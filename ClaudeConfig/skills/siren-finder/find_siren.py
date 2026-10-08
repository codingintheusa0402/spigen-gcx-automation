#!/usr/bin/env python3
"""
SIREN finder — finds SIREN-able cases (same product line-up × same defect with
claims + bad reviews >= THRESHOLD), drops the ones already in the CQ Emergency Net
registry, and reports the rest to Google Chat.

Sources (all read via the gws_shim OAuth token):
  claims   : Zendesk Raw Data_2026년 / '26년 전체문의' (1sjcCj_P4...) — Category '4. Product Issue',
             defect = '1차 Defect Reason or Inquiries'
  reviews  : APIFY_Axesso / 'SC' (1tMbA_msR...) — master Caspi+SC-scraper review table, ★1~3 only.
             No defect column there, so each review is classified once by the local Claude CLI
             into the same 1차 Defect taxonomy (cache: ~/.config/siren_finder/review_cls.json)
  registry : 2026_CQ_Spigen Issue Report Emergency Net Sheet_R01 (137K4hp...) — tabs
             Glass / Case / 전기전자 / 생활용품, cols SKU + 불량 유형 + 사유 ( 상세 )
  product  : product master 'Data' tab (1fx9K4r2T9...) — SKU/ASIN -> 기종명/모델명/생산업체

Usage:
  python3 find_siren.py                    # analyse, write candidates.json + chat preview, send nothing
  python3 find_siren.py --send             # also post the report to the webhook
  python3 find_siren.py --case N --export case.json   # siren-report data skeleton for candidate #N
Options: --since 2026-01-01 --threshold 5 --webhook URL --workers 6
"""
import argparse
import concurrent.futures as cf
import datetime
import json
import os
import re
import subprocess
import sys
import urllib.request
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lineup import ProductMaster, norm_model  # noqa: E402

from google.auth.transport.requests import Request  # noqa: E402
from google.oauth2.credentials import Credentials  # noqa: E402
from googleapiclient.discovery import build  # noqa: E402

TOKEN = os.path.expanduser("~/.config/gws_shim/token.json")
CLAUDE = os.path.expanduser("~/.local/bin/claude")
CACHE_DIR = os.path.expanduser("~/.config/siren_finder")
REVIEW_CLS = os.path.join(CACHE_DIR, "review_cls.json")
REGISTRY_CLS = os.path.join(CACHE_DIR, "registry_cls.json")
SECRETS = os.path.join(CACHE_DIR, "secrets.json")   # {"caspi_api_key": ...} chmod 600
CASPI_ENDPOINT = "https://caspilm.spigen.com/api/data-api/run"
Q_SALES = "pq_abe464407bdbd0de98"   # Amazon units by base SKU since date (FLAT_FILE_ALL_ORDERS_..., cancelled + amzn.gr resale excluded;
                                     # base SKU = REGEXP [A-Z]{3}[0-9]{5} since raw skus look like ACS09826PAN / ACS09826SGP)
OUT_DIR = os.path.join(CACHE_DIR, "runs")

CLAIMS_SHEET = ("1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc", "26년 전체문의")
REVIEWS_SHEET = ("1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw", "SC")
REGISTRY_ID = "137K4hpNfHoyb6PEb64gO7Wxi5b3CPlnMETQt-bbPKKE"
REGISTRY_TABS = ["Glass", "Case", "전기전자", "생활용품"]
PRODUCT_MASTER = ("1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo", "Data")
DEFAULT_WEBHOOK = json.load(open(os.path.expanduser("~/.config/gcx_webhooks.json")))["siren_private_room"]  # secret — kept out of git

GENERIC_MODELS = {"", "poweraccessories", "알수없음추가정보필요", "해당없음"}
DOMAINS = {"US": "amazon.com", "UK": "amazon.co.uk", "GB": "amazon.co.uk", "DE": "amazon.de",
           "FR": "amazon.fr", "IT": "amazon.it", "ES": "amazon.es", "JP": "amazon.co.jp", "IN": "amazon.in"}


def excluded_defect(d):
    """User rule 2026-09-28: never SIREN 황변, 배송 이슈, 중고품 배송 related defects."""
    return (not d) or ("황변" in d) or d.startswith("(Delivery Issue)") or ("배송" in d) or ("중고" in d)


def log(msg):
    print(f"[{datetime.datetime.now():%H:%M:%S}] {msg}", file=sys.stderr, flush=True)


def load_json(p, default):
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def save_json(p, obj):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, p)


def sheets():
    info = json.load(open(TOKEN))
    c = Credentials.from_authorized_user_info(info)
    c.refresh(Request())
    return build("sheets", "v4", credentials=c).spreadsheets()


def read_tab(svc, sid, tab):
    return svc.values().get(spreadsheetId=sid, range=f"'{tab}'").execute().get("values", [])


def col(header, row):
    ix = {k.strip(): i for i, k in enumerate(header)}
    return lambda k: (row[ix[k]] if k in ix and ix[k] < len(row) else "").replace("\xa0", "").strip()


def claude_json(prompt, model="sonnet", timeout=900):
    res = subprocess.run([CLAUDE, "-p", prompt, "--model", model, "--output-format", "text"],
                         capture_output=True, text=True, timeout=timeout)
    if res.returncode != 0:
        raise RuntimeError(f"claude CLI failed: {res.stderr.strip()[:300]}")
    m = re.search(r"\[.*\]", res.stdout, re.S)
    if not m:
        raise RuntimeError(f"no JSON array in claude output: {res.stdout[:300]}")
    return json.loads(m.group(0))


# ---------------------------------------------------------------- line-ups
class LineupIndex:
    def __init__(self, pm: ProductMaster):
        self.pm = pm
        self.names = {}   # lineup -> display name
        self.skus = defaultdict(set)
        self.asins = defaultdict(set)

    def of_claim(self, g):
        sku = g("★문의SKU")[:8].upper()
        lu = self.pm.lineup_from_sku(sku)
        if lu and lu[1] in GENERIC_MODELS:
            lu = ("", sku) if self.pm.by_sku.get(sku) else None   # power acc: SKU-level key
        if not lu:
            lu = self.pm.lineup_from_fields(g("Device"), g("Product Name"))
            if lu and lu[1] in GENERIC_MODELS:
                lu = None
        if lu:
            self._note(lu, sku, g("ASIN"), g("Device"), g("Product Name"))
        return lu

    def of_review(self, asin):
        rec = self.pm.by_asin.get(asin)
        if not rec:
            return None
        lu = self.pm.lineup_from_sku(rec["sku"])
        if lu and lu[1] in GENERIC_MODELS:
            lu = ("", rec["sku"])
        if lu:
            self._note(lu, rec["sku"], asin, rec["device"], rec["model"])
        return lu

    def _note(self, lu, sku, asin, device, model):
        if sku and sku in self.pm.by_sku:
            self.skus[lu].add(sku)
        if asin and asin.startswith("B0"):
            self.asins[lu].add(asin)
        if lu not in self.names:
            model = self.pm.by_sku.get(sku, {}).get("model") or model   # master name keeps line-ups distinct
            if lu[0]:
                self.names[lu] = f"{lu[0]} 시리즈용 {model or lu[1]}"
            else:
                rec = self.pm.by_sku.get(lu[1], {})
                name = rec.get("model") or model or lu[1]
                self.names[lu] = f"{name} ({lu[1]})" if rec else name


# ---------------------------------------------------------------- reviews
CLS_PROMPT = """You label Amazon bad reviews (★1-3) of Spigen products with the ONE defect type that best
matches the customer's main product complaint, choosing EXACTLY one string from the taxonomy below
(these are our Zendesk 1차 Defect names), or "none" if the review has no product defect
(price, shipping, seller, taste/design preference, wrong model ordered, vague dislike).
Match the product category: phone case -> (Case)_..., screen/camera protector -> (SP)_...,
chargers/cables/power banks -> (PAcc.)_..., straps/AirPods/Watch/other accessories -> (SDA)_... or (New Biz)_...
Return ONLY a JSON array, one object per review:
[{"id": "<REVIEW_ID>", "defect": "<taxonomy string or none>", "ko": "<Korean one-line summary, <= 50 chars>"}]

Taxonomy:
%s

Reviews:
%s"""


def classify_reviews(reviews, taxonomy, workers):
    cache = load_json(REVIEW_CLS, {})
    todo = [r for r in reviews if r["id"] not in cache]
    log(f"reviews: {len(reviews)} bad reviews, {len(todo)} need classification")
    tax = "\n".join(taxonomy)
    batches = [todo[i:i + 40] for i in range(0, len(todo), 40)]

    def run(batch):
        payload = [{"id": r["id"], "product": r["product"][:80], "title": r["title"][:150],
                    "text": r["text"][:500]} for r in batch]
        return claude_json(CLS_PROMPT % (tax, json.dumps(payload, ensure_ascii=False)))

    done = 0
    with cf.ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(run, b): b for b in batches}
        for f in cf.as_completed(futs):
            try:
                for it in f.result():
                    d = it.get("defect", "none")
                    cache[it["id"]] = {"defect": d if d in taxonomy else "none", "ko": it.get("ko", "")}
            except Exception as e:  # one bad batch must not kill the run; it retries next run
                log(f"batch failed ({e}); will retry next run")
            done += 1
            if done % 10 == 0 or done == len(batches):
                save_json(REVIEW_CLS, cache)
                log(f"classified batch {done}/{len(batches)}")
    save_json(REVIEW_CLS, cache)
    return cache


# ---------------------------------------------------------------- registry
REG_PROMPT = """For each SIREN candidate below, decide whether the SAME failure is already registered in the CQ
Emergency Net registry rows listed with it (rows are for the same product line-up SKUs).
registered=true ONLY if a row's 불량 유형 / 사유(상세) describes the same failure mode as the candidate's
defect (e.g. candidate 킥스탠드이슈 vs row "킥스탠드 고정 불량" = same). A row about a DIFFERENT failure on the
same product (e.g. candidate 재질 vs row 버튼부파손, candidate 형합 vs row 이염) is NOT a match -> false.
Return ONLY a JSON array: [{"key": "<key>", "registered": true|false, "row": "<tab No./일자 of the matching row, or empty>"}]

Candidates:
%s"""


def registry_rows(svc, pm):
    rows = []
    for tab in REGISTRY_TABS:
        v = read_tab(svc, REGISTRY_ID, tab)
        hdr_i = next((i for i, r in enumerate(v[:10]) if "SKU" in [c.strip() for c in r]), None)
        if hdr_i is None:
            continue
        hdr = v[hdr_i]
        for r in v[hdr_i + 1:]:
            g = col(hdr, r)
            sku = g("SKU")[:8].upper()
            if not sku:
                continue
            rows.append({"tab": tab, "no": g("No."), "date": g("일자"), "sku": sku,
                         "device": g("기종"), "type": g("불량 유형"), "detail": g("사유 ( 상세 )")[:200],
                         "lineup": pm.lineup_from_sku(sku)})
    return rows


def check_registry(cands, reg_rows, workers):
    cache = load_json(REGISTRY_CLS, {})
    by_lu = defaultdict(list)
    for r in reg_rows:
        if r["lineup"]:
            by_lu[tuple(r["lineup"])].append(r)
    todo = []
    for c in cands:
        rows = by_lu.get(c["lineup"], []) + [r for r in reg_rows if r["sku"] in c["skus"] and
                                              tuple(r["lineup"] or ()) != c["lineup"]]
        c["registry_rows"] = [f"{r['tab']} No.{r['no']} {r['date']} {r['sku']} [{r['type']}] {r['detail']}" for r in rows]
        if not rows:
            c["registered"], c["registered_row"] = False, ""
            continue
        sig = c["key"] + "|" + str(len(rows))
        if sig in cache:
            c["registered"], c["registered_row"] = cache[sig]["registered"], cache[sig]["row"]
        else:
            todo.append((sig, c))
    log(f"registry: {len(todo)} candidates need an LLM similarity check")
    chunks = [todo[i:i + 12] for i in range(0, len(todo), 12)]

    def run(chunk):
        text = "\n\n".join(f"key: {sig}\ncandidate: {c['name']} / defect {c['defect']}\nregistry rows:\n  "
                            + "\n  ".join(c["registry_rows"][:25]) for sig, c in chunk)
        return {x["key"]: x for x in claude_json(REG_PROMPT % text)}

    with cf.ThreadPoolExecutor(max_workers=workers) as ex:
        for chunk, fut in [(ch, ex.submit(run, ch)) for ch in chunks]:
            try:
                res = fut.result()
            except Exception as e:
                log(f"registry batch failed ({e}); treating as unregistered this run")
                res = {}
            for sig, c in chunk:
                x = res.get(sig)
                c["registered"] = bool(x and x.get("registered"))
                c["registered_row"] = (x or {}).get("row", "")
                if x:
                    cache[sig] = {"registered": c["registered"], "row": c["registered_row"]}
    save_json(REGISTRY_CLS, cache)


# ---------------------------------------------------------------- sales
def sales_by_sku(since):
    """{base_sku: {"units", "orders", "first"}} — Caspi S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL
    (all Amazon channels, non-cancelled) via the headless registered-query API (user rule 2026-09-28)."""
    key = load_json(SECRETS, {}).get("caspi_api_key")
    if not key:
        log("no caspi_api_key in secrets.json — skipping 판매량")
        return {}
    # limit must stay 1000: a 2000-row page came back silently cut (1,969 rows, nextOffset null)
    out, offset = {}, 0
    while True:
        body = json.dumps({"queryId": Q_SALES, "params": {"since_date": since}, "limit": 1000, "offset": offset}).encode()
        req = urllib.request.Request(CASPI_ENDPOINT, data=body, method="POST",
                                     headers={"x-api-key": key, "Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=180) as r:
            data = json.loads(r.read())
        for row in data["rows"]:
            out[row["BASE_SKU"].upper()] = {"units": int(float(row["UNITS"] or 0)),
                                            "orders": int(float(row["ORDERS"] or 0)), "first": row["FIRST_SALE"]}
        offset = (data.get("paging") or {}).get("nextOffset")
        if offset is None:
            return out


# ---------------------------------------------------------------- report
def sales_total(c, sales):
    return sum(sales.get(k, {}).get("units", 0) for k in c["skus"])


def write_sheet(svc_root, cands, sales, since, threshold):
    """Full candidate list -> new spreadsheet (tab 미등록 + tab 기등록). Returns its URL."""
    hdr = ["No.", "라인업(기종+제품)", "불량 유형", "합계", "클레임", "배드리뷰", "최초", "최근", "SKU",
           "판매량 합계(Amazon)", "클레임+리뷰율", "SKU별 판매량", "등록 여부", "매칭 등록 행",
           "클레임 링크", "리뷰 링크"]

    def rows(lst):
        out = [hdr]
        for i, c in enumerate(lst, 1):
            tot = sales_total(c, sales)
            out.append([i, c["name"], c["defect"].split("_", 1)[-1], c["total"], len(c["claims"]), len(c["reviews"]),
                        c["first"], c["last"], ", ".join(c["skus"]), tot or "",
                        f"{c['total'] / tot * 100:.2f}%" if tot else "",
                        ", ".join(f"{k} {sales.get(k, {}).get('units', 0):,}" for k in c["skus"]),
                        "기등록" if c["registered"] else "미등록", c.get("registered_row", ""),
                        "\n".join(f"https://spigenhelp.zendesk.com/agent/tickets/{t['id']}" for t in c["claims"])[:45000],
                        "\n".join(r["url"] for r in c["reviews"])[:45000]])
        return out

    open_c = [c for c in cands if not c["registered"]]
    reg_c = [c for c in cands if c["registered"]]
    title = f"SIREN 후보 리포트_{datetime.date.today():%y%m%d} (since {since}, ≥{threshold})"
    sp = svc_root.create(body={"properties": {"title": title},
                               "sheets": [{"properties": {"title": "미등록"}}, {"properties": {"title": "기등록"}}]},
                         fields="spreadsheetId,spreadsheetUrl").execute()
    svc_root.values().batchUpdate(spreadsheetId=sp["spreadsheetId"], body={"valueInputOption": "RAW", "data": [
        {"range": "'미등록'!A1", "values": rows(open_c)}, {"range": "'기등록'!A1", "values": rows(reg_c)}]}).execute()
    return sp["spreadsheetUrl"]


def chat_text(cands, since, threshold, total_found, registered_n, sales, sheet_url, top):
    head = (f"*[SIREN 후보 리포트]* {datetime.date.today():%Y-%m-%d}\n"
            f"기간 {since} ~ · 동일 라인업×동일 불량 클레임+배드리뷰 *{threshold}건 이상* {total_found}건 중 "
            f"CQ Emergency Net 기등록 {registered_n}건 제외 → *미등록 {len(cands)}건*\n"
            f"(제외 불량: 황변·배송/중고품 관련 · 판매량 = Caspi Amazon 주문, 기간 내·취소 제외)\n"
            f"전체 목록(링크·SKU별 판매량 포함): <{sheet_url}|SIREN 후보 시트>\n\n*상위 {min(top, len(cands))}건*")
    lines = []
    for i, c in enumerate(cands[:top], 1):
        tl = " ".join(f"<https://spigenhelp.zendesk.com/agent/tickets/{t['id']}|#{j + 1}>" for j, t in enumerate(c["claims"][:3]))
        rl = " ".join(f"<{r['url']}|R{j + 1}>" for j, r in enumerate(c["reviews"][:3]))
        tot = sales_total(c, sales)
        sl = f"판매량 {tot:,}개 · 율 {c['total'] / tot * 100:.2f}%" if tot else "판매량 -"
        lines.append(f"*{i}. {c['name']}* — {c['defect'].split('_', 1)[-1]} | *{c['total']}건* "
                     f"(클레임 {len(c['claims'])} · 리뷰 {len(c['reviews'])}) {c['first']}~{c['last']} | {sl}"
                     + (f"\n    클레임 {tl}" if tl else "") + (f"  리뷰 {rl}" if rl else ""))
    msgs, cur = [], head
    for ln in lines:
        if len(cur) + len(ln) + 2 > 3800:
            msgs.append(cur)
            cur = "(계속)"
        cur += "\n" + ln
    msgs.append(cur)
    return msgs


def post(webhook, text):
    req = urllib.request.Request(webhook, data=json.dumps({"text": text}, ensure_ascii=False).encode(),
                                 method="POST", headers={"Content-Type": "application/json; charset=UTF-8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read()).get("name")


def export_case(c, pm):
    """Skeleton data JSON for ~/.claude/skills/siren-report/build_siren_slides.py.
    detail_text/image_urls still need filling from Zendesk (see SKILL.md step 5)."""
    items = []
    for t in c["claims"]:
        items.append({"label": "클레임", "country": t["country"], "purchase_date": t["purchase"] or "-",
                      "inflow_date": t["created"], "days_elapsed": t["elapsed"], "link_label": t["id"],
                      "detail_text": f"[{t['sku']} / {t['asin']}] ", "image_urls": []})
    for r in c["reviews"]:
        items.append({"label": "배드 리뷰", "country": r["country"], "purchase_date": "-",
                      "inflow_date": r["created"], "days_elapsed": "-", "link_label": r["id"], "link_url": r["url"],
                      "detail_text": f"[{r['asin']} / ★{r['rating']}] {r['ko']}", "original_text": r["original"],
                      "image_urls": [r["image"]] if r["image"] else []})
    by_z, by_r = defaultdict(int), defaultdict(int)
    for t in c["claims"]:
        by_z[t["country"]] += 1
    for r in c["reviews"]:
        by_r[r["country"]] += 1
    skus = sorted(c["skus"])
    sku_asins = [pm.by_sku.get(s, {}).get("asin", "") for s in skus]
    # 제품명 link + Global 리뷰 평점 source: the case's own product-master ASIN (valid 10-char), never a claim-typed one
    own = [a for a in sku_asins if re.fullmatch(r"B0[A-Z0-9]{8}", a or "")]
    main_asin = own[0] if own else next((a for a in sorted(c["asins"]) if re.fullmatch(r"B0[A-Z0-9]{8}", a)), "")
    return {"issue_title": f"{c['name']} {c['defect'].split('_', 1)[-1]} 이슈", "author": "김지우",
            "product": {"device": c["lineup"][0] or "-", "product_name": c["name"], "defect_type": c["defect"].split("_", 1)[-1],
                        "skus": [{"sku": s, "asin": pm.by_sku.get(s, {}).get("asin", "")} for s in skus],
                        "manufacturer": "", "asin": main_asin},
            "overview": {"total_count": c["total"], "bad_review_count": len(c["reviews"]),
                         "zendesk_count": len(c["claims"]), "global_rating": "",  # "" -> siren-report fills it from Amazon (DE→US→UK→JP)
                         "by_country_bad_review": dict(by_r), "by_country_zendesk": dict(by_z)},
            "claims": items}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--since", default=f"{datetime.date.today().year}-01-01")
    ap.add_argument("--threshold", type=int, default=5)
    ap.add_argument("--send", action="store_true")
    ap.add_argument("--webhook", default=DEFAULT_WEBHOOK)
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--case", type=int)
    ap.add_argument("--export")
    ap.add_argument("--top", type=int, default=30, help="cases listed in the chat message (all go to the sheet)")
    a = ap.parse_args()

    svc = sheets()
    pm = ProductMaster(read_tab(svc, *PRODUCT_MASTER))
    idx = LineupIndex(pm)

    # claims
    cv = read_tab(svc, *CLAIMS_SHEET)
    groups = defaultdict(lambda: {"claims": [], "reviews": []})
    taxonomy = set()
    for r in cv[1:]:
        g = col(cv[0], r)
        if g("Category") != "4. Product Issue" or g("Ticket created - Date") < a.since:
            continue
        d = g("1차 Defect Reason or Inquiries")
        if d:
            taxonomy.add(d)
        if excluded_defect(d):
            continue
        lu = idx.of_claim(g)
        if not lu:
            continue
        created, purchase = g("Ticket created - Date"), g("Purchase Date - Date")
        try:
            elapsed = f"{(datetime.date.fromisoformat(created) - datetime.date.fromisoformat(purchase)).days}일"
        except ValueError:
            elapsed = "-"
        groups[(lu, d)]["claims"].append({"id": r[0].strip(), "created": created, "purchase": purchase,
                                          "elapsed": elapsed, "country": g("Country"), "sku": g("★문의SKU"),
                                          "asin": g("ASIN"), "device": g("Device"), "product": g("Product Name")})
    log(f"claims: {sum(len(v['claims']) for v in groups.values())} in {len(groups)} line-up×defect groups")

    # reviews
    sv = read_tab(svc, *REVIEWS_SHEET)
    bad = []
    for r in sv[1:]:
        g = col(sv[0], r)
        if g("Review Ratings") not in ("1", "2", "3") or not g("Review ID"):
            continue
        created = g("Created 날짜")
        if re.match(r"\d{4}-\d{2}-\d{2}$", created) and created < a.since:
            continue
        lu = idx.of_review(g("ASIN"))
        if not lu:
            continue
        mk = g("국가").upper()
        bad.append({"id": g("Review ID"), "lineup": lu, "asin": g("ASIN"), "rating": g("Review Ratings"),
                    "created": created, "country": "UK" if mk == "GB" else mk, "title": g("Review Title"),
                    "text": re.sub(r"<br\s*/?>", " ", g("본문")), "image": g("Image URL").split("|")[0],
                    "url": g("Review Link") or f"https://{DOMAINS.get(mk, 'amazon.com')}/gp/customer-reviews/{g('Review ID')}",
                    "product": idx.names.get(lu, "")})
    seen = set()
    bad = [b for b in bad if not (b["id"] in seen or seen.add(b["id"]))]
    cls = classify_reviews(bad, sorted(taxonomy), a.workers)
    for b in bad:
        c = cls.get(b["id"])
        if not c or c["defect"] == "none" or excluded_defect(c["defect"]):
            continue
        b["ko"] = c["ko"]
        b["original"] = f"{b['title']}\n{b['text']}"
        groups[(b["lineup"], c["defect"])]["reviews"].append(b)

    # candidates
    cands = []
    for (lu, d), v in groups.items():
        total = len(v["claims"]) + len(v["reviews"])
        if total < a.threshold:
            continue
        dates = sorted([x["created"] for x in v["claims"] + v["reviews"] if re.match(r"\d{4}-", x["created"])])
        cands.append({"key": f"{lu[0]}|{lu[1]}|{d}", "lineup": lu, "defect": d, "name": idx.names.get(lu, lu[1]),
                      "claims": sorted(v["claims"], key=lambda x: x["created"]),
                      "reviews": sorted(v["reviews"], key=lambda x: x["created"]),
                      "total": total, "skus": sorted(idx.skus[lu]), "asins": sorted(idx.asins[lu]),
                      "first": dates[0] if dates else "-", "last": dates[-1] if dates else "-"})
    cands.sort(key=lambda c: -c["total"])
    log(f"candidates >= {a.threshold}: {len(cands)}")
    check_registry(cands, registry_rows(svc, pm), a.workers)
    open_c = [c for c in cands if not c["registered"]]

    run_path = os.path.join(OUT_DIR, f"candidates_{datetime.date.today():%y%m%d}.json")
    save_json(run_path, [dict(c, lineup=list(c["lineup"])) for c in cands])
    log(f"wrote {run_path}")

    if a.case:
        c = open_c[a.case - 1]
        save_json(a.export, export_case(c, pm))
        log(f"exported case #{a.case} ({c['name']} / {c['defect']}) -> {a.export}")
        return

    sales = sales_by_sku(a.since)
    log(f"sales: {len(sales)} base SKUs")
    if a.send:
        url = write_sheet(svc, cands, sales, a.since, a.threshold)
        log(f"sheet: {url}")
    else:
        url = "(sheet created on --send)"
    msgs = chat_text(open_c, a.since, a.threshold, len(cands), len(cands) - len(open_c), sales, url, a.top)
    if a.send:
        for m in msgs:
            log(f"sent {post(a.webhook, m)}")
    else:
        print("\n\n---\n\n".join(msgs))


if __name__ == "__main__":
    main()
