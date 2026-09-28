#!/usr/bin/env python3
"""
이염/변색 claim + bad-review twice-weekly report (Mon/Fri 10:00 KST) -> Google Chat webhook.

Data comes from Caspi via two registered queries called headlessly through the data-api
endpoint (no Claude/MCP session needed):
  - pq_cde9dcd490cc850dad  Zendesk tickets tagged `_case__이염/변색` (1차 Defect), names decoded
  - pq_3cc3dd0f30778bf1f8  Spigen Amazon ★1~3 reviews passing a broad multilingual keyword
                           prefilter (discolor/stain/jeans/verfärb/変色/…)
Review candidates are then classified by the local Claude CLI (`claude -p`) — the keyword
prefilter alone is ~50% noise (stainless, 황변/yellowing, screen-protector smudges, …).
Each review is classified once and cached in state.json.

Report = items not yet reported (created within LOOKBACK_NEW_DAYS), grouped by base SKU,
with each group's 90-day cumulative claim+review count; groups at >= SIREN_THRESHOLD are
flagged for SIREN review.

Schedule: launchd runs this every 30 min (+ at load); the script self-gates to Mon/Fri
after 10:00 KST and sends at most once per day, so a Mac that was asleep/off at 10:00
catches up on the first tick after it wakes.

  python3 report.py            # gated scheduled run
  python3 report.py --force    # ignore the day/time/once-per-day gate
  python3 report.py --dry-run  # build + print the message, send nothing, don't touch state
"""
import argparse
import datetime
import json
import os
import re
import subprocess
import sys
import urllib.request
from collections import defaultdict

CASPI_ENDPOINT = "https://caspilm.spigen.com/api/data-api/run"
Q_CLAIMS = "pq_cde9dcd490cc850dad"
Q_REVIEWS = "pq_3cc3dd0f30778bf1f8"

CONF_DIR = os.path.expanduser("~/.config/discoloration_report")
SECRETS = os.path.join(CONF_DIR, "secrets.json")   # caspi_api_key, webhook_url (chmod 600)
STATE = os.path.join(CONF_DIR, "state.json")
CLAUDE_BIN = os.path.expanduser("~/.local/bin/claude")

CUMULATIVE_DAYS = 90
LOOKBACK_NEW_DAYS = 14
SIREN_THRESHOLD = 3
SEND_WEEKDAYS = (0, 4)  # Mon, Fri
SEND_HOUR = 10
CLASSIFY_BATCH = 25

DOMAINS = {"US": "amazon.com", "UK": "amazon.co.uk", "GB": "amazon.co.uk", "DE": "amazon.de",
           "FR": "amazon.fr", "IT": "amazon.it", "ES": "amazon.es", "JP": "amazon.co.jp",
           "IN": "amazon.in", "CA": "amazon.ca", "NL": "amazon.nl", "MX": "amazon.com.mx",
           "AU": "amazon.com.au"}
WEEKDAY_KO = "월화수목금토일"


def log(msg):
    print(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def save_state(state):
    tmp = STATE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE)


def caspi(query_id, params, api_key):
    rows, offset = [], 0
    while True:
        body = json.dumps({"queryId": query_id, "params": params, "limit": 1000, "offset": offset}).encode()
        req = urllib.request.Request(CASPI_ENDPOINT, data=body, method="POST",
                                     headers={"x-api-key": api_key, "Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read())
        rows.extend(data["rows"])
        nxt = (data.get("paging") or {}).get("nextOffset")
        if nxt is None:
            return rows
        offset = nxt


def base_sku(sku):
    return (sku or "").strip()[:8].upper() or None


def review_url(r):
    mk = (r.get("MARKETPLACE") or "").upper()
    return f"https://{DOMAINS.get(mk, 'amazon.com')}/gp/customer-reviews/{r['REVIEW_ID']}/ref_=bcr_shw_rev_dtl"


def clean_text(s, n):
    s = re.sub(r"<br\s*/?>", " ", s or "")
    return re.sub(r"\s+", " ", s).strip()[:n]


CLASSIFY_PROMPT = """You classify Amazon reviews of Spigen products (phone cases, straps, accessories).
Decide for each review whether its complaint is 이염/변색:
- 이염: color transferring onto the product (from jeans/denim, clothing, dye, pockets, hands) and not washing off
- 변색: the product's own color changing/darkening/fading/turning pink/blue/black/brown, or stains/grime that won't clean off
NOT a hit (hit=false):
- 황변: yellowing of clear/transparent/white cases over time (tracked separately)
- stainless steel mentions, screen-protector smudges/bubbles/residue, glue/sticky residue, dirt that cleans off,
  or reviews where discoloration is not actually the complaint.
Return ONLY a JSON array, one object per input review, no prose:
[{"id": "<REVIEW_ID>", "hit": true|false, "summary": "<Korean one-line summary of the complaint, <= 60 chars>"}]

Reviews:
"""


def classify(reviews):
    """{review_id: {"hit": bool, "summary": str}} via the local Claude CLI, in batches."""
    out = {}
    for i in range(0, len(reviews), CLASSIFY_BATCH):
        batch = reviews[i:i + CLASSIFY_BATCH]
        payload = [{"id": r["REVIEW_ID"], "product": clean_text(r.get("ASIN_TITLE"), 90),
                    "rating": r.get("REVIEW_RATING"), "title": clean_text(r.get("REVIEW_TITLE"), 150),
                    "text": clean_text(r.get("REVIEW_TEXT"), 700)} for r in batch]
        prompt = CLASSIFY_PROMPT + json.dumps(payload, ensure_ascii=False)
        res = subprocess.run([CLAUDE_BIN, "-p", prompt, "--model", "sonnet", "--output-format", "text"],
                             capture_output=True, text=True, timeout=600)
        if res.returncode != 0:
            raise RuntimeError(f"claude CLI failed: {res.stderr.strip()[:300]}")
        m = re.search(r"\[.*\]", res.stdout, re.S)
        if not m:
            raise RuntimeError(f"claude CLI returned no JSON array: {res.stdout[:300]}")
        for item in json.loads(m.group(0)):
            out[item["id"]] = {"hit": bool(item.get("hit")), "summary": item.get("summary", "")}
        missing = [r["REVIEW_ID"] for r in batch if r["REVIEW_ID"] not in out]
        if missing:
            raise RuntimeError(f"classifier skipped {len(missing)} reviews: {missing[:5]}")
        log(f"classified {min(i + CLASSIFY_BATCH, len(reviews))}/{len(reviews)}")
    return out


def build_message(now, new_claims, new_reviews, cum, labels, asins, cls, window_from):
    wd = WEEKDAY_KO[now.weekday()]
    head = (f"*[이염/변색 클레임·배드리뷰 정기 리포트]* {now:%Y-%m-%d}({wd}) {now:%H:%M}\n"
            f"신규 ({window_from:%m-%d} 이후 인입, 미보고분): 클레임 *{len(new_claims)}건* · "
            f"배드리뷰 *{len(new_reviews)}건*")
    groups = defaultdict(lambda: {"claims": [], "reviews": []})
    for c in new_claims:
        groups[base_sku(c.get("SKU")) or f"ASIN:{c.get('ASIN')}"]["claims"].append(c)
    for r in new_reviews:
        groups[base_sku(r.get("MAPPED_SKU")) or f"ASIN:{r.get('CHILD_ASIN')}"]["reviews"].append(r)
    siren = sorted(k for k, v in cum.items() if v["claims"] + v["reviews"] >= SIREN_THRESHOLD and k in groups)
    if siren:
        head += f"\n🚨 90일 누적 {SIREN_THRESHOLD}건 이상 (SIREN 검토 대상): *{len(siren)}개 제품*"
    if not groups:
        return head + "\n\n이번 기간 신규 이염/변색 클레임·배드리뷰가 없습니다."

    def total(k):
        v = cum.get(k, {"claims": 0, "reviews": 0})
        return v["claims"] + v["reviews"]

    blocks = [head]
    for k in sorted(groups, key=lambda k: (-total(k), k)):
        g, c90 = groups[k], cum.get(k, {"claims": 0, "reviews": 0})
        asin = asins.get(k)
        asin_part = f" · <https://www.amazon.de/dp/{asin}|{asin}>" if asin else ""
        flag = " 🚨SIREN 검토" if total(k) >= SIREN_THRESHOLD else ""
        lines = [f"\n*{labels.get(k, k)}* [{k if not k.startswith('ASIN:') else '-'}{asin_part}]\n"
                 f"신규 클레임 {len(g['claims'])} · 배드리뷰 {len(g['reviews'])} | "
                 f"90일 누적 {total(k)}건 (클레임 {c90['claims']} · 리뷰 {c90['reviews']}){flag}"]
        for c in g["claims"]:
            lines.append(f"*<https://spigenhelp.zendesk.com/agent/tickets/{c['TICKET_ID']}|[클레임]>* "
                         f"({(c.get('COUNTRY') or '-').upper()}) 인입 {c['CREATED']}"
                         + (f" · 구매 {c['PURCHASE_DATE']}" if c.get("PURCHASE_DATE") else ""))
        for r in g["reviews"]:
            mk = "UK" if r.get("MARKETPLACE") == "GB" else r.get("MARKETPLACE")
            lines.append(f"*<{review_url(r)}|[배드리뷰]>* ({mk} ★{r.get('REVIEW_RATING')}) {r['CREATED']} — "
                         f"{cls[r['REVIEW_ID']]['summary']}")
        blocks.append("\n".join(lines))
    return "\n".join(blocks)


def send(webhook, text, photos):
    payload = {"text": text}
    if photos:
        payload["cardsV2"] = [{"cardId": "photos", "card": {"sections": [{
            "header": f"배드리뷰 고객 사진 ({len(photos)}장)",
            "widgets": [{"grid": {"columnCount": 3, "items": [
                {"id": str(i), "image": {"imageUri": u}} for i, u in enumerate(photos)]}}]}]}}]
    req = urllib.request.Request(webhook, data=json.dumps(payload, ensure_ascii=False).encode(), method="POST",
                                 headers={"Content-Type": "application/json; charset=UTF-8"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read()).get("name")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    now = datetime.datetime.now()  # Mac clock is KST
    today = now.date()
    state = load_json(STATE, {"last_sent_date": None, "reported_claims": [], "reported_reviews": [],
                              "review_cls": {}})
    if not a.force and not a.dry_run:
        if today.weekday() not in SEND_WEEKDAYS or now.hour < SEND_HOUR:
            return
        if state.get("last_sent_date") == today.isoformat():
            return

    secrets = load_json(SECRETS, None)
    since90 = (today - datetime.timedelta(days=CUMULATIVE_DAYS)).isoformat()
    window_from = today - datetime.timedelta(days=LOOKBACK_NEW_DAYS)
    claims = caspi(Q_CLAIMS, {"since_date": since90}, secrets["caspi_api_key"])
    reviews = caspi(Q_REVIEWS, {"since_date": since90}, secrets["caspi_api_key"])
    log(f"fetched {len(claims)} claims, {len(reviews)} review candidates since {since90}")

    cls = state.setdefault("review_cls", {})
    todo = [r for r in reviews if r["REVIEW_ID"] not in cls]
    if todo:
        cls.update(classify(todo))
        save_state(state)  # cache classifications even on --dry-run (they never change)
    hits = [r for r in reviews if cls.get(r["REVIEW_ID"], {}).get("hit")]

    # 90-day cumulative per base SKU + display labels (claims give "Device용 Product")
    cum = defaultdict(lambda: {"claims": 0, "reviews": 0})
    labels, asins = {}, {}
    for c in claims:
        k = base_sku(c.get("SKU")) or f"ASIN:{c.get('ASIN')}"
        cum[k]["claims"] += 1
        if c.get("DEVICE") and c.get("PRODUCT"):
            labels.setdefault(k, f"{c['DEVICE']}용 {c['PRODUCT']}")
        if c.get("ASIN"):
            asins.setdefault(k, c["ASIN"])
    for r in hits:
        k = base_sku(r.get("MAPPED_SKU")) or f"ASIN:{r.get('CHILD_ASIN')}"
        cum[k]["reviews"] += 1
        labels.setdefault(k, clean_text(r.get("ASIN_TITLE"), 70))
        asins.setdefault(k, r.get("CHILD_ASIN"))

    reported_c, reported_r = set(state["reported_claims"]), set(state["reported_reviews"])
    new_claims = [c for c in claims if c["TICKET_ID"] not in reported_c and c["CREATED"] >= window_from.isoformat()]
    new_reviews = [r for r in hits if r["REVIEW_ID"] not in reported_r and r["CREATED"] >= window_from.isoformat()]
    text = build_message(now, new_claims, new_reviews, cum, labels, asins, cls, window_from)
    photos = [r["PRIMARY_PHOTO_URL"] for r in new_reviews if r.get("PRIMARY_PHOTO_URL")][:30]

    if a.dry_run:
        print(text)
        print(f"\n(photos: {len(photos)})")
        return
    name = send(secrets["webhook_url"], text, photos)
    log(f"sent {name}: {len(new_claims)} claims, {len(new_reviews)} reviews")
    state["reported_claims"] = sorted(reported_c | {c["TICKET_ID"] for c in new_claims})
    state["reported_reviews"] = sorted(reported_r | {r["REVIEW_ID"] for r in new_reviews})
    state["last_sent_date"] = today.isoformat()
    save_state(state)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"ERROR: {e}")
        sys.exit(1)
