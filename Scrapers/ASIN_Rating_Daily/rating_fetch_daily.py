#!/usr/bin/env python3
"""Unified daily ASIN rating/review-count fetch job.

Replaces the per-product paid `axesso_data/amazon-product-details-scraper` Apify tasks
with one actor (cosmic_meteorite/amazon-asin-rating-scraper, 2AuZCyLxKZuO2OyNY), called
once per marketplace across every product, and writes results directly into:

  Group A — the 8 review-monitoring books' country-code tabs (Score / Global Ratings cols).
            Market is resolved PER ROW from that row's own '국가' cell, never the tab name.
            Tabs whose Score/Global columns are still a live ARRAYFORMULA get a one-time
            migration (formula cleared, plain header text written) before static writes.
  Group B — 2 side-spreadsheets (iPhone18Fold sales sheet's 4 tabs, LJH's SKU세일즈/리뷰) —
            rating only, with a per-tab primary market + fallback cascade, and a cell Note
            whenever the value written came from a fallback market.

Usage:
  rating_fetch_daily.py --pilot-only            # GlxZ8 DE tab only, dry-run
  rating_fetch_daily.py --pilot-only --commit   # GlxZ8 DE tab only, writes for real
  rating_fetch_daily.py --commit                # full run across every book/tab
  rating_fetch_daily.py --commit --run-label "run 2/4"
"""
import argparse, json, os, sys, time, urllib.request
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_DIR = os.path.expanduser("~/.config/asin_rating_daily")
os.makedirs(STATE_DIR, exist_ok=True)
LOG_PATH = os.path.join(STATE_DIR, "run.log")

APIFY_ACTOR = "2AuZCyLxKZuO2OyNY"
APIFY_TOKEN = open(os.path.expanduser("~/.config/apify_asin_scraper/token")).read().strip()
APIFY_API = "https://api.apify.com/v2"

GWS_TOKEN_PATH = os.path.expanduser("~/.config/gws_shim/token.json")

ASIN_MASTER_ID = "1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo"
ASIN_MASTER_TAB = "Data"

PRIVATE_WEBHOOK = ("https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages"
                   "?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI"
                   "&token=t0A95TCkAqXjLGYBoZrgZSG_XuSSB9UDL3GsFJX_YDc")  # "siren_private_room"

# 국가 cell value (uppercased, trimmed) -> actor marketplace code
MARKET_CODE = {
    "US": "com", "COM": "com",
    "UK": "co.uk", "GB": "co.uk", "CO.UK": "co.uk",
    "DE": "de",
    "FR": "fr",
    "IT": "it",
    "ES": "es",
    "JP": "co.jp", "CO.JP": "co.jp",
    "IN": "in",
    "CA": "ca",
}

# ---- Group A: the 8 review-monitoring books. Columns are discovered live per tab —
# only the spreadsheet id + the list of country-code tabs to scan is hardcoded here.
GROUP_A_BOOKS = {
    "GlxZ8":     ("19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4", ["DE", "UK", "FR", "IT", "ES", "JP", "IN", "US"]),
    "Pixel11":   ("12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI", ["DE", "UK", "FR", "IT", "ES", "JP", "IN", "US"]),
    "iPh18":     ("1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU", ["DE", "UK", "FR", "IT", "ES", "JP", "IN", "US"]),
    "유지훈P":     ("1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw", ["US"]),  # single tab, multi-market via 국가 col
    "SDA":       ("1sxapIqJgXcJdeqyCf9bAxCNXrVMsVjsZE9QWPwEm0R4", ["DE", "UK", "FR", "IT", "ES", "JP", "US"]),
    "Auto_Acc":  ("1mEYb1b92D6BIOaSYkAnMit6THuw5ewtymhA-mSIVDfs", ["DE", "UK", "FR", "IT", "ES"]),
    "Power_Acc": ("1QC8Is6UvTnFXaOeXviKM_331i3Fo_CBIYx80VS696LI", ["IN", "DE", "UK", "FR", "IT", "ES", "US", "CA"]),
    "전략폰":      ("1yo8CbLhJkuxrf3eXbAqZCb6qBejZhSR3YOt7nFv97fw", ["IN", "DE", "UK", "FR", "IT", "ES"]),
}

# ---- Group B: rating-only, with primary market + fallback cascade + notes.
GROUP_B_TABS = [
    # (label, spreadsheet_id, tab, primary_market_code, fallback_codes_in_order)
    ("iPhone Fold",    "1XXIMDNVMPBut8WTCV-4gZt4DSxHmGaRR8GkRejAOGq0", "iPhone Fold",    "de", ["com", "co.uk", "co.jp"]),
    ("iPhone 18",      "1XXIMDNVMPBut8WTCV-4gZt4DSxHmGaRR8GkRejAOGq0", "iPhone 18",      "de", ["com", "co.uk", "co.jp"]),
    ("Apple ETC(26)",  "1XXIMDNVMPBut8WTCV-4gZt4DSxHmGaRR8GkRejAOGq0", "Apple ETC(26)",  "co.uk", ["de", "com", "co.jp"]),
    ("Pixel 11",       "1XXIMDNVMPBut8WTCV-4gZt4DSxHmGaRR8GkRejAOGq0", "Pixel 11",       "co.uk", ["de", "com", "co.jp"]),
    ("SKU세일즈/리뷰",     "1_9O8oTHt-yHewG0aGeioeM0psZ2kprf3lgH77-BFLs8", "SKU세일즈/리뷰",     "de", ["com", "co.uk", "co.jp"]),
]

MARKET_LABEL = {"com": "amazon.com", "co.uk": "amazon.co.uk", "de": "amazon.de", "fr": "amazon.fr",
                "it": "amazon.it", "es": "amazon.es", "co.jp": "amazon.co.jp", "in": "amazon.in", "ca": "amazon.ca"}


def log(msg):
    line = f"{datetime.now():%F %T} {msg}"
    print(line)
    with open(LOG_PATH, "a") as f:
        f.write(line + "\n")


# ---------------------------------------------------------------- Sheets API
def get_sheets_service():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    import httplib2
    from google_auth_httplib2 import AuthorizedHttp

    with open(GWS_TOKEN_PATH) as f:
        info = json.load(f)
    creds = Credentials(
        token=info.get("token"), refresh_token=info.get("refresh_token"),
        token_uri=info.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=info.get("client_id"), client_secret=info.get("client_secret"),
        scopes=info.get("scopes"),
    )
    creds.refresh(Request())
    info["token"] = creds.token
    with open(GWS_TOKEN_PATH, "w") as f:
        json.dump(info, f)
    http = AuthorizedHttp(creds, http=httplib2.Http(timeout=300))
    return build("sheets", "v4", http=http)


def col_letter(idx0):
    """0-based column index -> A1 letter(s)."""
    idx = idx0 + 1
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def col_idx(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


HEADER_ALIASES = {
    "asin": ["ASIN"],
    "sku": ["SKU", "SKU "],
    "country": ["국가"],
    "score": ["Score"],
    "global": ["Global Ratings", "Global ratings", "Global Rating"],
}


def find_header(svc, sid, tab, max_rows=10, want=("asin", "sku", "country", "score", "global")):
    """Scan the first max_rows rows for a header row containing ASIN (+ at least one more wanted field).
    Returns (header_row_1based, {field: col_letter}) or (None, {}) if not found."""
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{tab}'!A1:Z{max_rows}", valueRenderOption="UNFORMATTED_VALUE"
    ).execute(num_retries=8)
    rows = resp.get("values", [])
    for r_idx, row in enumerate(rows):
        norm = {str(c).strip(): c_idx for c_idx, c in enumerate(row) if isinstance(c, str) and c.strip()}
        found = {}
        for field in want:
            for alias in HEADER_ALIASES[field]:
                if alias in norm:
                    found[field] = col_letter(norm[alias])
                    break
        if "asin" in found:
            return r_idx + 1, found
    return None, {}


def read_formula(svc, sid, tab, a1_range):
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{tab}'!{a1_range}", valueRenderOption="FORMULA"
    ).execute(num_retries=8)
    vals = resp.get("values", [])
    return vals[0][0] if vals and vals[0] else ""


def last_data_row(svc, sid, tab, asin_col_letter, header_row):
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{tab}'!{asin_col_letter}{header_row+1}:{asin_col_letter}5000",
        valueRenderOption="UNFORMATTED_VALUE"
    ).execute(num_retries=8)
    vals = resp.get("values", [])
    last = header_row
    for i, row in enumerate(vals):
        if row and str(row[0]).strip():
            last = header_row + 1 + i
    return last


def migrate_arrayformula_if_present(svc, sid, tab, header_row, score_col, global_col, dry_run):
    """If Score/Global columns are driven by a live ARRAYFORMULA (owner cell at header_row or
    header_row-1), clear it and write plain header text. No-op if already static. Returns True
    if a migration happened (or would happen, under dry-run)."""
    migrated = False
    for col, label in [(score_col, "Score"), (global_col, "Global Ratings")]:
        if not col:
            continue
        owner_row = None
        for candidate_row in (header_row, max(1, header_row - 1)):
            f = read_formula(svc, sid, tab, f"{col}{candidate_row}")
            if isinstance(f, str) and f.startswith("=") and "ARRAYFORMULA" in f.upper():
                owner_row = candidate_row
                break
        if owner_row is None:
            continue
        migrated = True
        log(f"  [{tab}] migrating {label} col {col}: clearing ARRAYFORMULA at row {owner_row}, writing plain header")
        if not dry_run:
            svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{tab}'!{col}{owner_row}").execute(num_retries=8)
            svc.spreadsheets().values().update(
                spreadsheetId=sid, range=f"'{tab}'!{col}{header_row}",
                valueInputOption="RAW", body={"values": [[label]]}
            ).execute(num_retries=8)
    return migrated


def verify_no_errors(svc, sid, tab, score_col, global_col, header_row, last_row):
    """Pilot safety check: scan the Score/Global columns for #REF!/#ERROR!/#N/A after migration."""
    bad = []
    for col in [score_col, global_col]:
        if not col:
            continue
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{tab}'!{col}{header_row}:{col}{last_row}",
            valueRenderOption="FORMATTED_VALUE"
        ).execute(num_retries=8)
        for i, row in enumerate(resp.get("values", [])):
            v = row[0] if row else ""
            if isinstance(v, str) and v.startswith("#"):
                bad.append((f"{col}{header_row+i}", v))
    return bad


# ---------------------------------------------------------------- Apify
def apify_run_sync(asins, marketplace, timeout=280):
    if not asins:
        return []
    url = f"{APIFY_API}/acts/{APIFY_ACTOR}/run-sync-get-dataset-items?token={APIFY_TOKEN}&timeout={timeout}"
    body = json.dumps({"asins": asins, "marketplace": marketplace}).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout + 30) as r:
        return json.load(r)


def chunked(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def fetch_market(asins, marketplace, stats):
    """Call the actor for one marketplace, chunked, return {asin: {found, rating, ratingsCount}}."""
    out = {}
    uniq = sorted(set(asins))
    for chunk in chunked(uniq, 300):
        stats["actor_calls"] += 1
        stats["actor_calls_by_market"][marketplace] = stats["actor_calls_by_market"].get(marketplace, 0) + 1
        try:
            items = apify_run_sync(chunk, marketplace)
        except Exception as e:
            log(f"  ERROR calling actor for marketplace={marketplace} chunk of {len(chunk)}: {e}")
            stats["errors"].append(f"actor call {marketplace}: {e}")
            continue
        for it in items:
            asin = it.get("asin")
            if not asin:
                continue
            out[asin] = {
                "found": bool(it.get("found")) and it.get("rating") is not None,
                "rating": it.get("rating"),
                "ratingsCount": it.get("ratingsCount"),
            }
    return out


# ---------------------------------------------------------------- ASIN_Master fallback
def load_asin_master(svc):
    resp = svc.spreadsheets().values().get(
        spreadsheetId=ASIN_MASTER_ID, range=f"'{ASIN_MASTER_TAB}'!A1:J", valueRenderOption="UNFORMATTED_VALUE"
    ).execute(num_retries=8)
    rows = resp.get("values", [])
    if not rows:
        return {}, {}
    header = rows[0]
    asin_i = header.index("ASIN") if "ASIN" in header else None
    sku_i = header.index("SKU") if "SKU" in header else None
    asin_to_sku, sku_to_asin = {}, {}
    if asin_i is None or sku_i is None:
        return asin_to_sku, sku_to_asin
    for row in rows[1:]:
        a = row[asin_i].strip() if len(row) > asin_i and row[asin_i] else ""
        s = row[sku_i].strip() if len(row) > sku_i and row[sku_i] else ""
        if a and s:
            asin_to_sku[a] = s
            sku_to_asin[s] = a
    return asin_to_sku, sku_to_asin


# ---------------------------------------------------------------- Group A
def process_group_a_tab(svc, product, sid, tab, dry_run, pilot_mode, asin_to_sku, sku_to_asin, stats, needed_by_market):
    header_row, cols = find_header(svc, sid, tab)
    if header_row is None or "asin" not in cols:
        log(f"[{product}/{tab}] SKIP — no ASIN header found in first 10 rows")
        stats["skipped_tabs"].append(f"{product}/{tab} (no header)")
        return None

    asin_col = cols.get("asin")
    sku_col = cols.get("sku")
    country_col = cols.get("country")
    score_col = cols.get("score")
    global_col = cols.get("global")
    if not score_col or not global_col:
        log(f"[{product}/{tab}] SKIP — no Score/Global Ratings column found (header row {header_row})")
        stats["skipped_tabs"].append(f"{product}/{tab} (no Score/Global col)")
        return None

    last_row = last_data_row(svc, sid, tab, asin_col, header_row)
    if last_row <= header_row:
        log(f"[{product}/{tab}] no data rows, skipping")
        return None

    migrated = migrate_arrayformula_if_present(svc, sid, tab, header_row, score_col, global_col, dry_run)
    if migrated:
        stats["migrated_tabs"].append(f"{product}/{tab}")

    if pilot_mode and not dry_run:
        bad = verify_no_errors(svc, sid, tab, score_col, global_col, header_row, min(last_row, header_row + 30))
        if bad:
            log(f"[{product}/{tab}] PILOT CHECK FAILED — errors found: {bad}")
            stats["pilot_failed"] = True
            return None
        log(f"[{product}/{tab}] pilot check OK (no #REF!/#ERROR! in sample)")

    # Read ASIN/SKU/country columns for all data rows
    cols_needed = sorted(set(filter(None, [asin_col, sku_col, country_col])))
    read_range = f"'{tab}'!{cols_needed[0]}{header_row+1}:{cols_needed[-1]}{last_row}"
    resp = svc.spreadsheets().values().get(spreadsheetId=sid, range=read_range, valueRenderOption="UNFORMATTED_VALUE").execute(num_retries=8)
    rows = resp.get("values", [])

    def rel(letter):
        return col_idx(letter) - col_idx(cols_needed[0])

    row_info = []  # (row_number, asin, market_code)
    for i, row in enumerate(rows):
        row_num = header_row + 1 + i
        asin = str(row[rel(asin_col)]).strip() if len(row) > rel(asin_col) and row[rel(asin_col)] else ""
        sku = str(row[rel(sku_col)]).strip() if sku_col and len(row) > rel(sku_col) and row[rel(sku_col)] else ""
        country_raw = str(row[rel(country_col)]).strip().upper() if country_col and len(row) > rel(country_col) and row[rel(country_col)] else ""

        if not asin and sku and sku_to_asin.get(sku):
            asin = sku_to_asin[sku]
            stats["asin_master_resolved"] += 1
        if not sku and asin and asin_to_sku.get(asin):
            stats["asin_master_resolved"] += 1  # sku fill not written back here, just counted

        if not asin:
            stats["unresolved_rows"] += 1
            row_info.append((row_num, None, None))
            continue

        market = MARKET_CODE.get(country_raw)
        if not market:
            stats["unresolved_rows"] += 1
            row_info.append((row_num, asin, None))
            continue

        needed_by_market.setdefault(market, set()).add(asin)
        row_info.append((row_num, asin, market))

    return {"header_row": header_row, "last_row": last_row, "score_col": score_col,
            "global_col": global_col, "row_info": row_info}


def write_group_a_results(svc, sid, tab, info, results, dry_run, stats):
    score_col, global_col = info["score_col"], info["global_col"]
    values_rows = []
    for row_num, asin, market in info["row_info"]:
        rating, count = "", ""
        if asin and market:
            r = results.get(market, {}).get(asin)
            if r and r.get("found"):
                rating, count = r.get("rating", ""), r.get("ratingsCount", "")
        values_rows.append([rating, count])
    if not values_rows:
        return
    first_row = info["row_info"][0][0]
    last_row = info["row_info"][-1][0]
    stats["cells_written"] += len(values_rows) * 2
    if dry_run:
        return
    svc.spreadsheets().values().update(
        spreadsheetId=sid, range=f"'{tab}'!{score_col}{first_row}:{global_col}{last_row}",
        valueInputOption="RAW", body={"values": values_rows}
    ).execute(num_retries=8)


# ---------------------------------------------------------------- Group B
def process_group_b_tab(svc, label, sid, tab, primary, fallbacks, dry_run, stats, needed_by_market):
    header_row, cols = find_header(svc, sid, tab, want=("asin", "sku", "score"))
    if header_row is None or "asin" not in cols or "score" not in cols:
        log(f"[B:{label}] SKIP — header not found")
        stats["skipped_tabs"].append(f"B:{label} (no header)")
        return None
    asin_col, rating_col = cols["asin"], cols["score"]
    last_row = last_data_row(svc, sid, tab, asin_col, header_row)
    if last_row <= header_row:
        return None
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{tab}'!{asin_col}{header_row+1}:{asin_col}{last_row}",
        valueRenderOption="UNFORMATTED_VALUE"
    ).execute(num_retries=8)
    asins = [str(r[0]).strip() if r else "" for r in resp.get("values", [])]
    for a in asins:
        if a:
            needed_by_market.setdefault(primary, set()).add(a)
    return {"header_row": header_row, "last_row": last_row, "asin_col": asin_col,
            "rating_col": rating_col, "asins": asins, "primary": primary, "fallbacks": fallbacks}


def resolve_group_b_value(asin, primary, fallbacks, results):
    r = results.get(primary, {}).get(asin)
    if r and r.get("found"):
        return r.get("rating"), primary
    for fb in fallbacks:
        r = results.get(fb, {}).get(asin)
        if r and r.get("found"):
            return r.get("rating"), fb
    return None, None


def write_group_b_results(svc, sid, tab, info, results, dry_run, stats):
    rating_col = info["rating_col"]
    first_row = info["header_row"] + 1
    last_row = info["last_row"]
    cell_data = []
    fallback_used = 0
    for asin in info["asins"]:
        if not asin:
            cell_data.append({"userEnteredValue": {}, "note": None})
            continue
        rating, used_market = resolve_group_b_value(asin, info["primary"], info["fallbacks"], results)
        if rating is None:
            cell_data.append({"userEnteredValue": {}, "note": None})
            continue
        cell = {"userEnteredValue": {"numberValue": float(rating)}}
        if used_market != info["primary"]:
            cell["note"] = f"Rating from {MARKET_LABEL.get(used_market, used_market)} — {MARKET_LABEL.get(info['primary'], info['primary'])} not found"
            fallback_used += 1
        else:
            cell["note"] = None
        cell_data.append(cell)

    stats["cells_written"] += len(cell_data)
    stats["group_b_fallback_used"] += fallback_used
    if dry_run:
        return

    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute(num_retries=8)
    sheet_id = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == tab)
    col0 = col_idx(rating_col)
    req = {
        "requests": [{
            "updateCells": {
                "rows": [{"values": [c]} for c in cell_data],
                "fields": "userEnteredValue,note",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": first_row - 1, "endRowIndex": last_row,
                    "startColumnIndex": col0, "endColumnIndex": col0 + 1,
                }
            }
        }]
    }
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body=req).execute(num_retries=8)


# ---------------------------------------------------------------- Chat report
def post_chat_summary(stats, run_label, pilot_mode, dry_run):
    lines = [f"<b>ASIN Rating Daily — {run_label}</b>", f"{datetime.now():%Y-%m-%d %H:%M KST}"]
    if pilot_mode:
        lines.append("Mode: PILOT (GlxZ8/DE only)")
    if dry_run:
        lines.append("Mode: DRY RUN (nothing written)")
    lines.append(f"Actor calls: {stats['actor_calls']} ({stats['actor_calls_by_market']})")
    lines.append(f"Cells written: {stats['cells_written']}")
    lines.append(f"Tabs migrated off ARRAYFORMULA: {len(stats['migrated_tabs'])}" +
                 (f" — {stats['migrated_tabs']}" if stats['migrated_tabs'] else ""))
    lines.append(f"Tabs skipped: {len(stats['skipped_tabs'])}" + (f" — {stats['skipped_tabs']}" if stats['skipped_tabs'] else ""))
    lines.append(f"ASIN_Master fallback resolved: {stats['asin_master_resolved']}, still unresolved: {stats['unresolved_rows']}")
    lines.append(f"Group B fallback-market substitutions: {stats['group_b_fallback_used']}")
    if stats["errors"]:
        lines.append(f"Errors: {stats['errors']}")
    if stats.get("pilot_failed"):
        lines.append("⚠️ PILOT CHECK FAILED — rollout halted, see log")
    text = "\n".join(lines)
    card = {"cardsV2": [{"cardId": f"asin-rating-daily-{int(time.time())}",
                          "card": {"header": {"title": "ASIN Rating Daily"},
                                   "sections": [{"widgets": [{"textParagraph": {"text": text.replace(chr(10), '<br>')}}]}]}}]}
    try:
        req = urllib.request.Request(PRIVATE_WEBHOOK, json.dumps(card).encode(),
                                      headers={"Content-Type": "application/json; charset=UTF-8"})
        urllib.request.urlopen(req, timeout=15)
        log("Posted summary to private Chat room")
    except Exception as e:
        log(f"Chat post failed: {e}")


# ---------------------------------------------------------------- Main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--pilot-only", action="store_true")
    ap.add_argument("--run-label", default="run")
    args = ap.parse_args()
    dry_run = not args.commit

    stats = {"actor_calls": 0, "actor_calls_by_market": {}, "cells_written": 0,
             "migrated_tabs": [], "skipped_tabs": [], "asin_master_resolved": 0,
             "unresolved_rows": 0, "group_b_fallback_used": 0, "errors": [], "pilot_failed": False}

    log(f"=== rating_fetch_daily.py start — label={args.run_label} pilot={args.pilot_only} commit={args.commit} ===")
    svc = get_sheets_service()

    asin_to_sku, sku_to_asin = load_asin_master(svc)
    log(f"ASIN_Master loaded: {len(asin_to_sku)} ASIN->SKU mappings")

    needed_by_market = {}
    tab_infos = []  # (kind, product/label, sid, tab, info)

    if args.pilot_only:
        books = {"GlxZ8": (GROUP_A_BOOKS["GlxZ8"][0], ["DE"])}
    else:
        books = {k: v for k, v in GROUP_A_BOOKS.items()}

    for product, (sid, tabs) in books.items():
        for tab in tabs:
            pilot_mode = args.pilot_only
            info = process_group_a_tab(svc, product, sid, tab, dry_run, pilot_mode,
                                        asin_to_sku, sku_to_asin, stats, needed_by_market)
            if info:
                tab_infos.append(("A", product, sid, tab, info))
            if stats["pilot_failed"]:
                log("Halting further processing due to pilot failure.")
                post_chat_summary(stats, args.run_label, args.pilot_only, dry_run)
                return 1

    if not args.pilot_only:
        for label, sid, tab, primary, fallbacks in GROUP_B_TABS:
            info = process_group_b_tab(svc, label, sid, tab, primary, fallbacks, dry_run, stats, needed_by_market)
            if info:
                tab_infos.append(("B", label, sid, tab, info))

    # Fetch all primary-market needs
    results = {}
    for market, asins in needed_by_market.items():
        log(f"Fetching marketplace={market}: {len(asins)} ASINs")
        results[market] = fetch_market(list(asins), market, stats)

    # Group B fallback cascade: find not-found ASINs on primary, fetch fallbacks
    if not args.pilot_only:
        for kind, label, sid, tab, info in tab_infos:
            if kind != "B":
                continue
            primary = info["primary"]
            not_found = [a for a in info["asins"] if a and not (results.get(primary, {}).get(a) or {}).get("found")]
            for fb in info["fallbacks"]:
                pending = [a for a in not_found if not (results.get(fb, {}).get(a) or {}).get("found")]
                if not pending:
                    continue
                log(f"[B:{label}] fallback fetch marketplace={fb}: {len(pending)} ASINs")
                fetched = fetch_market(pending, fb, stats)
                results.setdefault(fb, {}).update(fetched)

    # Write results
    for kind, label, sid, tab, info in tab_infos:
        if kind == "A":
            write_group_a_results(svc, sid, tab, info, results, dry_run, stats)
        else:
            write_group_b_results(svc, sid, tab, info, results, dry_run, stats)

    log(f"=== done. stats={json.dumps(stats, ensure_ascii=False)} ===")
    post_chat_summary(stats, args.run_label, args.pilot_only, dry_run)
    return 0


if __name__ == "__main__":
    sys.exit(main())
