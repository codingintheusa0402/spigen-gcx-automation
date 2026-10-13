#!/usr/bin/env python3
"""Fetch the latest 500 Spigen Amazon reviews per marketplace (US/DE/IT/FR/ES/UK/JP/IN)
from Caspi via the registered data-api query, and upload them as a new
`CaspiLM_{yymmdd}` tab (KST run date) on the source review spreadsheet.

Columns match the SC scraper output exactly (ASIN … Ratings Count), so the tab
feeds straight into sc-review-propagate.

Usage:
  python3 fetch_caspi_reviews.py              # fetch + CSV + upload
  python3 fetch_caspi_reviews.py --no-upload  # fetch + CSV only

Caspi API key lookup (first hit wins; the key is never printed):
  1. env CASPI_API_KEY
  2. ~/.config/caspi_review_fetch/api_key   (plain text, one line)
  3. ~/.config/zendesk_inquiry_sync/credentials.json → "caspi_api_key"

(Source restored 2026-10-13 from the surviving .pyc after the .py went missing;
 same constants, same paging logic as Zendesk_Inquiry_Sync/sync.py.)
"""
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.request
from collections import Counter
from datetime import datetime, timedelta, timezone

ENDPOINT = "https://caspilm.spigen.com/api/data-api/run"
QUERY_ID = "pq_6df881b40267ebc62e"   # "Spigen Amazon 리뷰 최신 500건/마켓 (… SC 스크래퍼 컬럼 순서, 이미지 URL 포함)"
SPREADSHEET_ID = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw"
HEADER = ["ASIN", "Created 날짜", "사진 유무", "Reviewer", "Review Ratings", "Review Title",
          "본문", "국가", "Review Link", "Image URL", "Review ID", "Order ID",
          "Product Rating", "Ratings Count"]
PAGE = 2000
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
KST = timezone(timedelta(hours=9))


def api_key():
    k = os.environ.get("CASPI_API_KEY", "").strip()
    if k:
        return k
    p = os.path.expanduser("~/.config/caspi_review_fetch/api_key")
    if os.path.exists(p):
        with open(p) as f:
            k = f.read().strip()
        if k:
            return k
    p = os.path.expanduser("~/.config/zendesk_inquiry_sync/credentials.json")
    if os.path.exists(p):
        with open(p) as f:
            k = (json.load(f).get("caspi_api_key") or "").strip()
        if k:
            return k
    sys.exit("Caspi API key not found — set CASPI_API_KEY or write it to ~/.config/caspi_review_fetch/api_key")


def fetch_all(key):
    rows, offset, columns = [], 0, None
    while True:
        body = json.dumps({"queryId": QUERY_ID, "params": {}, "limit": PAGE, "offset": offset}).encode()
        for attempt in range(5):
            req = urllib.request.Request(ENDPOINT, data=body, method="POST",
                                         headers={"x-api-key": key, "Content-Type": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=300) as f:
                    data = json.load(f)
                break
            except (urllib.error.URLError, TimeoutError) as e:
                if isinstance(e, urllib.error.HTTPError) and e.code < 500 and e.code != 429:
                    sys.exit(f"Caspi API error {e.code}: {e.read()[:500]!r}")
                print(f"Caspi retry {attempt + 1}: {e}", flush=True)
                time.sleep(15 * (attempt + 1))
        else:
            raise RuntimeError("Caspi data-api failed after retries")
        got = data.get("rows") or []
        columns = columns or data.get("columns")
        rows.extend(got)
        nxt = (data.get("paging") or {}).get("nextOffset")
        if nxt is None and data.get("truncated") and got:
            # size-capped responses flag truncated=true with nextOffset=null → page by rows received
            nxt = offset + len(got)
        print(f"  page offset={offset}: +{len(got)} rows (truncated={data.get('truncated')}, next={nxt})", flush=True)
        if nxt is None or not got:
            return rows, columns
        offset = nxt


def to_list(row, columns):
    if isinstance(row, dict):
        if all(h in row for h in HEADER):
            return [row.get(h) for h in HEADER]
        cols = columns or list(row.keys())
        return [row.get(c) for c in cols[:len(HEADER)]]
    return list(row)[:len(HEADER)]


def norm(v, i):
    if v is None:
        return ""
    if i in (4, 12, 13):  # Review Ratings / Product Rating / Ratings Count → numbers
        try:
            f = float(v)
            return int(f) if f.is_integer() else f
        except (TypeError, ValueError):
            return v
    return str(v)


def upload(table, run_date):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    tpath = os.path.expanduser("~/.config/gws_shim/token.json")
    with open(tpath) as f:
        info = json.load(f)
    creds = Credentials.from_authorized_user_info(info)
    if not creds.valid:
        creds.refresh(Request())
        info["token"] = creds.token
        with open(tpath, "w") as f:
            json.dump(info, f)
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    title = f"CaspiLM_{run_date}"
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields="sheets.properties").execute()
    if any(s["properties"]["title"] == title for s in meta["sheets"]):
        sys.exit(f"tab {title} already exists — not overwriting")
    res = svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": [{"addSheet": {
        "properties": {"title": title, "gridProperties": {"rowCount": len(table) + 1, "columnCount": len(HEADER)}}}}]}).execute()
    gid = res["replies"][0]["addSheet"]["properties"]["sheetId"]
    svc.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1",
                                       valueInputOption="RAW", body={"values": [HEADER] + table}).execute()
    print(f"uploaded {len(table)} rows → {title} (gid {gid})")
    return title, gid


def main():
    run_date = datetime.now(KST).strftime("%y%m%d")
    print(f"Caspi review fetch — run date {run_date}")
    rows, columns = fetch_all(api_key())
    seen, table = set(), []
    rid_i = HEADER.index("Review ID")
    for r in rows:
        lst = [norm(v, i) for i, v in enumerate(to_list(r, columns))]
        rid = lst[rid_i]
        if not rid or rid in seen:
            continue
        seen.add(rid)
        table.append(lst)
    by_mkt = Counter(r[HEADER.index("국가")] for r in table)
    photos = sum(1 for r in table if r[2] == "Y")
    urls = sum(1 for r in table if r[9])
    print(f"{len(rows)} raw → {len(table)} unique Review IDs · per 국가: {dict(sorted(by_mkt.items()))}")
    print(f"사진 유무=Y: {photos} · with Image URL: {urls}")
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, f"CaspiLM_{run_date}.csv")
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        w.writerows(table)
    print(f"csv: {out}")
    if "--no-upload" not in sys.argv:
        upload(table, run_date)


if __name__ == "__main__":
    main()
