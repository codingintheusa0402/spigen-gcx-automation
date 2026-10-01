#!/usr/bin/env python3
"""
Fills Z col (판매량, EU+UK Amazon) on the '1-5점' tab of the iPhone18 /
Pixel11 / GlxZ8 review-monitoring spreadsheets with CUMULATIVE units sold
per base SKU from each product's start_date through Caspi's latest order day.

Runs unattended via launchd (see com.spigen.gcx.caspi-sales-backfill.plist),
fired every 30 min on weekdays. Each tick is a no-op unless it's within the
catch-up window (weekday, 08:00-23:59 KST) AND that product hasn't already
succeeded today -- so if the Mac was off/asleep during the normal 8-10AM
window, the first tick after it wakes catches up automatically.

Data source: Caspi registered query pq_01b6fe14503ee3e206
("판매량_EU_cumulative_since_start_date"), called via the headless data-api
endpoint (no Claude session needed) with a Caspi API key. It sums order
quantity from S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL
for EU+UK Amazon marketplaces (DE/FR/IT/ES/UK/NL/SE/PL/BE/IE), excluding
Cancelled, deduped per order+SKU, purchase-date >= start_date.

Until 2026-10-01 this used RESTOCK_INVENTORY_RECOMMENDATIONS_REPORT's
"Units Sold Last 30 Days" (latest minus a baseline snapshot). That is a
rolling 30-day window, not a cumulative count, and the baseline never
matched (FILE_DATE comes back as a day number), so Z showed last-30-day
FBA EU5 sales only.
"""
import json
import os
import sys
import datetime
import urllib.request
from zoneinfo import ZoneInfo

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request as GoogleAuthRequest
from googleapiclient.discovery import build

KST = ZoneInfo("Asia/Seoul")
SECRETS_PATH = os.path.expanduser("~/.config/caspi_sales_backfill/secrets.json")
STATE_PATH = os.path.expanduser("~/.config/caspi_sales_backfill/state.json")
GWS_TOKEN_PATH = os.path.expanduser("~/.config/gws_shim/token.json")
CASPI_ENDPOINT = "https://caspilm.spigen.com/api/data-api/run"
SOURCE_TABLE_NOTE = "S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL (EU+UK orders, excl. Cancelled)"

SHEET_TAB = "1-5점"
SKU_COL_IDX = 16  # Q
Z_COL_IDX = 25    # Z

PRODUCTS = [
    {
        "key": "iphone18",
        "label": "iPhone 18 Series",
        "ssid": "1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU",
        "sheet_id": 957652957,
        "start_date": "2026-09-18",
    },
    {
        "key": "pixel11",
        "label": "Pixel 11 Series",
        "ssid": "12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI",
        "sheet_id": 957652957,
        "start_date": "2026-08-18",
    },
    {
        "key": "glxz8",
        "label": "Galaxy Z Fold8/Flip8/Fold8 Ultra Series",
        "ssid": "19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4",
        "sheet_id": 957652957,
        "start_date": "2026-07-27",
    },
]


def log(msg):
    ts = datetime.datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def load_json(path, default):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return default


def save_state(state):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    with open(STATE_PATH, "w") as f:
        json.dump(state, f, indent=2)


def within_catchup_window(now):
    if now.weekday() >= 5:  # Sat=5, Sun=6
        return False
    return now.hour >= 8  # up to 23:59


def caspi_query(start_date, api_key, query_id):
    rows = []
    offset = 0
    while True:
        body = json.dumps({
            "queryId": query_id,
            "params": {"start_date": start_date},
            "limit": 1000,
            "offset": offset,
        }).encode()
        req = urllib.request.Request(
            CASPI_ENDPOINT, data=body,
            headers={"x-api-key": api_key, "Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read())
        rows.extend(data["rows"])
        nxt = (data.get("paging") or {}).get("nextOffset")
        if nxt is None:
            break
        offset = nxt
    return rows


def compute_units(rows):
    """{base_sku: cumulative units}. SKUs with no orders are absent → Z left blank."""
    return {r["BASE_SKU"]: int(float(r["UNITS"] or 0)) for r in rows}


def get_sheets_client():
    with open(GWS_TOKEN_PATH) as f:
        info = json.load(f)
    creds = Credentials.from_authorized_user_info(info)
    creds.refresh(GoogleAuthRequest())
    return build("sheets", "v4", credentials=creds)


def process_product(sheets, product, deltas, today_str):
    ssid = product["ssid"]
    meta = sheets.spreadsheets().get(spreadsheetId=ssid).execute()
    row_count = None
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] == product["sheet_id"]:
            row_count = s["properties"]["gridProperties"]["rowCount"]
            break
    if row_count is None:
        raise RuntimeError(f"sheet_id {product['sheet_id']} not found in {ssid}")

    res = sheets.spreadsheets().values().get(
        spreadsheetId=ssid, range=f"'{SHEET_TAB}'!A2:Z{row_count}"
    ).execute()
    data_rows = res.get("values", [])

    out_rows = []
    filled = 0
    for r in data_rows:
        sku = r[SKU_COL_IDX] if len(r) > SKU_COL_IDX else ""
        if sku and sku in deltas:
            out_rows.append([deltas[sku]])
            filled += 1
        else:
            out_rows.append([""])

    if not out_rows:
        log(f"  {product['key']}: no data rows found, skipping write")
        return

    end_row = 1 + len(out_rows)
    sheets.spreadsheets().values().update(
        spreadsheetId=ssid,
        range=f"'{SHEET_TAB}'!Z2:Z{end_row}",
        valueInputOption="RAW",
        body={"values": out_rows},
    ).execute()

    note = f"last updated {today_str} from {SOURCE_TABLE_NOTE} (accumulated from {product['start_date']})"
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=ssid,
        body={"requests": [{
            "updateCells": {
                "range": {
                    "sheetId": product["sheet_id"],
                    "startRowIndex": 0, "endRowIndex": 1,
                    "startColumnIndex": Z_COL_IDX, "endColumnIndex": Z_COL_IDX + 1,
                },
                "rows": [{"values": [{"note": note}]}],
                "fields": "note",
            }
        }]},
    ).execute()

    log(f"  {product['key']}: wrote {filled}/{len(out_rows)} rows")


def main():
    now = datetime.datetime.now(KST)
    today_str = now.strftime("%Y-%m-%d")

    if not within_catchup_window(now):
        return  # silent no-op outside weekday 08:00-23:59 KST

    state = load_json(STATE_PATH, {})
    pending = [p for p in PRODUCTS if state.get(p["key"], {}).get("last_success_date") != today_str]
    if not pending:
        return  # already ran successfully for all products today

    secrets = load_json(SECRETS_PATH, {})
    api_key = secrets.get("caspi_api_key")
    query_id = secrets.get("caspi_query_id_cumulative")
    if not api_key or not query_id:
        log("ERROR: missing Caspi API key/queryId in secrets.json")
        return

    log(f"Running for {len(pending)} pending product(s): {[p['key'] for p in pending]}")
    sheets = get_sheets_client()

    for product in pending:
        try:
            rows = caspi_query(product["start_date"], api_key, query_id)
            deltas = compute_units(rows)
            process_product(sheets, product, deltas, today_str)
            state.setdefault(product["key"], {})["last_success_date"] = today_str
            state[product["key"]]["last_run_at"] = now.isoformat()
            save_state(state)  # persist after each product so partial progress survives a crash
        except Exception as e:
            log(f"  {product['key']}: FAILED - {e!r}")

    log("Done.")


if __name__ == "__main__":
    main()
