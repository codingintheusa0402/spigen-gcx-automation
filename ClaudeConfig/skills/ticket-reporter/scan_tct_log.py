#!/usr/bin/env python3
"""Scan the Lazada/Shopee TCT log sheet for rows whose Status (col A, trimmed)
is exactly "Esc T2". Always reads the tab's real gridProperties.rowCount first
-- never a hardcoded/remembered row cap. See feedback_ticket_reporter_tct_log_full_range_scan
memory: this exact hardcoded-range bug has recurred twice and caused silent
"0 found" false negatives for hours/days at a time.

Usage:
    python3 scan_tct_log.py                  # scan both tabs, print JSON
    python3 scan_tct_log.py --tab "Lazada log"
"""
import argparse
import json
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

TOKEN_PATH = "/Users/kevinkim/.config/gws_shim/token.json"
SHEET_ID = "1HZ14uqTVeP7bGYZDu9v9Ve2C1xNY_m6dcSv-KMCoAKc"
TABS = ["Lazada log", "Shopee log"]


def _sheets_client():
    with open(TOKEN_PATH) as f:
        tok = json.load(f)
    creds = Credentials.from_authorized_user_info(tok)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("sheets", "v4", credentials=creds)


def scan_tab(svc, tab):
    # Always read the tab's real row count -- never hardcode/remember a cap.
    meta = svc.spreadsheets().get(
        spreadsheetId=SHEET_ID,
        fields="sheets.properties",
    ).execute()
    row_count = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab:
            row_count = s["properties"]["gridProperties"]["rowCount"]
            break
    if row_count is None:
        raise RuntimeError(f"Tab not found: {tab}")

    rng = f"'{tab}'!A5:C{row_count}"
    resp = svc.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=rng).execute()
    rows = resp.get("values", [])

    esc_t2 = []
    for i, row in enumerate(rows):
        status = row[0].strip() if len(row) > 0 else ""
        ticket_id = row[1].strip() if len(row) > 1 else ""
        order_id = row[2].strip() if len(row) > 2 else ""
        if status == "Esc T2":
            esc_t2.append({"row": i + 5, "ticket_id": ticket_id, "order_id": order_id})

    return {"tab": tab, "sheet_row_count": row_count, "rows_scanned": len(rows), "esc_t2": esc_t2}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tab", choices=TABS, help="scan only this tab (default: both)")
    args = ap.parse_args()

    svc = _sheets_client()
    tabs = [args.tab] if args.tab else TABS
    results = [scan_tab(svc, tab) for tab in tabs]
    print(json.dumps(results, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    sys.exit(main())
