#!/usr/bin/env python3
"""
Reads/marks the "Feedback" tab of the Ticket Reporter TicketQueue sheet — the /revision
channel from the Chat app (see TicketReporterCard/Code.gs handleRevisionFeedback_).

Usage:
  python3 check_feedback.py list              # print unapplied rows as JSON (row, ts, email, feedbackText)
  python3 check_feedback.py mark-applied <row_number>   # stamp that row's appliedAt (1-indexed data row, not header)

The ticket-reporter Claude session should run `list` at the start of every monitor tick;
for each unapplied row, read feedbackText, update SKILL.md's writing rules accordingly, then
call `mark-applied <row>` so the same feedback isn't re-read next tick.
"""
import datetime
import json
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

QUEUE_SHEET_ID = "1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M"
FEEDBACK_TAB = "Feedback"
GWS_SHIM_TOKEN = "/Users/kevinkim/.config/gws_shim/token.json"


def _sheets_client():
    with open(GWS_SHIM_TOKEN, encoding="utf-8") as f:
        info = json.load(f)
    creds = Credentials(
        token=info.get("token"),
        refresh_token=info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        scopes=info.get("scopes"),
    )
    creds.refresh(Request())
    info["token"] = creds.token
    with open(GWS_SHIM_TOKEN, "w", encoding="utf-8") as f:
        json.dump(info, f)
    return build("sheets", "v4", credentials=creds)


def list_unapplied():
    sheets = _sheets_client()
    try:
        result = sheets.spreadsheets().values().get(
            spreadsheetId=QUEUE_SHEET_ID, range=f"{FEEDBACK_TAB}!A2:D"
        ).execute()
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": str(e), "rows": []}))
        return
    values = result.get("values", [])
    rows = []
    for i, row in enumerate(values):
        row = row + [""] * (4 - len(row))
        ts, email, feedback_text, applied_at = row[:4]
        if not applied_at.strip():
            rows.append({
                "row": i + 2,  # 1-indexed sheet row (header is row 1)
                "ts": ts,
                "email": email,
                "feedbackText": feedback_text,
            })
    print(json.dumps({"rows": rows}, ensure_ascii=False, indent=1))


def mark_applied(row_number: int):
    sheets = _sheets_client()
    sheets.spreadsheets().values().update(
        spreadsheetId=QUEUE_SHEET_ID,
        range=f"{FEEDBACK_TAB}!D{row_number}",
        valueInputOption="RAW",
        body={"values": [[datetime.datetime.now().isoformat(timespec="seconds")]]},
    ).execute()
    print(f"Marked row {row_number} applied.")


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    cmd = sys.argv[1]
    if cmd == "list":
        list_unapplied()
        return 0
    if cmd == "mark-applied":
        if len(sys.argv) < 3:
            print("ERROR: mark-applied requires a row number", file=sys.stderr)
            return 2
        mark_applied(int(sys.argv[2]))
        return 0
    print(f"ERROR: unknown command {cmd!r}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
