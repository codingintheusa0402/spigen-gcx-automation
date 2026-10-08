#!/usr/bin/env python3
"""
Reads/marks the "BtwQueue" tab of the Ticket Reporter TicketQueue sheet — the /btw
side-channel Q&A from the Chat app (see TicketReporterCard/Code.gs handleBtwQuestion_).

/btw deliberately does NOT call an LLM API from Apps Script (사용자 지시 2026-10-07 — no
Anthropic API key in Script Properties; the question should be answered by an actual
Claude Code session instead). handleBtwQuestion_ just appends the question to this tab and
acks "접수됨" back to Chat. This script is how the ticket-reporter monitor session picks the
question back up on its next tick, answers it using its own reasoning/tools, and posts the
answer into the original thread (see Monitor mode section of this skill's SKILL.md).

Usage:
  python3 check_btw_queue.py list                       # print unanswered rows as JSON
  python3 check_btw_queue.py mark-answered <row_number>  # stamp that row's answeredAt (1-indexed data row, not header)

The ticket-reporter Claude session should run `list` at the start of every monitor tick;
for each unanswered row, answer `question` (plain text, matching the question's language),
post it with `send.py --file <answer.txt> --thread "<threadId>"` (no --ticket-id — this
never touches Zendesk), then call `mark-answered <row>` so the same question isn't
re-answered next tick. A row with an empty threadId can't be replied to in-thread — post a
new message instead by omitting --thread.
"""
import datetime
import json
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

QUEUE_SHEET_ID = "1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M"
BTW_TAB = "BtwQueue"
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


def list_unanswered():
    sheets = _sheets_client()
    try:
        result = sheets.spreadsheets().values().get(
            spreadsheetId=QUEUE_SHEET_ID, range=f"{BTW_TAB}!A2:E"
        ).execute()
    except Exception as e:  # noqa: BLE001
        # Most common cause: the tab doesn't exist yet because /btw has never been used —
        # Code.gs creates it lazily on first use. Empty list, not an error, in that case.
        print(json.dumps({"error": str(e), "rows": []}))
        return
    values = result.get("values", [])
    rows = []
    for i, row in enumerate(values):
        row = row + [""] * (5 - len(row))
        ts, question, thread_id, email, answered_at = row[:5]
        if not answered_at.strip():
            rows.append({
                "row": i + 2,  # 1-indexed sheet row (header is row 1)
                "ts": ts,
                "question": question,
                "threadId": thread_id,
                "email": email,
            })
    print(json.dumps({"rows": rows}, ensure_ascii=False, indent=1))


def mark_answered(row_number: int):
    sheets = _sheets_client()
    sheets.spreadsheets().values().update(
        spreadsheetId=QUEUE_SHEET_ID,
        range=f"{BTW_TAB}!E{row_number}",
        valueInputOption="RAW",
        body={"values": [[datetime.datetime.now().isoformat(timespec="seconds")]]},
    ).execute()
    print(f"Marked row {row_number} answered.")


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    cmd = sys.argv[1]
    if cmd == "list":
        list_unanswered()
        return 0
    if cmd == "mark-answered":
        if len(sys.argv) < 3:
            print("ERROR: mark-answered requires a row number", file=sys.stderr)
            return 2
        mark_answered(int(sys.argv[2]))
        return 0
    print(f"ERROR: unknown command {cmd!r}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
