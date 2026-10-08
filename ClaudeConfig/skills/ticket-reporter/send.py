#!/usr/bin/env python3
"""
Ticket Reporter — Google Chat webhook sender.

Two modes:

  1. Card mode (default, use --html): posts a cardsV2 card whose body is the
     report rendered with Google-Chat-supported HTML (<b>, <i>, <font color>,
     <a href>, <br>). This is the mode that supports RED / BLUE keyword color.
     Attachments from the ticket page are added as image widgets (--image) and
     as a link list (--attach "name|url").

  2. Legacy text mode (--file): posts {"text": ...} plain text. No color.

Usage:
  python3 send.py --html body.html \
      --image https://…/photo1.jpg --image https://…/photo2.png \
      --attach "clip.mp4|https://…/clip.mp4" \
      [--dry-run] [--webhook '<url>']

  python3 send.py --file report.txt            # legacy plain text
  cat report.txt | python3 send.py             # legacy plain text (stdin)

Notes
- Google Chat textParagraph caps at ~4096 chars; the body is split on <br>
  into multiple textParagraph widgets when longer.
- Zendesk attachment URLs usually render in an image widget (the token in the
  URL is the auth); the link list is always included as a fallback.
"""
import argparse
import datetime
import json, os
import re
import sys
import urllib.request
import urllib.error

DEFAULT_WEBHOOK = json.load(open(os.path.expanduser("~/.config/gcx_webhooks.json")))["ticket_reporter_room"]  # secret — kept out of git  # real chatroom (user, 2026-09-14). Old test room: spaces/AAQAc9NQmJQ

# TicketQueue sheet used by the Ticket Reporter Chat app (Code.gs lookupTicketByThread_) to
# resolve a plain thread reply back to a ticket id, without needing chat.bot API read access.
QUEUE_SHEET_ID = "1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M"
QUEUE_TAB = "TicketQueue"
GWS_SHIM_TOKEN = "/Users/kevinkim/.config/gws_shim/token.json"

RED = "#D93025"    # CS 미제공 / 거절 관련 강조
BLUE = "#1A73E8"   # CS 제공 관련 강조
MAX_TP = 3800      # keep each textParagraph safely under the 4096 cap


def _normalize(html: str) -> str:
    """Collapse file newlines to a single space so ONLY <br> controls line
    breaks. Google Chat's textParagraph renders a literal "\n" as a break, so
    `<br>\n` in the source would double-space every line."""
    html = re.sub(r"\s*[\r\n]+\s*", " ", html).strip()
    html = re.sub(r"\s*(<br\s*/?>)\s*", r"\1", html)  # tidy spaces around <br>
    return html


def _chunk_html(html: str):
    """Split an HTML body on <br> into <=MAX_TP-char pieces."""
    parts = re.split(r"(<br\s*/?>)", html)
    out, buf = [], ""
    for p in parts:
        if len(buf) + len(p) > MAX_TP and buf:
            out.append(buf)
            buf = ""
        buf += p
    if buf:
        out.append(buf)
    return out or [html]


def _plain(html: str) -> str:
    t = re.sub(r"<br\s*/?>", "\n", html)
    t = re.sub(r"<[^>]+>", "", t)
    return re.sub(r"[ \t]+\n", "\n", t).strip()


def build_card(html: str, images, attaches):
    widgets = [{"textParagraph": {"text": c}} for c in _chunk_html(html)]
    sections = [{"widgets": widgets}]

    extra = []
    if len(images) == 1:
        extra.append({"image": {"imageUrl": images[0], "altText": "ticket attachment"}})
    elif len(images) >= 2:
        # Grid widget instead of one `image` widget per row — keeps the card short when
        # many attachments exist (user rule 2026-09-16). Always 2 columns (user feedback
        # 2026-09-30, /revision row 23: 3-column thumbnails rendered too small — capping at
        # 2 columns regardless of image count keeps each thumbnail noticeably wider).
        # Note: grid.items[] does NOT accept a per-item onClick on this webhook endpoint
        # (confirmed: 400 "Unknown name onClick" at grid.items[N]) — only a single
        # grid-level onClick exists, which can't distinguish which thumbnail was tapped.
        # So the grid is thumbnails-only, and a "원본 보기" link list (each item numbered
        # to match its position) gives click-through to the full-size image instead.
        # cropStyle RECTANGLE_4_3 (not the default SQUARE) so a very tall/long attachment
        # isn't center-cropped down to a sliver (user feedback 2026-09-30, /revision row 25,
        # ticket #1000164250 — a long image was barely visible in the square thumbnail).
        extra.append({"grid": {
            "columnCount": 2,
            "items": [
                {
                    "id": str(i),
                    "image": {
                        "imageUri": url,
                        "cropStyle": {"type": "RECTANGLE_4_3"},
                        "borderStyle": {"type": "NO_BORDER"},
                    },
                }
                for i, url in enumerate(images)
            ],
        }})
        originals = " · ".join(
            f'<a href="{url}">원본{i + 1}</a>' for i, url in enumerate(images)
        )
        extra.append({"textParagraph": {"text": originals}})
    if attaches:
        links = "<br>".join(
            f'<a href="{u}">{n}</a>' for n, u in attaches
        )
        extra.append({"textParagraph": {"text": links}})
    if extra:
        sections.append({"header": "첨부 (Attachments)", "widgets": extra})

    # Card only — no separate "text" line above the app card (user rule 2026-09-10).
    return {
        "cardsV2": [{
            "cardId": "ticket-report",
            "card": {"sections": sections},
        }],
    }


def post(webhook: str, payload: dict):
    """Returns (exit_code, response_json_or_None)."""
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        webhook, data=data,
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode("utf-8", "replace")
            try:
                resp = json.loads(body)
                print("SENT:", resp.get("name", body[:200]))
                return 0, resp
            except json.JSONDecodeError:
                print("SENT:", body[:200])
                return 0, None
    except urllib.error.HTTPError as e:
        print(f"HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:800]}", file=sys.stderr)
        return 1, None
    except Exception as e:  # noqa: BLE001
        print(f"ERROR: {e}", file=sys.stderr)
        return 1, None


def write_queue_row(ticket_id: str, thread_name: str) -> None:
    """Writes {ticketId, threadId} into the TicketQueue sheet so the Ticket Reporter Chat
    app can resolve a plain thread reply back to this ticket (lookupTicketByThread_ in
    Code.gs). Best-effort: a failure here is logged but never fails the send itself."""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build

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

        sheets = build("sheets", "v4", credentials=creds)
        row = [
            str(ticket_id), "", "", "",
            datetime.datetime.now().isoformat(timespec="seconds"),
            "", "", thread_name,
        ]
        sheets.spreadsheets().values().append(
            spreadsheetId=QUEUE_SHEET_ID,
            range=f"{QUEUE_TAB}!A:H",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()
        print(f"QUEUED: ticket {ticket_id} -> {thread_name}")
    except Exception as e:  # noqa: BLE001
        print(f"WARN: could not write thread mapping ({e})", file=sys.stderr)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--html", help="path to report body rendered as Chat HTML (card mode)")
    ap.add_argument("--file", help="path to plain-text report (legacy text mode); omit to read stdin")
    ap.add_argument("--image", action="append", default=[], metavar="URL",
                    help="attachment image URL -> image widget (repeatable)")
    ap.add_argument("--attach", action="append", default=[], metavar="NAME|URL",
                    help="attachment link 'name|url' -> link list (repeatable)")
    ap.add_argument("--webhook", default=DEFAULT_WEBHOOK)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--ticket-id", help="records {ticketId, threadId} in the TicketQueue "
                     "sheet after a successful send, so a thread reply mentioning the app "
                     "later resolves back to this ticket. In --html mode this is normally "
                     "auto-derived from the report's own [Ticket Info.] agent/tickets/<id> "
                     "link (which always wins on a mismatch) — pass this explicitly only "
                     "for --file/stdin mode, where there is no link to parse")
    ap.add_argument("--thread", help="existing Chat thread resource name "
                     "(spaces/.../threads/...) to reply inside instead of starting a new "
                     "thread — e.g. a /btw question's threadId from BtwQueue. Adds "
                     "messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD to the "
                     "webhook URL automatically; works in both --html and --file/stdin mode.")
    args = ap.parse_args()

    attaches = []
    for a in args.attach:
        name, _, url = a.partition("|")
        if url:
            attaches.append((name.strip() or "attachment", url.strip()))

    ticket_id = args.ticket_id
    if args.html:
        html = _normalize(open(args.html, encoding="utf-8").read())
        if not html:
            print("ERROR: empty --html body", file=sys.stderr)
            return 2
        payload = build_card(html, args.image, attaches)
        # Always derive the ticket id from the report's own [Ticket Info.] link rather
        # than relying solely on a manually-passed --ticket-id (user rule 2026-09-16 —
        # every monitor-mode send before this had no --ticket-id, so no thread mapping
        # was ever written and thread replies couldn't resolve back to a ticket).
        m = re.search(r"agent/tickets/(\d+)", html)
        if m:
            if ticket_id and ticket_id != m.group(1):
                print(f"WARN: --ticket-id {ticket_id} != link ticket {m.group(1)} in body; "
                      f"using the link's id", file=sys.stderr)
            ticket_id = m.group(1)
    else:
        text = (open(args.file, encoding="utf-8").read() if args.file else sys.stdin.read()).strip()
        if not text:
            print("ERROR: empty report text", file=sys.stderr)
            return 2
        payload = {"text": text}

    webhook = args.webhook
    if args.thread:
        payload["thread"] = {"name": args.thread}
        if "messageReplyOption" not in webhook:
            sep = "&" if "?" in webhook else "?"
            webhook = webhook + sep + "messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD"

    if args.dry_run:
        print("--- DRY RUN (not sent) ---")
        print(json.dumps(payload, ensure_ascii=False, indent=1))
        return 0

    code, resp = post(webhook, payload)
    if code == 0 and ticket_id and resp:
        thread_name = (resp.get("thread") or {}).get("name")
        if thread_name:
            write_queue_row(ticket_id, thread_name)
        else:
            print("WARN: no thread name in webhook response; thread mapping not written",
                  file=sys.stderr)
    elif code == 0 and not ticket_id:
        print("WARN: no ticket id found (no --ticket-id and no agent/tickets/<id> link in "
              "body); thread mapping not written", file=sys.stderr)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
