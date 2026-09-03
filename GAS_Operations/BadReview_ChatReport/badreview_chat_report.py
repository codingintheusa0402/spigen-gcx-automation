#!/usr/bin/env python3
"""BadReview → Google Chat app-card report (Pixel 11 Series & Galaxy Z8 Series).

Standalone version of the Claude Code skills `pixel11-badreview-chat-report`,
`glxz8-badreview-chat-report` and `badreview-chat-broadcast`. Same cards, same
safety flow, no browser needed — it reads the sheets directly through the Sheets
API v4 using the local gws_shim OAuth token.

For each product it reads the `1-3점` tab and builds a Google Chat cardsV2 card:

  header.title    ✔️ M/D(요일) <product> 배드리뷰 (1~3점) (총 N건)
  header.subtitle 고객 리뷰 ★1~3점 · <range>            + product thumbnail
  "Top 5 인입사유(누적)"  2 columns by 대분류 (휴대폰보호필름 | 휴대폰케이스),
                          each a fixed 5 ranked lines "n위 이유 c건 (p%)"
  "오늘 M/D(요일) 최다 인입사유"  top 인입사유(tag) of rows whose Update 날짜 is
                                today, then a fixed 5-line breakdown
  [배드리뷰] button → the 1-3점 sheet

N (총 N건) = count of `1-3점` rows whose `Update 날짜` resolves to today (KST).
Both cards are padded to the same line counts so they render the same height.

USAGE
  # always test first — posts both cards to the TEST room only
  python3 badreview_chat_report.py --test

  # then, only after a human says yes, broadcast to every team room
  python3 badreview_chat_report.py --broadcast --yes
  python3 badreview_chat_report.py --broadcast --yes --only "ADS1,JP Sales"

  # other flags
  --date 2026-09-02     # override "today"
  --product pixel11|glxz8   # restrict to one product (default: both)
  --dry-run             # build + print the cards, send nothing
  --print-data          # also print the crunched numbers

The broadcast room list is ROOMS below (source of truth: the user's Google Chat
webhooks; mirrored in Claude memory `gcx_team_gchat_webhooks.md`). The internal GCX
team room is intentionally NOT in it.
"""
from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
import time
import urllib.error
import urllib.request

# --------------------------------------------------------------------------- auth
TOKEN_PATH = os.path.expanduser("~/.config/gws_shim/token.json")


def sheets_service():
    """Return an authorized Sheets API v4 client using the gws_shim token."""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    with open(TOKEN_PATH) as fh:
        data = json.load(fh)
    creds = Credentials(
        data["token"],
        refresh_token=data["refresh_token"],
        token_uri=data["token_uri"],
        client_id=data["client_id"],
        client_secret=data["client_secret"],
        scopes=data["scopes"],
    )
    creds.refresh(Request())
    data["token"] = creds.token
    with open(TOKEN_PATH, "w") as fh:
        json.dump(data, fh)
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# ----------------------------------------------------------------------- products
GID_13 = 970309432  # '1-3점' tab id is the same on both spreadsheets (informational)

KEY = "AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI"

PRODUCTS = {
    "pixel11": {
        "name": "Pixel 11 Series",
        "sheet_id": "12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI",
        "subtitle": "고객 리뷰 ★1~3점 · 26/8/18~26/11/18",
        "card_id": "pixel11-badreview-upload",
        "header_img": (
            "https://encrypted-tbn2.gstatic.com/shopping?q=tbn:ANd9GcQhY2aafxhQi-vGv0oxV5j0"
            "kiiOF2sGF0hwiXeEePaAI3DbRziTZcO4Z2sehnyCpp1_qSxCn_iAE4IZ0SlW9WftxRQLxykwNXmmsDn"
            "m3CQkubwlCmO7PL4F3JbUKGWpl1F6c2RuVw&usqp=CAc"
        ),
    },
    "glxz8": {
        "name": "Galaxy Z8 Series",
        "sheet_id": "19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4",
        "subtitle": "고객 리뷰 ★1~3점 · 26/7/27~26/10/27",
        "card_id": "glxz8-badreview-upload",
        "header_img": (
            "https://encrypted-tbn0.gstatic.com/shopping?q=tbn:ANd9GcRA_H2PEgYRDyPvE2kQ4RQ"
            "lhN4sOnJd5cIS90muwFk2pqDlNNPQGlXJ8DHo7ihE20nzMs6-C5AUao7n5SgfGH4vTuzzrg6Oh_4QGU"
            "SKReAVwWyaJK1DN9MI_TrAT8dE3Gq1-Rzobw&usqp=CAc"
        ),
    },
}

CATEGORIES = ("휴대폰보호필름", "휴대폰케이스")

KOR_WD = ["월", "화", "수", "목", "금", "토", "일"]

# ------------------------------------------------------------------------- rooms
# Each room: name, space id, and the default webhook token. `glxz8` is an OPTIONAL
# per-room override token used ONLY for the Galaxy Z8 card — same room, a separate
# incoming webhook. Pixel 11 always uses `token`. If a room has no `glxz8` key, the
# Z8 card goes through `token` too.
TEST_ROOM = {"name": "TEST", "sid": "AAQAc9NQmJQ",
             "token": "Nvngg3UoVU-M7TqqlC48NxP-SXRzXj9zWrIoqd4BJdo"}

ROOMS = [
    {"name": "GCX전략 x SDA (아마존직판)", "sid": "AAAAe96DDIs",
     "token": "fqRQJsNX1O8LDUyjmFsKdBJo5VCPXuFA2KX2OfAuLXk"},
    {"name": "GCX전략 x ADS1", "sid": "AAAAFjWzr40",
     "token": "Xv5J3ipKs_mIOem7OMHzhhmPwHcTrC-wDgYlMSZHAzs",
     "glxz8": "O6gHRrVCB3-X31BJgpUTncKObq2tCLEdk4cMXmwLiR0"},
    {"name": "GCX전략 x ADS2", "sid": "AAAATOmW7HU",
     "token": "GsprARTa_2ga2mkdz8lEFe2K4DTTRl7zfpBW6qEvlOU",
     "glxz8": "sgex588AiCc0FqI_XOEEHK-lOsWhRlw8QdVKkcDsM9g"},
    {"name": "GCX전략 x ADS3", "sid": "AAAAKwBoZPU",
     "token": "zW6cEhLwMozY2v3DvH9nvk4eFW8kSpwlx3MFLYbFFBE"},
    {"name": "GCX전략 x ADS5 (CP)", "sid": "AAAACMOrahk",
     "token": "sV1IIpqIWGQMZIItCEnHyDObAPAjsEZUvah2NKY4iC8",
     "glxz8": "uG4lD7n5oYOn3smRQo7zGGENRPgpeHSEMr-xU2RXPac"},
    {"name": "GCX전략 x JP Sales", "sid": "AAAA9VYH3s4",
     "token": "HLse4WgYcISsdtHF3zNYIi2I5BtEnR6zuGamlCF0cQY",
     "glxz8": "MzwxfPqlWxKLLI1t-R782YnF5RZco8r6m-NypiUZBeo"},
    {"name": "GCX전략 x IN Sales", "sid": "AAAAhqi-tNo",
     "token": "OwCl9xRwf3e4b9hFk6Ieu2h3RDMFv82TkfJmBcvSVvE",
     "glxz8": "PWEne2_aAL8eyK_VKtf4bdhb2mEjFwHInQxK0Nnnm9g"},
    {"name": "GCX전략 x 모바일제품개발팀", "sid": "AAAAwixNbdc",
     "token": "P2vgJp4v0mt1rbAJQmSROlDwmHnf5bKbNryqf_iWDYc",
     "glxz8": "ujAd5HUFylFysq0zGnTcYmqtb7P9GTLUKdHTUk5YHO8"},
    {"name": "실장님 & GCX", "sid": "AAQAb-u6r7s",
     "token": "AOJntA_PdElbBaGzQaCQhhr0aBvPAy1k3ImqQK0V9_E",
     "glxz8": "pyyBw1Gh-X4djUTd2utbHUxs7pVDJu2YCuce1Evf7tk"},
    {"name": "GCX x 클리어프로텍션 개발팀", "sid": "AAAAZcIQG8k",
     "token": "tb4sDPPPaWeP83HPMH0IUnz96T2D6azY1TAoXdiWqGg",
     "glxz8": "0Tw7OvEG60guBwAkLnjRLs6AWm6avHQBoMrewzP6cdE"},
    {"name": "[CQ] SPIGEN 국내&해외 CS", "sid": "AAAA45iXDL0",
     "token": "zh67JI0vK1DIeoet937rQ2byrOin9gV98FQddSSvfmY",
     "glxz8": "Bplrki7kUVeXMMCkkpplYn4QK1g-aMR0X-0RPYgqCOs"},
]

WEBHOOK = "https://chat.googleapis.com/v1/spaces/{sid}/messages?key=" + KEY + "&token={tok}"


def room_url(room: dict, product_key: str) -> str:
    """Webhook URL for a room + product. Z8 uses the room's `glxz8` override token
    when present; everything else uses the default `token`."""
    tok = room["glxz8"] if (product_key == "glxz8" and "glxz8" in room) else room["token"]
    return WEBHOOK.format(sid=room["sid"], tok=tok)


# --------------------------------------------------------------------- crunching
def _resolve_col(header: list[str], *names: str) -> int:
    for want in names:
        if want in header:
            return header.index(want)
    # forgiving fallback: prefix match (handles "인입사유(AI)  Acc. NN%" style headers)
    for want in names:
        for i, h in enumerate(header):
            if h.startswith(want):
                return i
    raise KeyError(f"none of {names} in header")


def _date_tuple(value: str):
    import re

    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", value or "")
    return (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def crunch(rows: list[list[str]], today: datetime.date) -> dict:
    header = rows[0]
    i_upd = _resolve_col(header, "Update 날짜", "Exported Date")
    i_tag = _resolve_col(header, "인입사유(tag)")
    i_cat = _resolve_col(header, "대분류")
    tkey = (today.year, today.month, today.day)

    today_count = 0
    today_tally: dict[str, int] = {}
    cat: dict[str, dict[str, int]] = {c: {} for c in CATEGORIES}

    for row in rows[1:]:
        if len(row) <= max(i_upd, i_tag, i_cat):
            row = row + [""] * (max(i_upd, i_tag, i_cat) + 1 - len(row))
        tag = (row[i_tag] or "").strip() or "(빈칸)"
        cc = (row[i_cat] or "").strip()
        if cc in cat:
            cat[cc][tag] = cat[cc].get(tag, 0) + 1
        if _date_tuple(row[i_upd]) == tkey:
            today_count += 1
            today_tally[tag] = today_tally.get(tag, 0) + 1

    def block(counts: dict[str, int]) -> dict:
        return {
            "tot": sum(counts.values()),
            "top5": sorted(counts.items(), key=lambda kv: -kv[1])[:5],
        }

    return {
        "todayCount": today_count,
        "todayTags": sorted(today_tally.items(), key=lambda kv: -kv[1]),
        "film": block(cat["휴대폰보호필름"]),
        "case": block(cat["휴대폰케이스"]),
    }


# ------------------------------------------------------------------ card builder
def _pad(lines: list[str], n: int = 5) -> str:
    lines = list(lines[:n]) + ["&nbsp;"] * max(0, n - len(lines))
    return "<br>".join(lines)


def _rank_widgets(rows: list, tot: int) -> list[dict]:
    """Top 5 rows as decoratedText widgets so rank / 인입사유 / 건수·% each sit at a
    fixed left edge and the two 대분류 columns line up."""
    out = []
    for i, (name, c) in enumerate(rows[:5]):
        pct = f"{round(c * 100 / tot)}%" if tot else "-"
        out.append({"decoratedText": {
            "topLabel": f"{i + 1}위",
            "text": f"<b>{name}</b>",
            "bottomLabel": f"{int(c)}건 · {pct}",
        }})
    while len(out) < 5:
        out.append({"decoratedText": {"topLabel": " ", "text": " ", "bottomLabel": " "}})
    return out


# 대분류 column-header colors (bright enough for Chat light + dark themes) so the
# header stands out from the 인입사유 rows under it.
CAT_COLORS = {"휴대폰보호필름": "#EA4335", "휴대폰케이스": "#4285F4"}


def _cat_column(label: str, block: dict) -> dict:
    tot = int(block.get("tot", 0))
    rows = [(n, int(c)) for n, c in block.get("top5", [])]
    color = CAT_COLORS.get(label, "#202124")
    return {
        "horizontalSizeStyle": "FILL_AVAILABLE_SPACE",
        "horizontalAlignment": "START",
        "verticalAlignment": "TOP",
        "widgets": [
            {"textParagraph": {"text": f'<b><font color="{color}">{label}</font></b>  ·  {tot}건'}},
            *_rank_widgets(rows, tot),
        ],
    }


def build_card(product_key: str, data: dict, today: datetime.date) -> dict:
    p = PRODUCTS[product_key]
    date_str = f"{today.month}/{today.day}({KOR_WD[today.weekday()]})"
    today_count = int(data["todayCount"])
    today_tags = [(n, int(c)) for n, c in data.get("todayTags", [])]
    link = f"https://docs.google.com/spreadsheets/d/{p['sheet_id']}/edit?gid={GID_13}#gid={GID_13}"

    title = f"✔️ {date_str} {p['name']} 배드리뷰 (1~3점) (총 {today_count}건)"

    top5_widgets = [{
        "columns": {
            "columnItems": [
                _cat_column("휴대폰보호필름", data.get("film", {})),
                _cat_column("휴대폰케이스", data.get("case", {})),
            ]
        }
    }]

    if today_tags:
        top_name, top_n = today_tags[0]
        lines = [f"{i + 1}. {n} &nbsp;{c}건" for i, (n, c) in enumerate(today_tags[:5])]
        if len(today_tags) > 5:
            lines[-1] += f" &nbsp;…외 {sum(c for _, c in today_tags[5:])}건"
        today_widgets = [
            {"decoratedText": {
                "topLabel": "인입사유(tag) 기준",
                "text": f"<b>{top_name}</b> — {top_n}건",
                "startIcon": {"knownIcon": "STAR"},
            }},
            {"textParagraph": {"text": _pad(lines, 5)}},
        ]
    else:
        today_widgets = [
            {"decoratedText": {
                "topLabel": "인입사유(tag) 기준",
                "text": "오늘 업로드된 배드리뷰 없음",
                "startIcon": {"knownIcon": "STAR"},
            }},
            {"textParagraph": {"text": _pad([], 5)}},
        ]

    today_widgets.append({"buttonList": {"buttons": [{
        "text": "배드리뷰",
        "onClick": {"openLink": {"url": link}},
    }]}})

    return {"cardsV2": [{
        "cardId": p["card_id"],
        "card": {
            "header": {
                "title": title,
                "subtitle": p["subtitle"],
                "imageUrl": p["header_img"],
                "imageType": "SQUARE",
            },
            "sections": [
                {"header": "Top 5 인입사유(누적)", "widgets": top5_widgets},
                {"header": f"오늘 {date_str} 최다 인입사유", "widgets": today_widgets},
            ],
        },
    }]}


# -------------------------------------------------------------------- transport
def post(url: str, card: dict) -> str:
    req = urllib.request.Request(
        url,
        data=json.dumps(card).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    try:
        resp = urllib.request.urlopen(req, timeout=60)
        return "OK   " + json.load(resp).get("name", "")
    except urllib.error.HTTPError as exc:
        return f"ERR {exc.code} {exc.read().decode()[:300]}"


# -------------------------------------------------------------------------- main
def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--test", action="store_true", help="post both cards to the TEST room only")
    mode.add_argument("--broadcast", action="store_true", help="post to every team room (needs --yes)")
    mode.add_argument("--dry-run", action="store_true", help="build + print the cards, send nothing")
    ap.add_argument("--yes", action="store_true", help="required confirmation for --broadcast")
    ap.add_argument("--only", help="comma-separated room-name substrings to restrict --broadcast")
    ap.add_argument("--product", choices=list(PRODUCTS), help="only this product (default: both)")
    ap.add_argument("--date", help="override today, YYYY-MM-DD")
    ap.add_argument("--print-data", action="store_true", help="also print the crunched numbers")
    args = ap.parse_args()

    if args.broadcast and not args.yes:
        ap.error("--broadcast requires --yes (test with --test first and get a human ok)")

    today = datetime.date.fromisoformat(args.date) if args.date else datetime.date.today()
    keys = [args.product] if args.product else list(PRODUCTS)

    svc = sheets_service()
    cards: list[tuple[str, str, dict]] = []  # (product_key, product_name, card)
    for key in keys:
        p = PRODUCTS[key]
        values = svc.spreadsheets().values().get(
            spreadsheetId=p["sheet_id"], range="1-3점!A:S"
        ).execute().get("values", [])
        if not values:
            print(f"!! {key}: '1-3점' returned no rows", file=sys.stderr)
            sys.exit(1)
        data = crunch(values, today)
        if args.print_data:
            print(f"# {key}: {json.dumps(data, ensure_ascii=False)}")
        if data["todayCount"] == 0:
            print(f"!! {key}: todayCount is 0 for {today} — card will say '없음'")
        cards.append((key, p["name"], build_card(key, data, today)))

    if args.dry_run:
        for _key, name, card in cards:
            print(f"\n===== {name} =====\n{json.dumps(card, ensure_ascii=False, indent=2)}")
        return

    if args.test:
        targets = [TEST_ROOM]
    else:
        targets = ROOMS
        if args.only:
            subs = [s.strip() for s in args.only.split(",") if s.strip()]
            targets = [r for r in targets if any(s in r["name"] for s in subs)]
        if not targets:
            ap.error("--only matched no rooms")

    print(f"date={today}  mode={'TEST' if args.test else 'BROADCAST'}  "
          f"rooms={len(targets)}  cards/room={len(cards)}\n")
    for room in targets:
        for key, cname, card in cards:
            print(f"[{room['name']}] {cname}: {post(room_url(room, key), card)}")
            time.sleep(1.0)


if __name__ == "__main__":
    main()
