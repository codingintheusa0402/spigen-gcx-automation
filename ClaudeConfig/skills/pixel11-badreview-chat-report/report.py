#!/usr/bin/env python3
"""Pixel 11 BadReview Google Chat Report — card builder + sender.

The authenticated CSV reading + number crunching happens in the browser
(see SKILL.md, step 3). This script only takes the small computed summary,
builds a Google Chat cardsV2 message, and POSTs it to the GCX team webhook.

Usage:
  report.py --data '<json>'                # build + send
  report.py --data '<json>' --dry-run      # print payload only, do not send
  report.py --data '<json>' --date 2026-09-01   # override "today"

--data JSON shape (all counts are ints), everything from the '1-3점' sheet:
  {
    "todayCount": 12,
    "todayTags": [["글라스깨짐",3],["재질",2], ...],   # 인입사유(tag) tally, Update 날짜 == today
    "film": {"tot": 82, "top5": [["후면카메라화질저하",10], ...]},  # 대분류 == 휴대폰보호필름
    "case": {"tot": 44, "top5": [["재질",11], ...]},              # 대분류 == 휴대폰케이스
    "recentAvg": 6.3   # OPTIONAL: avg daily count over the trailing 7 days (excl. today).
                       # Missing/0 → the total-count significance check is simply skipped.
  }

Significance highlighting (2026-09-14, user rule): a number is bold+red in the card when
it's "significant comparatively" — >= SIG_RATIO (1.5x) its natural comparison point:
  - 오늘 총 N건 vs `recentAvg` (trailing-7-day daily average, excl. today)
  - each Top-5(누적) column's 1위 count vs its own 2위 count
  - 오늘 최다 인입사유's top tag count vs the day's 2nd-place tag count
Only `decoratedText.text` renders the `<font color>`/`<b>` HTML subset — `topLabel` and
`bottomLabel` do NOT (verified empirically 2026-09-14 against the live webhook: a probe
card showed raw `<font>`/`<b>` tags literally in both). So the highlighted number always
lives in `.text`; `bottomLabel` stays plain (count·%, or a note like "최근 7일 평균 X건").
"""
import argparse, json, sys, datetime, urllib.request, urllib.error, os

WEBHOOK = json.load(open(os.path.expanduser("~/.config/gcx_webhooks.json")))["pixel11_badreview"]  # secret — kept out of git

SIG_RATIO = 1.5           # "significant" = at least this many times the comparison point
SIG_RED   = "#D93025"

SHEET_ID = "12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI"
GID_13   = "970309432"                     # '1-3점' tab
LINK_13  = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid={GID_13}#gid={GID_13}"
HEADER_IMG = ("https://encrypted-tbn2.gstatic.com/shopping?q=tbn:ANd9GcQhY2aafxhQi-vGv0oxV5j0"
             "kiiOF2sGF0hwiXeEePaAI3DbRziTZcO4Z2sehnyCpp1_qSxCn_iAE4IZ0SlW9WftxRQLxykwNXmmsDn"
             "m3CQkubwlCmO7PL4F3JbUKGWpl1F6c2RuVw&usqp=CAc")
KOR_WD   = ['월', '화', '수', '목', '금', '토', '일']


# 인입사유(tag) values kept out of every stat and every card (user rule 2026-09-08:
# exclude 긍정 리뷰 permanently). The step-3 JS already drops these rows before
# counting; this is a second guard in case pre-computed --data still carries them.
EXCLUDED_TAGS = frozenset({"긍정 리뷰"})


# Both cards are kept the SAME height by rendering every list at a fixed line count
# (blank &nbsp; lines pad the shortfall), so Pixel 11 and Z8 render identically.
def _pad(lines, n=5):
    lines = list(lines[:n])
    while len(lines) < n:
        lines.append("&nbsp;")
    return "<br>".join(lines)


# Top 5 rows as decoratedText widgets: rank / 인입사유 / 건수·% each sit at a fixed
# left edge (topLabel, text, bottomLabel) so the two 대분류 columns line up.
def _rank_widgets(rows, tot):
    out = []
    dominant = len(rows) >= 2 and rows[0][1] > 0 and rows[0][1] >= rows[1][1] * SIG_RATIO
    for i, (n, c) in enumerate(rows[:5]):
        pct = f"{round(c*100/tot)}%" if tot else "-"
        if i == 0 and dominant:
            text = f'<b>{n}</b> — <font color="{SIG_RED}"><b>{c}건</b></font>'
        else:
            text = f"<b>{n}</b>"
        out.append({"decoratedText": {
            "topLabel": f"{i+1}위",
            "text": text,
            "bottomLabel": f"{int(c)}건 · {pct}",
        }})
    while len(out) < 5:
        out.append({"decoratedText": {"topLabel": " ", "text": " ",
                                      "bottomLabel": " "}})
    return out


# 대분류 column-header colors (bright enough for Chat light + dark themes) so the
# header stands out from the 인입사유 rows under it.
CAT_COLORS = {"휴대폰보호필름": "#EA4335", "휴대폰케이스": "#4285F4"}


def _cat_column(label, block):
    tot  = int(block.get("tot", 0))
    rows = [(n, int(c)) for n, c in block.get("top5", []) if n not in EXCLUDED_TAGS]
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


def build_card(data, today):
    date_str = f"{today.month}/{today.day}({KOR_WD[today.weekday()]})"
    today_count = int(data["todayCount"])
    today_tags  = [(n, int(c)) for n, c in data.get("todayTags", []) if n not in EXCLUDED_TAGS]
    recent_avg  = data.get("recentAvg")
    total_sig   = recent_avg is not None and recent_avg > 0 and today_count >= recent_avg * SIG_RATIO

    title = f"✔️ {date_str} Pixel 11 Series 배드리뷰 (1~3점) (총 {today_count}건)"

    top5_widgets = [{
        "columns": {
            "columnItems": [
                _cat_column("휴대폰보호필름", data.get("film", {})),
                _cat_column("휴대폰케이스", data.get("case", {})),
            ]
        }
    }]

    # Always-present row so 오늘 총 N건 has a widget home to be highlighted in (the
    # card header can't render HTML/color — see module docstring).
    today_widgets = [{"decoratedText": {
        "topLabel": "오늘 총 업로드",
        "text": (f'<font color="{SIG_RED}"><b>{today_count}건</b></font>' if total_sig
                 else f"<b>{today_count}건</b>"),
        "bottomLabel": f"최근 7일 평균 {recent_avg:.1f}건" if recent_avg is not None else " ",
    }}]

    if today_tags:
        top_name, top_n = today_tags[0]
        today_dom = len(today_tags) >= 2 and top_n >= today_tags[1][1] * SIG_RATIO
        lines = [f"{i+1}. {n} &nbsp;{c}건" for i, (n, c) in enumerate(today_tags[:5])]
        if len(today_tags) > 5:
            lines[-1] += f" &nbsp;…외 {sum(c for _, c in today_tags[5:])}건"
        today_widgets += [
            {"decoratedText": {
                "topLabel": "인입사유(tag) 기준",
                "text": (f'<b>{top_name}</b> — <font color="{SIG_RED}"><b>{top_n}건</b></font>'
                         if today_dom else f"<b>{top_name}</b> — {top_n}건"),
                "startIcon": {"knownIcon": "STAR"},
            }},
            {"textParagraph": {"text": _pad(lines, 5)}},
        ]
    else:
        today_widgets += [
            {"decoratedText": {
                "topLabel": "인입사유(tag) 기준",
                "text": "오늘 업로드된 배드리뷰 없음",
                "startIcon": {"knownIcon": "STAR"},
            }},
            {"textParagraph": {"text": _pad([], 5)}},
        ]

    today_widgets.append({"buttonList": {"buttons": [{
        "text": "배드리뷰",
        "onClick": {"openLink": {"url": LINK_13}},
    }]}})

    return {"cardsV2": [{
        "cardId": "pixel11-badreview-upload",
        "card": {
            "header": {
                "title": title,
                "subtitle": "고객 리뷰 ★1~3점 · 26/8/18~26/11/18",
                "imageUrl": HEADER_IMG,
                "imageType": "SQUARE",
            },
            "sections": [
                {"header": "Top 5 인입사유(누적)", "widgets": top5_widgets},
                {"header": f"오늘 {date_str} 최다 인입사유", "widgets": today_widgets},
            ],
        },
    }]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="computed summary JSON (see module docstring)")
    ap.add_argument("--date", help="override today, YYYY-MM-DD")
    ap.add_argument("--webhook", help="override target Google Chat webhook URL")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    webhook = a.webhook or WEBHOOK

    today = datetime.date.fromisoformat(a.date) if a.date else datetime.date.today()
    data = json.loads(a.data)
    card = build_card(data, today)

    print(json.dumps(card, ensure_ascii=False, indent=2))

    if a.dry_run:
        print("\n[dry-run] not sent.")
        return

    req = urllib.request.Request(
        webhook,
        data=json.dumps(card).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=UTF-8"},
    )
    try:
        resp = urllib.request.urlopen(req, timeout=60)
        print("\ngchat OK:", json.load(resp).get("name"))
    except urllib.error.HTTPError as e:
        print("\ngchat ERR", e.code, e.read().decode())
        sys.exit(1)


if __name__ == "__main__":
    main()
