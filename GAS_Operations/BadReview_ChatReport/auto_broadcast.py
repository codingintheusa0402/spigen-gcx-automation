#!/usr/bin/env python3
"""
BadReview auto-broadcast — unattended weekday version of `badreview-chat-broadcast`.

Runs from a `launchd` LaunchAgent (see launchagent/README.md in this folder), every
weekday at 10:30 AM KST. NO test-send, NO confirmation prompt — by explicit user
request (2026-09-15): "make this a hook that runs every weekdays (except Korean
national holidays) and send automatically to all the chatrooms KST10:30AM without
test sending to my private chatroom". This script is the ONLY broadcast path that
skips the test-room gate — the interactive `badreview-chat-broadcast` skill keeps its
HARD RULE (test first, ask for confirmation) for every manually-triggered run.

Data source: Sheets API v4 with the gws_shim OAuth token (~/.config/gws_shim/token.json),
NOT the browser/gviz method the interactive skill uses — this runs unattended with no
Chrome session available. Card building and the room list are NOT duplicated here: this
script imports report.py from the three per-product skills and broadcast.py from
badreview-chat-broadcast, so all four stay the single source of truth for card layout,
significance-highlighting thresholds, and the 12-room list.

iPhone 18 (added 2026-09-21) reuses each room's `glxz8` webhook token, same as Z8 —
see `room_url()`/`SHARES_GLXZ8_TOKEN` in broadcast.py. It has no KR-gate (that's Z8-only).

Combined carousel format (2026-09-21): all 3 products now go out as ONE message per
room (a swipeable Cards v2 carousel, iPhone 18 → Z8 → Pixel 11 — see carousel.py in
badreview-chat-broadcast/), sent via each room's default `token` webhook, replacing
the old 3-separate-messages-per-room format. Because it's a single message, the Z8
KR-gate (see below) now holds back the ENTIRE carousel when it trips, not just Z8 —
there's no way to selectively omit one page from an already-sent message.

Usage:
  auto_broadcast.py                 # normal run (skips silently on holiday/weekend)
  auto_broadcast.py --dry-run       # crunch + build cards, print results, send nothing
  auto_broadcast.py --force         # send even if today is a weekend/holiday (manual testing)
  auto_broadcast.py --date 2026-09-16   # override "today" (KST) for testing
"""
import argparse, datetime, importlib.util, json, os, sys, time, urllib.error, urllib.parse, urllib.request

SKILLS = os.path.expanduser("~/.claude/skills")
GWS_TOKEN_PATH = os.path.expanduser("~/.config/gws_shim/token.json")
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")

SHEETS = {
    "pixel11":  "12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI",
    "glxz8":    "19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4",
    "iphone18": "1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU",
}

# Fallback if the Nager.Date API is unreachable at run time (network hiccup). Kept in
# sync manually from https://date.nager.at/api/v3/PublicHolidays/2026/KR (fetched
# 2026-09-15) MINUS 근로자의날/Labour Day (May 1 — not a 관공서 공휴일, so NOT skipped).
# Re-derive for next year before this list runs out; the API call is tried first every
# run specifically so this fallback rarely matters.
KR_HOLIDAYS_FALLBACK_2026 = {
    "2026-01-01", "2026-02-16", "2026-02-17", "2026-02-18", "2026-03-02",
    "2026-05-05", "2026-05-25", "2026-06-03", "2026-06-06", "2026-07-17",
    "2026-08-17", "2026-09-24", "2026-09-25", "2026-09-26", "2026-10-05",
    "2026-10-09", "2026-12-25",
}


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def log(msg):
    os.makedirs(LOG_DIR, exist_ok=True)
    line = f"[{datetime.datetime.now().isoformat(timespec='seconds')}] {msg}"
    print(line)
    with open(os.path.join(LOG_DIR, "auto_broadcast.log"), "a") as f:
        f.write(line + "\n")


def kr_holidays(year):
    """{'YYYY-MM-DD', ...} of Korean public (관공서) holidays for `year`."""
    try:
        url = f"https://date.nager.at/api/v3/PublicHolidays/{year}/KR"
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.load(r)
        # Nager's KR set already excludes Labour Day (근로자의날) as of the 2026-09-15
        # check — matches "national holiday" (관공서 공휴일) intent. If that ever
        # changes, filter by `types` here instead of trusting the raw list.
        return {d["date"] for d in data}
    except Exception as e:
        log(f"WARNING: Nager.Date holiday lookup failed ({e}); using hardcoded fallback")
        return KR_HOLIDAYS_FALLBACK_2026 if year == 2026 else set()


def refresh_gws_token():
    tok = json.load(open(GWS_TOKEN_PATH))
    d = urllib.parse.urlencode({
        "client_id": tok["client_id"], "client_secret": tok["client_secret"],
        "refresh_token": tok["refresh_token"], "grant_type": "refresh_token",
    }).encode()
    r = json.load(urllib.request.urlopen(urllib.request.Request(
        "https://oauth2.googleapis.com/token", d)))
    tok["token"] = r["access_token"]
    json.dump(tok, open(GWS_TOKEN_PATH, "w"), indent=2)
    return tok["token"]


def sheets_get(token, sheet_id, a1_range):
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/"
           f"{urllib.parse.quote(a1_range)}?valueRenderOption=FORMATTED_VALUE")
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token})
    return json.load(urllib.request.urlopen(req)).get("values", [])


def fetch_rows(token, sheet_id):
    rows = sheets_get(token, sheet_id, "'1-3점'!A1:Z")
    return rows[0], [r + [""] * (len(rows[0]) - len(r)) for r in rows[1:]]


def count_unfilled_today(header, body, today):
    """How many rows with Update 날짜 == today still have a blank 인입사유(tag) —
    the AI tagging agents can lag behind newly-added rows."""
    iU, iT = header.index("Update 날짜"), header.index("인입사유(tag)")
    return sum(1 for r in body if _parse_date(r[iU]) == today and not (r[iT] or "").strip())


def today_kr_count(header, body, today):
    """How many of today's rows are from Korea (국가(tag) == 'KR')."""
    iU, iN = header.index("Update 날짜"), header.index("국가(tag)")
    return sum(1 for r in body if _parse_date(r[iU]) == today and (r[iN] or "").strip().upper() == "KR")


def fetch_ready(token, sheet_id, today, label, max_retries=3, wait_s=600):
    """fetch_rows, retrying (per user rule, 2026-09-18) up to `max_retries` times,
    `wait_s` apart, while any of today's rows still have a blank 인입사유(tag). Sends
    whatever it has after the last retry either way — never blocks a scheduled run
    indefinitely — but the caller is told how many rows were still unfilled."""
    for attempt in range(max_retries + 1):
        header, body = fetch_rows(token, sheet_id)
        unfilled = count_unfilled_today(header, body, today)
        if unfilled == 0:
            return header, body, 0
        if attempt < max_retries:
            log(f"{label}: {unfilled} today-row(s) still missing 인입사유(tag) "
                f"(attempt {attempt + 1}/{max_retries + 1}) — waiting {wait_s // 60} min")
            time.sleep(wait_s)
    log(f"{label}: still {unfilled} today-row(s) unfilled after {max_retries} retries — sending anyway")
    return header, body, unfilled


def crunch(header, body, today):
    """Same shape/logic as the interactive skills' step-3 JS (todayCount, todayTags,
    recentAvg, film, case) — see pixel11-badreview-chat-report/SKILL.md."""
    iU, iT, iC = header.index("Update 날짜"), header.index("인입사유(tag)"), header.index("대분류")

    today_count = 0
    tal, day_counts = {}, {}
    cat = {"휴대폰보호필름": {}, "휴대폰케이스": {}}
    for r in body:
        tag = (r[iT] or "").strip() or "(빈칸)"
        if tag == "긍정 리뷰":
            continue
        cc = (r[iC] or "").strip()
        if cc in cat:
            cat[cc][tag] = cat[cc].get(tag, 0) + 1
        d = _parse_date(r[iU])
        if d:
            day_counts[d] = day_counts.get(d, 0) + 1
            if d == today:
                today_count += 1
                tal[tag] = tal.get(tag, 0) + 1

    recent_sum = sum(day_counts.get(today - datetime.timedelta(days=k), 0) for k in range(1, 8))

    def block(c):
        top5 = sorted(c.items(), key=lambda kv: -kv[1])[:5]
        return {"tot": sum(c.values()), "top5": top5}

    return {
        "todayCount": today_count,
        "todayTags": sorted(tal.items(), key=lambda kv: -kv[1]),
        "recentAvg": recent_sum / 7,
        "film": block(cat["휴대폰보호필름"]),
        "case": block(cat["휴대폰케이스"]),
    }


def _parse_date(s):
    import re
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s or "")
    return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="build cards, send nothing")
    ap.add_argument("--test-only", action="store_true",
                     help="send BOTH cards to the private test room ONLY, never the 12 live "
                          "rooms — implies --force (weekday/holiday skip doesn't apply to a "
                          "test send). ALWAYS use this, not a bare live run, to test anything "
                          "here — see the 2026-09-21 incident in AUTO_BROADCAST.md")
    ap.add_argument("--force", action="store_true", help="ignore the weekday/holiday skip")
    ap.add_argument("--ignore-kr-gate", action="store_true",
                     help="send the Z8 card to all rooms even if 0 KR reviews today "
                          "(after you've manually confirmed that's correct)")
    ap.add_argument("--date", help="override today (KST), YYYY-MM-DD")
    a = ap.parse_args()

    today = datetime.date.fromisoformat(a.date) if a.date else datetime.date.today()

    if not (a.force or a.test_only):
        if today.weekday() >= 5:  # 5=Sat, 6=Sun
            log(f"SKIP {today.isoformat()}: weekend")
            return
        if today.isoformat() in kr_holidays(today.year):
            log(f"SKIP {today.isoformat()}: Korean public holiday")
            return

    log(f"RUN {today.isoformat()}: fetching sheets")
    token = refresh_gws_token()
    px_header, px_body, px_unfilled = fetch_ready(token, SHEETS["pixel11"], today, "pixel11")
    z8_header, z8_body, z8_unfilled = fetch_ready(token, SHEETS["glxz8"], today, "glxz8")
    ip18_header, ip18_body, ip18_unfilled = fetch_ready(token, SHEETS["iphone18"], today, "iphone18")
    px_data = crunch(px_header, px_body, today)
    z8_data = crunch(z8_header, z8_body, today)
    ip18_data = crunch(ip18_header, ip18_body, today)
    if px_unfilled or z8_unfilled or ip18_unfilled:
        log(f"NOTE: sending with pixel11={px_unfilled} glxz8={z8_unfilled} iphone18={ip18_unfilled} "
            f"today-row(s) still missing 인입사유(tag) — retries exhausted")
    log(f"  pixel11  todayCount={px_data['todayCount']} recentAvg={px_data['recentAvg']:.2f}")
    log(f"  glxz8    todayCount={z8_data['todayCount']} recentAvg={z8_data['recentAvg']:.2f}")
    log(f"  iphone18 todayCount={ip18_data['todayCount']} recentAvg={ip18_data['recentAvg']:.2f}")

    px_report = _load(f"{SKILLS}/pixel11-badreview-chat-report/report.py", "auto_px_report")
    z8_report = _load(f"{SKILLS}/glxz8-badreview-chat-report/report.py", "auto_z8_report")
    ip18_report = _load(f"{SKILLS}/iphone18-badreview-chat-report/report.py", "auto_ip18_report")
    broadcast = _load(f"{SKILLS}/badreview-chat-broadcast/broadcast.py", "auto_broadcast_mod")
    carousel = _load(f"{SKILLS}/badreview-chat-broadcast/carousel.py", "auto_carousel_mod")

    px_card = px_report.build_card(px_data, today)
    z8_card = z8_report.build_card(z8_data, today)
    ip18_card = ip18_report.build_card(ip18_data, today)

    # KR is Z8's single largest country segment (Pixel 11 has none at all) and KR
    # reviews occasionally land after 11 AM — past this 10:30 run. Zero KR rows today
    # is therefore a signal the upload may still be incomplete, not necessarily a
    # real zero day. Per user rule (2026-09-18): don't auto-broadcast Z8 in that case
    # — alert the private room and hold it for a manual, confirmed resend instead.
    z8_kr_count = today_kr_count(z8_header, z8_body, today)
    # --test-only never touches the 12 live rooms, so there's nothing to hold back —
    # both cards go to the private room together, same as a normal test-send would.
    hold_z8 = z8_kr_count == 0 and not a.ignore_kr_gate and not a.test_only
    if z8_kr_count == 0:
        why = "--test-only, sending anyway" if a.test_only else \
              "--ignore-kr-gate set, sending anyway" if a.ignore_kr_gate else "holding Z8 broadcast"
        log(f"NOTE: 0 KR reviews in today's Z8 data — {why}")

    targets = [broadcast.TEST_ROOM] if a.test_only else broadcast.ROOMS

    header_imgs = {"pixel11": px_report.HEADER_IMG, "glxz8": z8_report.HEADER_IMG,
                    "iphone18": ip18_report.HEADER_IMG}
    cards = {"pixel11": px_card, "glxz8": z8_card, "iphone18": ip18_card}

    if a.dry_run:
        log(f"[dry-run] cards built, not sending. {'Test room' if a.test_only else 'Rooms'} "
            f"that would receive them:")
        for room in targets:
            log(f"  - {room['name']}")
        if hold_z8:
            log("[dry-run] would hold the ENTIRE carousel and alert the private room instead "
                "(0 KR reviews for Z8 today)")
        return

    if hold_z8:
        test_url = broadcast.BASE.format(sid=broadcast.TEST_ROOM["sid"], tok=broadcast.TEST_ROOM["token"])
        preview = carousel.build_carousel_message(cards, header_imgs, today,
                                                    card_id="badreview-carousel-held")
        alert = {
            "text": (f"⚠️ 배드리뷰 캐러셀 자동발송 전체 보류 — 오늘({today.isoformat()}) Z8 KR 리뷰 0건.\n"
                     f"KR 리뷰는 간혹 11시 이후 업로드되는 경우가 있어, 확인 후 수동 재발송이 필요합니다.\n"
                     f"(카드 3개 전부 아래 미리보기로 확인 가능)\n"
                     f"확인 후 재발송: python3 auto_broadcast.py --force --ignore-kr-gate"),
        }
        log("ALERT (0 KR reviews, entire carousel held): " + broadcast._post(test_url, alert))
        log("ALERT preview: " + broadcast._post(test_url, preview))
        log(f"DONE {today.isoformat()}: carousel HELD ENTIRELY (0 KR reviews for Z8) — "
            f"alert + preview posted to private room")
        return

    message = carousel.build_carousel_message(cards, header_imgs, today)
    for room in targets:
        url = broadcast.BASE.format(sid=room["sid"], tok=room["token"])
        log(f"[{room['name']}] carousel : " + broadcast._post(url, message))
        time.sleep(1.0)

    log(f"DONE {today.isoformat()}: sent to {len(targets)} {'test' if a.test_only else ''} room(s)")


if __name__ == "__main__":
    main()
