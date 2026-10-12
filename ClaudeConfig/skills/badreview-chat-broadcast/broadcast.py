#!/usr/bin/env python3
"""BadReview chat broadcast — sends the Pixel 11 + Galaxy Z8 + iPhone 18 배드리뷰 app
cards to Google Chat rooms.

Default format (2026-09-21): when all 3 products are sent (the normal daily
broadcast), they go out as ONE combined Cards v2 `carousel` message per room —
3 swipeable pages, iPhone 18 → Galaxy Z8 → Pixel 11 — via the room's default
`token` webhook (see `carousel.py`, same directory). A `--product` subset (used
for a one-off corrective resend of a single product) still sends that product's
own normal single card via the old per-product routing (`room_url`) — a 1-page
"carousel" isn't a real carousel, so this path is unchanged from before.

SAFETY: always run `--test-only` first (posts to the test room), show the user, and
get an explicit "yes" before running `--all`. See SKILL.md.

Usage:
  broadcast.py --test-only --px-data '<json>' --z8-data '<json>' --ip18-data '<json>'
  broadcast.py --all       --px-data ... --z8-data ... --ip18-data ...   # all team rooms
  broadcast.py --all --only "ADS1,JP Sales" --px-data ...   # subset (still needs all
      3 --*-data unless --product narrows which products are being sent)
  broadcast.py --all --exclude "리더들방" --product glxz8 --z8-data '<json>'
      # resend/correct just ONE product's card, skipping named room(s) — only that
      # product's --*-data is required
  ... optional: --date YYYY-MM-DD   (else today, local)

--px-data / --z8-data / --ip18-data JSON shapes are exactly what the product
report.py scripts take:
  {"todayCount": int, "todayTags": [[tag,c],...],
   "film": {"tot": int, "top5": [[tag,c],...]},
   "case": {"tot": int, "top5": [[tag,c],...]}}
"""
import argparse, importlib.util, json, os, sys, datetime, urllib.request, urllib.error, time

# All secrets (key, test room, and the full room list incl. tokens) live in the
# gitignored ~/.config/gcx_webhooks.json — this file is tracked in the (public) GCX
# repo, so no token may be hardcoded here (2026-10-12, see memory
# feedback_secrets_and_skills_in_git). Source of truth for room names/ids/tokens is
# still memory gcx_team_gchat_webhooks.md; keep that and this config in sync.
_CFG = json.load(open(os.path.expanduser("~/.config/gcx_webhooks.json")))

KEY = _CFG["badreview_broadcast_key"]
BASE = "https://chat.googleapis.com/v1/spaces/{sid}/messages?key=" + KEY + "&token={tok}"

TEST_ROOM = _CFG["badreview_test_room"]

# Broadcast targets — the rooms the user enumerated for "share the bad review scraped data".
# `token` = default webhook (used by the Pixel 11 card, and by Z8/iPhone18 where no override).
# `glxz8` = OPTIONAL per-room override token used for BOTH the Galaxy Z8 card AND (per
# explicit user instruction, 2026-09-21) the iPhone 18 card — same room, one separate
# incoming webhook shared by both products rather than a third dedicated one.
ROOMS = _CFG["badreview_rooms"]

# product_key -> (skill folder, --*-data dest, short log label). Order here is send order.
PRODUCTS = {
    "glxz8":    ("glxz8-badreview-chat-report", "z8_data", "Z8"),
    "iphone18": ("iphone18-badreview-chat-report", "ip18_data", "IP18"),
    "pixel11":  ("pixel11-badreview-chat-report", "px_data", "PX"),
}
# room_url product_keys that share the room's `glxz8` override token instead of the
# default `token` (2026-09-21: iPhone18 explicitly reuses Z8's per-room webhooks).
SHARES_GLXZ8_TOKEN = {"glxz8", "iphone18"}


def room_url(room, product_key):
    tok = room["glxz8"] if (product_key in SHARES_GLXZ8_TOKEN and "glxz8" in room) else room["token"]
    return BASE.format(sid=room["sid"], tok=tok)

SKILLS = os.path.expanduser("~/.claude/skills")


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


# Loaded dynamically (not `import carousel`) because this file is itself sometimes
# loaded dynamically by auto_broadcast.py, which doesn't add this directory to
# sys.path — a plain import only works when broadcast.py is run as the entry script.
carousel = _load(os.path.join(os.path.dirname(os.path.abspath(__file__)), "carousel.py"),
                  "badreview_carousel")


def _post(url, card):
    req = urllib.request.Request(url, data=json.dumps(card).encode("utf-8"),
                                 headers={"Content-Type": "application/json; charset=UTF-8"})
    try:
        r = urllib.request.urlopen(req, timeout=60)
        return "OK   " + json.load(r).get("name", "")
    except urllib.error.HTTPError as e:
        return f"ERR {e.code} {e.read().decode()[:300]}"


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--test-only", action="store_true", help="post to the TEST room only")
    g.add_argument("--all", action="store_true", help="post to every broadcast room")
    ap.add_argument("--only", help="comma-separated room-name substrings to restrict --all")
    ap.add_argument("--exclude", help="comma-separated room-name substrings to drop from the targets")
    ap.add_argument("--product", default="all",
                     help="comma-separated subset of pixel11,glxz8,iphone18 to send "
                          "(default: all three). 'both' is a legacy alias for pixel11,glxz8.")
    ap.add_argument("--px-data")
    ap.add_argument("--z8-data")
    ap.add_argument("--ip18-data")
    ap.add_argument("--date", help="override today, YYYY-MM-DD")
    a = ap.parse_args()

    if a.product == "all":
        wanted = list(PRODUCTS)
    elif a.product == "both":
        wanted = ["glxz8", "pixel11"]
    else:
        wanted = [p.strip() for p in a.product.split(",") if p.strip()]
        bad = [p for p in wanted if p not in PRODUCTS]
        if bad:
            ap.error(f"unknown --product value(s): {bad} (choices: {list(PRODUCTS)}, or 'all')")

    data_arg = {"pixel11": a.px_data, "glxz8": a.z8_data, "iphone18": a.ip18_data}
    for p in wanted:
        if not data_arg[p]:
            flag = {"pixel11": "--px-data", "glxz8": "--z8-data", "iphone18": "--ip18-data"}[p]
            ap.error(f"{flag} is required when sending {p} (--product includes it)")

    today = datetime.date.fromisoformat(a.date) if a.date else datetime.date.today()
    cards, mods = {}, {}
    for p in wanted:
        folder, _, _ = PRODUCTS[p]
        mod = _load(f"{SKILLS}/{folder}/report.py", f"{p}rep")
        mods[p] = mod
        cards[p] = mod.build_card(json.loads(data_arg[p]), today)

    full_broadcast = set(wanted) == set(PRODUCTS)

    if a.test_only:
        targets = [TEST_ROOM]
    else:
        targets = ROOMS
        if a.only:
            subs = [s.strip() for s in a.only.split(",") if s.strip()]
            targets = [r for r in targets if any(s in r["name"] for s in subs)]
        if a.exclude:
            subs = [s.strip() for s in a.exclude.split(",") if s.strip()]
            targets = [r for r in targets if not any(s in r["name"] for s in subs)]

    if not targets:
        print("no matching rooms", file=sys.stderr)
        sys.exit(1)

    print(f"date={today}  rooms={len(targets)}  product={','.join(wanted)}  "
          f"({'combined carousel, 1 msg/room' if full_broadcast else f'{len(wanted)} card(s) each'})\n")

    if full_broadcast:
        header_imgs = {p: mods[p].HEADER_IMG for p in wanted}
        message = carousel.build_carousel_message(cards, header_imgs, today)
        for room in targets:
            url = BASE.format(sid=room["sid"], tok=room["token"])
            print(f"[{room['name']}] carousel : {_post(url, message)}")
            time.sleep(1.0)
    else:
        for room in targets:
            for p in PRODUCTS:  # fixed send order: glxz8, iphone18, pixel11
                if p not in cards:
                    continue
                _, _, label = PRODUCTS[p]
                print(f"[{room['name']}] {label} : {_post(room_url(room, p), cards[p])}")
                time.sleep(1.0)


if __name__ == "__main__":
    main()
