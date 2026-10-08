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

KEY = "AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI"
BASE = "https://chat.googleapis.com/v1/spaces/{sid}/messages?key=" + KEY + "&token={tok}"

TEST_ROOM = {"name": "TEST (실장님&GCX test)", "sid": "AAQAc9NQmJQ",
             "token": "T_rTrPKTYq6biglb8kRL3GOVfQg3AAOH-JPKELutbAY"}

# Broadcast targets — the rooms the user enumerated for "share the bad review scraped data".
# (Source of truth: memory gcx_team_gchat_webhooks.md. Keep in sync.)
# `token` = default webhook (used by the Pixel 11 card, and by Z8/iPhone18 where no override).
# `glxz8` = OPTIONAL per-room override token used for BOTH the Galaxy Z8 card AND (per
# explicit user instruction, 2026-09-21) the iPhone 18 card — same room, one separate
# incoming webhook shared by both products rather than a third dedicated one.
ROOMS = [
 {"name": "GCX전략 x SDA (아마존직판)", "sid": "AAAAe96DDIs",
  "token": "fqRQJsNX1O8LDUyjmFsKdBJo5VCPXuFA2KX2OfAuLXk",
  "glxz8": "k5bXiLIJgrjaLdjzyQR1ae0QhNjYble8eLgDGnxnFUc"},
 {"name": "GCX전략 x ADS1", "sid": "AAAAFjWzr40",
  "token": "Xv5J3ipKs_mIOem7OMHzhhmPwHcTrC-wDgYlMSZHAzs",
  "glxz8": "O6gHRrVCB3-X31BJgpUTncKObq2tCLEdk4cMXmwLiR0"},
 {"name": "GCX전략 x ADS2", "sid": "AAAATOmW7HU",
  "token": "GsprARTa_2ga2mkdz8lEFe2K4DTTRl7zfpBW6qEvlOU",
  "glxz8": "sgex588AiCc0FqI_XOEEHK-lOsWhRlw8QdVKkcDsM9g"},
 {"name": "GCX전략 x ADS3", "sid": "AAAAKwBoZPU",
  "token": "zW6cEhLwMozY2v3DvH9nvk4eFW8kSpwlx3MFLYbFFBE",
  "glxz8": "IdYKJJunaaQKTzqFBU2BJaVAASk5MgIdLpzTyd0HC_0"},
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
 # 리더들방 — one webhook for every card (no per-product override).
 {"name": "리더들방", "sid": "AAAALqHjZHo",
  "token": "Z8696aOLrUlGk7zUUgjNywb8PIMbmKqrKaQd5JnofT0"},
 # GCX전략 Spigen x TCK — added 2026-09-29, no per-product override (same as 리더들방).
 {"name": "GCX전략 Spigen x TCK", "sid": "AAAAlxfKOYE",
  "token": "TtxdDkyX3Q5QGgKRWGKeedQk5R9z8IQ4mn_4m5oLhu4"},
]

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
