#!/usr/bin/env python3
"""
One-off for 2026-10-16: user flies ~10:30-11:30 KST, no wifi on Mac or server during
that window (and it could start earlier than 10:30 — boarding/airplane mode timing is
not guaranteed). Rather than waiting for a fixed clock time, poll the 3 BadReview
sheets every 2 min starting now and fire the REAL auto_broadcast.py the moment all of
today's rows have a 인입사유(tag) filled in — i.e. send as early as humanly possible
once the data is actually ready, to minimize exposure to an unknown wifi cutoff.
Hard deadline 10:20 (the user's own latest estimate for when data finishes crunching):
past that, send anyway regardless of unfilled rows — auto_broadcast.py's own
fetch_ready() retry logic handles that case the same way the normal 10:30 run would.
Ledger-gated by auto_broadcast.py itself, so no double-send risk vs. the regular
10:30 trigger (server cron or Mac fallback) if that also fires later.
"""
import datetime, importlib.util, os, sys, time

AB_DIR = "/Users/kevinkim/Desktop/GCX/GAS_Operations/BadReview_ChatReport"
spec = importlib.util.spec_from_file_location("auto_broadcast", os.path.join(AB_DIR, "auto_broadcast.py"))
ab = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ab)

DEADLINE = datetime.datetime.now().replace(hour=10, minute=20, second=0, microsecond=0)
POLL_S = 120


def log(msg):
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] {msg}", flush=True)


def ready():
    today = datetime.date.today()
    token = ab.refresh_gws_token()
    for name, sid in ab.SHEETS.items():
        header, body = ab.fetch_rows(token, sid)
        unfilled = ab.count_unfilled_today(header, body, today)
        if unfilled:
            log(f"{name}: {unfilled} today-row(s) still missing 인입사유(tag)")
            return False
    return True


log("poller starting — will trigger the real send as soon as all 3 sheets are ready, hard deadline 10:20")
while True:
    now = datetime.datetime.now()
    try:
        if ready():
            log("all 3 sheets ready — triggering real send now")
            break
    except Exception as e:
        log(f"readiness check failed ({e}) — will retry in {POLL_S}s")
    if now >= DEADLINE:
        log("hit 10:20 deadline — triggering send anyway (auto_broadcast.py's own retry "
            "logic covers any rows still unfilled)")
        break
    time.sleep(POLL_S)

os.chdir(AB_DIR)
os.execvp("/opt/homebrew/bin/python3", ["/opt/homebrew/bin/python3", "auto_broadcast.py"])
