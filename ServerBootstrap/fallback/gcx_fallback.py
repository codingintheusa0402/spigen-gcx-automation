#!/usr/bin/env python3
"""Mac fallback for the gcx-server scheduled jobs — Mac launchd com.spigen.gcx.server-fallback, every 5 min.

When the server misses its heartbeat (sheet, every 2 min) for > 7 min AND can't be reached over Tailscale
on 2 checks in a row, the Mac takes over: it runs the server's jobs below on the same schedule until the
heartbeat comes back, then hands back. Takeover / hand-back are posted to the GCX Server room.

Double sends are prevented by the shared run ledger (gcx_ledger.py): every job records its real runs there,
and the Mac only takes over once the server's heartbeat says its jobs use it ("ledger:1").

  gcx_fallback.py            # one tick (launchd)
  gcx_fallback.py --status   # print state
"""
import datetime, json, os, re, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
G = os.path.abspath(os.path.join(HERE, "..", ".."))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(G, "ServerBootstrap", "watchdog"))
import gcx_ledger  # noqa: E402

STATE = os.path.expanduser("~/.gcx-fallback.json")
RUN_DIR = os.path.expanduser("~/.gcx-fallback")            # pid files + per-job logs
PY = "/opt/homebrew/bin/python3"
DOWN_AFTER_MIN = 7                                        # heartbeat every 2 min → 3 missed
STREAK = 2                                                # consecutive bad checks before taking over
MAX_WINDOW_MIN = 10                                       # after a Mac sleep, don't replay hours of slots
SSH = [["ssh", "-o", "ConnectTimeout=8", "-o", "BatchMode=yes", "kevinkim@gcx-server", "true"],
       ["ssh", "-i", os.path.expanduser("~/.ssh/id_ed25519_gcx_server"), "-o", "ConnectTimeout=8",
        "-o", "BatchMode=yes", "user@claude-server", "echo ok"]]

# Same jobs + times as the server crontab (ServerBootstrap/crontab.txt). NOT here on purpose: git-autosync
# (the Mac has its own), watchdog/heartbeat (server-only), mock-trader (real-money trading engine).
JOBS = [
    ("BadReview broadcast",         "30 10 * * 1-5", "GAS_Operations/BadReview_ChatReport", ["auto_broadcast.py"]),
    ("BadReview retry (11:30)",     "30 11 * * 1-5", "GAS_Operations/BadReview_ChatReport", ["auto_broadcast.py", "--retry-if-held"]),
    ("BadReview catch-up",          "*/5 * * * *",   "GAS_Operations/BadReview_ChatReport", ["auto_broadcast.py", "--catchup"]),
    ("Caspi 판매량 backfill",        "*/30 * * * *",  "GAS_Operations/CaspiSalesBackfill",   ["backfill.py"]),
    ("Discoloration report",        "*/30 * * * *",  "GAS_Operations/DiscolorationReport",  ["report.py"]),
    ("Apify ASIN health check",     "0 10 * * *",    "Scrapers/apify-asin-rating-scraper",  ["health_check.py", "--autofix"]),
]


def log(msg):
    print(f"{time.strftime('%m-%d %H:%M:%S')} {msg}", flush=True)


# ---------- cron matching ----------
def _field(spec, val, lo, hi):
    for part in spec.split(","):
        rng, step = (part.split("/") + ["1"])[:2]
        a, b = (lo, hi) if rng == "*" else (map(int, rng.split("-")) if "-" in rng else (int(rng), int(rng)))
        if a <= val <= b and (val - a) % int(step) == 0:
            return True
    return False


def cron_match(expr, t):
    mi, h, dom, mon, dow = expr.split()
    cdow = (t.weekday() + 1) % 7                           # cron: 0=Sun
    return (_field(mi, t.minute, 0, 59) and _field(h, t.hour, 0, 23) and _field(dom, t.day, 1, 31)
            and _field(mon, t.month, 1, 12) and (_field(dow, cdow, 0, 7) or (cdow == 0 and _field(dow, 7, 0, 7))))


def due(expr, since, now):
    t = since.replace(second=0, microsecond=0) + datetime.timedelta(minutes=1)
    while t <= now:
        if cron_match(expr, t):
            return True
        t += datetime.timedelta(minutes=1)
    return False


# ---------- helpers ----------
def alert(title, text, level):
    try:
        from gcx_alert import alert as _a
        _a(title, text, level=level, host="Mac (fallback)")
    except Exception as e:
        log(f"alert failed: {e}")


def server_reachable():
    for cmd in SSH:
        try:
            if subprocess.run(cmd, capture_output=True, timeout=20).returncode == 0:
                return True
        except Exception:
            pass
    return False


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def running(name):
    pf = os.path.join(RUN_DIR, f"{slug(name)}.pid")
    try:
        os.kill(int(open(pf).read()), 0)
        return True
    except Exception:
        return False


def launch(name, cwd, args):
    os.makedirs(RUN_DIR, exist_ok=True)
    out = open(os.path.join(RUN_DIR, f"{cwd.split('/')[-1]}.log"), "a")
    out.write(f"\n=== {time.strftime('%F %T')} fallback run: {name} ===\n"); out.flush()
    env = dict(os.environ, PATH=f"{os.path.expanduser('~/.local/bin')}:/opt/homebrew/bin:/usr/bin:/bin")
    p = subprocess.Popen([PY] + args, cwd=os.path.join(G, cwd), stdout=out, stderr=subprocess.STDOUT,
                         stdin=subprocess.DEVNULL, env=env, start_new_session=True)
    open(os.path.join(RUN_DIR, f"{slug(name)}.pid"), "w").write(str(p.pid))


def main():
    try:
        st = json.load(open(STATE))
    except Exception:
        st = {"mode": "normal", "streak": 0, "ran": []}
    if "--status" in sys.argv:
        print(json.dumps(st, indent=1, ensure_ascii=False)); return
    now = datetime.datetime.now()

    try:
        row = gcx_ledger._svc().values().get(spreadsheetId=gcx_ledger._sid(), range="heartbeat!A2:E2").execute()["values"][0]
    except Exception as e:
        log(f"can't read heartbeat (Mac offline?): {e}")       # nothing safe to do without the ledger
        return
    age_min = (time.time() * 1000 - int(row[0])) / 60000
    armed = len(row) > 3 and "ledger" in row[3]
    down = age_min > DOWN_AFTER_MIN and not server_reachable()
    st["streak"] = st.get("streak", 0) + 1 if down else 0
    st["heartbeat_age_min"] = round(age_min, 1)

    if st["mode"] == "normal" and st["streak"] >= STREAK:
        if armed:
            st.update(mode="fallback", since=now.isoformat(timespec="minutes"), ran=[], last_tick=(now - datetime.timedelta(minutes=5)).isoformat())
            log("server down → taking over")
            alert("Server down — Mac took over the scheduled jobs",
                  f"No heartbeat from gcx-server for <b>{age_min:.0f} min</b> and it can't be reached over Tailscale.<br>"
                  f"Until it's back, this Mac runs: {', '.join(j[0] for j in JOBS)}.<br>"
                  "Nothing is sent twice — every run is checked against the shared run ledger.", "warn")
        elif not st.get("unarmed_alerted"):
            st["unarmed_alerted"] = True
            alert("Server down — Mac fallback NOT armed",
                  "The server is unreachable, but it hasn't yet run the version of the jobs that records runs in the "
                  "shared ledger, so the Mac won't take over (it could double-send). It arms itself automatically "
                  "the next time the server is online for a few minutes.", "bad")
    elif st["mode"] == "fallback" and age_min <= DOWN_AFTER_MIN:
        ran = st.get("ran", [])
        log("server back → handing back")
        alert("Server back — Mac handed the jobs back",
              f"Mac covered from {st.get('since', '?')[-5:]} to {now:%H:%M}. "
              + (f"Runs started on the Mac: {', '.join(ran[-20:])}." if ran else "No job was due in that time."), "ok")
        st.update(mode="normal", ran=[])
    if armed:
        st.pop("unarmed_alerted", None)

    if st["mode"] == "fallback":
        last = datetime.datetime.fromisoformat(st.get("last_tick") or now.isoformat())
        since = max(last, now - datetime.timedelta(minutes=MAX_WINDOW_MIN))
        for name, expr, cwd, args in JOBS:
            if due(expr, since, now) and not running(name):
                launch(name, cwd, args)
                log(f"started {name}")
                if "catch-up" not in name and "backfill" not in name and "Discoloration" not in name:
                    st.setdefault("ran", []).append(f"{now:%H:%M} {name}")
                elif not any(name in r for r in st.get("ran", [])):
                    st.setdefault("ran", []).append(f"{now:%H:%M} {name} (every few min while down)")
    st["last_tick"] = now.isoformat()
    json.dump(st, open(STATE, "w"), indent=1, ensure_ascii=False)


if __name__ == "__main__":
    main()
