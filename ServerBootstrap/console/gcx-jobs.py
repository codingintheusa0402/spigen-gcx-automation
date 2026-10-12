#!/usr/bin/env python3
"""Running + upcoming jobs on gcx-server for the GCX console. Prints: run|name|detail|logkey  and  next|name|detail
(logkey → gcx-logview.sh, opened when the row is clicked)"""
import datetime as dt, json, os, re, subprocess
from zoneinfo import ZoneInfo
from croniter import croniter

KST = ZoneInfo("Asia/Seoul")
NAMES = [  # (regex on command line, friendly name)
    (r"auto_broadcast\.py --catchup", "Bad-review catch-up"),
    (r"auto_broadcast\.py --retry-if-held", "Bad-review retry"),
    (r"auto_broadcast\.py", "Bad-review broadcast"),
    (r"CaspiSalesBackfill|backfill\.py", "Caspi 판매량 backfill"),
    (r"DiscolorationReport|report\.py", "이염/변색 report"),
    (r"scrape_sc_reviews\.py", "SC scraper"),
    (r"propagate\.py", "SC propagate"),
    (r"phase_g\.py", "Photo check (Phase G)"),
    (r"fill_asin_blanks|scheduler_gate", "SKU/ASIN filler"),
    (r"gcx-autosync", "Git auto-sync"),
    (r"health_check(_gate)?\.py", "Apify scraper health check"),
    (r"gcx_watchdog\.py", "Server health check (watchdog)"),
    (r"gcx_selfheal\.sh", "Self-heal check"),
    (r"gcx_heartbeat\.py", "Heartbeat"),
    (r"daily_cron_gate\.py", "SC propagate (daily gate)"),
    (r"zendesk_session\.py", "Zendesk login check"),
    (r"tmux_start", "Start tmux sessions"),
]
QUIET = {"Bad-review catch-up", "Git auto-sync", "Apify scraper health check", "Server health check (watchdog)", "Self-heal check", "Heartbeat"}          # every few minutes — not worth listing as "up next"

def name_of(cmd):
    for rx, n in NAMES:
        if re.search(rx, cmd):
            return n
    return None

now = dt.datetime.now(KST)
H = os.path.expanduser("~")


def sh(*a):
    return subprocess.run(list(a), capture_output=True, text=True).stdout.strip()


def since(ts):
    m = max(0, int((now - ts).total_seconds() // 60))
    return f"{m // 1440}d {m % 1440 // 60}h" if m >= 1440 else (f"{m // 60}h {m % 60}m" if m >= 60 else f"{m}m")


def unit_up(unit, user=True):
    """'up 3h 12m' for an active systemd unit, else None."""
    a = ["systemctl"] + (["--user"] if user else []) + ["show", unit, "-p", "ActiveState", "-p", "ActiveEnterTimestamp"]
    kv = dict(l.split("=", 1) for l in sh(*a).splitlines() if "=" in l)
    if kv.get("ActiveState") != "active":
        return None
    try:
        t = dt.datetime.strptime(kv["ActiveEnterTimestamp"].rsplit(" ", 1)[0], "%a %Y-%m-%d %H:%M:%S").replace(tzinfo=KST)
        return f"up {since(t)}"
    except Exception:
        return "up"


def ago(path):
    try:
        return since(dt.datetime.fromtimestamp(os.path.getmtime(path), KST)) + " ago"
    except Exception:
        return "never"


# ---- always-on services (listed first) ----
up = unit_up("mock-trader")
if up:
    try:
        plans = json.load(open(f"{H}/mock-trader/state/app/engine.json")).get("plans", {})
        act = [p for p in plans.values() if p.get("status") not in ("done", "merged", "halted", "settled")]
        real = sum(1 for p in act if p.get("mode") == "real")
        detail = f"{up} · {len(act)} plan(s) active ({real} real · {len(act) - real} mock)"
    except Exception:
        detail = up
    print(f"run|AI Trader engine|{detail}|svc:mock-trader:user")
elif sh("systemctl", "--user", "is-active", "mock-trader") == "activating":
    print("run|AI Trader engine|waiting — the Mac fallback is running it|svc:mock-trader:user")
if subprocess.run(["curl", "-s", "-m", "2", "-o", "/dev/null", "http://127.0.0.1:9222/json/version"]).returncode == 0:
    print("run|Zendesk Chrome (for ticket monitor)|open · CDP 9222|tmux:zendesk-chrome")
up = unit_up("gcx-relogin", user=False)
if up:
    print(f"run|Re-login page (Tailscale)|{up}|svc:gcx-relogin:system")
up = unit_up("gcx-owner-watch")
if up:
    print(f"run|Mac hand-over guard|{up}|svc:gcx-owner-watch:user")
up = unit_up("gcx-mesh")
if up:
    print(f"run|GCX Mesh for iPhone (server copy)|{up}|svc:gcx-mesh:user")
print("run|Health checks|watchdog " + ago(f"{H}/.gcx-watchdog.json") + " · heartbeat " + ago(f"{H}/.gcx-heartbeat.json")
      + " · self-heal every 2m · git sync " + ago(f"{H}/.gcx-autosync.heartbeat") + "|health")

# running: only real job processes (python/bash actually executing the job script) —
# not Claude sessions whose prompt merely mentions a script name
ps = subprocess.run(["ps", "-eo", "pid=,etimes=,args="], capture_output=True, text=True).stdout.splitlines()
seen = set()
for line in ps:
    pid, secs, args = (line.strip().split(None, 2) + ["", ""])[:3]
    argv = args.split()
    if not argv or os.path.basename(argv[0]) not in ("python3", "python", "bash", "sh"):
        continue
    script = next((a for a in argv[1:] if not a.startswith("-")), "")      # first non-flag arg = the script
    n = name_of(script) if script.endswith((".py", ".sh")) else None
    if n and n not in seen:
        seen.add(n); m = int(secs) // 60
        print((f"run|{n}|running {m // 60}h {m % 60}m" if m >= 60 else f"run|{n}|running {max(m, 0)}m") + f"|pid:{pid}")

# Claude task sessions: listed only while actually working (status busy); the ticket monitor is a standing loop
cmd = {l.strip().split(None, 2)[0]: l.strip().split(None, 2)[2] for l in ps if len(l.split(None, 2)) == 3}
sess_dir = os.path.expanduser("~/.claude/sessions")
for f in os.listdir(sess_dir) if os.path.isdir(sess_dir) else []:
    if not f.endswith(".json"):
        continue
    try:
        d = json.load(open(os.path.join(sess_dir, f)))
    except Exception:
        continue
    pid = str(d.get("pid") or f[:-5]); args = cmd.get(pid, "")
    if not args.startswith("claude"):
        continue                                   # process gone
    m = re.search(r"--remote-control (gcx-[\w-]+)", args)
    label = m.group(1)[4:] if m else (d.get("name") or "session")
    if "ticket-monitor" in label:
        print("run|Ticket monitor (Claude loop)|" + ("checking now" if d.get("status") == "busy" else "watching · every minute") + "|tmux:ticket-monitor")
    elif d.get("status") == "busy":
        print(f"run|Claude: {label}|working|" + ("tmux:main" if label == "main" else f"pid:{pid}"))
    elif label == "main":
        print("run|Claude: main (Remote Control)|idle · ready|tmux:main")

# upcoming: next fire time of each cron line
cron = subprocess.run(["crontab", "-l"], capture_output=True, text=True).stdout.splitlines()
ups = []
for l in cron:
    l = l.strip()
    if not l or l.startswith("#") or l.startswith("@") or "=" in l.split()[0]:
        continue
    parts = l.split(None, 5)
    if len(parts) < 6:
        continue
    n = name_of(parts[5])
    if not n or n in QUIET:
        continue
    nxt = croniter(" ".join(parts[:5]), now).get_next(dt.datetime)
    ups.append((nxt, n))
best = {}
for t, n in sorted(ups):
    best.setdefault(n, t)
for n, t in sorted(best.items(), key=lambda x: x[1])[:30]:
    mins = int((t - now).total_seconds() // 60)
    left = f"in {mins // 60}h {mins % 60}m" if mins >= 60 else f"in {mins}m"
    day = "" if t.date() == now.date() else ("tomorrow " if (t.date() - now.date()).days == 1 else t.strftime("%a "))
    print(f"next|{n}|{day}{t.strftime('%H:%M')} · {left}")
