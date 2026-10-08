#!/usr/bin/env python3
"""Running + upcoming jobs on gcx-server for the GCX console. Prints: run|name|detail  and  next|name|detail"""
import datetime as dt, os, re, subprocess
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
    (r"tmux_start", "Start tmux sessions"),
]
QUIET = {"Bad-review catch-up", "Git auto-sync"}          # every few minutes — not worth listing as "up next"

def name_of(cmd):
    for rx, n in NAMES:
        if re.search(rx, cmd):
            return n
    return None

now = dt.datetime.now(KST)

# running: job processes + Claude loop sessions
ps = subprocess.run(["ps", "-eo", "etimes=,args="], capture_output=True, text=True).stdout.splitlines()
seen = set()
for line in ps:
    secs, _, args = line.strip().partition(" ")
    if "python3" not in args and "bash" not in args:
        continue
    n = name_of(args)
    if n and n not in seen and "ps -eo" not in args:
        seen.add(n); m = int(secs) // 60
        print(f"run|{n}|running {m // 60}h {m % 60}m" if m >= 60 else f"run|{n}|running {m}m")
for line in ps:
    m = re.search(r"claude --remote-control (gcx-[\w-]+)", line)
    if m and "ticket-monitor" in m.group(1):
        print("run|Ticket monitor (Claude loop)|checks Zendesk every minute")

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
for n, t in sorted(best.items(), key=lambda x: x[1])[:6]:
    mins = int((t - now).total_seconds() // 60)
    left = f"in {mins // 60}h {mins % 60}m" if mins >= 60 else f"in {mins}m"
    day = "" if t.date() == now.date() else ("tomorrow " if (t.date() - now.date()).days == 1 else t.strftime("%a "))
    print(f"next|{n}|{day}{t.strftime('%H:%M')} · {left}")
