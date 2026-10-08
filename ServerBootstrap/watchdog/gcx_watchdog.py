#!/usr/bin/env python3
"""GCX Server watchdog — cron every 5 min on gcx-server.

Runs the same checks as the console's HEALTH & STATUS cards and alerts the GCX Server Chat room
once when a check goes bad (2 checks in a row, to ignore blips) and once when it recovers.
If the alert can't be sent (e.g. internet down) the state isn't saved, so it is retried next run.
"""
import json, os, re, subprocess, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gcx_alert import alert  # noqa: E402

STATE = os.path.expanduser("~/.gcx-watchdog.json")
BATTERY_STEPS = [70, 50, 30, 20, 10, 5, 1]       # alert once as the battery falls past each of these
HEALTH = os.path.expanduser("~/Desktop/GCX/ServerBootstrap/console/gcx-health.sh")
PS = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
LABEL = {"power": "Power", "tailscale": "Tailscale", "remote": "VNC / SSH", "disk": "Disk",
         "wsl": "Ubuntu", "claude": "Claude", "monitor": "Ticket monitor", "cron": "Scheduled jobs",
         "git": "Git sync", "sheets": "Google Sheets"}
HINT = {
    "power": "Plug the laptop in — on battery it will switch off when the battery runs out.",
    "monitor": "Open the GCX console → GCX Live → ticket-monitor tab, or ask Claude to restart the ticket monitor on the server.",
    "claude": "Claude Code is logged out — on the server run `claude` once (GCX console → New Claude Session) and log in.",
    "git": "The Mac and the server edited the same lines. Ask Claude: \"resolve the GCX git sync conflict\".",
    "cron": "The scheduled jobs stopped — ask Claude to check cron on the server.",
    "remote": "VNC or SSH service stopped — remote access may be lost.",
    "tailscale": "Tailscale is down — the Mac/iPhone can't reach the server.",
}


def windows_checks():
    ps = r'''
$b = Get-CimInstance Win32_Battery -EA SilentlyContinue
if ($b) { $p = $b.EstimatedChargeRemaining
  if ($b.BatteryStatus -ne 1) { "power|ok|plugged in ($p%)" } elseif ($p -le 25) { "power|bad|ON BATTERY $p%" } else { "power|warn|on battery $p%" } }
$v = (Get-Service tvnserver -EA SilentlyContinue).Status; $s = (Get-Service sshd -EA SilentlyContinue).Status
if ($v -eq "Running" -and $s -eq "Running") { "remote|ok|VNC + SSH running" } else { "remote|bad|VNC $v, SSH $s" }
$f = [math]::Round((Get-PSDrive C).Free / 1GB); if ($f -lt 10) { "disk|bad|C: $f GB free" } else { "disk|ok|C: $f GB free" }
try { $j = & "C:\Program Files\Tailscale\tailscale.exe" status --json | ConvertFrom-Json
      if ($j.BackendState -eq "Running") { "tailscale|ok|connected" } else { "tailscale|bad|$($j.BackendState)" } } catch { "tailscale|bad|not responding" }
'''
    try:
        out = subprocess.run([PS, "-NoProfile", "-Command", ps], capture_output=True, text=True, timeout=60).stdout
    except Exception as e:
        return [f"remote|warn|Windows checks failed: {e}"]
    return out.splitlines()


def main():
    lines = windows_checks()
    lines += subprocess.run(["bash", HEALTH], capture_output=True, text=True, timeout=60).stdout.splitlines()
    now = {}
    for l in lines:
        p = l.strip().split("|", 2)
        if len(p) == 3 and p[0] in LABEL:
            now[p[0]] = (p[1], p[2])
    try:
        st = json.load(open(STATE))
    except Exception:
        st = {}
    changed = False
    # power: alert only when the battery drops past a threshold (not on every unplug) — user rule 2026-10-08
    if "power" in now:
        lvl, txt = now.pop("power")
        pp = st.get("power", {"sent": [], "level": "ok"})
        m = re.search(r"(\d+)%", txt); pct = int(m.group(1)) if m else None
        sent = pp.get("sent", [])
        try:
            if lvl == "ok":                                   # plugged back in
                if sent:
                    alert("Power: plugged in again", txt, level="ok")
                sent = []
            elif pct is not None:
                due = [t for t in BATTERY_STEPS if pct <= t and t not in sent]
                if due:
                    t = min(due)
                    alert(f"Power: battery {pct}% (on battery)", HINT["power"], level="bad" if t <= 20 else "warn")
                    sent = sorted(set(sent) | set(due), reverse=True)
        except Exception as e:
            print(f"alert failed for power: {e}")
        st["power"] = {"level": lvl, "sent": sent, "text": txt, "alerted": "bad" if sent else "ok", "streak": 0}
        changed = True
    for k, (lvl, txt) in now.items():
        prev = st.get(k, {"level": "ok", "alerted": "ok", "streak": 0})
        streak = prev["streak"] + 1 if lvl == prev["level"] else 1
        need = 2                                            # 2 checks in a row, to ignore blips
        alerted = prev["alerted"]
        try:
            if lvl in ("bad", "warn") and alerted != lvl and streak >= need:
                alert(f"{LABEL[k]}: {txt}", HINT.get(k, ""), level=lvl); alerted = lvl
            elif lvl == "ok" and alerted in ("bad", "warn"):
                alert(f"{LABEL[k]} recovered", txt, level="ok"); alerted = "ok"
        except Exception as e:
            print(f"alert failed for {k}: {e}")                # keep old 'alerted' → retried next run
        st[k] = {"level": lvl, "alerted": alerted, "streak": streak, "text": txt}
        changed = True
    if changed:
        json.dump(st, open(STATE, "w"), indent=1)
    if "-v" in sys.argv:
        for k, v in st.items():
            print(f"{k:10} {v['level']:5} alerted={v['alerted']:5} {v['text']}")


if __name__ == "__main__":
    main()
