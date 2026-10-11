#!/usr/bin/env python3
"""Who runs the always-on services right now — gcx-server or the Mac (fallback)?

Services: `aitrader` (AI Trader engine — real money) and `ticket-monitor` (Zendesk ticket-reporter loop).
Exactly one machine may run each. The owner is a row `<svc>:owner` in the shared run ledger
(gcx_ledger.py); the Mac sets it to `mac` when it takes over and refreshes it every 5 min, and sets it back
to `server` after it has stopped its copy and pushed the state back. A `mac` owner older than LEASE_MIN
counts as expired (the Mac went to sleep / offline), so the server never waits forever.

Server side
  gcx_owner.py gate <svc>      block while the Mac owns <svc> (systemd ExecStartPre / tmux_start.sh)
  gcx_owner.py watch           loop (systemd --user gcx-owner-watch): stop the server's copy while the Mac
                               owns it, start it again once the Mac hands back
Mac side
  gcx_owner.py wrap <svc> -- cmd …   run cmd; pause it the moment the Mac wakes from sleep and only resume
                               if the Mac still owns <svc> — otherwise stop it (the server may be trading)
"""
import json, os, signal, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import gcx_ledger as L  # noqa: E402

LEASE_MIN = 15
HB_DOWN_MIN = 7
SERVICES = ("aitrader", "ticket-monitor")
MARK = os.path.expanduser("~/.gcx-owner")          # server: "<svc>.stopped" = stopped because the Mac owns it


def log(m):
    print(f"{time.strftime('%m-%d %H:%M:%S')} [owner] {m}", flush=True)


def owner(svc):
    """('mac'|'server', age_min). Raises if the ledger can't be read."""
    r = L.get(f"{svc}:owner")
    if not r:
        return "server", 0.0
    return (r["data"] or {}).get("owner", "server"), (time.time() * 1000 - r["updated"]) / 60000


def mac_owns(svc):
    o, age = owner(svc)
    return o == "mac" and age < LEASE_MIN


def set_owner(svc, who):
    return L.put(f"{svc}:owner", "done", {"owner": who})


def heartbeat_age_min():
    row = L._svc().values().get(spreadsheetId=L._sid(), range="heartbeat!A2:A2").execute()["values"][0]
    return (time.time() * 1000 - int(row[0])) / 60000


# ---------------- server side ----------------
TMUX_TM = ('cd ~ && export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0; '
           'claude --remote-control gcx-ticket-monitor "$(cat ~/.gcx-tasks/ticket-monitor.prompt)"; '
           'echo; echo [session ended]; exec bash')


def _sh(cmd):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True)


def server_running(svc):
    if svc == "aitrader":
        return _sh("systemctl --user is-active mock-trader").stdout.strip() in ("active", "activating")
    return "ticket-monitor" in _sh("tmux list-windows -t gcx -F '#W' 2>/dev/null").stdout.split()


def server_stop(svc):
    if svc == "aitrader":
        _sh("systemctl --user stop mock-trader")
    else:
        _sh("tmux kill-window -t gcx:ticket-monitor")


def server_start(svc):
    if svc == "aitrader":
        _sh("systemctl --user start --no-block mock-trader")
    elif os.path.exists(os.path.expanduser("~/.gcx-tasks/ticket-monitor.prompt")):
        _sh("tmux has-session -t gcx 2>/dev/null || tmux new-session -d -s gcx")
        subprocess.run(["tmux", "new-window", "-d", "-t", "gcx", "-n", "ticket-monitor", TMUX_TM])


def gate(svc):
    t0 = time.time()
    while True:
        try:
            if not mac_owns(svc):
                return 0
            log(f"{svc}: the Mac is running it (fallback) — waiting for hand-back")
        except Exception as e:
            if time.time() - t0 > 600:
                log(f"{svc}: ledger unreachable for 10 min ({e}) — starting anyway")
                return 0
        time.sleep(20)


def watch():
    os.makedirs(MARK, exist_ok=True)
    while True:
        for svc in SERVICES:
            try:
                mine = not mac_owns(svc)
            except Exception:
                continue                                    # can't tell — leave things as they are
            mark = os.path.join(MARK, f"{svc}.stopped")
            if not mine and server_running(svc):
                log(f"{svc}: Mac owns it → stopping the server copy")
                server_stop(svc); open(mark, "w").close()
            elif mine and os.path.exists(mark):
                log(f"{svc}: handed back → starting the server copy")
                server_start(svc); os.remove(mark)
        time.sleep(15)


# ---------------- Mac side ----------------
def wrap(svc, cmd):
    child = subprocess.Popen(cmd)
    stopped = False

    def finish(*_):
        try:
            os.kill(child.pid, signal.SIGCONT)
            child.terminate(); child.wait(30)
        except Exception:
            child.kill()
        sys.exit(0)
    signal.signal(signal.SIGTERM, finish); signal.signal(signal.SIGHUP, finish); signal.signal(signal.SIGINT, finish)

    last = lastcheck = time.time()
    while child.poll() is None:
        time.sleep(2)
        now = time.time()
        woke = now - last > 30                              # the Mac just woke up (or froze)
        last = now
        if woke and not stopped:
            os.kill(child.pid, signal.SIGSTOP); stopped = True
            log(f"{svc}: Mac woke up — paused until ownership is confirmed")
        if not (stopped or now - lastcheck > 60):
            continue
        lastcheck = now
        try:
            o, age = owner(svc)
            hb = heartbeat_age_min()
        except Exception as e:
            log(f"{svc}: can't check ownership ({e})"); continue      # stays paused if paused
        if o != "mac" or (hb <= HB_DOWN_MIN and age >= LEASE_MIN):
            log(f"{svc}: no longer the owner (owner={o}, lease {age:.0f} min, server heartbeat {hb:.0f} min) — stopping")
            finish()
        if stopped:
            os.kill(child.pid, signal.SIGCONT); stopped = False
            log(f"{svc}: still the owner — resumed")
    sys.exit(child.returncode)


if __name__ == "__main__":
    a = sys.argv[1:]
    if a[:1] == ["gate"]:
        sys.exit(gate(a[1]))
    if a[:1] == ["watch"]:
        watch()
    if a[:1] == ["wrap"]:
        wrap(a[1], a[a.index("--") + 1:])
    if a[:1] == ["status"]:
        for s in SERVICES:
            print(s, *owner(s))
