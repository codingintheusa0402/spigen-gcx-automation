"""Always-on services the Mac takes over when gcx-server is down (used by gcx_fallback.py).

  aitrader        AI Trader engine (~/mock-trader, REAL MONEY when armed). The Mac runs a copy of the code in
                  ~/.gcx-fallback/aitrader with the server's state (mirrored every 5 min while the server is up).
  ticket-monitor  ticket-reporter Monitor mode as a Claude session (claude-in-chrome, Chrome Profile 9), with the
                  server's state/processed.json so no ticket is reported twice.

Only one machine runs each service: ownership is in the ledger (gcx_owner.py). Hand-back = stop the Mac copy →
push its state to the server (while the server copy is held stopped by gcx-owner-watch) → owner=server.
"""
import glob, json, os, shutil, subprocess, time

import gcx_owner as O
from gcx_owner import SERVICES  # noqa: F401  (used by gcx_fallback.py)

FB = os.path.expanduser("~/.gcx-fallback")
MIRROR = os.path.join(FB, "mirror")
MT = os.path.expanduser("~/mock-trader")
AIT = os.path.join(FB, "aitrader")
SKILL_STATE = os.path.expanduser("~/.claude/skills/ticket-reporter/state")
HERE = os.path.dirname(os.path.abspath(__file__))
SSH = ["ssh", "-o", "ConnectTimeout=10", "-o", "BatchMode=yes", "kevinkim@gcx-server"]
RSH = "ssh -o ConnectTimeout=10 -o BatchMode=yes"
MAX_MIRROR_AGE = {"aitrader": 30, "ticket-monitor": 60}    # minutes; older → don't take over (could double-act)
SESSION = {"aitrader": "gcx-fb-aitrader", "ticket-monitor": "gcx-fb-ticket"}
LABEL = {"aitrader": "AI Trader engine", "ticket-monitor": "Zendesk ticket monitor"}
TICKET_PROMPT = ("/loop 1m Run the ticket-reporter skill in Monitor mode. You are the MAC FALLBACK monitor — the GCX "
                 "server is down. Use claude-in-chrome with Chrome Profile 9 as the skill says. state/processed.json "
                 "was just synced from the server: follow the skill's dedup rules exactly and never resend a ticket "
                 "that is already in it.")


def sh(cmd, timeout=120):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                              shell=isinstance(cmd, str))
    except subprocess.TimeoutExpired as e:
        return subprocess.CompletedProcess(cmd, 124, "", f"timeout: {e}")


def ssh(remote, timeout=120):
    return sh(SSH + [remote], timeout)


def session_alive(svc):
    return sh(["/opt/homebrew/bin/tmux", "has-session", "-t", SESSION[svc]]).returncode == 0


def stop_session(svc):
    sh(["/opt/homebrew/bin/tmux", "kill-session", "-t", SESSION[svc]])
    for _ in range(30):                                     # wrapper terminates the child on SIGHUP
        if sh(["pgrep", "-f", f"gcx_owner.py wrap {svc}"]).returncode != 0:
            return True
        time.sleep(2)
    sh(["pkill", "-9", "-f", f"gcx_owner.py wrap {svc}"])
    if svc == "aitrader":
        sh(["pkill", "-9", "-f", f"{AIT}"])
    return False


# ---------------- while the server is up ----------------
def mirror(st):
    """Copy the server's live state to the Mac (only while the server owns the services)."""
    os.makedirs(MIRROR, exist_ok=True)
    m = st.setdefault("mirror", {})
    r = sh(["rsync", "-az", "--delete", "--exclude", "logs/", "--exclude", "app/.lock", "--exclude", "app/port",
            "-e", RSH, "kevinkim@gcx-server:mock-trader/state/", os.path.join(MIRROR, "aitrader-state") + "/"])
    if r.returncode == 0:
        m["aitrader"] = time.time()
    os.makedirs(os.path.join(MIRROR, "ticket-reporter"), exist_ok=True)
    r = sh(["rsync", "-az", "-e", RSH, "--include", "*.json", "--exclude", "*",
            "kevinkim@gcx-server:.claude/skills/ticket-reporter/state/", os.path.join(MIRROR, "ticket-reporter") + "/"])
    if r.returncode == 0:
        m["ticket-monitor"] = time.time()
    return bool(m)


def ensure_gates(st, log):
    """Install the server half (install_server.sh) once, when the server is reachable."""
    if st.get("gates_ok"):
        return True
    try:
        if O.L.get("gates:installed"):
            st["gates_ok"] = True
            return True
    except Exception:
        return False
    if time.time() - st.get("gates_try", 0) < 1800:
        return False
    st["gates_try"] = time.time()
    r = ssh("cd ~/Desktop/GCX && git pull -q --rebase origin main >/dev/null 2>&1; "
            "bash ~/Desktop/GCX/ServerBootstrap/fallback/install_server.sh", 180)
    log(f"install_server.sh: rc={r.returncode} {r.stdout.strip()[-200:]} {r.stderr.strip()[-200:]}")
    st["gates_ok"] = r.returncode == 0
    return st["gates_ok"]


# ---------------- take over ----------------
def takeover(svc, st, log):
    """Start the Mac copy. Returns a status line for the Chat card."""
    s = st.setdefault("svc", {}).setdefault(svc, {})
    if not st.get("gates_ok"):
        return f"{LABEL[svc]}: NOT started — the server's hand-over guard isn't installed yet (installs itself next time the server is online)"
    age = (time.time() - st.get("mirror", {}).get(svc, 0)) / 60
    if not s.get("dirty") and age > MAX_MIRROR_AGE[svc]:
        return (f"{LABEL[svc]}: NOT started — the Mac's copy of the server's state is {age:.0f} min old "
                f"(limit {MAX_MIRROR_AGE[svc]}), running it could repeat what the server already did")
    if not O.set_owner(svc, "mac"):
        return f"{LABEL[svc]}: NOT started — couldn't record the hand-over in the ledger"
    tmux = "/opt/homebrew/bin/tmux"
    if svc == "aitrader":
        sh(["rsync", "-a", "--delete", "--exclude", "state", "--exclude", ".venv", "--exclude", "ios",
            "--exclude", "__pycache__", "--exclude", ".env", MT + "/", AIT + "/"])
        for n in (".venv", ".env"):
            if not os.path.lexists(os.path.join(AIT, n)):
                os.symlink(os.path.join(MT, n), os.path.join(AIT, n))
        if not s.get("dirty"):
            sh(["rsync", "-a", "--delete", os.path.join(MIRROR, "aitrader-state") + "/", os.path.join(AIT, "state") + "/"])
        os.makedirs(os.path.join(AIT, "state", "logs"), exist_ok=True)
        os.makedirs(os.path.join(AIT, "state", "app"), exist_ok=True)
        cmd = (f"cd {AIT} && AITRADER_HEADLESS=1 /opt/homebrew/bin/python3 {HERE}/gcx_owner.py wrap aitrader -- "
               f"{MT}/.venv/bin/python -u -m app.main >> {AIT}/state/logs/engine.log 2>&1")
    else:
        os.makedirs(SKILL_STATE, exist_ok=True)
        if not s.get("dirty"):
            bk = os.path.join(FB, f"ticket-state-mac-{time.strftime('%Y%m%d-%H%M')}")
            shutil.copytree(SKILL_STATE, bk, dirs_exist_ok=True)
            for f in glob.glob(os.path.join(MIRROR, "ticket-reporter", "*.json")):
                shutil.copy2(f, SKILL_STATE)
        cmd = (f"cd ~ && export PATH=$HOME/.local/bin:/opt/homebrew/bin:$PATH; /opt/homebrew/bin/python3 "
               f"{HERE}/gcx_owner.py wrap ticket-monitor -- claude --remote-control gcx-ticket-monitor-mac "
               f"{json.dumps(TICKET_PROMPT)}")
    sh([tmux, "kill-session", "-t", SESSION[svc]])
    r = sh([tmux, "new-session", "-d", "-s", SESSION[svc], cmd])
    if r.returncode != 0:
        O.set_owner(svc, "server")
        return f"{LABEL[svc]}: NOT started — {r.stderr.strip()[:200]}"
    s.update(active=True, dirty=True, since=time.strftime("%H:%M"))
    log(f"{svc}: started on the Mac")
    extra = " Watch it live: `tmux attach -t gcx-fb-ticket` or the 'gcx-ticket-monitor-mac' Remote Control session." \
        if svc == "ticket-monitor" else " The Mac's AI Trader app window now shows this engine (the iPhone app can't reach it until the server is back)."
    return f"{LABEL[svc]}: running on the Mac (state from {age:.0f} min ago).{extra}"


def keep(svc, st, log):
    """Each tick while the Mac owns svc: refresh the lease, restart the copy if it died, point the app window here."""
    O.set_owner(svc, "mac")
    if not session_alive(svc):
        log(f"{svc}: Mac copy not running — restarting")
        st["svc"][svc]["active"] = False
        takeover(svc, st, log)
    if svc == "aitrader":
        rj = os.path.join(MT, "state", "app", "remote.json")
        try:
            port = open(os.path.join(AIT, "state", "app", "port")).read().strip()
            tok = open(os.path.join(AIT, "state", "app", "mobile_token")).read().strip()
            if os.path.exists(rj) and not os.path.exists(rj + ".server"):
                shutil.copy2(rj, rj + ".server")
            want = {"url": f"http://127.0.0.1:{port}", "token": tok}
            if not os.path.exists(rj) or json.load(open(rj)) != want:
                json.dump(want, open(rj, "w")); os.chmod(rj, 0o600)
        except Exception:
            pass


# ---------------- hand back ----------------
def handback(svc, st, log):
    """Server is back and reachable: stop the Mac copy, push its state, owner=server. Returns a line or None (retry)."""
    s = st["svc"][svc]
    try:
        o, age = O.owner(svc)
    except Exception as e:
        log(f"{svc}: ledger unreachable for hand-back ({e})"); return None
    lease_ok = o == "mac" and age < O.LEASE_MIN
    # the server's copy must be held stopped by gcx-owner-watch before its state is overwritten
    if lease_ok:
        for _ in range(8):
            chk = ssh("systemctl --user is-active mock-trader" if svc == "aitrader"
                      else "tmux list-windows -t gcx -F '#W' 2>/dev/null | grep -qx ticket-monitor && echo running || echo stopped", 30)
            if chk.returncode == 255:
                return None                                  # ssh failed — keep the Mac copy running, retry next tick
            if chk.stdout.strip() not in ("active", "activating", "running"):
                break
            time.sleep(10)
        else:
            log(f"{svc}: server copy still running — not pushing state yet"); return None
    stop_session(svc)
    ts = time.strftime("%Y%m%d-%H%M")
    if svc == "aitrader":
        src = os.path.join(AIT, "state") + "/"
        dest = "mock-trader/state/" if lease_ok else f"mock-trader/state-mac-failover-{ts}/"
        r = sh(["rsync", "-az", "--exclude", "logs/", "--exclude", "app/.lock", "--exclude", "app/port", "--exclude", "app/remote.json",
                "-e", RSH, src, f"kevinkim@gcx-server:{dest}"], 300)
        sh(["rsync", "-az", "-e", RSH, os.path.join(AIT, "state", "logs") + "/",
            f"kevinkim@gcx-server:mock-trader/state/logs/mac-failover-{ts}/"], 300)
        rj = os.path.join(MT, "state", "app", "remote.json")
        if os.path.exists(rj + ".server"):
            os.replace(rj + ".server", rj)
    else:
        r = sh(["rsync", "-az", "-e", RSH, "--include", "*.json", "--exclude", "*", SKILL_STATE + "/",
                f"kevinkim@gcx-server:.gcx-fallback-in/ticket-{ts}/"], 120)
        if r.returncode == 0 and lease_ok:
            r = ssh(f"python3 ~/Desktop/GCX/ServerBootstrap/fallback/merge_state.py ~/.gcx-fallback-in/ticket-{ts} "
                    f"~/.claude/skills/ticket-reporter/state", 60)
    if r.returncode != 0:
        log(f"{svc}: pushing state failed: {r.stderr[-300:]}")
        O.set_owner(svc, "mac")                            # keep the server copy paused rather than run on stale state
        if s.get("push_alerted"):
            return None
        s["push_alerted"] = 1
        return (f"⚠️ {LABEL[svc]}: Mac copy stopped but its state could NOT be copied back ({r.stderr.strip()[-150:]}). "
                "The server copy stays paused; retrying every 5 min — or ask Claude to finish the hand-back.")
    O.set_owner(svc, "server")
    s.update(active=False, dirty=False, push_alerted=None)
    log(f"{svc}: handed back (lease_ok={lease_ok})")
    if lease_ok:
        return f"{LABEL[svc]}: back on the server with the Mac's latest state (Mac ran it since {s.get('since', '?')})."
    return (f"⚠️ {LABEL[svc]}: the Mac's lease ran out (Mac asleep/offline) so the server had already resumed with its own "
            f"state. The Mac's state was saved on the server for review ({'mock-trader/state-mac-failover-' + ts if svc == 'aitrader' else '~/.gcx-fallback-in/ticket-' + ts}) "
            + ("— check AI Trader positions vs Toss." if svc == "aitrader" else "— tickets the Mac reported may be reported once more."))
