#!/usr/bin/env python3
"""Safe "Resume a Claude session" for the GCX server desktop shortcut (gcx-launch.sh resume).

Plain `claude --resume` let you pick a conversation that is ALREADY running in another tmux tab
(e.g. the ticket monitor), which started a second live copy of its loop. This picker lists recent
conversations, marks the live ones with their tab, and:
  - live      -> switches to that tab (no second copy)
  - not live  -> resumes it here with --remote-control, tab renamed after it
Automatic `claude -p` runs (entrypoint sdk-cli, e.g. the AI Trader's news scoring) are hidden.
"""
import glob, json, os, subprocess, sys, time

HOME = os.path.expanduser("~")
PROJ = os.path.join(HOME, ".claude", "projects")


def sh(*a):
    return subprocess.run(a, capture_output=True, text=True).stdout


def live_sessions():
    """sessionId -> tmux window name, for every running claude that has a session file."""
    pane_of = {}
    for line in sh("tmux", "list-panes", "-a", "-F", "#{pane_pid}\t#{window_index}:#{window_name}").splitlines():
        pp, win = line.split("\t", 1)
        for k in sh("pgrep", "-P", pp).split():
            pane_of[k] = win
    out = {}
    for f in glob.glob(os.path.join(HOME, ".claude", "sessions", "*.json")):
        pid = os.path.basename(f)[:-5]
        if not os.path.exists(f"/proc/{pid}"):
            continue
        try:
            d = json.load(open(f))
        except Exception:
            continue
        if d.get("sessionId"):
            out[d["sessionId"]] = pane_of.get(pid, "outside tmux")
    return out


def conversations(limit=25):
    rows, seen = [], set()
    for f in sorted(glob.glob(os.path.join(PROJ, "*", "*.jsonl")), key=os.path.getmtime, reverse=True):
        if len(rows) >= limit:
            break
        if os.path.realpath(f) in seen:          # -Users-kevinkim is a symlink to -home-kevinkim
            continue
        seen.add(os.path.realpath(f))
        title = first = cwd = None
        auto = False
        try:
            with open(f, errors="ignore") as fh:
                for i, l in enumerate(fh):
                    if '"entrypoint":"sdk-cli"' in l:
                        auto = True
                        break
                    if '"custom-title"' in l or '"ai-title"' in l:
                        try:
                            d = json.loads(l); title = d.get("customTitle") or d.get("aiTitle") or title
                        except Exception:
                            pass
                    if cwd is None and '"cwd"' in l:
                        try:
                            cwd = json.loads(l).get("cwd")
                        except Exception:
                            pass
                    if first is None and '"type":"user"' in l:
                        try:
                            c = json.loads(l)["message"]["content"]
                            first = c if isinstance(c, str) else None
                        except Exception:
                            pass
        except OSError:
            continue
        if auto or not (title or first):
            continue
        rows.append(dict(sid=os.path.basename(f)[:-6], title=(title or first or "").replace("\n", " ")[:60],
                         cwd=cwd or HOME, t=os.path.getmtime(f)))
    return rows


def main():
    live, rows = live_sessions(), conversations()
    print("\n  Resume a Claude session  (LIVE = already running — you'll be switched to its tab)\n")
    for i, r in enumerate(rows, 1):
        tag = f"LIVE in tab {live[r['sid']]}" if r["sid"] in live else time.strftime("%m-%d %H:%M", time.localtime(r["t"]))
        print(f"  {i:>2}. {r['title']:<60}  {tag}")
    try:
        pick = input("\n  Number (Enter = cancel): ").strip()
    except EOFError:
        pick = ""
    if not pick.isdigit() or not 1 <= int(pick) <= len(rows):
        return 0
    r = rows[int(pick) - 1]
    if r["sid"] in live:
        win = live[r["sid"]]
        if win != "outside tmux":
            print(f"  Already running in tab {win} — switching there.")
            sh("tmux", "select-window", "-t", "gcx:" + win.split(":", 1)[0])
            time.sleep(1)
            return 0
        print("  Already running outside tmux — not starting a second copy.")
        time.sleep(3)
        return 0
    name = "".join(ch for ch in r["title"] if ch.isalnum() or ch in " -가힣")[:18].strip() or r["sid"][:8]
    sh("tmux", "rename-window", name)
    os.chdir(r["cwd"] if os.path.isdir(r["cwd"]) else HOME)
    os.execvp("claude", ["claude", "--resume", r["sid"], "--remote-control"])


if __name__ == "__main__":
    sys.exit(main())
