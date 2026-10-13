#!/usr/bin/env python3
"""Two-way merge of Claude auto-memory folders between the Mac and the server.

For each memory dir (-Users-kevinkim/memory <-> -home-kevinkim/memory, and the Desktop-GCX one):
  * a file that exists on one side only is copied to the other
  * a file that exists on both: the newer copy wins (the older one is kept in ~/.claude/memory_backup/<date>/)
  * MEMORY.md (the index) is merged: union of entries, each memory file listed once, Mac order first
Never deletes anything. Run from the Mac: python3 sync_memory.py
"""
import os, re, shutil, subprocess, sys, tempfile, time

H = "kevinkim@gcx-server"
SSH = ["ssh", "-o", "ConnectTimeout=25"]
PAIRS = [("-Users-kevinkim", "-home-kevinkim"), ("-Users-kevinkim-Desktop-GCX", "-home-kevinkim-Desktop-GCX")]
ROOT = os.path.expanduser("~/.claude/projects")
BACKUP = os.path.expanduser(f"~/.claude/memory_backup/{time.strftime('%Y%m%d-%H%M%S')}")
LINK = re.compile(r"\]\(([^)]+\.md)\)")


def merge_index(mac_text, srv_text):
    mac_lines = mac_text.rstrip("\n").split("\n")
    have = {m.group(1) for l in mac_lines for m in [LINK.search(l)] if m}
    extra = [l for l in srv_text.split("\n") if (m := LINK.search(l)) and m.group(1) not in have]
    if not extra:
        return mac_text, []
    out = list(mac_lines)
    fb = [l for l in extra if "feedback_" in l]
    pj = [l for l in extra if l not in fb]

    def insert(head, items):
        if head not in out:
            out.extend(["", head])
        i = out.index(head) + 1
        while i < len(out) and out[i].startswith("- "):
            i += 1
        for it in items:
            out.insert(i, it); i += 1
    if pj:
        insert("## Project", pj)
    if fb:
        insert("## Feedback", fb)
    return "\n".join(out) + "\n", extra


def main():
    total = {"to_mac": 0, "to_server": 0, "index": 0}
    for mac_d, srv_d in PAIRS:
        mac = f"{ROOT}/{mac_d}/memory"
        srv_rel = f".claude/projects/{srv_d}/memory"
        if not os.path.isdir(mac):
            continue
        tmp = tempfile.mkdtemp(prefix="srvmem_")
        r = subprocess.run(["rsync", "-az", "-e", " ".join(SSH), f"{H}:{srv_rel}/", f"{tmp}/"], capture_output=True, text=True)
        if r.returncode not in (0, 23):   # 23 = dir missing on server → treat as empty
            print(f"server unreachable for {srv_d}: {r.stderr.strip()[:200]}"); continue
        for n in os.listdir(tmp):
            if n.startswith(".") or n == "MEMORY.md":
                continue
            sp, mp = f"{tmp}/{n}", f"{mac}/{n}"
            if os.path.isdir(sp):
                continue
            if not os.path.exists(mp):
                shutil.copy2(sp, mp); total["to_mac"] += 1
            elif os.path.getmtime(sp) > os.path.getmtime(mp) + 2 and open(sp, "rb").read() != open(mp, "rb").read():
                os.makedirs(BACKUP, exist_ok=True); shutil.copy2(mp, f"{BACKUP}/{n}")
                shutil.copy2(sp, mp); total["to_mac"] += 1
        si, mi = f"{tmp}/MEMORY.md", f"{mac}/MEMORY.md"
        if os.path.exists(si) and os.path.exists(mi):
            merged, extra = merge_index(open(mi).read(), open(si).read())
            if extra:
                os.makedirs(BACKUP, exist_ok=True); shutil.copy2(mi, f"{BACKUP}/MEMORY.md")
                open(mi, "w").write(merged); total["index"] += len(extra)
        elif os.path.exists(si):
            shutil.copy2(si, mi)
        shutil.rmtree(tmp, ignore_errors=True)
        # the Mac folder is now the superset → push it (newest wins there too)
        subprocess.run(SSH + [H, f"mkdir -p ~/{srv_rel}"], capture_output=True)
        r = subprocess.run(["rsync", "-azu", "--exclude=.DS_Store", "-e", " ".join(SSH), f"{mac}/", f"{H}:{srv_rel}/"],
                           capture_output=True, text=True)
        r2 = subprocess.run(["rsync", "-az", "--exclude=.DS_Store", "-e", " ".join(SSH), mi, f"{H}:{srv_rel}/MEMORY.md"],
                            capture_output=True, text=True)   # the merged index always goes over
        total["to_server"] += 1 if r.returncode == 0 else 0
    print(f"memory merged: {total['to_mac']} file(s) taken from the server, {total['index']} index line(s) added, "
          f"{total['to_server']} folder(s) pushed; older copies in {BACKUP if os.path.isdir(BACKUP) else 'none needed'}")


if __name__ == "__main__":
    sys.exit(main())
