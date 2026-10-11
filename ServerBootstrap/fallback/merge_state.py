#!/usr/bin/env python3
"""Merge the Mac fallback's ticket-reporter state into the server's (hand-back), so nothing it reported is sent again.

  merge_state.py <dir with the Mac's *.json> <server state dir>
Lists → union (order kept), dicts → merged recursively, numbers → max, anything else → the Mac's value.
The server's file is backed up as <name>.before-mac-merge first.
"""
import glob, json, os, shutil, sys


def merge(a, b):
    if isinstance(a, dict) and isinstance(b, dict):
        out = dict(a)
        for k, v in b.items():
            out[k] = merge(a[k], v) if k in a else v
        return out
    if isinstance(a, list) and isinstance(b, list):
        seen, out = set(), []
        for x in a + b:
            key = json.dumps(x, sort_keys=True, ensure_ascii=False)
            if key not in seen:
                seen.add(key); out.append(x)
        return out
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return max(a, b)
    return b if b is not None else a


src, dst = sys.argv[1], sys.argv[2]
for f in glob.glob(os.path.join(src, "*.json")):
    t = os.path.join(dst, os.path.basename(f))
    mac = json.load(open(f))
    if os.path.exists(t):
        shutil.copy2(t, t + ".before-mac-merge")
        mac = merge(json.load(open(t)), mac)
    tmp = t + ".tmp"
    json.dump(mac, open(tmp, "w"), ensure_ascii=False, indent=1)
    os.replace(tmp, t)
    print("merged", os.path.basename(f))
