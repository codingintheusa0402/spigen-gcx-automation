#!/usr/bin/env python3
"""GCX Mesh — merge Claude session histories (~/.claude/projects/**/*.jsonl) between the MacBook and the server
(the Mac mini gcx-macmini since 2026-10-16; was the Windows laptop gcx-server).

Runs ON THE MAC (called by mesh_sync_mac.sh); the same file is used on the server over ssh for its side.
A session can be continued on both machines (e.g. resumed on the server while the Mac was off, and the Mac
copy is still open), so "newer file wins" would silently drop one side's conversation. Instead, entries are
merged: each side gets the lines it is missing, appended in the other side's order. Nothing is ever deleted
or rewritten, so a session that is live on either machine keeps writing safely.

Line identity = its "uuid" field, else a hash of the raw line. Only newline-terminated lines count (a live
writer's half-written last line is left for the next run). Automatic `claude -p` runs (entrypoint sdk-cli,
e.g. the AI Trader's news scoring) are not copied.

Project folders are renamed on the way: each machine's prefix is its home path (MacBook -Users-kevinkim…,
mini -Users-kevin…, old laptop -home-kevinkim…)  (memory/ is skipped).
Subcommands (internal):  list | keys <rel>... | emit (stdin: {rel: [keys]}) | append (stdin: stream)
"""
import hashlib, json, os, subprocess, sys

HOME = os.path.expanduser('~')
ROOT = os.path.join(HOME, '.claude', 'projects')
LOCAL_P = HOME.replace('/', '-')          # how Claude Code names project folders under this home
CACHE = os.path.join(HOME, '.gcx-mesh-jsonl-cache.json')
H = os.environ.get('MESH_SYNC_HOST', 'gcx-macmini')
REMOTE = os.environ.get('MESH_SYNC_REMOTE', '~/.gcx-mesh/mesh_jsonl_sync.py')   # mesh_sync_mac.sh copies this file there


def to_local(rel):   # rel uses the neutral prefix "K"
    return os.path.join(ROOT, LOCAL_P + rel[1:])


def walk():
    out = {}
    for d in os.listdir(ROOT):
        full = os.path.join(ROOT, d)
        rest = d[len(LOCAL_P):]
        # -Users-kevin must not swallow -Users-kevinkim…: the prefix has to end at a path separator
        if not d.startswith(LOCAL_P) or (rest and rest[0] != '-') or os.path.islink(full) or not os.path.isdir(full):
            continue
        for dp, dns, fns in os.walk(full):
            dns[:] = [x for x in dns if x != 'memory']
            for f in fns:
                if f.endswith('.jsonl'):
                    p = os.path.join(dp, f)
                    out['K' + os.path.relpath(p, ROOT)[len(LOCAL_P):]] = os.path.getsize(p)
    return out


def is_sdk(path):
    try:
        with open(path, 'rb') as fh:
            for _ in range(8):
                ln = fh.readline()
                if not ln:
                    break
                if b'"entrypoint":"sdk-' in ln:
                    return True
    except OSError:
        pass
    return False


def key(line):
    try:
        u = json.loads(line).get('uuid')
        if u:
            return u
    except Exception:
        pass
    return 'h:' + hashlib.sha1(line).hexdigest()


def keys_of(rel, cache):
    """Keys of all complete lines. Files only grow, so a cached prefix is reused."""
    p = to_local(rel)
    c = cache.get(rel)
    size = os.path.getsize(p)
    if not c or c['size'] > size:
        c = {'size': 0, 'keys': []}
    with open(p, 'rb') as fh:
        fh.seek(c['size'])
        data = fh.read()
    end = data.rfind(b'\n') + 1
    new = [key(ln) for ln in data[:end].split(b'\n')[:-1] if ln.strip()]
    c = {'size': c['size'] + end, 'keys': c['keys'] + new}
    cache[rel] = c
    return c['keys']


def load_cache():
    try:
        return json.load(open(CACHE))
    except Exception:
        return {}


def save_cache(cache, live):
    cache = {k: v for k, v in cache.items() if k in live}
    tmp = CACHE + '.tmp'
    json.dump(cache, open(tmp, 'w'))
    os.replace(tmp, CACHE)


def emit(want):
    """want = {rel: [keys the other side already has]} -> stream of lines the other side lacks."""
    out = sys.stdout.buffer
    for rel, have in want.items():
        have = set(have)
        p = to_local(rel)
        with open(p, 'rb') as fh:
            data = fh.read()
        data = data[:data.rfind(b'\n') + 1]
        lines = [ln for ln in data.split(b'\n')[:-1] if ln.strip() and key(ln) not in have]
        if lines:
            out.write(json.dumps({'rel': rel, 'n': len(lines)}).encode() + b'\n')
            for ln in lines:
                out.write(ln + b'\n')


def append(stream):
    n_files = n_lines = 0
    while True:
        hdr = stream.readline()
        if not hdr:
            break
        h = json.loads(hdr)
        lines = [stream.readline() for _ in range(h['n'])]
        p = to_local(h['rel'])
        os.makedirs(os.path.dirname(p), exist_ok=True)
        if os.path.exists(p) and os.path.getsize(p):
            with open(p, 'rb') as fh:
                fh.seek(-1, 2)
                if fh.read(1) != b'\n':        # a live writer is mid-line: try again next run
                    continue
        with open(p, 'ab') as fh:
            fh.write(b''.join(lines))
        n_files += 1
        n_lines += len(lines)
    print(json.dumps({'files': n_files, 'lines': n_lines}))


def ssh(args, inp=None):
    r = subprocess.run(['ssh', '-o', 'BatchMode=yes', '-o', 'ConnectTimeout=8', H, 'python3 ' + REMOTE + ' ' + args],
                       input=inp, capture_output=True, timeout=600)
    if r.returncode:
        raise RuntimeError(r.stderr.decode()[-400:])
    return r.stdout


def main_mac(dry=False):
    cache = load_cache()
    mine, theirs = walk(), json.loads(ssh('list'))
    both = [r for r in mine if r in theirs and mine[r] != theirs[r]]
    only_mine = [r for r in mine if r not in theirs and not is_sdk(to_local(r))]
    only_theirs = [r for r in theirs if r not in mine]
    # their keys, in one call (cached on the server side too)
    tk = json.loads(ssh('keys', json.dumps(both).encode())) if both else {}
    mk = {r: keys_of(r, cache) for r in both}
    to_srv = {r: tk[r] for r in both if set(mk[r]) - set(tk[r])}
    to_mac = {r: mk[r] for r in both if set(tk[r]) - set(mk[r])}
    forks = [r for r in to_srv if r in to_mac]
    for r in only_mine:
        to_srv[r] = []
    for r in only_theirs:
        to_mac[r] = []
    summary = {'push_files': len(to_srv), 'pull_files': len(to_mac), 'forks_merged': forks}
    if dry:
        summary['push'] = sorted(to_srv)
        summary['pull'] = sorted(to_mac)
        print(json.dumps(summary, indent=1))
        return
    if to_srv:
        buf = subprocess.run([sys.executable, __file__, 'emit'], input=json.dumps(to_srv).encode(),
                             capture_output=True).stdout
        summary['pushed'] = json.loads(ssh('append', buf))
    if to_mac:
        stream = ssh('emit-sdk-filtered', json.dumps(to_mac).encode())
        import io
        old = sys.stdout
        sys.stdout = io.StringIO()
        append(io.BytesIO(stream))
        summary['pulled'] = json.loads(sys.stdout.getvalue())
        sys.stdout = old
    save_cache(cache, mine)
    print(json.dumps(summary))


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'sync'
    if cmd == 'list':
        print(json.dumps(walk()))
    elif cmd == 'keys':
        cache = load_cache()
        rels = json.load(sys.stdin)
        print(json.dumps({r: keys_of(r, cache) for r in rels}))
        save_cache(cache, walk())
    elif cmd == 'emit':
        emit(json.load(sys.stdin))
    elif cmd == 'emit-sdk-filtered':   # server side: never hand automatic sdk runs that only exist here
        want = json.load(sys.stdin)
        emit({r: v for r, v in want.items() if v or not is_sdk(to_local(r))})
    elif cmd == 'append':
        append(sys.stdin.buffer)
    elif cmd in ('sync', 'dry'):
        main_mac(dry=cmd == 'dry')
