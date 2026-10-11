#!/usr/bin/env python3
"""Shared run ledger for jobs that run on gcx-server, or on the Mac as a fallback when the server is down.

Each job's local state (ran_ markers, state.json) lives on whichever machine ran it, so without this the
Mac taking over for a dead server — or the server coming back after the Mac covered for it — could send
the same Chat report twice. Both machines record every real run here first:

  tab `ledger` of the private "GCX Server Heartbeat" sheet (~/.config/gcx_heartbeat_sheet.txt)
  A key (e.g. badreview:2026-10-12) · B status · C host (server|mac) · D updated ms · E data (json)

status: running → done / sent / held / partial. A `running` row from the other host blocks for
STALE_MIN; after that it's treated as a crashed run and can be taken over.

  from gcx_ledger import claim, put, get
  if not claim(f"badreview:{today}"): return     # another host ran / is running it
  ...; put(f"badreview:{today}", "sent")

If the sheet can't be reached: the server proceeds (keeps today's behaviour), the Mac doesn't (it is only
a fallback — never risk a double send to the live rooms).
"""
import json, os, time

SHEET_FILE = os.path.expanduser("~/.config/gcx_heartbeat_sheet.txt")
TOKEN = os.path.expanduser("~/.config/gws_shim/token.json")
TAB = "ledger"
STALE_MIN = 60                     # a real badreview run takes ≤ ~45 min
KEEP_DAYS = 14
IS_SERVER = os.path.exists("/proc/version") and "microsoft" in open("/proc/version").read().lower()
HOST = "server" if IS_SERVER else "mac"
_svc_cache = None


def log(msg):
    print(f"[ledger {HOST}] {msg}", flush=True)


def _svc():
    global _svc_cache
    if _svc_cache:
        return _svc_cache
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    creds = Credentials.from_authorized_user_info(json.load(open(TOKEN)))
    if not creds.valid:
        creds.refresh(Request())
        with open(TOKEN, "w") as f:
            f.write(creds.to_json())
    _svc_cache = build("sheets", "v4", credentials=creds, cache_discovery=False).spreadsheets()
    return _svc_cache


def _sid():
    return open(SHEET_FILE).read().strip()


def _rows():
    s = _svc()
    try:
        return s.values().get(spreadsheetId=_sid(), range=f"{TAB}!A2:E").execute().get("values", [])
    except Exception as e:
        if "Unable to parse range" not in str(e):
            raise
        s.batchUpdate(spreadsheetId=_sid(), body={"requests": [{"addSheet": {"properties": {"title": TAB}}}]}).execute()
        s.values().update(spreadsheetId=_sid(), range=f"{TAB}!A1:E1", valueInputOption="RAW",
                          body={"values": [["key", "status", "host", "updated_ms", "data"]]}).execute()
        return []


def _parse(r):
    r = r + [""] * (5 - len(r))
    try:
        data = json.loads(r[4]) if r[4] else None
    except Exception:
        data = None
    return {"key": r[0], "status": r[1], "host": r[2], "updated": int(r[3] or 0), "data": data}


def get(key):
    hit = None
    for r in _rows():
        if r and r[0] == key:
            hit = _parse(r)          # last one wins
    return hit


def put(key, status, data=None):
    """Update the key's row in place (or append it). Row-level writes, so jobs writing at the same
    minute don't clobber each other. Errors are logged, not raised."""
    try:
        rows = _rows()
        now = int(time.time() * 1000)
        new = [key, status, HOST, str(now), json.dumps(data, ensure_ascii=False) if data is not None else ""]
        s = _svc()
        idx = [i for i, r in enumerate(rows) if r and r[0] == key]
        if idx:
            n = idx[-1] + 2
            s.values().update(spreadsheetId=_sid(), range=f"{TAB}!A{n}:E{n}", valueInputOption="RAW",
                              body={"values": [new]}).execute()
        else:
            s.values().append(spreadsheetId=_sid(), range=f"{TAB}!A:E", valueInputOption="RAW",
                              insertDataOption="INSERT_ROWS", body={"values": [new]}).execute()
            # tidy: blank out dated rows older than KEEP_DAYS (":state" rows are kept forever)
            cut = now - KEEP_DAYS * 86400000
            old = [i + 2 for i, r in enumerate(rows) if r and r[0] and not r[0].endswith(":state") and _parse(r)["updated"] < cut]
            if old:
                s.values().batchClear(spreadsheetId=_sid(), body={"ranges": [f"{TAB}!A{n}:E{n}" for n in old]}).execute()
        return True
    except Exception as e:
        log(f"put {key} failed: {e}")
        return False


def claim(key, allow=(), fail_open=None):
    """True → this host may run `key` now (row set to running). False → done elsewhere / running elsewhere.
    `allow`: statuses that still permit a run (e.g. ("held",) for the 11:30 retry)."""
    fail_open = IS_SERVER if fail_open is None else fail_open
    try:
        row = get(key)
    except Exception as e:
        log(f"ledger unreachable ({e}) — {'running anyway' if fail_open else 'not running'}")
        return fail_open
    if row and row["status"] not in allow:
        fresh = time.time() * 1000 - row["updated"] < STALE_MIN * 60000
        if row["status"] in ("done", "sent", "held"):
            log(f"skip {key}: already {row['status']} on {row['host']}")
            return False
        if row["status"] == "running" and row["host"] != HOST and fresh:
            log(f"skip {key}: running on {row['host']}")
            return False
    if not put(key, "running", row and row["data"]):
        return fail_open
    return True


def status(key):
    try:
        row = get(key)
        return row and row["status"]
    except Exception:
        return None
