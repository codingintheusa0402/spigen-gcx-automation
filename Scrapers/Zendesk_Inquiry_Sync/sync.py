#!/usr/bin/env python3
"""Zendesk → '26년 전체문의' sheet sync.

Appends Solved (incl. Closed) Zendesk tickets from Caspi (S3.ZENDESK.*), created in 2026
KST, 3 GCX channels, that are not yet in the sheet (dedupe by Ticket ID in col A). Writes A:AD only; formula columns
AE~ are extended by copying the last existing row's formulas down.

Subcommands:
  run [--dry-run] [--solved-only] [--until-yesterday] [--scheduled]   fetch + append (scheduled = self-gated by schedule)
      --until-yesterday: only tickets created up to yesterday (Ticket created - Date ≤ yesterday KST);
                         today's tickets wait for tomorrow's run. Dedupe makes it cumulative.
  setup --caspi-key K --query-id Q [--client-secret F] [--chat-webhook URL]   store per-user credentials
      --chat-webhook: Google Chat incoming webhook; scheduled runs post "Zendesk Raw Data 업데이트 완료"
                      (or a failure notice, once per day)
      --intake-query-id: Caspi queryId of intake_query.sql → notice also shows yesterday's intake
                      (after unscheduled days, e.g. Monday on a weekdays schedule: Fri~Sun total)
                      (all statuses: 완료 = Solved·Closed / 처리 중 = the rest)
  schedule --days mon,thu --time 09:00 [--until-yesterday]   install/update the launchd job (Windows: Task Scheduler)
  unschedule                      remove the launchd job (Windows: Task Scheduler task)
  status                          show config, schedule, last run
"""
import argparse, datetime, json, os, subprocess, sys, time
IS_WIN = sys.platform == "win32"
if IS_WIN:
    import msvcrt
else:
    import fcntl
from xml.sax.saxutils import escape
import urllib.error, urllib.request

HERE = os.path.dirname(os.path.realpath(__file__))
CFG_DIR = os.path.expanduser("~/.config/zendesk_inquiry_sync")
CREDS_PATH = os.path.join(CFG_DIR, "credentials.json")      # caspi api key + queryId
GTOKEN_PATH = os.path.join(CFG_DIR, "google_token.json")    # authorized-user OAuth json
GWS_SHIM_TOKEN = os.path.expanduser("~/.config/gws_shim/token.json")  # fallback
SCHED_PATH = os.path.join(CFG_DIR, "schedule.json")
STATE_PATH = os.path.join(CFG_DIR, "state.json")
LOCK_PATH = os.path.join(CFG_DIR, "run.lock")
LOG_DIR = (os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "zendesk-inquiry-sync", "Logs")
           if IS_WIN else os.path.expanduser("~/Library/Logs/zendesk-inquiry-sync"))
LABEL = "com.spigen.gcx.zendesk-inquiry-sync"
PLIST = os.path.expanduser(f"~/Library/LaunchAgents/{LABEL}.plist")
WIN_TASK = "Spigen GCX Zendesk Inquiry Sync"
# how often the scheduled job wakes to check (launchd StartInterval 1800 / Task Scheduler 5 min,
# so a PC switched on after the scheduled time catches up within minutes)
TICK_MIN = 5 if IS_WIN else 30

SPREADSHEET_ID = os.environ.get("ZIS_SPREADSHEET_ID", "1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc")  # env override = test copy
SHEET_GID = 1597176315                                       # '26년 전체문의'
SHEET_TITLE = "26년 전체문의"                                  # fallback if the tab is re-created
CASPI_ENDPOINT = "https://caspilm.spigen.com/api/data-api/run"
# First ticket of 2026 (KST). Caspi's created_at is a UTC *date*, so the year boundary is
# set by ticket ID (IDs are monotonic) instead of by date.
MIN_TICKET_ID = 1000132837
# Zendesk auto-flips Solved → Closed after a few days; Closed = solved & locked. Default keeps
# both so tickets that close between two runs are never lost. `run --solved-only` = strict.
STATUSES = {"solved", "closed"}
CHANNELS = {"SQ_website", "Amazon Buyer Message", "spigen.support@spigen.com"}
# The sheet only ever held tickets with ≥1 agent reply (verified 2026-09-30: 0 of 22,368 rows
# have 0 replies). This also drops merged duplicates (closed_by_merge), untriaged Amazon
# buyer messages and auto-closed noise.
MIN_AGENT_REPLIES = 1
KST = datetime.timezone(datetime.timedelta(hours=9))
DAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

# Sheet header row A:AD, exactly as it exists. Col A is labelled "Last updated" but
# holds the Ticket ID (Zendesk Explore export quirk) — it's the dedupe key.
HEADER = ["Last updated", "Ticket updated", "Ticket created - Date", "Country", "Brand(상세)",
          "Category", "Requested Channel type", "1차 Defect Reason or Inquiries",
          "2차 Defect Reason or Inquiries", "Device", "Product Name",
          "Power Accessories _categories", "Product Name: Caseology", "Purchase Date - Date",
          "ESC. 사유", "Tier 2 Escalated", "❗사진/영상 유무❗", "최종 고객 대응", "Order ID", "ASIN",
          "★문의SKU", "T1 → T2", "T2 → T3", "✅전체 주문 (Product Issue, 아크테크X)", "❎전체 환불",
          "✳️슈피겐 환불", "Agent replies brackets", "Agent replies", "Agent wait time (min)",
          "On-hold time (min)"]
NCOL = len(HEADER)  # 30 → A:AD
DATE_COLS = {1, 2, 13}
NUMERIC_COLS = set(range(21, 26)) | {27, 28, 29}  # V..Z, AB..AD (AA brackets stays text: "3-5")

def log(*a):
    print(datetime.datetime.now(KST).strftime("[%Y-%m-%d %H:%M:%S KST]"), *a, flush=True)


def load_json(path, default=None):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return default


def save_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    os.chmod(path, 0o600)


# ---------------------------------------------------------------- Caspi (Zendesk data)

def caspi_creds():
    c = load_json(CREDS_PATH, {})
    key = os.environ.get("CASPI_API_KEY") or c.get("caspi_api_key")
    qid = os.environ.get("CASPI_QUERY_ID") or c.get("caspi_query_id")
    if not (key and qid):
        sys.exit(f"Caspi API key / queryId missing in {CREDS_PATH} — see SKILL.md 'Setup' "
                 "(register query.sql via Caspi data_api, issue a key, then `sync.py setup`).")
    return key, qid


def caspi_fetch(key, qid, params=None):
    rows, offset = [], 0
    params = params or {"min_ticket_id": str(MIN_TICKET_ID)}
    while True:
        body = json.dumps({"queryId": qid, "params": params,
                           "limit": 2000, "offset": offset}).encode()
        for attempt in range(5):
            req = urllib.request.Request(CASPI_ENDPOINT, data=body, method="POST",
                                         headers={"x-api-key": key, "Content-Type": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=300) as f:
                    data = json.loads(f.read())
                break
            except (urllib.error.URLError, TimeoutError) as e:
                if isinstance(e, urllib.error.HTTPError) and e.code < 500 and e.code != 429:
                    sys.exit(f"Caspi API error {e.code}: {e.read()[:500]!r}")
                log(f"Caspi retry {attempt + 1}: {e}"); time.sleep(15 * (attempt + 1))
        else:
            raise RuntimeError("Caspi data-api failed after retries")
        got = data.get("rows") or []
        rows.extend(got)
        nxt = (data.get("paging") or {}).get("nextOffset")
        if nxt is None and data.get("truncated") and got:
            # responses are size-capped (~100KB ≈ 300 rows): the server flags truncated=true
            # but returns nextOffset=null, so keep paging by the rows actually received.
            nxt = offset + len(got)
        if nxt is None or not got:
            log(f"Caspi: {len(rows)} rows fetched")
            return rows
        offset = nxt


def intake_days(yesterday, sched_days):
    """Days whose intake this run reports. Day D is normally reported by the run on D+1; if D+1 has no
    scheduled run, the next run picks it up (weekdays schedule → Monday reports Fri~Sun)."""
    end = start = datetime.date.fromisoformat(yesterday)
    while DAYS[start.weekday()] not in sched_days and end - start < datetime.timedelta(days=6):
        start -= datetime.timedelta(days=1)
    return [(start + datetime.timedelta(days=i)).isoformat()
            for i in range((end - start).days + 1)]


def intake_counts(days):
    """Tickets created on `days` (all statuses) → (total, done, in_progress), or None if not configured/failed.
    Notification-only: never fails the run."""
    qid = os.environ.get("CASPI_INTAKE_QUERY_ID") or (load_json(CREDS_PATH, {}) or {}).get("caspi_intake_query_id")
    if not qid:
        return None
    try:
        key, _ = caspi_creds()
        by = {}
        for day in days:
            for r in caspi_fetch(key, qid, {"created_date": day}):
                by[r["STATUS"]] = by.get(r["STATUS"], 0) + int(float(r["N"]))
    except (Exception, SystemExit) as e:
        log(f"intake count failed: {e}")
        return None
    done = sum(n for st, n in by.items() if st in STATUSES)
    total = sum(by.values())
    log(f"intake {days[0]}~{days[-1]}: {by}")
    return total, done, total - done


def date_serial(s):
    """'YYYY-MM-DD' → Sheets date serial (days since 1899-12-30)."""
    d = datetime.date.fromisoformat(s[:10])
    return (d - datetime.date(1899, 12, 30)).days


def bracket(n):
    if n in (None, ""):
        return ""
    n = int(float(n))
    return "0" if n <= 0 else "1" if n == 1 else "2" if n == 2 else "3-5" if n <= 5 else ">5"


def build_row(r):
    g = lambda k: "" if r.get(k) is None else str(r[k])
    t2 = g("T2").lower()
    row = [
        int(r["TICKET_ID"]),
        date_serial(g("UPDATED")) if g("UPDATED") else "",
        date_serial(g("CREATED")) if g("CREATED") else "",
        g("COUNTRY"), g("BRAND"), g("CATEGORY"), g("CHANNEL"), g("DEFECT1"), g("DEFECT2"),
        g("DEVICE"), g("PRODUCT"), g("PACC_CAT"), g("CASEOLOGY"),
        date_serial(g("PURCHASE")) if g("PURCHASE")[:4].isdigit() else g("PURCHASE"),
        g("ESC"),
        True if t2 == "true" else False if t2 == "false" else "",
        g("PHOTO"), g("FINAL_RESP"), g("ORDER_ID"), g("ASIN"), g("SKU"),
        g("T1T2"), g("T2T3"), g("ORDERS_TOTAL"), g("REFUNDS_TOTAL"), g("REFUNDS_SPIGEN"),
        bracket(r.get("REPLIES")), g("REPLIES"), g("AGENT_WAIT"), g("ON_HOLD"),
    ]
    for j in NUMERIC_COLS:
        v = row[j]
        if isinstance(v, str) and v.strip().lstrip("-").isdigit():
            row[j] = int(v)
    assert len(row) == NCOL
    return row


def fetch_candidates(existing_ids, statuses, created_max=None):
    key, qid = caspi_creds()
    rows, stats = {}, {"seen": 0, "skipped_status": 0, "skipped_created_today": 0, "skipped_channel": 0,
                       "skipped_no_agent_reply": 0, "already_in_sheet": 0}
    for r in caspi_fetch(key, qid):
        stats["seen"] += 1
        if r.get("STATUS") not in statuses:
            stats["skipped_status"] += 1; continue
        if created_max and (r.get("CREATED") or "9999") > created_max:
            stats["skipped_created_today"] += 1; continue
        if str(r["TICKET_ID"]) in existing_ids:
            stats["already_in_sheet"] += 1; continue
        if r.get("CHANNEL") not in CHANNELS:
            stats["skipped_channel"] += 1; continue
        if r.get("REPLIES") in (None, "") or int(float(r["REPLIES"])) < MIN_AGENT_REPLIES:
            stats["skipped_no_agent_reply"] += 1; continue
        rows[int(r["TICKET_ID"])] = build_row(r)
    return [rows[k] for k in sorted(rows)], stats


# ---------------------------------------------------------------- Google Sheets

def sheets_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    path = os.environ.get("GOOGLE_TOKEN_PATH") or GTOKEN_PATH
    if not os.path.exists(path) and os.path.exists(GWS_SHIM_TOKEN):
        path = GWS_SHIM_TOKEN   # team's existing Sheets token, if this machine has one
    info = load_json(path)
    if not info:
        sys.exit(f"Google token missing at {path} — run `{sys.argv[0]} setup`.")
    creds = Credentials.from_authorized_user_info(info)
    if not creds.valid:
        creds.refresh(Request())
        info["token"] = creds.token
        save_json(path, info)
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def sheet_props(svc):
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields="sheets.properties").execute()
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] == SHEET_GID:
            return s["properties"]
    # tab was re-created (new gid) → fall back to its name
    for s in meta["sheets"]:
        if s["properties"]["title"] == SHEET_TITLE:
            log(f"tab gid {SHEET_GID} not found — using '{SHEET_TITLE}' (gid {s['properties']['sheetId']})")
            return s["properties"]
    sys.exit(f"Tab gid {SHEET_GID} / '{SHEET_TITLE}' not found in spreadsheet.")


def append_rows(svc, props, rows):
    title, gid = props["title"], props["sheetId"]
    ncols_grid = props["gridProperties"]["columnCount"]
    q = f"'{title}'"
    col_a = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{q}!A:A").execute().get("values", [])
    last = len(col_a)                    # 1-based index of last row with a Ticket ID
    grid_rows = props["gridProperties"]["rowCount"]
    first_new = last + 1
    end = last + len(rows)
    if end > grid_rows:
        svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": [
            {"appendDimension": {"sheetId": gid, "dimension": "ROWS", "length": end - grid_rows}}]}).execute()

    src = {"sheetId": gid, "startRowIndex": last - 1, "endRowIndex": last}
    dst = {"sheetId": gid, "startRowIndex": first_new - 1, "endRowIndex": end}
    reqs = [
        # formats (date columns etc.) for A:AD from the last data row
        {"copyPaste": {"source": {**src, "startColumnIndex": 0, "endColumnIndex": NCOL},
                       "destination": {**dst, "startColumnIndex": 0, "endColumnIndex": NCOL},
                       "pasteType": "PASTE_FORMAT"}},
    ]
    if ncols_grid > NCOL:
        # formula columns AE~ : copy formulas + format down (relative refs shift per row)
        reqs.append({"copyPaste": {"source": {**src, "startColumnIndex": NCOL, "endColumnIndex": ncols_grid},
                                   "destination": {**dst, "startColumnIndex": NCOL, "endColumnIndex": ncols_grid},
                                   "pasteType": "PASTE_NORMAL"}})
    svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": reqs}).execute()

    # RAW so IDs/order numbers/"3-5" are never auto-parsed; dates are pre-converted serials.
    for i in range(0, len(rows), 5000):
        chunk = rows[i:i + 5000]
        r0 = first_new + i
        svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"{q}!A{r0}:AD{r0 + len(chunk) - 1}",
            valueInputOption="RAW", body={"values": chunk}).execute()
    return first_new, end


# ---------------------------------------------------------------- Google Chat notify

def kst_stamp(dt):
    """10/1(수) 09:02"""
    return f"{dt.month}/{dt.day}({'월화수목금토일'[dt.weekday()]}) {dt:%H:%M}"


def sheet_url(props):
    return f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={props['sheetId']}"


def notify(text):
    """Post to the Google Chat webhook from credentials.json (optional; never fails the run)."""
    url = (load_json(CREDS_PATH, {}) or {}).get("chat_webhook")
    if not url:
        return
    req = urllib.request.Request(url, data=json.dumps({"text": text}).encode(), method="POST",
                                 headers={"Content-Type": "application/json; charset=UTF-8"})
    try:
        urllib.request.urlopen(req, timeout=30).read()
        log("Google Chat notified")
    except Exception as e:
        log(f"Google Chat notify failed: {e}")


# ---------------------------------------------------------------- commands

def cmd_run(args):
    os.makedirs(CFG_DIR, exist_ok=True)
    state = load_json(STATE_PATH, {})
    now = datetime.datetime.now(KST)
    if args.scheduled:
        sched = load_json(SCHED_PATH)
        if not sched:
            return log("no schedule configured — skip")
        args.until_yesterday = args.until_yesterday or sched.get("mode") == "until_yesterday"
        today = now.strftime("%Y-%m-%d")
        if DAYS[now.weekday()] not in sched["days"] or now.strftime("%H:%M") < sched["time"]:
            return
        if state.get("last_scheduled_run") == today:
            return

    lock = open(LOCK_PATH, "w")
    try:
        if IS_WIN:
            msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (BlockingIOError, OSError):
        return log("another run is in progress — skip")

    svc = sheets_service()
    props = sheet_props(svc)
    q = f"'{props['title']}'"
    header = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{q}!A1:AD1").execute()["values"][0]
    if header != HEADER:
        diff = [(i, a, b) for i, (a, b) in enumerate(zip(HEADER, header)) if a != b]
        sys.exit(f"Header A1:AD1 changed — refusing to write. Diffs (idx, expected, found): {diff or (len(HEADER), len(header))}")

    col_a = svc.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{q}!A2:A").execute().get("values", [])
    existing = {r[0].strip() for r in col_a if r and r[0].strip()}
    log(f"sheet '{props['title']}': {len(existing)} existing Ticket IDs")

    statuses = {"solved"} if args.solved_only else STATUSES
    created_max = (now.date() - datetime.timedelta(days=1)).isoformat() if args.until_yesterday else None
    if created_max:
        log(f"created date filter: ≤ {created_max}")
    rows, stats = fetch_candidates(existing, statuses, created_max)
    log(f"stats: {stats} → {len(rows)} new {'/'.join(sorted(statuses))} tickets to append")

    if args.dry_run:
        for r in rows[:10]:
            log("  would append:", r[:8])
        return log("dry run — nothing written")

    if rows:
        first, end = append_rows(svc, props, rows)
        log(f"appended rows {first}–{end} (Ticket IDs {rows[0][0]}…{rows[-1][0]})")
    md = lambda d: f"{int(d[5:7])}/{int(d[8:])}"
    scope = f"Solved·Closed, ~{md(created_max)} 생성분" if created_max else "Solved·Closed"
    msg = ["*✅ Zendesk Raw Data 업데이트 완료*", "", f"• 실행: {kst_stamp(now)}"]
    days = intake_days(created_max, sched["days"]) if created_max and args.scheduled else []
    intake = intake_counts(days) if days else None
    if intake:
        span = md(days[0]) if len(days) == 1 else f"{md(days[0])}~{md(days[-1])}"
        msg.append(f"• {span} 인입: {intake[0]}건 (완료 {intake[1]} / 처리 중 {intake[2]})")
    msg += [f"• 시트 추가: {len(rows)}건 ({scope})" if rows else f"• 시트 추가: 없음 ({scope})",
            "", f"📊 <{sheet_url(props)}|{props['title']} 바로가기>"]
    state.update({"last_run": now.isoformat(timespec="seconds"), "last_appended": len(rows)})
    if args.scheduled:
        state["last_scheduled_run"] = now.strftime("%Y-%m-%d")
    save_json(STATE_PATH, state)
    if args.scheduled:
        notify("\n".join(msg))


def cmd_setup(args):
    os.makedirs(CFG_DIR, exist_ok=True)
    creds = load_json(CREDS_PATH, {})
    creds["caspi_api_key"] = args.caspi_key or creds.get("caspi_api_key") or input("Caspi API key (ak_…): ").strip()
    creds["caspi_query_id"] = args.query_id or creds.get("caspi_query_id") or input("Caspi queryId (pq_…): ").strip()
    if args.chat_webhook:
        creds["chat_webhook"] = args.chat_webhook
    if args.intake_query_id:
        creds["caspi_intake_query_id"] = args.intake_query_id
    save_json(CREDS_PATH, creds)
    log(f"Caspi credentials saved → {CREDS_PATH}")

    if load_json(GTOKEN_PATH) and not args.client_secret:
        return log(f"Google token already present → {GTOKEN_PATH}")
    if not args.client_secret:
        sys.exit("Pass --client-secret <OAuth desktop client JSON> to authorize Google Sheets "
                 "(ask the repo owner for the GCP 'gcxbot' desktop client file).")
    from google_auth_oauthlib.flow import InstalledAppFlow
    flow = InstalledAppFlow.from_client_secrets_file(
        args.client_secret, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    c = flow.run_local_server(port=0)
    save_json(GTOKEN_PATH, json.loads(c.to_json()))
    log(f"Google token saved → {GTOKEN_PATH}")


def cmd_schedule(args):
    days = [d.strip().lower()[:3] for d in args.days.split(",") if d.strip()]
    if args.days.strip().lower() in ("daily", "everyday"):
        days = DAYS[:]
    elif args.days.strip().lower() == "weekdays":
        days = DAYS[:5]
    bad = [d for d in days if d not in DAYS]
    if bad or not days:
        sys.exit(f"--days must be comma-separated from {DAYS} (or daily / weekdays); got {bad or args.days}")
    hh, mm = args.time.split(":")
    t = f"{int(hh):02d}:{int(mm):02d}"
    save_json(SCHED_PATH, {"days": days, "time": t, "timezone": "Asia/Seoul",
                           "mode": "until_yesterday" if args.until_yesterday else "all"})

    os.makedirs(LOG_DIR, exist_ok=True)
    if IS_WIN:
        return schedule_windows(days, t)
    args_xml = "".join(f"<string>{escape(a)}</string>" for a in
                       [sys.executable, os.path.realpath(__file__), "run", "--scheduled"])
    # StartInterval 1800: tick every 30 min; `run --scheduled` self-gates on days/time and
    # runs once per scheduled day → a sleeping Mac catches up on wake.
    plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>{LABEL}</string>
  <key>ProgramArguments</key><array>{args_xml}</array>
  <key>EnvironmentVariables</key><dict><key>PATH</key><string>/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin</string></dict>
  <key>StartInterval</key><integer>1800</integer>
  <key>RunAtLoad</key><true/>
  <key>StandardOutPath</key><string>{escape(os.path.join(LOG_DIR, "sync.log"))}</string>
  <key>StandardErrorPath</key><string>{escape(os.path.join(LOG_DIR, "sync.err.log"))}</string>
</dict></plist>
"""
    os.makedirs(os.path.dirname(PLIST), exist_ok=True)
    domain = f"gui/{os.getuid()}"
    subprocess.run(["launchctl", "bootout", f"{domain}/{LABEL}"], capture_output=True)
    with open(PLIST, "w") as f:
        f.write(plist)
    subprocess.run(["launchctl", "bootstrap", domain, PLIST], check=True)
    log(f"scheduled: {', '.join(days)} at {t} KST ({len(days)}x/week) → {PLIST}")


def schedule_windows(days, t):
    # Task Scheduler: tick every TICK_MIN min (like launchd StartInterval); `run --scheduled`
    # self-gates on days/time, so a PC that was off/asleep catches up the same day.
    # pythonw = no console window; output goes to LOG_DIR/sync.log (see main()).
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pyw if os.path.exists(pyw) else sys.executable
    tr = f'"{exe}" "{os.path.realpath(__file__)}" run --scheduled'
    subprocess.run(["schtasks", "/Delete", "/TN", WIN_TASK, "/F"], capture_output=True)
    subprocess.run(["schtasks", "/Create", "/TN", WIN_TASK, "/TR", tr, "/SC", "MINUTE",
                    "/MO", str(TICK_MIN), "/F"], check=True, capture_output=True)
    # schtasks defaults skip runs on battery (laptops) and don't catch up missed triggers
    ps = (f"$s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries "
          f"-StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew; "
          f"Set-ScheduledTask -TaskName '{WIN_TASK}' -Settings $s | Out-Null")
    subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=True, capture_output=True)
    log(f"scheduled: {', '.join(days)} at {t} KST ({len(days)}x/week) → Task Scheduler '{WIN_TASK}'")


def cmd_unschedule(args):
    if IS_WIN:
        subprocess.run(["schtasks", "/Delete", "/TN", WIN_TASK, "/F"], capture_output=True)
        if os.path.exists(SCHED_PATH):
            os.remove(SCHED_PATH)
        return log("schedule removed")
    subprocess.run(["launchctl", "bootout", f"gui/{os.getuid()}/{LABEL}"], capture_output=True)
    if os.path.exists(PLIST):
        os.remove(PLIST)
    if os.path.exists(SCHED_PATH):
        os.remove(SCHED_PATH)
    log("schedule removed")


def cmd_status(args):
    print("credentials :", "OK" if (load_json(CREDS_PATH) or {}).get("caspi_api_key") or os.environ.get("CASPI_API_KEY") else "MISSING", CREDS_PATH)
    print("google token:", "OK" if load_json(os.environ.get("GOOGLE_TOKEN_PATH") or GTOKEN_PATH) or load_json(GWS_SHIM_TOKEN) else "MISSING")
    print("schedule    :", load_json(SCHED_PATH) or "none")
    if IS_WIN:
        loaded = subprocess.run(["schtasks", "/Query", "/TN", WIN_TASK], capture_output=True).returncode == 0
        print("task sched  :", "registered" if loaded else "not registered")
    else:
        loaded = subprocess.run(["launchctl", "print", f"gui/{os.getuid()}/{LABEL}"], capture_output=True).returncode == 0
        print("launchd     :", "loaded" if loaded else "not loaded")
    print("state       :", load_json(STATE_PATH, {}))
    print("logs        :", LOG_DIR)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run")
    r.add_argument("--dry-run", action="store_true"); r.add_argument("--scheduled", action="store_true")
    r.add_argument("--solved-only", action="store_true", help="strict: exclude Closed tickets")
    r.add_argument("--until-yesterday", action="store_true", help="only tickets created up to yesterday")
    s = sub.add_parser("setup")
    s.add_argument("--caspi-key"); s.add_argument("--query-id"); s.add_argument("--client-secret")
    s.add_argument("--chat-webhook", help="Google Chat incoming webhook URL for run notifications")
    s.add_argument("--intake-query-id", help="Caspi queryId of intake_query.sql (yesterday's intake line in the notice)")
    sc = sub.add_parser("schedule")
    sc.add_argument("--days", required=True, help="e.g. mon,thu | weekdays | daily")
    sc.add_argument("--time", required=True, help="HH:MM, KST")
    sc.add_argument("--until-yesterday", action="store_true", help="scheduled runs use `run --until-yesterday`")
    sub.add_parser("unschedule"); sub.add_parser("status")
    a = p.parse_args()
    if IS_WIN and a.cmd == "run" and a.scheduled:
        # pythonw has no console: send output to the log file (launchd does this on macOS)
        os.makedirs(LOG_DIR, exist_ok=True)
        sys.stdout = sys.stderr = open(os.path.join(LOG_DIR, "sync.log"), "a", encoding="utf-8")
    try:
        {"run": cmd_run, "setup": cmd_setup, "schedule": cmd_schedule,
         "unschedule": cmd_unschedule, "status": cmd_status}[a.cmd](a)
    except (Exception, SystemExit) as e:
        if a.cmd == "run" and a.scheduled and not (isinstance(e, SystemExit) and e.code in (None, 0)):
            # the job retries every TICK_MIN min until it succeeds — notify only once per day
            state, today = load_json(STATE_PATH, {}), datetime.datetime.now(KST).strftime("%Y-%m-%d")
            if state.get("last_fail_notified") != today:
                notify("\n".join(["*⚠️ Zendesk Raw Data 업데이트 실패*", "",
                                  f"📅 {kst_stamp(datetime.datetime.now(KST))}",
                                  f"❗ {str(e)[:300]}",
                                  f"🔁 {TICK_MIN}분마다 자동 재시도 중 · 로그: {LOG_DIR}"]))
                state["last_fail_notified"] = today
                save_json(STATE_PATH, state)
        raise


if __name__ == "__main__":
    main()
