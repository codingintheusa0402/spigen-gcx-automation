#!/usr/bin/env python3
"""Weekday + KR-holiday gate for the daily sc-review-propagate cron job.

Called every 5 min by cron (not a single exact-time shot) so a missed slot
(server was down/asleep at 8:30) self-catches-up the moment the machine is
back, instead of silently skipping the whole day — added 2026-10-12 after a
27h Modern Standby outage ate both this job's 8:30 run and the Apify health
check's 10:00 run on the same day. Only actually launches once per working
day (state_cron_gate.json, git-synced between hosts), from 08:30 KST onward.

2026-10-14: the Mac runs this same gate too (launchd com.spigen.gcx.sc-propagate-gate,
every 5 min). Ownership per day: the server, unless it is unreachable over Tailscale
ssh at launch time, or unless the state file's "force_mac_dates" lists today — then
the Mac launches and the server skips. The shared state file keeps it to one launch.

Skips on Korean public holidays, computed via the `holidays` package (South
Korea, "public" category) in the dedicated venv ~/.venvs/gcx-holidays —
EXCEPT 제헌절 (Constitution Day, Jul 17), which the library mislabels: it has
not been a non-working public holiday since 2008.

If today is a working day and we're past 08:30 and haven't already launched
today, launches the daily pipeline as a fresh Claude Code session in tmux
window gcx:sc-propagate-daily (same mechanism as run-on-server's
run_on_server.sh), which fetches from Caspi, runs propagate.py through all
phases, then phase_g.py --commit for the image backfill — architected
exactly per the sc-review-propagate skill.
"""
import json
import os
import subprocess
import sys
from datetime import date, datetime, time as dtime
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")
IS_MAC = sys.platform == "darwin"          # 2026-10-14: the same gate also runs on the Mac (launchd, every 5 min)
# 2026-10-17: the Mac mini (user kevin) is the host; the MacBook is fallback only; the Linux laptop server is retired.
IS_MINI = IS_MAC and os.path.expanduser("~") == "/Users/kevin"
HOST_LABEL = "Mac mini server" if IS_MINI else ("MacBook" if IS_MAC else "GCX server")
PROJECT_DIR = os.path.expanduser("~").replace("/", "-")
EXCLUDE_HOLIDAYS = {"07-17"}  # 제헌절 — no longer a non-working day since 2008
TASK_NAME = "sc-propagate-daily"
TASK_DIR = os.path.expanduser("~/.gcx-tasks")
TARGET_TIME = dtime(8, 30)
STATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "state_cron_gate.json")

sys.path.insert(0, os.path.expanduser("~/Desktop/GCX/ServerBootstrap/watchdog"))

PROMPT = """Run today's full sc-review-propagate pipeline end to end, autonomously, no confirmation needed. This is the standing daily routine (see memory sc_master_sheet_propagation.md, caspilm_review_fetch_workflow.md, and ~/Desktop/GCX/Scrapers/SC_Master_Propagate/SKILL.md — read them, don't re-derive the approach). You are running on the __HOST__ via a cron-triggered session with no prior context, so figure out today's date yourself first.

1. Get today's date in KST (`TZ=Asia/Seoul date +%y%m%d` and `+%Y-%m-%d`).

2. **Freshness check first**: `SELECT MAX(REVIEW_CREATED_AT), MAX(LOAD_TS), COUNT(*) FROM SQ.TAHOE.AMAZON_REVIEWS` via the CaspiLM MCP run_query tool. If the newest review is more than 2 days old, or LOAD_TS is more than 24h old, this means Caspi's data is stale (this has happened before — a multi-day ingestion stall). Still proceed with the fetch (it'll just find fewer/no new rows), but flag the staleness explicitly and prominently in the Phase E Chat notification text so a human notices, same as a real finding would be reported — don't silently note it only in your own output.

2b. **Already-ran check (added 2026-10-13)**: another scheduler (the Mac) may already have run this same pipeline today. Before writing anything, list the source spreadsheet's tabs and `drive.revisions.list` on `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw` for today's kjw@spigen.com revisions. If `CaspiLM_<yymmdd>` already exists and today's revisions show the funnel/cleanup already happened, SKIP every phase that is already complete — above all do NOT re-post the Phase E Google Chat card — and only verify (`--all-products` = 0 pending, `--refresh-tem --commit` is always safe, `phase_g.py` dry run), then report what you skipped and why.

3. **CaspiLM fetch + upload (steps 3 and 4 are one script)**: `cd ~/Desktop/GCX/Scrapers/Caspi_Review_Fetch && python3 fetch_caspi_reviews.py` — it pulls the latest 500 reviews per marketplace (US/DE/IT/FR/ES/UK/JP/IN, Spigen, SC-scraper 14-column schema, photo URLs `|`-joined, GB→UK) through the registered Caspi data-api query, dedupes by Review ID, writes a CSV under `out/`, and uploads the new tab `CaspiLM_<yymmdd>` (header included) to spreadsheet `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw` via gws_shim OAuth. It reads its own API key — NEVER run shell commands that mention the api_key path (not even `test -f`); the auto-mode classifier blocks them and the session hangs for hours. Expect 4,000 rows (500 × 8). Only if the script itself fails, fall back to the CaspiLM MCP `run_query` (one call per marketplace, `LIMIT 500`; large results land as JSON files under `~/.claude/projects/__PROJECT_DIR__/<session>/tool-results/`) and upload with googleapiclient.

4. (merged into step 3 — the script uploads the tab itself.)

5. **Run the full propagate.py pipeline** from ~/Desktop/GCX/Scrapers/SC_Master_Propagate/ (all via Bash, these are plain scripts — no need to drive them interactively):
   - `python3 propagate.py --new-sheet CaspiLM_<yymmdd>` (dry run) then `--commit` (Phase A)
   - `python3 propagate.py --all-products` (dry-run worklist)
   - `python3 propagate.py --product X --commit` for each of the 8 active products, append-at-bottom books first (GlxZ8, Pixel11, iPh18, SDA, Auto_Acc, Power_Acc, 전략폰), then 유지훈P last (insert-at-top)
   - `python3 propagate.py --refresh-tem --commit` (Phase C)
   - `python3 propagate.py --all-products` again — must end at 0 pending everywhere; if a product shows new pending on a second check (known Sheets-recalc lag), commit it and re-check once more
   - `python3 propagate.py --cleanup --notify --new-sheet CaspiLM_<yymmdd> --commit` (Phase F + E — posts the completion card to the GCX Chat webhook; mention the freshness-staleness flag here too if applicable, per step 2)

5b. **Caspi stale → also run the live SC scraper (user rule 2026-10-15: "if no fresher exist, use sc scraper")**: if step 2 found Caspi stale (newest review > 2 days old or no load in 24h), after the Caspi pass run `cd ~/Desktop/GCX/Scrapers/SC_Review_Scraper && python3 scrape_sc_reviews.py < /dev/null` (all markets, defaults; it uploads `SC_<yymmdd>`), then repeat the step-5 propagate sequence with `--new-sheet SC_<yymmdd>` (Phase A → B → C → pending check), and run the single Phase F+E `--cleanup --notify --commit` with `--new-sheet SC_<yymmdd>` and `--alert "Caspi 데이터 미갱신 → 로컬 SC 스크레이퍼로 수집"` ONCE at the very end (never post the card twice). On the Mac mini the scraper attaches to the unified Amazon Chrome (CDP :9223, launchd com.gcx.amazon-chrome) — never launch another browser/profile. If the scraper log shows "Not logged in" / LOGIN REDIRECT / `selection_required` for a marketplace, do not wait for it: stop the scraper, finish with the markets that worked, and say in the card alert which Seller Central login the user must redo via Screen Sharing.

6. **Phase G — already scripted, just run it**: `python3 phase_g.py --commit` from the same directory (reads today's KST date automatically; override with `--date <yyyy-mm-dd>` only if needed). It's a plain Playwright script (on the Mac mini it attaches to the unified Amazon Chrome on :9223). If it reports amazon.es rows skipped (needs login), just note that in your final summary — never log in yourself.

7. If you hit any propagate.py or phase_g.py bug, fix it, verify live, and note it in your summary. If you're unsure of an existing live formula/structure you're about to depend on, verify it live rather than assuming from memory notes, since those can go stale.

Keep your final chat-visible summary concise — the real report is the Chat card Phase E already sends; you don't need to repeat it all, just confirm completion or flag anything that needs human attention (staleness, login walls, unresolved errors).
"""


def is_working_day(today: date) -> bool:
    if today.weekday() >= 5:  # Sat/Sun
        return False
    mmdd = today.strftime("%m-%d")
    if mmdd in EXCLUDE_HOLIDAYS:
        return True  # explicitly override the library's mislabeled entry
    try:
        import holidays
        kr = holidays.SouthKorea(years=today.year, categories=("public",))
    except ImportError:
        print("WARN: holidays package not available, skipping holiday check", file=sys.stderr)
        return True
    return today not in kr


def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH) as f:
            return json.load(f)
    return {}


def save_state(state):
    with open(STATE_PATH, "w") as f:
        json.dump(state, f)


def notify_catchup(today_str, now, late_minutes):
    try:
        from gcx_alert import alert
        alert(
            "sc-review-propagate 지연 실행 (catch-up)",
            f"오늘({today_str}) 08:30 예정이던 자동 실행이 서버 다운/절전 등으로 누락되어, "
            f"{now.strftime('%H:%M')} KST ({late_minutes:.0f}분 지연)에 catch-up 으로 실행을 시작합니다.",
            level="warn",
        )
    except Exception as e:
        print(f"WARN: catch-up chat alert failed: {e}", file=sys.stderr)


def mini_reachable():
    """Is the Mac mini host up? Probed from the MacBook over Tailscale ssh (~/.ssh/config host gcx-macmini)."""
    try:
        return subprocess.run(["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=8", "gcx-macmini", "true"],
                              capture_output=True, timeout=25).returncode == 0
    except Exception:
        return False


def owner_is_other_host(state, today_str):
    """Both hosts share this gate (and the git-synced state file). Exactly one of them launches per day:
    the server unless it is unreachable at launch time, or unless state["force_mac_dates"] lists today
    (set 2026-10-14 by the user: "use my mac tomorrow"). Returns a skip reason, or None to launch here."""
    # 2026-10-17: the Mac mini owns every day; the MacBook launches only if force_mac_dates names today or the
    # mini is unreachable; the retired Linux laptop server never launches.
    forced = today_str in (state.get("force_mac_dates") or [])
    if IS_MINI:
        return "force_mac_dates names today → the MacBook owns today's run, mini skips" if forced else None
    if IS_MAC:
        if forced:
            return None
        if mini_reachable():
            return "Mac mini reachable → the mini owns today's run, MacBook skips"
        return None
    return "Linux laptop server retired 2026-10-16 → the Mac mini owns the run"


def launch_pipeline():
    os.makedirs(TASK_DIR, exist_ok=True)
    prompt_path = os.path.join(TASK_DIR, f"{TASK_NAME}.prompt")
    with open(prompt_path, "w") as f:
        f.write(PROMPT.replace("__HOST__", HOST_LABEL).replace("__PROJECT_DIR__", PROJECT_DIR))

    if IS_MAC:
        launch_cmd = f"""
export PATH=$HOME/.local/bin:/opt/homebrew/bin:$PATH
tmux has-session -t gcx 2>/dev/null || tmux new-session -d -s gcx -c "$HOME"
tmux kill-window -t gcx:{TASK_NAME} 2>/dev/null
tmux new-window -t gcx: -n {TASK_NAME} 'cd ~ && export PATH=$HOME/.local/bin:/opt/homebrew/bin:$PATH; claude --remote-control gcx-{TASK_NAME}-{"mini" if IS_MINI else "mac"} "$(cat {prompt_path})"; echo; echo "[session ended]"; exec zsh'
"""
    else:
        launch_cmd = f"""
bash ~/tmux_start.sh >/dev/null 2>&1
tmux kill-window -t gcx:{TASK_NAME} 2>/dev/null
tmux new-window -t gcx: -n {TASK_NAME} 'cd ~ && export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0; claude --remote-control gcx-{TASK_NAME} "$(cat {prompt_path})"; echo; echo "[session ended]"; exec bash'
"""
    subprocess.run(["bash", "-c", launch_cmd], check=True)
    print(f"launched gcx:{TASK_NAME}")


def main():
    now = datetime.now(KST)
    today = now.date()
    today_str = today.isoformat()

    if not is_working_day(today):
        reason = "weekend" if today.weekday() >= 5 else "KR public holiday"
        print(f"{today} skipped ({reason})")
        return

    if now.time() < TARGET_TIME:
        print(f"{today} not yet {TARGET_TIME} KST, skipping this tick")
        return

    state = load_state()
    if state.get("last_run_date") == today_str:
        print(f"{today} already launched today, skipping")
        return
    skip = owner_is_other_host(state, today_str)
    if skip:
        print(f"{today} {skip}")
        return

    # Catch-up detection: are we more than ~10 min past the target time?
    target_dt = datetime.combine(today, TARGET_TIME, tzinfo=KST)
    late_minutes = (now - target_dt).total_seconds() / 60
    is_catchup = late_minutes > 10

    state["last_run_date"] = today_str
    state["last_run_at"] = now.isoformat()
    state["was_catchup"] = is_catchup
    state["last_run_host"] = HOST_LABEL
    save_state(state)

    if is_catchup:
        print(f"{today} CATCH-UP launch ({late_minutes:.0f} min late) — alerting Chat")
        notify_catchup(today_str, now, late_minutes)
    else:
        print(f"{today} on-time launch")

    try:
        launch_pipeline()
    except Exception:
        # Launch failed -> un-mark today so the next 5-min tick retries instead of silently skipping.
        state["last_run_date"] = None
        save_state(state)
        raise


if __name__ == "__main__":
    main()
