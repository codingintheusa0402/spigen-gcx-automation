#!/usr/bin/env python3
"""Weekday + KR-holiday gate for the daily sc-review-propagate cron job.

Runs on the GCX server (cron, weekdays 8:30 AM KST — the day-of-week filter
is already in the crontab entry itself; this script's own weekday check is
just defense-in-depth). Skips on Korean public holidays, computed via the
`holidays` package (South Korea, "public" category) in the dedicated venv
~/.venvs/gcx-holidays — EXCEPT 제헌절 (Constitution Day, Jul 17), which the
library mislabels: it has not been a non-working public holiday since 2008.

If today is a working day, launches the daily pipeline as a fresh Claude
Code session in tmux window gcx:sc-propagate-daily (same mechanism as
run-on-server's run_on_server.sh), which fetches from Caspi, runs
propagate.py through all phases, then phase_g.py --commit for the image
backfill — architected exactly per the sc-review-propagate skill.
"""
import os
import subprocess
import sys
from datetime import date
from zoneinfo import ZoneInfo
from datetime import datetime

KST = ZoneInfo("Asia/Seoul")
EXCLUDE_HOLIDAYS = {"07-17"}  # 제헌절 — no longer a non-working day since 2008
TASK_NAME = "sc-propagate-daily"
TASK_DIR = os.path.expanduser("~/.gcx-tasks")

PROMPT = """Run today's full sc-review-propagate pipeline end to end, autonomously, no confirmation needed. This is the standing daily routine (see memory sc_master_sheet_propagation.md, caspilm_review_fetch_workflow.md, and ~/Desktop/GCX/Scrapers/SC_Master_Propagate/SKILL.md — read them, don't re-derive the approach). You are running on the GCX server via a cron-triggered session with no prior context, so figure out today's date yourself first.

1. Get today's date in KST (`TZ=Asia/Seoul date +%y%m%d` and `+%Y-%m-%d`).

2. **Freshness check first**: `SELECT MAX(REVIEW_CREATED_AT), MAX(LOAD_TS), COUNT(*) FROM SQ.TAHOE.AMAZON_REVIEWS` via the CaspiLM MCP run_query tool. If the newest review is more than 2 days old, or LOAD_TS is more than 24h old, this means Caspi's data is stale (this has happened before — a multi-day ingestion stall). Still proceed with the fetch (it'll just find fewer/no new rows), but flag the staleness explicitly and prominently in the Phase E Chat notification text so a human notices, same as a real finding would be reported — don't silently note it only in your own output.

3. **CaspiLM fetch**: latest 500 reviews per marketplace for the 8 in-scope markets US/DE/IT/FR/ES/GB/JP/IN, BRAND_NAME='Spigen'. Standard 14-column schema (ASIN | Created 날짜 | 사진 유무 | Reviewer | Review Ratings | Review Title | 본문 | 국가 | Review Link | Image URL | Review ID | Order ID | Product Rating | Ratings Count), LISTAGG(PHOTO_URL,'|') join on AMAZON_REVIEW_PHOTOS via REVIEW_ID. GB → 국가 "UK". Split per-marketplace and further by hash/offset if truncated — check actual response sizes, don't assume a fixed split works; verify rowCount/truncated on every slice; dedupe by REVIEW_ID when combining.

4. Upload the combined dataset as a new tab `CaspiLM_<yymmdd>` in spreadsheet `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw` via gws_shim OAuth (~/.config/gws_shim/token.json) + googleapiclient, header row included.

5. **Run the full propagate.py pipeline** from ~/Desktop/GCX/Scrapers/SC_Master_Propagate/ (all via Bash, these are plain scripts — no need to drive them interactively):
   - `python3 propagate.py --new-sheet CaspiLM_<yymmdd>` (dry run) then `--commit` (Phase A)
   - `python3 propagate.py --all-products` (dry-run worklist)
   - `python3 propagate.py --product X --commit` for each of the 8 active products, append-at-bottom books first (GlxZ8, Pixel11, iPh18, SDA, Auto_Acc, Power_Acc, 전략폰), then 유지훈P last (insert-at-top)
   - `python3 propagate.py --refresh-tem --commit` (Phase C)
   - `python3 propagate.py --all-products` again — must end at 0 pending everywhere; if a product shows new pending on a second check (known Sheets-recalc lag), commit it and re-check once more
   - `python3 propagate.py --cleanup --notify --new-sheet CaspiLM_<yymmdd> --commit` (Phase F + E — posts the completion card to the GCX Chat webhook; mention the freshness-staleness flag here too if applicable, per step 2)

6. **Phase G — already scripted, just run it**: `python3 phase_g.py --commit` from the same directory (reads today's KST date automatically; override with `--date <yyyy-mm-dd>` only if needed). This replaces the old claude-in-chrome-driven version — it's a plain Playwright script using the dedicated ~/.chrome-phaseg-profile. If it reports amazon.es rows skipped (needs login), just note that in your final summary — never log in yourself.

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


def launch_pipeline():
    os.makedirs(TASK_DIR, exist_ok=True)
    prompt_path = os.path.join(TASK_DIR, f"{TASK_NAME}.prompt")
    with open(prompt_path, "w") as f:
        f.write(PROMPT)

    launch_cmd = f"""
bash ~/tmux_start.sh >/dev/null 2>&1
tmux kill-window -t gcx:{TASK_NAME} 2>/dev/null
tmux new-window -t gcx -n {TASK_NAME} 'cd ~ && export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0; claude --remote-control gcx-{TASK_NAME} "$(cat {prompt_path})"; echo; echo "[session ended]"; exec bash'
"""
    subprocess.run(["bash", "-c", launch_cmd], check=True)
    print(f"launched gcx:{TASK_NAME}")


def main():
    today = datetime.now(KST).date()
    if not is_working_day(today):
        reason = "weekend" if today.weekday() >= 5 else "KR public holiday"
        print(f"{today} skipped ({reason})")
        return
    print(f"{today} is a working day — launching pipeline")
    launch_pipeline()


if __name__ == "__main__":
    main()
