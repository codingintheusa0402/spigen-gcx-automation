#!/usr/bin/env python3
"""Catch-up gate for the daily Apify actor health check.

Called every 5 min by cron (not a single exact-time shot) so a missed slot
(server was down/asleep at 10:00) self-catches-up the moment the machine is
back, instead of silently skipping the whole day — added 2026-10-12 after a
27h Modern Standby outage ate both this job's 10:00 run and
sc-review-propagate's 8:30 run on the same day (2026-10-11). Only actually
runs health_check.py once per day (state_cron_gate.json), from 10:00 KST
onward. Runs every day (not just weekdays) — matches the original crontab.
"""
import json
import os
import subprocess
import sys
from datetime import date, datetime, time as dtime
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")
HERE = os.path.dirname(os.path.abspath(__file__))
TARGET_TIME = dtime(10, 0)
STATE_PATH = os.path.join(HERE, "state_cron_gate.json")
VENV_PYTHON = os.path.expanduser("~/.venvs/apify-asin/bin/python3")
HEALTH_SCRIPT = os.path.join(HERE, "health_check.py")
CRON_LOG = os.path.expanduser("~/.config/apify_asin_scraper/cron.log")

sys.path.insert(0, os.path.expanduser("~/Desktop/GCX/ServerBootstrap/watchdog"))


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
            "Apify 헬스체크 지연 실행 (catch-up)",
            f"오늘({today_str}) 10:00 예정이던 자동 헬스체크가 서버 다운/절전 등으로 누락되어, "
            f"{now.strftime('%H:%M')} KST ({late_minutes:.0f}분 지연)에 catch-up 으로 실행합니다.",
            level="warn",
        )
    except Exception as e:
        print(f"WARN: catch-up chat alert failed: {e}", file=sys.stderr)


def main():
    now = datetime.now(KST)
    today = now.date()
    today_str = today.isoformat()

    if now.time() < TARGET_TIME:
        print(f"{today} not yet {TARGET_TIME} KST, skipping this tick")
        return

    state = load_state()
    if state.get("last_run_date") == today_str:
        print(f"{today} already ran today, skipping")
        return

    target_dt = datetime.combine(today, TARGET_TIME, tzinfo=KST)
    late_minutes = (now - target_dt).total_seconds() / 60
    is_catchup = late_minutes > 10

    state["last_run_date"] = today_str
    state["last_run_at"] = now.isoformat()
    state["was_catchup"] = is_catchup
    save_state(state)

    if is_catchup:
        print(f"{today} CATCH-UP run ({late_minutes:.0f} min late) — alerting Chat")
        notify_catchup(today_str, now, late_minutes)
    else:
        print(f"{today} on-time run")

    with open(CRON_LOG, "a") as logf:
        subprocess.run([VENV_PYTHON, HEALTH_SCRIPT, "--autofix"], stdout=logf, stderr=subprocess.STDOUT, cwd=HERE)


if __name__ == "__main__":
    main()
