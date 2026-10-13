#!/usr/bin/env python3
"""Weekday gate for the daily ASIN rating/review-count fetch job.

Called every 5 min by cron (not a single exact-time shot) so a missed slot
(server was down/asleep at the target time) self-catches-up the moment the
machine is back, same pattern as sc-review-propagate's daily_cron_gate.py
and the Apify actor's health_check_gate.py. Only actually launches once per
weekday (state file below), from TARGET_TIME onward.

This job is INTENTIONALLY independent of sc-review-propagate's own schedule
and state file — it does not check, wait for, or depend on that job having
run. It is its own standalone daily job.
"""
import json
import os
import subprocess
import sys
from datetime import date, datetime, time as dtime
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")
TARGET_TIME = dtime(9, 0)
HERE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(HERE, "state_rating_gate.json")
LOG_PATH = os.path.expanduser("~/.config/asin_rating_daily/gate.log")
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

sys.path.insert(0, os.path.expanduser("~/Desktop/GCX/ServerBootstrap/watchdog"))


def log(msg):
    line = f"{datetime.now():%F %T} {msg}"
    print(line)
    with open(LOG_PATH, "a") as f:
        f.write(line + "\n")


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
            "ASIN Rating Daily 지연 실행 (catch-up)",
            f"오늘({today_str}) {TARGET_TIME.strftime('%H:%M')} 예정이던 자동 실행이 "
            f"서버 다운/절전 등으로 누락되어, {now.strftime('%H:%M')} KST "
            f"({late_minutes:.0f}분 지연)에 catch-up 으로 실행합니다.",
            level="warn",
        )
    except Exception as e:
        print(f"WARN: catch-up alert failed: {e}", file=sys.stderr)


def run_job(run_label):
    script = os.path.join(HERE, "rating_fetch_daily.py")
    # No -I: on gcx-server the Google client libs are user-site installs, which -I hides.
    subprocess.run([sys.executable, script, "--commit", "--run-label", run_label],
                    check=False)


def main():
    now = datetime.now(KST)
    today = now.date()
    today_str = today.isoformat()

    if today.weekday() >= 5:
        print(f"{today} skipped (weekend)")
        return

    if now.time() < TARGET_TIME:
        print(f"{today} not yet {TARGET_TIME} KST, skipping this tick")
        return

    state = load_state()
    if state.get("last_run_date") == today_str:
        print(f"{today} already launched today, skipping")
        return

    target_dt = datetime.combine(today, TARGET_TIME, tzinfo=KST)
    late_minutes = (now - target_dt).total_seconds() / 60
    is_catchup = late_minutes > 10

    state["last_run_date"] = today_str
    state["last_run_at"] = now.isoformat()
    state["was_catchup"] = is_catchup
    save_state(state)

    if is_catchup:
        log(f"{today} CATCH-UP launch ({late_minutes:.0f} min late)")
        notify_catchup(today_str, now, late_minutes)
        run_job(f"daily catch-up ({late_minutes:.0f}min late)")
    else:
        log(f"{today} on-time launch")
        run_job("daily")


if __name__ == "__main__":
    main()
