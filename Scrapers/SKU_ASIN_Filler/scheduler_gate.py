"""
Called every 5 minutes by launchd. Runs fill_asin_blanks.py exactly once per
weekday, only from 9:30 KST onward. If the Mac was off/asleep at 9:30, the
next time this fires after wake (still within the same day) it catches up
and runs then -- satisfying "retry every 5 min till this runs after 9:30".
"""
import json
import os
import subprocess
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

STATE_PATH = os.path.expanduser("~/Desktop/GCX/Scrapers/SKU_ASIN_Filler/last_run.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
KST = ZoneInfo("Asia/Seoul")


def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH) as f:
            return json.load(f)
    return {}


def save_state(state):
    with open(STATE_PATH, "w") as f:
        json.dump(state, f)


def main():
    now = datetime.now(KST)
    today_str = now.strftime("%Y-%m-%d")

    if now.weekday() >= 5:  # 5=Sat, 6=Sun
        return

    if now.hour < 9 or (now.hour == 9 and now.minute < 30):
        return

    state = load_state()
    if state.get("last_run_date") == today_str:
        return  # already ran today

    state["last_run_date"] = today_str
    state["last_run_at"] = now.isoformat()
    save_state(state)

    subprocess.run(
        [sys.executable, os.path.join(SCRIPT_DIR, "fill_asin_blanks.py")],
        check=False,
    )


if __name__ == "__main__":
    main()
