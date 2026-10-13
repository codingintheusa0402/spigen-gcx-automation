#!/bin/bash
# One-off helper for 2026-10-13: runs #2, #3, #4 of today's 4-run test cadence at
# fixed start times 1h, 2h, 3h after BASE_EPOCH (= run #1's finish), then exits.
# Not part of the permanent schedule — that lives on gcx-server (rating_fetch_gate.py
# via crontab). Launch under `caffeinate -i` so the Mac can't idle-sleep mid-test
# (run #1 lost ~4h to a sleep gap before this was added).
#   usage: caffeinate -i bash today_hourly_test_runs.sh <BASE_EPOCH>
cd "$(dirname "$0")"
BASE=${1:?base epoch required}
for n in 2 3 4; do
  target=$((BASE + (n - 1) * 3600))
  now=$(date +%s)
  [ "$target" -gt "$now" ] && sleep $((target - now))
  python3 -I rating_fetch_daily.py --commit --run-label "run $n/4 (today, hourly test)" >> ~/.config/asin_rating_daily/run.log 2>&1
done
