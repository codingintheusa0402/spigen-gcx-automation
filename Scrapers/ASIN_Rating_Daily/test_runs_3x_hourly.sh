#!/bin/bash
# One-off test cadence: 3 full --commit runs, the first immediately, then +1h and +2h, then exits.
# Not the permanent schedule (that's rating_fetch_gate.py via crontab on gcx-server).
#   usage: caffeinate -i bash test_runs_3x_hourly.sh
cd "$(dirname "$0")"
BASE=$(date +%s)
for n in 1 2 3; do
  target=$((BASE + (n - 1) * 3600))
  now=$(date +%s)
  [ "$target" -gt "$now" ] && sleep $((target - now))
  python3 -I rating_fetch_daily.py --commit --run-label "test $n/3 (hourly, Spigen org)" >> ~/.config/asin_rating_daily/run.log 2>&1
done
