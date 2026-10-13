#!/bin/bash
# One-off test cadence: full --commit runs that execute the saved Apify tasks (Spigen org) and write
# their results, at BASE+0h/+1h/+2h for each run number given, then exits.
# Not the permanent schedule (Apify schedule 08:00 + rating_fetch_gate.py 08:20 on gcx-server).
#   usage: caffeinate -i bash test_runs_3x_hourly.sh <BASE_EPOCH> [run numbers, default "1 2 3"]
cd "$(dirname "$0")"
BASE=${1:-$(date +%s)}
RUNS=${2:-"1 2 3"}
for n in $RUNS; do
  target=$((BASE + (n - 1) * 3600))
  now=$(date +%s)
  [ "$target" -gt "$now" ] && sleep $((target - now))
  python3 -I rating_fetch_daily.py --commit --source tasks-run --run-label "test $n/3 (per-sheet Apify tasks, Spigen org)" >> ~/.config/asin_rating_daily/run.log 2>&1
done
