#!/bin/bash
# One-off helper for 2026-10-13: runs #2, #3, #4 of today's 4-run test cadence,
# exactly 1 hour apart, then exits. Not part of the permanent schedule — the
# permanent job lives on gcx-server (rating_fetch_gate.py via crontab).
#
# Wrapped in `caffeinate -i` (invoke this script itself via
# `caffeinate -i bash today_hourly_test_runs.sh`) so the Mac can't idle-sleep
# mid-test — run #1 hit a ~4h sleep gap before this was added.
cd "$(dirname "$0")"
for n in 2 3 4; do
  sleep 3600
  python3 -I rating_fetch_daily.py --commit --run-label "run $n/4 (today, hourly test)" >> ~/.config/asin_rating_daily/run.log 2>&1
done
