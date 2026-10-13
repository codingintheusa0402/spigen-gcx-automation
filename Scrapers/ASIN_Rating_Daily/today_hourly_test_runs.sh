#!/bin/bash
# One-off helper for 2026-10-13: runs #2, #3, #4 of today's 4-run test cadence,
# exactly 1 hour apart, then exits. Not part of the permanent schedule — the
# permanent job is the weekday crontab entry (see crontab -l / README.md).
cd "$(dirname "$0")"
for n in 2 3 4; do
  sleep 3600
  python3 -I rating_fetch_daily.py --commit --run-label "run $n/4 (today, hourly test)" >> ~/.config/asin_rating_daily/run.log 2>&1
done
