#!/bin/bash
# One-off: user is flying 10:30-11:30 KST today, no network on Mac or server during that
# window. BadReview data finishes crunching ~10:20, so send it right before the window
# closes instead of waiting for the normal 10:30 trigger that might miss connectivity.
# Ledger-gated (auto_broadcast.py's gcx_ledger.claim), so the regular 10:30 fallback tick
# or server cron will just skip once this has already sent — no double broadcast.
TARGET_EPOCH=$(date -j -f "%Y-%m-%d %H:%M:%S" "$(date +%Y-%m-%d) 10:22:00" +%s)
NOW_EPOCH=$(date +%s)
WAIT=$((TARGET_EPOCH - NOW_EPOCH))
if [ "$WAIT" -gt 0 ]; then
  sleep "$WAIT"
fi
cd /Users/kevinkim/Desktop/GCX/GAS_Operations/BadReview_ChatReport
/opt/homebrew/bin/python3 auto_broadcast.py
