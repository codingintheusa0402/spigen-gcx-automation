#!/bin/bash
# Hourly wrapper for scrape_sc_reviews.py, invoked by the
# com.spigen.sc_scraper.hourly LaunchAgent (StartInterval=3600, i.e. every
# exactly 1 hour). Skips this cycle if a previous run is still in progress
# instead of launching a second Chrome instance on the same scraper profile
# (which would conflict). All scrape output still goes to /tmp/sc_scraper.log;
# this script's own start/skip/finish lines go to /tmp/sc_scraper_hourly.log.

SCRIPT_DIR="/Users/kevinkim/Desktop/GCX/Scrapers/SC_Review_Scraper"
LOG="/tmp/sc_scraper.log"
CRON_LOG="/tmp/sc_scraper_hourly.log"

echo "$(date '+%Y-%m-%d %H:%M:%S') — hourly trigger fired" >> "$CRON_LOG"

if pgrep -f "scrape_sc_reviews.py" > /dev/null; then
  echo "$(date '+%Y-%m-%d %H:%M:%S') — previous run still active, skipping this cycle" >> "$CRON_LOG"
  exit 0
fi

cd "$SCRIPT_DIR" || exit 1
unset SC_SCRAPER_CREDENTIALS_FILE SC_SCRAPER_HEADLESS SC_SCRAPER_OUT_DIR

/opt/homebrew/bin/python3 scrape_sc_reviews.py > "$LOG" 2>&1
echo "$(date '+%Y-%m-%d %H:%M:%S') — run finished (exit $?)" >> "$CRON_LOG"
