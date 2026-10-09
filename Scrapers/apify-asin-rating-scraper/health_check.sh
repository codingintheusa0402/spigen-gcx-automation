#!/bin/zsh
# Daily health check for the published Apify actor: runs it with default input on the
# owner account and checks Apify's own status for the actor. Alerts via macOS notification.
TOKEN=$(cat ~/.config/apify_asin_scraper/token)
ACTOR=2AuZCyLxKZuO2OyNY
LOG=~/Library/Logs/apify_asin_scraper_health.log
alert() { osascript -e "display notification \"$1\" with title \"Apify ASIN scraper\" sound name \"Basso\""; }

items=$(curl -s -X POST "https://api.apify.com/v2/acts/$ACTOR/run-sync-get-dataset-items?token=$TOKEN&timeout=280" \
  -H 'Content-Type: application/json' -d '{}')
ok=$(echo "$items" | python3 -c "import sys,json
try: d=json.load(sys.stdin)
except Exception: print(0); sys.exit()
print(sum(1 for r in d if isinstance(r,dict) and r.get('found') and r.get('rating') and r.get('ratingsCount')))")
actor_state=$(curl -s "https://api.apify.com/v2/acts/$ACTOR?token=$TOKEN" | python3 -c "import sys,json
d=json.load(sys.stdin)['data']; print('maintenance' if d.get('isUnderMaintenance') or d.get('notice') not in (None,'NONE') else 'ok', d.get('isPublic'))")
echo "$(date '+%F %T') valid_items=$ok actor=$actor_state" >> $LOG
if [[ "$ok" -lt 2 ]]; then alert "Test run FAILED ($ok/2 valid items) — check $LOG"; fi
if [[ "$actor_state" == maintenance* ]]; then alert "Apify flagged the actor (under maintenance) — fix before it is hidden"; fi
