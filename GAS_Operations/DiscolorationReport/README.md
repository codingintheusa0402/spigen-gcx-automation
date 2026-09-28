# DiscolorationReport — 이염/변색 클레임·배드리뷰 정기 리포트

Twice-weekly (Mon/Fri 10:00 KST) Google Chat report of new 이염/변색 Zendesk claims and
Amazon bad reviews, grouped by base SKU with each product's 90-day cumulative count
(🚨 flag at ≥3 = SIREN review candidate).

- **Data**: Caspi registered queries via the headless `data-api/run` endpoint
  - `pq_cde9dcd490cc850dad` — Zendesk tickets tagged `_case__이염/변색`, device/product names decoded
  - `pq_3cc3dd0f30778bf1f8` — Spigen ★1–3 reviews passing a multilingual keyword prefilter
- **Review classification**: local `claude -p --model sonnet` (keyword prefilter alone is ~50% noise;
  황변/yellowing is excluded). Each review is classified once and cached.
- **Schedule**: `~/Library/LaunchAgents/com.spigen.gcx.discoloration-report.plist`, every 30 min +
  RunAtLoad; `report.py` self-gates to Mon/Fri ≥10:00 and once per day, so a sleeping/off Mac
  catches up on the next tick.
- **Config/state** (outside the repo): `~/.config/discoloration_report/secrets.json`
  (`caspi_api_key`, `webhook_url`, chmod 600) and `state.json` (reported IDs, classification cache).
- "New" = not yet reported and created within the last 14 days (catches Caspi's ~1-day Zendesk load lag).

```
python3 report.py --dry-run   # print message, send nothing
python3 report.py --force     # send now regardless of day/time gate
```
Logs: `logs/launchd.{out,err}.log`.
