# Caspi 판매량 Backfill

`backfill.py` fills column Z (판매량, EU+UK Amazon FBA) on the `1-5점` tab of the
three review-monitoring spreadsheets (iPhone 18 / Pixel 11 / Galaxy Z Fold8-Flip8-Fold8
Ultra Series), from each product's fixed launch-date baseline through Caspi's latest
available daily snapshot.

Runs unattended via launchd (`~/Library/LaunchAgents/com.spigen.gcx.caspi-sales-backfill.plist`),
fired every 30 min. Each tick is a no-op unless it's a weekday between 08:00-23:59 KST
**and** a given product hasn't already succeeded that day — so a normal day writes once
around 8-9AM, and if the Mac was off/asleep through the usual window, the first tick
after it wakes catches up later the same day rather than skipping it.

## Data source

Caspi registered query `pq_bc4163d82dce2e4d38` ("판매량_EU_backfill_delta_by_baseline_date"),
called headlessly via `POST https://caspilm.spigen.com/api/data-api/run` with an API key
— no live Claude session needed. Underlying table:
`S3.AMAZON_SELLER.RESTOCK_INVENTORY_RECOMMENDATIONS_REPORT` (Amazon Seller Central FBA
restock report, DE/FR/ES/IT/GB only — no US/JP/IN despite those countries appearing
elsewhere in the sheet).

Per SKU: `delta = SUM(Units Sold Last 30 Days @ latest snapshot) − SUM(... @ baseline date)`,
grouped by the 8-char base SKU across all country/color variants. Negative deltas (return/
report-recalc noise) are floored to 0; a SKU with no baseline-date row is treated as
baseline=0 (its full latest value counts); a SKU with zero Caspi rows at all is left blank.

## Config

Per-product spreadsheet ID, sheet ID, and baseline/start date live in the `PRODUCTS` list
at the top of `backfill.py`. Secrets (Caspi API key + registered queryId) and per-day
success state live outside the repo at `~/.config/caspi_sales_backfill/` (`secrets.json`,
`state.json`) — never committed.

Full method + caveats (this is a directional estimate, not Caspi's canonical `실판매`
source): see memory `caspi_판매량_backfill_workflow.md`.

## Manual run

```
python3 backfill.py
```

Safe to re-run any time — it's a no-op for any product that already succeeded today.
