# Caspi 판매량 Backfill

`backfill.py` fills column **Z (판매량, EU+UK Amazon)** on the `1-5점` tab of the three
review-monitoring spreadsheets (iPhone 18 / Pixel 11 / Galaxy Z Fold8-Flip8-Fold8 Ultra
Series) with **cumulative units sold per base SKU** from each product's start date
through Caspi's latest order day.

## Screenshots

![iPhone 18 `1-5점` tab: SKU column Q that the backfill keys on (order IDs blurred)](docs/sku_column.jpg)
*iPhone 18 `1-5점` tab: SKU column Q that the backfill keys on (order IDs blurred)*

## Schedule

Since 2026-10-07 it runs on the 24/7 GCX server crontab (`ServerBootstrap/crontab.txt`),
every 30 min:

```
*/30 * * * *  cd $G/GAS_Operations/CaspiSalesBackfill && python3 backfill.py
```

The old Mac LaunchAgent (`com.spigen.gcx.caspi-sales-backfill.plist`) is parked in
`~/Library/LaunchAgents/disabled-moved-to-server/` — don't run both.

Each tick is a no-op unless it's a **weekday 08:00–23:59 KST** and a product hasn't
already succeeded today (`state.json` → `last_success_date`). So a normal day writes once
around 08:00, and a missed window is caught up later the same day.

## How it works

1. For each pending product, calls the Caspi registered query
   **`판매량_EU_cumulative_since_start_date`** (queryId read from `secrets.json` →
   `caspi_query_id_cumulative`; param `start_date`) via
   `POST https://caspilm.spigen.com/api/data-api/run` (paged by `nextOffset`).
   Source table: `S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL` —
   EU+UK marketplaces (DE/FR/IT/ES/UK/NL/SE/PL/BE/IE), Cancelled excluded, deduped per
   order+SKU, purchase date ≥ `start_date`. Returns `BASE_SKU`, `UNITS`.
2. Reads `1-5점!A2:Z`, matches column **Q** (SKU) to `BASE_SKU`, writes `Z2:Z` (RAW).
   SKUs with no orders are left blank.
3. Sets a note on the Z header cell: `last updated <date> from <table> (accumulated from <start_date>)`.
4. Saves state after each product, so a crash keeps partial progress.

> **2026-10-01 change:** until then Z used `RESTOCK_INVENTORY_RECOMMENDATIONS_REPORT`
> "Units Sold Last 30 Days" (latest minus a baseline snapshot). That was a rolling 30-day
> FBA EU5 number, not cumulative, and the baseline never matched — replaced by the
> cumulative orders query above.

## Config

`PRODUCTS` at the top of `backfill.py`:

| key | Spreadsheet ID | `1-5점` sheetId | start_date |
|-----|----------------|-----------------|------------|
| `iphone18` | `1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU` | `957652957` | 2026-09-18 |
| `pixel11` | `12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI` | `957652957` | 2026-08-18 |
| `glxz8` | `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4` | `957652957` | 2026-07-27 |

Outside the repo (never committed):

- `~/.config/caspi_sales_backfill/secrets.json` — `caspi_api_key`, `caspi_query_id_cumulative`
- `~/.config/caspi_sales_backfill/state.json` — per-product `last_success_date`, `last_run_at`
- `~/.config/gws_shim/token.json` — Sheets API OAuth token (`kjw@spigen.com`, GCP `gcxbot`)

Logs: `logs/launchd.{out,err}.log` (gitignored). This is a directional estimate, not
Caspi's canonical `실판매` source — see memory `caspi_판매량_backfill_workflow.md`.

## Manual run

```
python3 backfill.py
```

Writes to the live sheets (there is no dry-run flag). Safe to re-run — no-op for any
product that already succeeded today; delete that product's entry in `state.json` to
force a rewrite.
