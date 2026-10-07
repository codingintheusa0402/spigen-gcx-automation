# iPhone18Fold_Rating_Apify

> Sibling project: [SKUSales_Rating_Apify](../SKUSales_Rating_Apify/) — same design, different spreadsheet. See [AmazonDE_RatingScraper_README.md](../AmazonDE_RatingScraper_README.md) for the shared architecture.

Container-bound Google Apps Script for the "주요 디바이스별 아마존 세일즈 현황" spreadsheet. Runs the Apify task `product-details-scraper-iphone18-fold` (axesso_data/amazon-product-details-scraper actor — same actor as [SKUSales_Rating_Apify](../SKUSales_Rating_Apify/)) every weekday at 8AM KST and writes each product's current Amazon.de rating back into the sheets, matched by ASIN.

**Linked spreadsheet:** `1XXIMDNVMPBut8WTCV-4gZt4DSxHmGaRR8GkRejAOGq0`
**Target sheets:** `iPhone 18`, `iPhone Fold`, `Apple ETC(26)`, `Pixel 11` — all four share the same layout: ASIN in col **B**, rating written to col **E** ("Rating") starting at row **6**
**Apify task:** `CLQ3G6Sokyr7AJtQi` (`product-details-scraper-iphone18-fold`) — one task, pre-loaded with amazon.de URLs for every ASIN across all four sheets (266 URLs as of 2026-09-29, after Apple ETC(26)'s ASINs were filled in; kept in sync by `syncNewAsinsToApifyTask()`)

## Screenshots

![`iPhone 18` tab: Rating column E refreshed per ASIN (sales columns cropped)](docs/rating_sheet.jpg)
*`iPhone 18` tab: Rating column E refreshed per ASIN (sales columns cropped)*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Apify run lifecycle (start task → poll → write ratings to all four sheets), daily weekday trigger setup |
| `config.js` | `SHEETS` array (per-sheet ASIN/rating column config), task ID, `_getToken()` |
| `appsscript.json` | GAS manifest |

---

## Flow

1. `runApifyRatingRefreshNow()` starts an async run of Apify task `CLQ3G6Sokyr7AJtQi` and schedules a recurring 1-minute poller (`pollRatingRunAndWrite`).
2. The poller checks run status; once `SUCCEEDED`, it fetches the run's dataset (`asin`, `productRating`), builds one ASIN → rating map, and applies it to **all four** sheets' `E6:E<lastRow>`, each matched against its own `B` column (ASIN), writing `=HYPERLINK("https://www.amazon.de/dp/<ASIN>", <rating>)`. A row whose ASIN got no (or empty) result in that run is left completely untouched — an existing rating is only ever overwritten by a fresh non-empty scrape result, never blanked.
3. `productRating` comes back locale-formatted (e.g. `"4,5 von 5 Sternen"`) — `_extractRatingValue_()` takes the leading number and normalizes the decimal comma to a dot (e.g. `4.5`).
4. `dailyWeekdayKickoff()` is called by a daily trigger at ~08:00 Asia/Seoul; it checks the ISO weekday and skips (no run started) on Saturday/Sunday, otherwise calls `runApifyRatingRefreshNow()`.
5. `reprocessDataset(datasetId)` re-applies an already-completed run's dataset to all four sheets without starting a new Apify run — useful after a write-logic fix.
6. `syncNewAsinsToApifyTask()` scans all four sheets' ASIN columns for any ASIN not yet covered by the task's `input.urls`, appends `https://www.amazon.de/dp/<ASIN>` for each one, and PUTs the task via the Apify API. It also runs at the start of every rating run (`_startRatingRun_()`), so ASINs added to any of the four sheets are scraped in the very next daily run; a sync failure there is logged and the run proceeds with the existing URL list. The weekly trigger (`setupAsinSyncTrigger()`, every Sunday ~07:00 Asia/Seoul) remains as a backstop. `_looksLikeAsin_()` requires exactly 10 uppercase-alphanumeric characters, so placeholder text some rows carry instead of a real ASIN is never sent to Apify.

---

## Script Properties required

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (Project Settings → Script Properties in the Apps Script editor) |
| `RATING_LAST_RUN_ID`, `RATING_LAST_DATASET_ID`, `RATING_LAST_POLL_STARTED_AT_MS` | Run state written/cleared by the code |

No Google Chat notification — completion is a sheet toast + execution log only. Polling gives up after `CONFIG.pollMaxMinutes` (180). A missing tab in `SHEETS` makes the write step throw (`Sheet "<name>" not found`), while the ASIN sync just skips it.

---

## One-time setup (after `clasp push`)

1. Set the `APIFY_TOKEN` script property.
2. Run **Apify Rating → Install Daily Weekday 8AM Trigger** (or call `setupDailyWeekdayTrigger()` once) — requires an OAuth authorization prompt the first time.
3. Run **Apify Rating → Install Weekly ASIN-Sync Trigger** (or call `setupAsinSyncTrigger()` once) to install the Sunday ASIN-sync trigger.
4. Optionally run **Apify Rating → Run Now (refresh ratings)** for an immediate refresh.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/iPhone18Fold_Rating_Apify
clasp push --force
```
