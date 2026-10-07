# SKUSales_Rating_Apify

> Sibling project: [iPhone18Fold_Rating_Apify](../iPhone18Fold_Rating_Apify/) — same design, different spreadsheet. See [AmazonDE_RatingScraper_README.md](../AmazonDE_RatingScraper_README.md) for the shared architecture.

Container-bound Google Apps Script for the "해외사업부문 개발 발의 품목_사후관리 (세일즈, 리뷰, 필드테스트)" spreadsheet. Runs the Apify task `product-details-scraper-ljh` (axesso_data/amazon-product-details-scraper actor, pre-loaded with each row's Amazon.de URL) every Monday 8AM KST and writes each product's current rating back into the sheet, matched by ASIN.

**Linked spreadsheet:** `1_9O8oTHt-yHewG0aGeioeM0psZ2kprf3lgH77-BFLs8`
**Target sheet:** `SKU세일즈/리뷰` — ASIN in col **G**, rating written to col **I** starting at row **7**
**Apify task:** `QWtvKi7oXZ6YYR92G` (`product-details-scraper-ljh`)

## Screenshots

![`SKU세일즈/리뷰` tab: rating column I refreshed per ASIN (sales columns cropped)](docs/rating_sheet.jpg)
*`SKU세일즈/리뷰` tab: rating column I refreshed per ASIN (sales columns cropped)*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Apify run lifecycle (start task → poll → write ratings), weekly trigger setup |
| `config.js` | Sheet/column config, task ID, `_getToken()` |
| `appsscript.json` | GAS manifest |

---

## Flow

1. `runApifyRatingRefreshNow()` starts an async run of Apify task `QWtvKi7oXZ6YYR92G` and schedules a recurring 1-minute poller (`pollRatingRunAndWrite`).
2. The poller checks run status; once `SUCCEEDED`, it fetches the run's dataset (`asin`, `productRating` fields from the axesso actor), builds an ASIN → rating map, and writes `SKU세일즈/리뷰!I7:I<lastRow>` row-by-row matched against `SKU세일즈/리뷰!G7:G<lastRow>` (ASIN) as `=HYPERLINK("https://www.amazon.de/dp/<ASIN>", <rating>)`. A row whose ASIN got no (or empty) result in that run is left completely untouched — an existing rating is only ever overwritten by a fresh non-empty scrape result, never blanked.
3. `productRating` comes back locale-formatted (e.g. `"4,5 von 5 Sternen"` for amazon.de) — `_extractRatingValue_()` takes the leading number and normalizes the decimal comma to a dot (e.g. `4.5`).
4. A time-based trigger calls `runApifyRatingRefreshNow` every Monday ~08:00 Asia/Seoul (see `setupWeeklyTrigger()`).
5. `reprocessDataset(datasetId)` re-applies an already-completed run's dataset to the sheet without starting a new Apify run — useful after a write-logic fix.
6. `syncNewAsinsToApifyTask()` scans the sheet's ASIN column for any ASIN not yet covered by the task's `input.urls`, appends `https://www.amazon.de/dp/<ASIN>` for each one, and PUTs the task via the Apify API. A weekly trigger (`setupAsinSyncTrigger()`, every Sunday ~07:00 Asia/Seoul — before Monday's kickoff) runs this automatically, so rows added to the sheet during the week get picked up without any manual step. `_looksLikeAsin_()` requires exactly 10 uppercase-alphanumeric characters, so placeholder text some rows carry instead of a real ASIN (e.g. `"TBU"`, `"미판매"`) is never sent to Apify.

---

## Script Properties required

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (Project Settings → Script Properties in the Apps Script editor) |
| `RATING_LAST_RUN_ID`, `RATING_LAST_DATASET_ID`, `RATING_LAST_POLL_STARTED_AT_MS` | Run state written/cleared by the code (a pending run ID makes "Run Now" just re-arm the poller instead of starting a second run) |

No Google Chat notification — completion is a sheet toast + execution log only. Polling gives up after `CONFIG.pollMaxMinutes` (180).

---

## One-time setup (after `clasp push`)

In the Apps Script editor (or via the spreadsheet's **Apify Rating** menu once the sheet is reloaded):

1. Set the `APIFY_TOKEN` script property.
2. Run **Apify Rating → Install Weekly Monday 8AM Trigger** (or call `setupWeeklyTrigger()` once) to install the recurring kickoff trigger. This requires an OAuth authorization prompt the first time.
3. Run **Apify Rating → Install Weekly ASIN-Sync Trigger** (or call `setupAsinSyncTrigger()` once) to install the Sunday ASIN-sync trigger.
4. Optionally run **Apify Rating → Run Now (refresh ratings)** to do an immediate refresh.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/SKUSales_Rating_Apify
clasp push --force
```
