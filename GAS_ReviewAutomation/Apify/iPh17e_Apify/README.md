# iPh17e_Apify (iPhone 17e)

Container-bound Google Apps Script for the iPhone 17e review spreadsheet. It starts saved Apify tasks (Amazon review scrape + product rating/review-count scrape), polls them with a recurring 1-minute trigger, writes the results into the sheet and posts a completion message to Google Chat. Same code as [SDA_Apify](../SDA_Apify/) / [Auto_Acc_Apify](../Auto_Acc_Apify/) except for the product task ID and the manifest timezone.

**Script ID:** `1oIR6d9_cjLXpRLMvpVfU0WSIPOZba3S1m7DGc1eAx8VmPMzxqjm7eU-6`
**Linked spreadsheet:** `16xRJHH7Ynii4erNOn_905ST4CZs6OLpOYTof4uqsGsQ` (iPhone 17e review sheet — code uses `SpreadsheetApp.getActive()`)

## Screenshots

![`Product` tab filled by the Apify product task](docs/product_tab.jpg)
*`Product` tab filled by the Apify product task*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Review-scrape run lifecycle — start task, poll status, write a new dated sheet, Chat notification, Excel-export helpers, `FILTER_WHITE_ROWS()` |
| `Config.js` | `CONFIG` object, `PREFERRED_HEADERS` column order, `CHAT_WEBHOOK_URL`, `getPollDelayMs_()` |
| `Product.js` | Product-level scrape (rating / review count per ASIN) → `Product` sheet; recurring poller |
| `Trigger.js` | `createApifyWeekdayTriggers()` / `deleteApifyWeekdayTriggers()` (see Known issues) |
| `UI.js` | `onOpen()` menu + `menuRunProduct()` handler |
| `appsscript.json` | GAS manifest — timezone **`America/New_York`** (unlike the SDA/Auto_Acc twins) |

---

## How it works

### Product scrape (menu-driven)
Open the spreadsheet → **Apify → Product → Run Product (auto polling)**.

1. Starts Apify task `PRODUCT.taskIdOrSlug` (`MlqquJP8seUKlFezP`) with its saved input and creates a recurring `pollProductRunAndWrite` trigger (every 1 min).
2. On `SUCCEEDED`: fetches `asin,countReview,productRating,url,title,globalReviews`, clears and rewrites the **`Product`** tab (`country, asin, title, countReview, productRating, url`), posts "Apify Product Scraping Completed." to Chat, deletes the poller.
3. Failure statuses or 180-min timeout → state cleared, poller deleted. **Cancel Product Polling** removes the poller.

### Review scrape
`startApifyRunAndSchedulePoll()` starts task `TvUlCaUpNvjgC23g5`; `pollApifyRunAndWrite()` keeps only `statusCode=200` / `FOUND` items, writes them into a **new tab `Apify_yyMMdd`** (date in `Asia/Seoul`; suffix `_2`… if taken), dedupes by `username + reviewTitle + reviewDescription`, and posts the sheet + `.xlsx` export links to Chat. No menu item; see Known issues.

### Misc utility functions

| Function | Purpose |
|----------|---------|
| `FILTER_WHITE_ROWS()` | Custom function: returns rows from the `신제품 라인업` sheet (cols A:I) whose column-A cell has a white (`#ffffff`) background |

---

## Config (`Config.js`)

| Key | Value |
|-----|-------|
| `actorTaskIdOrSlug` | `TvUlCaUpNvjgC23g5` (review task — same as SDA / Auto_Acc) |
| `PRODUCT.taskIdOrSlug` | `MlqquJP8seUKlFezP` (iPhone 17e product task) |
| `pollIntervalMinutes` / `pollMaxMinutes` | 1 / 180 |
| `POLL_DELAY` | 2 hours (production) / 2 minutes (test mode) |
| `CHAT_WEBHOOK_URL` | TCK GCX Spigen Google Chat space |

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (required) |
| `APIFY_LAST_*`, `PRODUCT_LAST_*` (`RUN_ID`, `DATASET_ID`, `POLL_STARTED_AT_MS`) | Run state written/cleared by code |

---

## Known issues (verified against current code)

- `createApifyWeekdayTriggers()` schedules one-off Mon–Fri triggers for `runApifyNowAndPollAfter2Hours`, which is **not defined** in the project — those triggers would fail. Also, because the manifest timezone is `America/New_York`, its `setHours(4)` would mean 04:00 ET, not 04:00 KST as the comment says.
- `startApifyRunAndSchedulePoll()` does not create the review poller (`_ensureRecurringPoller_()` is never called), so `pollApifyRunAndWrite` must be run manually.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/iPh17e_Apify
clasp push --force
```
