# SDA_Apify (Screen & Display Accessories)

Container-bound Google Apps Script for the SDA review spreadsheet. It starts saved Apify tasks (Amazon review scrape + product rating/review-count scrape), polls them with a recurring 1-minute trigger, writes the results into the sheet and posts a completion message to Google Chat. Same Config/Trigger/Product/UI split-file pattern as [Auto_Acc_Apify](../Auto_Acc_Apify/) and [iPh17e_Apify](../iPh17e_Apify/).

**Script ID:** `1rUwC_XwGUjZvu5ileGo_jAhihoMaio4wfdtjSnqO0VIclG0OByVDyoco`
**Linked spreadsheet:** `1sxapIqJgXcJdeqyCf9bAxCNXrVMsVjsZE9QWPwEm0R4` (SDA review sheet — the code itself uses `SpreadsheetApp.getActive()`, no ID constant)

> ⚠️ `Auto_Acc_Apify`'s local `.clasp.json` points at this **same** script ID, so `clasp push` from that folder overwrites this project. All six code files are currently byte-identical between the two folders (so it's harmless today), but they must not diverge until Auto_Acc's real script ID is re-identified. This README's script ID is the authoritative one.

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
| `appsscript.json` | GAS manifest (timezone `Asia/Seoul`, V8) |

---

## How it works

### Product scrape (menu-driven)
Open the spreadsheet → **Apify → Product → Run Product (auto polling)**.

1. Starts Apify task `PRODUCT.taskIdOrSlug` (`Rs0CN69AhiwPYkt3H`) with its saved input and creates a recurring `pollProductRunAndWrite` trigger (every 1 min).
2. On `SUCCEEDED`: fetches `asin,countReview,productRating,url,title,globalReviews`, clears and rewrites the **`Product`** tab (`country, asin, title, countReview, productRating, url`), posts "Apify Product Scraping Completed." to Chat, deletes the poller.
3. Failure statuses or 180-min timeout → state cleared, poller deleted, toast only. **Cancel Product Polling** removes the poller.

### Review scrape
`startApifyRunAndSchedulePoll()` starts task `TvUlCaUpNvjgC23g5`; `pollApifyRunAndWrite()` keeps only `statusCode=200` / `FOUND` items, writes them flattened into a **new tab `Apify_yyMMdd`** (suffix `_2`… if taken), dedupes by `username + reviewTitle + reviewDescription`, and posts the sheet + `.xlsx` export links to Chat. No menu item; see Known issues.

### Misc utility functions

| Function | Purpose |
|----------|---------|
| `FILTER_WHITE_ROWS()` | Custom function: returns rows from the `신제품 라인업` sheet (cols A:I) whose column-A cell has a white (`#ffffff`) background |

---

## Config (`Config.js`)

| Key | Value |
|-----|-------|
| `actorTaskIdOrSlug` | `TvUlCaUpNvjgC23g5` (review task — same as Auto_Acc / iPh17e) |
| `PRODUCT.taskIdOrSlug` | `Rs0CN69AhiwPYkt3H` (product task) |
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

- `createApifyWeekdayTriggers()` schedules one-off 04:00 Mon–Fri triggers (tomorrow → `2026-12-31`) for `runApifyNowAndPollAfter2Hours`, which is **not defined** in the project — those triggers would fail. The review scrape has no working automatic schedule.
- `startApifyRunAndSchedulePoll()` does not create the review poller (`_ensureRecurringPoller_()` is never called), so `pollApifyRunAndWrite` must be run manually.
- scriptId shared with Auto_Acc_Apify (above).

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/SDA_Apify
clasp push --force
```
