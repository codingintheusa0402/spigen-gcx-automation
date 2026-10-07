# Auto_Acc_Apify (Auto Accessories — Apify trigger)

Container-bound Google Apps Script for the Auto Accessories review spreadsheet. It starts saved Apify tasks (Amazon review scrape + product rating/review-count scrape), polls them with a recurring 1-minute trigger, writes the results into the bound spreadsheet and posts a completion message to Google Chat. Config/Trigger/Product/UI split-file architecture — the standard pattern for this repo's simpler per-product Apify triggers (contrast with [GlxZ8_Apify](../GlxZ8_Apify/) / [iPh17e_Monday](../iPh17e_Monday/), which bundle Monday sync + Gemini helpers into the same project).

> ⚠️ **scriptId conflict (unresolved):** this folder's local `.clasp.json` (gitignored) holds the **same** script ID as [`SDA_Apify`](../SDA_Apify/)'s `.clasp.json` (`1rUwC_…yoco`). `clasp push` from either folder overwrites the same live GAS project. The code in both folders is currently byte-identical, so a push is harmless today, but Auto_Acc's real script ID has not been re-identified — treat SDA as the owner of that ID and don't diverge the two folders until this is fixed.

## Screenshots

![Same GAS project as SDA_Apify: its `Product` tab](../SDA_Apify/docs/product_tab.jpg)
*Same GAS project as SDA_Apify: its `Product` tab*

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

### Product scrape (menu-driven — the path that works today)
1. **Apify → Product → Run Product (auto polling)** → `menuRunProduct()` → `runProductNowAndPollRecurring()`.
2. Starts Apify task `PRODUCT.taskIdOrSlug` (`Rs0CN69AhiwPYkt3H`) with its saved input, stores run state in Script Properties, creates a recurring `pollProductRunAndWrite` trigger (every `CONFIG.pollIntervalMinutes` = 1 min).
3. On `SUCCEEDED`: fetches only `asin,countReview,productRating,url,title,globalReviews`, clears and rewrites the **`Product`** tab (`country, asin, title, countReview, productRating, url`; country parsed from the Amazon domain), posts "Apify Product Scraping Completed." to Chat, deletes the poller.
4. On `FAILED`/`ABORTED`/`TIMED-OUT` or after `pollMaxMinutes` (180): state cleared, poller deleted, toast only.
5. **Cancel Product Polling** deletes the poller trigger.

### Review scrape
- `startApifyRunAndSchedulePoll()` starts task `CONFIG.actorTaskIdOrSlug` (`TvUlCaUpNvjgC23g5`); `pollApifyRunAndWrite()` waits for `SUCCEEDED`, keeps only items with `statusCode=200` / `statusMessage="FOUND"`, writes them flattened into a **new tab `Apify_yyMMdd`** (`_2`, `_3`… if it exists), dedupes by `username + reviewTitle + reviewDescription`, and posts the sheet link + `.xlsx` export link to Chat.
- There is no menu item for this path, and `startApifyRunAndSchedulePoll()` does not create the poller itself (`_ensureRecurringPoller_()` is never called) — see Known issues.

### Misc utility functions

| Function | Purpose |
|----------|---------|
| `FILTER_WHITE_ROWS()` | Custom function: returns rows from the `신제품 라인업` sheet (cols A:I) whose column-A cell has a white (`#ffffff`) background |

---

## Config (`Config.js`)

| Key | Value |
|-----|-------|
| `CONFIG.sheetBaseName` | `Apify` (review tabs → `Apify_yyMMdd`) |
| `CONFIG.actorTaskIdOrSlug` | `TvUlCaUpNvjgC23g5` (review task — shared with SDA_Apify / iPh17e_Apify) |
| `PRODUCT.taskIdOrSlug` (`Product.js`) | `Rs0CN69AhiwPYkt3H` (product task — shared with SDA_Apify) |
| `CONFIG.pollIntervalMinutes` / `pollMaxMinutes` | 1 / 180 |
| `CONFIG.POLL_DELAY` | 2 h (production) / 2 min (`TEST_MODE`) — only used by `getPollDelayMs_()` |
| `CHAT_WEBHOOK_URL` | TCK GCX Spigen Google Chat space (hard-coded in `Config.js`; a commented-out private-test-room URL sits above it) |

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (required) |
| `APIFY_LAST_RUN_ID`, `APIFY_LAST_DATASET_ID`, `APIFY_LAST_POLL_STARTED_AT_MS` | Review-run state (written/cleared by code) |
| `PRODUCT_LAST_RUN_ID`, `PRODUCT_LAST_DATASET_ID`, `PRODUCT_LAST_POLL_STARTED_AT_MS` | Product-run state (written/cleared by code) |

---

## Known issues (verified against current code)

- **scriptId shared with SDA_Apify** — see the warning at the top.
- **Weekday trigger points at a function that doesn't exist.** `createApifyWeekdayTriggers()` creates one-off triggers at 04:00 (script timezone) Mon–Fri from tomorrow to `2026-12-31` for handler `runApifyNowAndPollAfter2Hours`, but no such function is defined anywhere in the project, so those triggers would fail with "Script function not found". The review scrape therefore has no working automatic schedule in this code.
- **Review poller is never scheduled.** Running `startApifyRunAndSchedulePoll()` manually only starts the Apify run; `pollApifyRunAndWrite` must be run by hand (or a trigger added) to write the result.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/Auto_Acc_Apify
clasp push --force   # ⚠️ pushes to the SDA_Apify script ID — see warning above
```
