# 유지훈P_Apify

Container-bound Google Apps Script for the 유지훈P review spreadsheet. A daily 04:00 KST trigger starts two saved Apify tasks (Amazon review scrape + product rating/review-count scrape), polls them every minute, writes the results into the sheet, stamps the run time on the `US` tab and posts completion messages to Google Chat. It also sends a Chat alert whenever column K of the `US` tab is edited, and provides the Gemini-based `=DR()` defect-classification custom function.

**Script ID:** `1UZ5NzqtjTa5nHW17w0vOOgbrFG6Zn2TpxVqkfFc4WdZHfR7upLVz_lmO`
**Linked spreadsheet:** `1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw` (`SPREADSHEET_ID` in `Alert.js`)

## Screenshots

![`Product` tab filled by the Apify product task](docs/product_tab.jpg)
*`Product` tab filled by the Apify product task*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Review-scrape lifecycle — start task, poll, write new dated sheet, Chat notification, Excel-export helpers (same as the SDA/Auto_Acc pattern + US-tab timestamp) |
| `Product.js` | Product scrape (rating / review count per ASIN) → `Product` tab; recurring poller |
| `Trigger.js` | `getSpreadsheetId_()`, `_writeTimestampToUsSheet_()`, `dailyScrapeJob()`, `createDailyTriggers()` |
| `Alert.js` | `SPREADSHEET_ID`, `CHAT_WEBHOOK_URL`, installable onEdit alert for `US!K`, `testSendChat()`, `listTriggers()` |
| `Gemini.js` | `=DR(inputText, category)` custom function (Gemini defect classifier, reads the `Defect` tab) |
| `Config.js` | `CONFIG` object only (task ID, timezone, poll interval/timeout) |
| `UI.js` | `onOpen()` menu + `menuRunProduct()` handler |
| `appsscript.json` | GAS manifest (timezone `Asia/Seoul`, V8) |

---

## How it works

### Daily scrape (`Trigger.js`)
- `createDailyTriggers()` (run once in the editor) installs a daily **04:00 Asia/Seoul** trigger for `dailyScrapeJob()`.
- `dailyScrapeJob()` → `startApifyRunAndSchedulePoll()` + `_ensureRecurringPoller_()` (review task) and `runProductNowAndPollRecurring()` (product task). Each poller runs every 1 min, gives up after 180 min.
- **Review result:** only `statusCode=200` / `FOUND` items, flattened into a **new tab `Apify_yyMMdd`** (`_2`… if taken), deduped by `username + reviewTitle + reviewDescription`; Chat gets the sheet + `.xlsx` export links.
- **Product result:** the **`Product`** tab is cleared and rewritten (`country, asin, title, countReview, productRating, url`); Chat gets a completion message.
- After each successful write, `_writeTimestampToUsSheet_()` writes the current KST time into row 2 of the `US` tab's `Timestamp` column (falls back to column J).
- Manual: **Apify → Product → Run Product (auto polling)** / **Cancel Product Polling**.

### Column-K edit alert (`Alert.js`)
`createEditTrigger()` (run once) installs an installable onEdit trigger → `onEditAlertToGoogleChat(e)`: a single-cell edit in `US!K` posts old/new value, cell and sheet link to Chat.

### `=DR(inputText, category)` (`Gemini.js`)
Classifies a review into one of the defect labels listed for that category in the **`Defect`** tab (A=category, B=label, C=description; `모니터링대상아님` excluded; `태블릿케이스` mapped to `휴대폰케이스`). Order: script cache (key prefix `DR_v19_`, 6 h) → `keywordFallback_()` → Gemini models `gemini-3.1-flash-lite`, `gemini-2.5-flash-lite`, `gemini-3.5-flash`. `testDR()` is an editor diagnostic.

---

## Config

| Key | Value |
|-----|-------|
| `CONFIG.actorTaskIdOrSlug` | `vUlCaUpNvjgC23g5T` (review task) |
| `PRODUCT.taskIdOrSlug` (`Product.js`) | `09gwMIgOKhyyDrH3g` (product task) |
| `CONFIG.pollIntervalMinutes` / `pollMaxMinutes` | 1 / 180 |
| `TARGET_SHEET_NAME` / `TARGET_COLUMN` (`Alert.js`) | `US` / 11 (K) |
| `CHAT_WEBHOOK_URL` (`Alert.js`) | TCK GCX Spigen Google Chat space (used for both edit alerts and scrape notifications) |

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (required) |
| `GEMINI_API_KEY` | Gemini API key for `=DR()` |
| `APIFY_LAST_*`, `PRODUCT_LAST_*` (`RUN_ID`, `DATASET_ID`, `POLL_STARTED_AT_MS`) | Run state written/cleared by code |

---

## Known issues (verified against local code — not checked against the live GAS project)

- **`PREFERRED_HEADERS` is not defined anywhere in this project** (the other per-product folders define it in `Config.js`). `_collectHeadersFromFlat()` / `_overwriteSheet()` reference it, so the review write step would throw `ReferenceError` as soon as a run succeeds. The product write path doesn't use it.
- **Review task ID looks mistyped:** `vUlCaUpNvjgC23g5T` is the shared review task `TvUlCaUpNvjgC23g5` with the leading `T` moved to the end. If so, the review start call fails (the error is logged and `dailyScrapeJob` continues with the product scrape).
- `getPollDelayMs_()` and `FILTER_WHITE_ROWS()` from the sibling projects are absent here (`CONFIG.POLL_DELAY` exists but is unused).

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/유지훈P_Apify
clasp push --force
```
