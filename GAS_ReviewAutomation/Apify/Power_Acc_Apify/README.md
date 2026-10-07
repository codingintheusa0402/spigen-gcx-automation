# Power_Acc_Apify (Power Accessories)

Container-bound Google Apps Script for the Power Accessories review spreadsheet (Power Acc. CustomerReviews ★1~3). It runs the Apify review-scrape task into a dated `Apify_yyMMdd` tab and the Apify **Product** task into a `Product` tab, posting a completion message to Google Chat, and carries the shared `=DR(text, category)` Gemini classifier and `=countAnyColoredCells()` helper. The Sheet → Monday uploader files were copied in from the Galaxy S26 project but are **not configured** for this sheet (see Known issues).

**Script ID:** `1xQhzIcvjP2n2zp5RFDOwckB5HOnBCvJXLlw1aRjaXF8ie2JLKX8MHVWt`
**Linked spreadsheet:** `1QC8Is6UvTnFXaOeXviKM_331i3Fo_CBIYx80VS696LI` (hard-coded in `getSpreadsheetId_()`; also `parentId` in `.clasp.json`)

## Screenshots

![Power Acc. `1-3점` tab (reviewer names blurred)](docs/sheet_1-3.jpg)
*Power Acc. `1-3점` tab (reviewer names blurred)*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Review-scrape flow: `startApifyRunAndSchedulePoll()` → recurring `pollApifyRunAndWrite()` → new `Apify_yyMMdd` tab (filters to `statusCode=200` / `FOUND`) → Chat post with Sheet + Excel export link; shared Apify helpers |
| `GProducts.js` | Product flow variant: task `fh2EbyE6tR2J9Lp26`, fields `asin,countReview,productRating,statusCode,title,url`, rating normalization, writes `Product` via `_overwriteSheet` |
| `Product.js` | Older copy of the Product flow (same task ID, `PRODUCT_SHEET_HEADERS` writer) — **duplicates `GProducts.js`** |
| `trigger.js` | `createApifyWeekdayTriggers()` / `deleteApifyWeekdayTriggers()` — one-off Mon–Fri 04:00 KST triggers |
| `UI.js` | `onOpen()` menus, uploader dialog RPCs, `uiRunProductNow()` |
| `Main.js` | Sheet → Monday core (copied from Glx26), `countAnyColoredCells()` |
| `uploader_sidebar.html` | Uploader dialog HTML |
| `Gemini.js` | `DR(inputText, category)`, `clearDRCache()` |
| `config.js` | `CONFIG`, `getSpreadsheetId_()`, `CHAT_WEBHOOK_URL`, `getPollDelayMs_()`, `PREFERRED_HEADERS` |
| `국내.js` | `updateScoreSheet()` — same leftover as in `iPh17e_Monday` (writes to a different spreadsheet) |
| `appsscript.json` | GAS manifest (V8, Asia/Seoul) |

---

## Config (`config.js` / `GProducts.js`)

| Key | Value |
|-----|-------|
| `CONFIG.actorTaskIdOrSlug` (review task) | `TgpoNoMcN4a5bYsyX` |
| `CONFIG.sheetBaseName` | `Apify` → tabs named `Apify_yyMMdd` (`_2`, `_3` … if taken) |
| `CONFIG.pollIntervalMinutes` / `pollMaxMinutes` | `1` / `180` |
| `CONFIG.POLL_DELAY` / `TEST_POLL_DELAY` / `TEST_MODE` | 2h / 2min / `false` (only used by `getPollDelayMs_()`, not by the recurring poller) |
| `PRODUCT.taskIdOrSlug` | `fh2EbyE6tR2J9Lp26` (`Product` tab) |
| `CHAT_WEBHOOK_URL` | TCK GCX Spigen Google Chat space (hard-coded in `config.js`; value not reproduced here) |

---

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| CX Upload | Open Uploader | `openUploaderDialog()` (not functional — see below) |
| CX Upload → Apify Product | Run Product Now / Cancel Product Polling | `uiRunProductNow()` / `cancelProductPolling()` |
| AI Tools | Summarize selected cells / Run Defect GPT | `uiRunSummarize` / `uiRunDefectGPT` — not defined |

Script-editor entry points: `startApifyRunAndSchedulePoll()` (start review scrape), `pollApifyRunAndWrite()`, `createApifyWeekdayTriggers()`. Note: nothing calls `_ensureRecurringPoller_()`, so starting a review run does **not** create the 1-minute poll trigger by itself — the poller must be created separately (the former wrapper `runApifyNowAndPollAfter2Hours` that did this no longer exists).

Custom functions: `=DR(text, category)` (Defect tab, Gemini `gemini-3.1-flash-lite` → fallbacks, 6h cache), `=countAnyColoredCells("A1:B10")` (share of non-empty cells with non-white background).

---

## Known issues (verified against code, 2026-10-07)

- **Duplicate `const PRODUCT`** in both `Product.js` and `GProducts.js`. In Apps Script V8 all files share one global scope, so this should fail with `Identifier 'PRODUCT' has already been declared` when the project loads (every menu/trigger/custom function). Remove one of the two files before relying on this project. If the live GAS project still works, it differs from this repo.
- **Weekday triggers are broken:** `trigger.js` creates one-off triggers for `runApifyNowAndPollAfter2Hours`, but that function is not defined anywhere, so each trigger fails. Window: tomorrow → `2026-12-31`, 04:00 KST, Mon–Fri.
- **Monday upload path is unconfigured:** `UI.js` / `Main.js` reference `UPLOAD_SHEET_ID`, `UPLOAD_SHEET_NAME`, `BOARD_ID`, `LINK_COLUMN_ID`, `COLUMN_OVERRIDES_BY_TITLE`, `DATE_COL_INDEX_1BASED` etc., none of which are defined here; `chooseGroupFromModel()` still maps Galaxy S26 groups.
- `uiRunSummarize` / `uiRunDefectGPT` are referenced by the AI Tools menu but not defined.
- In `GProducts.js` the Chat message link has a stray space (`.../d/<id> /edit#gid=…`).
- `clearDRCache()` calls `removeAll([])` (no-op).
- Earlier README listed task `TvUlCaUpNvjgC23g5` and an "Apify → Product" menu — both were wrong.

---

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (required) |
| `GEMINI_API_KEY` | Gemini API key (for `=DR()`) |
| `MONDAY_API_KEY` | Only needed if the Monday uploader is ever configured |
| `APIFY_LAST_*`, `PRODUCT_LAST_*` | Run state, written/cleared by the script |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/Power_Acc_Apify
clasp push --force
```
