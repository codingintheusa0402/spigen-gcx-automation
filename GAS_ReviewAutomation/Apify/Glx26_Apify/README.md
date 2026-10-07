# Glx26_Apify (Galaxy S26 — Apify Product + Monday.com + DR())

Container-bound Google Apps Script for the **Galaxy S26 리뷰 모니터링 스프레드시트**. It does three things inside that sheet: (1) runs the Apify **product-details** task and writes per-ASIN rating / review count to the `Product` tab, (2) uploads new 1–3★ reviews from the `1-3점` tab to the monday.com board **📌Galaxy S26 Case+CP** through a sidebar uploader, and (3) provides the `=DR()` custom function that classifies review text into a GCX 인입사유 with Gemini. This is the original template that [GlxZ8_Apify](../GlxZ8_Apify/), [Pixel11_Apify](../Pixel11_Apify/) and [iPhone18_Apify](../iPhone18_Apify/) were copied from.

The daily review scrape itself (Apify → `1-3점`) is driven by `MasterTrigger/` (masterDailyJob), not by this script.

**Script ID:** `1sPKcHgYy8kEqrp6Ra_FSw3vpnIVlJgB5dNeFLVTzZrZoptmEeA8lnrMm`
**Linked spreadsheet:** `1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g` (Galaxy S26 review sheet)
**Monday.com board:** `18399593191` (📌Galaxy S26 Case+CP)

## Screenshots

![Glx26 `1-3점` tab (reviewer names blurred)](docs/sheet_1-3.jpg)
*Glx26 `1-3점` tab (reviewer names blurred)*

---

## Files

| File | Purpose |
|------|---------|
| `UI.js` | `onOpen()` menus, `_getToken()`, uploader dialog + sidebar RPCs (`getPlanMeta`, `getPlanPage`, `uploadOnePlannedItem`), `uiRunProductNow` |
| `uploader_sidebar.html` | "Sheet → Monday Uploader" modeless dialog (date picker → Prepare → Run upload) |
| `main.js` | Sheet → monday core: board schema fetch, existing-link dedupe, column mapping, batched `create_item` with retry/backoff, `chooseGroupFromModel()` |
| `config.js` | Sheet/board IDs, monday column IDs, header names, group titles, `CONFIG` (poll interval/timeout/timezone) |
| `Products.js` | Apify product-details task run + recurring poller → `Product` tab |
| `Gemini.js` | `=DR()` custom function (Gemini 인입사유 classifier) + `clearDRCache`, `DEBUG_DR` |
| `Code.js` | Shared Apify helpers (dataset fetch, flatten → sheet writer, poll triggers, Excel export URL, `_postToGoogleChat`). Also holds the legacy review-run starter `startApifyRunAndSchedulePoll` / `pollApifyRunAndWrite` (see Known issues) |
| `국내.js` | `updateScoreSheet()` — legacy helper that copies the 국내 고객배드리뷰 sheet into another book's `1-3점` (hardcoded URLs, not S26-specific) |
| `appsscript.json` | Manifest (V8, Asia/Seoul; scopes: container UI, spreadsheets, drive, external_request, scriptapp) |

---

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| **CX Upload** | Open Uploader | `openUploaderDialog` — sidebar: pick a date (`yyyy.mm.dd`) → rows of `1-3점` whose **Update 날짜** (col 15 / O) equals that date are planned in pages of 250, skipping rows already on the board (dedupe by Review Link), then created one by one |
| CX Upload → Apify Product | Run Product Now | `uiRunProductNow` → `runProductNowAndPollRecurring()` |
| CX Upload → Apify Product | Cancel Product Polling | `cancelProductPolling` |
| **AI Tools** | Summarize selected cells / Run Defect GPT | `uiRunSummarize` / `uiRunDefectGPT` — **not defined in this codebase** (menu items error) |

Other public functions: `syncSheetToMonday()` (non-paged, whole-sheet upload, logs only), `syncSheetToMondayInteractive()`, `countAnyColoredCells(rangeA1)` (custom function), `testProductFetchFix()`.

### `=DR(text, category)` custom function
- Reads the `Defect` tab (A=Category, B=인입사유, C=정의) — an `IMPORTRANGE` mirror of the master `GCX 인입사유` sheet. Edit the master, never the mirror.
- Sends the whole category's label+정의 list to Gemini (`gemini-3.1-flash-lite` → `gemini-2.5-flash-lite` → `gemini-3.5-flash` fallback) and returns the exact label (strict, then loose match) or `''`.
- Cached 6 h in `CacheService` under key prefix `DR_v28_`. Bump the prefix when 정의 change so cached results re-evaluate.

### Monday upload behaviour
- Item name = `Review Title`; `Review Link` → link column `link_mm0fkspz`; 클레임/리뷰 status (`color_mm0f7bwq`) defaults to `리뷰`.
- Group = `chooseGroupFromModel(기종명)` → `Galaxy S26` / `Galaxy S26 Plus` / `Galaxy S26 Ultra`, else the board's first group.
- Skipped column types: mirror, subtasks, integration, board_relation. `DRY_RUN = false`.

---

## Triggers

No fixed schedule. `runProductNowAndPollRecurring()` creates a **time-based trigger every `CONFIG.pollIntervalMinutes` (1) min** on `pollProductRunAndWrite`, which deletes itself on SUCCEEDED / FAILED / ABORTED / TIMED-OUT or after `CONFIG.pollMaxMinutes` (180). On OOM (exitCode 137) it retries once with a fresh 4096 MB run. Results are deduped by ASIN+URL and overwrite the `Product` tab.

## Config (`config.js` / `Products.js`)

| Key | Value |
|-----|-------|
| `UPLOAD_SHEET_ID` / `UPLOAD_SHEET_NAME` | `1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g` / `1-3점` |
| `DATE_COL_INDEX_1BASED` | `15` (Update 날짜) |
| `BOARD_ID` | `18399593191` |
| `GROUP_TITLES` | Galaxy S26, Galaxy S26 Plus, Galaxy S26 Ultra |
| `PRODUCT.taskIdOrSlug` | `g2egqkSfZrtZ8f4ts` (Apify product-details task) → tab `Product` |
| `CONFIG` | `pollIntervalMinutes: 1`, `pollMaxMinutes: 180`, `timezone: 'Asia/Seoul'` |

## Script Properties required

| Key | Used by |
|-----|---------|
| `APIFY_TOKEN` | Product task run/poll |
| `MONDAY_API_KEY` | Uploader / `syncSheetToMonday` |
| `GEMINI_API_KEY` | `=DR()` |

Internal state keys written by the poller: `PRODUCT_LAST_RUN_ID`, `PRODUCT_LAST_DATASET_ID`, `PRODUCT_LAST_POLL_STARTED_AT_MS`, `PRODUCT_OOM_RETRIED`.

---

## Recent changes (since 2026-08)

- **2026-08-10** — Removed the `keywordFallback_()` English-substring shortcut that ran before Gemini in `DR()` (it ignored category and emitted invalid labels). Cache key → `DR_v28_` (same across Glx26/GlxZ8/Pixel11/iPhone18).
- **2026-08-06** — Cache key bumped `DR_v21_` → `DR_v25_` to pick up the master 인입사유 정의 fixes; `clearDRCache()` no-op fix carried over.

## Known issues

- Missing vs. the later copies: no `'formula'` column exclusion (monday rejects writes to formula columns — fixed in GlxZ8/Pixel11/iPhone18) and no `reapplyDateColumns()` pass for boards whose automation overwrites Created/Update 날짜. Only matters if the S26 board has those.
- `CHAT_WEBHOOK_URL` is not defined anywhere in the repo, so the Product "completed" Chat post fails inside a try/catch (logged as `Chat notify failed`).
- `startApifyRunAndSchedulePoll()` needs `CONFIG.actorTaskIdOrSlug`, which `CONFIG` doesn't have — the legacy in-script review run is dead code (MasterTrigger does it).
- `uiRunSummarize`, `uiRunDefectGPT`, `translateTextAuto` are referenced but undefined; the 자동번역 column is therefore silently left blank.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/Glx26_Apify
clasp push --force
```
