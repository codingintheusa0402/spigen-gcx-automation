# Gemini_DR — Google Apps Script (Glx26 bound)

Container-bound Apps Script project on the **Galaxy S26 review spreadsheet** (`1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g`). Its main job is the **`=DR()` custom function**, which uses Gemini to classify review text (본문) into one defect label (인입사유) from the sheet's `Defect` tab. The same project also hosts a monday.com upload sidebar for the `1-3점` tab, an Apify Product scraper, and a one-off 국내 배드리뷰 copy helper.

clasp script ID: `1sPKcHgYy8kEqrp6Ra_FSw3vpnIVlJgB5dNeFLVTzZrZoptmEeA8lnrMm` (`.clasp.json` is local only — not tracked in git).

## Screenshots

![`=dr(G2,S2)` in the 인입사유(AI) column next to the human 인입사유 label](docs/dr_formula.jpg)
*`=dr(G2,S2)` in the 인입사유(AI) column next to the human 인입사유 label*

---

## Features

### 1. `=DR()` — Gemini 인입사유 classifier (`Gemini.js`)

```
=DR(본문셀)            // classify against every Defect row
=DR(본문셀, 대분류)     // only Defect rows whose col A == 대분류
```

**Flow:**
1. Cache lookup — `CacheService` script cache, 6 h TTL, key = `DR_CACHE_VERSION` (`DR_v22_`) + base64(text|category). Bump the version string to invalidate everything.
2. Load the `Defect` sheet (A `대분류`, B `label`, C `description`); prompt lists `label: description`.
3. Keyword fast path (`keywordFallback_`): heavy/bulky → 두꺼움, yellow → 황변, button → 버튼불량, attach/difficult → 부착어려움, scratch → 스크래치 (only if that label exists in the filtered list). Note: these fast-path hits are **not** cached.
4. Gemini call (`temperature 0`, `maxOutputTokens 20`, thinking off), trying models in order: `gemini-3.1-flash-lite` → `gemini-2.5-flash-lite` → `gemini-3.5-flash`.
5. Strict normalized match → loose contains match against the label list; result cached. No match → `''`. Exceptions → `ERROR: …`.

Utilities: `clearDRCache()` (note: `removeAll([])` is effectively a no-op — bump `DR_CACHE_VERSION` instead), `testDRBatch()` (classifies rows 2–21 of `1-3점`, writes into the `인입사유(AI)` column and logs token usage), `DEBUG_DR(text, category)` (returns JSON trace).

### 2. Monday.com upload sidebar (`main.js`, `UI.js`, `uploader_sidebar.html`)

**CX Upload → Open Uploader** opens a sidebar that uploads one day's rows from `1-3점` to the **📌Galaxy S26 Case+CP** board (`BOARD_ID 18399593191`).

- Filters by `Update 날짜` (col 15) = the date picked in the sidebar; scans in pages of 250 rows.
- Skips rows without `Review Title` / `Review Link`; dedupes against existing board items by `Review Link` (`link_mm0fkspz`).
- Group routing by `기종명/모델명/Model` → `Galaxy S26` / `Galaxy S26 Plus` / `Galaxy S26 Ultra` (fallback: first group).
- Default `클레임/리뷰` status = `리뷰`; `본문` also written to the `자동번역` column (target `ko`).
- `syncSheetToMonday()` / `syncSheetToMonday_core()` = non-sidebar version (no date filter).

### 3. Apify Product scraper (`Products.js`)

**CX Upload → Apify Product → Run Product Now** starts Apify task `hhYN1b5uTF8x8yk4Q`, creates a 1-min recurring trigger for `pollProductRunAndWrite()`, and overwrites the `Product` sheet on success. **Cancel Product Polling** deletes the trigger.

### 4. Apify review scraper (`Code.js`) — legacy, currently broken

`startApifyRunAndSchedulePoll()` / `pollApifyRunAndWrite()` (1-min poller, 180-min timeout) would write a dated `Apify_YYMMDD` sheet, dedupe, and post to Chat. Daily review scraping for Glx26 actually runs from **MasterTrigger** (`masterDailyJob`), not here.

### 5. 국내 배드리뷰 copy (`국내.js`)

`updateScoreSheet()` (manual) copies `국내 고객배드리뷰` from spreadsheet `1UVXNdfYlGxCCkhwhcmRZsxsHvctiAi6DX3LmuMk-DjU` into the `1-3점` tab of spreadsheet `1qs03gqcnDo9t94BrqPCcN0nYjCAE43sL7BmS8bB7kOQ`, remapping columns, country `KOREA`/`KR`, today's KST date. It **clears the whole target tab first** (header included — output has no header row).

---

## Menus (`onOpen`)

- `CX Upload` → Open Uploader · Apify Product ▸ Run Product Now / Cancel Product Polling
- `AI Tools` → Summarize selected cells / Run Defect GPT (selected cells)

## Script Properties

| Key | Used by |
|---|---|
| `GEMINI_API_KEY` | `DR()` |
| `APIFY_TOKEN` | Product / review scrapers |
| `MONDAY_API_KEY` | monday uploader |

Runtime state (auto-managed): `PRODUCT_LAST_RUN_ID`, `PRODUCT_LAST_DATASET_ID`, `PRODUCT_LAST_POLL_STARTED_AT_MS`, `APIFY_LAST_*`.

## Config (`config.js`)

`UPLOAD_SHEET_ID` / `UPLOAD_SHEET_NAME` (`1-3점`), `DATE_COL_INDEX_1BASED` (15), `BOARD_ID`, monday column IDs (`LINK_/CLAIM_REVIEW_/COUNTRY_/PHOTO_/CHANNEL_COLUMN_ID`), `DRY_RUN` (false), `GROUP_TITLES`, `PREFERRED_HEADERS`, `CONFIG` (`pollIntervalMinutes 1`, `pollMaxMinutes 180`, `timezone Asia/Seoul`).

## Known issues (as of current code)

- `Code.js` review scraper: `CONFIG.actorTaskIdOrSlug` is not defined in `config.js`, so `startApifyRunAndSchedulePoll()` throws immediately; `_postToGoogleChat` references a global `CHAT_WEBHOOK_URL` constant that is not defined anywhere (would be a ReferenceError, caught and logged).
- `AI Tools` menu items point at `uiRunSummarize` / `uiRunDefectGPT`, which do not exist in this project — clicking them errors.
- `=DR()` results are cached for 6 h, so edits to the `Defect` sheet don't show up until the cache expires or `DR_CACHE_VERSION` is bumped.

## Deploy

```bash
cd Gemini_DR
clasp push --force
```

> Always push from `Gemini_DR/` — never from `Apify/APIFY_Axesso/` or `MasterTrigger/`, which are different script projects.
