# iPh17e_Monday (iPhone 17e — Sheet → Monday.com upload)

Container-bound Google Apps Script for the iPhone 17e review spreadsheet. Its main job is pushing bad-review rows from the `1-3점` tab into the **📌iPhone 17e Case+CP** Monday board (`18419272697`), skipping rows whose Review Link is already on the board. It also carries the Apify **Product** task runner (per-ASIN rating / review count → `Product` tab) and the `=DR(text, category)` Gemini 인입사유 classifier. Same bundle pattern as [GlxZ8_Apify](../GlxZ8_Apify/) / [Glx26_Apify](../Glx26_Apify/); the review scrape itself lives in [iPh17e_Apify](../iPh17e_Apify/).

**Script ID:** `19uV2r_JqkysVWdtCxcbT8rlLNnGX-BPZwP4kuvvR33-kppOmB5HDbZyY`
**Linked spreadsheet:** `16xRJHH7Ynii4erNOn_905ST4CZs6OLpOYTof4uqsGsQ` (tab `1-3점`)
**Monday.com board:** `18419272697` (📌iPhone 17e Case+CP, single group `iPhone 17e`)

## Screenshots

![iPhone 17e `1-3점` tab that the uploader sends to monday (reviewer names blurred)](docs/sheet_1-3.jpg)
*iPhone 17e `1-3점` tab that the uploader sends to monday (reviewer names blurred)*

---

## Files

| File | Purpose |
|------|---------|
| `main.js` | Sheet → Monday core: `syncSheetToMonday_core()`, board meta (cached in Script Properties), existing-link scan, row → column-value formatting, batched `create_item` with retry/backoff |
| `UI.js` | `onOpen()` menus, uploader dialog + sidebar RPCs (`getPlanMeta`, `getPlanPage`, `uploadOnePlannedItem`), `uiRunProductNow()` |
| `uploader_sidebar.html` | "Sheet → Monday Uploader" modeless dialog (date-filtered, paged plan → upload) |
| `config.js` | Sheet ID/tab, `BOARD_ID`, all Monday column IDs, `COLUMN_OVERRIDES_BY_TITLE`, behavior flags, `CONFIG` |
| `Products.js` | Apify Product task (`FnpokY7cRrNbI7EIe`) → recurring 1-min poll → `Product` tab |
| `Code.js` | Shared Apify helpers + legacy review-scrape flow (not wired to menu) |
| `Gemini.js` | `DR(inputText, category)`, `clearDRCache()`, `DEBUG_DR()` |
| `국내.js` | `updateScoreSheet()` — copies `국내 고객배드리뷰` rows from another spreadsheet into a `1-3점` tab (see Known issues) |
| `appsscript.json` | GAS manifest (V8, Asia/Seoul) |

---

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| CX Upload | Open Uploader | `openUploaderDialog()` |
| CX Upload → Apify Product | Run Product Now / Cancel Product Polling | `uiRunProductNow()` / `cancelProductPolling()` |
| AI Tools | Summarize selected cells / Run Defect GPT | `uiRunSummarize` / `uiRunDefectGPT` — **not defined in code** (menu items error) |

Other runnable functions: `syncSheetToMonday()` (full upload, logs result), `syncSheetToMondayInteractive()`.

### Upload logic

- Reads displayed values of `UPLOAD_SHEET_ID` / `1-3점`. Item name = `Review Title`; rows without a title or `Review Link` are skipped.
- Dedup: all existing `Review Link` values on the board (`link_mm4nf4xm`) are fetched and normalized; matching rows are skipped.
- Sheet headers are mapped to board columns via `COLUMN_OVERRIDES_BY_TITLE` (Review Link, Image URL, ASIN, Review Ratings, Reviewer, 본문, 인입사유, 키워드 (AI 요약), 인입사유(AI), Review ID, SKU, 기종명, 모델명, 대분류, 생산업체, 원산지, Created/Update 날짜, 사진 유무, 국가).
- `클레임/리뷰` defaults to `Review`. All items go into group `iPhone 17e`.
- The uploader dialog filters by the `Update 날짜` column (`DATE_COL_INDEX_1BASED = 15`) and plans in pages of 250 rows.
- `DRY_RUN = false`.

### Product run

Same as the other Apify projects: POST `actor-tasks/FnpokY7cRrNbI7EIe/runs`, trigger `pollProductRunAndWrite` every 1 min (max 180 min), rewrite the `Product` tab (`country | asin | title | countReview | productRating | url`). No daily/scheduled trigger is created.

### `=DR(text, category)`

Picks a 인입사유 from the `Defect` tab rows whose col A equals `category` (col B label, col C description). Keyword fast-path (`keywordFallback_`) first, then Gemini `gemini-3.1-flash-lite` → `gemini-2.5-flash-lite` → `gemini-3.5-flash`. Cached 6h (`DR_v21_` key prefix). Note: `clearDRCache()` calls `removeAll([])`, which is a no-op — to invalidate, bump the cache-key prefix.

---

## Known issues

- `uiRunSummarize` / `uiRunDefectGPT` (AI Tools menu) are referenced but not defined.
- `CHAT_WEBHOOK_URL` is not defined in this project, so the Product-completion Chat post fails (caught and logged; the sheet write still happens).
- Legacy `startApifyRunAndSchedulePoll()` in `Code.js` needs `CONFIG.actorTaskIdOrSlug`, which is not defined.
- `국내.js` `updateScoreSheet()` reads `1UVXNdfYlGxCCkhwhcmRZsxsHvctiAi6DX3LmuMk-DjU` / `국내 고객배드리뷰` and **clears + overwrites** `1-3점` in `1qs03gqcnDo9t94BrqPCcN0nYjCAE43sL7BmS8bB7kOQ` — not the iPhone 17e sheet. It looks like a shared leftover; don't run it from here unless that's intended.

---

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token |
| `MONDAY_API_KEY` | Monday.com API token |
| `GEMINI_API_KEY` | Gemini API key (for `=DR()`) |
| `BOARD_META_<boardId>` | Cached board columns/groups (written by script) |
| `PRODUCT_LAST_RUN_ID` etc. | Product run state (written/cleared by script) |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/iPh17e_Monday
clasp push --force
```
