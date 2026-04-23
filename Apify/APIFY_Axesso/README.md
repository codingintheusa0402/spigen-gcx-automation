# APIFY_Axesso

Google Apps Script project for Apify/Axesso Amazon review scraping and daily distribution into Spigen product spreadsheets.

## Files

| File | Purpose |
|------|---------|
| `Master.gs` | Daily automation — filters, deduplicates, and distributes reviews into destination spreadsheets |
| `Code.gs` | Apify run lifecycle — start task, poll status, write results to source sheet |
| `Product.gs` | Product-level Apify run — aggregate rating/review count per ASIN |

---

## Master.gs

### Trigger

Time-based trigger — runs `dailyJob()` once per day (KST).

### Flow

```
dailyJob()
  │
  ├─ step1_deleteNumberedSheets()      Delete conflict/dated sheets not matching today
  ├─ step2_dedupDatedSheets()          Deduplicate today's dated sheets by Review ID
  ├─ step2b_updateTemSheet()           Refresh `tem` sheet with all active Review IDs
  └─ Per-config loop (SHEET_CONFIGS)
        ├─ has15=true  → _processFilterSheet_()
        └─ has15=false → _processTo13_()
```

### SHEET_CONFIGS field reference

| Field | Description |
|-------|-------------|
| `filterSheet` | Tab name in source spreadsheet holding raw scraped reviews. Must have a named filter view `"finalize"`. |
| `destId` | Google Spreadsheet ID of the destination workbook |
| `countries` | Set of country codes to include (`"US"`, `"UK"`, `"DE"`, `"FR"`, `"ES"`, `"IT"`, `"JP"`, `"IN"`) |
| `numCols` | Number of source columns to copy |
| `has15` | `true` = dest has `1-5점` + `1-3점`; `false` = dest has `1-3점` only |
| `seriesFilter` | Extra column filter e.g. `{ colLetter: "Q", contains: "S26" }`. `null` = skip |
| `temCol` | Column name in `tem` sheet for this product's Review IDs |
| `insertAtTop` | (`has15=false`) `true` = insert at row 2; `false` = append at bottom |
| `ratingFilter` | (`has15=false`) Allowed rating values e.g. `[1,2,3]`. `null` = all |
| `drFormula` | `true` = write `=dr()` into `인입사유(AI)` column after paste |
| `pasteReviewId` | `true` = copy Review ID from source; `false` = leave blank (dest has formula) |

### Current product configs (as of 2026-04-01)

| filterSheet | has15 | Countries | Notes |
|-------------|-------|-----------|-------|
| `Glx26_filter` | true | US FR ES JP UK IN DE IT | seriesFilter Q ⊇ "S26" |
| `iPh17e_filter` | true | US FR ES JP UK IN DE IT | seriesFilter N ⊇ "17e" |
| `Pixel10a_filter` | true | US FR ES JP UK IN DE IT | seriesFilter N ⊇ "10a" |
| `SDA_filter` | false | FR ES JP UK DE IT | |
| `Auto_Acc_filter` | false | FR ES UK DE IT | |
| `Power_Acc_filter` | false | IN | |
| `전략폰_filter` | false | IN | |
| `유지훈P_filter` | false | US FR ES JP UK IN DE IT | insertAtTop, ratingFilter 1-3 |

### Key column names

| Column | Purpose |
|--------|---------|
| `Review ID` | Dedup key |
| `Country` | Country filter |
| `Content` | Source body text |
| `본문` | Dest body column (for `=dr()` formula) |
| `대분류` | Dest category column (for `=dr()` formula) |
| `인입사유(AI)` | AI classification formula target in `1-3점` |
| `키워드 (AI 요약)` | AI summary formula column in `1-5점` |
| `Update 날짜` | Date written on copy (KST) |
| `Rating` | Used by `ratingFilter` |

**Source spreadsheet:** `SRC_ID = 1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`

---

## Code.gs

Manages the Apify run lifecycle for review scraping — starts a task run, polls for completion, writes results to the source spreadsheet, and notifies via Google Chat.

### Key functions

| Function | Description |
|----------|-------------|
| `startApifyRunAndSchedulePoll()` | Starts the Apify task run; saves run ID to Script Properties |
| `pollApifyRunAndWrite()` | Recurring trigger — checks status, writes results when `SUCCEEDED` |
| `_getToken()` | Reads `APIFY_TOKEN` from Script Properties |
| `_postToGoogleChat(text)` | Posts notification to `CHAT_WEBHOOK_URL` |

### Script Properties required

| Property | Value |
|----------|-------|
| `APIFY_TOKEN` | Your Apify API token |
| `CHAT_WEBHOOK_URL` | Google Chat incoming webhook URL (optional) |

---

## Product.gs

Fetches product-level aggregate data (rating, review count) from Apify and writes to a Product sheet.

### Output sheet columns

| Column | Source |
|--------|--------|
| `country` | Parsed from URL (`amazon.de` → `DE`) |
| `asin` | ASIN from Apify dataset |
| `title` | Product title |
| `countReview` | Total review count |
| `productRating` | Aggregate star rating |
| `url` | Full Amazon product URL |

### Key functions

| Function | Description |
|----------|-------------|
| `runProductNowAndPollRecurring()` | Entry point — starts run if none pending, schedules recurring poll |
| `startProductRun_()` | Starts the Apify task run for product data |
| `_scheduleRecurringProductPoll_()` | Creates time-based trigger to poll until complete |

**Task ID:** `fh2EbyE6tR2J9Lp26` · **Sheet base name:** `Product`
