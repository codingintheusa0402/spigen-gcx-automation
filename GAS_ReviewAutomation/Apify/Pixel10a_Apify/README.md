# Pixel10a_Apify (Google Pixel 10a)

Container-bound Google Apps Script for the Pixel 10a review spreadsheet. From the sheet menu it starts an Apify **Product** task (per-ASIN rating / review count), polls the run every minute, writes the result into the `Product` tab and posts a completion message to Google Chat. It also provides the `=DR()` custom function (Gemini-based 인입사유 classification against the `Defect` tab).

**Script ID:** `1Ah4m3-STEzY7tfURUhgRIDub-m7sI8NKE54mF2YUtWXKQE1ehHkAh1MI`
**Linked spreadsheet:** `1BpeGq5gIr4tNsPZmnHr19NNY6pQ6sb2_H-v3V9-It4E` (Pixel 10a review sheet; also `UPLOAD_SHEET_ID` in `config.js`)

## Screenshots

![`Product` tab filled by the Apify product task](docs/product_tab.jpg)
*`Product` tab filled by the Apify product task*

---

## Files

| File | Purpose |
|------|---------|
| `Products.js` | **Active flow.** `PRODUCT` config (Apify task `O8wZO63UsU9CTr7ph`), start run → recurring poll → write `Product` tab → Chat notify |
| `Code.js` | Shared Apify helpers (token, Chat post, paginated dataset fetch, flatten/overwrite, Excel export URL) + legacy review-scrape flow (`startApifyRunAndSchedulePoll` / `pollApifyRunAndWrite`) |
| `UI.js` | `onOpen()` menu + `menuRunProduct()` confirm dialog |
| `Gemini.js` | `DR(inputText)` custom function — Gemini 인입사유 classifier, 6h cache |
| `config.js` | `UPLOAD_SHEET_ID` / `UPLOAD_SHEET_NAME`, `CHAT_WEBHOOK_URL`, `CONFIG` (poll interval/max, timezone), `PREFERRED_HEADERS`, group mapping |
| `appsscript.json` | GAS manifest (V8, Asia/Seoul) |

---

## How it works

1. **Apify → Product → Run Product (auto polling)** → confirm dialog → `runProductNowAndPollRecurring()`.
   If `PRODUCT_LAST_RUN_ID` is already set, no new run is started; the poller is just re-ensured.
2. `startProductRun_()` POSTs to `actor-tasks/O8wZO63UsU9CTr7ph/runs` (task's saved input) and stores the run state in Script Properties.
3. A time-based trigger `pollProductRunAndWrite` runs **every `CONFIG.pollIntervalMinutes` (1) minute** until the run ends or `pollMaxMinutes` (180) elapses.
4. On `SUCCEEDED`: fetches only `asin,countReview,productRating,url,title,globalReviews`, clears and rewrites the **`Product`** tab with `country | asin | title | countReview | productRating | url` (country parsed from the Amazon domain), then posts "Apify Product Scraping Completed" to `CHAT_WEBHOOK_URL` and deletes the poll trigger.
5. On `FAILED` / `ABORTED` / `TIMED-OUT` / timeout: state + trigger are cleaned up and a toast is shown.

No daily/scheduled trigger is created by this project — runs are manual from the menu.

### `=DR(text)`

Classifies a review text into one 인입사유 from the `Defect` tab (col B label, col C description). Tries `gemini-3.1-flash-lite` → `gemini-2.5-flash-lite` → `gemini-3.5-flash`; results cached 6h (`DR_v3_` key prefix). Requires `GEMINI_API_KEY`.

---

## Config (`config.js` / `Products.js`)

| Key | Value |
|-----|-------|
| `PRODUCT.taskIdOrSlug` | `O8wZO63UsU9CTr7ph` |
| `PRODUCT.sheetBaseName` | `Product` |
| `CONFIG.pollIntervalMinutes` / `pollMaxMinutes` | `1` / `180` |
| `CONFIG.timezone` | `Asia/Seoul` |
| `CHAT_WEBHOOK_URL` | TCK GCX Spigen Google Chat space (hard-coded in `config.js`; value not reproduced here) |

---

## Known issues / leftovers

- **Legacy review flow is not runnable:** `startApifyRunAndSchedulePoll()` (Code.js) requires `CONFIG.actorTaskIdOrSlug`, which is no longer defined in `config.js`, so it throws `CONFIG.actorTaskIdOrSlug missing.` It is not wired to the menu. (Earlier README listed task `TvUlCaUpNvjgC23g5` / a 2-hour `POLL_DELAY` — neither exists in the current code.)
- `config.js` header comment, `GROUP_TITLES` (`Galaxy S26` / `Plus` / `Ultra`) and `MODEL_HEADER_CANDIDATES` are copy-over from the Galaxy S26 project and are unused here (no Monday sync code in this project).
- The Chat webhook URL (with key/token) is hard-coded in `config.js`; moving it to a Script Property would be safer.

---

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token (required) |
| `GEMINI_API_KEY` | Gemini API key (for `=DR()`) |
| `PRODUCT_LAST_RUN_ID`, `PRODUCT_LAST_DATASET_ID`, `PRODUCT_LAST_POLL_STARTED_AT_MS` | Run state, written/cleared by the script |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/Pixel10a_Apify
clasp push --force
```
