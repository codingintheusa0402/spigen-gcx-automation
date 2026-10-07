# 전략폰_Apify (Strategy Phones — Apify trigger)

Container-bound Google Apps Script for the 전략폰 review spreadsheet. From the sheet menu it starts the Apify **Product** task (per-ASIN rating / review count), polls every minute and rewrites the `Product` tab. It also contains the generic review-scrape flow (dated `Apify_yyMMdd` tab + Excel export link) and the `=FILTER_WHITE_ROWS()` custom function. There is no Monday push code in this project.

**Script ID:** `1b9TGawGmcDUj0sm2OswDEfkUjSIi8C0b90vsMqkmkTGiYpUlH_TEtoZN`
**Linked spreadsheet:** `1yo8CbLhJkuxrf3eXbAqZCb6qBejZhSR3YOt7nFv97fw` (Spigen_전략폰 Series_CustomerReviews (★1~3) 2026_GCX) — the code uses the *active* spreadsheet (`getSpreadsheetId_()`), not a constant.

> ⚠️ **Known config bug (still present 2026-10-07) — `config.js` still targets Galaxy S26, not 전략폰.**
> `UPLOAD_SHEET_ID` = `1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g` (Galaxy S26 sheet), `BOARD_ID` = `18399593191` (📌Galaxy S26 Case+CP), S26 column IDs and `GROUP_TITLES` = Galaxy S26 / Plus / Ultra — copied from `Glx26_Apify` and never adapted (same class of bug as the old `Pixel11_Apify` Z8 leak).
> Currently **latent**: no file in this project reads these constants (there is no `main.js` / uploader), so nothing is written to the S26 sheet or board today. Fix them before adding any Monday-push code here.

## Screenshots

![`Product` tab filled by the Apify product task](docs/product_tab.jpg)
*`Product` tab filled by the Apify product task*

---

## Files

| File | Purpose |
|------|---------|
| `Products.js` | **Active flow.** `PRODUCT` config (Apify task `cOhcLe65h2fPhKRkH`), start → recurring poll → `Product` tab → Chat notify |
| `Code.js` | Shared Apify helpers + review-scrape flow (`startApifyRunAndSchedulePoll` / `pollApifyRunAndWrite`), Excel export, `FILTER_WHITE_ROWS()` |
| `UI.js` | `onOpen()` menu + `menuRunProduct()` confirm dialog |
| `config.js` | Sheet/board/column constants (**S26 — see bug above**), `PREFERRED_HEADERS`, `CONFIG` (poll 1 min / max 180 min / Asia/Seoul) |
| `trigger.js` | `createApifyWeekdayTriggers()` / `deleteApifyWeekdayTriggers()` |
| `appsscript.json` | GAS manifest |

---

## Usage

Open the linked spreadsheet → **Apify → Product → Run Product (auto polling)** (`menuRunProduct`). The run uses the task's saved input; `pollProductRunAndWrite` runs every minute until `SUCCEEDED` (clears + rewrites `Product`: `country | asin | title | countReview | productRating | url`) or failure / 180-minute timeout. **Cancel Product Polling** removes the poll trigger.

### Custom function

| Function | Purpose |
|----------|---------|
| `FILTER_WHITE_ROWS()` | Returns rows from the `신제품 라인업` tab (cols A:I) whose column-A cell has a white (`#ffffff`) background |

---

## Other known issues

- **Weekday triggers are dead:** `APIFY_TRIGGER_WINDOW` ends `2026-01-14` (05:30 KST, Mon–Fri) and the handler `runApifyNowAndPollAfter2Hours` is commented out in `Code.js`. Running `createApifyWeekdayTriggers()` now creates nothing useful.
- **Review-scrape flow not runnable:** `startApifyRunAndSchedulePoll()` requires `CONFIG.actorTaskIdOrSlug`, which isn't defined.
- **No Chat notification:** `CHAT_WEBHOOK_URL` is not defined in this project; the Product-completion post fails inside a try/catch (logged only).

---

## Script Properties

| Key | Description |
|-----|-------------|
| `APIFY_TOKEN` | Apify API token |
| `PRODUCT_LAST_RUN_ID` etc. | Run state, written/cleared by the script |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/전략폰_Apify
clasp push --force
```
