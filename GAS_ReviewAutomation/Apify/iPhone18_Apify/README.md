# iPhone18_Apify (iPhone 18 — Apify Product + Monday.com + DR())

Container-bound Google Apps Script for the **iPhone 18 Series 리뷰 모니터링 스프레드시트**. Copied from [GlxZ8_Apify](../GlxZ8_Apify/) on 2026-09-14 and repointed at the iPhone 18 sheet/board: it uploads new 1–3★ reviews from `1-3점` to the monday.com board **📌iPhone 18 Case+CP** via a sidebar uploader, runs the Apify product-details task into the `Product` tab, provides the `=DR()` Gemini 인입사유 classifier, and carries a copy of `DefectDefsPatch.js`.

The daily review scrape itself (Apify → `1-3점`) is driven by `MasterTrigger/`, not by this script. (The sheet's SOP tab still says "GLX Z8" — stale text; this is the iPhone 18 book.)

**Script ID:** `1u89bbw5SO41F_d4qyGIhKZRHEQzgv4XD2hcdaSHuLPc2Vetpwo3Gz_BO`
**Linked spreadsheet:** `1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU`
**Monday.com board:** `18430082360` (📌iPhone 18 Case+CP) — same column IDs as the Z8/S26 board template

## Screenshots

![iPhone 18 `1-3점` tab (reviewer names blurred)](docs/sheet_1-3.jpg)
*iPhone 18 `1-3점` tab (reviewer names blurred)*

---

## Differences from GlxZ8_Apify

| Area | iPhone18 |
|------|----------|
| `config.js` | `UPLOAD_SHEET_ID`, `BOARD_ID` above; `GROUP_TITLES` = iPhone 18 Pro / iPhone 18 Pro Max / iPhone Duo |
| `main.js` | `chooseGroupFromModel()`: `duo` → `pro max` → `pro` (mapping checked against 기종명 values in the 신제품 라인업 sheet). Anything else (e.g. a plain "iPhone 18") falls to the board's first group |
| `Products.js` | **Identical to GlxZ8** — see Known issues |

Everything else (`UI.js`, `uploader_sidebar.html` incl. the Created/Update 날짜 re-apply pass, `Code.js`, `Gemini.js`, `DefectDefsPatch.js`, `국내.js`) is byte-identical to GlxZ8 — see the [GlxZ8 README](../GlxZ8_Apify/README.md) for details.

## Files

`Code.js`, `config.js`, `main.js`, `UI.js`, `uploader_sidebar.html`, `Products.js`, `Gemini.js`, `DefectDefsPatch.js`, `국내.js`, `appsscript.json` — same roles as in GlxZ8. `DefectDefsPatch.js` edits the shared master `GCX 인입사유`, so run it from one project only (GlxZ8), not here.

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| **CX Upload** | Open Uploader | `openUploaderDialog` — `1-3점` rows whose Update 날짜 (col 15) = chosen date → board 18430082360, deduped by Review Link, dates re-applied after creation |
| CX Upload → Apify Product | Run Product Now / Cancel Product Polling | `uiRunProductNow` / `cancelProductPolling` |
| **AI Tools** | Summarize selected cells / Run Defect GPT | **undefined in code** |

- `=DR(text, category)` — reads the `Defect` tab (IMPORTRANGE mirror of master `GCX 인입사유`), cache key `DR_v28_`.

## Triggers

Only the Product poller: every-1-minute trigger on `pollProductRunAndWrite`, self-deleting on completion / failure / 180 min timeout.

## Script Properties required

`APIFY_TOKEN`, `MONDAY_API_KEY`, `GEMINI_API_KEY` (already present in the copied GAS project per the 2026-09-14 commit).

---

## History

- **2026-09-14** — Project added (duplicated from GlxZ8; only sheet ID, board ID, group titles and `chooseGroupFromModel()` changed).

## Known issues

- **`Products.js` still uses GlxZ8's product-details task** (`PRODUCT.taskIdOrSlug = 'w3jI45UjmyFLwVHek'`) — the same copy-paste leftover fixed for Pixel11 on 2026-08-10. "Run Product Now" will overwrite this book's `Product` tab with **Galaxy Z8** ASIN ratings/review counts until it is pointed at an iPhone 18 task.
- Inherited from the template: `CHAT_WEBHOOK_URL` undefined (Product Chat post fails silently), dead `startApifyRunAndSchedulePoll()`, undefined `uiRunSummarize` / `uiRunDefectGPT` / `translateTextAuto`.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/iPhone18_Apify
clasp push --force
```
