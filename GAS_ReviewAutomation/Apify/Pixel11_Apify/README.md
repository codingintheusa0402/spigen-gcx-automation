# Pixel11_Apify (Google Pixel 11 — Apify Product + Monday.com + DR())

Container-bound Google Apps Script for the **Google Pixel 11 Series 리뷰 모니터링 스프레드시트**. Copied from [GlxZ8_Apify](../GlxZ8_Apify/) (created 2026-08-05) and repointed at the Pixel 11 sheet/board: it runs the Apify product-details task into the `Product` tab, uploads new 1–3★ reviews from `1-3점` to the monday.com board **📌Pixel 11 Case+CP** via a sidebar uploader, and provides the `=DR()` Gemini 인입사유 classifier.

The daily review scrape itself (Apify → `1-3점`) is driven by `MasterTrigger/`, not by this script.

**Script ID:** `1vB_8lzDdEW6K8so6mJL1HeeFn4NkofipdscR_hpUPC4_oyEcVPhsmCvU`
**Linked spreadsheet:** `12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI`
**Monday.com board:** `18425190666` (📌Pixel 11 Case+CP) — exact column-ID duplicate of the Z8/S26 board template

## Screenshots

![Pixel 11 `1-3점` tab (reviewer names blurred)](docs/sheet_1-3.jpg)
*Pixel 11 `1-3점` tab (reviewer names blurred)*

---

## Differences from GlxZ8_Apify

| Area | Pixel11 |
|------|---------|
| `config.js` | `UPLOAD_SHEET_ID`, `BOARD_ID` above; `GROUP_TITLES` = Pixel 11 / Pixel 11 Pro / Pixel 11 Pro XL / Pixel 11 Pro Fold |
| `main.js` | `chooseGroupFromModel()` most-specific first: `pro fold` → `pro xl` → `pro` → `pixel 11` |
| `Products.js` | `PRODUCT.taskIdOrSlug = 'n8VIp2RTIYxAaXL8S'` (Pixel 11's own product-details task) |
| `DefectDefsPatch.js` | Not included (the 정의 master is shared; patch from GlxZ8) |

Everything else (`UI.js`, `uploader_sidebar.html` incl. the Created/Update 날짜 re-apply pass, `Code.js`, `Gemini.js`, `국내.js`) is byte-identical to GlxZ8 — see the [GlxZ8 README](../GlxZ8_Apify/README.md) for how each piece works.

## Files

`Code.js`, `config.js`, `main.js`, `UI.js`, `uploader_sidebar.html`, `Products.js`, `Gemini.js`, `국내.js`, `appsscript.json` — same roles as in GlxZ8.

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| **CX Upload** | Open Uploader | `openUploaderDialog` — `1-3점` rows whose Update 날짜 (col 15) = chosen date → board 18425190666, deduped by Review Link, dates re-applied after creation |
| CX Upload → Apify Product | Run Product Now / Cancel Product Polling | `uiRunProductNow` / `cancelProductPolling` |
| **AI Tools** | Summarize selected cells / Run Defect GPT | **undefined in code** |

- `=DR(text, category)` — reads the `Defect` tab (IMPORTRANGE mirror of master `GCX 인입사유`), cache key `DR_v28_`.

## Triggers

Only the Product poller: every-1-minute trigger on `pollProductRunAndWrite`, self-deleting on completion / failure / 180 min timeout.

## Script Properties required

`APIFY_TOKEN`, `MONDAY_API_KEY`, `GEMINI_API_KEY`

---

## History (since 2026-08)

- **2026-08-10** — `Products.js` task ID was still GlxZ8's product task (`w3jI45UjmyFLwVHek`); fixed to Pixel 11's own `n8VIp2RTIYxAaXL8S`. `DR()` `keywordFallback_()` removed, cache key → `DR_v28_`.
- **2026-08-06** — **Z8-leak fix.** `config.js` was an uncorrected copy of GlxZ8's: `UPLOAD_SHEET_ID` pointed at the Z8 sheet and `BOARD_ID` at 18421346787. Before the fix, 2 Galaxy Z Fold 8 reviews were pushed live onto the Pixel 11 board (left in place for the user to decide). Also rewrote `chooseGroupFromModel()` (it only matched `fold 8`/`flip 8`, so every Pixel review would fall to the first group) and synced `Gemini.js`.

## Known issues

- **Never runtime-verified.** The 2026-08-06 / 08-10 fixes are code-reviewed only: the script had never been authorized under the clasp account, so `clasp run` / Execution API calls were blocked pending a one-time interactive OAuth consent. No later commit records a verified run. First real run should be watched (check that new items land in the right Pixel 11 group).
- The 2 mis-routed Z Fold 8 items on board 18425190666 may still be there (not deleted by the fix).
- Inherited from the template: `CHAT_WEBHOOK_URL` undefined (Product Chat post fails silently), dead `startApifyRunAndSchedulePoll()`, undefined `uiRunSummarize` / `uiRunDefectGPT` / `translateTextAuto`.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/Pixel11_Apify
clasp push --force
```
