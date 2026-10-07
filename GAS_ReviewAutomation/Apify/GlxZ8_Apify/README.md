# GlxZ8_Apify (Galaxy Z8 — Apify Product + Monday.com + DR())

Container-bound Google Apps Script for the **Galaxy Z Fold8 / Flip8 / Fold8 Ultra 리뷰 모니터링 스프레드시트**. Copy of [Glx26_Apify](../Glx26_Apify/)'s bundle repointed at the Z8 sheet/board: it runs the Apify product-details task into the `Product` tab, uploads new 1–3★ reviews from `1-3점` to the monday.com board **📌Galaxy Z8 Case+CP** via a sidebar uploader, and provides the `=DR()` Gemini 인입사유 classifier. It also holds `DefectDefsPatch.js`, the script used to tune the shared GCX 인입사유 정의 master. This version is the template that [Pixel11_Apify](../Pixel11_Apify/) and [iPhone18_Apify](../iPhone18_Apify/) were copied from.

The daily review scrape itself (Apify → `1-3점`) is driven by `MasterTrigger/`, not by this script.

**Script ID:** `1iTOJijA5cwc35U_ECEpTXclgPWll7ivCmOK5zYXprwK7QKcwUGWmlsh8`
**Linked spreadsheet:** `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4`
**Monday.com board:** `18421346787` (📌Galaxy Z8 Case+CP) — same column IDs as the S26 board

## Screenshots

![GlxZ8 `1-3점` tab filled by the Apify review run (reviewer names blurred)](docs/sheet_1-3.jpg)
*GlxZ8 `1-3점` tab filled by the Apify review run (reviewer names blurred)*

---

## Differences from Glx26_Apify

| Area | GlxZ8 |
|------|-------|
| `config.js` | Z8 sheet/board IDs; `GROUP_TITLES` = Galaxy Z Fold 8 / Galaxy Z Flip 8 / Galaxy Z Fold 8 Ultra; adds `DATE_CREATED_COLUMN_ID` (`date_mm0f80th`), `DATE_UPDATE_COLUMN_ID` (`date_mm0fyd4a`), `DATE_FIX_DELAY_MS` (3000) |
| `main.js` | `chooseGroupFromModel()` matches `fold 8 ultra` → `flip 8` → `fold 8` (Ultra first, handles combo labels like "Z Fold 8 Ultra / Z Fold 7"); `'formula'` columns excluded from writes |
| `UI.js` | `'formula'` excluded in `getPlanPage`; `uploadOnePlannedItem` returns `itemId` + `dateFix`; new `reapplyDateColumns(entries)` |
| `uploader_sidebar.html` | After the create loop, calls `reapplyDateColumns` once to re-write Created/Update 날짜 (the board's "when item created" automation stamps today's date over them) |
| `Products.js` | `PRODUCT.taskIdOrSlug = 'w3jI45UjmyFLwVHek'` |
| `DefectDefsPatch.js` | Extra file (below) |

`Code.js`, `Gemini.js`, `국내.js` are byte-identical to Glx26.

---

## Files

| File | Purpose |
|------|---------|
| `UI.js` | `onOpen()` menus, `_getToken()`, uploader RPCs (`getPlanMeta`, `getPlanPage`, `uploadOnePlannedItem`, `reapplyDateColumns`), `uiRunProductNow` |
| `uploader_sidebar.html` | "Sheet → Monday Uploader" dialog (date → Prepare → Run upload → date re-apply) |
| `main.js` | Sheet → monday core (`syncSheetToMonday_core`, batched create with retry/backoff, link dedupe, group routing) |
| `config.js` | Sheet/board/column IDs, header names, group titles, `CONFIG` |
| `Products.js` | Apify product-details task + recurring poller → `Product` tab |
| `Gemini.js` | `=DR()` custom function, `clearDRCache`, `DEBUG_DR` |
| `DefectDefsPatch.js` | `patchDefectDefinitions()` / `auditDefectDefinitions()` — patches 정의 (col C) in the master |
| `Code.js` | Shared Apify helpers + legacy review-run starter (dead, see Known issues) |
| `국내.js` | `updateScoreSheet()` — legacy 국내 고객배드리뷰 copier with hardcoded URLs; not Z8-specific |
| `appsscript.json` | Manifest (V8, Asia/Seoul) |

## Menus / entry points

| Menu | Item | Function |
|------|------|----------|
| **CX Upload** | Open Uploader | `openUploaderDialog` — rows of `1-3점` whose Update 날짜 (col 15) = chosen date, deduped by Review Link against the board, created one by one, then dates re-applied |
| CX Upload → Apify Product | Run Product Now / Cancel Product Polling | `uiRunProductNow` / `cancelProductPolling` |
| **AI Tools** | Summarize selected cells / Run Defect GPT | `uiRunSummarize` / `uiRunDefectGPT` — **undefined in code** |

- `=DR(text, category)` — reads the `Defect` tab (IMPORTRANGE mirror of master `GCX 인입사유`), Gemini fallback chain `gemini-3.1-flash-lite` → `gemini-2.5-flash-lite` → `gemini-3.5-flash`, 6 h cache, key prefix `DR_v28_`.
- `patchDefectDefinitions()` — writes `DEFS_PATCH` (replace/append by (Category, 인입사유) pair, idempotent) to master `1jC2k5Cssj4nYI6R1gMlqtpyVTeDUIc-hTirdY3Ij8mk` → `GCX 인입사유`. Never writes the `Defect` mirror (would break the IMPORTRANGE array). Changes propagate to every monitoring book's `Defect` tab; bump `DR_v` afterwards. `auditDefectDefinitions()` is read-only.

## Triggers

Only the Product poller: `runProductNowAndPollRecurring()` creates an every-1-minute trigger on `pollProductRunAndWrite` (self-deletes on finish, failure, or after 180 min; one OOM retry at 4096 MB).

## Script Properties required

| Key | Used by |
|-----|---------|
| `APIFY_TOKEN` | Product task |
| `MONDAY_API_KEY` | Uploader |
| `GEMINI_API_KEY` | `=DR()` |

---

## Recent changes (since 2026-08)

- **2026-08-10** — `keywordFallback_()` pre-Gemini shortcut removed from `DR()`; cache key → `DR_v28_`.
- **2026-08-10** — 2nd 인입사유 정의 tuning round in `DefectDefsPatch.js`: filled 9 blank 휴대폰케이스 정의, guards for 필름간섭 / 코팅벗겨짐 / MagSafe 부착어려움 / 필름 부착어려움. Measured 63/70 (90.0%) on a larger sample (net −1 vs. before).
- **2026-08-06** — First 정의 patch: DR() 일치율 57.4% → 97.9% on 47 tagged rows.

## Known issues

- `CHAT_WEBHOOK_URL` is not defined in the repo → Product "completed" Chat post fails silently (caught).
- `startApifyRunAndSchedulePoll()` needs `CONFIG.actorTaskIdOrSlug`, which doesn't exist — dead code.
- `uiRunSummarize`, `uiRunDefectGPT`, `translateTextAuto` are undefined; 자동번역 column is left blank.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/Apify/GlxZ8_Apify
clasp push --force
```
