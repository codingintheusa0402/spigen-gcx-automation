# MasterTrigger

Google Apps Script project that runs the **daily Amazon review pipeline** (`masterDailyJob`): it starts every Apify/Axesso review-scraping task, writes each run into a dated tab of the source spreadsheet, then distributes the new reviews into the per-product monitoring spreadsheets (`1-5점` / `1-3점`), adding Gemini 키워드 summaries and 인입사유(AI) classifications.

**Script ID:** `1AWrX0Xl8feD-AzYRbGVBb9kLRQra2ppE547i_Ghys4lLU9l28pkMUf9O` (bound to the source spreadsheet)
**Source spreadsheet (`SRC_ID`):** `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`

> This is the canonical `clasp` folder for the master daily job. Always push from here — never from `Apify/APIFY_Axesso/`, which is an older copy with the **same scriptId** and would overwrite this project.

## Screenshots

![GlxZ8 `1-3점` tab after the daily job: 키워드 (AI 요약) and 인입사유(AI) filled per review](docs/ai_columns.jpg)
*GlxZ8 `1-3점` tab after the daily job: 키워드 (AI 요약) and 인입사유(AI) filled per review*

---

## Files

| File | Purpose |
|------|---------|
| `trigger.js` | `masterDailyJob()` entry point, `createTriggers()`, working-day/holiday check, trigger-expiry status card (`sendAllTriggerStatus`) |
| `Apify.js` | `APIFY_TASKS`, `runAllScrapers()`, 1-min `pollApifyRuns()` poller, dated-sheet writer, repair/retry helpers |
| `Master.js` | `SHEET_CONFIGS`, `dailyJob()` distribution, Gemini summary/classifier helpers, `fixLegacyDRFormulas()` |
| `Sheet_Automation.js` | Legacy helpers (`dedupeSheetByReviewId_`, standalone `tem` updater) — not called by the daily flow |
| `appsscript.json` | Manifest: `Asia/Seoul`, advanced Sheets v4 service, Calendar read-only scope (holiday check) |

---

## How it works

```
masterDailyJob()                         (time trigger, 04:00 KST working days)
  ├─ skip if weekend / Korean public holiday (Google holiday calendar, fail-open)
  ├─ retryPendingMaterializations_()     retry yesterday's failed sheet writes (max 3 attempts)
  ├─ runAllScrapers()                    POST every APIFY_TASKS task, save run state, add 1-min poller
  └─ sendAllTriggerStatus()              "Trigger Timeline" countdown bars → GCX Chat room

pollApifyRuns()                          (every 1 min until all runs finish)
  ├─ SUCCEEDED → createResultSheet_()    <Prefix>_yymmdd tab in SRC; drops Axesso *_PENALTY_n rows,
  │                                      header = union of keys, arrays → comma-joined, 1000-row chunks + 4 retries
  ├─ write failure → saved to APIFY_PENDING_MATERIALIZE (retried next day)
  └─ all finished → remove poller → dailyJob()

dailyJob()
  ├─ step1_deleteNumberedSheets()        delete *_yymmdd_n, "conflict", and non-today dated tabs
  ├─ step2_dedupDatedSheets()            dedupe today's dated tabs by Review ID
  ├─ step2b_updateTemSheet()             refresh `tem` tab with each dest's existing Review IDs
  └─ per SHEET_CONFIGS entry
        has15 = true  → _processFilterSheet_()  → dest 1-5점 (+ 키워드 (AI 요약), Update 날짜),
                                                   then 인입사유(AI) values on today's 1-3점 rows
        has15 = false → _processTo13_()         → dest 1-3점 only (optional insertAtTop / ratingFilter)
```

Each config reads the `finalize` filter view on `<X>_filter` in the source sheet (date cutoff + hidden-value exclusions), filters by `countries` / `seriesFilter`, and dedupes against the dest `Review ID` column plus the `tem` column.

The AI columns are written as **plain values** via direct Gemini calls (`gemini-3.1-flash-lite` → `gemini-2.5-flash-lite`, categories from the dest workbook's `Defect` tab) — not `=dr()` formulas. `fixLegacyDRFormulas()` converts leftover `=dr()` formulas to values in the old Glx26/iPh17e/Pixel10a/유지훈P books.

---

## APIFY_TASKS (current)

| Key | Apify task ID | Sheet prefix |
|---|---|---|
| SDA | `gIrI2jQcTdSG87VOu` | `SDA` |
| PowerAcc | `TgpoNoMcN4a5bYsyX` | `Power_Acc` |
| AutoAcc | `1jctsYj5oMnIkssv2` | `Auto_Acc` |
| 전략폰 | `Vv859ksggWODaIzoN` | `전략폰` |
| 유지훈P | `cskwDlRo3TY9TLsiQ` | `유지훈P` |
| GlxZ8 | `0imX0G3cKe75WJwXX` | `GlxZ8` |
| Pixel11 | `Oe0fI8QvYl7fSRhCu` | `Pixel11` |

Retired (commented out): Pixel10a, Glx26, iPh17e.

## SHEET_CONFIGS (current)

| filterSheet | Dest spreadsheet ID | has15 | Countries | Notes |
|---|---|---|---|---|
| `GlxZ8_filter` | `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4` | true | US FR ES JP UK IN DE IT | numCols 11, drFormula, pasteReviewId |
| `Pixel11_filter` | `12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI` | true | US FR ES JP UK IN DE IT | numCols 11, drFormula, pasteReviewId (added 2026-08-10) |
| `SDA_filter` | `1sxapIqJgXcJdeqyCf9bAxCNXrVMsVjsZE9QWPwEm0R4` | false | FR ES JP UK DE IT | |
| `Auto_Acc_filter` | `1mEYb1b92D6BIOaSYkAnMit6THuw5ewtymhA-mSIVDfs` | false | FR ES UK DE IT | |
| `Power_Acc_filter` | `1QC8Is6UvTnFXaOeXviKM_331i3Fo_CBIYx80VS696LI` | false | FR ES UK DE IT IN | |
| `전략폰_filter` | `1yo8CbLhJkuxrf3eXbAqZCb6qBejZhSR3YOt7nFv97fw` | false | IN | |
| `유지훈P_filter` | `1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw` | false | all 8 | insertAtTop, ratingFilter 1–3, drFormula |

Retired 2026-07-28 (commented out): `Glx26_filter`, `iPh17e_filter`, `Pixel10a_filter`. Field reference is in the comment block above `SHEET_CONFIGS` in `Master.js`.

---

## Entry points

| Function | Use |
|---|---|
| `masterDailyJob()` | Daily trigger target |
| `createTriggers()` | **Deletes all project triggers**, then creates one-off 04:00 KST `masterDailyJob` triggers for every working day until `MASTER_END_DATE` |
| `runAllScrapers()` / menu **Apify → Run ALL Scrapers** | Start scraping manually |
| `dailyJob()`, `dailyJob_GlxZ8()`, `dailyJob_Pixel11()`, `dailyJob_SDA()` … | Distribution only (all / single product) |
| `repairGlxZ8Sheet()`, `repair유지훈PSheet()`, `repairLatestRun(key)` | Rebuild a dated tab from the task's last run |
| `diagnoseAndClearStuckApifyState()`, `clearRunDoneProperties()` | Clear stuck poll state / `APIFY_RUN_DONE_*` property pile-up |
| `testSendTriggerStatus()` | Send the trigger-status card to `STATUS_WEBHOOK` |

## Script Properties

`APIFY_TOKEN`, `GEMINI_API_KEY`. Runtime state: `APIFY_MULTI_RUN_STATE`, `APIFY_RUN_DONE_<runId>`, `APIFY_PENDING_MATERIALIZE`.
Chat webhooks (`STATUS_WEBHOOK`, `GCX_WEBHOOK`) are constants in `trigger.js`.

## Known issues

- **Trigger expiry:** `MASTER_END_DATE` in the tracked code is `2026-08-16` (also the `endDate` shown for Apify Master / 오전보고 / TCT시트 보고 in `TRIGGER_PROJECTS`). `createTriggers()` only schedules one-off triggers up to that date, so unless it was re-run with a later date in the GAS editor, no `masterDailyJob` triggers exist after 2026-08-16. Check the live project's triggers and bump + re-run `createTriggers()` as needed (then `clasp pull` to keep git in sync).
- `fetchLatestGlx26Run()` still references `APIFY_TASKS.Glx26`, which is commented out — it throws if run.
- Legacy `Sheet_Automation.js` tem map points `유지훈P` at the Auto_Acc spreadsheet ID; unused by the daily flow (step2b uses `SHEET_CONFIGS`).

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_ReviewAutomation/MasterTrigger
clasp pull            # if anything was edited in the GAS editor
clasp push --force

cd ~/Desktop/GCX
git add GAS_ReviewAutomation/MasterTrigger/
git commit -m "feat(master-trigger): ..."
git push
```
