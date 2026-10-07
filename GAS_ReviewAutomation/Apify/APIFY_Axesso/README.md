# APIFY_Axesso (legacy snapshot)

Older copy of the Google Apps Script project that scrapes Amazon reviews via Apify/Axesso tasks and distributes them into the Spigen product monitoring spreadsheets. **It is superseded by [MasterTrigger](../../MasterTrigger/)** and kept only for reference/history — the code here was last changed on 2026-07-27 and has drifted from what is deployed.

> ⚠️ **Do not `clasp push` from this folder.** Its `.clasp.json` has the **same scriptId as MasterTrigger** (`1AWrX0Xl8feD-AzYRbGVBb9kLRQra2ppE547i_Ghys4lLU9l28pkMUf9O`), so a push here would overwrite the live `masterDailyJob` project with this stale code (re-enabling retired products, removing the pending-write retry, and reverting the AI columns to `=dr()` formulas).

**Source spreadsheet:** `SRC_ID = 1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`

## Screenshots

![Output of the live pipeline (GlxZ8 `1-3점` tab, reviewer names blurred)](../GlxZ8_Apify/docs/sheet_1-3.jpg)
*Output of the live pipeline (GlxZ8 `1-3점` tab, reviewer names blurred)*

## Files

| File | Purpose |
|------|---------|
| `Master.js` | `dailyJob()` — filters, dedupes and distributes reviews into destination spreadsheets |
| `Apify.js` | `runAllScrapers()` + 1-min `pollApifyRuns()` → dated `<Prefix>_yymmdd` source tabs |
| `Sheet_Automation.js` | `dedupeSheetByReviewId_()` + legacy standalone `tem` updater |
| `trigger.js` | `masterDailyJob()`, `createTriggers()`, holiday-aware trigger-countdown Chat card |

## Flow (as in this snapshot)

```
masterDailyJob()            skip weekends/KR holidays → runAllScrapers() → sendAllTriggerStatus()
pollApifyRuns()             SUCCEEDED → createResultSheet_() ; all done → dailyJob()
dailyJob()
  ├─ step1_deleteNumberedSheets / step2_dedupDatedSheets / step2b_updateTemSheet
  └─ SHEET_CONFIGS loop: has15 → _processFilterSheet_()  |  !has15 → _processTo13_()
```

## How this snapshot differs from MasterTrigger (current)

| | APIFY_Axesso (here) | MasterTrigger |
|---|---|---|
| `APIFY_TASKS` | SDA, PowerAcc, AutoAcc, 전략폰, 유지훈P, **Pixel10a, Glx26, iPh17e** | SDA, PowerAcc, AutoAcc, 전략폰, 유지훈P, **GlxZ8, Pixel11** |
| `SHEET_CONFIGS` | Glx26, iPh17e, Pixel10a, SDA, Auto_Acc, Power_Acc, 전략폰, 유지훈P | GlxZ8, Pixel11, SDA, Auto_Acc, Power_Acc, 전략폰, 유지훈P |
| AI columns | `=dr()` formula in 1-3점 `인입사유(AI)`, formula in 1-5점 `키워드 (AI 요약)` | Static values from direct Gemini calls |
| `createTriggers()` time | 09:00 KST | 04:00 KST |
| `MASTER_END_DATE` | `2026-05-08` (expired) | `2026-08-16` |
| Failed-write retry (`APIFY_PENDING_MATERIALIZE`), penalty-row filter, chunked writes, repair helpers | — | yes |

`SHEET_CONFIGS` fields (`filterSheet`, `destId`, `countries`, `numCols`, `has15`, `seriesFilter`, `temCol`, `insertAtTop`, `ratingFilter`, `drFormula`, `pasteReviewId`) are documented in the comment block at the top of `Master.js` and are the same as MasterTrigger's.

## Script Properties

`APIFY_TOKEN`. Chat webhooks are constants in `trigger.js` (`STATUS_WEBHOOK`, `GCX_WEBHOOK`), not Script Properties.
