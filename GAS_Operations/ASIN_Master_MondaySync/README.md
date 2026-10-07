# ASIN_Master_MondaySync

Container-bound Google Apps Script for the **ASIN_Master(먼데이보드)** spreadsheet (`1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo`). Two independent features live in one `Code.js`:

1. **Monday.com board sync** — pulls Monday board `7606389164` into the sheet (the `Data` tab is the SKU/product lookup that [GCX Reply](../../GAS_Zendesk/GCXReply_GAS/) reads).
2. **ABM_Relay_Log cleanup** — prunes old rows from the `ABM_Relay_Log` tab in the same spreadsheet.

**Script ID:** `1WwwnwKuPbpdTGG6Uozx1Mr0AZU9mD_hc3UzHq82e_t-GjRe2eQ_4K-Fp` (editor title "ASIN_Master (먼데이보드)")

## Screenshots

![ASIN_Master `Data` tab synced from the monday board](docs/data_tab.jpg)
*ASIN_Master `Data` tab synced from the monday board*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Both features below |
| `appsscript.json` | GAS manifest (V8, Asia/Seoul) |

---

## 1. Monday.com sync

Same modeless-dialog sync engine used by [Monday_CX_Board](../Monday_CX_Board/). `syncMondayBoardToSheet(reqId)`:

1. Fetches board columns via GraphQL.
2. Pass 1 — all items (`items_page` cursor, `PAGE_LIMIT` = 500 per page) with typed display fragments (`MirrorValue.display_value`, `StatusValue.label`, …).
3. Pass 2 — re-fetches formula columns (often empty on pass 1).
4. Pass 3 — targeted per-item `items(ids:[…])` fetch for any still-empty formula cells, chunked.
5. Writes `item_id, name, <columns…>` to the **active sheet**, preserving header formatting (`RESPECT_SHEET_FORMATS = true`). Zendesk integration values are normalised to a clean ticket URL.

| Key | Value |
|-----|-------|
| `BOARD_ID` | `7606389164` |
| `PAGE_LIMIT` | `500` (page size) |
| `RUN_LOCK_KEY` | `MONDAY_SYNC_LOCK` (Script Property used as a run lock) |
| `RUN_LOCK_STALE_MS` | 20 min |

> ⚠️ The target is `getActiveSheet()`, not a fixed tab name — when running manually, open the `Data` tab first.

### Usage / schedule

- Manual: open the spreadsheet → **Monday.com → 업데이트하기** (live-log modeless dialog, progress via `CacheService.getUserCache()` polled every 600 ms).
- Daily: a time-driven trigger on `syncMondayBoardToSheet` (created in the Apps Script UI, not in code).

### Run-lock self-heal (2026-08-12)

GAS hard-kills executions past its ~6 min limit and skips `finally`, which used to leave `MONDAY_SYNC_LOCK` set forever — every later run failed with *"Another Monday sync is already running."* (stuck for ~a month). `_acquireRunLock_` now treats a lock older than `RUN_LOCK_STALE_MS` (20 min) as orphaned and proceeds. The same fix was ported to [Monday_CX_Board](../Monday_CX_Board/) on 2026-09-17.

---

## 2. ABM_Relay_Log cleanup

Fully independent of the sync above — operates on a different tab (`ABM_Relay_Log`) written by [GCXReply_GAS](../../GAS_Zendesk/GCXReply_GAS/)'s `upsertAbmRelayLog_()` (columns `Timestamp, RelayKey, TicketId, CommentId, CaseId, Marketplace, Status, Attempts, LastError, MessageText`). That log only grows (one row per ABM relay, upserted in place, never pruned), and every write does a full-sheet scan to find a matching row — so an unbounded log slows down ABM reply relaying for every agent, not just sheet size. This prunes rows older than the retention window.

| Key | Value |
|-----|-------|
| `ABM_LOG_SHEET_NAME_` | `ABM_Relay_Log` |
| `ABM_LOG_RETENTION_DAYS_` | `15` |

### Functions

| Function | Description |
|----------|-------------|
| `cleanupOldAbmRelayLogRows()` | Deletes rows whose Timestamp (col A) is older than 15 days. Deletes contiguous runs highest-row-first so row numbers of still-queued deletions never shift. |
| `installAbmRelayLogCleanupTrigger()` | One-time setup — run manually from the Apps Script editor (not headlessly; trigger creation needs an OAuth consent click). Installs a daily trigger at ~4am Asia/Seoul. Safe to re-run (no-ops if already installed). |

---

## Triggers (live)

| Handler | Schedule | Created by |
|---------|----------|-----------|
| `syncMondayBoardToSheet` | daily | Apps Script UI |
| `cleanupOldAbmRelayLogRows` | daily ~04:00 KST | `installAbmRelayLogCleanupTrigger()` |

---

## Script Properties

| Key | Used by |
|-----|---------|
| `MONDAY_API_KEY` | Monday.com sync (or `MONDAY_API_KEY_HARDCODED` in `Code.js`, normally empty) |
| `MONDAY_SYNC_LOCK` | Written/cleared automatically as the run lock |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/ASIN_Master_MondaySync
clasp push --force
```

Container-bound + time-driven trigger only — no web-app deployment to bump.

> ⚠️ `../TriggerAlert/` has a `.clasp.json` pointing at **this same script ID** with an older copy of the sync code. Never `clasp push` from `TriggerAlert/` — it would overwrite this project and drop the ABM cleanup + run-lock fix.
