# Monday_CX_Board

Container-bound Google Apps Script that syncs Monday.com board `5669388007` into the `Data` tab of the **SKU_Master(먼데이보드)** spreadsheet (`1JijzoYw9aDW-9Jx_OJOBByMVCX8-0OtNxb-0Y82EZJI`), with a modeless dialog UI (Monday.com branding + live log). Same sync engine as the Monday-sync half of [ASIN_Master_MondaySync](../ASIN_Master_MondaySync/), minus the ABM log cleanup feature.

**Script ID:** `14WU3CBKFHvoBAKXOPEemCLZZoUSfKTISUQLwullDbM_cIAdhkbBxpO0S` (editor title "SKU_Master(먼데이보드)")

## Screenshots

![SKU_Master `Data` tab synced from the monday board](docs/data_tab.jpg)
*SKU_Master `Data` tab synced from the monday board*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | Full sync engine + UI |
| `appsscript.json` | GAS manifest (V8, Asia/Seoul) |

---

## How it works

`syncMondayBoardToSheet(reqId)`:

1. Acquires the run lock (`MONDAY_SYNC_LOCK` Script Property; locks older than 20 min are treated as orphaned).
2. Fetches board columns via GraphQL.
3. Pass 1 — fetches all items with a safe fragment set (universal fields + typed extras like `MirrorValue.display_value`, `StatusValue.label`, etc.), paginated via `items_page` cursor (`PAGE_LIMIT` = 500 per page).
4. Pass 2 — re-fetches formula columns specifically (formula values often come back empty on pass 1).
5. Pass 3 — for any item still missing formula values, a targeted root-level `items(ids:[…])` fetch, chunked.
6. Writes `item_id, name, <columns…>` to the **`Data`** sheet (falls back to the active sheet if `Data` doesn't exist), preserving header row formatting.

Progress is streamed to the modeless dialog via `CacheService.getUserCache()`, polled every 600ms from the client-side dialog HTML. A full run takes ~6 min at current board size.

| Key | Value |
|-----|-------|
| `BOARD_ID` | `5669388007` |
| Target tab | `Data` |
| `RUN_LOCK_STALE_MS` | 20 min |

### Entry points

| Function | Description |
|----------|-------------|
| `onOpen()` | Adds **Monday.com → 업데이트하기** menu |
| `showMondaySyncDialog()` | Opens the live-log dialog and starts a sync |
| `syncMondayBoardToSheet(reqId)` | The sync itself (also the trigger handler) |
| `clearRunLock()` | Manually deletes `MONDAY_SYNC_LOCK` if a run is ever stuck |

### Schedule

Daily time-driven trigger on `syncMondayBoardToSheet` at ~05:01 KST (created in the Apps Script UI).

### Recent changes

- **2026-09-17** — every run since 2026-09-11 failed with *"Another Monday sync is already running."*: an execution killed past the GAS limit skipped `finally` and left the lock set. Added the stale-lock check (`RUN_LOCK_STALE_MS`) + `clearRunLock()`, same fix as ASIN_Master_MondaySync (2026-08-12). Other copies of this template (per-product boards) may still lack it.

---

## Script Properties

| Key | Used by |
|-----|---------|
| `MONDAY_API_KEY` | Board sync |
| `MONDAY_SYNC_LOCK` | Run lock (auto-managed) |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/Monday_CX_Board
clasp push --force
```
