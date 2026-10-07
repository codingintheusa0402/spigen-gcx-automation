# TriggerAlert

⚠️ **Legacy / do-not-push folder.** Despite the name, this project contains no trigger-alerting logic — `Code.js` is an **older copy of the Monday.com → Sheet sync** template (board `7606389164`), and its `.clasp.json` points at the **same script ID as [ASIN_Master_MondaySync](../ASIN_Master_MondaySync/)** (`1WwwnwKuPbpdTGG6Uozx1Mr0AZU9mD_hc3UzHq82e_t-GjRe2eQ_4K-Fp`). Running `clasp push` here would overwrite the live ASIN_Master project and remove its ABM_Relay_Log cleanup and its stale run-lock fix. Make changes in `ASIN_Master_MondaySync/` (or `Monday_CX_Board/`) instead.

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | `syncMondayBoardToSheet()`, modeless-dialog UI HTML, progress logger (pre-2026-08-12 version) |
| `appsscript.json` | GAS manifest |

---

## What the code does

Identical to the Monday-sync half of ASIN_Master_MondaySync as it was before 2026-08-12:

- `onOpen()` → **Monday.com → 업데이트하기** → `showMondaySyncDialog()` → `syncMondayBoardToSheet(reqId)`.
- 3-pass GraphQL fetch (items → formula columns → per-item fill-in), written to the **active sheet** with header formatting preserved.

| Constant | Value | Notes |
|----------|-------|-------|
| `BOARD_ID` | `7606389164` | Same board as ASIN_Master |
| `MONDAY_API_KEY_HARDCODED` | `''` | Otherwise Script Property `MONDAY_API_KEY` |
| `PAGE_LIMIT` | `500` | `items_page` page size (not a total cap) |
| `RESPECT_SHEET_FORMATS` | `true` | Preserve existing cell formatting on write |
| `RUN_LOCK_KEY` | `MONDAY_SYNC_LOCK` | ⚠️ No staleness check — a killed run leaves the lock set forever (the bug fixed in ASIN_Master_MondaySync / Monday_CX_Board) |

No ABM_Relay_Log cleanup, no `RUN_LOCK_STALE_MS`, no triggers in code.

---

## Deployment

Don't. If this folder is ever repurposed, first point `.clasp.json` at its own new script ID.
