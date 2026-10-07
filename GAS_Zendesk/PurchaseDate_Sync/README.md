# PurchaseDate_Sync — Zendesk → monday.com "Purchase Date" bridge

Fills the **Purchase Date** date column on the Case+CP monday boards
[Galaxy Z8 (18421346787)](https://spigen.monday.com/boards/18421346787),
[Pixel 11 (18425190666)](https://spigen.monday.com/boards/18425190666), and
[iPhone 18 Series (18430082360)](https://spigen.monday.com/boards/18430082360)
from the Zendesk custom ticket field **Purchase Date** (field id `360019586172`).
Boards are listed in `MONDAY_BOARD_IDS` in `Code.js` — they share the same
column ids, so adding a new series board is a one-line change.

**Script ID:** `1f5Ky77H9LT0-9R6ZIuCsoWuZwXqJHGcQbs_RvV4kRbVBKYnMghLnilL1` (standalone; not bound to any sheet)
**Triggers page:** <https://script.google.com/home/projects/1f5Ky77H9LT0-9R6ZIuCsoWuZwXqJHGcQbs_RvV4kRbVBKYnMghLnilL1/triggers>

## Screenshots

![Twice-daily `scheduledPurchaseDateSync` executions](docs/executions.jpg)
*Twice-daily `scheduledPurchaseDateSync` executions*

## Why

The native monday↔Zendesk integration cannot map Zendesk **custom** fields to
board columns — for a date column its dropdown only offers Zendesk's system
fields (`Created at` / `Due at` / `Updated at`). So the Purchase Date entered
by agents on the ticket never reaches the board.

## How  (rewritten 2026-09-08 — scheduled batch; switched to twice-daily 2026-09-15)

Previously a Zendesk webhook hit `doPost` on every "Purchase Date changed"
event. In practice it fired ~5–6×/min around the clock and each call walked
the **entire board up to 4×** to locate one item — by far the biggest
consumer of the monday.com account's API budget (~tens of thousands of calls
/day). It is now a scheduled batch:

```
Time trigger twice a day, at each hour in SYNC_HOURS_KST (default 9am/9pm KST)
  └─ scheduledPurchaseDateSync()   (script-lock guarded; overlapping tick = no-op)
       └─ _runPurchaseDateSync_()
            ├─ ONE walk of each board in MONDAY_BOARD_IDS (500 items/page)
            │  → every item with a linked Zendesk ticket + its current
            │  Purchase Date cell
            ├─ Zendesk tickets/show_many (100 ids/call) → Purchase Date per ticket
            └─ change_multiple_column_values → date_mm59ejfp
               ONLY for items where the ticket's date differs from the board
```

`MONDAY_CALLS_MAX_PER_RUN = 200` is a hard ceiling — `mondayGql_` throws once a
single run passes it (200 × 2 runs/day = 400/day absolute worst case, still
far under the old ~40,000/day webhook). A run that hits the cap stops cleanly
and the next run picks up the remainder, so nothing is lost; a genuine
loop/bug fails loudly instead of repeating the 2026-09 runaway.

`doPost` is now a **no-op** — deactivate the Zendesk trigger + webhook.

## Components

| Piece | Where |
|---|---|
| GAS project | `PurchaseDate_Sync` (standalone) |
| Time trigger | twice daily at `SYNC_HOURS_KST` (default 9am/9pm KST) `scheduledPurchaseDateSync` — installed by `setupPurchaseDateTriggers` |
| Zendesk webhook | "monday Purchase Date Sync" → **deactivate** (endpoint is a no-op now) |
| Zendesk trigger | "monday Purchase Date Sync" → **deactivate** |
| monday boards | `18421346787` (Z8), `18425190666` (Pixel 11, added 2026-09-11), `18430082360` (iPhone 18 Series, added 2026-09-15) — columns `date_mm59ejfp` (Purchase Date), `integration_mm0fzmv0` (Zendesk Ticket) on all |

## Configuration (constants in `Code.js`)

There are no Script Properties. Everything is a constant at the top of `Code.js`:

| Constant | Value / meaning |
|---|---|
| `MONDAY_BOARD_IDS` | Boards to walk. **Add every new series Case+CP board here.** Pixel 11 and iPhone 18 were both missed at first. |
| `MONDAY_DATE_COL` / `MONDAY_TICKET_COL` | `date_mm59ejfp` / `integration_mm0fzmv0` (must be the same on every board) |
| `ZD_PURCHASE_DATE_FIELD` | `360019586172` |
| `SYNC_HOURS_KST` | `[9, 21]`. Re-run `setupPurchaseDateTriggers` after changing it. |
| `MONDAY_CALLS_MAX_PER_RUN` | `200` |
| `BOARD_PAGE_LIMIT` / `ZD_SHOW_MANY_CHUNK` / `WRITE_SLEEP_MS` | `500` / `100` / `120` ms |
| `ZENDESK_EMAIL`, `ZENDESK_TOKEN`, `MONDAY_TOKEN`, `WEBHOOK_SECRET` | Credentials. See [Secrets](#secrets). |

Each run logs one summary line (`updated / unchanged / ticket has no purchase date / failed`, plus the monday call count). To check what happened, look at the Executions page.

## Functions (GAS editor → Run)

- `setupPurchaseDateTriggers` — **run once** to install the twice-daily
  trigger(s) (needs the OAuth consent click; 403s if invoked headlessly).
  Re-run-safe; also clears any stale `backfillPurchaseDates` timer.
- `scheduledPurchaseDateSync` — the batch sync itself; also runnable on demand.
- `backfillPurchaseDates` — alias for `scheduledPurchaseDateSync` (kept for
  older docs). Now also refreshes items whose date changed, not just empty ones.
- `testSyncOneTicket` — report what the batch would do for one hardcoded ticket.

## Deploy

```bash
cd ~/Desktop/GCX/GAS_Zendesk/PurchaseDate_Sync
clasp push --force
```

`clasp push` is enough here. The live path is the time trigger, which always runs `@HEAD`, so **no `clasp deploy` is needed**. The old web-app deployment is pinned but only serves the no-op `doPost`. `clasp run` fails (no linked standard GCP project), so run functions from the script.google.com editor.

## History

- **2026-09-15:** iPhone 18 Series board added. Schedule moved from every 15 min to twice daily at 9:00/21:00 KST. Per-run cap raised 50 → 200.
- **2026-09-11:** Pixel 11 board added; `MONDAY_BOARD_ID` became the `MONDAY_BOARD_IDS` list.
- **2026-09-08:** the per-ticket webhook was replaced with a single-walk batch. Added the per-run monday call cap and the script lock.
- **2026-08-20:** a leftover every-minute `backfillPurchaseDates` trigger was found and removed (`setupPurchaseDateTriggers` now clears it).

## Cutover steps (one-time, from the 2026-09-08 rewrite)

1. `clasp push` (or paste `Code.js` into the editor).
2. Run `setupPurchaseDateTriggers` once from the editor; approve OAuth.
3. Run `scheduledPurchaseDateSync` once manually to confirm it completes clean.
4. In Zendesk Admin: deactivate the **trigger** and the **webhook** named
   "monday Purchase Date Sync".
5. (Optional) Apps Script → Deploy → Manage deployments → archive the Web App
   deployment. Leaving it published is harmless — `doPost` does nothing.

## Secrets

Zendesk API token, monday API token, and the (now-unused) webhook secret live
in `Code.js` constants. These were exposed in an assistant chat — **rotate the
Zendesk and monday tokens** and replace the constants. Do not copy them into docs.
