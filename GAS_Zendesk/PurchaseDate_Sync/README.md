# PurchaseDate_Sync — Zendesk → monday.com "Purchase Date" bridge

Fills the **Purchase Date** date column on the Case+CP monday boards
[Galaxy Z8 (18421346787)](https://spigen.monday.com/boards/18421346787) and
[Pixel 11 (18425190666)](https://spigen.monday.com/boards/18425190666)
from the Zendesk custom ticket field **Purchase Date** (field id `360019586172`).
Boards are listed in `MONDAY_BOARD_IDS` in `Code.js` — they share the same
column ids, so adding a new series board is a one-line change.

## Why

The native monday↔Zendesk integration cannot map Zendesk **custom** fields to
board columns — for a date column its dropdown only offers Zendesk's system
fields (`Created at` / `Due at` / `Updated at`). So the Purchase Date entered
by agents on the ticket never reaches the board.

## How  (rewritten 2026-09-08 — scheduled batch)

Previously a Zendesk webhook hit `doPost` on every "Purchase Date changed"
event. In practice it fired ~5–6×/min around the clock and each call walked
the **entire board up to 4×** to locate one item — by far the biggest
consumer of the monday.com account's API budget (~tens of thousands of calls
/day). It is now a scheduled batch:

```
Time trigger every SYNC_EVERY_MINUTES min (15 → ~96 runs/day)
  └─ scheduledPurchaseDateSync()   (script-lock guarded; overlapping tick = no-op)
       └─ _runPurchaseDateSync_()
            ├─ ONE walk of each board in MONDAY_BOARD_IDS (500 items/page)
            │  → every item with a linked Zendesk ticket + its current
            │  Purchase Date cell
            ├─ Zendesk tickets/show_many (100 ids/call) → Purchase Date per ticket
            └─ change_multiple_column_values → date_mm59ejfp
               ONLY for items where the ticket's date differs from the board
```

The Z8 board holds ~1,040 items (**3 pages**) and the Pixel 11 board ~300
(**1 page**), so a run costs **~4 monday reads + only-changed writes** →
**~400–500 monday calls/day** across 96 runs (was ~40,000/day via the webhook).

`MONDAY_CALLS_MAX_PER_RUN = 50` is a hard ceiling — `mondayGql_` throws once a
single run passes it (50 × 96 = 4,800/day absolute worst case). A run that hits
the cap stops cleanly and the next run picks up the remainder, so nothing is
lost; a genuine loop/bug fails loudly instead of repeating the 2026-09 runaway.

`doPost` is now a **no-op** — deactivate the Zendesk trigger + webhook.

## Components

| Piece | Where |
|---|---|
| GAS project | `PurchaseDate_Sync` (standalone) |
| Time trigger | every 15 min `scheduledPurchaseDateSync` (~96/day) — installed by `setupPurchaseDateTriggers` |
| Zendesk webhook | "monday Purchase Date Sync" → **deactivate** (endpoint is a no-op now) |
| Zendesk trigger | "monday Purchase Date Sync" → **deactivate** |
| monday boards | `18421346787` (Z8), `18425190666` (Pixel 11, added 2026-09-11) — columns `date_mm59ejfp` (Purchase Date), `integration_mm0fzmv0` (Zendesk Ticket) on both |

## Functions (GAS editor → Run)

- `setupPurchaseDateTriggers` — **run once** to install the every-15-min trigger
  (needs the OAuth consent click; 403s if invoked headlessly). Re-run-safe;
  also clears any stale `backfillPurchaseDates` timer.
- `scheduledPurchaseDateSync` — the batch sync itself; also runnable on demand.
- `backfillPurchaseDates` — alias for `scheduledPurchaseDateSync` (kept for
  older docs). Now also refreshes items whose date changed, not just empty ones.
- `testSyncOneTicket` — report what the batch would do for one hardcoded ticket.

## Cutover steps

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
