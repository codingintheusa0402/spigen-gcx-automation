# TCTChatLog_GCX

Google Apps Script project bound to the Spigen GCX TCT (Third-party Channel Team) chat log
spreadsheet (`Lazada log` / `Shopee log` tabs). Sends escalation alerts to Google Chat when a
row status changes to `Esc T2`, sends a weekday close report, syncs `Listing Issue` rows to a
monday.com board, and manages the voucher dropdown on the log sheets.

**Script ID:** `15E1ZJabPc7bQ4aKptxmPfCzOXryCDibhtYdZpKb6x1R-U6G_e8eHeqBm`
**Spreadsheet:** `1HZ14uqTVeP7bGYZDu9v9Ve2C1xNY_m6dcSv-KMCoAKc` (container-bound; tabs `Lazada log`, `Shopee log`)
**monday board:** `18409753446`, group `group_mm3jz6yk`

## Screenshots

![`Lazada log` tab (IDs, agents and customer columns blurred)](docs/lazada_log.jpg)
*`Lazada log` tab (IDs, agents and customer columns blurred)*

---

## Files

| File | Purpose |
|------|---------|
| `chat_alrt_main.js` | `onEdit` trigger — detects `Esc T2` status in col A and posts a Google Chat card to the platform-specific webhook |
| `send_daily_stat.js` | `sendDailyEscT2()` — counts `Esc T2` rows per sheet and sends the `TCT Chat Log_GCX 마감보고` cardsV2 card |
| `trigger.js` | Time-based trigger scheduling for `sendDailyEscT2` (self-perpetuating auto-extend) |
| `Monday_sync.js` | `syncListingIssuesToMonday()` — pushes `Listing Issue` rows to monday.com |
| `auto_suggestion_Q_col.js` | Sets dropdown validation on `Lazada log!Q2:Q9000` (`Provide 100% / 50% / 10% voucher`) |
| `appsscript.json` | GAS manifest (`Asia/Seoul`, V8) |

---

## Escalation alert flow (`chat_alrt_main.js`)

1. `onEdit` fires on any edit in **col A**, rows 5+.
2. If col A is `Esc T2` and the key `<sheet>_<orderId>_<platform>_<status>` isn't already in
   the `SENT_IDS` Script Property:
   - Collects all `Esc T2` rows on the same sheet (`getEscT2Rows`)
   - Posts a legacy `cards` payload: `@Lim TCT <sheet> Esc T2 Created:` + one
     `Open Claim on <row>` button per row (deep link to `A<row>`)
3. Routes by sheet name + platform value in **col E**:
   - `Lazada log` + `Lazada…` → `WEBHOOK_LAZADA`
   - `Shopee log` + `Shopee…` → `WEBHOOK_SHOPEE`
   (both currently point at the same TCT room; the commented-out lines point at the private test room)
4. A row is only recorded in `SENT_IDS` if a message actually went out — a sheet/platform
   mismatch no longer silently swallows the row. `resetSentAlerts()` clears the dedupe list.

Current header layout (row 4): `A Status | B Ticket ID | C Order ID | D Date | E Platform | …`.
Column D used to be Platform until a `Date` column was inserted before it — if alerts stop firing
again, re-check this layout first with `sheet.getRange(4, 1, 1, 16).getValues()`.

`testSendEscT2ToTestWebhook()` sends the current `Esc T2` rows of both tabs to `WEBHOOK_TEST`.

---

## Daily close report (`send_daily_stat.js`)

`sendDailyEscT2()` posts a cardsV2 card (`cardId: tct-daily-summary`, title
`TCT Chat Log_GCX 마감보고`) with `날짜: yyyy-MM-dd`, the `Esc T2` count per tab, and a
`시트 링크` button per tab. It is sent to **`WEBHOOK_TEST`** — despite the name, that constant
is the live destination for this report (the T2 ticket room), not a sandbox.

Schedule: weekdays, **Thu 15:30 / other weekdays 17:30 KST** (see Deployment).

---

## Listing Issue → monday.com (`Monday_sync.js`)

`syncListingIssuesToMonday()` scans both tabs from row 5, keeps rows whose Category column is
`Listing Issue`, and creates one monday item per new `OrderID_SKU` key on board `18409753446`
(item name `[<Platform>] <SKU>`, status `오류 접수`; SKU / Platform / Device / Model / Errors /
Order number / Product group / Date columns mapped in `CFG.MON`). Processed keys are kept in the
`processedKeys` Script Property; `resetProcessedKeys()` clears it.

> Work in progress (uncommitted at the time of writing): the column map is being shifted one
> column right for the 7/7 B-column insert, the monday token moved to the `MONDAY_TOKEN` Script
> Property, the sheet opened by ID, a weekday 08:00–17:30 KST run window added, and
> `setupSyncTrigger()` installing a 30-minute trigger. Check `git diff` before relying on the
> column indexes above.

---

## Script Properties

| Key | Used by |
|-----|---------|
| `SENT_IDS` | `Esc T2` alert dedupe list (JSON array) |
| `ESCT2_TRIGGER_SCHEDULE_END` | end of the currently-scheduled `sendDailyEscT2` trigger window |
| `processedKeys` | monday sync dedupe map (JSON object) |
| `MONDAY_TOKEN` | monday API token (once the WIP Monday_sync change lands) |

Webhook URLs (`WEBHOOK_LAZADA`, `WEBHOOK_SHOPEE`, `WEBHOOK_TEST`) are constants at the top of
`chat_alrt_main.js`, format `https://chat.googleapis.com/v1/spaces/<SPACE_ID>/messages?key=<KEY>&token=<TOKEN>`.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/TCTChatLog_GCX
clasp push --force
```

Run `setupAutoExtendEscT2Trigger()` once in the GAS editor. It installs a
daily 6AM self-check (`autoExtendEscT2Triggers`) that rolls the
`sendDailyEscT2` schedule (weekday, Thu 15:30 / other weekdays 17:30 KST)
forward 30 days whenever fewer than 3 days remain — full wipe + rebuild each
time, so no duplicate triggers accumulate. Runs forever until the installed
trigger is manually removed. (`createEscT2Triggers()` still exists for a
one-off manual regen with a hardcoded end date.)

`onEdit` is a simple trigger (no install needed). Run `setVoucherDropdown_LazadaLog()` manually
if the col-Q dropdown needs re-applying.

Related: replies to TCT-log tickets in the Ticket Reporter Chat app write back into this sheet
(Voucher / Memo / GCX STATUS / Status → `Esc T1`) — see `../TicketReporterCard/README.md`.
