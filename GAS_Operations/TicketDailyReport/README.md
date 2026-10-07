# TicketDailyReport

Google Apps Script project that generates the Spigen GCX daily Zendesk ticket report (weekdays 9AM KST) — fetches ticket views from the Zendesk API, updates the `All_Graph` / `K_시트` sheets, posts the `All_Graph` chart image and the `K_시트` pending-ticket table card to the Ticket T2 Google Chat room, and serves a `/report` Chat app for browsing past days.

**Script ID:** `1GNowLPF82wfWIrHc1Q9xIDr_L0Kz5iv_YxuNRb7Pzioe9OpsS_O_J5L6`  
**Linked spreadsheet:** `10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo`

## Screenshots

![`All_Graph` tab: daily New / Open / Pending chart used in the report card](docs/all_graph.jpg)
*`All_Graph` tab: daily New / Open / Pending chart used in the report card*

---

## Files

| File | Purpose |
|------|---------|
| `main.js` | `runZendeskDailyJob()` — declares the ordered step list, runs it via `runResumableJob_()` |
| `retryRunner.js` | Resilient runner: per-step retry, resume-from-crash, auto catch-up scheduling |
| `ZendeskAPI.js` | Zendesk API helpers — `zendeskApiGet_()` (retrying GET), tickets-by-view, tagger option→name map, `fetchZendeskViewToKsheet()` |
| `All_GraphGen.js` | `appendZendeskDailyStatus()` (one `All_Graph` row/day), `collapseOldRowsIfNeeded()`, `getKoreanFormattedDate()` |
| `K_시트Gen.js` | `kSheetToChat()` — builds + posts the `K_시트` monospace-table cardsV2 card; shared row-cleanup / table helpers |
| `KSheetHistory.js` | `archiveKSheetHistory()` — appends today's `K_시트` rows to the `K_시트_history` sheet (source for the `/report` date picker) |
| `chatApp.js` | Google Chat app — `/report` slash command → `onMessage` shows an "Update 날짜" date-picker card; `loadDayReport` handles the button |
| `sendChat.js` | `all_GraphChartToGoogleChat()` — exports the first `All_Graph` chart as PNG, uploads it to freeimage.host (`uploadToFreeImageHost`), posts a `[오전 보고]` image card |
| `testingFunc.js` | Old manual test helpers (all commented out) |
| `trigger.js` | `setupAutoExtendZendeskTrigger()` / `autoExtendZendeskTriggers()` / legacy `createTriggers()` |
| `appsscript.json` | GAS manifest — **must contain `"chat": {}`** (see the `/report` section) |

---

## `/report` Chat app

Type `/report` in a DM with the **T2 Report** app (or a space it's in) to get a
card: pick an `Update 날짜` from the dropdown and see that day's `K_시트`
pending-ticket table, read from the `K_시트_history` sheet that
`archiveKSheetHistory()` appends to on every weekday run.

### Wiring (all in the `gcxbot` GCP project — `console.cloud.google.com`, project `gcxbot` / number `64325928759`)

1. **Google Chat API** enabled.
2. **`appsscript.json` must have `"chat": {}`.** Without it, *no* deployment of
   this project registers as a Chat app — Chat has nowhere to send the event,
   so `/report` returns **"T2 Report not responding"** with **zero** Apps Script
   executions and zero `chat_app` logs. This was the bug on 2026-09-08.
3. Deploy → **New deployment → type Add-on** → copy the **Deployment ID**
   (a *versioned* deployment, not Head).
4. Chat API → **Configuration**:
   - App status **LIVE**, name `T2 Report`, `/report` slash command (id 1),
     Interactive features on, "Join spaces" on.
   - **Connection settings → Apps Script → paste the Deployment ID from step 3.**
   - Visibility: the people/groups who should see it (currently `kjw@spigen.com`).
5. After any `chatApp.js` change: `clasp push`, make a **new** Add-on deployment,
   and update the Deployment ID in Connection settings (`clasp push` alone does
   **not** move the versioned deployment Chat points at). Allow ~1–3 min to
   propagate, then re-test `/report`.

`test_chatPicker` (editor → Run) logs the card JSON and triggers the OAuth
consent — run it once after granting any new scope.

---

## Daily job flow (`runZendeskDailyJob`)

`runZendeskDailyJob()` hands an ordered list of named steps to `runResumableJob_()`:

| # | Step (`name`) | Does | Critical? |
|---|---------------|------|-----------|
| 1 | `fetchZendeskViewToSheet`    | fills `Zendesk_Daily` from 9 Zendesk views | yes |
| 2 | `fetchZendeskViewToKsheet_A` | fills `K_시트` B5:In table | yes |
| 3 | `appendZendeskDailyStatus`   | counts new/open/pending → **one** `All_Graph` row/day, clears `Zendesk_Daily` | yes |
| 4 | `all_GraphChartToGoogleChat` | posts the `All_Graph` chart image to Google Chat | yes |
| 5 | `collapseOldRowsIfNeeded`    | collapses rows past 4 weeks | **no** (logged & skipped on failure) |
| 6 | `fetchZendeskViewToKsheet_B` | rebuilds `K_시트` for `P_시트` | yes |
| 7 | `archiveKSheetHistory`       | appends today's `K_시트` rows to `K_시트_history` (idempotent per date) | **no** |
| 8 | `kSheetToChat`               | posts the `K_시트` pending-ticket table card to Google Chat | yes |

Step 1 reads views `360102672972`, `360121546552`, `360121545992`, `360108214671`,
`28990066416793`, `360096790151`, `19940259463705`, `360103290632`, `37502606662809`.
The `K_시트` card shows at most 25 rows (note line if truncated), header title from `K_시트!B3`,
`All: <G3 or sum> | KJW: <n>` totals and a `Start Zendesk` button (filter `360103290632`).

### `K_시트` table columns (view `49523632520985`, `pending` tickets only)

| Col | Field | Source |
|-----|-------|--------|
| B | No. | sequence |
| C | Country | field `4513936822297` |
| D | Brand | field `5495572594201` |
| E | Category | field `900006613446` |
| F | Qty | group count |
| G | Owner | PIC — always `KJW` (LYS left the company 2026-09-10; was week-alternating LYS/KJW by country group) |
| H | Device | field `360022185671` (raw value → agent-UI name) |
| I | 1차 Defect Reason or Inquiries | field `360022182831` — shown on the Chat card as **인입사유** |

Rows are grouped by Country + Brand + Category + Device + Reason; `Qty` is the group size.

The Chat card renders these as one column-aligned `<pre>` monospace table
(`# | CC | Brand | Category | Device | 인입사유 | Qty | PIC`). Card-only cleanups
applied in `kSheetToChat()` (the sheet keeps its raw values):

- Brand: strip `spigen_` / `Spigen` / `(` / `)`, then upper-case → `NEW BIZ`, `SDA`
- Category: drop the leading `N. ` → `Product Inquiry`
- 인입사유: drop the leading `(XXX)_` prefix → `대량구매문의`

---

## Resilience (`retryRunner.js`)

Transient failures — most often `Exception: Address unavailable: https://spigenhelp.zendesk.com/api/v2/views/.../tickets.json` — hit this job a few times a month. Handling:

1. **Zendesk fetch retry** — `getZendeskTicketsByView()` retries each call up to `ZENDESK_FETCH_ATTEMPTS` (3), backoff 0/2s/4s, on network error, HTTP 429, HTTP 5xx, or unparseable JSON. Non-transient 4xx (auth / not found) fail fast.
2. **Per-step retry** — each step is retried up to `MAX_STEP_ATTEMPTS` (3), backoff 0/5s/15s.
3. **Resume from crash** — every finished step is recorded in Script Properties under `dailyJobProgress:<KST date>`. If the execution dies, the next run **skips the finished steps and resumes at the one that crashed** — so the chart / table are never re-posted and the `All_Graph` row is never added twice. Progress is wiped on full success and pruned when the date rolls over.
4. **Same-day guard** — `appendZendeskDailyStatus()` checks column A of the last `All_Graph` row. If it already equals today's label it **overwrites that row** instead of appending; if `Zendesk_Daily` is already cleared it leaves the row alone (no zero-out). This holds even if the saved progress is lost.
5. **Auto catch-up** — after retries are exhausted on a critical step, a one-off trigger re-runs `runZendeskDailyJob` in `RESCHEDULE_DELAY_MS` (5 min), up to `MAX_RESCHEDULES` (3) per day. These triggers are tracked by id and cleaned up on the next run / on success. Set `ENABLE_AUTO_RESCHEDULE = false` to rely on the normal daily trigger / a manual re-run instead.

**Manual recovery:** run `runZendeskDailyJob` again — it resumes where it stopped. To force a clean full run, use `runZendeskDailyJob_forceFresh()` (clears today's progress + catch-up triggers first).

---

## Configuration

Zendesk credentials and the spreadsheet ID are constants at the top of `ZendeskAPI.js`:

| Constant | Value |
|----------|-------|
| `ZENDESK_SUBDOMAIN` | `spigenhelp` |
| `ZENDESK_EMAIL` | `kjw@spigen.com` |
| `ZENDESK_TOKEN` | `<ZENDESK_API_TOKEN>` (set in file) |
| `SPREADSHEET_ID` | `10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo` |

- `webhookUrl` (top of `main.js`) — Ticket T2 room (`spaces/AAQAdqYt1ro`); a commented-out line
  points at the private test room `spaces/AAQAc9NQmJQ`. Swap the comment to test.
- freeimage.host upload key — inline in `uploadToFreeImageHost()` (`sendChat.js`).
- Script Properties (written by the code, nothing to set by hand): `dailyJobProgress:<yyyy-MM-dd>`,
  `dailyJobCatchupTriggerIds`, `ZENDESK_TRIGGER_SCHEDULE_END`.

Sheets touched (all in `SPREADSHEET_ID`): `Zendesk_Daily`, `All_Graph`, `K_시트`, `K_시트_history`.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/TicketDailyReport
clasp push --force
```

Run `setupAutoExtendZendeskTrigger()` once in the GAS editor. It installs a
daily 6AM self-check (`autoExtendZendeskTriggers`) that rolls the
`runZendeskDailyJob` schedule (weekday 9AM KST) forward 30 days whenever
fewer than 3 days remain — full wipe + rebuild each time, so no duplicate
triggers accumulate. Runs forever until the installed trigger is manually
removed. (`createTriggers()` still exists for a one-off manual regen with a
hardcoded end date.)
