# MCF_Tracking

Google Apps Script project for Spigen GCX — Multi-Channel Fulfillment (MCF) order tracking, fee
backfill, stock lookup and a daily Google Chat report, bound to the `MCF 발송 로그` spreadsheet. It
resolves Amazon tracking numbers via SP-API custom formulas (`=AMZTK()` / `=AMZTK_JP()`), writes
settled MCF fulfillment fees into col Y from Settlement Reports, keeps the `GStore MCF 발송 로그`
tab's tracking column filled, and serves the web-app endpoint the GCX Reply Tampermonkey script uses
to mark a row as MCF and read back its Order ID.

**GAS Script ID:** `1kDfEUVEEJ7TCA3HOMF6EYFTjbeZZKIeg_X84wCbLT1-tQqJI2ZlPUCxp` (bound)  
**Spreadsheet:** `1g6a-S7eeA1oY19aTEFhTNAyp2A5nLqLNkPRqOqriWfc` — tabs `MCF 발송 로그` (gid 1608794212,
header row 3, data from row 4), `GStore MCF 발송 로그` (data from row 5), hidden `_SettlementFeeCache`

## Screenshots

![`MCF 발송 로그` tab (customer names and tracking numbers blurred)](docs/mcf_log.jpg)
*`MCF 발송 로그` tab (customer names and tracking numbers blurred)*

---

## Active triggers (live script)

| Handler | Schedule | What it does |
|---|---|---|
| `backfillMCFFeesRecent()` | every 30 min | Col Y fees from the last ~89 days of Settlement Reports + `getFulfillmentPreview` estimate fallback |
| `retryR429Errors()` | hourly | Re-fetches col R cells showing 429, then calls `resolveBlankTrackingNumbers()` (GStore top-up + both sheets) |
| `MCFReporter()` | weekdays 9 AM KST | Chat card listing rows with a Tracking No. but empty col S |
| `onEdit_mcf(e)` | installable on-edit | Row automation (see `onEdit` section) |

Triggers were created from the Apps Script **Triggers** UI (there is no installer in the code except
`triggerTester()`); manage them there. `dailyTrackingMaintenance()` was removed on 2026-07-31.

## Files

| File | Purpose |
|------|---------|
| `sp-api.js` | SP-API auth (LWA + AWS SigV4), core fetch/retry, custom formulas (`AMZTK`, `AMZTK_JP`, `MCFFee`, `MCFFee_JP`, `MCFFeeDebug`, `MCFFeeDebugWide`, `getMcfStockByAsin`), Settlement Report fee map + `_SettlementFeeCache`, cache-clear utilities (`clearMcfFeeCache`, `clearAmztkCache`, `clearSettlementScanCache`, `pruneFeeCacheToGcxOnly`) |
| `autoFill.js` | `onEdit_mcf`, fee backfills (`backfillMCFFees`, `backfillMCFFeesRecent`, `fillPreviewEstimatesForRange`, `retryZeroTransportationFees`, `recoverAllZeroTransportationFees`, `auditYColumnFees`), tracking maintenance (`retryR429Errors`, `resolveBlankTrackingNumbers`, `installGStoreTrackingFormulas`, `unfreezeAmztkFormulas` / `freezeAmztkFormulas`). `backfillTrackingNumbers()`, `freezeTrackingColumnR()`, `dailyTrackingMaintenance()` are **disabled** stubs (see below) |
| `tamperMonkey.js` | Web app `doGet` used by GCX Reply's MCF autofill (see below) |
| `main.js` | `MCFReporter` — daily Google Chat card of rows missing a tracking entry |
| `MCFGen.js` | Archived — MCF order creation / stock-check helpers, **entirely commented out** |
| `triggerGen.js` | `triggerTester()` — schedules a one-off `MCFReporter` run 1 minute out |
| `_diag.js` | `diagCheckAug12PlusRows()` — read-only col Y diagnostic |
| `_oneTime*Backfill.js` | One-time col Y fee backfills from locally parsed settlement reports (May 2026, Jul–Aug, JP Oct–Dec, UK, UK July). Each says "run once, then delete" |
| `appsscript.json` | Manifest (`Asia/Seoul`, V8, scopes spreadsheets / external_request / scriptapp; web app `USER_DEPLOYING`, `ANYONE_ANONYMOUS`) |

---

## Custom Sheet Formulas (`sp-api.js`)

### `=AMZTK(orderId)`
Returns the tracking number for an EU MCF order. Tries EU endpoint first, falls back to FE.

### `=AMZTK_JP(orderId)`
Same as `AMZTK` but tries FE (Japan/AU/SG) first.

### `=MCFFee(orderId)` / `=MCFFee(orderId, sentDate)` / `=MCFFee(sentDate, orderId)`
Returns the actual settled MCF fulfillment fee via the Finances API. Always uses the actual charged amount (not an estimate). Returns blank until the order settles (usually a few days after shipment) — retries automatically on the next sheet recalculation.

**For bulk use, prefer `backfillMCFFees()` over this formula** — ARRAYFORMULA is not supported and many simultaneous formula calls will queue indefinitely.

| Usage | Example | Notes |
|-------|---------|-------|
| 1-arg | `=MCFFee(Q2)` | Searches last 180 days |
| 2-arg | `=MCFFee(Q2, P2)` or `=MCFFee(P2, Q2)` | `Q2` = orderId, `P2` = sent date (yyyy-mm-dd) — either argument order works, detected by shape (a `Date`/`yyyy-mm-dd` value is always treated as sentDate). Providing sentDate is faster, less bandwidth. |

```
=IF(Q2<>"", MCFFee(Q2, P2), "")
```

> Different row blocks in this sheet have historically used both argument orders — `MCFFee()`/`MCFFee_JP()` detect which argument is the date vs. the orderId by value shape rather than trusting position, so both work correctly regardless of which convention a given row's formula was copied from.

**Lookup strategy:** two strategies, all SP-API calls use 1 attempt (no retry sleep → no 30s timeout):

**Strategy A — targeted** (for Amazon marketplace–linked orders): `getFulfillmentOrderRaw` → `displayableOrderId` (Amazon 3-7-7 format) → `GET /finances/v0/orders/{id}/financialEvents` → tiny single-order response.

**Strategy B — date-range scan** (for seller-created MCF orders with GCX orderId): `listFinancialEvents` for `sentDate → sentDate+60d` window (2 pages max), match `SellerOrderId = GCX orderId`. Confirmed: `backfillMCFFees()` matches fees this way, so Finances API stores `SellerOrderId = GCX-XX-XXXXXX-XX`.

> **Why 1 attempt / 2 pages?** With 15+ concurrent formula cells, retry sleeps (5s × 3 = 15s) cause 30s GAS timeout. Instead: on 429, the cell returns `''` cached for 90 s and retries automatically with fewer concurrent peers. Events older than 2 pages: run `backfillMCFFees()` which uses 20 pages + full retries in a single sequential process.
>
> `=MCFFee(Q35)` (1-arg) and `=MCFFee(P35, Q35)` (2-arg) both work. sentDate (P col) narrows the Finances API window for Strategy B.

- On 429 QuotaExceeded: returns blank and retries after 90 seconds automatically.
- Tries EU endpoint first, falls back to FE.
- Required SP-API roles: **Amazon Fulfillment** + **Finance and Accounting**.

> **GCX order fee-type quirk:** GCX-prefixed orders' Finances API fee lines aren't tagged with a
> `FeeType` containing `FBA`/`FULFILLMENT` (unlike real Amazon marketplace orders), so the standard
> fee-type filter silently dropped their real fee data to `0`. `_sumMcfFeeFromShipments()` and
> `_buildFeeMapForWindow()` sum **all** fee-line types for any `SellerOrderId` starting with `GCX`
> — this was fixed once already (commits `5fe6def`/`eed853c`, 2026-04-24/27) but got lost in a
> later revert; restored 2026-07-28. Also fixed: only the *first* matching `ShipmentEvent` was
> ever checked — split-shipment orders whose fee posted in a later event were misread as `0`;
> now sums across every matching event.
>
> **Important — CacheService, not just code, has to be cleared:** `MCFFee()`/`MCFFee_JP()` cache
> their result in `CacheService` for up to 6h, independent of script deployments. A cell that
> cached a wrong `0` under the pre-fix code keeps serving that stale value for the rest of its
> TTL even after `clasp push` — the cache check happens before the (now-correct) computation
> logic ever runs. Run `clearMcfFeeCache()` once after any MCFFee-related code fix so cells
> actually recompute instead of replaying stale cached results.

### `=MCFFee_JP(orderId)` / `=MCFFee_JP(sentDate, orderId)`
Same as `MCFFee` but tries FE (Japan/AU/SG) first.

### `getMcfStockByAsin(asin, marketplaceId)`
Returns available FBA inventory count for a given ASIN and marketplace ID. Used internally by `autoFill.js`.

> **Cache TTL:** Found values (tracking number or fee) → 6 hours. Empty/not-yet-settled → 10 minutes (retried). 429 QuotaExceeded → 90 seconds (auto-retry). Permanent errors (403) → 6 hours.

---

## SP-API Setup (Script Properties)

Set these in **Extensions → Apps Script → Project Settings → Script Properties**:

| Key | Description |
|-----|-------------|
| `LWA_CLIENT_ID` | EU LWA client ID |
| `LWA_CLIENT_SECRET` | EU LWA client secret |
| `LWA_REFRESH_TOKEN` | EU LWA refresh token |
| `LWA_CLIENT_ID_JP` | JP LWA client ID (falls back to `LWA_CLIENT_ID`) |
| `LWA_CLIENT_SECRET_JP` | JP LWA client secret |
| `LWA_REFRESH_TOKEN_JP` | JP LWA refresh token |
| `AWS_ACCESS_KEY_ID` | AWS access key for SigV4 signing |
| `AWS_SECRET_ACCESS_KEY` | AWS secret key |
| `AWS_SESSION_TOKEN` | (Optional) STS session token for assumed roles |
| `SPAPI_HOST_EU` | Defaults to `sellingpartnerapi-eu.amazon.com` |
| `SPAPI_HOST_FE` | Defaults to `sellingpartnerapi-fe.amazon.com` |
| `SPAPI_REGION_EU` | Defaults to `eu-west-1` |
| `SPAPI_REGION_FE` | Defaults to `us-west-2` |
| `SETTLEMENT_SCANNED_<endpoint>` | Written by the script — settlement report IDs already scanned (reset with `clearSettlementScanCache()`) |

**Required SP-API roles:**
- `Amazon Fulfillment` — tracking lookup (`AMZTK`), stock lookup, `MCFFee` (both methods)
- `Finance and Accounting` — `MCFFee("FinancesAPI", ...)` only

---

## Backfill Functions (`autoFill.js`)

### `backfillMCFFees()`
Writes MCF fulfillment fees as **static values** into col Y (`Transportation Fee`). Use this instead of `=MCFFee(...)` formulas — bulk formula calls queue indefinitely and ARRAYFORMULA is not supported for API-calling functions.

**Batch fetch approach** (fast): finds the earliest sent date across all pending rows, then fetches all Finances API events in 60-day windows for EU and FE endpoints — typically 2–3 API calls total regardless of how many pending rows there are.

- Reads col Q (order ID), col P (sent date), col B (region) from row 4 down
- Skips rows that already have a valid fee value (never overwrites)
- Rows marked `RETRY` or `ERR:` are retried on the next run
- On 429 during a window fetch: sleeps 15s and retries once
- On unsettled orders: leaves blank — next run retries

```javascript
// Run once manually in GAS editor, or set a daily time-based trigger:
backfillMCFFees()
```

### `backfillTrackingNumbers()` — DISABLED 2026-07-31
~~Same pattern for tracking numbers → writes static values into col Z.~~ This assumed col Z
(`BF_COL_RESULT`) was a free static cache. It isn't — see the "Root cause" note below for what
actually happened. The function now throws immediately instead of running. Left in the file for
reference; do not re-enable without re-reading that section and picking a real target column.

### `retryR429Errors()`
Scans col R (row 1108+) for cells whose live `=AMZTK()`/`=AMZTK_JP()` formula is currently
showing a 429 QuotaExceeded error, and re-fetches those specific orders directly.

- **Tracking number found** → freezes the cell to a static `=HYPERLINK(...)` (same URL format
  the live formula produces). This is deliberate: confirmed live that Sheets' custom-function
  engine does not reliably re-run `AMZTK()`/`AMZTK_JP()` just because Apps Script rewrites the
  identical formula text via the API — the cache was primed correctly but the cell kept showing
  the stale 429 text. Freezing to a static value is the only way to guarantee the sheet actually
  displays the result, and it also drops the row out of the 429 filter so future runs progress to
  new rows instead of re-fetching the same already-resolved orders every hour.
- **No tracking number yet** (order genuinely not ready, no error) → primes the cache `AMZTK`
  reads from, then forces a real recalculation by writing a placeholder formula and then the
  original formula back (two distinct writes — Sheets only re-triggers custom functions on an
  actual content change), so the stale error clears to blank.

Bounded so it can never call SP-API forever in one run:
- stops after 40 rows per execution (`RETRY_R_MAX_ROWS_PER_RUN`)
- aborts the whole run after 5 consecutive rows still come back 429 (`RETRY_R_MAX_CONSEC_429`) —
  quota is clearly still exhausted, so it stops and lets the next hourly trigger pick up where
  it left off, instead of hammering SP-API in a tight loop

An hourly time-based trigger for `retryR429Errors()` is already installed on the live script
(set up via the Apps Script Triggers UI / a one-off installer that has since been removed from
the codebase). Manage or remove it from **Apps Script editor → Triggers** (clock icon).

```javascript
// Can also be run manually in GAS editor:
retryR429Errors()
```

**2026-07-30 diagnosis (later found incomplete — see the 2026-07-31 correction right below):** the
429s weren't primarily an SP-API problem, they were a *design* problem. As of 2026-07-30, **2,489
of 3,755** col R cells still held a live `=AMZTK()`/`=AMZTK_JP()` formula. Every time Google Sheets
does a broad recalculation (opening the file, any edit, etc.) it fires a burst of dozens-to-
hundreds of these custom-function calls **simultaneously** — confirmed live in the Executions log,
~50+ AMZTK/AMZTK_JP calls within the same 3-second window — which floods the FBA Outbound API
(`GetFulfillmentOrder`) and produces fresh 429s faster than this hourly, 40-row-capped,
row-1108+-only retry could ever keep up with. The diagnosis at the time *assumed* col Z
(`BF_COL_RESULT`) held 2,817 already-resolved tracking numbers sitting unused, and built
`freezeTrackingColumnR()` to copy those into col R as static values.

**That assumption was wrong, and acting on it corrupted col R the same day.** Col Z is not a
tracking-number cache — it's a live `ARRAYFORMULA` anchored at Z3 that computes
`Product Price − Transportation Fee` for the whole column (confirmed 2026-07-31 by reading Z3's
actual formula, and by scanning all 3,755 rows: 0 held anything that looked like a tracking
number, 2,613 (70%) were plain prices). `freezeTrackingColumnR()`'s first run froze those price
values straight into col R, displaying numbers like `17.99` where a tracking number like
`JJD000390016584418318` should have been. Caught and reported immediately by the sheet owner;
reverted live via `unfreezeAmztkFormulas()` (1,448 cells restored), confirmed clean by a full
column scan against a "< 500" price-magnitude heuristic (0 remaining cells flagged).

**That heuristic itself had a blind spot, and it hid a second corrupted pocket.** The "< 500"
threshold was tuned for EUR/GBP-style decimal prices (e.g. `17.99`) and completely missed col Z's
JPY-region margin values, which are bare integers with no decimal point and are routinely in the
thousands (e.g. `2799`, `4712`, `3222`) — comfortably over 500. Every JP row's col R had been
frozen the same way and displayed a plain yen amount instead of a real Yamato tracking number
(confirmed via a raw `GetFulfillmentOrder` dump: the actual tracking numbers, e.g.
`371434845460`, were never anywhere near `2799`). Reported by the sheet owner again with a
screenshot. Fixed by replacing the magnitude threshold with a shape-based check in
`_looksLikePriceNotTracking()`: any value containing a decimal point, or any bare digit string
under 8 digits, is treated as a price/margin, never a tracking number (real tracking numbers on
this sheet are always either alphanumeric with a country prefix, or a 12+ digit all-numeric
string). `unfreezeAmztkFormulas()` was also fixed to call this shared function instead of its own
separate, independently-wrong inline check — the two had drifted apart, which is exactly how the
JP pocket went undetected through the first "clean" verification. Re-running
`unfreezeAmztkFormulas()` with the fixed heuristic restored **73 more cells**; a follow-up scan
confirmed 0 remaining price-like frozen cells anywhere in col R, and spot-checked JP rows now
resolve to real Yamato tracking numbers.

### `freezeTrackingColumnR()`, `backfillTrackingNumbers()`, `dailyTrackingMaintenance()` — DISABLED 2026-07-31
All three are now stubs that throw immediately instead of running, so a stray manual run (or a
re-added trigger) can't repeat the corruption. The `dailyTrackingMaintenance()` daily trigger that
caused the incident has already been deleted from **Apps Script editor → Triggers**. Only
`backfillMCFFeesRecent()` (30 min) and `retryR429Errors()` (hourly) triggers remain active.

### `resolveBlankTrackingNumbers()` — companion to `retryR429Errors()`, added 2026-08-28
Fixes a *separate* failure mode from the 429/corruption saga above: `retryR429Errors()` only
detects cells showing literal `"SP-API error 429"`/`"QuotaExceeded"` text. But `AMZTK()`/
`AMZTK_JP()` treat SP-API's `"Unable to get order info"` (400 InvalidInput) as "not shipped yet"
and silently return `''` — no error text ever lands in the cell, so `retryR429Errors()`'s
text-match filter can never see or retry these rows.

**Confirmed live** (direct SP-API calls, bypassing GAS entirely) that self-created MCF fulfillment
orders stop being resolvable via `GetFulfillmentOrder` roughly **~1 day after creation** —
same-day orders returned 200 OK, every order from the day before onward returned 400
`"Unable to get order info"`, **including orders the sheet already shows a real tracking number
for**. Per the team's own observed norm, a real tracking number only appears after at least
~24h, so any row whose live formula happened to (re)run before that window opened, or after the
order aged out of `GetFulfillmentOrder`, is stuck blank forever — the same order id 400s
identically on every future call, with no way back.

Called from the end of `retryR429Errors()` (rides the same hourly trigger — no new trigger
needed). Two passes over every row with a live `AMZTK()`/`AMZTK_JP()` formula in col R (formula
text contains `"AMZTK("` — distinguishes a still-live formula from an already-frozen static
`=HYPERLINK(...)` cell):

1. **Lock-in** — a cell already *displaying* a real tracking number is frozen to a static
   `=HYPERLINK(...)` immediately. Needed because `AMZTK()`'s own success cache is only 6h — once
   that expires, the very next recalculation re-queries the same now-aged-out order and silently
   blanks a previously-correct cell. This pass is pure reads/writes, no API cost, so it always
   runs in full regardless of the row cap below.
2. **Backfill** — a genuinely blank cell (skipping rows younger than `BLANK_R_MIN_AGE_HOURS`,
   24h) is retried directly. Resolved → frozen the same way as pass 1. Still unresolved past
   `BLANK_R_GIVE_UP_DAYS` (3 days) → the cell is frozen to a literal blank (formula removed) so it
   stops silently re-querying SP-API — and burning quota — for an order that will never resolve.
   Still unresolved but younger than that → left alone, retried again next hour. Bounded to
   `BLANK_R_MAX_ROWS_PER_RUN` (60) actual API calls per run, with the same
   `BLANK_R_MAX_CONSEC_429`-consecutive-429 abort as `retryR429Errors()`.

Scans the **whole sheet** from `BF_START_ROW`, not just `RETRY_R_START_ROW+` — this failure mode
isn't confined to any particular row range.

`BLANK_R_GIVE_UP_DAYS` (3) is a judgment call, not a measured constant — gives real stragglers a
wide margin past the observed ~24h cutoff before permanently giving up on a row. Adjust if real
data suggests otherwise.

**Not yet independently re-verified live post-deploy** — worth checking `Logger` output after the
next hourly run, and whether previously-blank rows in the 1700s–1900s range (reported stuck as of
2026-08-28) start resolving or freezing to blank as expected.

**Do not re-enable these** without first deciding what column (if any) can safely hold a resolved
tracking-number cache — col Z cannot, since it's genuinely in use for the profit-margin figure.
Options not yet evaluated: use a new/unused column as the real cache; skip caching entirely and
throttle the live `AMZTK()`/`AMZTK_JP()` calls directly (e.g. stagger recalculation, add a sheet
add-on level rate limiter); or keep relying on `retryR429Errors()` alone and simply widen its
row range past the stale `RETRY_R_START_ROW = 1108` assumption (live 429s were observed as early
as row 29). This is an open design question — raise it with the sheet owner before building
anything new here.

### `retryZeroTransportationFees()`
Cleans up the col Y "literal 0" backlog left over from the GCX fee-type filtering bug (see the
`MCFFee()` note above). `backfillMCFFees()` skips any row whose col Y already has a non-empty
value — including `"0"` — so those rows never get reprocessed on their own even after the fix.

Clears up to 40 zero-fee cells per run (`RETRY_ZERO_FEE_MAX_ROWS_PER_RUN`) back to blank, then
runs `backfillMCFFees()` so they get recomputed correctly. Reuses all of `backfillMCFFees()`'s
own safeguards (429 sleep-and-retry per window) — no separate retry logic needed. Bounded to 40
rows/run so one execution can't try to reprocess the whole historical backlog (which could span
many months of Finances API windows) at once; run hourly and it works through the backlog
gradually.

```javascript
// Run manually, or set up an hourly trigger via Apps Script editor → Triggers:
retryZeroTransportationFees()
```

---

## Col Y fee source: Settlement Reports, not Finances API (rebuilt 2026-07-30)

`backfillMCFFeesRecent()` is installed as a 30-min time-driven trigger (Head deployment).

### Why the Finances API doesn't work for this sheet's orders
Live testing across a diverse sample of orders (spanning 2025-03 → 2026-07) confirmed Amazon's
Finances API has **no fee record at all** for these self-created ("Non-Amazon `<country>`"
marketplace) MCF shipments — which is nearly every row in this sheet:

- `getFulfillmentOrderRaw(orderId)` succeeds and returns a real `fulfillmentShipments[].amazonShipmentId`
  + tracking number — the order genuinely exists and shipped.
- But `GET /finances/v0/financialEvents` for the matching window (checked all 35 event-list types,
  not just `ShipmentEventList`) never contains the GCX order id, the resolved `displayableOrderId`
  (usually identical to the GCX id — Amazon doesn't reassign a separate order number), or the
  `amazonShipmentId`, anywhere.
- `GET /finances/v0/orders/{orderId}/financialEvents` using the GCX id directly returns a
  structurally valid but completely empty `FinancialEvents` object — Amazon just doesn't index
  these shipments under any ID we have via this API surface.
- Even the standard **Settlement Report scanned by `SellerOrderId`/`order-id`** doesn't match —
  these shipments post under a synthetic `S02-...` pseudo order id there, never the GCX id.

### What actually works: `merchant-order-id` in the Settlement Report
The `GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2` report (Reports API, not Finances API) has a
**separate `merchant-order-id` column that preserves the sheet's exact GCX id**, alongside a
proper fee line — confirmed live:
```
merchant-order-id=GCX-FR-260716-2695  amount-type=ItemFees  amount-description=FBAPerUnitFulfillmentFee  amount=-7,50
```
`_buildSettlementFeeMap(ep, sinceDate)` (sp-api.js) lists settlement reports created in the window,
downloads + parses each (tab-delimited, gzip sometimes), and sums `amount-type === 'ItemFees'` rows
per `merchant-order-id` — this is a **direct match** against the sheet's Q-column value, no numeric
alias or `displayableOrderId` resolution needed (unlike the old Finances-API approach).
`backfillMCFFeesRecent()` calls this directly now instead of the old windowed Finances API scan.

First real run after this fix (2026-07-30): **109 rows written** with real fee values (e.g. 3.48,
4.96, 7.5, 8.35, 16.28 €/£) out of 695 pending, confirmed visually in the sheet.

### Constraints that shape the design
- The reports-listing endpoint (`GET /reports/2021-06-30/reports`) rejects `createdSince` older
  than ~90 days ("RequestedFromDate ... is more than 90 days old") — so this can only ever cover
  the last ~89 days, which is exactly why `backfillMCFFeesRecent()` (not the unbounded
  `backfillMCFFees()`) is the one wired to the 30-min trigger. `days` beyond 89 is clamped down.
- This account settles very frequently — **109 settlement reports** were found within one 89-day
  window in testing, some with 200k+ lines. `_buildSettlementFeeMap()` sorts newest-first and caps
  to the 40 most recent reports per run, plus a 4.5-minute wall-clock budget that bails early and
  logs how many were scanned — both so one run can't blow past GAS's execution time limit. Rows
  whose fee posted in an older/unscanned batch simply resolve on a later 30-min run once it rotates
  through more of the window (already-filled rows are always skipped, so this converges over time
  without reprocessing anything).
- The report-**document** download endpoint (`GET /reports/.../documents/{id}`) 429s in bursts —
  confirmed live (~20 successful downloads then a wall of `QuotaExceeded`). Each download is paced
  800ms apart and wrapped in `spapiFetchWithRetry` (one retry, 15s wait); a report that still fails
  is skipped and picked up on a later run rather than blocking the rest.
- Settlement report amounts use a locale decimal **comma** in EU reports (`"7,50"`, not `"7.50"`) —
  `_parseSettlementAmount()` handles both.
- `nextToken` must be sent **alone** on paginated listing calls — combining it with the original
  filter params 400s with `"NextToken cannot be specified with other input parameters"`.

### Estimate fallback for un-settled but shipped orders (added 2026-08-13)
Confirmed live 2026-08-11: for a genuinely recent order (shipped within the last ~2 weeks),
**neither** the settlement report **nor** the Finances API has any fee data yet — Amazon simply
hasn't generated the settlement data at all, regardless of source. Tested directly: 0/103 blank
rows matched against a live Finances API scan that found 1,590 *other* orders' fees in the same
window.

So `backfillMCFFeesRecent()` now has a second pass after the settlement-map lookup: for any row
still unmatched, if **col R (Tracking Number) already holds a real value** (the order has
genuinely shipped — an `AMZTK`/`AMZTK_JP` formula resolved), it calls `_fetchMcfFeePreview()`
(`getFulfillmentPreview`, see `MCFFee` docs above) and writes that as an **estimate** — a fresh
shipping-quote using today's rates, not the actual historical charge. Rows with no tracking number
yet (order hasn't shipped) are left alone; there's nothing to estimate.

Every estimate-written cell gets a note starting with `FEE_ESTIMATE_NOTE_PREFIX`
(`'ESTIMATE (getFulfillmentPreview)'`). The "already filled, skip" check at the top of
`backfillMCFFeesRecent()` treats a cell carrying this note as **still pending**, not done — so on
a later run, once the settlement report actually covers that order, the real fee overwrites the
estimate and `.setNote('')` clears the tag. No separate cleanup step or second trigger needed —
the same 30-min run does both the real lookup and the estimate fallback, and self-corrects once
real data exists.

Capped to 30 estimate calls per run (2 SP-API calls each: `getFulfillmentOrderRaw` + preview),
paced 500ms apart, own 90s time budget and 5-consecutive-429 circuit breaker — separate from the
settlement-scan's own budget above, so a burst of newly-shipped rows can't blow past GAS's
execution limit on top of whatever the settlement scan already used that run.

There's also a standalone one-off version for backfilling a specific row range manually:
`fillPreviewEstimatesForRange(startRow, endRow)` — same estimate + note-tagging logic, run once
from the editor (e.g. `fillPreviewEstimatesForRange(2782, 2909)`).

### Bug fixed 2026-08-14: estimate fallback stalled on the same old rows forever
Because an estimate-tagged cell stays "pending" every run (by design — so a real fee can overwrite
it later), it competes with never-estimated rows for the same 30-row/run cap above. The candidate
list was built in plain ascending row order with no distinction between "never estimated" and
"already has a stale estimate, just needs re-checking" — so once there were more stale-estimate
rows than the 30/run cap, the same low-numbered rows consumed the **entire** budget every single
cycle, forever. Confirmed live: 24h after the first estimate pass (rows 2782-2823), the trigger was
still only ever re-touching that same block every 30 min — with a 100% completion rate and zero
errors — while orders shipped on 2026-08-12 through 2026-08-14 never got a first estimate at all.
This is the same class of bug as the settlement-report scan's own prioritization fix above (oldest-
unscanned-first) — a cap without prioritization silently starves whatever's ranked below it.

**Fix:** `needsEstimate` is now split into `neverEstimated` (rows with no `FEE_ESTIMATE_NOTE_PREFIX`
note yet) and `alreadyEstimated` (a refresh, not a first fill), concatenated in that order before
slicing to the 30/run cap — so new orders always get their first estimate before any budget is
spent re-estimating rows that already have one. Verified live: the next run immediately advanced
past the stale 2782-2823 block, and orders sent 2026-08-12 through 2026-08-14 showed real estimated
values within one cycle.

**How to apply:** any future cap-and-slice pattern over a "still pending" list needs the same
prioritization — a row that's `pending` for a "needs refresh" reason should never rank ahead of a
row that's `pending` because it has never been touched at all.

### Bug fixed 2026-07-30: `backfillMCFFeesRecent()` crashed on every scheduled run
Apps Script time-driven triggers call the handler with a trigger event object as the first
argument (not `undefined`) — `days = days || 90` let that object through as a truthy value, so
`days * 24 * 3600 * 1000` evaluated to `NaN`, `cutoff` became an `Invalid Date`, and
`cutoff.toISOString()` threw immediately, before any SP-API call was made. This was firing every
5 minutes (not 30) and failing 100% of the time — confirmed in the Executions log:
`RangeError: Invalid time value at backfillMCFFeesRecent(autoFill:669:28)`. Manual "Run" clicks in
the editor never hit this (no event object passed), which is why a handful of col Y values existed
despite the trigger never working. Fixed by guarding `days` to require an actual finite number;
the old 5-min trigger was deleted and replaced with a correct 30-min one.

---

## `onEdit` Automation (`autoFill.js` → `onEdit_mcf`)

Installable on-edit trigger; only acts on the `MCF 발송 로그` tab, rows 4+:

| Trigger column | Action |
|---------------|--------|
| Col I (9) | Writes today's date to col M if empty |
| Col N (14) | Sets col S to `Pending` if col S is empty |
| Col U (21) | Writes today's date to col P (if empty) and sets col S to `MCF` |
| Col Y (25) | Writes today's date to col T (if empty) and sets col W to `MCF` |
| Col F (6), Col AB (28) = `STOCK`, Col W (23) = `RUN` | Call `updateMcfStockForRow` / `runStockCheckOnly` / `processMCFRow` — ⚠️ these live in the commented-out `MCFGen.js`, so these branches currently throw `ReferenceError` (logged, no sheet change) |

---

## Web app for GCX Reply (`tamperMonkey.js`)

`doGet(e)` — JSON endpoint called by the GCX Reply Tampermonkey script (Seller Central MCF flow).

| Param | Meaning |
|---|---|
| `email` (required) | Customer email; matched against col H after stripping any `+<uuid>` segment on both sides (ABM proxy addresses are per-case — fixed 2026-08-14) |
| `match` | `first` (default) or other match mode for duplicate emails |
| `action=markMcf` + `person` | Writes 담당자 (col U; isolated so an off-list name in the strict dropdown can't abort the rest — fixed 2026-08-18), date (col P) and Status `MCF`, flushes, and returns the Order ID the col Q formula generates (`GCX-{B}-{YYMMDD}-{A}`) |
| (no action) | Polls col Q for up to 10 s and returns the Order ID |

Deployed as a web app (`USER_DEPLOYING`, anonymous access); after changes, redeploy the existing
deployment so the URL the Tampermonkey script calls stays the same.

---

## Daily Report (`main.js`)

`MCFReporter` runs on a time-based trigger (weekdays 9 AM KST).  
It scans the `MCF 발송 로그` sheet for rows where col R is filled but col S is empty and posts a cardsV2
card (`MCF Daily Report`, subtitle `Tracking No. 미기입: N개 (date)`, one "N행 이동" row-jump button per
row) to the GCX T2 ESC. Ticket space, retrying up to 5 times. The webhook URL is currently hard-coded in
`main.js` (`<WEBHOOK_URL>`; a commented-out line points at the private test space) — moving it to a
Script Property is recommended.

### Setting triggers

```javascript
// In GAS editor, run once:
triggerTester()    // schedules MCFReporter 1 minute from now for testing
```

---

## Version Control & Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/MCF_Tracking
clasp pull            # pull latest from GAS
clasp push --force    # push changes (custom formulas + triggers use HEAD)

cd ~/Desktop/GCX
git add GAS_Operations/MCF_Tracking/
git commit -m "..."
git push
```

`clasp push` does not move the web-app deployment — for `tamperMonkey.js` changes also run
`clasp deploy -i <deploymentId>`.

## GStore MCF 발송 로그 — Tracking Number (col P)

Added 2026-10-07. Same scheme as "MCF 발송 로그" col R, on the Google Store replacement sheet:
B = MCF 발송 국가, N = 발송일자, O = MCF Order ID, **P = Tracking Number**, data from row 5.

- P5:P1000 hold a live formula (`installGStoreTrackingFormulas()` re-creates it in empty cells only):
  `=IF(OR(B5="",O5=""),"",IF(B5<>"JP",HYPERLINK("https://www.swiship.de/track?id="&AMZTK(O5),AMZTK(O5)),HYPERLINK("https://www.swiship.jp/track?id="&AMZTK_JP(O5),AMZTK_JP(O5))))`
- `resolveBlankTrackingNumbers()` (hourly, via `retryR429Errors`) now loops over `TRACKING_SHEETS`
  (both sheets) and freezes resolved cells to a static `=HYPERLINK(...)`, retries blanks older than
  24h, and gives up after 3 days. For GStore it also retries 429/`ERR:` cells itself (there's no
  `retryR429Errors` pass for that sheet). The SP-API call budget (60/run) is shared across both sheets.
- Each hourly run first tops up the formula in the 200 empty P cells below the last row with an MCF Order ID
  (`_topUpGStoreTrackingFormulas_`), so new rows never need it pasted by hand. GStore is processed before
  the main sheet so the shared lookup budget can't starve it.
- Order ids are trimmed before the SP-API call — a trailing `\r` in the cell produced a bare HTML 400
  (`<title>x</title>`) that looked like a JP credential problem.
- Cells showing `EU ERR:`/`JP ERR:` (non-429) are no longer frozen as if they were tracking numbers —
  they're retried instead.
