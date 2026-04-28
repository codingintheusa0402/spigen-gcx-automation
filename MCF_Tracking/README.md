# MCF_Tracking

Google Apps Script project for Spigen GCX — Multi-Channel Fulfillment (MCF) order tracking, stock lookup, fee estimation, and daily reporting to Google Chat.

**GAS Script ID:** `1kDfEUVEEJ7TCA3HOMF6EYFTjbeZZKIeg_X84wCbLT1-tQqJI2ZlPUCxp`  
**Linked sheet:** `MCF 발송 로그` (Spreadsheet ID: `1g6a-S7eeA1oY19aTEFhTNAyp2A5nLqLNkPRqOqriWfc`)

---

## Files

| File | Purpose |
|------|---------|
| `sp-api.js` | SP-API auth (LWA + AWS SigV4), core fetch, retry logic, and all custom sheet formulas (`AMZTK`, `AMZTK_JP`, `MCFFee`, `MCFFee_JP`, `getMcfStockByAsin`, `MCFFeeDebug`) |
| `autoFill.js` | `onEdit` trigger, `backfillTrackingNumbers()`, and `backfillMCFFees()` — batch server-side fee/tracking writes to col Y/Z |
| `main.js` | `MCFReporter` — daily Google Chat card alert listing rows missing a tracking number |
| `MCFGen.js` | (Archived / commented out) MCF order creation and stock-check helpers via SP-API |
| `triggerGen.js` | `triggerGen()` — sets weekday 9 AM KST time-based triggers for `MCFReporter`; `triggerTester()` schedules a test run 1 minute out |
| `tamperMonkey.js` | TamperMonkey-related helpers |
| `appsscript.json` | GAS manifest (timezone, OAuth scopes) |

---

## Custom Sheet Formulas (`sp-api.js`)

### `=AMZTK(orderId)`
Returns the tracking number for an EU MCF order. Tries EU endpoint first, falls back to FE.

### `=AMZTK_JP(orderId)`
Same as `AMZTK` but tries FE (Japan/AU/SG) first.

### `=MCFFee(orderId)` / `=MCFFee(sentDate, orderId)`
Returns the actual settled MCF fulfillment fee via the Finances API. Always uses the actual charged amount (not an estimate). Returns blank until the order settles (usually a few days after shipment) — retries automatically on the next sheet recalculation.

**For bulk use, prefer `backfillMCFFees()` over this formula** — ARRAYFORMULA is not supported and many simultaneous formula calls will queue indefinitely.

| Usage | Example | Notes |
|-------|---------|-------|
| 1-arg | `=MCFFee(Q2)` | Searches last 180 days |
| 2-arg | `=MCFFee(P2, Q2)` | `P2` = sent date (yyyy-mm-dd), `Q2` = orderId — faster, less bandwidth |

```
=IF(Q2<>"", MCFFee(P2, Q2), "")
```

**Lookup strategy:** single Finances API date-range scan — no FBA Outbound call.
1. Builds date window: `sentDate → sentDate+60d` (or last 180d if no sentDate)
2. Calls `listFinancialEvents` for the window (up to 5 pages × 100 events)
3. Matches `SellerOrderId` / `AmazonOrderId` against the GCX order ID
4. Sums `OrderFeeList` (MCF order-level fees) + `ShipmentFeeList` + `ItemFeeList`

> **Why no FBA Outbound call?** Calling `getFulfillmentOrderRaw` per-cell with 15+ concurrent formulas hits SP-API rate limits and causes 30s GAS timeout. The Finances API scan alone is fast enough (~1-2s per cell). For orders that settle under a different Amazon-format ID (`S02-XXXXXXX-XXXXXXX`), run **`backfillMCFFees()`** — it resolves `displayableOrderId` in a single batch call.

- On 429 QuotaExceeded: returns blank and retries after 90 seconds automatically.
- Tries EU endpoint first, falls back to FE.
- Required SP-API roles: **Amazon Fulfillment** + **Finance and Accounting**.

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

### `backfillTrackingNumbers()`
Same pattern for tracking numbers → writes static values into col Z.

---

## `onEdit` Automation (`autoFill.js`)

Fires on any edit in the `MCF 발송 로그` sheet (rows 4+):

| Trigger column | Action |
|---------------|--------|
| Col I (9) | Writes today's date to col M if empty |
| Col N (14) | Sets col S to `Pending` if col S is empty |
| Col U (21) | Writes today's date to col P (if empty) and sets col S to `MCF` |
| Col F (6) | Calls `updateMcfStockForRow` to refresh stock in col H |
| Col Y (25) | Writes today's date to col T (if empty) and sets col W to `MCF` |
| Col AB (28) = `STOCK` | Runs stock check only (`runStockCheckOnly`) |
| Col W (23) = `RUN` | Runs full MCF row processing (`processMCFRow`) |

---

## Daily Report (`main.js`)

`MCFReporter` runs on a time-based trigger (weekdays 9 AM KST, set by `triggerGen`).  
It scans the `MCF 발송 로그` sheet for rows where col R is filled but col S is empty (order sent, tracking not yet entered) and posts a Google Chat card to the GCX T2 ESC. Ticket space with direct row-jump links.

### Setting triggers

```javascript
// In GAS editor, run once:
triggerGen()       // sets weekday 9AM triggers up to the hardcoded end date
triggerTester()    // schedules MCFReporter 1 minute from now for testing
```

---

## Version Control & Deployment

```bash
# Pull latest from GAS
cd ~/Desktop/GCX/MCF_Tracking
clasp pull

# Push changes to GAS
clasp push

# Commit and push to GitHub
cd ~/Desktop/GCX
git add MCF_Tracking/
git commit -m "..."
git push
```
