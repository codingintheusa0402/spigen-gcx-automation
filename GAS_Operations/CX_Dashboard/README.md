# CX_Dashboard

Container-bound Google Apps Script project that pulls live Amazon SP-API data (EU + FE regions) into a Google Spreadsheet — marketplaces, orders, sales metrics, customer feedback, and FBA inventory — via an **SP-API Dashboard** menu. Also exposes SP-API data as custom sheet formulas (`=SPORDERS()`, `=SPSALES()`, `=SPINVENTORY_ASIN()`, etc.).

**Script ID:** `1FIyZcgVPPlVE_A5zrFB01khTt-5QeYT6xAtgcLGuh-PHXByt_1ZTz84S`
**Linked spreadsheet:** `1C3QOyhjGk-zMKr0H8-lwEMNmzk128ijuiLfSUmsEUm4`

## Screenshots

![SP-API Dashboard menu (as of 2026-10-07 the SP-API formulas return `LWA token fetch failed: 401 invalid_client`, so the data tabs are empty)](docs/menu.jpg)
*SP-API Dashboard menu (as of 2026-10-07 the SP-API formulas return `LWA token fetch failed: 401 invalid_client`, so the data tabs are empty)*

---

## Files

| File | Purpose |
|------|---------|
| `sp-api.js` | SP-API auth (LWA refresh-token → access token, AWS SigV4 signing), EU/FE endpoint resolver, `spapiFetch` / `spapiFetchWithRetry` |
| `formulas.js` | Custom sheet formulas (results cached in `CacheService`) |
| `menu.js` | `onOpen()` menu, `setupSheets()`, Refresh functions, `Config` tab reader |
| `appsscript.json` | GAS manifest (scopes: spreadsheets, external_request, script.scriptapp) |

---

## Menu actions (SP-API Dashboard)

| Menu item | Function | Writes tab | SP-API endpoint |
|-----------|----------|-----------|-----------------|
| **Setup Sheets** | `setupSheets()` | Creates/resets `Config`, creates empty `Marketplaces`, `Orders`, `Order Items`, `Sales Metrics`, `Feedback`, `Inventory` | — |
| **Refresh Marketplaces** | `refreshMarketplaces()` | `Marketplaces` | `/sellers/v1/marketplaceParticipations` (EU + FE) |
| **Refresh Orders (all)** | `refreshOrders()` | `Orders` | `/orders/v0/orders` (paginated, up to `MAX_PAGES`) |
| **Refresh Sales Metrics** | `refreshSalesMetrics()` | `Sales Metrics` | `/sales/v1/orderMetrics` |
| **Refresh Feedback** | `refreshFeedback()` | `Feedback` | `/customer-feedback/2024-06-01/feedbacks` |
| **Refresh Inventory** | `refreshInventory()` | `Inventory` | `/fba/inventory/v1/summaries` |

Each Refresh clears the target tab and rewrites it with a bold, frozen header row. No time-driven triggers — everything is manual.

---

## Custom formulas

| Formula | Returns |
|---------|---------|
| `=SPMARKETPLACES()` | All marketplace participations (cached 1 h) |
| `=SPORDERS(marketplaceId, startDate, endDate, [status])` | Up to 100 orders (BuyerEmail masked) |
| `=SPORDERITEMS(orderId, [endpoint])` | Line items for one order |
| `=SPSALES(marketplaceId, startDate, endDate, [granularity], [fulfillmentNetwork])` | Interval / Units / Orders / TotalSales … |
| `=SPFEEDBACK(marketplaceId, startDate, endDate)` | Customer feedback entries |
| `=SPINVENTORY(marketplaceId)` | FBA inventory per SKU |
| `=SPINVENTORY_ASIN(marketplaceId, [asins], [includeZero])` | FBA inventory summed per ASIN; `asins` may be a cell range or comma list |
| `=SPINVENTORY_DEBUG(marketplaceId)` | Raw first page of the inventory API, for diagnostics |

---

## Config sheet settings

After running **Setup Sheets**, edit the `Config` tab (col A = key, col B = value):

| Setting | Default | Notes |
|---------|---------|-------|
| `MARKETPLACE_ID` | `A1F83G8C2ARO7P` (UK) | Run `=SPMARKETPLACES()` to find your IDs |
| `START_DATE` | `2025-01-01` | Filter start (YYYY-MM-DD) |
| `END_DATE` | today | Filter end |
| `SALES_GRANULARITY` | `Day` | Day / Week / Month / Year / Total / Hour |
| `FULFILLMENT_NETWORK` | `All` | All / AFN / MFN |
| `MAX_PAGES` | `10` | Pagination cap for Refresh Orders (100 orders/page) |

> Only EU and FE endpoints are wired. JP / AU / SG marketplace IDs (`A1VC38T7YXB528`, `A39IBJ37TRP1C6`, `A19VAU5U5O7RUS`) route to FE with the `_JP` LWA profile; everything else routes to EU. NA marketplaces are not supported.

---

## Script Properties (SP-API credentials)

Set in **Extensions → Apps Script → Project Settings → Script Properties** (names only — never commit values):

| Key | Description |
|-----|-------------|
| `LWA_CLIENT_ID` / `LWA_CLIENT_SECRET` / `LWA_REFRESH_TOKEN` | EU LWA app |
| `LWA_CLIENT_ID_JP` / `LWA_CLIENT_SECRET_JP` / `LWA_REFRESH_TOKEN_JP` | FE LWA app (falls back to the EU values if unset) |
| `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` | AWS keys for SigV4 |
| `AWS_SESSION_TOKEN` | Optional |
| `SPAPI_HOST_EU` / `SPAPI_REGION_EU` | Optional overrides (default `sellingpartnerapi-eu.amazon.com` / `eu-west-1`) |
| `SPAPI_HOST_FE` / `SPAPI_REGION_FE` | Optional overrides (default `sellingpartnerapi-fe.amazon.com` / `us-west-2`) |

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/CX_Dashboard
clasp push --force
```
