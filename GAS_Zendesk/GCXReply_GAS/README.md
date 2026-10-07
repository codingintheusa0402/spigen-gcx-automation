# GCXReply_GAS

Backend for the **GCX Reply** Tampermonkey script ([Browser_Extensions/tampermonkey_scripts](../../Browser_Extensions/tampermonkey_scripts/)). It is a Google Apps Script Web App that does signed Amazon SP-API order lookups and Google Sheet product lookups, so the userscript never has to hold AWS/LWA credentials client-side. It also owns the **ABM relay log** that GCX Reply and [ABM_TicketMerge](../ABM_TicketMerge/) use to make sure every agent reply to an Amazon Buyer Message reaches the buyer exactly once.

**Current backend version:** `Code.js` v2.7.1 (2026-10-01), paired with userscript GCX Reply v3.7.2
**Script ID:** `1xi5UnQ9zcPm_QjiahGvfTwZXTDSsCFH_EZkA4E_8jwON3M7HSKXHIGks`
**Web app (pinned deployment):** `https://script.google.com/macros/s/AKfycbw2Vdwk197LXB6oUAzuHS8sKamD5uqKZJDLvcHzbftWJk-M65XV1fAnTqiZo7ZEm4hk/exec`. This URL is hardcoded as `GAS_URL` in the userscript and as `GCX_GAS_URL` in ABM_TicketMerge.

## Screenshots

![`ABM_Relay_Log` tab (case IDs blurred)](docs/abm_relay_log.jpg)
*`ABM_Relay_Log` tab (case IDs blurred)*

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | The deployed Web App (the current live backend) |
| `appsscript.json` | GAS manifest: web app runs as `USER_DEPLOYING`, access `ANYONE_ANONYMOUS`; execution API `ANYONE` |
| `sp-api-proxy.py` | Local FastAPI proxy that mirrors `Code.js`'s SigV4 signing in Python, for testing SP-API calls without redeploying |
| `.sp-api-config.example.json` | Template for `sp-api-proxy.py`'s local config (placeholders only) |
| `v*.gs` (~150 files, `v1.9.7` → `v3.7.2`) | Byte-for-byte archive copies of the **userscript** `GCX Reply.user.js`, one per released version. They are not backend code: `.claspignore` excludes `v*.gs` (and `sp-api-proxy.py`), so they never reach the GAS project and live only in git. See [Archive workflow](#archive-workflow). |

---

## `Code.js` — Web App

### `doGet` endpoints

| Request | Returns |
|---------|---------|
| `?orderId=XXX-XXXXXXX-XXXXXXX` | Order, items, shipping address, buyer info (RDT-signed), buyer 2-year purchase/refund stats. Also returns product info for the first item's ASIN. Cached 90 s per order. |
| `?asin=XXX` | Product info (ASIN Master + per-marketplace sheets) |
| `?orderId=…&asin=…` | Both |
| `?action=inferReason&review=…&category=…` | AI 인입사유 (DR) label |
| `?action=abmRelayPending&clientVersion=…` | Undelivered ABM relay rows (rows with attachments are only handed to clients ≥ `3.3.12`) |
| `?action=abmRelayStatus&ticketId=…` | **All** relay rows for one ticket (ABM_TicketMerge's reconciliation uses this to dedup) |
| `?action=abmRelayAll&limit=N` | Newest N rows of the log (default 50) |

### `doPost` actions (JSON body `{action: …}`)

| Action | Purpose |
|--------|---------|
| `logAbmRelay` | Upsert a row keyed by `RelayKey` (`${ticketId}_${startTimeMs}`), so each reply gets its own row |
| `setAbmRelayStatus` | Mark a row delivered / failed (the panel's "Mark delivered") |
| `deleteAbmRelayRow` | Remove a row |
| `claimAbmRelay` | Atomic claim before a browser re-sends a queued row, so two browsers can't both send it |
| `claimAbmSend` | Guard on the **initial** send. Keyed by ticket + normalized-text hash, held for 5 min under a `LockService` lock (since 2026-08-19). If the lock times out it fails **open**: a missed real reply is worse than a rare duplicate. |

### Features

| Feature | Detail |
|---------|--------|
| SP-API order lookup | SigV4-signed requests tried across 4 region configs in turn: EU, FE/Japan, NA, India. India uses the EU endpoint but has its own Seller Central account and refresh token. An expired LWA token is cleared and retried once. |
| Product lookup | Reads the `SHEET_ID` `Data` tab (SKU, 모델명, 브랜드, 제조사명, 기종명, 색상명, 대분류, 생산업체, 원산지정보) plus per-marketplace tabs (`DE`,`NL`,`SE`,`ES`,`UK`,`FR`,`IT`,`JP`,`IN`,`SG`) in `MARKET_SS_ID` |
| Product index (v2.7.0 / v2.7.1, 2026-10-01) | All product and marketplace data is precomputed into 128 hashed ScriptCache buckets, so `?asin=` is one cache read (~1.6 s instead of 8–16 s). `refreshProductIndex` runs on a 15-min trigger but only rebuilds when the index is more than 4 h old, because product sheets change weekly to monthly. Requests never build the index: if it is missing or over 6 h old they fall back to live sheet reads. **After editing ASIN Master or a market sheet, run `forceRefreshProductIndex`.** Userscript v3.7.1 also caches each ASIN's response in GM storage for 6 h (the manual Product button bypasses that cache). |
| AI 인입사유 (DR) | `inferReason` classifies a review or claim text against the `GCX 인입사유` sheet (`DEFECT_SS_ID`) using Gemini: `gemini-2.5-flash-lite`, falling back to `gemini-2.5-flash`. Results are cached 6 h (`DR_v24_` key prefix). |
| ABM relay log | `ABM_Relay_Log` tab in the `SHEET_ID` spreadsheet. Columns: `Timestamp, RelayKey, TicketId, CommentId, CaseId, Marketplace, Status, Attempts, LastError, MessageText`. Missing columns are migrated in place automatically. [ASIN_Master_MondaySync](../../GAS_Operations/ASIN_Master_MondaySync/) prunes it daily. |
| Keep-warm | `keepWarm` runs every 5 min (install once with `setupKeepWarmTrigger`). It pings the cache and refreshes the product index if it is stale. |

### Triggers (install once from the editor)

| Function | Schedule | Installer |
|----------|----------|-----------|
| `keepWarm` | every 5 min | `setupKeepWarmTrigger` |
| `refreshProductIndex` | every 15 min (rebuilds only if > 4 h old) | `setupProductIndexTrigger` |

Manual helpers: `forceRefreshProductIndex`, `diagJP` (JP LWA diagnostic), `testInferReason`, `updateFeedbackSheet` (writes release notes into the `GCX Reply 피드백` sheet), `fixProductSheetData`.

### Script Properties required

| Property | Covers |
|----------|--------|
| `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` | SigV4 signing |
| `LWA_CLIENT_ID` / `LWA_CLIENT_SECRET` / `LWA_REFRESH_TOKEN` | EU + NA |
| `LWA_CLIENT_ID_JP` / `LWA_CLIENT_SECRET_JP` / `LWA_REFRESH_TOKEN_JP` | Japan (FE) |
| `LWA_CLIENT_ID_IN` / `LWA_CLIENT_SECRET_IN` / `LWA_REFRESH_TOKEN_IN` | India |
| `GEMINI_API_KEY` | `inferReason` (returns empty if unset) |

### Recent changes (since 2026-08)

- **2026-10-01, v2.7.1:** the product index rebuilds only when older than 4 h (it used to rebuild every 15 min). Added `forceRefreshProductIndex`. Evicting a bucket clears the index metadata, so the next trigger run rebuilds it.
- **2026-10-01, v2.7.0:** added the precomputed product index (above). Output was verified identical to the old path on 38 ASINs.
- **2026-08-19:** `claimAbmSend_` check-and-set is now wrapped in `LockService`. Before this, two calls 785 ms apart could both pass the check and send the same reply twice.
- **2026-08-12:** new `claimAbmSend` endpoint guards the userscript's initial ABM send against firing twice for one reply.

---

## Deployment

The userscript and ABM_TicketMerge call the **pinned** `/exec` deployment above. `clasp push` alone only updates `@HEAD`, so after any `Code.js` change also cut a new version onto that deployment:

```bash
cd ~/Desktop/GCX/GAS_Zendesk/GCXReply_GAS
clasp push --force
clasp deploy -i AKfycbw2Vdwk197LXB6oUAzuHS8sKamD5uqKZJDLvcHzbftWJk-M65XV1fAnTqiZo7ZEm4hk -d "v2.x.y: <what changed>"
clasp deployments   # confirm the version number advanced
```

If only a new `v*.gs` archive was added and `Code.js` didn't change, `clasp push` reports "already up to date". That is expected (archives are claspignored), and no redeploy is needed.

---

## Archive workflow

After every change to `GCX Reply.user.js`, copy it **verbatim** to `v{@version}.gs` in this folder, for example `v3.7.2.gs`. Old archives are kept, never deleted. Commit the new archive together with the userscript change. This gives a point-in-time copy of every released userscript version that doesn't depend on a git checkout.

---

## `sp-api-proxy.py`

Local dev tool with the same SigV4/LWA signing logic as `Code.js`. It is a small FastAPI/uvicorn app on `http://127.0.0.1:5050` with two routes: `GET /order/{order_id}` and `GET /debug/{order_id}`. Use it to test SP-API responses from a terminal without going through the deployed GAS web app.

```bash
pip3 install fastapi uvicorn requests
cp .sp-api-config.example.json ~/.sp-api-config.json   # fill in your own credentials (never commit)
python3 ~/Desktop/GCX/GAS_Zendesk/GCXReply_GAS/sp-api-proxy.py
curl http://127.0.0.1:5050/order/XXX-XXXXXXX-XXXXXXX
```
