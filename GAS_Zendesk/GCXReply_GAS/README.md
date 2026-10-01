# GCXReply_GAS

Backend for the **GCX Reply** Tampermonkey script ([Browser_Extensions/tampermonkey_scripts](../../Browser_Extensions/tampermonkey_scripts/)) — a Google Apps Script Web App that does signed Amazon SP-API order lookups and Google Sheet product lookups so the userscript doesn't need to hold AWS/LWA credentials client-side.

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | The deployed Web App — current live backend |
| `appsscript.json` | GAS manifest |
| `sp-api-proxy.py` | Local Flask proxy mirroring `Code.js`'s SigV4 signing logic in Python, for testing SP-API calls without redeploying the GAS web app |
| `.sp-api-config.example.json` | Template for `sp-api-proxy.py`'s local config (copy to `.sp-api-config.json`, gitignored) |
| `v*.gs` (~100+ files) | Versioned reference-copy archive of every past GCX Reply script version (`v1.9.7` → current) — see [Archive workflow](#archive-workflow) below |

---

## `Code.js` — Web App

**Endpoint:** `?orderId=XXX` \| `?asin=XXX` \| `?orderId=XXX&asin=XXX` \| `?action=inferReason&review=...&category=...`
**Deploy as:** Execute as Me, Access: Anyone (or Anyone anonymous)

| Feature | Detail |
|---------|--------|
| SP-API order lookup | SigV4-signed requests across 4 region configs (EU, FE/Japan, NA, India — India shares the EU endpoint but a separate Seller Central account/refresh token) |
| Product lookup | Reads `SHEET_ID` `Data` tab (SKU, 모델명, 브랜드, etc.) plus per-marketplace sheets (`DE`,`NL`,`SE`,`ES`,`UK`,`FR`,`IT`,`JP`,`IN`,`SG`) |
| Product index (v2.7.0) | `refreshProductIndex` (15-min trigger, installed once via `setupProductIndexTrigger`; `keepWarm` also refreshes it) pre-builds all product + marketplace data into 128 hashed ScriptCache buckets so `?asin=` is one cache read (~1.6 s vs 8–16 s). Requests never build it; if missing or >2 h old they fall back to the original live sheet reads |
| AI 인입사유 (DR) | `inferReason` action — Gemini-based defect/reason classification (`gemini-2.5-flash-lite` → `gemini-2.5-flash` fallback), cached |
| ABM relay logging | `upsertAbmRelayLog_()` writes/updates rows in the `ASIN_Master_MondaySync` spreadsheet's `ABM_Relay_Log` tab, which [ASIN_Master_MondaySync](../../GAS_Operations/ASIN_Master_MondaySync/) prunes on a daily trigger |

### Script Properties required

| Property | Covers |
|----------|--------|
| `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` | SigV4 signing |
| `LWA_CLIENT_ID` / `LWA_CLIENT_SECRET` / `LWA_REFRESH_TOKEN` | EU + NA |
| `LWA_CLIENT_ID_JP` / `LWA_CLIENT_SECRET_JP` / `LWA_REFRESH_TOKEN_JP` | Japan (FE) |
| `LWA_CLIENT_ID_IN` / `LWA_CLIENT_SECRET_IN` / `LWA_REFRESH_TOKEN_IN` | India |

---

## Archive workflow

Per project convention: after every `Code.js` edit, delete the previous `v*.gs` file(s) if superseded and add a new `v{version}.gs` snapshot of the full script, then `clasp push --force`. This keeps a full point-in-time history of every deployed GCX Reply backend version independent of git history (useful for rolling back a live deployment without a git checkout).

---

## `sp-api-proxy.py`

Local dev tool — same SigV4/LWA signing logic as `Code.js`, exposed as a small local Flask app (`get_order`, `debug_order` routes) for testing SP-API responses directly from a terminal, without going through the deployed GAS web app.

```bash
cp .sp-api-config.example.json .sp-api-config.json   # fill in credentials
cd ~/Desktop/GCX/GAS_Zendesk/GCXReply_GAS
./sp-api-proxy.py
```

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Zendesk/GCXReply_GAS
clasp push --force
```
