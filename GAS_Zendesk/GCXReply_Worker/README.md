# GCXReply_Worker — fast read path for the GCX Reply panel

Cloudflare Worker (`gcx-reply`, https://gcx-reply.kjw-55b.workers.dev, account kjw@spigen.com,
Workers Paid) that answers the GCX Reply Tampermonkey panel's two read lookups with the
**same JSON as the GCXReply_GAS web app**:

| Request | Same as GAS | Typical time |
|---|---|---|
| `GET /?orderId=XXX` | `?orderId=` (SP-API order + address + buyerInfo) | 0.2–1.2 s (GAS 3–12 s) |
| `GET /?asin=XXX` | `?asin=` (product index) | ~0.2 s (GAS ~1.6 s) |
| `GET /?orderId=…&asin=…` | combined | |
| `&fresh=1` | skip the 90 s / prefetch cache (panel's manual **Lookup** button) | |

Requests need header `x-gcx-key` (secret `GCX_KEY`). The panel reads that key from the Zendesk
dynamic-content item `gcx_reply_worker_key` (id 63025656454809) — only signed-in agents can read it,
so the key is never in the public repo. Everything else (ABM relay log/claims, inferReason, MCF)
stays on GAS, and the panel falls back to GAS whenever the Worker errors, is slow (>3.5 s), or
the key is missing.

## How data gets here
- **Orders**: SP-API directly (ported 1:1 from `GCXReply_GAS/Code.js` v2.7.3 — keep them in sync).
  90 s per-instance cache, plus a 2 h prefetched copy in KV (`pf:<orderId>`).
- **Products**: GAS `buildProductIndex_` (every 4 h, and `forceRefreshProductIndex`) POSTs the index to
  `/admin/pidx` (header `x-push-key` = secret `PIDX_PUSH_KEY`), stored in KV key `pidx`. Missing or
  >6 h old → the Worker asks GAS `?asin=` instead. GAS reads `WORKER_URL` / `WORKER_PUSH_KEY` from its
  Script Properties (set once via the `setWorkerConfig` doPost action, which requires a proof only the
  SP-API secret holder can compute).
- **Prefetch (Phase 3)**: Zendesk trigger *"GCX Reply – prefetch order data (Worker)"* (id 63026144521881,
  notification only — changes nothing on the ticket) fires on ticket create / end-user update when the
  Order ID field (360021934132) is set → webhook *"GCX Reply Worker – order prefetch"*
  (01M4D620JPAJZ4N52S92F1ZGGZ, Bearer = secret `ZD_WEBHOOK_TOKEN`) → `POST /prefetch`.

## Autocomplete telemetry (TM v3.9.1+)
The panel's Smart Autocomplete POSTs small de-identified event batches to `POST /ac/log` (header
`x-gcx-key`). Stored in KV as `aclog:<day>:<agentId>:<tab>` (60-day TTL). Nothing here touches the
order/product read path.

**Dashboard (no tools needed):** `GET /ac/log/view?k=<ACLOG_VIEW_KEY>&days=7` renders the summary as an HTML page for the whole team + per agent (key in `~/.config/gcx_reply_worker/secrets.json`, read-only, de-identified data only; keep the link private). `DELETE /ac/log?agent=<id>` (x-push-key) purges one agent's batches (test cleanup).

Read it back as JSON (header `x-push-key` = `PIDX_PUSH_KEY`, value in `~/.config/gcx_reply_worker/secrets.json`):
```bash
PK=$(python3 -c "import json;print(json.load(open('$HOME/.config/gcx_reply_worker/secrets.json'))['PIDX_PUSH_KEY'])")
curl -s "https://gcx-reply.kjw-55b.workers.dev/ac/log?days=7" -H "x-push-key: $PK" | python3 -m json.tool   # summary
curl -s "https://gcx-reply.kjw-55b.workers.dev/ac/log?days=7&raw=1" -H "x-push-key: $PK" > events.json     # every event
```
Summary fields: `overall` / `byAgent` / `byLevel` (a paragraph, b sentence, c n-gram) / `byKind` (p reply,
n note) / `bySource` (0 typed, 1 history, 2 macro, -1 n-gram) each with episode counts per outcome,
`acceptRate` (tab+word), `usefulRate` (+typed-through), accepted chars and compute-time p50/p95;
`messages` (chars, accepted share, keystroke coverage); `topAccepted` / `topIgnored` suggestion texts
(ignored ones include samples of what the agent wrote instead); `agents` (last version / corpus size).

## Secrets / bindings
`wrangler secret put` (values live only in Cloudflare and `~/.config/gcx_reply_worker/` on the Mac,
never in git): `GCX_KEY`, `PIDX_PUSH_KEY`, `ZD_WEBHOOK_TOKEN`, `ACLOG_VIEW_KEY`, `AWS_ACCESS_KEY_ID`,
`AWS_SECRET_ACCESS_KEY`, `LWA_CLIENT_ID[_JP|_IN]`, `LWA_CLIENT_SECRET[_JP|_IN]`, `LWA_REFRESH_TOKEN[_JP|_IN]`.
KV binding `GCX_KV`. Var `GAS_URL`.

## Deploy / operate
```bash
cd GAS_Zendesk/GCXReply_Worker
npx wrangler deploy            # code
npx wrangler tail              # live logs
npx wrangler secret put NAME   # rotate a secret (e.g. after an LWA secret rotation — IN expires 2026-12-12)
```
Rotating `GCX_KEY`: update the secret **and** the Zendesk dynamic-content item; panels drop a
rejected key (401) and re-read it automatically, falling back to GAS meanwhile.

**Turning it off**: nothing to change in the panel — delete/disable the Worker (or the dynamic-content
item) and every panel silently uses GAS again.
