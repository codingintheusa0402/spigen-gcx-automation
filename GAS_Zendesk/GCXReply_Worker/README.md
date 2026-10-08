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

## Secrets / bindings
`wrangler secret put` (values live only in Cloudflare and `~/.config/gcx_reply_worker/` on the Mac,
never in git): `GCX_KEY`, `PIDX_PUSH_KEY`, `ZD_WEBHOOK_TOKEN`, `AWS_ACCESS_KEY_ID`,
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
