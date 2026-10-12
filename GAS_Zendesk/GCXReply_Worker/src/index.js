// GCX Reply — Cloudflare Worker (fast read path for the GCX Reply panel)
//
// Serves the SAME JSON as the GCXReply_GAS web app for its two read endpoints,
// so the Tampermonkey panel can use either interchangeably:
//   GET /?orderId=XXX            → { order, items, itemsStatus, itemsError, rdtStatus, rdtError,
//                                    address, buyer, orderCount, totalPurchases, totalRefunds, region }
//   GET /?asin=XXX               → { product, productSource, allSources, marketplaces }
//   GET /?orderId=XXX&asin=XXX   → both merged (same as GAS doGet)
// Everything else (ABM relay log/claims, inferReason, MCF) stays on GAS.
//
// Ported 1:1 from GCXReply_GAS/Code.js v2.7.3 (fetchOrderDataFresh_, findOrderRegion_,
// fetchBuyerPurchaseStats_, fetchBuyerRefundCount_, lookupFromIndex_). Any change to
// those GAS functions must be mirrored here — the panel falls back to GAS whenever
// this Worker errors, but silently-different data would not be caught by that.
//
// Admin / integration endpoints (separate secrets):
//   POST /admin/pidx   ← GAS buildProductIndex_ pushes the product index (x-push-key: PIDX_PUSH_KEY)
//   POST /ac/log       ← panel autocomplete telemetry batches (x-gcx-key); GET /ac/log?days=7[&raw=1] (x-push-key) reads them,
//                        DELETE /ac/log?agent=<id> (x-push-key) purges one agent, GET /ac/log/view?k=<ACLOG_VIEW_KEY> = HTML dashboard
//   POST /prefetch     ← Zendesk webhook on ticket create / end-user update (Bearer ZD_WEBHOOK_TOKEN)
//   GET  /health       ← no auth, no data
//
// Secrets (wrangler secret put): GCX_KEY, PIDX_PUSH_KEY, ZD_WEBHOOK_TOKEN, AWS_ACCESS_KEY_ID,
//   AWS_SECRET_ACCESS_KEY, LWA_CLIENT_ID[_JP|_IN], LWA_CLIENT_SECRET[_JP|_IN], LWA_REFRESH_TOKEN[_JP|_IN]
// Bindings: KV namespace GCX_KV (product index + prefetched orders). Var: GAS_URL.

const REGIONS = [
  { endpoint: 'https://sellingpartnerapi-eu.amazon.com', region: 'eu-west-1', cred: 'main' },
  { endpoint: 'https://sellingpartnerapi-fe.amazon.com', region: 'us-west-2', cred: 'jp'   },
  { endpoint: 'https://sellingpartnerapi-na.amazon.com', region: 'us-east-1', cred: 'main' },
  { endpoint: 'https://sellingpartnerapi-eu.amazon.com', region: 'eu-west-1', cred: 'in'   },
];

const MARKETPLACE_MAP = [
  ['.com.sg', 'A19VAU5U5O7RUS'], ['.com.au', 'A39IBJ37TRP1C6'], ['.com.mx', 'A1AM78C64UM0Y8'],
  ['.com.tr', 'A33AVAJ2PDY3EV'], ['.co.uk',  'A1F83G8C2ARO7P'], ['.co.jp',  'A1VC38T7YXB528'],
  ['.de',     'A1PA6795UKMFR9'], ['.fr',     'A13V1IB3VIYZZH'], ['.it',     'APJ6JRA9NG5V4'],
  ['.es',     'A1RKKUPIHCS9HS'], ['.nl',     'A1805IZSGTT6HS'], ['.pl',     'AZ1PBY3F3E3AE'],
  ['.se',     'A2NODRKZP88ZB9'], ['.be',     'AMEN7PMS3EDWL'],  ['.in',     'A21TJRUUN4KGV'],
  ['.ca',     'A2EUQ1WTGCTBG2'], ['.tr',     'A33AVAJ2PDY3EV'], ['.com',    'ATVPDKIKX0DER'],
];

const PRODUCT_COLS = ['SKU','모델명','브랜드','제조사명','기종명','색상명','대분류','생산업체','원산지정보'];
const PIDX_BUCKETS = 128;
const PIDX_MAX_AGE_MS = 6 * 60 * 60 * 1000;
const ORDER_CACHE_SEC = 90;              // same as GAS ord2_ cache
const PREFETCH_TTL_SEC = 2 * 60 * 60;    // prefetched (webhook) orders
const ORDER_RE = /^\d{3}-\d{7}-\d{7}$/;
const ASIN_RE = /^B[A-Z0-9]{9}$/;

// ── Small cache — where GAS used CacheService ────────────────────────────────
// Per-isolate memory for everything; small long-lived keys (LWA token, region
// hints, the "RDT unavailable" flag) are also kept in KV so new isolates start
// warm. (The Cache API is a no-op on *.workers.dev, so it isn't used.)
const _mem = new Map();
const KV_PERSIST = /^(lwa_|oreg_|nordt_)/;
async function cacheGet(env, key) {
  const m = _mem.get(key);
  if (m && m.exp > Date.now()) return m.v;
  if (KV_PERSIST.test(key)) {
    const v = await env.GCX_KV.get('c:' + key);
    if (v !== null) { _mem.set(key, { v, exp: Date.now() + 60000 }); return v; }
  }
  return null;
}
async function cachePut(env, key, value, ttlSec) {
  _mem.set(key, { v: value, exp: Date.now() + ttlSec * 1000 });
  if (_mem.size > 2000) { for (const [k, e] of _mem) if (e.exp <= Date.now()) _mem.delete(k); }
  if (KV_PERSIST.test(key)) await env.GCX_KV.put('c:' + key, value, { expirationTtl: Math.max(60, Math.floor(ttlSec)) });
}
async function cacheDel(env, key) {
  _mem.delete(key);
  if (KV_PERSIST.test(key)) await env.GCX_KV.delete('c:' + key);
}

function marketplaceId(salesChannel) {
  if (!salesChannel) return null;
  const s = salesChannel.toLowerCase();
  const m = MARKETPLACE_MAP.find(([suffix]) => s.includes(suffix));
  return m ? m[1] : null;
}

// ── LWA token (memory + Cache API) ───────────────────────────────────────────
const _lwaMem = {};
async function getLwaToken(env, cred) {
  const mem = _lwaMem[cred];
  if (mem && mem.exp > Date.now()) return mem.token;
  const hit = await cacheGet(env, 'lwa_' + cred);
  if (hit) { _lwaMem[cred] = { token: hit, exp: Date.now() + 60000 }; return hit; }
  const sfx = cred === 'jp' ? '_JP' : cred === 'in' ? '_IN' : '';
  const resp = await fetch('https://api.amazon.com/auth/o2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: env['LWA_REFRESH_TOKEN' + sfx],
      client_id: env['LWA_CLIENT_ID' + sfx],
      client_secret: env['LWA_CLIENT_SECRET' + sfx],
    }),
  });
  const txt = await resp.text();
  let d; try { d = JSON.parse(txt); } catch { d = {}; }
  if (!d.access_token) throw new Error('LWA failed: ' + txt);
  const ttl = Math.max(60, Math.min(d.expires_in - 300, 21600));
  await cachePut(env, 'lwa_' + cred, d.access_token, ttl);
  _lwaMem[cred] = { token: d.access_token, exp: Date.now() + Math.min(ttl, 600) * 1000 };
  return d.access_token;
}
async function dropLwaToken(env, cred) { delete _lwaMem[cred]; await cacheDel(env, 'lwa_' + cred); }

// ── AWS SigV4 (WebCrypto) ────────────────────────────────────────────────────
const enc = new TextEncoder();
const hex = buf => [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
async function sha256Hex(msg) { return hex(await crypto.subtle.digest('SHA-256', enc.encode(msg))); }
async function hmac(key, msg) {
  const k = await crypto.subtle.importKey('raw', typeof key === 'string' ? enc.encode(key) : key,
    { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return crypto.subtle.sign('HMAC', k, enc.encode(msg));
}
async function signingKey(secret, dateStamp, region) {
  const kDate = await hmac('AWS4' + secret, dateStamp);
  const kRegion = await hmac(kDate, region);
  const kService = await hmac(kRegion, 'execute-api');
  return hmac(kService, 'aws4_request');
}
function amzDates() {
  const iso = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, ''); // yyyyMMddTHHmmssZ
  return { amzDate: iso, dateStamp: iso.slice(0, 8) };
}

async function spApiRequest(env, endpoint, region, cred, method, fullPath, { token, body } = {}) {
  const accessToken = token || await getLwaToken(env, cred);
  const host = endpoint.replace('https://', '');
  const { amzDate, dateStamp } = amzDates();
  const qIdx = fullPath.indexOf('?');
  const uriPath = qIdx >= 0 ? fullPath.slice(0, qIdx) : fullPath;
  const rawQuery = qIdx >= 0 ? fullPath.slice(qIdx + 1) : '';
  const canonQuery = rawQuery
    ? rawQuery.split('&').map(pair => {
        const eq = pair.indexOf('=');
        const k = eq >= 0 ? pair.slice(0, eq) : pair;
        const v = eq >= 0 ? pair.slice(eq + 1) : '';
        return encodeURIComponent(decodeURIComponent(k)) + '=' + encodeURIComponent(decodeURIComponent(v));
      }).sort().join('&')
    : '';
  const bodyStr = body !== undefined ? JSON.stringify(body) : '';
  const signHdrs = { host, 'x-amz-access-token': accessToken, 'x-amz-date': amzDate };
  if (method === 'POST') signHdrs['content-type'] = 'application/json';
  const keys = Object.keys(signHdrs).sort();
  const canonHdrs = keys.map(k => k + ':' + signHdrs[k]).join('\n') + '\n';
  const signedHdrs = keys.join(';');
  const canonReq = [method, uriPath, canonQuery, canonHdrs, signedHdrs, await sha256Hex(bodyStr)].join('\n');
  const scope = `${dateStamp}/${region}/execute-api/aws4_request`;
  const sts = ['AWS4-HMAC-SHA256', amzDate, scope, await sha256Hex(canonReq)].join('\n');
  const sig = hex(await hmac(await signingKey(env.AWS_SECRET_ACCESS_KEY, dateStamp, region), sts));
  const headers = {
    'x-amz-access-token': accessToken, 'x-amz-date': amzDate,
    'Authorization': `AWS4-HMAC-SHA256 Credential=${env.AWS_ACCESS_KEY_ID}/${scope}, SignedHeaders=${signedHdrs}, Signature=${sig}`,
  };
  if (method === 'POST') headers['Content-Type'] = 'application/json';
  const res = await fetch(endpoint + fullPath, { method, headers, body: method === 'POST' ? bodyStr : undefined });
  return { status: res.status, body: await res.text() };
}
const spApiGet = (env, ep, rg, cred, path, token) => spApiRequest(env, ep, rg, cred, 'GET', path, { token });

async function getRdt(env, endpoint, region, cred, orderId) {
  const r = await spApiRequest(env, endpoint, region, cred, 'POST', '/tokens/2021-03-01/restrictedDataToken', {
    body: {
      restrictedResources: [
        { method: 'GET', path: `/orders/v0/orders/${orderId}/items` },
        { method: 'GET', path: `/orders/v0/orders/${orderId}/buyerInfo`, dataElements: ['buyerInfo'] },
      ],
    },
  });
  if (r.status !== 200) return { token: null, status: r.status, error: r.body };
  try { return { token: JSON.parse(r.body).restrictedDataToken || null, status: r.status, error: null }; }
  catch { return { token: null, status: r.status, error: r.body }; }
}

// ── Buyer stats (only runs when SP-API returns a BuyerEmail — mirrors GAS) ──
async function fetchBuyerPurchaseStats(env, endpoint, region, cred, salesChannel, buyerEmail) {
  const mpId = marketplaceId(salesChannel);
  if (!mpId || !buyerEmail) return null;
  const cacheKey = 'bstat_' + btoa(unescape(encodeURIComponent(buyerEmail))).replace(/[+/=]/g, '').slice(0, 50);
  const hit = await cacheGet(env, cacheKey);
  if (hit) { try { return JSON.parse(hit); } catch {} }

  const createdAfter = new Date(Date.now() - 2 * 365.25 * 24 * 3600 * 1000).toISOString().slice(0, 19) + 'Z';
  let totalPurchases = 0, nextToken = null, page = 0;
  const orderIds = [];
  do {
    const path = nextToken
      ? `/orders/v0/orders?NextToken=${encodeURIComponent(nextToken)}`
      : `/orders/v0/orders?MarketplaceIds=${encodeURIComponent(mpId)}&BuyerEmail=${encodeURIComponent(buyerEmail)}&CreatedAfter=${encodeURIComponent(createdAfter)}&MaxResultsPerPage=100`;
    const r = await spApiGet(env, endpoint, region, cred, path);
    if (r.status !== 200) break;
    try {
      const d = JSON.parse(r.body);
      const orders = d.payload?.Orders || [];
      totalPurchases += orders.length;
      orders.forEach(o => { if (o.AmazonOrderId) orderIds.push(o.AmazonOrderId); });
      nextToken = d.payload?.NextToken || null;
    } catch { break; }
    page++;
  } while (nextToken && page < 5);

  const totalRefunds = await fetchBuyerRefundCount(env, endpoint, region, cred, orderIds);
  const result = { totalPurchases, totalRefunds };
  try { await cachePut(env, cacheKey, JSON.stringify(result), 300); } catch {}
  return result;
}

async function fetchBuyerRefundCount(env, endpoint, region, cred, orderIds) {
  if (!orderIds || !orderIds.length) return 0;
  const BATCH = 30;
  let count = 0;
  const uncached = [];
  for (const id of orderIds) {
    const hit = await cacheGet(env, 'fin_' + id);
    if (hit !== null) { if (hit === '1') count++; } else uncached.push(id);
  }
  for (let i = 0; i < uncached.length; i += BATCH) {
    const batch = uncached.slice(i, i + BATCH);
    let results;
    try { results = await Promise.all(batch.map(id => spApiGet(env, endpoint, region, cred, `/finances/v0/orders/${id}/financialEvents`))); }
    catch { break; }
    let financeUnavailable = false;
    for (let j = 0; j < results.length; j++) {
      const code = results[j].status;
      if (code === 403) { financeUnavailable = true; break; }
      if (code === 200) {
        try {
          const ev = JSON.parse(results[j].body).payload?.FinancialEvents;
          const refunded = (ev?.RefundEventList?.length > 0) || (ev?.GuaranteeClaimEventList?.length > 0) || (ev?.ChargebackEventList?.length > 0);
          await cachePut(env, 'fin_' + batch[j], refunded ? '1' : '0', 3600);
          if (refunded) count++;
        } catch {}
      }
    }
    if (financeUnavailable) break;
    if (i + BATCH < uncached.length) await new Promise(r => setTimeout(r, 2000));
  }
  return count;
}

// ── Order lookup ─────────────────────────────────────────────────────────────
async function findOrderRegion(env, orderId) {
  const regionErrors = [];
  const hintKey = 'oreg_' + orderId.slice(0, 3);
  const hint = Number(await cacheGet(env, hintKey));
  const order_ = REGIONS.map((r, i) => i);
  if (hint > 0 && hint < REGIONS.length) { order_.splice(hint, 1); order_.unshift(hint); }
  for (const ri of order_) {
    const { endpoint, region, cred } = REGIONS[ri];
    let r;
    try { r = await spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}`); }
    catch (e) { regionErrors.push(`${cred}:LWA(${e.message})`); continue; }
    if (r.status === 403 && r.body.includes('expired')) {
      await dropLwaToken(env, cred);
      try { r = await spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}`); }
      catch (e) { regionErrors.push(`${cred}:LWA-retry(${e.message})`); continue; }
    }
    if (r.status !== 200) {
      const detail = (r.status === 403 && r.body.includes('expired')) ? '(auth-revoked)' : '';
      regionErrors.push(`${cred}:${r.status}${detail}`);
      continue;
    }
    const order = JSON.parse(r.body).payload || {};
    if (!order.AmazonOrderId) {
      regionErrors.push(`${cred}:200-noId(${r.body.substring(0, 120).replace(/\s+/g, ' ')})`);
      continue;
    }
    if (hint !== ri) { try { await cachePut(env, hintKey, String(ri), 604800); } catch {} }
    return { endpoint, region, cred, order };
  }
  throw new Error('Order not found — ' + regionErrors.join(' | '));
}

async function fetchOrderDataFresh(env, orderId) {
  const { endpoint, region, cred, order } = await findOrderRegion(env, orderId);
  // Same as GAS v2.7.3: skip only the always-failing RDT + items calls; buyerInfo
  // is always fetched (it returns BuyerName for many orders even without an RDT).
  const noRdtKey = 'nordt_v2_' + cred;
  let known = null;
  try { known = JSON.parse((await cacheGet(env, noRdtKey)) || 'null'); } catch {}

  let rdtResult, itemsR, addrR, buyerR;
  if (known) {
    rdtResult = { token: null, status: known.rdtStatus, error: known.rdtError };
    itemsR = { status: known.itemsStatus, body: known.itemsError };
    [addrR, buyerR] = await Promise.all([
      spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}/address`),
      spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}/buyerInfo`),
    ]);
  } else {
    rdtResult = await getRdt(env, endpoint, region, cred, orderId);
    const rdtToken = rdtResult.token || undefined;
    [itemsR, addrR, buyerR] = await Promise.all([
      spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}/items`, rdtToken),
      spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}/address`),
      spApiGet(env, endpoint, region, cred, `/orders/v0/orders/${orderId}/buyerInfo`, rdtToken),
    ]);
    if (!rdtResult.token && rdtResult.status === 400 && itemsR.status === 403) {
      try {
        await cachePut(env, noRdtKey, JSON.stringify({
          rdtStatus: rdtResult.status, rdtError: rdtResult.error, itemsStatus: itemsR.status, itemsError: itemsR.body,
        }), 21600);
      } catch {}
    }
  }

  const buyer = buyerR.status === 200 ? JSON.parse(buyerR.body).payload || {} : {};
  const stats = await fetchBuyerPurchaseStats(env, endpoint, region, cred, order.SalesChannel, buyer.BuyerEmail || null);
  return {
    order,
    items: itemsR.status === 200 ? JSON.parse(itemsR.body).payload?.OrderItems || [] : [],
    itemsStatus: itemsR.status,
    itemsError: itemsR.body,
    rdtStatus: rdtResult.status,
    rdtError: rdtResult.error,
    address: addrR.status === 200 ? JSON.parse(addrR.body).payload?.ShippingAddress || {} : {},
    buyer,
    orderCount: stats ? stats.totalPurchases : null,
    totalPurchases: stats ? stats.totalPurchases : null,
    totalRefunds: stats ? stats.totalRefunds : null,
    region,
  };
}

// 90 s per-colo cache (= GAS ord2_), then 2 h webhook-prefetched copy in KV.
async function fetchOrderData(env, ctx, orderId, fresh) {
  const ck = 'ord2_' + orderId;
  if (!fresh) {
    const hit = await cacheGet(env, ck);
    if (hit) { try { return JSON.parse(hit); } catch {} }
    const pf = await env.GCX_KV.get('pf:' + orderId, 'json');
    if (pf && pf.data && Date.now() - pf.ts < PREFETCH_TTL_SEC * 1000) return pf.data;
  }
  const result = await fetchOrderDataFresh(env, orderId);
  ctx.waitUntil(cachePut(env, ck, JSON.stringify(result), ORDER_CACHE_SEC).catch(() => {}));
  return result;
}

// ── Product index (pushed by GAS buildProductIndex_) ─────────────────────────
let _pidxMem = null; // { builtAt, buckets: {key: parsedObj|string}, loadedAt }
function pidxBucketKey(asin) {
  let h = 0;
  for (let i = 0; i < asin.length; i++) h = (h * 31 + asin.charCodeAt(i)) | 0;
  return 'pidx_v1_' + (((h % PIDX_BUCKETS) + PIDX_BUCKETS) % PIDX_BUCKETS);
}
async function loadIndex(env) {
  if (_pidxMem && Date.now() - _pidxMem.loadedAt < 5 * 60 * 1000) return _pidxMem;
  const raw = await env.GCX_KV.get('pidx', 'json');
  _pidxMem = raw ? { builtAt: raw.builtAt, buckets: raw.buckets, loadedAt: Date.now() } : null;
  return _pidxMem;
}
function pidxExpand(vals) {
  if (!vals) return null;
  const o = {};
  PRODUCT_COLS.forEach((c, i) => { if (vals[i] !== null) o[c] = vals[i]; });
  return o;
}
async function lookupFromIndex(env, asin) {
  if (!ASIN_RE.test(asin)) return null;
  const idx = await loadIndex(env);
  if (!idx || Date.now() - idx.builtAt > PIDX_MAX_AGE_MS) return null;
  const k = pidxBucketKey(asin);
  let bucket = idx.buckets[k];
  if (bucket === undefined) return null;
  if (typeof bucket === 'string') { bucket = JSON.parse(bucket); idx.buckets[k] = bucket; }
  const rec = bucket[asin] || {};
  const sheet1 = pidxExpand(rec.s1);
  const sheet2 = pidxExpand(rec.s2);
  let product = sheet1 || sheet2;
  let productSource = sheet1 ? 'sheet1' : sheet2 ? 'sheet2' : null;
  if (!product && rec.p) {
    product = {
      'SKU': '', '모델명': rec.p[1], '브랜드': '',
      '제조사명': '', '기종명': rec.p[0], '색상명': '',
      '대분류': '', '생산업체': '', '원산지정보': '',
    };
    productSource = 'market';
  }
  return {
    product, productSource,
    allSources: { sheet1: sheet1 || null, sheet2: sheet2 || null },
    marketplaces: (rec.m || []).map(([name, gid, cell]) => ({ name, gid, cell })),
  };
}
// Index unavailable → ask GAS (its own index or live path), so the answer is always GAS-identical.
async function lookupProductFull(env, asin) {
  const hit = await lookupFromIndex(env, asin);
  if (hit) return hit;
  const r = await fetch(`${env.GAS_URL}?asin=${encodeURIComponent(asin)}`, { redirect: 'follow' });
  const d = JSON.parse(await r.text());
  if (d.error) throw new Error(d.error);
  return { product: d.product, productSource: d.productSource, allSources: d.allSources, marketplaces: d.marketplaces || [] };
}

// ── HTTP ─────────────────────────────────────────────────────────────────────
const json = (data, status = 200) => new Response(JSON.stringify(data), {
  status, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
});
function safeEq(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length || !a.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

async function handleRead(request, env, ctx, url) {
  const key = request.headers.get('x-gcx-key') || '';
  if (!safeEq(key, env.GCX_KEY)) return json({ error: 'unauthorized' }, 401);
  const p = url.searchParams;
  const orderId = p.get('orderId');
  const asin = p.get('asin');
  const fresh = p.get('fresh') === '1';
  if (!orderId && !asin) return json({ error: 'Provide orderId and/or asin parameter' });
  try {
    const result = {};
    if (orderId) {
      if (!ORDER_RE.test(orderId)) return json({ error: 'Invalid order ID format' });
      const orderData = await fetchOrderData(env, ctx, orderId, fresh);
      Object.assign(result, orderData);
      const itemAsin = !asin && orderData.items && orderData.items[0] ? orderData.items[0].ASIN : null;
      if (itemAsin) Object.assign(result, await lookupProductFull(env, itemAsin));
    }
    if (asin) Object.assign(result, await lookupProductFull(env, asin));
    return json(result);
  } catch (err) {
    return json({ error: err.message });
  }
}

async function handlePidxPush(request, env) {
  if (!safeEq(request.headers.get('x-push-key') || '', env.PIDX_PUSH_KEY)) return json({ error: 'unauthorized' }, 401);
  const body = await request.json();
  if (!body || !body.buckets || !body.builtAt) return json({ error: 'bad payload' }, 400);
  await env.GCX_KV.put('pidx', JSON.stringify({ builtAt: body.builtAt, buckets: body.buckets }));
  _pidxMem = null;
  return json({ ok: true, buckets: Object.keys(body.buckets).length });
}

// Zendesk webhook: { order_id, ticket_id }. Fetch in the background and keep it
// in KV for 2 h so the agent's first open is a KV read instead of SP-API calls.
async function handlePrefetch(request, env, ctx) {
  const auth = request.headers.get('authorization') || '';
  if (!safeEq(auth.replace(/^Bearer\s+/i, ''), env.ZD_WEBHOOK_TOKEN)) return json({ error: 'unauthorized' }, 401);
  let body = {};
  try { body = await request.json(); } catch {}
  const orderId = String(body.order_id || '').trim();
  if (!ORDER_RE.test(orderId)) return json({ ok: true, skipped: 'no order id' });
  ctx.waitUntil((async () => {
    try {
      const data = await fetchOrderDataFresh(env, orderId);
      await env.GCX_KV.put('pf:' + orderId, JSON.stringify({ ts: Date.now(), data }), { expirationTtl: PREFETCH_TTL_SEC });
    } catch (e) { console.log('prefetch failed', orderId, e.message); }
  })());
  return json({ ok: true });
}

// ── Autocomplete telemetry (TM v3.9.1) ──────────────────────────────────────
// The panel POSTs small batches of de-identified autocomplete events (see the
// "Telemetry" block in GCX Reply.user.js). Stored per day / agent / tab in KV
// (`aclog:<day>:<agentId>:<tab>`, 60-day TTL). Read back with
//   GET /ac/log?days=7            → aggregate summary
//   GET /ac/log?days=7&raw=1      → every event
// (x-push-key: PIDX_PUSH_KEY). Nothing here is on the panel's lookup path.
const ACLOG_TTL_SEC = 60 * 86400;
const ACLOG_MAX_PER_KEY = 6000;
const ACLOG_STR_KEYS = new Set(['e', 'k', 'lv', 'o', 'sg', 'at', 'v', 'os']);
const ACLOG_NUM_KEYS = new Set(['t', 'src', 'sl', 'tl', 'ms', 'tt', 'ac', 'w', 'len', 'keys', 'nosug', 'ep', 'acc', 'dur', 'c', 'hist', 'macro']);
function sanitizeAcEvent(e) {
  if (!e || typeof e !== 'object') return null;
  const out = {};
  for (const k of Object.keys(e)) {
    if (ACLOG_STR_KEYS.has(k) && typeof e[k] === 'string') out[k] = e[k].slice(0, 200);
    else if (ACLOG_NUM_KEYS.has(k) && typeof e[k] === 'number' && isFinite(e[k])) out[k] = e[k];
  }
  if (!out.e || !out.t) return null;
  return out;
}
async function handleAcLogWrite(request, env, ctx) {
  if (!safeEq(request.headers.get('x-gcx-key') || '', env.GCX_KEY)) return json({ error: 'unauthorized' }, 401);
  let body;
  try { body = await request.json(); } catch { return json({ error: 'bad json' }, 400); }
  const agent = String(body.agent || 'anon').replace(/[^\w-]/g, '').slice(0, 24) || 'anon';
  const tab = String(body.tab || '0').replace(/[^\w-]/g, '').slice(0, 16) || '0';
  const ver = String(body.ver || '').slice(0, 16);
  const events = (Array.isArray(body.events) ? body.events : []).slice(0, 500).map(sanitizeAcEvent).filter(Boolean)
    .map(e => Object.assign(e, { a: agent, v: e.v || ver }));
  if (!events.length) return json({ ok: true, stored: 0 });
  const day = new Date().toISOString().slice(0, 10);
  const key = `aclog:${day}:${agent}:${tab}`;
  ctx.waitUntil((async () => {
    let existing = [];
    try { existing = (await env.GCX_KV.get(key, 'json')) || []; } catch {}
    if (!Array.isArray(existing)) existing = [];
    const merged = existing.concat(events);
    if (merged.length > ACLOG_MAX_PER_KEY) merged.splice(0, merged.length - ACLOG_MAX_PER_KEY);
    await env.GCX_KV.put(key, JSON.stringify(merged), { expirationTtl: ACLOG_TTL_SEC });
  })());
  return json({ ok: true, stored: events.length });
}
function pct(arr, p) {
  if (!arr.length) return null;
  const s = arr.slice().sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.floor(p * s.length))];
}
function summarizeAcLog(events) {
  const eps = events.filter(e => e.e === 's');
  const msgs = events.filter(e => e.e === 'm');
  const hbs = events.filter(e => e.e === 'h');
  const ACCEPT = o => o === 'tab' || o === 'word';
  const bucket = () => ({ episodes: 0, tab: 0, word: 0, esc: 0, typed: 0, ign: 0, other: 0, acceptedChars: 0, typedThroughChars: 0, ms: [] });
  const fill = (b, e) => {
    b.episodes++;
    if (e.o === 'tab') b.tab++; else if (e.o === 'word' || (e.w > 0 && e.o !== 'esc')) b.word++; else if (e.o === 'esc') b.esc++; else if (e.o === 'typed') b.typed++; else if (e.o === 'ign' || e.o === 'swap') b.ign++; else b.other++;
    b.acceptedChars += e.ac || 0; b.typedThroughChars += e.tt || 0; if (typeof e.ms === 'number') b.ms.push(e.ms);
  };
  const finish = b => ({ episodes: b.episodes, tab: b.tab, word: b.word, esc: b.esc, typed: b.typed, ign: b.ign, other: b.other,
    acceptRate: b.episodes ? +((b.tab + b.word) / b.episodes).toFixed(3) : null,
    usefulRate: b.episodes ? +((b.tab + b.word + b.typed) / b.episodes).toFixed(3) : null,
    acceptedChars: b.acceptedChars, typedThroughChars: b.typedThroughChars,
    msP50: pct(b.ms, 0.5), msP95: pct(b.ms, 0.95), msMax: b.ms.length ? Math.max(...b.ms) : null });
  const group = (keyFn) => { const m = {}; for (const e of eps) { const k = keyFn(e) ?? 'na'; (m[k] = m[k] || bucket()); fill(m[k], e); } const o = {}; for (const k in m) o[k] = finish(m[k]); return o; };
  const total = bucket(); eps.forEach(e => fill(total, e));
  const top = (filter, n) => { const m = {}; for (const e of eps) { if (!filter(e) || !e.sg) continue; const k = e.sg; m[k] = m[k] || { sg: k, n: 0, at: [] }; m[k].n++; if (e.at && m[k].at.length < 3 && !m[k].at.includes(e.at)) m[k].at.push(e.at); } return Object.values(m).sort((a, b) => b.n - a.n).slice(0, n); };
  const msgAgg = { messages: msgs.length, chars: 0, acceptedChars: 0, keystrokes: 0, nosug: 0, withAccept: 0 };
  for (const m of msgs) { msgAgg.chars += m.len || 0; msgAgg.acceptedChars += m.acc || 0; msgAgg.keystrokes += m.keys || 0; msgAgg.nosug += m.nosug || 0; if (m.acc > 0) msgAgg.withAccept++; }
  msgAgg.acceptedShare = msgAgg.chars ? +(msgAgg.acceptedChars / msgAgg.chars).toFixed(3) : null;
  msgAgg.coverage = msgAgg.keystrokes ? +(1 - msgAgg.nosug / msgAgg.keystrokes).toFixed(3) : null;
  const agents = {};
  for (const h of hbs) { const a = agents[h.a] = agents[h.a] || { loads: 0 }; a.loads++; if (!a.last || h.t > a.last) { a.last = h.t; a.version = h.v; a.corpus = h.c; a.hist = h.hist; a.macro = h.macro; a.os = h.os; } }
  return {
    events: events.length, agentsSeen: Object.keys(agents).length,
    overall: finish(total), byAgent: group(e => e.a), byLevel: group(e => e.lv), byKind: group(e => e.k), bySource: group(e => e.src),
    messages: msgAgg, messagesByKind: { p: msgs.filter(m => m.k === 'p').length, n: msgs.filter(m => m.k === 'n').length },
    topAccepted: top(e => ACCEPT(e.o) || e.o === 'typed', 15),
    topIgnored: top(e => e.o === 'ign' || e.o === 'swap' || e.o === 'esc', 25),
    agents,
  };
}
async function handleAcLogRead(request, env, url) {
  if (!safeEq(request.headers.get('x-push-key') || '', env.PIDX_PUSH_KEY)) return json({ error: 'unauthorized' }, 401);
  const days = Math.min(60, Math.max(1, parseInt(url.searchParams.get('days') || '7', 10) || 7));
  const since = Date.now() - days * 86400000;
  const names = [];
  let cursor;
  do {
    const r = await env.GCX_KV.list({ prefix: 'aclog:', cursor });
    for (const k of r.keys) { const day = Date.parse(k.name.split(':')[1]); if (day >= since - 86400000) names.push(k.name); }
    cursor = r.list_complete ? null : r.cursor;
  } while (cursor);
  const all = [];
  for (let i = 0; i < names.length; i += 10) {
    const chunk = await Promise.all(names.slice(i, i + 10).map(n => env.GCX_KV.get(n, 'json').catch(() => null)));
    for (const arr of chunk) if (Array.isArray(arr)) all.push(...arr);
  }
  const events = all.filter(e => e && e.t >= since).sort((a, b) => a.t - b.t);
  if (url.searchParams.get('raw') === '1') return json({ days, keys: names.length, events });
  return json(Object.assign({ days, keys: names.length }, summarizeAcLog(events)));
}
// Shared loader for the read/view handlers.
async function loadAcLogEvents(env, days) {
  const since = Date.now() - days * 86400000;
  const names = [];
  let cursor;
  do {
    const r = await env.GCX_KV.list({ prefix: 'aclog:', cursor });
    for (const k of r.keys) { const day = Date.parse(k.name.split(':')[1]); if (day >= since - 86400000) names.push(k.name); }
    cursor = r.list_complete ? null : r.cursor;
  } while (cursor);
  const all = [];
  for (let i = 0; i < names.length; i += 10) {
    const chunk = await Promise.all(names.slice(i, i + 10).map(n => env.GCX_KV.get(n, 'json').catch(() => null)));
    for (const arr of chunk) if (Array.isArray(arr)) all.push(...arr);
  }
  return { names, events: all.filter(e => e && e.t >= since).sort((a, b) => a.t - b.t) };
}
// Admin: DELETE /ac/log?agent=<id> (x-push-key) removes that agent's batches (test data cleanup).
async function handleAcLogPurge(request, env, url) {
  if (!safeEq(request.headers.get('x-push-key') || '', env.PIDX_PUSH_KEY)) return json({ error: 'unauthorized' }, 401);
  const agent = String(url.searchParams.get('agent') || '').replace(/[^\w-]/g, '');
  if (!agent) return json({ error: 'agent required' }, 400);
  const names = [];
  let cursor;
  do {
    const r = await env.GCX_KV.list({ prefix: 'aclog:', cursor });
    for (const k of r.keys) if (k.name.split(':')[2] === agent) names.push(k.name);
    cursor = r.list_complete ? null : r.cursor;
  } while (cursor);
  await Promise.all(names.map(n => env.GCX_KV.delete(n)));
  return json({ ok: true, deleted: names });
}
// Human-readable page: GET /ac/log/view?k=<ACLOG_VIEW_KEY>&days=7 — read-only, de-identified data only.
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function renderAcLogPage(days, sum) {
  const pctS = v => v == null ? '–' : (v * 100).toFixed(0) + '%';
  const outcomeCols = ['episodes', 'tab', 'word', 'typed', 'esc', 'ign', 'other', 'acceptRate', 'usefulRate', 'acceptedChars', 'msP50', 'msP95'];
  const table = (title, rows, labelMap) => {
    const keys = Object.keys(rows);
    if (!keys.length) return `<h3>${esc(title)}</h3><p class="muted">데이터 없음</p>`;
    let h = `<h3>${esc(title)}</h3><table><thead><tr><th></th>${outcomeCols.map(c => `<th>${c}</th>`).join('')}</tr></thead><tbody>`;
    for (const k of keys) {
      const r = rows[k];
      h += `<tr><td class="lbl">${esc(labelMap ? (labelMap[k] || k) : k)}</td>` + outcomeCols.map(c => `<td>${c.endsWith('Rate') ? pctS(r[c]) : esc(r[c] == null ? '–' : r[c])}</td>`).join('') + '</tr>';
    }
    return h + '</tbody></table>';
  };
  const list = (title, items) => `<h3>${esc(title)}</h3>` + (items.length ? '<ol>' + items.map(i => `<li><span class="n">${i.n}×</span> <code>${esc(i.sg)}</code>${i.at && i.at.length ? `<div class="at">대신 입력: ${i.at.map(a => `<code>${esc(a)}</code>`).join(' · ')}</div>` : ''}</li>`).join('') + '</ol>' : '<p class="muted">데이터 없음</p>');
  const o = sum.overall, m = sum.messages;
  const agentsRows = Object.entries(sum.agents || {}).map(([id, a]) => `<tr><td>${esc(id)}</td><td>${esc(a.version)}</td><td>${a.loads}</td><td>${a.corpus} (내 답변 ${a.hist} / 매크로 ${a.macro})</td><td>${esc(a.os)}</td><td>${new Date(a.last).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })}</td></tr>`).join('');
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>GCX Reply 자동완성 로그</title>
<style>body{font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;margin:0;padding:24px;background:#f6f7fb;color:#1c1c1e}h1{font-size:20px;margin:0 0 4px}h3{margin:28px 0 8px;font-size:15px}.muted{color:#888}.cards{display:flex;flex-wrap:wrap;gap:12px;margin:16px 0}.card{background:#fff;border-radius:12px;padding:12px 16px;min-width:150px;box-shadow:0 1px 3px rgba(0,0,0,.06)}.card b{display:block;font-size:22px}.card span{color:#666;font-size:12px}table{border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.06);font-size:13px}th,td{padding:6px 10px;text-align:right;border-bottom:1px solid #eef}th{background:#eef0f7;font-weight:600}td.lbl,th:first-child{text-align:left}ol{padding-left:22px}li{margin:4px 0}code{background:#fff;border:1px solid #e3e5ee;border-radius:5px;padding:1px 5px;font-size:12.5px}.n{color:#5b6dc8;font-weight:600}.at{color:#777;font-size:12px;margin-left:4px}nav a{margin-right:10px}.help{background:#fff;border-radius:10px;padding:10px 14px;font-size:12.5px;color:#555;margin-top:8px}</style></head><body>
<h1>GCX Reply 자동완성 로그</h1><div class="muted">최근 ${days}일 · 이벤트 ${sum.events}개 · 상담원 ${sum.agentsSeen}명 · ${new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })} 기준</div>
<nav>기간: <a href="?days=1">1일</a><a href="?days=7">7일</a><a href="?days=14">14일</a><a href="?days=30">30일</a><a href="?days=60">60일</a></nav>
<div class="cards"><div class="card"><b>${o.episodes}</b><span>제안 표시 횟수</span></div><div class="card"><b>${pctS(o.acceptRate)}</b><span>채택률 (Tab/단어)</span></div><div class="card"><b>${pctS(o.usefulRate)}</b><span>유효율 (채택 + 그대로 타이핑)</span></div><div class="card"><b>${o.acceptedChars}</b><span>Tab으로 입력된 글자 수</span></div><div class="card"><b>${m.messages}</b><span>전송된 답변/메모</span></div><div class="card"><b>${pctS(m.acceptedShare)}</b><span>전송 글자 중 자동완성 비율</span></div><div class="card"><b>${pctS(m.coverage)}</b><span>제안이 있던 입력 비율</span></div><div class="card"><b>${o.msP95 == null ? '–' : o.msP95 + ' ms'}</b><span>계산 시간 p95</span></div></div>
<div class="help">tab = Tab으로 전체 채택 · word = Alt+→ 단어 채택 · typed = 제안과 같은 내용을 직접 타이핑 · esc = Esc로 닫음 · ign = 다른 내용을 입력(무시) · other = 탭 이동/전송/포커스 아웃 등. a = 문단 매칭, b = 문장 매칭, c = 단어 예측.</div>
${table('상담원별', sum.byAgent)}${table('매칭 방식별', sum.byLevel, { a: 'a 문단', b: 'b 문장', c: 'c 단어 예측' })}${table('채널별', sum.byKind, { p: '공개 답변', n: '내부 메모' })}${table('학습 출처별', sum.bySource, { 0: '직접 보낸 답변', 1: '과거 티켓', 2: '매크로', '-1': '단어 예측' })}
${list('가장 많이 채택된 제안', sum.topAccepted || [])}${list('가장 많이 무시된 제안', sum.topIgnored || [])}
<h3>상담원 (최근 접속)</h3><table><thead><tr><th>ID</th><th>버전</th><th>페이지 로드</th><th>학습 문장</th><th>OS</th><th>마지막</th></tr></thead><tbody>${agentsRows || '<tr><td colspan="6" class="muted">없음</td></tr>'}</tbody></table>
</body></html>`;
}
async function handleAcLogView(request, env, url) {
  const k = url.searchParams.get('k') || '';
  if (!env.ACLOG_VIEW_KEY || !safeEq(k, env.ACLOG_VIEW_KEY)) return new Response('unauthorized', { status: 401 });
  const days = Math.min(60, Math.max(1, parseInt(url.searchParams.get('days') || '7', 10) || 7));
  const { events } = await loadAcLogEvents(env, days);
  const html = renderAcLogPage(days, summarizeAcLog(events));
  return new Response(html, { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'X-Robots-Tag': 'noindex' } });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === '/health') return json({ ok: true });
    if (request.method === 'GET' && url.pathname === '/') return handleRead(request, env, ctx, url);
    if (request.method === 'POST' && url.pathname === '/admin/pidx') return handlePidxPush(request, env);
    if (request.method === 'POST' && url.pathname === '/prefetch') return handlePrefetch(request, env, ctx);
    if (request.method === 'POST' && url.pathname === '/ac/log') return handleAcLogWrite(request, env, ctx);
    if (request.method === 'GET' && url.pathname === '/ac/log') return handleAcLogRead(request, env, url);
    if (request.method === 'DELETE' && url.pathname === '/ac/log') return handleAcLogPurge(request, env, url);
    if (request.method === 'GET' && url.pathname === '/ac/log/view') return handleAcLogView(request, env, url);
    return json({ error: 'not found' }, 404);
  },
};
