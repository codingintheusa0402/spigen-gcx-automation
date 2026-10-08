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

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === '/health') return json({ ok: true });
    if (request.method === 'GET' && url.pathname === '/') return handleRead(request, env, ctx, url);
    if (request.method === 'POST' && url.pathname === '/admin/pidx') return handlePidxPush(request, env);
    if (request.method === 'POST' && url.pathname === '/prefetch') return handlePrefetch(request, env, ctx);
    return json({ error: 'not found' }, 404);
  },
};
