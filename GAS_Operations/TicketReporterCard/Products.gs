/**
 * `/products <검색어>` and `/products-similar <검색어>` — product lookup cards.
 *
 * /products          SKU master + ASIN master match (SKU, ASIN or keywords such as
 *                    "iPhone 17 Pro 울트라 하이브리드"), plus FBA inventory per marketplace
 *                    (SP-API getInventorySummaries) and an Amazon thumbnail (Catalog Items API).
 * /products-similar  the first match's device (SKU master 기종 / ASIN master 기종명), then every
 *                    other SKU made for that same device — "compatible with the same phone".
 *
 * SP-API credentials are copies of the MCF_Tracking project's Script Properties (사용자 지시
 * 2026-10-17: use only those) under the same names — LWA_CLIENT_ID / LWA_CLIENT_SECRET /
 * LWA_REFRESH_TOKEN for NA + EU, LWA_*_JP for FE. SigV4 is not needed (Amazon dropped it).
 *
 * A Chat add-on reply has to come back within ~30 s, so: the master index is cached (6 h),
 * each region's marketplace participation is cached (6 h), and every SP-API request of one
 * kind goes out together through UrlFetchApp.fetchAll with one throttled retry for 429s.
 */

var SKU_MASTER_ID = '1JijzoYw9aDW-9Jx_OJOBByMVCX8-0OtNxb-0Y82EZJI';   // SKU_Master(먼데이보드) › Data
var ASIN_MASTER_ID = '1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo';  // ASIN_Master(먼데이보드) › Data
var PRODUCTS_MAX = 8;            // product rows per card (/products)
var SIMILAR_MAX = 12;            // product rows per card (/products-similar)
var PRODUCT_INDEX_CACHE = 'PRODUCT_INDEX_V1';

var SPAPI_REGIONS = {
  NA: { host: 'sellingpartnerapi-na.amazon.com', creds: '' },
  EU: { host: 'sellingpartnerapi-eu.amazon.com', creds: '' },
  FE: { host: 'sellingpartnerapi-fe.amazon.com', creds: '_JP' }
};

// Marketplace id → [code, region, amazon domain]. Only the ones the seller participates in are queried.
var MARKETPLACES = {
  ATVPDKIKX0DER: ['US', 'NA', 'amazon.com'],
  A2EUQ1WTGCTBG2: ['CA', 'NA', 'amazon.ca'],
  A1AM78C64UM0Y8: ['MX', 'NA', 'amazon.com.mx'],
  A2Q3Y263D00KWC: ['BR', 'NA', 'amazon.com.br'],
  A1F83G8C2ARO7P: ['UK', 'EU', 'amazon.co.uk'],
  A1PA6795UKMFR9: ['DE', 'EU', 'amazon.de'],
  A13V1IB3VIYZZH: ['FR', 'EU', 'amazon.fr'],
  APJ6JRA9NG5V4: ['IT', 'EU', 'amazon.it'],
  A1RKKUPIHCS9HS: ['ES', 'EU', 'amazon.es'],
  A1805IZSGTT6HS: ['NL', 'EU', 'amazon.nl'],
  A2NODRKZP88ZB9: ['SE', 'EU', 'amazon.se'],
  A1C3SOZRARQ6R3: ['PL', 'EU', 'amazon.pl'],
  AMEN7PMS3EDWL: ['BE', 'EU', 'amazon.com.be'],
  A28R8C7NBKEWEA: ['IE', 'EU', 'amazon.ie'],
  A33AVAJ2PDY3EV: ['TR', 'EU', 'amazon.com.tr'],
  A2VIGQ35RCS4UG: ['AE', 'EU', 'amazon.ae'],
  A17E79C6D8DWNP: ['SA', 'EU', 'amazon.sa'],
  ARBP9OOSHTCHU: ['EG', 'EU', 'amazon.eg'],
  A21TJRUUN4KGV: ['IN', 'EU', 'amazon.in'],
  A1VC38T7YXB528: ['JP', 'FE', 'amazon.co.jp'],
  A39IBJ37TRP1C6: ['AU', 'FE', 'amazon.com.au'],
  A19VAU5U5O7RUS: ['SG', 'FE', 'amazon.sg']
};
var CATALOG_ORDER = ['ATVPDKIKX0DER', 'A1PA6795UKMFR9', 'A1F83G8C2ARO7P', 'A1VC38T7YXB528'];  // thumbnail lookup order

/* ===================== Command handlers ===================== */

function handleProductsCommand_(query, similar) {
  query = String(query || '').trim();
  var cmd = similar ? '/products-similar' : '/products';
  if (!query) {
    return chatCreate_({ text: '⚠️ ' + cmd + ' 뒤에 SKU, ASIN 또는 제품명을 입력해 주세요. 예: "' + cmd + ' ACS08008", "' + cmd + ' iPhone 17 Pro 울트라 하이브리드"' });
  }
  var t0 = Date.now();
  var index = loadProductIndex_();
  var matches = searchProducts_(index, query);
  var title, subtitle, list, total;

  if (!similar) {
    total = matches.length;
    list = matches.slice(0, PRODUCTS_MAX);
    if (!list.length && /^[A-Z]{3}\d{5}$/i.test(query)) list = [{ sku: query.toUpperCase(), asins: [] }];  // not in the masters — still ask SP-API
    if (!list.length && /^B0[A-Z0-9]{8}$/i.test(query)) list = [{ sku: '', asins: [query.toUpperCase()] }];
    title = '🔎 ' + query;
    subtitle = '마스터 일치 ' + total + '건' + (total > list.length ? ' · 상위 ' + list.length + '건 표시' : '');
  } else {
    var base = matches.filter(function (p) { return isRealDevice_(p.device); })[0];
    if (!base) {
      return chatCreate_({ text: '⚠️ "' + query + '" 에서 기종(호환 기기)을 찾지 못했습니다. SKU나 ASIN으로 다시 시도해 주세요.' });
    }
    var devs = deviceSet_(base.device);
    var same = index.filter(function (p) {
      return p.sku !== base.sku && deviceSet_(p.device).some(function (d) { return devs.indexOf(d) >= 0; });
    });
    same.sort(function (a, b) {
      return (a.status === 'B') - (b.status === 'B') ||
             String(a.category).localeCompare(String(b.category)) ||
             String(a.model).localeCompare(String(b.model)) || a.sku.localeCompare(b.sku);
    });
    total = same.length;
    list = [base].concat(same.slice(0, SIMILAR_MAX - 1));
    title = '📱 ' + base.device + ' 호환 제품';
    subtitle = '기준 ' + base.sku + ' · 같은 기종 ' + total + '건' + (total > list.length - 1 ? ' · 상위 ' + (list.length - 1) + '건 표시' : '');
  }

  if (!list.length) {
    return chatCreate_({ text: '⚠️ "' + query + '" 와 일치하는 제품이 SKU master / ASIN master 에 없습니다.' });
  }

  var notes = [];
  var inv = {}, thumbs = {};
  try { inv = fetchInventory_(list.map(function (p) { return p.sku; }).filter(String), notes); }
  catch (e) { notes.push('재고 조회 실패: ' + e.message); }
  list.forEach(function (p) {   // ASINs seen in FBA inventory fill in products the ASIN master doesn't know
    var seen = (inv[p.sku] && inv[p.sku].asins) || [];
    seen.forEach(function (a) { if (p.asins.indexOf(a) < 0) p.asins.push(a); });
  });
  try { thumbs = fetchThumbnails_(list, notes); }
  catch (e) { notes.push('이미지 조회 실패: ' + e.message); }

  var card = buildProductsCard_(title, subtitle, list, inv, thumbs, similar, notes, Date.now() - t0);
  return chatCreate_({ cardsV2: [card] });
}

/* ===================== Master index ===================== */

/** [{sku, asins[], model, device, nameKo, nameEn, brand, color, category, status}] from both masters, cached 6 h. */
function loadProductIndex_() {
  var cached = cacheGetChunked_(PRODUCT_INDEX_CACHE);
  if (cached) return JSON.parse(cached);

  var bySku = {};
  var rows = readColumns_(SKU_MASTER_ID, ['name', '모델명', '자재내역(KO)', '자재내역(EN)', '브랜드', '기종', '색상', '대분류', '품목상태']);
  rows.forEach(function (r) {
    var sku = String(r[0] || '').trim().toUpperCase();
    if (!sku) return;
    bySku[sku] = { sku: sku, asins: [], model: r[1], nameKo: r[2], nameEn: r[3], brand: r[4], device: r[5], color: r[6], category: r[7], status: r[8] };
  });
  rows = readColumns_(ASIN_MASTER_ID, ['SKU', 'ASIN', '모델명', '자재내역', '자재내역(EN)', '브랜드', '기종명', '색상명', '대분류']);
  rows.forEach(function (r) {
    var sku = String(r[0] || '').trim().toUpperCase();
    var asin = String(r[1] || '').trim().toUpperCase();
    if (!sku) return;
    var p = bySku[sku];
    if (!p) p = bySku[sku] = { sku: sku, asins: [], model: r[2], nameKo: r[3], nameEn: r[4], brand: r[5], device: r[6], color: r[7], category: r[8], status: '' };
    if (asin && p.asins.indexOf(asin) < 0) p.asins.push(asin);
    if (!isRealDevice_(p.device) && isRealDevice_(r[6])) p.device = r[6];
  });

  var index = Object.keys(bySku).map(function (k) {
    var p = bySku[k];
    ['model', 'nameKo', 'nameEn', 'brand', 'device', 'color', 'category', 'status'].forEach(function (f) { p[f] = String(p[f] || '').trim(); });
    return p;
  });
  cachePutChunked_(PRODUCT_INDEX_CACHE, JSON.stringify(index), 6 * 3600);
  return index;
}

/** Reads the named header columns of a spreadsheet's Data tab (row 1 = header) → rows of values in that order. */
function readColumns_(spreadsheetId, headers) {
  var sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Data');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var cols = headers.map(function (h) {
    var i = head.indexOf(h);
    return i < 0 ? null : sh.getRange(2, i + 1, last - 1, 1).getValues();
  });
  var out = [];
  for (var r = 0; r < last - 1; r++) {
    out.push(cols.map(function (c) { return c ? c[r][0] : ''; }));
  }
  return out;
}

/** SKU / ASIN exact match first; otherwise every keyword must appear somewhere in the product's text. */
function searchProducts_(index, query) {
  var q = query.trim().toUpperCase();
  var exact = index.filter(function (p) { return p.sku === q || p.asins.indexOf(q) >= 0; });
  if (exact.length) return exact;

  var terms = query.toLowerCase().split(/\s+/).filter(String);
  var hits = index.filter(function (p) {
    var hay = [p.sku, p.asins.join(' '), p.model, p.nameKo, p.nameEn, p.brand, p.device, p.color, p.category].join(' ').toLowerCase();
    return terms.every(function (t) { return hay.indexOf(t) >= 0; });
  });
  hits.sort(function (a, b) {   // 품목상태 A before B, ones with an ASIN, then SKU
    return (a.status === 'B') - (b.status === 'B') || (b.asins.length > 0) - (a.asins.length > 0) || a.sku.localeCompare(b.sku);
  });
  return hits;
}

function isRealDevice_(d) {
  d = String(d || '').trim();
  return !!d && !/^(common|공용|미정|-|n\/a)$/i.test(d);
}

/**
 * Every device a 기종 value fits, normalised: "iPhone 18 Pro Max/17 Pro Max" → ["iphone 18 pro max",
 * "iphone 17 pro max"] (a later part borrows the first part's first word),
 * "Model Y/3" → ["model y", "model 3"], "#iPad Air 11" (2024)/" → ["ipad air 11" (2024)"].
 */
function deviceSet_(d) {
  var parts = String(d || '').replace(/#/g, '').split('/').map(function (x) { return x.replace(/\s+/g, ' ').trim(); }).filter(String);
  if (!parts.length || !isRealDevice_(parts[0])) return [];
  var family = parts[0].split(' ')[0];   // iPhone / Galaxy / Model / Pixel …
  return parts.map(function (x, i) {
    return (i && x.split(' ')[0].toLowerCase() !== family.toLowerCase() ? family + ' ' + x : x).toLowerCase();
  });
}

/* ===================== SP-API ===================== */

function spapiToken_(region) {
  var suffix = SPAPI_REGIONS[region].creds;
  var cache = CacheService.getScriptCache();
  var key = 'SPAPI_LWA' + suffix;
  var hit = cache.get(key);
  if (hit) return hit;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LWA_CLIENT_ID' + suffix), secret = props.getProperty('LWA_CLIENT_SECRET' + suffix),
      refresh = props.getProperty('LWA_REFRESH_TOKEN' + suffix);
  if (!id || !secret || !refresh) throw new Error('Script Properties에 LWA_*' + suffix + ' 값이 없습니다');
  var res = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'refresh_token', refresh_token: refresh, client_id: id, client_secret: secret }
  });
  if (res.getResponseCode() !== 200) throw new Error('LWA ' + region + ' ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 150));
  var tok = JSON.parse(res.getContentText()).access_token;
  cache.put(key, tok, 3300);
  return tok;
}

/** Runs [{region, path, query}] in parallel; 429/5xx get one retry ~1 s later. → parsed JSON or {error}. */
function spapiFetchAll_(reqs) {
  var tokens = {};
  var build = function (r) {
    if (!tokens[r.region]) tokens[r.region] = spapiToken_(r.region);
    return {
      url: 'https://' + SPAPI_REGIONS[r.region].host + r.path + (r.query ? '?' + r.query : ''),
      method: 'get', muteHttpExceptions: true,
      headers: { 'x-amz-access-token': tokens[r.region], 'user-agent': 'GCX-TicketReporter/1.0 (Language=AppsScript)' }
    };
  };
  var out = new Array(reqs.length);
  var pending = reqs.map(function (r, i) { return i; });
  for (var attempt = 0; attempt < 3 && pending.length; attempt++) {
    if (attempt) Utilities.sleep(1100 * attempt);
    var resps = UrlFetchApp.fetchAll(pending.map(function (i) { return build(reqs[i]); }));
    var retry = [];
    resps.forEach(function (res, k) {
      var i = pending[k], code = res.getResponseCode();
      if ((code === 429 || code >= 500) && attempt < 2) { retry.push(i); return; }
      var body = {};
      try { body = JSON.parse(res.getContentText()); } catch (e) {}
      out[i] = code === 200 ? body : { error: code + ' ' + ((body.errors && body.errors[0] && body.errors[0].message) || res.getContentText().slice(0, 120)) };
    });
    pending = retry;
  }
  return out;
}

/** Marketplace ids this seller sells in, per region, cached 6 h. Regions whose token fails are skipped (noted). */
function activeMarketplaces_(notes) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('SPAPI_MKTS_V1');
  if (hit) return JSON.parse(hit);
  var regions = Object.keys(SPAPI_REGIONS), reqs = [], used = [];
  regions.forEach(function (rg) {
    try { spapiToken_(rg); reqs.push({ region: rg, path: '/sellers/v1/marketplaceParticipations' }); used.push(rg); }
    catch (e) { notes.push(rg + ' 리전 인증 실패 — ' + e.message); }
  });
  var res = spapiFetchAll_(reqs), ids = [];
  res.forEach(function (r, i) {
    if (r.error) { notes.push(used[i] + ' 마켓 목록 조회 실패 — ' + r.error); return; }
    (r.payload || []).forEach(function (mp) {
      var id = mp.marketplace && mp.marketplace.id;
      if (MARKETPLACES[id] && MARKETPLACES[id][1] === used[i] && mp.participation && mp.participation.isParticipating && ids.indexOf(id) < 0) ids.push(id);
    });
  });
  if (ids.length) cache.put('SPAPI_MKTS_V1', JSON.stringify(ids), 6 * 3600);
  return ids;
}

/**
 * FBA inventory for up to 50 seller SKUs in every active marketplace.
 * → { SKU: { asins: [...], byMkt: { US: {f: fulfillable, t: total, in: inbound}, ... } } }
 */
function fetchInventory_(skus, notes) {
  var out = {};
  if (!skus.length) return out;
  var mkts = activeMarketplaces_(notes);
  var reqs = mkts.map(function (id) {
    return {
      region: MARKETPLACES[id][1], mkt: id,
      path: '/fba/inventory/v1/summaries',
      query: 'details=true&granularityType=Marketplace&granularityId=' + id + '&marketplaceIds=' + id +
             '&sellerSkus=' + encodeURIComponent(skus.slice(0, 50).join(','))
    };
  });
  var failed = [];
  spapiFetchAll_(reqs).forEach(function (r, i) {
    var code = MARKETPLACES[reqs[i].mkt][0];
    if (r.error) { failed.push(code); return; }
    ((r.payload && r.payload.inventorySummaries) || []).forEach(function (s) {
      var sku = String(s.sellerSku || '').toUpperCase();
      var o = out[sku] = out[sku] || { asins: [], byMkt: {} };
      if (s.asin && o.asins.indexOf(s.asin) < 0) o.asins.push(s.asin);
      var d = s.inventoryDetails || {};
      var inbound = (d.inboundWorkingQuantity || 0) + (d.inboundShippedQuantity || 0) + (d.inboundReceivingQuantity || 0);
      o.byMkt[code] = { f: d.fulfillableQuantity || 0, t: s.totalQuantity || 0, in: inbound };
    });
  });
  if (failed.length) notes.push('재고 조회 실패 마켓: ' + failed.join(', '));
  return out;
}

/** ASIN → { img, mkt } using Catalog Items searchCatalogItems (20 ASINs per call), US → DE → UK → JP. */
function fetchThumbnails_(list, notes) {
  var want = [];
  list.forEach(function (p) { if (p.asins[0] && want.indexOf(p.asins[0]) < 0) want.push(p.asins[0]); });
  var found = {};
  var active = activeMarketplaces_(notes);
  for (var m = 0; m < CATALOG_ORDER.length && want.length; m++) {
    var id = CATALOG_ORDER[m];
    if (active.indexOf(id) < 0) continue;
    var chunks = [];
    for (var c = 0; c < want.length; c += 20) chunks.push(want.slice(c, c + 20));
    var res = spapiFetchAll_(chunks.map(function (ch) {
      return {
        region: MARKETPLACES[id][1], path: '/catalog/2022-04-01/items',
        query: 'identifiers=' + ch.join(',') + '&identifiersType=ASIN&marketplaceIds=' + id + '&includedData=images,summaries&pageSize=20'
      };
    }));
    res.forEach(function (r) {
      (r.items || []).forEach(function (it) {
        var img = pickThumb_(it.images);
        if (img && !found[it.asin]) found[it.asin] = { img: img, mkt: id };
      });
    });
    want = want.filter(function (a) { return !found[a]; });
  }
  return found;
}

/** Smallest MAIN image that's still ≥ 150 px wide. */
function pickThumb_(imagesByMkt) {
  var imgs = [];
  (imagesByMkt || []).forEach(function (g) { imgs = imgs.concat(g.images || []); });
  var main = imgs.filter(function (i) { return i.variant === 'MAIN'; });
  if (!main.length) main = imgs;
  main.sort(function (a, b) { return a.width - b.width; });
  var ok = main.filter(function (i) { return i.width >= 150; });
  return (ok[0] || main[main.length - 1] || {}).link || '';
}

/* ===================== Card ===================== */

function buildProductsCard_(title, subtitle, list, inv, thumbs, similar, notes, ms) {
  var sections = list.map(function (p, i) {
    var asin = p.asins[0] || '';
    var th = asin && thumbs[asin];
    var domain = th ? MARKETPLACES[th.mkt][2] : 'amazon.com';
    var lines = [];
    var name = [p.brand, p.model, p.color].filter(String).join(' · ');
    if (name) lines.push('<b>' + escapeHtml_(name) + '</b>');
    if (p.nameKo) lines.push(escapeHtml_(p.nameKo));
    var meta = [];
    if (p.device) meta.push('📱 ' + escapeHtml_(p.device));
    if (p.category) meta.push(escapeHtml_(p.category));
    if (meta.length) lines.push('<font color="#5f6368">' + meta.join(' · ') + '</font>');
    if (p.asins.length) {
      lines.push('ASIN ' + p.asins.slice(0, 3).map(function (a) {
        return '<a href="https://www.' + domain + '/dp/' + a + '">' + a + '</a>';
      }).join(', ') + (p.asins.length > 3 ? ' 외 ' + (p.asins.length - 3) : ''));
    }
    lines.push(inventoryLine_(inv[p.sku]));

    var textCol = { textParagraph: { text: lines.join('<br>') } };
    var widget = th ? {
      columns: { columnItems: [
        { horizontalSizeStyle: 'FILL_MINIMUM_SPACE', horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
          widgets: [{ image: { imageUrl: th.img, altText: p.sku, onClick: { openLink: { url: 'https://www.' + domain + '/dp/' + asin } } } }] },
        { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', horizontalAlignment: 'START', verticalAlignment: 'TOP',
          widgets: [textCol] }
      ] }
    } : textCol;

    var header = (similar && i === 0 ? '⭐ 기준 · ' : '') + (p.sku || asin);
    return { header: header, widgets: [widget] };
  });

  var footer = [];
  if (notes.length) footer.push('⚠️ ' + notes.map(escapeHtml_).join('<br>⚠️ '));
  footer.push('<font color="#5f6368">재고 = FBA 판매가능 수량 (괄호: 입고중) · 출처 SKU master / ASIN master / SP-API · ' + (ms / 1000).toFixed(1) + '초</font>');
  footer.push('<a href="https://docs.google.com/spreadsheets/d/' + SKU_MASTER_ID + '/edit">SKU master</a> · ' +
              '<a href="https://docs.google.com/spreadsheets/d/' + ASIN_MASTER_ID + '/edit">ASIN master</a>');
  sections.push({ widgets: [{ textParagraph: { text: footer.join('<br>') } }] });

  return {
    cardId: 'products-' + new Date().getTime(),
    card: { header: { title: title, subtitle: subtitle }, sections: sections }
  };
}

/** "재고 US 1,234 · DE 56 (+200) · JP 0" — marketplaces where the SKU has an FBA listing, biggest first. */
function inventoryLine_(o) {
  if (!o || !Object.keys(o.byMkt).length) return '📦 FBA 재고 정보 없음';
  var codes = Object.keys(o.byMkt);
  codes.sort(function (a, b) { return o.byMkt[b].f - o.byMkt[a].f || a.localeCompare(b); });
  var sum = 0;
  var parts = codes.map(function (c) {
    var m = o.byMkt[c];
    sum += m.f;
    var s = c + ' ' + (m.f ? '<b>' + m.f.toLocaleString() + '</b>' : '0');
    return m.in ? s + ' (+' + m.in.toLocaleString() + ')' : s;
  });
  return '📦 재고 ' + sum.toLocaleString() + ' — ' + parts.join(' · ');
}

/* ===================== CacheService chunking (100 KB per value) ===================== */

function cachePutChunked_(key, str, ttl) {
  var cache = CacheService.getScriptCache(), size = 90000, parts = {}, n = Math.ceil(str.length / size);
  for (var i = 0; i < n; i++) parts[key + '_' + i] = str.slice(i * size, (i + 1) * size);
  parts[key + '_n'] = String(n);
  cache.putAll(parts, ttl);
}

function cacheGetChunked_(key) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(key + '_n') || 0);
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var got = cache.getAll(keys), out = '';
  for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; out += got[keys[j]]; }
  return out;
}

/** Editor helper: drop cached index / marketplaces / tokens (after a master-sheet fix or a credential change). */
function clearProductsCache() {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(PRODUCT_INDEX_CACHE + '_n') || 0), keys = [PRODUCT_INDEX_CACHE + '_n', 'SPAPI_MKTS_V1', 'SPAPI_LWA', 'SPAPI_LWA_JP'];
  for (var i = 0; i < n; i++) keys.push(PRODUCT_INDEX_CACHE + '_' + i);
  cache.removeAll(keys);
}

/** Editor/test helper: logs the card JSON a query would produce. */
function testProducts() {
  Logger.log(JSON.stringify(handleProductsCommand_('ACS08008', false)).slice(0, 4000));
  Logger.log(JSON.stringify(handleProductsCommand_('ACS08008', true)).slice(0, 4000));
}
