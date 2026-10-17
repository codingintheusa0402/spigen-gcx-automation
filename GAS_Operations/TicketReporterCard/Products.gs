/**
 * `/products <검색어>` and `/products-similar <검색어>` — product lookup cards (사용자 지시 2026-10-17).
 *
 * /products          SKU master + ASIN master match (SKU, ASIN or keywords such as
 *                    "iPhone 17 Pro 울트라 하이브리드"), with FBA stock per country and an Amazon thumbnail.
 * /products-similar  the first match's device (SKU master 기종 / ASIN master 기종명), then every
 *                    other SKU made for that same device — "compatible with the same phone".
 *                    Multi-device 기종 values ("iPhone 18 Pro Max/17 Pro Max") match on any device.
 *
 * SP-API credentials are copies of the MCF_Tracking project's Script Properties (사용자 지시: use
 * only those), same names: LWA_CLIENT_ID / LWA_CLIENT_SECRET / LWA_REFRESH_TOKEN (EU app) and
 * LWA_*_JP (JP app). Checked 2026-10-17: they reach EU (DE GB FR IT ES NL SE PL BE IE TR AE SA)
 * and JP; NA answers 403, and the JP app has no Catalog role, so thumbnails come from DE / UK.
 *
 * Why a snapshot: Amazon seller SKUs are the master SKU plus a channel suffix (ACS08008UKO,
 * ACS08008PAN, ACS08008EUP, ACS08008F, …JPP) or a legacy code (000CG20879PAN), so the inventory
 * API's exact sellerSkus filter can't find them, and paging every SKU in every marketplace takes
 * minutes. refreshInventorySnapshot() (hourly trigger) reads two FBA reports instead —
 * GET_AFN_INVENTORY_DATA_BY_COUNTRY (EU, per country) and GET_AFN_INVENTORY_DATA (JP) — and folds
 * them by master SKU (or ASIN for legacy codes) into the script cache + the InventorySnapshot tab
 * of the queue spreadsheet. A Chat reply must come back within ~30 s, so the command only reads it.
 */

var SKU_MASTER_ID = '1JijzoYw9aDW-9Jx_OJOBByMVCX8-0OtNxb-0Y82EZJI';   // SKU_Master(먼데이보드) › Data
var ASIN_MASTER_ID = '1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo';  // ASIN_Master(먼데이보드) › Data
var PRODUCTS_MAX = 8;            // product rows per card (/products)
var SIMILAR_MAX = 12;            // product rows per card (/products-similar, incl. the base product)
var PRODUCT_INDEX_CACHE = 'PRODUCT_INDEX_V1';
var INVENTORY_CACHE = 'INVENTORY_SNAPSHOT_V1';
var INVENTORY_TAB = 'InventorySnapshot';

var SPAPI_REGIONS = {
  EU: { host: 'sellingpartnerapi-eu.amazon.com', creds: '' },
  FE: { host: 'sellingpartnerapi-fe.amazon.com', creds: '_JP' }
};
var INVENTORY_REPORTS = [
  { region: 'EU', type: 'GET_AFN_INVENTORY_DATA_BY_COUNTRY', marketplace: 'A1PA6795UKMFR9' },
  { region: 'FE', type: 'GET_AFN_INVENTORY_DATA', marketplace: 'A1VC38T7YXB528', country: 'JP' }
];
var CATALOG_MARKETS = [['A1PA6795UKMFR9', 'amazon.de'], ['A1F83G8C2ARO7P', 'amazon.co.uk']];  // thumbnail lookup order
var COUNTRY_ORDER = ['DE', 'GB', 'FR', 'IT', 'ES', 'PL', 'SE', 'NL', 'TR', 'JP'];

/* ===================== Command handler ===================== */

function handleProductsCommand_(query, similar) {
  query = String(query || '').trim();
  var cmd = similar ? '/products-similar' : '/products';
  if (!query) {
    return chatCreate_({ text: '⚠️ ' + cmd + ' 뒤에 SKU, ASIN 또는 제품명을 입력해 주세요. 예: "' + cmd + ' ACS08008", "' + cmd + ' iPhone 17 Pro 울트라 하이브리드"' });
  }
  var t0 = Date.now(), notes = [];
  var index = loadProductIndex_();
  var snap = loadInventorySnapshot_();
  if (!snap) notes.push('재고 스냅샷이 아직 없습니다 (refreshInventorySnapshot 미실행)');
  var stockOf = function (p) { return stockFor_(p, snap).sum; };
  var matches = searchProducts_(index, query, stockOf);
  var title, subtitle, list, total;

  if (!similar) {
    total = matches.length;
    list = matches.slice(0, PRODUCTS_MAX);
    if (!list.length && /^[A-Z]{3}\d{5}$/i.test(query)) list = [{ sku: query.toUpperCase(), asins: [] }];  // not in the masters — FBA may still know it
    if (!list.length && /^B0[A-Z0-9]{8}$/i.test(query)) list = [{ sku: '', asins: [query.toUpperCase()] }];
    title = '🔎 ' + query;
    subtitle = '마스터 일치 ' + total + '건' + (total > list.length ? ' · 상위 ' + list.length + '건 표시 (재고순)' : '');
  } else {
    var base = matches.filter(function (p) { return isRealDevice_(p.device); })[0];
    if (!base) {
      if (matches.length) {   // found, but 기종 is Common/blank — a universal accessory has no device to match on
        return chatCreate_({ text: '⚠️ "' + query + '" (' + matches[0].sku + ') 는 기종이 "' + (matches[0].device || '미지정') + '" 인 공용 제품이라 같은 기종 제품을 찾을 수 없습니다. /products 로 조회해 주세요.' });
      }
      return chatCreate_({ text: '⚠️ "' + query + '" 와 일치하는 제품이 SKU master / ASIN master 에 없습니다.' });
    }
    var devs = deviceSet_(base.device);
    var same = index.filter(function (p) {
      return p.sku !== base.sku && deviceSet_(p.device).some(function (d) { return devs.indexOf(d) >= 0; });
    });
    same.forEach(function (p) { p._stock = stockOf(p); });
    same.sort(function (a, b) {   // in stock first, then by category / model so a case's siblings sit together
      return (b._stock > 0) - (a._stock > 0) || (a.status === 'B') - (b.status === 'B') ||
             a.category.localeCompare(b.category) || b._stock - a._stock || a.model.localeCompare(b.model) || a.sku.localeCompare(b.sku);
    });
    total = same.length;
    list = [base].concat(same.slice(0, SIMILAR_MAX - 1));
    title = '📱 ' + base.device + ' 호환 제품';
    subtitle = '기준 ' + base.sku + ' · 같은 기종 ' + total + '건' + (total > list.length - 1 ? ' · 재고 있는 순 상위 ' + (list.length - 1) + '건' : '');
  }

  if (!list.length) {
    return chatCreate_({ text: '⚠️ "' + query + '" 와 일치하는 제품이 SKU master / ASIN master 에 없습니다.' });
  }

  var inv = {};
  list.forEach(function (p) {
    inv[p.sku || p.asins[0]] = stockFor_(p, snap);
    inv[p.sku || p.asins[0]].asins.forEach(function (a) { if (p.asins.indexOf(a) < 0) p.asins.push(a); });  // FBA knows ASINs the master lacks
  });
  var thumbs = {};
  try { thumbs = fetchThumbnails_(list); }
  catch (e) { notes.push('이미지 조회 실패: ' + e.message); }

  return chatCreate_({ cardsV2: [buildProductsCard_(title, subtitle, list, inv, thumbs, similar, snap, notes, Date.now() - t0)] });
}

/* ===================== Master index ===================== */

/** [{sku, asins[], model, device, nameKo, nameEn, brand, color, category, status}] from both masters, cached 6 h. */
function loadProductIndex_() {
  var cached = cacheGetChunked_(PRODUCT_INDEX_CACHE);
  if (cached) return JSON.parse(cached);
  return buildProductIndex_();
}

function buildProductIndex_() {
  var bySku = {};
  readColumns_(SKU_MASTER_ID, ['name', '모델명', '자재내역(KO)', '자재내역(EN)', '브랜드', '기종', '색상', '대분류', '품목상태']).forEach(function (r) {
    var sku = String(r[0] || '').trim().toUpperCase();
    if (sku) bySku[sku] = { sku: sku, asins: [], model: r[1], nameKo: r[2], nameEn: r[3], brand: r[4], device: r[5], color: r[6], category: r[7], status: r[8] };
  });
  readColumns_(ASIN_MASTER_ID, ['SKU', 'ASIN', '모델명', '자재내역', '자재내역(EN)', '브랜드', '기종명', '색상명', '대분류']).forEach(function (r) {
    var sku = String(r[0] || '').trim().toUpperCase();
    var asin = String(r[1] || '').trim().toUpperCase();
    if (!sku) return;
    var p = bySku[sku];
    if (!p) p = bySku[sku] = { sku: sku, asins: [], model: r[2], nameKo: r[3], nameEn: r[4], brand: r[5], device: r[6], color: r[7], category: r[8], status: '' };
    if (/^B0[A-Z0-9]{8}$/.test(asin) && p.asins.indexOf(asin) < 0) p.asins.push(asin);
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
  for (var r = 0; r < last - 1; r++) out.push(cols.map(function (c) { return c ? c[r][0] : ''; }));
  return out;
}

/** SKU / ASIN exact match first; otherwise every keyword must appear somewhere in the product's text. */
function searchProducts_(index, query, stockOf) {
  var q = query.trim().toUpperCase();
  var exact = index.filter(function (p) { return p.sku === q || p.asins.indexOf(q) >= 0; });
  if (exact.length) return exact;

  var terms = query.toLowerCase().split(/\s+/).filter(String);
  var hits = index.filter(function (p) {
    var hay = [p.sku, p.asins.join(' '), p.model, p.nameKo, p.nameEn, p.brand, p.device, p.color, p.category].join(' ').toLowerCase();
    return terms.every(function (t) { return hay.indexOf(t) >= 0; });
  });
  hits.forEach(function (p) { p._stock = stockOf(p); });
  hits.sort(function (a, b) {   // most FBA stock first, then 품목상태 A before B, then SKU
    return b._stock - a._stock || (a.status === 'B') - (b.status === 'B') || a.sku.localeCompare(b.sku);
  });
  return hits;
}

function isRealDevice_(d) {
  d = String(d || '').trim();
  return !!d && !/^(common|공용|미정|-|n\/a)$/i.test(d);
}

/**
 * Every device a 기종 value fits, normalised: "iPhone 18 Pro Max/17 Pro Max" → ["iphone 18 pro max",
 * "iphone 17 pro max"] (a later part borrows the first part's first word), "Model Y/3" →
 * ["model y", "model 3"], "#iPad Air 11" (2024)/" → ["ipad air 11" (2024)"].
 */
function deviceSet_(d) {
  var parts = String(d || '').replace(/#/g, '').split('/').map(function (x) { return x.replace(/\s+/g, ' ').trim(); }).filter(String);
  if (!parts.length || !isRealDevice_(parts[0])) return [];
  var family = parts[0].split(' ')[0];   // iPhone / Galaxy / Model / Pixel …
  return parts.map(function (x, i) {
    return (i && x.split(' ')[0].toLowerCase() !== family.toLowerCase() ? family + ' ' + x : x).toLowerCase();
  });
}

/* ===================== Inventory snapshot ===================== */

/**
 * Trigger target (hourly). Never waits on Amazon: for each FBA inventory report type it folds the
 * newest DONE report (normally the one the previous run requested) and then requests a fresh one
 * for the next run, unless one is already queued. Stores { ts, by: { KEY: { c: {DE: n, …, JP: n},
 * a: [asins] } } } where KEY = master SKU (seller SKU's leading AAA00000) or "ASIN:<asin>" for
 * legacy seller SKUs, and refreshes the master index cache so Chat users never wait on the
 * 15k-row sheet read. (A first version polled inside one run; the EU report took > 4.5 min.)
 */
function refreshInventorySnapshot() {
  var by = {}, times = [];
  INVENTORY_REPORTS.forEach(function (cfg) {
    var q = 'reportTypes=' + cfg.type + '&marketplaceIds=' + cfg.marketplace + '&pageSize=10';
    var done = (spapiRequest_(cfg.region, 'get', '/reports/2021-06-30/reports', q + '&processingStatuses=DONE').reports || [])
      .sort(function (x, y) { return y.createdTime.localeCompare(x.createdTime); })[0];
    if (done) {
      foldInventoryReport_(by, cfg, downloadReport_(cfg.region, done.reportDocumentId));
      times.push(done.createdTime);
    }
    var queued = spapiRequest_(cfg.region, 'get', '/reports/2021-06-30/reports', q + '&processingStatuses=IN_QUEUE,IN_PROGRESS').reports || [];
    if (!queued.length) {
      try { spapiRequest_(cfg.region, 'post', '/reports/2021-06-30/reports', null, { reportType: cfg.type, marketplaceIds: [cfg.marketplace] }); }
      catch (e) { Logger.log('report request skipped (next run retries): %s', e.message); }   // createReport quota ≈ 1/min
    }
  });
  if (times.length < INVENTORY_REPORTS.length) {
    Logger.log('no finished report yet for every type — requested; snapshot left as is');
    return;
  }

  var snap = { ts: times.sort()[0], by: by };   // oldest of the two reports = how fresh the numbers are
  var json = JSON.stringify(snap);
  cachePutChunked_(INVENTORY_CACHE, json, 6 * 3600);
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (id) {   // durable copy (cache can be evicted): 45k-char chunks down column A
    var ss = SpreadsheetApp.openById(id);
    var sh = ss.getSheetByName(INVENTORY_TAB) || ss.insertSheet(INVENTORY_TAB);
    var rows = [];
    for (var i = 0; i < json.length; i += 45000) rows.push([json.slice(i, i + 45000)]);
    sh.clearContents();
    sh.getRange(1, 1, rows.length, 1).setValues(rows);
  }
  buildProductIndex_();
  Logger.log('inventory snapshot as of %s: %s keys, %s chars', snap.ts, Object.keys(by).length, json.length);
}

function foldInventoryReport_(by, cfg, rows) {
  var head = rows[0];
  var iSku = head.indexOf('seller-sku'), iAsin = head.indexOf('asin'), iCountry = head.indexOf('country'),
      iQty = cfg.country ? head.indexOf('Quantity Available') : head.indexOf('quantity-for-local-fulfillment'),
      iCond = head.indexOf('Warehouse-Condition-code');
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    if (iCond >= 0 && row[iCond] !== 'SELLABLE') continue;
    var sku = String(row[iSku] || '').trim().toUpperCase(), asin = String(row[iAsin] || '').trim();
    var m = sku.match(/^[A-Z]{3}\d{5}/);
    var key = m ? m[0] : (asin ? 'ASIN:' + asin : null);
    if (!key) continue;
    var country = cfg.country || row[iCountry];
    var o = by[key] = by[key] || { c: {}, a: [] };
    o.c[country] = (o.c[country] || 0) + (Number(row[iQty]) || 0);
    if (/^B0[A-Z0-9]{8}$/.test(asin) && o.a.indexOf(asin) < 0) o.a.push(asin);
  }
}

function downloadReport_(region, documentId) {
  var doc = spapiRequest_(region, 'get', '/reports/2021-06-30/documents/' + documentId);
  var blob = UrlFetchApp.fetch(doc.url).getBlob();
  if (doc.compressionAlgorithm === 'GZIP') blob = Utilities.ungzip(blob.setContentType('application/x-gzip'));
  return Utilities.parseCsv(blob.getDataAsString('UTF-8'), '\t');
}

function loadInventorySnapshot_() {
  var json = cacheGetChunked_(INVENTORY_CACHE);
  if (!json) {
    var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
    var sh = id && SpreadsheetApp.openById(id).getSheetByName(INVENTORY_TAB);
    if (!sh || sh.getLastRow() < 1) return null;
    json = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
    cachePutChunked_(INVENTORY_CACHE, json, 3600);
  }
  return JSON.parse(json);
}

/** Merged stock for a product: its master SKU key plus any legacy seller SKUs found under its ASINs. */
function stockFor_(p, snap) {
  var out = { c: {}, asins: [], sum: 0 };
  if (!snap) return out;
  var keys = [p.sku].concat((p.asins || []).map(function (a) { return 'ASIN:' + a; }));
  keys.forEach(function (k) {
    var o = k && snap.by[k];
    if (!o) return;
    Object.keys(o.c).forEach(function (c) { out.c[c] = (out.c[c] || 0) + o.c[c]; out.sum += o.c[c]; });
    o.a.forEach(function (a) { if (out.asins.indexOf(a) < 0) out.asins.push(a); });
  });
  return out;
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

function spapiRequest_(region, method, path, query, body) {
  var opts = {
    method: method, muteHttpExceptions: true,
    headers: { 'x-amz-access-token': spapiToken_(region), 'user-agent': 'GCX-TicketReporter/1.0 (Language=AppsScript)' }
  };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  var url = 'https://' + SPAPI_REGIONS[region].host + path + (query ? '?' + query : '');
  for (var attempt = 0; ; attempt++) {
    var res = UrlFetchApp.fetch(url, opts), code = res.getResponseCode();
    if ((code === 429 || code >= 500) && attempt < 3) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    if (code >= 300) throw new Error('SP-API ' + method.toUpperCase() + ' ' + path + ' → ' + code + ' ' + res.getContentText().slice(0, 200));
    return JSON.parse(res.getContentText());
  }
}

/** ASIN → { img, domain } via Catalog Items searchCatalogItems (20 ASINs per call), DE then UK. */
function fetchThumbnails_(list) {
  var want = [];
  list.forEach(function (p) { p.asins.slice(0, 3).forEach(function (a) { if (want.indexOf(a) < 0) want.push(a); }); });
  var found = {};
  CATALOG_MARKETS.forEach(function (mk) {
    want = want.filter(function (a) { return !found[a]; });
    for (var c = 0; c < want.length; c += 20) {
      var res = spapiRequest_('EU', 'get', '/catalog/2022-04-01/items',
        'identifiers=' + want.slice(c, c + 20).join(',') + '&identifiersType=ASIN&marketplaceIds=' + mk[0] + '&includedData=images&pageSize=20');
      (res.items || []).forEach(function (it) {
        var img = pickThumb_(it.images);
        if (img && !found[it.asin]) found[it.asin] = { img: img, domain: mk[1] };
      });
    }
  });
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

function buildProductsCard_(title, subtitle, list, inv, thumbs, similar, snap, notes, ms) {
  var sections = list.map(function (p, i) {
    var asin = p.asins.filter(function (a) { return thumbs[a]; })[0] || p.asins[0] || '';
    var th = asin && thumbs[asin];
    var domain = th ? th.domain : 'amazon.de';
    var lines = [];
    var name = [p.brand, p.model, p.color].filter(Boolean).join(' · ');
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
    lines.push(inventoryLine_(inv[p.sku || p.asins[0]]));

    var textCol = { textParagraph: { text: lines.join('<br>') } };
    var widget = th ? {
      columns: { columnItems: [
        { horizontalSizeStyle: 'FILL_MINIMUM_SPACE', horizontalAlignment: 'CENTER', verticalAlignment: 'CENTER',
          widgets: [{ image: { imageUrl: th.img, altText: p.sku || asin, onClick: { openLink: { url: 'https://www.' + domain + '/dp/' + asin } } } }] },
        { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', horizontalAlignment: 'START', verticalAlignment: 'CENTER',
          widgets: [textCol] }
      ] }
    } : textCol;
    return { header: (similar && i === 0 ? '⭐ 기준 · ' : '') + (p.sku || asin), widgets: [widget] };
  });

  var footer = [];
  if (notes.length) footer.push('⚠️ ' + notes.map(escapeHtml_).join('<br>⚠️ '));
  var asOf = snap ? Utilities.formatDate(new Date(snap.ts), 'Asia/Seoul', 'M/d HH:mm') : '-';
  footer.push('<font color="#5f6368">📦 = FBA 판매가능 재고 (EU 국가별 + JP, ' + asOf + ' 기준 · NA 미지원) · ' + (ms / 1000).toFixed(1) + '초</font>');
  footer.push('<a href="https://docs.google.com/spreadsheets/d/' + SKU_MASTER_ID + '/edit">SKU master</a> · ' +
              '<a href="https://docs.google.com/spreadsheets/d/' + ASIN_MASTER_ID + '/edit">ASIN master</a>');
  sections.push({ widgets: [{ textParagraph: { text: footer.join('<br>') } }] });

  return {
    cardId: 'products-' + new Date().getTime(),
    card: { header: { title: title, subtitle: subtitle }, sections: sections }
  };
}

/** "📦 1,234 — DE 800 · GB 300 · JP 134" — countries with stock, in COUNTRY_ORDER. */
function inventoryLine_(o) {
  if (!o || !o.sum) return '📦 FBA 재고 없음';
  var codes = Object.keys(o.c).filter(function (c) { return o.c[c] > 0; });
  codes.sort(function (a, b) { return (COUNTRY_ORDER.indexOf(a) + 99) % 99 - (COUNTRY_ORDER.indexOf(b) + 99) % 99 || a.localeCompare(b); });
  return '📦 <b>' + o.sum.toLocaleString() + '</b> — ' + codes.map(function (c) { return c + ' ' + o.c[c].toLocaleString(); }).join(' · ');
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

/** Editor/test helper: logs the card JSON for a sample query. */
function testProducts() {
  Logger.log(JSON.stringify(handleProductsCommand_('ACS08008', false)).slice(0, 3000));
  Logger.log(JSON.stringify(handleProductsCommand_('ACS08008', true)).slice(0, 3000));
}
