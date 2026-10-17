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
 *
 * Seller Central listings (사용자 지시 2026-10-18): FBA reports only hold SKUs with FBA stock, so a SKU
 * that is listed but inactive / out of stock / merchant-fulfilled showed no ASIN at all.
 * refreshListingsSnapshot() (6-hourly trigger, requests fresh reports once a day) folds
 * GET_MERCHANT_LISTINGS_ALL_DATA — everything in Manage All Inventory — for each EU marketplace + JP
 * into SKU → ASINs + active/inactive per country. Anything still without an ASIN is looked up live in
 * the Amazon catalog by the SKU master's EAN/UPC (finds e.g. CYRILL items listed by another account).
 */

var SKU_MASTER_ID = '1JijzoYw9aDW-9Jx_OJOBByMVCX8-0OtNxb-0Y82EZJI';   // SKU_Master(먼데이보드) › Data
var ASIN_MASTER_ID = '1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo';  // ASIN_Master(먼데이보드) › Data
var PRODUCTS_MAX = 8;            // product rows per card (/products)
var SIMILAR_MAX = 12;            // product rows per card (/products-similar, incl. the base product)
var PRODUCT_INDEX_CACHE = 'PRODUCT_INDEX_V3';   // V2: + ean / upc, V3: + dept
var INVENTORY_CACHE = 'INVENTORY_SNAPSHOT_V1';
var INVENTORY_TAB = 'InventorySnapshot';
var LISTINGS_CACHE = 'LISTINGS_SNAPSHOT_V1';
var LISTINGS_TAB = 'ListingsSnapshot';
var LISTINGS_IDS_PROP = 'LISTINGS_REPORT_IDS';   // report ids the stored listings snapshot was built from

var SPAPI_REGIONS = {
  EU: { host: 'sellingpartnerapi-eu.amazon.com', creds: '' },
  FE: { host: 'sellingpartnerapi-fe.amazon.com', creds: '_JP' }
};
var INVENTORY_REPORTS = [
  { region: 'EU', type: 'GET_AFN_INVENTORY_DATA_BY_COUNTRY', marketplace: 'A1PA6795UKMFR9' },
  { region: 'FE', type: 'GET_AFN_INVENTORY_DATA', marketplace: 'A1VC38T7YXB528', country: 'JP' }
];
var CATALOG_MARKETS = [['A1PA6795UKMFR9', 'amazon.de'], ['A1F83G8C2ARO7P', 'amazon.co.uk']];  // thumbnail lookup order
// [marketplace id, country code, region] for the Seller Central listings reports
var LISTING_MARKETS = [
  ['A1PA6795UKMFR9', 'DE', 'EU'], ['A1F83G8C2ARO7P', 'GB', 'EU'], ['A13V1IB3VIYZZH', 'FR', 'EU'], ['APJ6JRA9NG5V4', 'IT', 'EU'],
  ['A1RKKUPIHCS9HS', 'ES', 'EU'], ['A1805IZSGTT6HS', 'NL', 'EU'], ['A2NODRKZP88ZB9', 'SE', 'EU'], ['A1C3SOZRARQ6R3', 'PL', 'EU'],
  ['AMEN7PMS3EDWL', 'BE', 'EU'], ['A28R8C7NBKEWEA', 'IE', 'EU'], ['A33AVAJ2PDY3EV', 'TR', 'EU'], ['A2VIGQ35RCS4UG', 'AE', 'EU'],
  ['A17E79C6D8DWNP', 'SA', 'EU'], ['A1VC38T7YXB528', 'JP', 'FE']
];
var COUNTRY_ORDER = ['DE', 'GB', 'FR', 'IT', 'ES', 'PL', 'SE', 'NL', 'BE', 'IE', 'TR', 'AE', 'SA', 'JP'];

/* ===================== Command handler ===================== */

function handleProductsCommand_(query, similar) {
  query = String(query || '').trim();
  var cmd = similar ? '/products-similar' : '/products';
  if (!query) {
    return chatCreate_({ text: '⚠️ ' + cmd + ' 뒤에 SKU, ASIN 또는 제품명을 입력해 주세요. 예: "' + cmd + ' ACS08008", "' + cmd + ' iPhone 17 Pro 울트라 하이브리드"' });
  }
  var t0 = Date.now(), notes = [];
  var index = loadProductIndex_();
  var snap = loadSnapshot_(INVENTORY_CACHE, INVENTORY_TAB);
  if (!snap) notes.push('재고 스냅샷이 아직 없습니다 (refreshInventorySnapshot 미실행)');
  var lsnap = loadSnapshot_(LISTINGS_CACHE, LISTINGS_TAB);
  if (!lsnap) notes.push('Seller Central 리스팅 스냅샷이 아직 없습니다 (refreshListingsSnapshot 미실행)');
  var stockOf = function (p) { return stockFor_(p, snap).sum; };
  var matches = searchProducts_(index, query, stockOf);
  if (!matches.length && /^B0[A-Z0-9]{8}$/i.test(query)) matches = productsForAsin_(index, query.toUpperCase(), [snap, lsnap]);
  var title, subtitle, list, total;

  if (!similar) {
    total = matches.length;
    list = matches.slice(0, PRODUCTS_MAX);
    if (!list.length && /^[A-Z]{3}\d{5}$/i.test(query)) list = [{ sku: query.toUpperCase(), asins: [] }];  // not in the masters — FBA may still know it
    if (!list.length && /^B0[A-Z0-9]{8}$/i.test(query)) list = [{ sku: '', asins: [query.toUpperCase()] }];
    title = '🔎 ' + query;
    subtitle = '마스터 일치 ' + total + '건' + (total > list.length ? ' · 상위 ' + list.length + '건 표시 (재고순)' : '');
  } else {
    var sim = similarProducts_(index, query, matches, stockOf);
    if (sim.error) return chatCreate_({ text: sim.error });
    if (!sim.same.length && !sim.base) {
      return chatCreate_({ text: '⚠️ ' + sim.label + (sim.cat ? ' · ' + sim.cat.label : '') + ' 에 해당하는 제품이 SKU master / ASIN master 에 없습니다.' +
        (sim.cat && sim.cat.key === 'sda' ? ' (SDA = 담당부서코드 A003 제품 중 기종이 ' + sim.label + ' 로 등록된 것)' : '') });
    }
    total = sim.same.length;
    list = (sim.base ? [sim.base] : []).concat(sim.same.slice(0, SIMILAR_MAX - (sim.base ? 1 : 0)));
    var shown = list.length - (sim.base ? 1 : 0);
    title = '📱 ' + sim.label + ' 호환 ' + (sim.cat ? sim.cat.label : '제품');
    subtitle = (sim.base ? '기준 ' + sim.base.sku + ' · ' : '') + '같은 기종 ' + total + '건' + (total > shown ? ' · 재고 있는 순 상위 ' + shown + '건' : '');
  }

  if (!list.length) {
    return chatCreate_({ text: '⚠️ "' + query + '" 와 일치하는 제품이 SKU master / ASIN master 에 없습니다.' });
  }

  // ASINs the masters lack: FBA stock → Seller Central listings → catalog by EAN/UPC (each step only adds)
  var inv = {}, lst = {};
  list.forEach(function (p) {
    p.asins = p.asins.slice();
    var k = p.sku || p.asins[0];
    inv[k] = stockFor_(p, snap);
    addAsins_(p, inv[k].asins);
    lst[k] = listingFor_(p, lsnap);
    addAsins_(p, lst[k].asins);
  });
  try {
    var byCode = asinsByBarcode_(list.filter(function (p) { return !p.asins.length && (p.ean || p.upc); }));
    list.forEach(function (p) {
      var hit = byCode[p.ean] || byCode[p.upc];
      if (!hit || p.asins.length) return;
      p.viaBarcode = true;
      addAsins_(p, hit);
      var k = p.sku || p.asins[0];
      lst[k] = listingFor_(p, lsnap);   // the catalog ASIN may still be ours under a legacy seller SKU
      inv[k] = stockFor_(p, snap);
    });
  } catch (e) { notes.push('EAN 카탈로그 조회 실패: ' + e.message); }
  var thumbs = {};
  try { thumbs = fetchThumbnails_(list); }
  catch (e) { notes.push('이미지 조회 실패: ' + e.message); }

  return chatCreate_({ cardsV2: [buildProductsCard_(title, subtitle, list, inv, lst, thumbs, similar, snap, lsnap, notes, Date.now() - t0)] });
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
  readColumns_(SKU_MASTER_ID, ['name', '모델명', '자재내역(KO)', '자재내역(EN)', '브랜드', '기종', '색상', '대분류', '품목상태', 'EAN바코드', 'UPC바코드', '담당부서코드']).forEach(function (r) {
    var sku = String(r[0] || '').trim().toUpperCase();
    if (sku) bySku[sku] = { sku: sku, asins: [], model: r[1], nameKo: r[2], nameEn: r[3], brand: r[4], device: r[5], color: r[6], category: r[7], status: r[8], ean: r[9], upc: r[10], dept: r[11] };
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
    ['model', 'nameKo', 'nameEn', 'brand', 'device', 'color', 'category', 'status', 'ean', 'upc', 'dept'].forEach(function (f) { p[f] = String(p[f] || '').trim(); });
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

/**
 * Category words (사용자 지시 2026-10-18) that narrow either command: "case" → 대분류 …케이스;
 * "screen protector" / "SP" / "CP" / 필름 / 글라스 → 대분류 …보호필름 (camera protectors included);
 * "SDA" → 담당부서코드 A003, the Smart Device Acc. team (every SKU in the SDA review sheet is A003).
 */
var CATEGORY_WORDS = [
  { key: 'case', label: '케이스', words: ['case', 'cases', 'cover', '케이스', '커버'], test: function (p) { return /케이스$/.test(p.category); } },
  { key: 'sp', label: '보호필름 (SP·CP)', words: ['screen protector', 'screen protectors', 'camera protector', 'lens protector', 'protector', 'protectors',
      'screen', 'sp', 'cp', 'glass', 'film', '필름', '보호필름', '글라스', '강화유리', '액정필름', '카메라필름'], test: function (p) { return /보호필름$/.test(p.category); } },
  { key: 'sda', label: 'SDA', words: ['sda', 'smart device acc', 'smart device accessories'], test: function (p) { return p.dept === 'A003'; } }
];

// Common spellings / Korean names folded into the English device words before matching
var DEVICE_ALIASES = { iph: 'iphone', ip: 'iphone', ipone: 'iphone', iphon: 'iphone', '아이폰': 'iphone', '갤럭시': 'galaxy', '갤': 'galaxy',
  '픽셀': 'pixel', '프로': 'pro', '맥스': 'max', '울트라': 'ultra', '플러스': 'plus', '에어': 'air', '폴드': 'fold', '플립': 'flip' };

/**
 * Splits a query into { device, cat, rest }: the longest known device it names (typo-tolerant, numbers
 * exact), a category word, and the remaining words. "ipone 17 pro case" → iphone 17 pro / case / "".
 */
function parseProductQuery_(index, query) {
  var words = splitWords_(query)
    .map(function (w) { return DEVICE_ALIASES[w] || w; });
  var cat = null;
  CATEGORY_WORDS.forEach(function (c) {   // longest phrase first so "screen protector" beats "screen"
    if (cat) return;
    c.words.slice().sort(function (x, y) { return y.split(' ').length - x.split(' ').length; }).some(function (w) {
      var ws = w.split(' ');
      for (var i = 0; i + ws.length <= words.length; i++) {
        if (ws.every(function (t, k) { return wordMatch_(words[i + k], t); })) { words.splice(i, ws.length); cat = c; return true; }
      }
      return false;
    });
  });
  var best = null, at = -1;
  var devs = knownDevices_(index);
  Object.keys(devs).forEach(function (d) {
    var ds = d.split(' ');
    if (best && ds.length <= best.split(' ').length) return;   // longest device wins ("iphone 17 pro max" over "iphone 17 pro")
    for (var i = 0; i + ds.length <= words.length; i++) {
      if (ds.every(function (t, k) { return wordMatch_(words[i + k], t); })) {
        // a following word that would make it a longer device ("max") means this isn't the device named
        best = d; at = i; return;
      }
    }
  });
  if (best) {
    var next = words[at + best.split(' ').length];
    var longer = next && Object.keys(devs).some(function (d) { return d.indexOf(best + ' ') === 0 && wordMatch_(next, d.split(' ')[best.split(' ').length]); });
    if (longer) { best = null; } else { words.splice(at, best.split(' ').length); }
  }
  return { device: best ? devs[best] : null, cat: cat, rest: words.join(' ') };
}

/** Same word allowing small typos: numbers / ≤2 letters exact, 3–5 letters 1 edit, longer 2 edits. */
function wordMatch_(a, b) {
  if (a === b) return true;
  if (!a || !b || /\d/.test(a) || /\d/.test(b) || /[^\x00-\x7f]/.test(a + b)) return false;
  var n = Math.min(a.length, b.length);
  var max = n <= 2 ? 0 : n <= 5 ? 1 : 2;
  return max > 0 && Math.abs(a.length - b.length) <= max && editDistance_(a, b) <= max;
}

function editDistance_(a, b) {   // Damerau (adjacent swaps count as one edit, "ipohne")
  var d = [];
  for (var i = 0; i <= a.length; i++) { d[i] = [i]; }
  for (var j = 1; j <= b.length; j++) d[0][j] = j;
  for (i = 1; i <= a.length; i++) {
    for (j = 1; j <= b.length; j++) {
      var c = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + c);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
    }
  }
  return d[a.length][b.length];
}

/** { split form ("galaxy watch ultra 2") → device as in deviceSet_ ("galaxy watch ultra2") } */
function knownDevices_(index) {
  if (index._devices) return index._devices;
  var seen = {};
  index.forEach(function (p) { deviceSet_(p.device).forEach(function (d) { seen[splitWords_(d).join(' ')] = d; }); });
  return (index._devices = seen);
}

/** Lower-cased words with letters and digits split ("iphone17pro" → iphone 17 pro) and aliases applied. */
function splitWords_(s) {
  return normText_(s).replace(/([a-z])(\d)/g, '$1 $2').replace(/(\d)([a-z]{2,})/g, '$1 $2').split(' ').filter(String)
    .map(function (w) { return DEVICE_ALIASES[w] || w; });
}

/** Same words in the same order, each allowing a small typo ("ultra hybird" = "Ultra Hybrid"). */
function sameWords_(a, b) {
  var x = a.split(' '), y = normText_(b).split(' ');
  return x.length === y.length && x.every(function (w, i) { return wordMatch_(w, y[i]); });
}

/**
 * Exact matching (사용자 지시 2026-10-18): a SKU / ASIN / EAN must match exactly. Free text = a device
 * (whole words, typo-tolerant — "iPhone 17 Pro" never takes in 17 Pro Max or 11 Pro), an optional
 * category word, and the rest, which must equal a model name or match as (typo-tolerant) whole words.
 */
function searchProducts_(index, query, stockOf) {
  var q = query.trim().toUpperCase();
  if (/^[A-Z]{3}\d{5}$/.test(q) || /^B0[A-Z0-9]{8}$/.test(q) || /^\d{12,13}$/.test(q)) {
    return index.filter(function (p) { return p.sku === q || p.asins.indexOf(q) >= 0 || p.ean === q || p.upc === q; });
  }
  var pq = parseProductQuery_(index, query);
  if (!pq.device && !pq.rest) return [];   // a category word alone is not a search
  var pool = index.filter(function (p) {
    return (!pq.device || deviceSet_(p.device).indexOf(pq.device) >= 0) && (!pq.cat || pq.cat.test(p));
  });
  var hits = pool;
  if (pq.rest) {
    var exactModel = pool.filter(function (p) { return sameWords_(pq.rest, p.model); });
    hits = exactModel.length ? exactModel : pool.filter(function (p) {
      var en = normText_([p.model, p.nameEn, p.brand, p.color, p.category, p.sku].join(' ')).split(' ');
      var ko = normText_(p.nameKo + ' ' + p.category);
      return pq.rest.split(' ').every(function (t) {
        return /[^\x00-\x7f]/.test(t) ? ko.indexOf(t) >= 0 : en.some(function (w) { return wordMatch_(t, w); });
      });
    });
  }
  hits.forEach(function (p) { p._stock = stockOf(p); });
  hits.sort(function (a, b) {   // most FBA stock first, then 품목상태 A before B, then SKU
    return b._stock - a._stock || (a.status === 'B') - (b.status === 'B') || a.sku.localeCompare(b.sku);
  });
  return hits;
}

/**
 * Same-device products. The device is the one named in the query ("/products -similar ipone 17 pro case"),
 * else the device(s) of the exactly matched product; a category word narrows either.
 * → { base, label, cat, same[] } or { error }.
 */
function similarProducts_(index, query, matches, stockOf) {
  var q = query.trim().toUpperCase();
  var isId = /^[A-Z]{3}\d{5}|^B0[A-Z0-9]{8}|^\d{12,13}/.test(q);
  var pq = parseProductQuery_(index, isId ? query.replace(/^\S+/, '') : query);
  var devs, base = null, label;
  if (!isId && pq.device && !pq.rest) {
    devs = [pq.device];
    label = deviceLabel_(index, pq.device);
  } else {
    if (isId) matches = searchProducts_(index, query.trim().split(/\s+/)[0], stockOf);
    base = matches.filter(function (p) { return isRealDevice_(p.device); })[0];
    if (!base) {
      if (matches.length) {   // found, but 기종 is Common/blank — a universal accessory has no device to match on
        return { error: '⚠️ "' + query + '" (' + matches[0].sku + ') 는 기종이 "' + (matches[0].device || '미지정') + '" 인 공용 제품이라 같은 기종 제품을 찾을 수 없습니다. /products 로 조회해 주세요.' };
      }
      return { error: '⚠️ "' + query + '" 에서 제품이나 기종을 찾지 못했습니다. SKU, ASIN 또는 기종명(예: iPhone 17 Pro, 뒤에 case / screen protector / SDA)으로 입력해 주세요.' };
    }
    devs = deviceSet_(base.device);
    label = base.device;
  }
  var same = index.filter(function (p) {
    return (!base || p.sku !== base.sku) && (!pq.cat || pq.cat.test(p)) &&
           deviceSet_(p.device).some(function (d) { return devs.indexOf(d) >= 0; });
  });
  same.forEach(function (p) { p._stock = stockOf(p); });
  same.sort(function (a, b) {   // in stock first, then by category / model so a case's siblings sit together
    return (b._stock > 0) - (a._stock > 0) || (a.status === 'B') - (b.status === 'B') ||
           a.category.localeCompare(b.category) || b._stock - a._stock || a.model.localeCompare(b.model) || a.sku.localeCompare(b.sku);
  });
  return { base: base, label: label, cat: pq.cat, same: same };
}

/** The 기종 text as written for a normalised device, e.g. "iphone 17 pro" → "iPhone 17 Pro". */
function deviceLabel_(index, dev) {
  for (var i = 0; i < index.length; i++) {
    var set = deviceSet_(index[i].device);
    if (set.length === 1 && set[0] === dev) return String(index[i].device).replace(/#|\/$/g, '').trim();
  }
  return dev.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

function normText_(s) { return String(s || '').toLowerCase().replace(/[#,()\[\]"]/g, ' ').replace(/\s+/g, ' ').trim(); }

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
  saveSnapshot_(INVENTORY_CACHE, INVENTORY_TAB, snap);
  buildProductIndex_();
  Logger.log('inventory snapshot as of %s: %s keys', snap.ts, Object.keys(by).length);
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

/** Script cache (6 h) + a durable copy in the queue spreadsheet tab (cache can be evicted), 45k-char chunks down column A. */
function saveSnapshot_(cacheKey, tab, obj) {
  var json = JSON.stringify(obj);
  cachePutChunked_(cacheKey, json, 6 * 3600);
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return;
  var ss = SpreadsheetApp.openById(id);
  var sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
  var rows = [];
  for (var i = 0; i < json.length; i += 45000) rows.push([json.slice(i, i + 45000)]);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function loadSnapshot_(cacheKey, tab) {
  var json = cacheGetChunked_(cacheKey);
  if (!json) {
    var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
    var sh = id && SpreadsheetApp.openById(id).getSheetByName(tab);
    if (!sh || sh.getLastRow() < 1) return null;
    json = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
    cachePutChunked_(cacheKey, json, 3600);
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

/* ===================== Seller Central listings snapshot ===================== */

/**
 * Trigger target (every 6 h). Per marketplace: requests a GET_MERCHANT_LISTINGS_ALL_DATA report when
 * none was asked for in the last 20 h (so ≈ daily — the EU ones take ~8 min to build), and when the
 * newest finished reports differ from the ones the stored snapshot came from, rebuilds
 * { ts, by: { KEY: { a: [asins], on: 'DE GB', off: 'IT' } } } (KEY as in the inventory snapshot).
 */
function refreshListingsSnapshot() {
  try { refreshListingsSnapshot_(); }
  catch (e) {
    if (!/→ 429/.test(e.message)) throw e;
    Logger.log('SP-API quota hit — next run continues: %s', e.message.slice(0, 120));   // getReports refills 1 per 45 s
  }
}

function refreshListingsSnapshot_() {
  var since = new Date(Date.now() - 20 * 3600 * 1000).toISOString();
  var window = new Date(Date.now() - 72 * 3600 * 1000).toISOString();
  var reps = [];   // one getReports call per region (quota: burst 10, then 1 per 45 s)
  ['EU', 'FE'].forEach(function (rg) {
    // no marketplaceIds filter (max 10 per call; EU has 13) — matched per marketplace below
    reps = reps.concat(spapiRequest_(rg, 'get', '/reports/2021-06-30/reports', 'reportTypes=GET_MERCHANT_LISTINGS_ALL_DATA' +
      '&pageSize=100&createdSince=' + encodeURIComponent(window)).reports || []);
  });
  if (!reps.length && PropertiesService.getScriptProperties().getProperty(LISTINGS_IDS_PROP)) Logger.log('no listings reports in 72 h');
  reps.sort(function (x, y) { return y.createdTime.localeCompare(x.createdTime); });
  var latest = [];
  LISTING_MARKETS.forEach(function (m) {
    var mine = reps.filter(function (r) { return (r.marketplaceIds || []).indexOf(m[0]) >= 0; });
    var done = mine.filter(function (r) { return r.processingStatus === 'DONE'; })[0];
    if (done) latest.push({ m: m, rep: done });
    var recent = mine.filter(function (r) { return r.createdTime >= since && r.processingStatus !== 'CANCELLED' && r.processingStatus !== 'FATAL'; });
    if (!recent.length) {
      try { spapiRequest_(m[2], 'post', '/reports/2021-06-30/reports', null, { reportType: 'GET_MERCHANT_LISTINGS_ALL_DATA', marketplaceIds: [m[0]] }); }
      catch (e) { Logger.log('%s listings request skipped (next run retries): %s', m[1], e.message); }
    }
  });
  var ids = latest.map(function (l) { return l.rep.reportId; }).sort().join(',');
  var props = PropertiesService.getScriptProperties();
  if (!latest.length || ids === props.getProperty(LISTINGS_IDS_PROP)) {
    Logger.log('listings: nothing new (%s markets with a finished report)', latest.length);
    return;
  }
  var by = {};
  latest.forEach(function (l) {
    var doc = spapiRequest_(l.m[2], 'get', '/reports/2021-06-30/documents/' + l.rep.reportDocumentId);
    var blob = UrlFetchApp.fetch(doc.url).getBlob();
    if (doc.compressionAlgorithm === 'GZIP') blob = Utilities.ungzip(blob.setContentType('application/x-gzip'));
    foldListingsReport_(by, l.m[1], blob.getDataAsString('UTF-8'));   // only ASCII columns are read, so encoding doesn't matter
  });
  var ts = latest.map(function (l) { return l.rep.createdTime; }).sort()[0];
  var markets = latest.map(function (l) { return l.m[1]; });
  saveSnapshot_(LISTINGS_CACHE, LISTINGS_TAB, { ts: ts, markets: markets, by: by });
  props.setProperty(LISTINGS_IDS_PROP, ids);
  Logger.log('listings snapshot (%s) as of %s: %s keys', markets.join(' '), ts, Object.keys(by).length);
}

/** Header-driven TSV fold (EU and JP reports have different columns; JP has no asin1, only product-id). */
function foldListingsReport_(by, code, text) {
  var lines = text.split(/\r?\n/);
  var head = lines[0].replace(/^\uFEFF/, '').split('\t');
  var col = function (names) { for (var i = 0; i < names.length; i++) { var j = head.indexOf(names[i]); if (j >= 0) return j; } return -1; };
  var iSku = col(['seller-sku']), iAsin = col(['asin1']), iPid = col(['product-id']), iPidType = col(['product-id-type']), iStatus = col(['status']);
  for (var r = 1; r < lines.length; r++) {
    if (!lines[r]) continue;
    var c = lines[r].split('\t');
    var sku = String(c[iSku] || '').trim().toUpperCase().replace(/^AMZN\.GR\./, '');   // Grade & Resell copies
    var asin = iAsin >= 0 ? c[iAsin] : (c[iPidType] === '1' ? c[iPid] : '');
    asin = String(asin || '').trim().toUpperCase();
    var m = sku.match(/^[A-Z]{3}\d{5}/);
    var key = m ? m[0] : (/^B0[A-Z0-9]{8}$/.test(asin) ? 'ASIN:' + asin : null);
    if (!key) continue;
    var o = by[key] = by[key] || { a: [], on: '', off: '' };
    if (/^B0[A-Z0-9]{8}$/.test(asin) && o.a.indexOf(asin) < 0) o.a.push(asin);
    var active = String(c[iStatus] || '').trim() === 'Active';
    var on = o.on ? o.on.split(' ') : [], off = o.off ? o.off.split(' ') : [];
    if (active && on.indexOf(code) < 0) { on.push(code); off = off.filter(function (x) { return x !== code; }); }
    if (!active && on.indexOf(code) < 0 && off.indexOf(code) < 0) off.push(code);
    o.on = on.join(' '); o.off = off.join(' ');
  }
}

/** Listing state for a product: its SKU key plus "ASIN:" keys for every ASIN it has. */
function listingFor_(p, lsnap) {
  var out = { on: [], off: [], asins: [], known: !!lsnap };
  if (!lsnap) return out;
  [p.sku].concat(p.asins.map(function (a) { return 'ASIN:' + a; })).forEach(function (k) {
    var o = k && lsnap.by[k];
    if (!o) return;
    addTo_(out.asins, o.a);
    addTo_(out.on, o.on ? o.on.split(' ') : []);
    addTo_(out.off, o.off ? o.off.split(' ') : []);
  });
  out.off = out.off.filter(function (c) { return out.on.indexOf(c) < 0; });
  return out;
}

/** An ASIN the masters don't know: find which SKU key carries it in the FBA / listings snapshots. */
function productsForAsin_(index, asin, snaps) {
  for (var s = 0; s < snaps.length; s++) {
    var by = snaps[s] && snaps[s].by;
    if (!by) continue;
    for (var k in by) {
      if (by[k].a.indexOf(asin) < 0) continue;
      if (k.indexOf('ASIN:') === 0) return [{ sku: '', asins: [asin], model: '', nameKo: '', nameEn: '', brand: '', device: '', color: '', category: '', status: '' }];
      var hit = index.filter(function (p) { return p.sku === k; });
      return hit.length ? hit : [{ sku: k, asins: [asin], model: '', nameKo: '', nameEn: '', brand: '', device: '', color: '', category: '', status: '' }];
    }
  }
  return [];
}

/** EAN/UPC → [ASIN] via Catalog Items (20 identifiers per call), DE then UK. */
function asinsByBarcode_(list) {
  var out = {};
  [['EAN', 'ean'], ['UPC', 'upc']].forEach(function (t) {
    var codes = [];
    list.forEach(function (p) { var v = String(p[t[1]] || '').trim(); if (/^\d{12,13}$/.test(v) && codes.indexOf(v) < 0 && !out[p.ean] && !out[p.upc]) codes.push(v); });
    CATALOG_MARKETS.forEach(function (mk) {
      var want = codes.filter(function (c) { return !out[c]; });
      for (var i = 0; i < want.length; i += 20) {
        var res = spapiRequest_('EU', 'get', '/catalog/2022-04-01/items',
          'identifiers=' + want.slice(i, i + 20).join(',') + '&identifiersType=' + t[0] + '&marketplaceIds=' + mk[0] + '&includedData=identifiers&pageSize=20');
        (res.items || []).forEach(function (it) {
          (it.identifiers || []).forEach(function (g) {
            (g.identifiers || []).forEach(function (id) {
              var v = String(id.identifier || '').replace(/^0+(?=\d{12,13}$)/, '');
              want.forEach(function (c) { if (v === c || v === '0' + c) { out[c] = out[c] || []; if (out[c].indexOf(it.asin) < 0) out[c].push(it.asin); } });
            });
          });
        });
      }
    });
  });
  return out;
}

function addAsins_(p, asins) { addTo_(p.asins, asins); }
function addTo_(arr, items) { (items || []).forEach(function (x) { if (x && arr.indexOf(x) < 0) arr.push(x); }); }

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

function buildProductsCard_(title, subtitle, list, inv, lst, thumbs, similar, snap, lsnap, notes, ms) {
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
      }).join(', ') + (p.asins.length > 3 ? ' 외 ' + (p.asins.length - 3) : '') +
        (p.viaBarcode ? ' <font color="#5f6368">(EAN/UPC 카탈로그 매칭)</font>' : ''));
    } else {
      lines.push('<font color="#5f6368">ASIN 없음 — 마스터·Seller Central 리스팅·EAN 카탈로그 모두 미검색</font>');
    }
    lines.push(listingLine_(lst[p.sku || p.asins[0]]));
    lines.push(inventoryLine_(inv[p.sku || p.asins[0]]));

    var textCol = { textParagraph: { text: lines.filter(String).join('<br>') } };
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
  var lAsOf = lsnap ? Utilities.formatDate(new Date(lsnap.ts), 'Asia/Seoul', 'M/d HH:mm') + ', ' + lsnap.markets.join('·') : '-';
  footer.push('<font color="#5f6368">🏷️ = Seller Central 리스팅 (' + lAsOf + ') · 📦 = FBA 판매가능 재고 (EU 국가별 + JP, ' + asOf + ' 기준) · NA 미지원 · ' + (ms / 1000).toFixed(1) + '초</font>');
  footer.push('<a href="https://docs.google.com/spreadsheets/d/' + SKU_MASTER_ID + '/edit">SKU master</a> · ' +
              '<a href="https://docs.google.com/spreadsheets/d/' + ASIN_MASTER_ID + '/edit">ASIN master</a>');
  sections.push({ widgets: [{ textParagraph: { text: footer.join('<br>') } }] });

  return {
    cardId: 'products-' + new Date().getTime(),
    card: { header: { title: title, subtitle: subtitle }, sections: sections }
  };
}

/** "🏷️ 활성 DE·GB·FR · 비활성 IT" — listing state per country in Seller Central. */
function listingLine_(o) {
  if (!o || !o.known) return '';
  if (!o.on.length && !o.off.length) return '🏷️ Seller Central 리스팅 없음 (EU·JP)';
  var order = function (a, b) { return COUNTRY_ORDER.indexOf(a) - COUNTRY_ORDER.indexOf(b); };
  var parts = [];
  if (o.on.length) parts.push('활성 ' + o.on.sort(order).join('·'));
  if (o.off.length) parts.push('<font color="#5f6368">비활성 ' + o.off.sort(order).join('·') + '</font>');
  return '🏷️ ' + parts.join(' · ');
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
