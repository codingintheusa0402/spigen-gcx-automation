/**
 * Query understanding for /products and /products-similar (사용자 지시 2026-10-18: exact matches for
 * /products, same-device matches for -similar, category words, typos accepted). Built from a pass over
 * the whole SKU master (8,976 rows) and ASIN master (6,744 rows):
 *
 * Identifiers people paste — all matched exactly, several at once allowed ("ACS08008, ACS08009"):
 *   master SKU AAA00000 (ACS/AGL/AMP/ACP/AFA…), legacy 000AA00000 / A00AA00000,
 *   seller SKU = master SKU + channel suffix (ACS08008UKO, …PAN, …EUP, …F, …JPP, amzn.gr.… copies) → its SKU,
 *   ASIN B0…, an Amazon URL (/dp/…, /gp/product/…), EAN/UPC (12–13 digits).
 *   (The ASIN master's ASIN column is junk on ~1/3 of rows — "B" or a monday id — so only B0… counts.)
 *
 * 기종 strings — 559 distinct, written several ways; deviceSet_ turns each into the devices it fits:
 *   multi-device "iPhone 18 Pro Max/17 Pro Max", "Galaxy Z Fold 8/Z Flip 8", "Galaxy Watch Ultra2/Ultra",
 *   "Pixel Watch 5/Watch 4/Watch 3(41mm)", "iPhone SE(2022/2020)/8/7", "Model Y/3"; a later part takes
 *   the words of the first part that come before the matching word (Ultra2 ↔ Ultra, Z Fold ↔ Z Flip) or,
 *   for a part starting with a number, before the first number ("Apple Watch Series 10" → "… 9").
 *   Slashes inside parentheses are alternatives, not devices: "iPad mini (2024/2021)",
 *   "Galaxy Watch9(44mm/40mm)", "MacBook Air 15" (M3,2024/M2,2023)". Size / year / chip qualifiers are
 *   optional: the bare device ("ipad mini", "apple watch series 10") matches every size and year, and
 *   naming one ("ipad mini 2024", "galaxy watch9 44mm") narrows to it. "#…/" decorations are ignored.
 *   Everything is compared as split words — letters/digits apart ("iphone17pro", "watch9", "s26+" =
 *   "s 26 plus") — so typed and stored forms meet in the middle.
 *
 * Free text = [device] [category words] [brand] [model] [colour], in any order, each typo-tolerant
 * (numbers exact). Model + colour must match exactly when such a product exists ("Ultra Hybrid" never
 * returns "Ultra Hybrid Mag Fit"); otherwise every word must appear in the product's names. Korean
 * device / model names are mapped to the English words (아이폰 17 프로 울하 → iphone 17 pro ultra hybrid),
 * and any other Korean word is looked up in 자재내역(KO).
 */

/** Category words → which products they keep. Several in one query = any of them. 대분류 / 자재 그룹 내역 / 담당부서코드 based. */
var CATEGORY_WORDS = [
  { key: 'case', label: '케이스', words: ['case', 'cases', 'cover', 'covers', '케이스', '커버'],
    test: function (p) { return /케이스$/.test(p.category) || /케이스/.test(p.group); } },
  { key: 'sp', label: '보호필름 (SP·CP)', words: ['screen protector', 'screen protectors', 'camera protector', 'camera protectors', 'lens protector',
      'protector', 'protectors', 'screen', 'sp', 'cp', 'glass', 'film', 'tempered glass', '필름', '보호필름', '액정필름', '글라스', '강화유리', '카메라필름'],
    test: function (p) { return /보호필름$/.test(p.category); } },
  { key: 'camera', label: '카메라 보호필름', words: ['camera', 'lens', 'optik', '카메라', '렌즈'],
    test: function (p) { return /보호필름$/.test(p.category) && /optik|camera|lens/i.test(p.model + ' ' + p.nameEn) || /카메라/.test(p.nameKo); } },
  { key: 'sda', label: 'SDA', words: ['sda', 'smart device acc', 'smart device accessories'], test: function (p) { return p.dept === 'A003'; } },
  { key: 'band', label: '밴드·스트랩', words: ['band', 'bands', 'strap', 'straps', 'watch band', '밴드', '스트랩', '시계줄'],
    test: function (p) { return p.category === '워치ACC' || /밴드/.test(p.group) || /band|strap/i.test(p.nameEn); } },
  { key: 'charger', label: '충전기', words: ['charger', 'chargers', 'charging', 'wireless charger', 'magsafe charger', '충전기', '무선충전기'],
    test: function (p) { return p.category === '충전기' || /충전기/.test(p.group); } },
  { key: 'cable', label: '케이블·어댑터', words: ['cable', 'cables', 'adapter', 'adaptor', 'gender', '케이블', '어댑터', '젠더'],
    test: function (p) { return p.category === '케이블/어댑터'; } },
  { key: 'battery', label: '배터리', words: ['battery', 'power bank', 'powerbank', '보조배터리', '배터리'], test: function (p) { return p.category === '배터리'; } },
  { key: 'mount', label: '거치대·스탠드', words: ['mount', 'holder', 'stand', 'car mount', 'grip', '거치대', '스탠드', '그립'],
    test: function (p) { return p.category === '거치대/스탠드' || /거치대|스탠드/.test(p.group); } },
  { key: 'car', label: '차량용품', words: ['car', 'tesla', 'vehicle', '차량', '차량용'], test: function (p) { return /^차량/.test(p.category); } },
  { key: 'audio', label: '음향기기·액세서리', words: ['earphone', 'earphones', 'earbuds', 'headphone', 'audio', '이어폰', '헤드폰'],
    test: function (p) { return /^음향기기/.test(p.category); } },
  { key: 'wallet', label: '패션잡화 (지갑·카드·가방)', words: ['wallet', 'card holder', 'card', 'bag', 'pouch', '지갑', '카드', '가방', '파우치'],
    test: function (p) { return p.category === '패션잡화/뷰티'; } }
];

var BRAND_WORDS = [
  { key: 'Spigen', words: ['spigen', '슈피겐'] }, { key: 'Caseology', words: ['caseology', '케이스올로지', '케올'] },
  { key: 'CYRILL', words: ['cyrill', 'cyril', '씨릴'] }, { key: 'Spigen Essential', words: ['essential'] },
  { key: 'PowerArc', words: ['powerarc'] }, { key: 'Spigen Audio', words: ['spigen audio'] }, { key: 'Legato', words: ['legato'] },
  { key: 'MAG-iT', words: ['mag-it', 'magit'] }, { key: 'MediSerin', words: ['mediserin'] }
];

// Spellings, short forms and Korean names folded into the English words before matching (one word → words).
var WORD_ALIASES = {
  iph: 'iphone', ip: 'iphone', ipone: 'iphone', iphon: 'iphone', '아이폰': 'iphone', '갤럭시': 'galaxy', '갤': 'galaxy', '픽셀': 'pixel',
  '아이패드': 'ipad', '애플워치': 'apple watch', '갤워치': 'galaxy watch', '에어팟': 'airpods', '맥북': 'macbook', '테슬라': 'tesla',
  '프로': 'pro', '맥스': 'max', '울트라': 'ultra', '플러스': 'plus', '에어': 'air', '폴드': 'fold', '플립': 'flip', '미니': 'mini',
  '울하': 'ultra hybrid', '울트라하이브리드': 'ultra hybrid', '하이브리드': 'hybrid', '맥핏': 'mag fit', '리퀴드에어': 'liquid air',
  '러기드아머': 'rugged armor', '터프아머': 'tough armor', '슬림아머': 'slim armor', '코어아머': 'core armor', '씬핏': 'thin fit',
  '크리스탈하이브리드': 'crystal hybrid', '크리스탈플렉스': 'crystal flex', '나노팝': 'nano pop', '옵틱': 'optik', '아머': 'armor',
  '블랙': 'black', '화이트': 'white', '클리어': 'clear', '크리스탈클리어': 'crystal clear', '매트블랙': 'matte black', '네이비': 'navy',
  '그레이': 'gray', '실버': 'silver', '블루': 'blue', '그린': 'green', '핑크': 'pink', '퍼플': 'purple', '베이지': 'beige'
};

var ID_PATTERNS = {
  sku: /^(?:AMZN\.GR\.)?([A-Z]{3}\d{5}|\d{3}[A-Z]{2}\d{5}|[A-Z]\d{2}[A-Z]{2}\d{5})(?:[A-Z]{0,5}|[-_].*)?$/,   // master / seller SKU
  asin: /^B0[A-Z0-9]{8}$/,
  url: /(?:\/dp\/|\/gp\/product\/|\/product-reviews\/|[?&]asin=)(B0[A-Z0-9]{8})/i,
  barcode: /^\d{12,13}$/
};

/* ===================== Identifiers ===================== */

/** "ACS08008, B0BYMJ3M29 https://amazon.de/dp/…" → [{type, value}] when EVERY token is an identifier, else null. */
function parseIdentifiers_(query) {
  var toks = String(query || '').trim().split(/[\s,;]+/).filter(String);
  if (!toks.length) return null;
  var out = [];
  for (var i = 0; i < toks.length; i++) {
    var t = toks[i], u = t.toUpperCase(), m;
    if ((m = t.match(ID_PATTERNS.url))) out.push({ type: 'asin', value: m[1].toUpperCase() });
    else if (ID_PATTERNS.asin.test(u)) out.push({ type: 'asin', value: u });
    else if (ID_PATTERNS.barcode.test(u)) out.push({ type: 'barcode', value: u });
    else if ((m = u.match(ID_PATTERNS.sku))) out.push({ type: 'sku', value: m[1], raw: u });
    else return null;
  }
  return out;
}

/** Exact products for identifiers, in the order typed; unknown ones come back as bare {sku}/{asins} stubs. */
function productsForIds_(index, ids) {
  var out = [], seen = {};
  ids.forEach(function (id) {
    var hits = index.filter(function (p) {
      return id.type === 'sku' ? p.sku === id.value : id.type === 'asin' ? p.asins.indexOf(id.value) >= 0 : (p.ean === id.value || p.upc === id.value);
    });
    if (!hits.length && id.type === 'sku') hits = [stubProduct_(id.value, [])];
    if (!hits.length && id.type === 'asin') hits = [stubProduct_('', [id.value])];
    hits.forEach(function (p) { var k = p.sku || p.asins[0]; if (!seen[k]) { seen[k] = 1; out.push(p); } });
  });
  return out;
}

function stubProduct_(sku, asins) {
  return { sku: sku, asins: asins, model: '', nameKo: '', nameEn: '', brand: '', device: '', devs: [], color: '', category: '', group: '', status: '', dept: '', stub: true };
}

/* ===================== Devices ===================== */

/**
 * Every device a 기종 value fits, as split words, with and without its qualifiers:
 *   "iPhone 18 Pro Max/17 Pro Max"            → iphone 18 pro max | iphone 17 pro max
 *   "#Apple Watch Series 10(42mm)/9(41mm)/"   → apple watch series 10 | … 10 · 42 mm | apple watch series 9 | … 9 · 41 mm
 *   "iPad mini (2024/2021)"                   → ipad mini | ipad mini · 2024 | ipad mini · 2021
 *   "Galaxy Watch Ultra2/Ultra"               → galaxy watch ultra 2 | galaxy watch ultra
 *   "Galaxy S26+/S26"                         → galaxy s 26 plus | galaxy s 26
 */
function deviceSet_(d) {
  d = String(d || '').replace(/#/g, '').trim();
  if (!isRealDevice_(d)) return [];
  var parts = [], depth = 0, cur = '';
  for (var i = 0; i < d.length; i++) {   // split on "/" outside parentheses
    var ch = d[i];
    if (ch === '(') depth++;
    if (ch === ')') depth = Math.max(0, depth - 1);
    if (ch === '/' && !depth) { parts.push(cur); cur = ''; } else cur += ch;
  }
  parts.push(cur);
  parts = parts.map(function (x) { return x.trim(); }).filter(String);
  if (!parts.length) return [];

  var parsed = parts.map(function (x) {
    var m = x.match(/^(.*?)\s*\((.*)\)\s*(.*)$/);   // base (alt/alt) trailing
    if (!m) return { base: x, quals: [] };
    return { base: (m[1] + ' ' + m[3]).trim(), quals: m[2].split('/').map(function (q) { return q.trim(); }).filter(String) };
  });
  var first = parsed[0].base.split(/\s+/);
  var out = [];
  parsed.forEach(function (pp, i) {
    var base = pp.base;
    if (i) {
      var w = base.split(/\s+/);
      var k = -1;
      if (/^\d/.test(w[0])) {   // "17 Pro Max" after "iPhone 18 Pro Max": prefix = words before the first number
        for (var j = 0; j < first.length; j++) if (/\d/.test(first[j])) { k = j; break; }
        if (k < 0) k = 1;
      } else {                  // "Z Flip 8" after "Galaxy Z Fold 8", "Ultra" after "Galaxy Watch Ultra2"
        for (j = 0; j < first.length; j++) if (stem_(first[j]) === stem_(w[0])) { k = j; break; }
        if (k < 0) k = 1;
      }
      base = first.slice(0, k).concat(w).join(' ');
    }
    var b = splitWords_(base).join(' ');
    if (b && out.indexOf(b) < 0) out.push(b);
    pp.quals.forEach(function (q) {
      var bq = b + ' · ' + splitWords_(q).join(' ');   // " · " marks a qualifier (size / year / chip)
      if (out.indexOf(bq) < 0) out.push(bq);
    });
  });
  return out;
}

/** The entries that say exactly what a product fits: qualified ones replace their bare device. */
function mostSpecific_(devs) {
  return (devs || []).filter(function (d) { return d.indexOf(' · ') >= 0 || !devs.some(function (o) { return o.indexOf(d + ' · ') === 0; }); });
}

/**
 * Does product p fit any target device? A bare target ("apple watch series 10") takes every size of it;
 * a qualified one ("… · 42 mm") takes that size plus products made for the device without a size.
 */
function devFits_(p, targets) {
  var devs = p.devs || [];
  return targets.some(function (t) {
    if (devs.indexOf(t) >= 0) return true;
    var bare = t.split(' · ')[0];
    return bare !== t && mostSpecific_(devs).indexOf(bare) >= 0;
  });
}

function stem_(w) { return String(w || '').toLowerCase().replace(/[\d+]/g, ''); }

function isRealDevice_(d) {
  d = String(d || '').replace(/#|\//g, '').trim();
  return !!d && !/^(common|공용|미정|-|n\/?a)$/i.test(d);
}

/** { device words → true } for every device in the index (cached on the index array for this execution). */
function knownDevices_(index) {
  if (index._devices) return index._devices;
  var seen = {};
  index.forEach(function (p) { (p.devs || []).forEach(function (d) { seen[d] = true; }); });
  return (index._devices = seen);
}

/** Display text for a device: "iphone 17 pro" → "iPhone 17 Pro", "ipad mini · 2024" → "iPad mini (2024)" (casing from the master). */
function deviceLabel_(index, dev) {
  var bare = dev.split(' · ')[0], qual = dev.split(' · ')[1];
  var best = null;
  for (var i = 0; i < index.length; i++) {
    var ds = index[i].devs || [];
    if (ds[0] !== bare) continue;
    var raw = String(index[i].device).replace(/#/g, '').split('/')[0].replace(/\(.*$/, '').trim();   // first device, no qualifiers
    if (!best || raw.length < best.length) best = raw;
  }
  best = best || bare.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  return qual ? best + ' (' + qual.replace(/(\d) (mm|nd|rd|th|st)\b/g, '$1$2') + ')' : best;
}

/* ===================== Free-text query ===================== */

/** Lower-cased words: punctuation dropped, letters/digits apart, "+" → plus, aliases expanded. */
function splitWords_(s) {
  var t = String(s || '').toLowerCase().replace(/["#,()\[\]]/g, ' ').replace(/\+/g, ' plus ')
    .replace(/([a-z])(\d)/g, '$1 $2').replace(/(\d)([a-z]{2,})/g, '$1 $2');
  var out = [];
  t.split(/\s+/).filter(String).forEach(function (w) {
    var a = WORD_ALIASES[w];
    (a ? a.split(' ') : [w]).forEach(function (x) { out.push(x); });
  });
  return out;
}

/**
 * Query → { device, cats[], brand, words[] }: the longest device named (typo-tolerant, numbers exact),
 * any category words, a brand, and what's left (model / colour / anything else).
 */
function parseProductQuery_(index, query) {
  var words = splitWords_(query);
  var take = function (phrases) {   // first phrase found (longest first) is removed from words → true
    phrases = phrases.slice().sort(function (x, y) { return y.split(' ').length - x.split(' ').length; });
    for (var n = 0; n < phrases.length; n++) {
      var ws = splitWords_(phrases[n]);
      for (var i = 0; i + ws.length <= words.length; i++) {
        if (ws.every(function (t, k) { return wordMatch_(words[i + k], t); })) { words.splice(i, ws.length); return true; }
      }
    }
    return false;
  };

  var device = null, devs = Object.keys(knownDevices_(index));
  devs.sort(function (a, b) { return b.split(' ').length - a.split(' ').length || b.length - a.length; });
  for (var d = 0; d < devs.length && !device; d++) {
    var ds = devs[d].replace(/ · /g, ' ').split(' ');
    for (var i = 0; i + ds.length <= words.length; i++) {
      if (ds.every(function (t, k) { return wordMatch_(words[i + k], t); })) { device = devs[d]; words.splice(i, ds.length); break; }
    }
  }
  var cats = CATEGORY_WORDS.filter(function (c) { return take(c.words); });
  var brand = null;
  BRAND_WORDS.some(function (b) { if (take(b.words)) { brand = b.key; return true; } return false; });
  return { device: device, cats: cats, brand: brand, words: words };
}

/** Same word allowing small typos: numbers / ≤2 letters exact, 3–5 letters 1 edit, longer 2 edits. */
function wordMatch_(a, b) {
  if (a === b) return true;
  if (!a || !b || /\d/.test(a) || /\d/.test(b) || /[^\x00-\x7f]/.test(a + b)) return false;
  var n = Math.min(a.length, b.length);
  var max = n <= 2 ? 0 : n <= 5 ? 1 : 2;
  return max > 0 && Math.abs(a.length - b.length) <= max && editDistance_(a, b) <= max;
}

function editDistance_(a, b) {   // Damerau: an adjacent swap is one edit ("iphnoe")
  var d = [];
  for (var i = 0; i <= a.length; i++) d[i] = [i];
  for (var j = 1; j <= b.length; j++) d[0][j] = j;
  for (i = 1; i <= a.length; i++) {
    for (j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
    }
  }
  return d[a.length][b.length];
}

/** Does a product fit the parsed device / categories / brand? */
function productFits_(p, pq) {
  if (pq.device && !devFits_(p, [pq.device])) return false;
  if (pq.cats.length && !pq.cats.some(function (c) { return c.test(p); })) return false;
  if (pq.brand && p.brand !== pq.brand && !(pq.brand === 'Spigen' && /^Spigen/.test(p.brand))) return false;
  return true;
}

/**
 * How well the leftover words name this product: 2 = exactly its model (+ its colour words), 1 = every
 * word appears in its names (Korean words in 자재내역), 0 = no.
 */
function wordScore_(p, words) {
  if (!words.length) return 2;
  var model = splitWords_(p.model), color = splitWords_(p.color);
  for (var i = 0; i + model.length <= words.length; i++) {
    if (!model.length || !model.every(function (t, k) { return wordMatch_(words[i + k], t); })) continue;
    var rest = words.slice(0, i).concat(words.slice(i + model.length));
    if (rest.every(function (w) { return color.some(function (c) { return wordMatch_(w, c); }); })) return 2;
  }
  var en = splitWords_([p.model, p.nameEn, p.color, p.sku].join(' '));
  var ko = String(p.nameKo || '').toLowerCase().replace(/\s+/g, '');
  return words.every(function (w) {
    return /[^\x00-\x7f]/.test(w) ? ko.indexOf(w) >= 0 : en.some(function (t) { return wordMatch_(w, t); });
  }) ? 1 : 0;
}

/* ===================== The two searches ===================== */

/** /products: identifiers exactly; free text → exact model tier when one exists, else all-words tier. */
function searchProducts_(index, query, stockOf) {
  var ids = parseIdentifiers_(query);
  if (ids) return productsForIds_(index, ids).filter(function (p) { return !p.stub; });
  var pq = parseProductQuery_(index, query);
  if (!pq.device && !pq.words.length && !pq.brand) return [];   // a category word alone is not a search
  var scored = index.filter(function (p) { return productFits_(p, pq); })
    .map(function (p) { return { p: p, s: wordScore_(p, pq.words) }; }).filter(function (x) { return x.s > 0; });
  var top = scored.some(function (x) { return x.s === 2; }) ? 2 : 1;
  var hits = scored.filter(function (x) { return x.s === top; }).map(function (x) { return x.p; });
  hits.forEach(function (p) { p._stock = stockOf(p); });
  hits.sort(function (a, b) {   // most FBA stock first, then 품목상태 A before B, then SKU
    return b._stock - a._stock || (a.status === 'B') - (b.status === 'B') || a.sku.localeCompare(b.sku);
  });
  return hits;
}

/**
 * /products-similar: products for the same device. The device is the one named ("ipone 17 pro case"),
 * else the device(s) of the identified / matched product. Category words and a brand narrow it.
 * → { base, label, cats, same[] } or { error }.
 */
function similarProducts_(index, query, stockOf) {
  var toks = String(query || '').trim().split(/\s+/);
  var ids = parseIdentifiers_(toks[0]);
  var pq = parseProductQuery_(index, ids ? toks.slice(1).join(' ') : query);
  var devs, base = null, label;
  if (!ids && pq.device && !pq.words.length) {
    devs = [pq.device];
    label = deviceLabel_(index, pq.device);
  } else {
    var matches = ids ? productsForIds_(index, ids).filter(function (p) { return !p.stub; }) : searchProducts_(index, query, stockOf);
    base = matches.filter(function (p) { return (p.devs || []).length; })[0];
    if (!base) {
      if (matches.length) {   // found, but 기종 is Common/blank — a universal accessory has no device to match on
        return { error: '⚠️ "' + query + '" (' + matches[0].sku + ') 는 기종이 "' + (matches[0].device || '미지정') + '" 인 공용 제품이라 같은 기종 제품을 찾을 수 없습니다. /products 로 조회해 주세요.' };
      }
      return { error: '⚠️ "' + query + '" 에서 제품이나 기종을 찾지 못했습니다. SKU, ASIN 또는 기종명(예: iPhone 17 Pro, 뒤에 case / screen protector / SDA)으로 입력해 주세요.' };
    }
    devs = mostSpecific_(base.devs);   // a 42 mm watch case's siblings are 42 mm or size-free, never 46 mm
    label = String(base.device).replace(/#/g, '').replace(/\/+$/, '');
  }
  var filt = { device: null, cats: pq.cats, brand: pq.brand };
  var same = index.filter(function (p) {
    return (!base || p.sku !== base.sku) && productFits_(p, filt) && devFits_(p, devs);
  });
  same.forEach(function (p) { p._stock = stockOf(p); });
  same.sort(function (a, b) {   // in stock first, then by category / model so a case's siblings sit together
    return (b._stock > 0) - (a._stock > 0) || (a.status === 'B') - (b.status === 'B') ||
           a.category.localeCompare(b.category) || b._stock - a._stock || a.model.localeCompare(b.model) || a.sku.localeCompare(b.sku);
  });
  return { base: base, label: label, cats: pq.cats, brand: pq.brand, same: same };
}
