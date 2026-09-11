/**
 * BadReview — Google Chat app (interactive twin of ../badreview_chat_report.py).
 *
 * One deployment = one product (see Config.gs → APP_PRODUCT):
 *   chat_app/      김지우 Kevin 글로벌CX전략팀 → Pixel 11 Series
 *   chat_app_jane/ 나아름 Jane 글로벌CX전략팀  → Galaxy Z8 Series
 *
 * Control card:  시작일 / 종료일 (DATE_ONLY; default = earliest `Update 날짜` on the
 *                '1-3점' tab → today), 국가 dropdown, 기종 dropdown, [조회].
 * Report card:   ✔️ M/D(요일)~M/D(요일) <product> 배드리뷰 (1~3점) (총 N건)
 *                "Top 5 인입사유" 2 cols by 대분류 (red 보호필름 / blue 케이스)
 *                "기간 내 최다 인입사유" + fixed 5-line list, [배드리뷰] link button.
 * EVERY number is scoped to the picked range + 국가/기종 filters (unlike the webhook
 * card, whose Top 5 is cumulative).
 *
 * Runs as a Google Workspace add-on (Chat API config checkbox, irreversible per GCP
 * project) → replies must be wrapped in the add-on action envelope; see chatCreate_.
 */

var PRODUCTS = {
  pixel11: {
    name: 'Pixel 11 Series',
    sheetId: '12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI',
    subtitle: '고객 리뷰 ★1~3점 · 26/8/18~26/11/18',
    cardId: 'pixel11-badreview',
    img: 'https://encrypted-tbn2.gstatic.com/shopping?q=tbn:ANd9GcQhY2aafxhQi-vGv0oxV5j0' +
         'kiiOF2sGF0hwiXeEePaAI3DbRziTZcO4Z2sehnyCpp1_qSxCn_iAE4IZ0SlW9WftxRQLxykwNXmmsDn' +
         'm3CQkubwlCmO7PL4F3JbUKGWpl1F6c2RuVw&usqp=CAc'
  },
  glxz8: {
    name: 'Galaxy Z8 Series',
    sheetId: '19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4',
    subtitle: '고객 리뷰 ★1~3점 · 26/7/27~26/10/27',
    cardId: 'glxz8-badreview',
    img: 'https://encrypted-tbn0.gstatic.com/shopping?q=tbn:ANd9GcRA_H2PEgYRDyPvE2kQ4RQ' +
         'lhN4sOnJd5cIS90muwFk2pqDlNNPQGlXJ8DHo7ihE20nzMs6-C5AUao7n5SgfGH4vTuzzrg6Oh_4QGU' +
         'SKReAVwWyaJK1DN9MI_TrAT8dE3Gq1-Rzobw&usqp=CAc'
  }
};

var GID_13 = 970309432;                                   // '1-3점' tab
var COL = { date: 'Update 날짜', dateAlt: 'Exported Date', tag: '인입사유(tag)',
            cat: '대분류', country: '국가(tag)', device: '기종명' };
var CAT_COLOR = { '휴대폰보호필름': '#EA4335', '휴대폰케이스': '#4285F4' };
// 인입사유(tag) values dropped before any counting (user rule 2026-09-08:
// exclude 긍정 리뷰 from every stat and every card, permanently). Mirror of
// EXCLUDED_TAGS in ../badreview_chat_report.py — keep in sync.
var EXCLUDED_TAGS = ['긍정 리뷰'];
var ALL = '__all__';                                      // dropdown "전체"
var WD = ['일', '월', '화', '수', '목', '금', '토'];      // JS getDay(): 0 = Sun
var TZ = 'Asia/Seoul';

/* ===================== Chat event handlers ===================== */

function chatCreate_(message) {
  return { hostAppDataAction: { chatDataAction: { createMessageAction: { message: message } } } };
}
function chatUpdate_(message) {
  return { hostAppDataAction: { chatDataAction: { updateMessageAction: { message: message } } } };
}

function onMessage(event) {
  var q = defaultQuery_();
  var range = parseRangeFromText_(messageText_(event));
  if (range) { q.start = range[0]; q.end = range[1]; }
  return chatCreate_({ cardsV2: buildCards_(q) });
}

/** Message text under either event shape (classic: event.message; add-on: event.chat.messagePayload). */
function messageText_(event) {
  var m = (event && event.message) ||
          (event && event.chat && event.chat.messagePayload && event.chat.messagePayload.message) || {};
  return m.argumentText || m.text || '';
}

function onAddToSpace(event) {
  return chatCreate_({
    text: PRODUCTS[APP_PRODUCT].name + ' 배드리뷰(1~3점) 리포트 앱입니다. 기간·국가·기종을 고르고 [조회]하세요.',
    cardsV2: buildCards_(defaultQuery_())
  });
}

function onRemoveFromSpace(event) {}

/** Classic-mode shim; add-on mode calls refreshReport directly. */
function onCardClick(event) {
  var fn = (event.common && event.common.invokedFunction) ||
           (event.action && event.action.actionMethodName) || '';
  if (fn === 'refreshReport') return refreshReport(event);
  return chatUpdate_({ cardsV2: buildCards_(defaultQuery_()) });
}

/** [조회] → onClick.action.function = "refreshReport". */
function refreshReport(event) {
  var inputs = formInputs_(event);
  // The pickers render empty after every re-render (valueMsEpoch is ignored in add-on
  // mode), so the card's current query travels in the button's action.parameters and is
  // the fallback whenever a picker was left blank — re-pressing [조회] keeps the range.
  var prev = actionParams_(event);
  var q = defaultQuery_();
  if (prev.start) q.start = toDate_(prev.start) || q.start;
  if (prev.end)   q.end   = toDate_(prev.end)   || q.end;
  var s = extractDateMs_(inputs, 'startDate'), e = extractDateMs_(inputs, 'endDate');
  if (s != null) q.start = utcMsToLocalDate_(s);
  if (e != null) q.end = utcMsToLocalDate_(e);
  if (q.start > q.end) { var t = q.start; q.start = q.end; q.end = t; }
  q.country = extractString_(inputs, 'country') || prev.country || ALL;
  q.device  = extractString_(inputs, 'device')  || prev.device  || ALL;
  return chatUpdate_({ cardsV2: buildCards_(q) });
}

/** onClick.action.parameters → {key: value}, under either event shape. */
function actionParams_(event) {
  var list = (event.common && event.common.parameters) ||
             (event.commonEventObject && event.commonEventObject.parameters) ||
             (event.action && event.action.parameters) || {};
  if (Array.isArray(list)) {
    var o = {};
    list.forEach(function (p) { o[p.key] = p.value; });
    return o;
  }
  return list;   // commonEventObject.parameters is already a {key: value} map
}

/* ===================== Query defaults ===================== */

/** { start: earliest Update 날짜 on the sheet, end: today (KST), country/device: 전체 } */
function defaultQuery_() {
  var sheet = loadSheet_();
  return { start: sheet.minDate || todayKst_(), end: todayKst_(), country: ALL, device: ALL };
}

/* ===================== Form input parsing ===================== */

/**
 * Chat interaction events: event.common.formInputs[name][""] → {dateInput|stringInputs}
 * Workspace add-on event object: event.commonEventObject.formInputs[name] → same inner shape.
 */
function formInputs_(event) {
  return (event.common && event.common.formInputs) ||
         (event.commonEventObject && event.commonEventObject.formInputs) ||
         event.formInputs || {};
}

function inputField_(inputs, name) {
  var v = inputs && inputs[name];
  if (!v) return null;
  return v[''] || v;
}

function extractDateMs_(inputs, name) {
  var f = inputField_(inputs, name);
  if (!f) return null;
  var di = f.dateInput || f.dateTimeInput;
  if (di && di.msSinceEpoch != null && di.msSinceEpoch !== '') return Number(di.msSinceEpoch);
  return null;
}

function extractString_(inputs, name) {
  var f = inputField_(inputs, name);
  if (!f) return '';
  var si = f.stringInputs;
  if (si && si.value && si.value.length) return String(si.value[0]);
  return '';
}

/** DATE_ONLY gives UTC midnight of the picked day; keep that calendar day. */
function utcMsToLocalDate_(ms) {
  var d = new Date(Number(ms));
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

/* ===================== Card building ===================== */

function buildCards_(q) {
  var sheet = loadSheet_();
  return [controlCard_(q, sheet), reportCard_(q, sheet)];
}

function controlCard_(q, sheet) {
  var p = PRODUCTS[APP_PRODUCT];
  var toMs = function (d) { return String(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); };
  var items = function (values, selected, allLabel) {
    var out = [{ text: allLabel, value: ALL, selected: selected === ALL }];
    values.forEach(function (v) { out.push({ text: v, value: v, selected: selected === v }); });
    return out;
  };
  return {
    cardId: 'badreview-control',
    card: {
      header: {
        title: p.name + ' 배드리뷰 리포트 조회',
        subtitle: '기본 기간 ' + fmtYmd_(sheet.minDate || todayKst_()) + ' ~ ' + fmtYmd_(todayKst_()) +
                  ' · 기간·국가·기종을 고르고 [조회]'
      },
      sections: [{
        widgets: [
          { columns: { columnItems: [
            { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [
              { dateTimePicker: { name: 'startDate', label: '시작일 (Update 날짜)',
                                  type: 'DATE_ONLY', valueMsEpoch: toMs(q.start) } } ] },
            { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [
              { dateTimePicker: { name: 'endDate', label: '종료일 (Update 날짜)',
                                  type: 'DATE_ONLY', valueMsEpoch: toMs(q.end) } } ] }
          ]}},
          { columns: { columnItems: [
            { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [
              { selectionInput: { name: 'country', label: '국가', type: 'DROPDOWN',
                                  items: items(sheet.countries, q.country, '전체 국가') } } ] },
            { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', widgets: [
              { selectionInput: { name: 'device', label: '기종', type: 'DROPDOWN',
                                  items: items(sheet.devices, q.device, '전체 기종') } } ] }
          ]}},
          { buttonList: { buttons: [
              { text: '조회', type: 'FILLED', onClick: { action: {
                  function: 'refreshReport',
                  parameters: [                       // current query → fallback for blank pickers
                    { key: 'start',   value: fmtYmd_(q.start) },
                    { key: 'end',     value: fmtYmd_(q.end) },
                    { key: 'country', value: q.country },
                    { key: 'device',  value: q.device }
                  ]
              } } }
          ]}}
        ]
      }]
    }
  };
}

function reportCard_(q, sheet) {
  var p = PRODUCTS[APP_PRODUCT];
  var data = crunch_(sheet, q);
  var period = fmtMd_(q.start) + (sameDay_(q.start, q.end) ? '' : '~' + fmtMd_(q.end));
  var filt = [];
  if (q.country !== ALL) filt.push(q.country);
  if (q.device !== ALL) filt.push(q.device);
  var link = 'https://docs.google.com/spreadsheets/d/' + p.sheetId +
             '/edit?gid=' + GID_13 + '#gid=' + GID_13;

  var top5Section = {
    header: 'Top 5 인입사유 (' + period + (filt.length ? ' · ' + filt.join(' · ') : '') + ')',
    widgets: [{ columns: { columnItems: [
      catColumn_('휴대폰보호필름', data.film),
      catColumn_('휴대폰케이스', data.box)
    ]}}]
  };

  var dayWidgets;
  if (data.tags.length) {
    var lines = data.tags.slice(0, 5).map(function (t, i) {
      return (i + 1) + '. ' + t[0] + ' &nbsp;' + t[1] + '건';
    });
    if (data.tags.length > 5) {
      var extra = data.tags.slice(5).reduce(function (s, x) { return s + x[1]; }, 0);
      lines[lines.length - 1] += ' &nbsp;…외 ' + extra + '건';
    }
    dayWidgets = [
      { decoratedText: {
          topLabel: '인입사유(tag) 기준',
          text: '<b>' + data.tags[0][0] + '</b> — ' + data.tags[0][1] + '건',
          startIcon: { knownIcon: 'STAR' }
      }},
      { textParagraph: { text: pad_(lines, 5) } }
    ];
  } else {
    dayWidgets = [
      { decoratedText: {
          topLabel: '인입사유(tag) 기준',
          text: '조건에 맞는 배드리뷰 없음',
          startIcon: { knownIcon: 'STAR' }
      }},
      { textParagraph: { text: pad_([], 5) } }
    ];
  }
  dayWidgets.push({ buttonList: { buttons: [
    { text: '배드리뷰', onClick: { openLink: { url: link } } }
  ]}});

  return {
    cardId: p.cardId,
    card: {
      header: {
        title: '✔️ ' + period + ' ' + p.name + ' 배드리뷰 (1~3점) (총 ' + data.count + '건)',
        subtitle: p.subtitle + (filt.length ? ' · ' + filt.join(' · ') : ''),
        imageUrl: p.img,
        imageType: 'SQUARE'
      },
      sections: [ top5Section, { header: period + ' 최다 인입사유', widgets: dayWidgets } ]
    }
  };
}

function catColumn_(label, blk) {
  var color = CAT_COLOR[label] || '#202124';
  var widgets = [{ textParagraph: {
    text: '<b><font color="' + color + '">' + label + '</font></b>  ·  ' + blk.tot + '건'
  }}];
  for (var i = 0; i < 5; i++) {
    if (i < blk.top5.length) {
      var n = blk.top5[i][0], c = blk.top5[i][1];
      var pct = blk.tot ? Math.round(c * 100 / blk.tot) + '%' : '-';
      widgets.push({ decoratedText: {
        topLabel: (i + 1) + '위', text: '<b>' + n + '</b>', bottomLabel: c + '건 · ' + pct
      }});
    } else {
      widgets.push({ decoratedText: { topLabel: ' ', text: ' ', bottomLabel: ' ' } });
    }
  }
  return {
    horizontalSizeStyle: 'FILL_AVAILABLE_SPACE',
    horizontalAlignment: 'START',
    verticalAlignment: 'TOP',
    widgets: widgets
  };
}

/* ===================== Sheet loading & crunching ===================== */

var SHEET_CACHE_ = null;   // per-execution memo (one Sheets read per interaction)

/**
 * Reads the product's '1-3점' tab once and normalises the rows:
 *   rows: [{date: Date|null, tag, cat, country, device}] (excluded tags dropped)
 *   minDate: earliest Update 날짜; countries / devices: sorted distinct values.
 */
function loadSheet_() {
  if (SHEET_CACHE_) return SHEET_CACHE_;
  var p = PRODUCTS[APP_PRODUCT];
  var vals = SpreadsheetApp.openById(p.sheetId).getSheetByName('1-3점').getDataRange().getValues();
  var H = vals[0];
  var iU = H.indexOf(COL.date); if (iU < 0) iU = H.indexOf(COL.dateAlt);
  var iT = H.indexOf(COL.tag), iC = H.indexOf(COL.cat);
  var iN = H.indexOf(COL.country), iD = H.indexOf(COL.device);

  var rows = [], minDate = null, countries = {}, devices = {};
  for (var r = 1; r < vals.length; r++) {
    var row = vals[r];
    var tag = String(row[iT] || '').trim() || '(빈칸)';
    if (EXCLUDED_TAGS.indexOf(tag) !== -1) continue;
    var d = toDate_(row[iU]);
    var country = iN >= 0 ? String(row[iN] || '').trim() : '';
    var device  = iD >= 0 ? String(row[iD] || '').trim() : '';
    if (country) countries[country] = (countries[country] || 0) + 1;
    if (device)  devices[device]   = (devices[device]   || 0) + 1;
    if (d && (!minDate || d < minDate)) minDate = d;
    rows.push({ date: d, tag: tag, cat: String(row[iC] || '').trim(), country: country, device: device });
  }
  var byCount = function (o) { return Object.keys(o).sort(function (a, b) { return o[b] - o[a]; }); };
  SHEET_CACHE_ = { rows: rows, minDate: minDate, countries: byCount(countries), devices: byCount(devices) };
  return SHEET_CACHE_;
}

/** Counts within [q.start, q.end] (inclusive, calendar days) and the 국가/기종 filters. */
function crunch_(sheet, q) {
  var lo = dayKey_(q.start), hi = dayKey_(q.end);
  var count = 0, tally = {};
  var cat = { '휴대폰보호필름': {}, '휴대폰케이스': {} };
  sheet.rows.forEach(function (x) {
    if (!x.date) return;
    var k = dayKey_(x.date);
    if (k < lo || k > hi) return;
    if (q.country !== ALL && x.country !== q.country) return;
    if (q.device  !== ALL && x.device  !== q.device)  return;
    count++;
    tally[x.tag] = (tally[x.tag] || 0) + 1;
    if (cat[x.cat]) cat[x.cat][x.tag] = (cat[x.cat][x.tag] || 0) + 1;
  });
  return {
    count: count,
    tags: sortDesc_(tally),
    film: block_(cat['휴대폰보호필름']),
    box: block_(cat['휴대폰케이스'])
  };
}

/** Sheet cell (Date object or "2026. 8. 19"-style text) → local Date at midnight, or null. */
function toDate_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return new Date(cell.getFullYear(), cell.getMonth(), cell.getDate());
  }
  var m = String(cell || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function dayKey_(d) { return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate(); }
function sameDay_(a, b) { return dayKey_(a) === dayKey_(b); }

function block_(obj) {
  var rows = sortDesc_(obj);
  return { tot: rows.reduce(function (s, x) { return s + x[1]; }, 0), top5: rows.slice(0, 5) };
}

function sortDesc_(obj) {
  return Object.keys(obj)
    .map(function (k) { return [k, obj[k]]; })
    .sort(function (a, b) { return b[1] - a[1]; });
}

/* ===================== Small helpers ===================== */

function todayKst_() {
  var s = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd').split('-');
  return new Date(Number(s[0]), Number(s[1]) - 1, Number(s[2]));
}

function fmtMd_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ')';
}

function fmtYmd_(d) {
  return d.getFullYear() + '. ' + (d.getMonth() + 1) + '. ' + d.getDate();
}

function pad_(lines, n) {
  var out = lines.slice(0, n);
  while (out.length < n) out.push('&nbsp;');
  return out.join('<br>');
}

/**
 * Message text → [start, end] or null.
 *   "9/1~9/11", "9/1 - 9/11", "2026-09-01~2026-09-11" → that range
 *   "9/11"                                            → single day
 */
function parseRangeFromText_(text) {
  var dates = [];
  var re = /(\d{4})\D(\d{1,2})\D(\d{1,2})|(\d{1,2})\s*[\/.]\s*(\d{1,2})/g, m;
  var y = todayKst_().getFullYear();
  while ((m = re.exec(String(text))) && dates.length < 2) {
    dates.push(m[1] ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
                    : new Date(y, Number(m[4]) - 1, Number(m[5])));
  }
  if (!dates.length) return null;
  if (dates.length === 1) return [dates[0], dates[0]];
  return dates[0] <= dates[1] ? [dates[0], dates[1]] : [dates[1], dates[0]];
}
