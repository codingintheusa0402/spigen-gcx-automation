/*************************************************
 * CONFIG — Pixel 11 Case+CP (Monday → Sheet, full refresh)
 * Sibling of GlxZ8_MondayToSheet (Galaxy Z8). Same sheet layout, different board.
 *************************************************/
const BOARD_ID = 18425190666; // 📌Pixel 11 Case+CP
const SHEET_NAME = '해외&국내 리뷰+클레임 데이터';
const PAGE_LIMIT = 500;
const RUN_TRIGGER_HOUR = 17; // 5PM KST

/**
 * Sheet header (in order) ↔ Monday column id.
 *   colId=null      → Monday item name (title field).
 *   transform(v)    → optional post-processing of the display value before it is written.
 *
 * NOTE (Pixel 11 board): unlike the Galaxy Z8 board there is no "사진 유무 (Clean)" formula column
 * (formula_mm7184wc). The same Y/N cleanup is applied here in-script on the "사진 유무" status
 * column (color_mm0ffgz8) so the sheet output stays identical to the Z8 sheet.
 */
const COLUMN_MAP = [
  { header: 'Order ID',       colId: null },
  { header: 'Created 날짜',    colId: 'date_mm0f80th' },
  { header: 'Purchased 날짜',  colId: 'date_mm59ejfp' },
  { header: 'ASIN',           colId: 'text_mm0f1q4h' },
  { header: 'SKU',             colId: 'lookup_mm0fv615' },
  { header: '대분류',          colId: 'lookup_mm0ffq8f' },
  { header: '인입사유',        colId: 'formula_mm0g81mb' },
  { header: '국가',            colId: 'formula_mm25vbf0' },
  { header: '기종명',          colId: 'lookup_mm0f6j81' },
  { header: '모델명',          colId: 'lookup_mm0fn79' },
  { header: '색상명',          colId: 'lookup_mm0fg6ja' },
  { header: '클레임/리뷰',     colId: 'color_mm0f7bwq' },
  { header: '생산업체',        colId: 'lookup_mm0feh3b' },
  { header: '원산지',          colId: 'lookup_mm0fahcy' },
  { header: '고객 대응',       colId: 'color_mm0fjzar' },
  { header: 'Review Link',    colId: 'link_mm0fkspz' },
  { header: 'Zendesk Ticket', colId: 'integration_mm0fzmv0' },
  { header: '데이터 출처',     colId: 'formula_mm5hrmzb' },
  { header: '사진 유무',       colId: 'color_mm0ffgz8', transform: _photoYN_ },
  { header: 'Review Ratings', colId: 'text_mm0fn5c0' }
];

const FORMULA_COL_IDS = COLUMN_MAP
  .filter(function(c){ return c.colId && c.colId.indexOf('formula_') === 0; })
  .map(function(c){ return c.colId; });

const HEADER = ['item_id'].concat(COLUMN_MAP.map(function(c){ return c.header; }));

/** Mirrors the Z8 board formula: SUBSTITUTE(SUBSTITUTE({사진 유무}, "yes", "Y"), "no", "N") */
function _photoYN_(v){
  return String(v == null ? '' : v).replace(/yes/g, 'Y').replace(/no/g, 'N');
}

/** CacheService keys used to pass progress from the running sync to the popup dialog. */
const SYNC_LOG_CACHE_KEY = 'MONDAY_SYNC_LOG';
const SYNC_DONE_CACHE_KEY = 'MONDAY_SYNC_DONE';
const SYNC_CACHE_TTL_SEC = 21600; // 6h (CacheService max)

/*************************************************
 * MENU
 *************************************************/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Monday.com')
    .addItem('지금 동기화 (전체 새로고침)', 'openSyncDialogAndRun')
    .addItem('일일 자동 실행 설정 (최초 1회)', 'setupDailyTrigger')
    .addItem('보드 컬럼 ID 목록 보기 (진단용)', 'listBoardColumns')
    .addItem('항목 1개 전체 컬럼 덤프 (진단용)', 'dumpItemColumns')
    .addToUi();
}

/*************************************************
 * DIAGNOSTIC — dump board's actual column ids/titles/types
 *************************************************/
function listBoardColumns(){
  const apiKey = _getMondayApiKey_();
  const query = 'query ($boardId: [ID!]) { boards(ids: $boardId) { columns { id title type } } }';
  const resp = _mondayFetch_(apiKey, query, { boardId: BOARD_ID });
  const cols = (resp.data && resp.data.boards && resp.data.boards[0] && resp.data.boards[0].columns) || [];
  const rows = [['id', 'title', 'type']].concat(cols.map(function(c){ return [c.id, c.title, c.type]; }));
  _writeDiagSheet_('_diag_columns', rows);
  SpreadsheetApp.getUi().alert('완료! "_diag_columns" 탭을 확인하세요. (총 ' + cols.length + '개 컬럼)');
}

/*************************************************
 * DIAGNOSTIC — dump one item's ALL column_values (id/title/type/text/value)
 * Use this to spot duplicate-named columns (e.g. two "데이터 출처" columns).
 *************************************************/
function dumpItemColumns(){
  const ui = SpreadsheetApp.getUi();
  const resp1 = ui.prompt('진단할 Monday item_id 입력 (시트 A열 값)', ui.ButtonSet.OK_CANCEL);
  if (resp1.getSelectedButton() !== ui.Button.OK) return;
  const itemId = resp1.getResponseText().trim();
  if (!itemId) return;

  const apiKey = _getMondayApiKey_();
  const query = 'query ($itemIds: [ID!]) { items(ids: $itemIds) { id name column_values { id type text value column { title } } } }';
  const resp = _mondayFetch_(apiKey, query, { itemIds: [itemId] });
  const it = resp.data && resp.data.items && resp.data.items[0];
  if (!it) { ui.alert('Item not found: ' + itemId); return; }
  const rows = [['id', 'title', 'type', 'text', 'value']].concat(
    (it.column_values || []).map(function(cv){
      return [cv.id, cv.column && cv.column.title, cv.type, cv.text, cv.value];
    })
  );
  _writeDiagSheet_('_diag_item_' + it.id, rows);
  ui.alert('완료! "_diag_item_' + it.id + '" 탭을 확인하세요. (항목: ' + it.name + ')');
}

function _writeDiagSheet_(name, rows){
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  if (rows.length) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  ss.setActiveSheet(sh);
}

/*************************************************
 * MANUAL RUN — show progress dialog, then run the sync
 *************************************************/
function openSyncDialogAndRun(){
  _resetSyncLog_();
  const html = HtmlService.createHtmlOutput(_syncDialogHtml_())
    .setWidth(560)
    .setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, 'Monday.com 동기화');
  syncMondayToSheet();
}

/** Polled by the dialog's client-side JS. */
function getSyncLog(){
  const cache = CacheService.getScriptCache();
  const logRaw = cache.get(SYNC_LOG_CACHE_KEY);
  const doneRaw = cache.get(SYNC_DONE_CACHE_KEY);
  return {
    lines: logRaw ? JSON.parse(logRaw) : [],
    done: !!doneRaw,
    result: doneRaw ? JSON.parse(doneRaw) : null
  };
}

function _resetSyncLog_(){
  const cache = CacheService.getScriptCache();
  cache.remove(SYNC_LOG_CACHE_KEY);
  cache.remove(SYNC_DONE_CACHE_KEY);
}

function _logAppend_(msg){
  const cache = CacheService.getScriptCache();
  const ts = Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm:ss');
  const line = '[' + ts + '] ' + msg;
  const existing = cache.get(SYNC_LOG_CACHE_KEY);
  const arr = existing ? JSON.parse(existing) : [];
  arr.push(line);
  cache.put(SYNC_LOG_CACHE_KEY, JSON.stringify(arr), SYNC_CACHE_TTL_SEC);
  Logger.log(line);
}

function _markSyncDone_(result){
  const cache = CacheService.getScriptCache();
  cache.put(SYNC_DONE_CACHE_KEY, JSON.stringify(result), SYNC_CACHE_TTL_SEC);
}

function _syncDialogHtml_(){
  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    '  body{font-family:Arial,Helvetica,sans-serif;font-size:13px;margin:0;padding:16px;color:#202124;}' +
    '  #status{font-weight:bold;margin-bottom:8px;}' +
    '  #status.running{color:#1a73e8;}' +
    '  #status.ok{color:#188038;}' +
    '  #status.error{color:#d93025;}' +
    '  #log{background:#f1f3f4;border:1px solid #dadce0;border-radius:6px;padding:10px;' +
    '       height:280px;overflow-y:auto;white-space:pre-wrap;font-family:Consolas,Menlo,monospace;font-size:12px;}' +
    '  #closeBtn{margin-top:12px;padding:6px 16px;cursor:pointer;}' +
    '</style></head><body>' +
    '  <div id="status" class="running">⏳ 동기화 진행 중...</div>' +
    '  <div id="log"></div>' +
    '  <button id="closeBtn" onclick="google.script.host.close()">닫기</button>' +
    '  <script>' +
    '    var timer = null;' +
    '    function poll(){' +
    '      google.script.run.withSuccessHandler(render).withFailureHandler(onFail).getSyncLog();' +
    '    }' +
    '    function render(data){' +
    '      var logEl = document.getElementById("log");' +
    '      logEl.textContent = data.lines.join("\\n");' +
    '      logEl.scrollTop = logEl.scrollHeight;' +
    '      if (data.done){' +
    '        clearInterval(timer);' +
    '        var statusEl = document.getElementById("status");' +
    '        if (data.result && data.result.ok){' +
    '          statusEl.className = "ok";' +
    '          statusEl.textContent = "✅ 완료: " + data.result.total + "행 동기화됨 (" + data.result.elapsed + "초)";' +
    '        } else {' +
    '          statusEl.className = "error";' +
    '          statusEl.textContent = "❌ 오류: " + (data.result && data.result.message ? data.result.message : "알 수 없는 오류");' +
    '        }' +
    '      }' +
    '    }' +
    '    function onFail(err){' +
    '      clearInterval(timer);' +
    '      var statusEl = document.getElementById("status");' +
    '      statusEl.className = "error";' +
    '      statusEl.textContent = "❌ 로그 조회 실패: " + err.message;' +
    '    }' +
    '    poll();' +
    '    timer = setInterval(poll, 800);' +
    '  </script>' +
    '</body></html>';
}

/*************************************************
 * MAIN — full refresh: fetch ALL Monday items, replace sheet contents
 * (no UI calls here — this also runs unattended via the daily trigger)
 *************************************************/
function syncMondayToSheet(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Another sync is already running. Skipped.');
    _logAppend_('⚠️ 이미 다른 동기화가 실행 중입니다. 건너뜁니다.');
    _markSyncDone_({ ok: false, message: '다른 동기화가 이미 실행 중입니다.' });
    return;
  }
  const startTime = new Date();
  try {
    _logAppend_('동기화 시작...');
    const apiKey = _getMondayApiKey_();
    const sheet = _getSheet_();

    _logAppend_('Monday 보드에서 전체 항목 조회 중...');
    const colIds = COLUMN_MAP.filter(function(c){ return c.colId; }).map(function(c){ return c.colId; });
    const items = _fetchAllItems_(apiKey, colIds);
    _logAppend_('Monday에서 총 ' + items.length + '개 항목 조회 완료.');

    var formulaOverrides = {};
    if (FORMULA_COL_IDS.length){
      const missingIds = items
        .filter(function(it){ return _needsFormulaRetry_(it); })
        .map(function(it){ return it.id; });
      if (missingIds.length) {
        _logAppend_('formula 컬럼(인입사유/국가 등) 재조회 중... (' + missingIds.length + '건)');
        formulaOverrides = _fetchColumnsForItems_(apiKey, missingIds, FORMULA_COL_IDS);
      }
    }

    _logAppend_('시트에 기록할 행 생성 중...');
    const rows = items.map(function(it){ return _buildRow_(it, formulaOverrides[it.id]); });

    _logAppend_('기존 시트 데이터를 새 데이터로 교체 중... (Monday에서 삭제된 항목은 함께 제거됩니다)');
    _replaceSheetData_(sheet, rows);

    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log('Sync complete: ' + rows.length + ' row(s) written.');
    _logAppend_('✅ 완료: ' + rows.length + '행 기록됨. (' + elapsed + '초 소요)');
    _markSyncDone_({ ok: true, total: rows.length, elapsed: elapsed });
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    Logger.log('ERROR: ' + msg);
    _logAppend_('❌ 오류 발생: ' + msg);
    _markSyncDone_({ ok: false, message: msg });
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * DAILY TRIGGER SETUP (run once manually)
 *************************************************/
function setupDailyTrigger(){
  ['syncMondayToSheet', 'appendNewMondayItems'].forEach(function(handlerName){
    ScriptApp.getProjectTriggers().forEach(function(t){
      if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
    });
  });
  ScriptApp.newTrigger('syncMondayToSheet')
    .timeBased()
    .atHour(RUN_TRIGGER_HOUR)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();
  Logger.log('Daily trigger installed: ~' + RUN_TRIGGER_HOUR + ':00 Asia/Seoul');
}

/*************************************************
 * SHEET HELPERS
 *************************************************/
function _getSheet_(){
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
}

/** Overwrites the header and replaces all data rows below it in one shot. */
function _replaceSheetData_(sheet, rows){
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), HEADER.length);
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  if (rows.length) sheet.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
}

/*************************************************
 * MONDAY FETCH (paged, only requested columns)
 *************************************************/
function _cvFragments_(){
  return 'id type text value ' +
    '... on MirrorValue { display_value } ' +
    '... on FormulaValue { display_value } ' +
    '... on StatusValue { label } ' +
    '... on LinkValue { url text } ' +
    '... on NumbersValue { number } ' +
    '... on DateValue { date } ' +
    '... on BoardRelationValue { linked_items { name } }';
}

function _fetchAllItems_(apiKey, colIds){
  const fr = _cvFragments_();
  const query =
    'query ($boardId: [ID!], $pageCursor: String, $colIds: [String!]) {' +
    '  boards(ids: $boardId) {' +
    '    items_page(limit: ' + PAGE_LIMIT + ', cursor: $pageCursor) {' +
    '      items { id name column_values(ids: $colIds) { ' + fr + ' } } cursor } } }';
  var cursor = null, all = [];
  do {
    const resp = _mondayFetch_(apiKey, query, { boardId: BOARD_ID, pageCursor: cursor, colIds: colIds });
    const page = resp.data && resp.data.boards && resp.data.boards[0] && resp.data.boards[0].items_page;
    const batch = (page && page.items) ? page.items : [];
    all = all.concat(batch);
    cursor = page && page.cursor;
    if (cursor) Utilities.sleep(120);
  } while (cursor);
  return all;
}

function _needsFormulaRetry_(item){
  const map = {};
  (item.column_values || []).forEach(function(cv){ map[cv.id] = cv; });
  return FORMULA_COL_IDS.some(function(cid){ return _isEmpty_(map[cid]); });
}

function _fetchColumnsForItems_(apiKey, itemIds, colIds){
  const fr = _cvFragments_();
  const byItem = {};
  const CHUNK = 50;
  for (var i = 0; i < itemIds.length; i += CHUNK) {
    const slice = itemIds.slice(i, i + CHUNK);
    const query =
      'query ($itemIds: [ID!], $colIds: [String!]) {' +
      '  items(ids: $itemIds) { id column_values(ids: $colIds) { ' + fr + ' } } }';
    const resp = _mondayFetch_(apiKey, query, { itemIds: slice, colIds: colIds });
    const its = (resp.data && resp.data.items) || [];
    its.forEach(function(it){
      const map = byItem[it.id] || (byItem[it.id] = {});
      (it.column_values || []).forEach(function(cv){ map[cv.id] = cv; });
    });
    Utilities.sleep(120);
  }
  return byItem;
}

/*************************************************
 * ROW BUILDING / VALUE DISPLAY
 *************************************************/
function _buildRow_(item, formulaOverrides){
  const cvMap = {};
  (item.column_values || []).forEach(function(cv){ cvMap[cv.id] = cv; });
  const cells = COLUMN_MAP.map(function(c){
    if (c.colId === null) return item.name; // Order ID
    var cv = cvMap[c.colId];
    if (FORMULA_COL_IDS.indexOf(c.colId) >= 0 && _isEmpty_(cv) && formulaOverrides && formulaOverrides[c.colId]) {
      cv = formulaOverrides[c.colId];
    }
    const v = _displayValue_(cv);
    return (typeof c.transform === 'function') ? c.transform(v) : v;
  });
  return [String(item.id)].concat(cells);
}

function _displayValue_(cv){
  if (!cv) return '';
  if (_has_(cv.display_value)) return _clean_(cv.display_value);
  if (_has_(cv.label)) return _clean_(cv.label);
  if (_has_(cv.url)) return _clean_(cv.url);
  if (_has_(cv.number)) return String(cv.number);
  if (_has_(cv.date)) return String(cv.date);
  if (Array.isArray(cv.linked_items) && cv.linked_items.length) {
    const names = cv.linked_items.map(function(li){ return li && li.name; }).filter(Boolean);
    if (names.length) return names.join(', ');
  }
  if (_has_(cv.text)) return _clean_(cv.text);
  if (_has_(cv.value)) {
    const parsed = _tryParse_(cv.value);
    if (parsed && typeof parsed === 'object') {
      if (parsed.api_ticket_url) return _cleanZendeskUrl_(String(parsed.api_ticket_url));
      if (_has_(parsed.display_value)) return _clean_(parsed.display_value);
      if (_has_(parsed.text)) return _clean_(parsed.text);
      if (_has_(parsed.label)) return _clean_(parsed.label);
      if (_has_(parsed.url)) return _clean_(parsed.url);
      if (parsed.date) return String(parsed.date);
      if (parsed.number != null) return String(parsed.number);
      return _clean_(JSON.stringify(parsed));
    }
    return _clean_(String(cv.value).replace(/^"(.*)"$/, '$1'));
  }
  return '';
}

function _isEmpty_(cv){
  if (!cv) return true;
  return !(_has_(cv.display_value) || _has_(cv.label) || _has_(cv.url) ||
    _has_(cv.number) || _has_(cv.date) || _has_(cv.text) || _has_(cv.value));
}

function _cleanZendeskUrl_(u){
  try {
    const id = (u.match(/\/tickets\/(\d+)/) || [])[1];
    const m = u.match(/^https?:\/\/([^\/]+)/);
    if (id && m) return 'https://' + m[1] + '/tickets/' + id;
  } catch (e) {}
  return String(u || '').replace(/\.json$/i, '');
}

function _has_(v){ return v != null && String(v).trim() !== ''; }
function _clean_(s){ return String(s == null ? '' : s).trim(); }
function _tryParse_(s){ try { return typeof s === 'string' ? JSON.parse(s) : s; } catch (e) { return null; } }

/*************************************************
 * MONDAY API (GraphQL)
 *************************************************/
function _mondayFetch_(apiKey, query, variables){
  const resp = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: apiKey },
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Monday API HTTP ' + code + ': ' + resp.getContentText());
  const json = JSON.parse(resp.getContentText());
  if (json.errors && json.errors.length) throw new Error('Monday GraphQL error: ' + JSON.stringify(json.errors));
  return json;
}

function _getMondayApiKey_(){
  const key = PropertiesService.getScriptProperties().getProperty('MONDAY_API_KEY');
  if (!key) throw new Error('Missing Script Property MONDAY_API_KEY. Project Settings → Script Properties에서 추가하세요.');
  return key;
}