// =============================================
// Lazada/Shopee 리스팅 오류 → Monday.com 자동 연동
// 수정: 7월 7일 B열 삽입 반영 + openById 수정
// 추가 수정: 토큰을 Script Properties로 이동, 트리거를 평일 8:00-17:30 30분 간격으로 제한
// =============================================

var CFG = {
  MONDAY_TOKEN: PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN'),
  SPREADSHEET_ID: '1HZ14uqTVeP7bGYZDu9v9Ve2C1xNY_m6dcSv-KMCoAKc',
  BOARD_ID: '18409753446',
  GROUP_ID: 'group_mm3jz6yk',

  SHEETS: [
    { name: 'Lazada log', dataStart: 5 },
    { name: 'Shopee log', dataStart: 5 }
  ],

  // 7월 7일 B열 삽입 후 기준 (0-based 인덱스)
  COL: {
    ORDER_ID:      2,   // C열
    DATE:          3,   // D열
    PLATFORM:      4,   // E열
    PRODUCT_GROUP: 8,   // I열
    SKU:           9,   // J열
    DEVICE:        10,  // K열
    MODEL:         11,  // L열
    CATEGORY:      12,  // M열 ← 'Listing Issue' 필터
    ERRORS:        14   // O열
  },

  MON: {
    SKU:           'text_mm2nk668',
    PLATFORM:      'text_mm3rpny8',
    DEVICE:        'text_mm2n5yy2',
    MODEL:         'text_mm2nvwkb',
    ERRORS:        'long_text_mm3gzttm',
    ORDER_NUM:     'long_text_mm39c4yp',
    PRODUCT_GROUP: 'text_mm3rwqdv',
    DATE:          'date4',
    STATUS:        'status'
  },

  // 실행 허용 시간대 (Asia/Seoul 기준)
  SYNC_WINDOW: {
    WEEKDAYS_ONLY: true, // 토=6/일=7 제외
    START_HOUR: 8,       // 08:00 부터
    START_MINUTE: 0,
    END_HOUR: 17,        // 17:30 미만까지 (오후 5시 30분)
    END_MINUTE: 30
  }
};

// ── 현재 시각이 실행 허용 시간대인지 확인 ──
function isWithinSyncWindow_() {
  var now = new Date();
  var day = Number(Utilities.formatDate(now, 'Asia/Seoul', 'u')); // 1=Mon ... 7=Sun
  var hour = Number(Utilities.formatDate(now, 'Asia/Seoul', 'H'));
  var minute = Number(Utilities.formatDate(now, 'Asia/Seoul', 'm'));

  if (CFG.SYNC_WINDOW.WEEKDAYS_ONLY && (day === 6 || day === 7)) return false;

  var nowMinutes = hour * 60 + minute;
  var startMinutes = CFG.SYNC_WINDOW.START_HOUR * 60 + CFG.SYNC_WINDOW.START_MINUTE;
  var endMinutes = CFG.SYNC_WINDOW.END_HOUR * 60 + CFG.SYNC_WINDOW.END_MINUTE;

  if (nowMinutes < startMinutes || nowMinutes >= endMinutes) return false;
  return true;
}

// ── ▶ 메인 실행 함수 ──
function syncListingIssuesToMonday() {
  if (!isWithinSyncWindow_()) {
    Logger.log('⏭ 실행 시간대 아님 (평일 08:00-17:30 KST 외) — 스킵');
    return;
  }

  if (!CFG.MONDAY_TOKEN) {
    Logger.log('🚨 MONDAY_TOKEN Script Property가 설정되지 않았습니다.');
    return;
  }

  try {
    var ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    var props = PropertiesService.getScriptProperties();
    var processedKeys = JSON.parse(props.getProperty('processedKeys') || '{}');
    var newKeys = {};

    CFG.SHEETS.forEach(function(sheetCfg) {
      try {
        var sheet = ss.getSheetByName(sheetCfg.name);
        if (!sheet) {
          Logger.log('⚠️ 시트 없음: ' + sheetCfg.name);
          return;
        }

        var lastRow = sheet.getLastRow();
        if (lastRow < sheetCfg.dataStart) {
          Logger.log('⚠️ 데이터 없음: ' + sheetCfg.name);
          return;
        }

        var rowCount = lastRow - sheetCfg.dataStart + 1;
        var data = sheet.getRange(sheetCfg.dataStart, 1, rowCount, 15).getValues();
        var processed = 0;

        for (var i = 0; i < data.length; i++) {
          var row = data[i];

          // Listing Issue 필터
          var category = String(row[CFG.COL.CATEGORY] || '').trim();
          if (category !== 'Listing Issue') continue;

          var orderId = String(row[CFG.COL.ORDER_ID] || '').trim();
          var sku     = String(row[CFG.COL.SKU]      || '').trim();
          if (!orderId || !sku) {
            Logger.log('⚠️ 행 ' + (sheetCfg.dataStart + i) + ': OrderID 또는 SKU 없음');
            continue;
          }

          var key = orderId + '_' + sku;
          if (processedKeys[key] || newKeys[key]) {
            Logger.log('⏭ 중복 스킵: ' + key);
            continue;
          }

          try {
            var platform = String(row[CFG.COL.PLATFORM]      || '').trim();
            var pg       = String(row[CFG.COL.PRODUCT_GROUP] || '').trim();
            var device   = String(row[CFG.COL.DEVICE]        || '').trim();
            var model    = String(row[CFG.COL.MODEL]         || '').trim();
            var errors   = String(row[CFG.COL.ERRORS]        || '').trim();

            var dateVal = row[CFG.COL.DATE];
            var dateStr = '';
            if (dateVal instanceof Date && !isNaN(dateVal)) {
              dateStr = Utilities.formatDate(dateVal, 'Asia/Seoul', 'yyyy-MM-dd');
            } else if (dateVal) {
              dateStr = String(dateVal).substring(0, 10);
            }

            var itemName = '[' + platform + '] ' + sku;

            var cols = {};
            cols[CFG.MON.SKU]           = sku;
            cols[CFG.MON.PLATFORM]      = platform;
            cols[CFG.MON.DEVICE]        = device;
            cols[CFG.MON.MODEL]         = model;
            cols[CFG.MON.ERRORS]        = { text: errors };
            cols[CFG.MON.ORDER_NUM]     = { text: orderId };
            cols[CFG.MON.PRODUCT_GROUP] = pg;
            if (dateStr) cols[CFG.MON.DATE] = { date: dateStr };
            cols[CFG.MON.STATUS]        = { label: '오류 접수' };

            createMondayItem(itemName, cols);

            newKeys[key] = true;
            processed++;
            Logger.log('✅ ' + itemName + ' | 키: ' + key);
            Utilities.sleep(300);

          } catch(rowErr) {
            Logger.log('❌ 행 ' + (sheetCfg.dataStart + i) + ' 처리 실패: ' + rowErr.message);
          }
        }

        Logger.log('[' + sheetCfg.name + '] 처리: ' + processed + '건');

      } catch(sheetErr) {
        Logger.log('❌ 시트 오류 [' + sheetCfg.name + ']: ' + sheetErr.message);
      }
    });

    // 새 키 저장
    Object.keys(newKeys).forEach(function(k) { processedKeys[k] = true; });
    props.setProperty('processedKeys', JSON.stringify(processedKeys));
    Logger.log('=== 완료 | 누적 키 수: ' + Object.keys(processedKeys).length + ' ===');

  } catch(e) {
    Logger.log('🚨 치명적 오류: ' + e.message + '\n스택: ' + e.stack);
  }
}

// ── 처리 키 초기화 (필요시만 실행) ──
function resetProcessedKeys() {
  PropertiesService.getScriptProperties().deleteProperty('processedKeys');
  Logger.log('✅ 초기화 완료');
}

// ── Monday.com 아이템 생성 ──
function createMondayItem(name, cols) {
  var query = 'mutation { create_item('
    + 'board_id: ' + CFG.BOARD_ID + ', '
    + 'group_id: "' + CFG.GROUP_ID + '", '
    + 'item_name: ' + JSON.stringify(name) + ', '
    + 'column_values: ' + JSON.stringify(JSON.stringify(cols))
    + ') { id } }';

  var response = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + CFG.MONDAY_TOKEN,
      'Content-Type': 'application/json',
      'API-Version': '2024-01'
    },
    payload: JSON.stringify({ query: query }),
    muteHttpExceptions: true
  });

  var statusCode = response.getResponseCode();
  var body = response.getContentText();

  if (statusCode !== 200) {
    throw new Error('HTTP ' + statusCode + ': ' + body);
  }

  var json = JSON.parse(body);
  if (json.errors) {
    throw new Error('Monday API 오류: ' + JSON.stringify(json.errors));
  }

  return json;
}

// ── 트리거 재설정: 기존 syncListingIssuesToMonday 트리거를 모두 지우고
//     30분 간격 트리거 1개로 재생성 (실제 평일 8:00-17:30 제한은 isWithinSyncWindow_ 가드로 처리) ──
function setupSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'syncListingIssuesToMonday') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncListingIssuesToMonday')
    .timeBased()
    .everyMinutes(30)
    .create();

  Logger.log('✅ 트리거 재설정 완료: 30분 간격 (함수 내부에서 평일 08:00-17:30 KST만 실제 실행)');
}