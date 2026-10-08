// ── 디버깅 전용: Shopee log 17789행이 왜 Monday로 안 들어갔는지 원인 확인 ──
// Monday_sync.gs와 같은 프로젝트에 이 파일 하나 추가하고 debugRow17789() 실행
function debugRow17789() {
  var ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Lazada log');
  var targetRow = 17789;

  Logger.log('=== 기본 정보 ===');
  Logger.log('lastRow: ' + sheet.getLastRow());
  Logger.log('dataStart: 5, targetRow: ' + targetRow);
  if (targetRow > sheet.getLastRow()) {
    Logger.log('🚨 17789행이 getLastRow() 범위를 벗어남 — 스크립트가 아예 이 행을 읽지 않고 있습니다.');
    return;
  }
  if (targetRow < 5) {
    Logger.log('🚨 17789행이 dataStart(5)보다 위쪽입니다.');
    return;
  }

  var rowValues = sheet.getRange(targetRow, 1, 1, 15).getValues()[0];

  var category = String(rowValues[CFG.COL.CATEGORY] || '');
  var orderId  = String(rowValues[CFG.COL.ORDER_ID] || '');
  var sku      = String(rowValues[CFG.COL.SKU] || '');
  var platform = String(rowValues[CFG.COL.PLATFORM] || '');

  Logger.log('=== 원본 값 (trim 전) ===');
  Logger.log('CATEGORY(M열) raw: "' + category + '" (length=' + category.length + ')');
  Logger.log('ORDER_ID(C열) raw: "' + orderId + '"');
  Logger.log('SKU(J열) raw: "' + sku + '"');
  Logger.log('PLATFORM(E열) raw: "' + platform + '"');

  var categoryTrim = category.trim();
  var orderIdTrim = orderId.trim();
  var skuTrim = sku.trim();

  Logger.log('=== 필터 체크 ===');
  Logger.log('category.trim() === "Listing Issue" ? ' + (categoryTrim === 'Listing Issue'));
  if (categoryTrim !== 'Listing Issue') {
    Logger.log('  → 실제 값 charCode 목록: ' + categoryTrim.split('').map(function(c){ return c.charCodeAt(0); }).join(','));
  }
  Logger.log('orderId 존재? ' + !!orderIdTrim);
  Logger.log('sku 존재? ' + !!skuTrim);

  if (categoryTrim !== 'Listing Issue' || !orderIdTrim || !skuTrim) {
    Logger.log('🚨 여기서 필터링되어 Monday로 안 갑니다 (카테고리 불일치 또는 필수값 누락).');
    return;
  }

  var key = orderIdTrim + '_' + skuTrim;
  var props = PropertiesService.getScriptProperties();
  var processedKeys = JSON.parse(props.getProperty('processedKeys') || '{}');

  Logger.log('=== 중복 체크 ===');
  Logger.log('생성된 key: ' + key);
  Logger.log('이미 processedKeys에 있음? ' + !!processedKeys[key]);
  if (processedKeys[key]) {
    Logger.log('🚨 이미 처리된 것으로 기록되어 스킵됩니다. (동일 OrderID+SKU 조합이 예전에 이미 처리됨 — 다른 행이었을 수도 있음)');
    return;
  }

  Logger.log('=== 토큰/시간대 체크 ===');
  Logger.log('MONDAY_TOKEN 설정됨? ' + !!CFG.MONDAY_TOKEN);
  Logger.log('지금 isWithinSyncWindow_()? ' + isWithinSyncWindow_());

  Logger.log('✅ 여기까지 통과 — 이 행은 조건상 정상적으로 Monday로 보내져야 합니다. createMondayItem 호출 시 API 에러가 났을 가능성을 보려면 Executions 로그에서 17789행 관련 "❌ 행 17789 처리 실패" 메시지를 찾아보세요.');
}