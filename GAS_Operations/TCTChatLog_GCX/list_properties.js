// ── 진단용: Script Properties에 저장된 모든 키를 정확한 길이와 함께 출력 ──
function listAllScriptProperties() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var keys = Object.keys(props);

  Logger.log('=== 총 ' + keys.length + '개 속성 ===');
  keys.forEach(function(k) {
    var v = props[k];
    Logger.log('KEY: ["' + k + '"]  (key length=' + k.length + ')  value length=' + v.length + '  value 앞 20자: "' + v.substring(0, 20) + '"');
  });

  Logger.log('=== 직접 getProperty("MONDAY_TOKEN") 테스트 ===');
  var direct = PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN');
  Logger.log('결과: ' + (direct === null ? 'null (못 찾음)' : '찾음, length=' + direct.length));
}