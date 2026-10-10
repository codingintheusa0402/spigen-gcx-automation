/**
 * GCX Server Monitor — watches the GCX server from Google's cloud, so a dead/off server is still noticed.
 * The server writes a heartbeat (epoch ms) to the "GCX Server Heartbeat" sheet every 2 min.
 * check() runs every 5 min: no heartbeat for > down_after_min → 🔴 DOWN card (once);
 * heartbeat back without a reboot → 🟢 reachable again (the server itself reports boots/resume).
 * Webhook lives in the sheet's config tab (private file, not in git).
 * One-time: run setup() from the editor and authorize.
 */
var SHEET_ID = '19Nta2cWyUc8tp4TR2X-ripJYNEPP4_BPzrUu0aE7rIY';

function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('check').timeBased().everyMinutes(5).create();
  check();
  Logger.log('GCX Server Monitor armed — check() every 5 min');
}

function cfg_() {
  var rows = SpreadsheetApp.openById(SHEET_ID).getSheetByName('config').getDataRange().getValues(), c = {};
  rows.forEach(function (r) { c[r[0]] = r[1]; });
  return c;
}

function fmt_(ms) { return Utilities.formatDate(new Date(ms), 'Asia/Seoul', 'MM-dd HH:mm'); }
function dur_(ms) { var m = Math.round(ms / 60000); return m < 60 ? m + ' min' : Math.floor(m / 60) + ' h ' + (m % 60) + ' min'; }

function post_(webhook, icon, title, text) {
  var card = { cardsV2: [{ cardId: 'gcx-mon-' + Date.now(), card: {
    header: { title: icon + ' ' + title, subtitle: 'GCX Server · cloud monitor · ' + fmt_(Date.now()) },
    sections: [{ widgets: [{ textParagraph: { text: text } }] }] } }] };
  UrlFetchApp.fetch(webhook, { method: 'post', contentType: 'application/json; charset=UTF-8',
                               payload: JSON.stringify(card), muteHttpExceptions: true });
}

function check() {
  var c = cfg_(), props = PropertiesService.getScriptProperties();
  var row = SpreadsheetApp.openById(SHEET_ID).getSheetByName('heartbeat').getRange('A2:D2').getValues()[0];
  var hb = Number(row[0]) || 0, boot = Number(row[1]) || 0, note = row[3] || '';
  var now = Date.now(), downAfter = (Number(c.down_after_min) || 7) * 60000;
  var down = props.getProperty('down') === '1', downSince = Number(props.getProperty('downSince')) || 0;

  if (!down && hb && now - hb > downAfter) {
    post_(c.webhook, '🔴', 'GCX Server DOWN',
      'No heartbeat since <b>' + fmt_(hb) + '</b> (' + dur_(now - hb) + ' ago).<br>' +
      'Possible causes: power off / battery empty, Windows crashed or frozen, restart stuck, or internet lost.<br>' +
      'It restarts by itself if possible — you\'ll get 🔵 booting and 🟢 back-up messages. If not within ~15 min: check the laptop power and network.');
    props.setProperties({ down: '1', downSince: String(hb) });
  } else if (down && hb && now - hb <= downAfter) {
    if (!(boot > downSince)) {                        // no reboot in between → it was a network/heartbeat gap
      post_(c.webhook, '🟢', 'GCX Server reachable again',
        'Heartbeat is back after <b>' + dur_(hb - downSince) + '</b> of silence — the server did not restart (likely a network gap). ' + note);
    }                                                 // after a reboot the server posts its own 🟢 back-up message
    props.setProperties({ down: '0', downSince: '0' });
  }
}
