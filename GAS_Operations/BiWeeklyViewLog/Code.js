// Bi-Weekly report view tracker. Share the web-app URL (optionally ?r=<report code>) instead of the deck link.
// Each open logs a row in the "Visits" sheet; the page sends a heartbeat every 30s while visible,
// which updates "마지막 활동" and "체류 시간(분)". Runs as the deployer, so viewers need no sheet access;
// viewers are identified by their Spigen Google account (web app access = domain).
const LOG_SHEET_ID = '1NSMiMwz_4nd6NlOv8rvYeyGvBDmCRZefsWD0Ux0PfDg';
const REPORTS = {   // report code -> deck id (add a line every period)
  '261002': '1quCr9Xj-pSsVXKrYuaEOq0LPILZMN2LPUwBkY1f_GFI'
};
const DEFAULT_REPORT = '261002';
const HEARTBEAT_SEC = 30;

function doGet(e) {
  const code = (e && e.parameter && e.parameter.r) || DEFAULT_REPORT;
  const deck = REPORTS[code] || REPORTS[DEFAULT_REPORT];
  const t = HtmlService.createTemplateFromFile('Page');
  t.code = code; t.deck = deck; t.beat = HEARTBEAT_SEC;
  return t.evaluate().setTitle(code + ' GCX Bi-weekly Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function _sheet() { return SpreadsheetApp.openById(LOG_SHEET_ID).getSheetByName('Visits'); }

function _name(email) {
  try {
    const r = People.People.searchDirectoryPeople({ query: email, readMask: 'names', sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'] });
    const p = (r.people || [])[0];
    if (p && p.names && p.names[0]) return p.names[0].displayName;
  } catch (err) {}
  return email.split('@')[0];
}

// Called once when the page loads. Returns the visit id used by heartbeats.
function startVisit(code, ua) {
  const email = Session.getActiveUser().getEmail() || '(unknown)';
  const id = Utilities.getUuid(), now = new Date();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { _sheet().appendRow([id, _name(email), email, now, now, 0, code, String(ua || '').slice(0, 150)]); }
  finally { lock.releaseLock(); }
  return id;
}

// Heartbeat while the page is visible: last activity = now, duration = minutes since start.
function beat(id) {
  const sh = _sheet();
  const f = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  if (!f) return;
  const r = f.getRow(), start = sh.getRange(r, 4).getValue(), now = new Date();
  sh.getRange(r, 5, 1, 2).setValues([[now, Math.round((now - start) / 600) / 100]]);
}

// Run once from the editor so the deployer authorizes Sheets + directory access.
function authorize() { _sheet(); _name(Session.getEffectiveUser().getEmail()); Logger.log('ok'); }
