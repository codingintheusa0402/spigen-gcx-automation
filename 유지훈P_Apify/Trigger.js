// ── Spreadsheet ID helper (SPREADSHEET_ID defined in Alert.js) ───────────────
function getSpreadsheetId_() {
  return SPREADSHEET_ID;
}

// ── Write current KST time to the "Timestamp" column in the US sheet ─────────
function _writeTimestampToUsSheet_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('US');
    if (!sheet) { Logger.log('_writeTimestampToUsSheet_: US sheet not found'); return; }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colIdx = headers.indexOf('Timestamp'); // 0-based
    if (colIdx === -1) { Logger.log('_writeTimestampToUsSheet_: Timestamp header not found'); return; }

    const tz = (typeof CONFIG !== 'undefined' && CONFIG.timezone) ? CONFIG.timezone : 'Asia/Seoul';
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(2, colIdx + 1).setValue(now);
    Logger.log('_writeTimestampToUsSheet_: wrote "%s" to US col %s (row 2)', now, colIdx + 1);
  } catch (e) {
    Logger.log('_writeTimestampToUsSheet_ failed: ' + (e && e.message ? e.message : e));
  }
}

// ── Daily entry point: starts both scrapers ───────────────────────────────────
function dailyScrapeJob() {
  // Review scraper
  try {
    startApifyRunAndSchedulePoll();
    _ensureRecurringPoller_();
    Logger.log('dailyScrapeJob: review scraper started');
  } catch (e) {
    Logger.log('dailyScrapeJob: review scraper start failed — ' + (e && e.message ? e.message : e));
  }

  // Product scraper
  try {
    runProductNowAndPollRecurring();
    Logger.log('dailyScrapeJob: product scraper started');
  } catch (e) {
    Logger.log('dailyScrapeJob: product scraper start failed — ' + (e && e.message ? e.message : e));
  }
}

// ── Run once in the GAS editor to install a daily 09:00 KST trigger ──────────
function createDailyTriggers() {
  // Remove any existing dailyScrapeJob triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyScrapeJob') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('dailyScrapeJob')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();

  Logger.log('Daily trigger created: dailyScrapeJob at 09:00 KST every day');
}
