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
    const tz = (typeof CONFIG !== 'undefined' && CONFIG.timezone) ? CONFIG.timezone : 'Asia/Seoul';
    const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    // Find Timestamp col by header (case-insensitive, trimmed); fall back to col J (10)
    let col = 10;
    if (lastCol >= 1) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const idx = headers.findIndex(h => String(h).trim().toLowerCase() === 'timestamp');
      if (idx !== -1) col = idx + 1;
    }

    sheet.getRange(2, col).setValue(now);
    Logger.log('_writeTimestampToUsSheet_: wrote "%s" to US col %s (row 2)', now, col);
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

  // Product/rating scraper — DISABLED 2026-10-13: replaced by the unified
  // daily ASIN rating-fetch job (actor 2AuZCyLxKZuO2OyNY). The old paid
  // axesso_data task (09gwMIgOKhyyDrH3g) was costing money for data no
  // longer needed from here. Trigger left installed on purpose; this just
  // no-ops the product-scraper half of the daily job.
  Logger.log('dailyScrapeJob: product scraper skipped (decommissioned 2026-10-13, superseded by rating_fetch_daily.py)');
}

// ── Run once in the GAS editor to install a daily 04:00 KST trigger ──────────
function createDailyTriggers() {
  // Remove any existing dailyScrapeJob triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyScrapeJob') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('dailyScrapeJob')
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .inTimezone('Asia/Seoul')
    .create();

  Logger.log('Daily trigger created: dailyScrapeJob at 04:00 KST every day');
}
