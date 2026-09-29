/**********************************************************
 * MENU
 **********************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Apify Rating')
    .addItem('Run Now (refresh ratings)', 'runApifyRatingRefreshNow')
    .addItem('Cancel Polling', 'cancelRatingPolling')
    .addSeparator()
    .addItem('Install Daily Weekday 8AM Trigger', 'setupDailyWeekdayTrigger')
    .addItem('Remove Daily Trigger', 'removeDailyWeekdayTrigger')
    .addSeparator()
    .addItem('Sync New ASINs Now', 'syncNewAsinsToApifyTask')
    .addItem('Install Weekly ASIN-Sync Trigger', 'setupAsinSyncTrigger')
    .addItem('Remove Weekly ASIN-Sync Trigger', 'removeAsinSyncTrigger')
    .addToUi();
}


/**********************************************************
 * ENTRY POINTS
 **********************************************************/

// The function the daily time-based trigger calls — skips Sat/Sun, otherwise
// kicks off the same refresh as the manual "Run Now" menu item.
function dailyWeekdayKickoff() {
  const dow = Number(Utilities.formatDate(new Date(), CONFIG.timezone, 'u')); // 1=Mon ... 7=Sun
  if (dow === 6 || dow === 7) {
    Logger.log('Skipping weekend run (ISO day=%s).', dow);
    return;
  }
  runApifyRatingRefreshNow();
}

// Manual "Run Now" from the menu, and also called by dailyWeekdayKickoff() on weekdays.
function runApifyRatingRefreshNow() {
  const props = PropertiesService.getScriptProperties();
  const pendingRunId = props.getProperty('RATING_LAST_RUN_ID');

  if (pendingRunId) {
    Logger.log('A rating run is already pending (runId=%s). Ensuring poller is scheduled.', pendingRunId);
  } else {
    _startRatingRun_();
  }
  _scheduleRecurringRatingPoll_();
}

function cancelRatingPolling() {
  _deleteTriggersByHandler_('pollRatingRunAndWrite');
  _cleanupRatingState_();
}


/**********************************************************
 * START RUN
 **********************************************************/
function _startRatingRun_() {
  const ss = SpreadsheetApp.getActive();
  const token = _getToken();

  // Pull in any ASINs added to the sheets since the last sync so they're
  // scraped in this same run, not only after the weekly sync trigger fires.
  // A sync failure must not block the refresh of already-covered ASINs.
  try {
    syncNewAsinsToApifyTask();
  } catch (e) {
    Logger.log('ASIN sync before run failed (continuing with existing task URLs): ' + e);
  }

  const url = `https://api.apify.com/v2/actor-tasks/${encodeURIComponent(TASK_ID)}/runs?token=${encodeURIComponent(token)}`;
  Logger.log('Starting rating run (async): ' + url.replace(/token=[^&]+/, 'token=***'));

  const resp = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  if (code >= 400) {
    throw new Error(`Failed to start run: HTTP ${code}: ${body.slice(0, 1000)}`);
  }

  const data = JSON.parse(body).data;
  if (!data || !data.id) throw new Error('Start run response missing run id.');

  _rememberRatingRun_(data.id, data.defaultDatasetId || null);

  const msg = `Apify rating run started. runId=${data.id}`;
  Logger.log(msg);
  try { ss.toast(msg, 'Apify Rating', 5); } catch (e) { /* no UI context (trigger run) */ }
}


/**********************************************************
 * POLLER (invoked by recurring time-based trigger)
 **********************************************************/
function pollRatingRunAndWrite() {
  const ss = SpreadsheetApp.getActive();
  const token = _getToken();
  const props = PropertiesService.getScriptProperties();

  const runId = props.getProperty('RATING_LAST_RUN_ID');
  const datasetIdFromStart = props.getProperty('RATING_LAST_DATASET_ID');
  const startedAtMs = Number(props.getProperty('RATING_LAST_POLL_STARTED_AT_MS')) || Date.now();

  if (!runId) {
    _deleteTriggersByHandler_('pollRatingRunAndWrite');
    return;
  }

  const elapsedMin = (Date.now() - startedAtMs) / 60000;
  if (elapsedMin > CONFIG.pollMaxMinutes) {
    _cleanupRatingState_();
    _deleteTriggersByHandler_('pollRatingRunAndWrite');
    Logger.log('Rating polling timed out after %s min.', CONFIG.pollMaxMinutes);
    return;
  }

  const runUrl = `https://api.apify.com/v2/actor-runs/${encodeURIComponent(runId)}?token=${encodeURIComponent(token)}`;
  const runResp = UrlFetchApp.fetch(runUrl, { method: 'get', muteHttpExceptions: true });
  if (runResp.getResponseCode() >= 400) return;

  const runData = JSON.parse(runResp.getContentText()).data;
  const status = runData.status;
  const datasetId = runData.defaultDatasetId || datasetIdFromStart;

  Logger.log('Rating run status=%s, datasetId=%s', status, datasetId);

  if (status === 'SUCCEEDED') {
    if (!datasetId) {
      _cleanupRatingState_();
      _deleteTriggersByHandler_('pollRatingRunAndWrite');
      return;
    }

    const ratingByAsin = _fetchRatingDatasetItems_(datasetId, token);
    let written = 0;
    for (const sheetCfg of SHEETS) {
      written += _writeRatingsToSheet_(sheetCfg, ratingByAsin);
    }

    Logger.log('Rating refresh done: %s dataset item(s), %s row(s) written across %s sheet(s).',
      ratingByAsin.size, written, SHEETS.length);
    try {
      ss.toast(`Rating refresh done: ${written} row(s) updated.`, 'Apify Rating', 8);
    } catch (e) { /* no UI context */ }

    _cleanupRatingState_();
    _deleteTriggersByHandler_('pollRatingRunAndWrite');
    return;
  }

  if (['FAILED', 'ABORTED', 'TIMED-OUT'].includes(status)) {
    _cleanupRatingState_();
    _deleteTriggersByHandler_('pollRatingRunAndWrite');
    Logger.log('Rating run ended with status=%s', status);
  }
}


/**********************************************************
 * DATASET FETCH
 **********************************************************/
function _fetchRatingDatasetItems_(datasetId, token) {
  const baseUrl = `https://api.apify.com/v2/datasets/${datasetId}/items?clean=true&format=json&token=${token}`;
  let offset = 0;
  const limit = 1000;
  let allItems = [];

  while (true) {
    const url = `${baseUrl}&offset=${offset}&limit=${limit}`;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) {
      throw new Error('Dataset fetch failed: ' + resp.getContentText());
    }

    const items = JSON.parse(resp.getContentText());
    if (!items || items.length === 0) break;

    allItems = allItems.concat(items.map(i => ({
      asin: i.asin || '',
      productRating: _extractRatingValue_(i.productRating)
    })));

    offset += items.length;
    if (items.length < limit) break;
  }

  // Same ASIN can appear more than once (the task's URL list has duplicate
  // ASIN entries mirroring duplicate rows in the sheets) — last write wins.
  const byAsin = new Map();
  for (const item of allItems) {
    if (item.asin) byAsin.set(item.asin, item.productRating);
  }
  return byAsin;
}

// Rating text comes back locale-formatted (e.g. "4,5 von 5 Sternen" for amazon.de).
// Just take the leading number and normalize the decimal separator.
function _extractRatingValue_(raw) {
  if (!raw) return '';
  const s = String(raw).trim().replace(',', '.');
  const m = s.match(/^\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : '';
}


/**********************************************************
 * WRITE TO SHEET — match by ASIN → Rating col, hyperlinked to the
 * product's amazon.de page. A row whose ASIN got no (or empty) result in
 * this run keeps whatever was already in its cell — only a fresh
 * non-empty scrape result overwrites a rating.
 **********************************************************/
function _writeRatingsToSheet_(sheetCfg, ratingByAsin) {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const sheet = ss.getSheetByName(sheetCfg.name);
  if (!sheet) throw new Error(`Sheet "${sheetCfg.name}" not found.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < sheetCfg.dataStartRow) return 0;

  const numRows = lastRow - sheetCfg.dataStartRow + 1;
  const asinValues = sheet.getRange(sheetCfg.dataStartRow, sheetCfg.asinCol, numRows, 1).getValues();

  const ratingRange = sheet.getRange(sheetCfg.dataStartRow, sheetCfg.ratingCol, numRows, 1);
  const existingFormulas = ratingRange.getFormulas();
  const existingValues = ratingRange.getValues();

  let updated = 0;
  const output = asinValues.map((row, i) => {
    const asin = String(row[0] || '').trim();
    if (asin && ratingByAsin.has(asin)) {
      const rating = ratingByAsin.get(asin);
      if (rating !== '') {
        updated++;
        return [`=HYPERLINK("https://www.amazon.de/dp/${asin}",${rating})`];
      }
    }
    return [existingFormulas[i][0] || existingValues[i][0]];
  });

  ratingRange.setFormulas(output);
  return updated;
}

// Manual recovery: reapply an already-scraped dataset without starting a new
// Apify run (e.g. after fixing the write logic). Run from the editor with the
// dataset ID from the Apify console or a prior execution log line.
function reprocessDataset(datasetId) {
  const token = _getToken();
  const ratingByAsin = _fetchRatingDatasetItems_(datasetId, token);
  let written = 0;
  for (const sheetCfg of SHEETS) {
    written += _writeRatingsToSheet_(sheetCfg, ratingByAsin);
  }
  Logger.log('Reprocessed dataset %s: %s row(s) updated across %s sheet(s).', datasetId, written, SHEETS.length);
}


/**********************************************************
 * DAILY (WEEKDAY-ONLY) TRIGGER
 **********************************************************/
function setupDailyWeekdayTrigger() {
  removeDailyWeekdayTrigger();
  ScriptApp.newTrigger('dailyWeekdayKickoff')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.dailyHour)
    .nearMinute(0)
    .inTimezone(CONFIG.timezone)
    .create();
  Logger.log('Daily trigger installed: every day ~%s:00 %s (dailyWeekdayKickoff skips Sat/Sun).',
    CONFIG.dailyHour, CONFIG.timezone);
}

function removeDailyWeekdayTrigger() {
  _deleteTriggersByHandler_('dailyWeekdayKickoff');
}


/**********************************************************
 * WEEKLY ASIN SYNC — scans every configured sheet for ASINs not yet
 * covered by the Apify task's URL list and appends them. Runs from a
 * persistent time-based trigger, so it fires even with no session open.
 **********************************************************/
// Amazon ASINs are always exactly 10 uppercase-alphanumeric characters.
// The ASIN column also holds non-ASIN placeholder text on some rows
// (e.g. "TBU", "미판매" = not-yet-updated / discontinued) — never scrape those.
function _looksLikeAsin_(value) {
  return /^[A-Z0-9]{10}$/.test(value);
}

function syncNewAsinsToApifyTask() {
  const token = _getToken();

  const getUrl = `https://api.apify.com/v2/actor-tasks/${encodeURIComponent(TASK_ID)}?token=${encodeURIComponent(token)}`;
  const getResp = UrlFetchApp.fetch(getUrl, { muteHttpExceptions: true });
  if (getResp.getResponseCode() >= 400) {
    throw new Error(`Failed to read task: HTTP ${getResp.getResponseCode()}: ${getResp.getContentText().slice(0, 500)}`);
  }
  const task = JSON.parse(getResp.getContentText()).data;
  const currentUrls = (task.input && task.input.urls) || [];
  const coveredAsins = new Set(currentUrls.map(u => u.split('/dp/')[1]).filter(Boolean));

  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  const newUrls = [];
  for (const sheetCfg of SHEETS) {
    const sheet = ss.getSheetByName(sheetCfg.name);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < sheetCfg.dataStartRow) continue;
    const numRows = lastRow - sheetCfg.dataStartRow + 1;
    const asinValues = sheet.getRange(sheetCfg.dataStartRow, sheetCfg.asinCol, numRows, 1).getValues();
    for (const row of asinValues) {
      const asin = String(row[0] || '').trim();
      if (asin && _looksLikeAsin_(asin) && !coveredAsins.has(asin)) {
        newUrls.push(`https://www.amazon.de/dp/${asin}`);
        coveredAsins.add(asin); // dedup within this same sync pass too
      }
    }
  }

  if (newUrls.length === 0) {
    Logger.log('ASIN sync: no new ASINs found across %s sheet(s).', SHEETS.length);
    return;
  }

  const putUrl = `https://api.apify.com/v2/actor-tasks/${encodeURIComponent(TASK_ID)}?token=${encodeURIComponent(token)}`;
  const putResp = UrlFetchApp.fetch(putUrl, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify({ input: { urls: currentUrls.concat(newUrls) } }),
    muteHttpExceptions: true
  });
  if (putResp.getResponseCode() >= 400) {
    throw new Error(`Failed to update task: HTTP ${putResp.getResponseCode()}: ${putResp.getContentText().slice(0, 500)}`);
  }

  Logger.log('ASIN sync: added %s new ASIN(s) to task %s (total urls now %s).',
    newUrls.length, TASK_ID, currentUrls.length + newUrls.length);
}

function setupAsinSyncTrigger() {
  removeAsinSyncTrigger();
  ScriptApp.newTrigger('syncNewAsinsToApifyTask')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay[CONFIG.asinSyncWeekDay])
    .atHour(CONFIG.asinSyncHour)
    .nearMinute(0)
    .inTimezone(CONFIG.timezone)
    .create();
  Logger.log('ASIN-sync trigger installed: every %s ~%s:00 %s.',
    CONFIG.asinSyncWeekDay, CONFIG.asinSyncHour, CONFIG.timezone);
}

function removeAsinSyncTrigger() {
  _deleteTriggersByHandler_('syncNewAsinsToApifyTask');
}


/**********************************************************
 * STATE / TRIGGER HELPERS
 **********************************************************/
function _rememberRatingRun_(runId, datasetId) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('RATING_LAST_RUN_ID', runId);
  if (datasetId) props.setProperty('RATING_LAST_DATASET_ID', datasetId);
  props.setProperty('RATING_LAST_POLL_STARTED_AT_MS', String(Date.now()));
}

function _cleanupRatingState_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('RATING_LAST_RUN_ID');
  props.deleteProperty('RATING_LAST_DATASET_ID');
  props.deleteProperty('RATING_LAST_POLL_STARTED_AT_MS');
}

function _scheduleRecurringRatingPoll_() {
  _deleteTriggersByHandler_('pollRatingRunAndWrite');
  const every = Math.max(1, Number(CONFIG.pollIntervalMinutes || 1));
  ScriptApp.newTrigger('pollRatingRunAndWrite')
    .timeBased()
    .everyMinutes(every)
    .create();
}

function _deleteTriggersByHandler_(handlerName) {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  }
}
