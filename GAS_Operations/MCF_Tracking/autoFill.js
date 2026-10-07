/***** ========= BACKFILL CONFIG ========= *****/
// Adjust these if your sheet layout changes.
var BF_SHEET_NAME  = 'MCF 발송 로그';
var BF_START_ROW   = 4;    // first data row (skip headers)
var BF_COL_REGION  = 2;    // B — "JP" triggers FE-first, anything else EU-first
var BF_COL_ORDER   = 17;   // Q — sellerFulfillmentOrderId
var BF_COL_SENT    = 16;   // P — MCF sent date (yyyy-mm-dd)
var BF_COL_RESULT  = 26;   // Z — NOT a tracking number cache (confirmed 2026-07-31: it's a live
                            //     ARRAYFORMULA computing "Product Price − Shipping Fee", spilled
                            //     from Z3; verified across all 3,755 data rows — 0 held a formula
                            //     at the row level, 0 looked like a tracking number, 2,613 (70%)
                            //     were plain prices). This comment used to claim otherwise — that
                            //     was wrong and caused the col R corruption incident. Do not read
                            //     from or write to this column expecting tracking data.
var BF_COL_FEE     = 25;   // Y — Transportation Fee (€, ¥, £) written by backfillMCFFees()

// Cell-note marker for col Y values written by fillPreviewEstimatesForRange() — an ESTIMATE
// (getFulfillmentPreview quote), not the actual settled fee. backfillMCFFeesRecent() checks for
// this prefix so it knows to overwrite an estimate once real settlement data arrives, instead of
// treating any non-blank cell as permanently done.
var FEE_ESTIMATE_NOTE_PREFIX = 'ESTIMATE (getFulfillmentPreview)';
function _isFeeEstimateNote(note) {
  return String(note || '').indexOf(FEE_ESTIMATE_NOTE_PREFIX) === 0;
}

/**
 * DISABLED 2026-07-31 — do not re-enable without picking a different target column.
 *
 * This function assumed BF_COL_RESULT (Z) was an available static cache for tracking numbers.
 * Confirmed false: Z is a live ARRAYFORMULA (anchored at Z3) computing
 * "Product Price − Transportation Fee" across the entire column — there is no free column here
 * to write tracking numbers into. Every "already has a value" row this function skipped was
 * actually skipping a live profit-margin number, and every row where Z looked resolvable via
 * freezeTrackingColumnR() got a price frozen into col R as a fake tracking number — that's the
 * corruption incident from earlier today. Left in place (not deleted) for reference; throws
 * immediately so a manual run from the editor can't repeat the damage.
 */
function backfillTrackingNumbers() {
  throw new Error('backfillTrackingNumbers() is disabled: col Z (BF_COL_RESULT) is a live profit-margin ' +
    'ARRAYFORMULA, not a tracking-number cache. See the function doc comment before re-enabling.');
  // eslint-disable-next-line no-unreachable
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  _warmLwaTokens(); // fetch EU+JP tokens once before the loop to avoid per-row rate-limit 401s

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows  = lastRow - BF_START_ROW + 1;
  var orderIds = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();
  var regions  = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var existing = sheet.getRange(BF_START_ROW, BF_COL_RESULT, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var current = existing[i][0];
    // Already has a valid tracking number — never overwrite.
    if (current && !_isErrorValue(current)) continue;

    var isJP      = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    var endpoints = isJP ? ['FE', 'EU'] : ['EU', 'FE'];

    var tn = '';
    var retries429b = 0;
    var fetchOkB = false;
    while (!fetchOkB) {
      try {
        var tracks = _tracksWithFallbacks(orderId, endpoints);
        tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
          ? tracks[0].trackingNumber.trim() : '';
        Utilities.sleep(400);
        fetchOkB = true;
      } catch (e) {
        if (_isRateLimit429(e) && retries429b < 3) {
          retries429b++;
          Logger.log('Row ' + (BF_START_ROW + i) + ': 429 — waiting 45 s (retry ' + retries429b + '/3)');
          Utilities.sleep(45000);
        } else {
          var errMsg = (isJP ? 'JP' : 'EU') + ' ERR: ' + (e.message || e);
          sheet.getRange(BF_START_ROW + i, BF_COL_RESULT).setValue(errMsg);
          Logger.log('Row ' + (BF_START_ROW + i) + ': ' + errMsg);
          fetchOkB = true; // stop retrying, move on
        }
      }
    }

    if (tn) {
      sheet.getRange(BF_START_ROW + i, BF_COL_RESULT).setValue(tn);
      Logger.log('Row ' + (BF_START_ROW + i) + ': wrote ' + tn);
    }
  }
}

/***** ========= RETRY 429 ERRORS IN COL R ========= *****/
var RETRY_R_START_ROW        = 1108; // only rows from here down — matches the reported error range
var RETRY_R_COL              = 18;   // R — live =AMZTK()/=AMZTK_JP() tracking formula
var RETRY_R_MAX_ROWS_PER_RUN = 40;   // hard cap so one execution can't hammer SP-API across hundreds of rows
var RETRY_R_MAX_CONSEC_429   = 5;    // abort the run early if quota is clearly still exhausted

/**
 * Finds cells in col R (MCF 발송 로그, row RETRY_R_START_ROW+) whose live =AMZTK()/=AMZTK_JP()
 * formula is currently showing a 429 QuotaExceeded error, and re-fetches those specific orders
 * directly.
 *
 * - Found a tracking number: freezes the cell to a static =HYPERLINK(...) (same URL format the
 *   live formula produces). This is deliberate, not just a recalc trick — confirmed live that
 *   Sheets' custom-function engine does NOT reliably re-run AMZTK()/AMZTK_JP() just because
 *   Apps Script rewrites the identical formula text via the API (cache was primed correctly but
 *   the cell kept showing the stale 429 text), so writing the final known-correct value directly
 *   is the only way to guarantee the sheet actually displays it — and it also means the row won't
 *   match the 429 filter on the next run, so later runs progress to new rows instead of
 *   re-fetching the same already-resolved orders every hour.
 * - No tracking number yet (order genuinely not ready, no error): primes the cache AMZTK reads
 *   from, then forces a real recalculation by writing a placeholder formula and then the original
 *   formula back — two distinct writes, since Sheets only re-triggers custom functions on an
 *   actual content change — so the stale error clears to blank instead of sitting there forever.
 *
 * Bounded two ways so a single run can never call the API forever:
 *   1) stops after RETRY_R_MAX_ROWS_PER_RUN rows
 *   2) aborts the whole run after RETRY_R_MAX_CONSEC_429 consecutive rows still come back 429 —
 *      quota is clearly still exhausted, so this run stops instead of hammering it; the next
 *      hourly trigger picks up where this one left off
 *
 * Run manually, or via the hourly trigger already installed on the live script
 * (Apps Script editor → Triggers).
 */
function retryR429Errors() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < RETRY_R_START_ROW) return;

  var numRows   = lastRow - RETRY_R_START_ROW + 1;
  var rValues   = sheet.getRange(RETRY_R_START_ROW, RETRY_R_COL, numRows, 1).getDisplayValues();
  var rFormulas = sheet.getRange(RETRY_R_START_ROW, RETRY_R_COL, numRows, 1).getFormulas();
  var regions   = sheet.getRange(RETRY_R_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var orderIds  = sheet.getRange(RETRY_R_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();

  _warmLwaTokens();

  var cache = CacheService.getScriptCache();
  var processed = 0, fixed = 0, cleared = 0, stillFailing = 0, consec429 = 0;

  for (var i = 0; i < numRows; i++) {
    if (processed >= RETRY_R_MAX_ROWS_PER_RUN) {
      Logger.log('retryR429Errors: hit per-run cap (' + RETRY_R_MAX_ROWS_PER_RUN + ' rows) — remaining rows retried next hour.');
      break;
    }

    if (!_is429ErrorValue(rValues[i][0])) continue;

    var formula = rFormulas[i][0];
    if (!formula) continue; // not a live formula cell — nothing to retry

    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var row       = RETRY_R_START_ROW + i;
    var isJP      = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    var endpoints = isJP ? ['FE', 'EU'] : ['EU', 'FE'];
    var cacheKey  = (isJP ? 'AMZTK_JP_' : 'AMZTK_') + orderId;

    processed++;

    var tn = '', got429 = false;
    for (var attempt = 0; attempt < 3; attempt++) {
      try {
        var tracks = _tracksWithFallbacks(orderId, endpoints);
        tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
          ? tracks[0].trackingNumber.trim() : '';
        got429 = false;
        break;
      } catch (e) {
        if (_isRateLimit429(e) && attempt < 2) {
          got429 = true;
          Logger.log('Row ' + row + ': 429 — waiting 45 s (attempt ' + (attempt + 1) + '/3)');
          Utilities.sleep(45000);
          continue;
        }
        got429 = _isRateLimit429(e);
        break; // non-429, or 429 retries exhausted — stop retrying this row
      }
    }

    if (got429) {
      consec429++;
      stillFailing++;
      Logger.log('Row ' + row + ': still 429 after retries (' + consec429 + '/' + RETRY_R_MAX_CONSEC_429 + ' consecutive)');
      if (consec429 >= RETRY_R_MAX_CONSEC_429) {
        Logger.log('retryR429Errors: quota still exhausted after ' + consec429 + ' consecutive 429s — stopping this run.');
        break;
      }
      continue;
    }
    consec429 = 0;

    if (tn) {
      // Freeze to a static HYPERLINK — guaranteed to render, and drops out of the 429 filter
      // on future runs so this row is never re-fetched again.
      var domain = isJP ? 'jp' : 'de';
      var url    = 'https://www.swiship.' + domain + '/track?id=' + tn;
      sheet.getRange(row, RETRY_R_COL).setFormula('=HYPERLINK("' + url + '","' + tn + '")');
      fixed++;
      Logger.log('Row ' + row + ': fixed → ' + tn);
    } else {
      // Order not ready yet — prime the cache AMZTK reads from, then force a genuine
      // recalculation (placeholder write, then the real formula) so the stale 429 text clears.
      cache.put(cacheKey, '', 600);
      var cell = sheet.getRange(row, RETRY_R_COL);
      cell.setFormula('=NA()');
      SpreadsheetApp.flush();
      cell.setFormula(formula);
      cleared++;
      Logger.log('Row ' + row + ': no tracking number yet — cleared stale error');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('retryR429Errors done — processed: ' + processed + ', fixed: ' + fixed + ', cleared: ' + cleared + ', still 429: ' + stillFailing);

  // Rides this same hourly trigger — see resolveBlankTrackingNumbers() doc comment below for why
  // this is a SEPARATE failure mode retryR429Errors above can't see. Best-effort: a failure here
  // must never mask the retry summary already logged above.
  try {
    resolveBlankTrackingNumbers();
  } catch (e) {
    Logger.log('resolveBlankTrackingNumbers (called from retryR429Errors) failed: ' + (e.message || e));
  }
}

function _is429ErrorValue(v) {
  var s = String(v || '');
  return s.indexOf('SP-API error 429') >= 0 || s.indexOf('QuotaExceeded') >= 0;
}

/***** ========= RESOLVE BLANK TRACKING CELLS (col R) — companion to retryR429Errors ========= ****
 * Root cause (confirmed live 2026-08-28, direct SP-API testing outside GAS): self-created MCF
 * fulfillment orders stop being resolvable via GetFulfillmentOrder roughly ~1 day after creation
 * (tested: today's orders 200 OK, yesterday's and older all 400 "Unable to get order info" —
 * INCLUDING orders the sheet already shows a real, correct tracking number for). AMZTK()/
 * AMZTK_JP() treat that 400 as "not shipped yet" (see _isNoOrderInfoError in sp-api.js) and
 * silently return '' — no error text ever lands in the cell, so retryR429Errors()'s
 * _is429ErrorValue() text-match filter can never see or retry these rows. Per the user
 * (2026-08-28): a real tracking number normally only appears after AT LEAST ~24h, so any row
 * whose live formula happened to (re)run before that window, or after the order aged out of
 * GetFulfillmentOrder, is stuck blank FOREVER — no amount of future recalculation can ever fix
 * it, since the same order id will 400 the exact same way every time from then on.
 *
 * This function, called from the end of retryR429Errors() (same hourly trigger, no new trigger
 * needed), does two independent passes over every row with a live AMZTK()/AMZTK_JP() formula in
 * col R (identified by the formula text containing "AMZTK(" — distinguishes a still-live formula
 * from an already-frozen static =HYPERLINK(...) cell, which contains no such call):
 *
 *   1) LOCK-IN: a cell already DISPLAYING a real tracking number is frozen to a static
 *      =HYPERLINK(...) right away. Needed because AMZTK()'s own success cache is only 6h
 *      (sp-api.js) — once that expires, the very next recalculation re-queries the same
 *      now-aged-out order and silently blanks a previously-correct cell. Freezing removes the
 *      live formula so this can never happen to an already-good value again.
 *   2) BACKFILL: a genuinely blank cell (skipping rows younger than BLANK_R_MIN_AGE_HOURS, to
 *      give the normal ~24h+ window a chance) is retried directly. Resolved → frozen the same
 *      way as (1). Still unresolved AND the row is older than BLANK_R_GIVE_UP_DAYS → the cell is
 *      frozen to a literal blank (formula removed) so it stops silently re-querying SP-API (and
 *      burning quota) on every future recalculation for an order that will never resolve. Still
 *      unresolved but not yet past that age → left alone, tried again next hour.
 *
 * Scans the WHOLE sheet from BF_START_ROW (not just RETRY_R_START_ROW+, unlike retryR429Errors) —
 * this failure mode isn't confined to any particular row range. Only pass (2) makes API calls, and
 * only up to BLANK_R_MAX_ROWS_PER_RUN of them, with the same consecutive-429 abort as
 * retryR429Errors — pass (1) is pure sheet reads/writes, no quota cost, so it always runs in full.
 *
 * BLANK_R_GIVE_UP_DAYS is a judgment call, not a measured constant — 3 days gives real stragglers
 * a wide margin past the observed ~24h cutoff before permanently giving up on a row. Adjust if
 * real data suggests otherwise.
 */
var BLANK_R_MAX_ROWS_PER_RUN = 60;   // cap on actual SP-API calls per run (pass 2 only)
var BLANK_R_MAX_CONSEC_429   = 5;    // abort early if quota is clearly still exhausted
var BLANK_R_MIN_AGE_HOURS    = 24;   // don't even attempt a row shipped more recently than this
var BLANK_R_GIVE_UP_DAYS     = 3;    // past this age, stop retrying and freeze the cell blank

/***** "GStore MCF 발송 로그" — Google Store MCF replacements, same live-formula + freeze scheme as
 * "MCF 발송 로그" col R, just different columns. Its P cells hold
 *   =IF(OR(B5="",O5=""),"",IF(B5<>"JP",HYPERLINK("https://www.swiship.de/track?id="&AMZTK(O5),AMZTK(O5)),
 *                                      HYPERLINK("https://www.swiship.jp/track?id="&AMZTK_JP(O5),AMZTK_JP(O5))))
 * (written by installGStoreTrackingFormulas()), and resolveBlankTrackingNumbers() freezes them.
 */
var GSTORE_SHEET_NAME = 'GStore MCF 발송 로그';
var GSTORE_START_ROW  = 5;    // rows 1-4 are notes + header
var GSTORE_COL_REGION = 2;    // B — MCF 발송 국가
var GSTORE_COL_SENT   = 14;   // N — 발송일자
var GSTORE_COL_ORDER  = 15;   // O — MCF Order ID
var GSTORE_COL_TRACK  = 16;   // P — Tracking Number

// Sheets resolveBlankTrackingNumbers() covers. retry429: "MCF 발송 로그" leaves 429-text cells to
// retryR429Errors(); GStore has no such companion, so its 429/ERR cells are retried here directly.
var TRACKING_SHEETS = [
  // GStore first: it's tiny, so it always gets its lookups before the main sheet can use up the shared budget.
  { name: GSTORE_SHEET_NAME, startRow: GSTORE_START_ROW, colRegion: GSTORE_COL_REGION, colSent: GSTORE_COL_SENT,
    colOrder: GSTORE_COL_ORDER, colTrack: GSTORE_COL_TRACK, retry429: true },
  { name: BF_SHEET_NAME, startRow: BF_START_ROW, colRegion: BF_COL_REGION, colSent: BF_COL_SENT,
    colOrder: BF_COL_ORDER, colTrack: RETRY_R_COL, retry429: false }
];

function resolveBlankTrackingNumbers() {
  var budget = { apiCalls: 0 }; // BLANK_R_MAX_ROWS_PER_RUN is shared across all sheets
  _warmLwaTokens();
  try {
    _topUpGStoreTrackingFormulas_();
  } catch (e) {
    Logger.log('_topUpGStoreTrackingFormulas_ failed: ' + (e.message || e));
  }
  for (var s = 0; s < TRACKING_SHEETS.length; s++) {
    try {
      _resolveBlankTrackingForSheet_(TRACKING_SHEETS[s], budget);
    } catch (e) {
      Logger.log('resolveBlankTrackingNumbers [' + TRACKING_SHEETS[s].name + '] failed: ' + (e.message || e));
    }
  }
}

function _resolveBlankTrackingForSheet_(cfg, budget) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) throw new Error('Sheet not found: ' + cfg.name);

  var lastRow = sheet.getLastRow();
  if (lastRow < cfg.startRow) return;

  var numRows   = lastRow - cfg.startRow + 1;
  var rValues   = sheet.getRange(cfg.startRow, cfg.colTrack,  numRows, 1).getDisplayValues();
  var rFormulas = sheet.getRange(cfg.startRow, cfg.colTrack,  numRows, 1).getFormulas();
  var regions   = sheet.getRange(cfg.startRow, cfg.colRegion, numRows, 1).getValues();
  var orderIds  = sheet.getRange(cfg.startRow, cfg.colOrder,  numRows, 1).getValues();
  var sentDates = sheet.getRange(cfg.startRow, cfg.colSent,   numRows, 1).getValues();

  var now = new Date();
  var frozenGood = 0, fixed = 0, gaveUp = 0, consec429 = 0, calls = 0;

  for (var i = 0; i < numRows; i++) {
    var formula = rFormulas[i][0];
    // No formula at all, or already frozen to a static value (no AMZTK() call left in it) —
    // nothing for this function to do either way.
    if (!formula || formula.indexOf('AMZTK') < 0) continue;

    var row     = cfg.startRow + i;
    var display = String(rValues[i][0] || '').trim();
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var isJP      = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    var endpoints = isJP ? ['FE', 'EU'] : ['EU', 'FE'];
    var domain    = isJP ? 'jp' : 'de';
    var isErr     = _is429ErrorValue(display) || _isErrorValue(display);

    // --- Pass 1: already resolved, just needs freezing before its cache can expire ---
    if (display && !isErr) {
      var url = 'https://www.swiship.' + domain + '/track?id=' + display;
      sheet.getRange(row, cfg.colTrack).setFormula('=HYPERLINK("' + url + '","' + display + '")');
      frozenGood++;
      continue;
    }

    // 429-error-text cells on "MCF 발송 로그" are retryR429Errors()'s job — don't double-handle them.
    if (display && _is429ErrorValue(display) && !cfg.retry429) continue;

    // --- Pass 2: blank (or error) — only attempt if old enough, and within this run's API budget ---
    var ageHours = _ageHours_(sentDates[i][0], now);
    if (ageHours === null || ageHours < BLANK_R_MIN_AGE_HOURS) continue;
    if (budget.apiCalls >= BLANK_R_MAX_ROWS_PER_RUN) continue;

    budget.apiCalls++;
    calls++;
    var tn = '', got429 = false, gotNoInfo = false;
    try {
      var tracks = _tracksWithFallbacks(orderId, endpoints);
      tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
        ? tracks[0].trackingNumber.trim() : '';
    } catch (e) {
      if (_isRateLimit429(e))        { got429 = true; }
      else if (_isNoOrderInfoError(e)) { gotNoInfo = true; }
      else { Logger.log('resolveBlankTrackingNumbers [' + cfg.name + '] row ' + row + ': unexpected error — ' + (e.message || e)); }
    }

    if (got429) {
      consec429++;
      if (consec429 >= BLANK_R_MAX_CONSEC_429) {
        Logger.log('resolveBlankTrackingNumbers [' + cfg.name + ']: quota exhausted after ' + consec429 + ' consecutive 429s — stopping this run.');
        break;
      }
      continue;
    }
    consec429 = 0;

    if (tn) {
      var url2 = 'https://www.swiship.' + domain + '/track?id=' + tn;
      sheet.getRange(row, cfg.colTrack).setFormula('=HYPERLINK("' + url2 + '","' + tn + '")');
      fixed++;
      Logger.log('resolveBlankTrackingNumbers [' + cfg.name + '] row ' + row + ': resolved → ' + tn);
      continue;
    }

    if (gotNoInfo && ageHours >= BLANK_R_GIVE_UP_DAYS * 24) {
      sheet.getRange(row, cfg.colTrack).setValue('');
      gaveUp++;
      Logger.log('resolveBlankTrackingNumbers [' + cfg.name + '] row ' + row + ': gave up (order info permanently unavailable, ' +
        Math.round(ageHours / 24) + 'd old) — froze blank to stop future retries');
    }
    // else: genuinely still not ready, or a transient error — leave the live formula for next run
  }

  SpreadsheetApp.flush();
  Logger.log('resolveBlankTrackingNumbers [' + cfg.name + '] done — locked-in good: ' + frozenGood +
    ', newly resolved: ' + fixed + ', gave up: ' + gaveUp + ', SP-API calls made: ' + calls);
}

// Sent-date cell → hours since then. Accepts a Date, or a date serial number (cell not date-formatted).
function _ageHours_(v, now) {
  if (v instanceof Date) return (now - v) / 3600000;
  if (typeof v === 'number' && v > 0) return (now - new Date(Date.UTC(1899, 11, 30) + v * 86400000)) / 3600000;
  return null;
}

/**
 * Writes the live AMZTK tracking formula into "GStore MCF 발송 로그" col P, empty cells only — never
 * overwrites a frozen tracking number or anything typed by hand.
 *   installGStoreTrackingFormulas()  — manual, rows GSTORE_START_ROW..GSTORE_FORMULA_LAST_ROW
 *   _topUpGStoreTrackingFormulas_()  — hourly, the GSTORE_TOPUP_ROWS rows BELOW the last data row, so
 *     new rows always have the formula waiting. Stays below the data on purpose: a row the hourly job
 *     gave up on is frozen blank, and must not get a fresh formula (and fresh retries) every hour.
 */
var GSTORE_FORMULA_LAST_ROW = 1000;
var GSTORE_TOPUP_ROWS       = 200;

function installGStoreTrackingFormulas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GSTORE_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + GSTORE_SHEET_NAME);
  var written = _writeGStoreTrackingFormulas_(sheet, GSTORE_START_ROW, GSTORE_FORMULA_LAST_ROW);
  Logger.log('installGStoreTrackingFormulas: wrote ' + written + ' formula(s) into ' + GSTORE_SHEET_NAME + ' col P');
}

function _topUpGStoreTrackingFormulas_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GSTORE_SHEET_NAME);
  if (!sheet) return;
  // Last row with an MCF Order ID (getLastRow() would count the pre-filled formula rows).
  var n   = Math.max(sheet.getLastRow() - GSTORE_START_ROW + 1, 1);
  var ids = sheet.getRange(GSTORE_START_ROW, GSTORE_COL_ORDER, n, 1).getValues();
  var lastData = GSTORE_START_ROW - 1;
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]).trim() !== '') { lastData = GSTORE_START_ROW + i; break; }
  }
  var from = lastData + 1, to = lastData + GSTORE_TOPUP_ROWS;
  if (to > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), to - sheet.getMaxRows());
  var written = _writeGStoreTrackingFormulas_(sheet, from, to);
  if (written) Logger.log('_topUpGStoreTrackingFormulas_: wrote ' + written + ' formula(s), rows ' + from + '-' + to);
}

function _writeGStoreTrackingFormulas_(sheet, fromRow, toRow) {
  var n = toRow - fromRow + 1;
  if (n <= 0) return 0;
  var cells = sheet.getRange(fromRow, GSTORE_COL_TRACK, n, 1);
  var vals = cells.getValues(), fs = cells.getFormulas();
  var written = 0, runStart = -1, run = [];
  function flushRun() {
    if (run.length) sheet.getRange(fromRow + runStart, GSTORE_COL_TRACK, run.length, 1).setFormulas(run);
    written += run.length; run = []; runStart = -1;
  }
  for (var i = 0; i <= n; i++) {
    var empty = i < n && !fs[i][0] && String(vals[i][0]).trim() === '';
    if (!empty) { flushRun(); continue; }
    if (runStart < 0) runStart = i;
    var r = fromRow + i;
    run.push(['=IF(OR(B' + r + '="",O' + r + '=""),"",IF(B' + r + '<>"JP",' +
      'HYPERLINK("https://www.swiship.de/track?id="&AMZTK(O' + r + '),AMZTK(O' + r + ')),' +
      'HYPERLINK("https://www.swiship.jp/track?id="&AMZTK_JP(O' + r + '),AMZTK_JP(O' + r + '))))']);
  }
  return written;
}

/***** ========= RETRY $0 TRANSPORTATION FEES IN COL Y ========= *****/
var RETRY_ZERO_FEE_MAX_ROWS_PER_RUN = 40; // hard cap so one run only clears/reprocesses a bounded batch

/**
 * Finds cells in col Y (Transportation Fee) showing a literal 0 — the symptom of the GCX
 * fee-type filtering bug (see _sumMcfFeeFromShipments) that was fixed in sp-api.js, but had
 * already written wrong 0s into the sheet before the fix landed. backfillMCFFees() skips any
 * row whose col Y already has a non-empty value — including "0" — so those rows never get
 * reprocessed on their own even after the fix.
 *
 * Clears up to RETRY_ZERO_FEE_MAX_ROWS_PER_RUN of those 0-cells back to blank (making them
 * "pending" under backfillMCFFees()'s own skip logic), then runs backfillMCFFees() so they
 * get recomputed with the fixed fee-type logic. All of backfillMCFFees()'s own safeguards
 * (429 sleep-and-retry per window, moving on rather than hanging) apply as-is — no separate
 * retry/backoff logic needed here.
 *
 * Bounded so a single run can't try to clear and reprocess the whole historical backlog at
 * once, which would risk a long-running execution across many months of Finances API windows.
 * Run hourly and it works through the backlog gradually.
 *
 * Run manually, or set up an hourly time-based trigger for this function via
 * Apps Script editor → Triggers.
 */
function retryZeroTransportationFees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows = lastRow - BF_START_ROW + 1;
  var feeVals = sheet.getRange(BF_START_ROW, BF_COL_FEE, numRows, 1).getValues();

  var cleared = 0;
  for (var i = 0; i < numRows && cleared < RETRY_ZERO_FEE_MAX_ROWS_PER_RUN; i++) {
    var v = feeVals[i][0];
    if (v === 0 || v === '0') {
      sheet.getRange(BF_START_ROW + i, BF_COL_FEE).setValue('');
      cleared++;
    }
  }

  if (!cleared) {
    Logger.log('retryZeroTransportationFees: no zero-fee cells found — nothing to do.');
    return;
  }

  Logger.log('retryZeroTransportationFees: cleared ' + cleared + ' zero-fee cell(s), running backfillMCFFees()...');
  backfillMCFFees();
}

/**
 * ONE-TIME recovery variant of retryZeroTransportationFees() — clears ALL zero-fee cells in
 * col Y (uncapped) in a single batched write, then runs backfillMCFFees() once.
 *
 * Use this instead of the hourly-safe retryZeroTransportationFees() (40 rows/run) when there's
 * a large one-off backlog: backfillMCFFees()'s cost is dominated by the date-range window
 * fetches (a handful of calls covering the whole range), not by how many rows get matched
 * against the result — clearing only a small slice per run means paying that same window-fetch
 * cost repeatedly while only a few rows benefit each time. Clearing everything at once lets a
 * single backfillMCFFees() run resolve the whole backlog in one pass (assuming SP-API quota
 * allows it to complete).
 *
 * Run manually from the Apps Script editor once SP-API's EU quota has recovered — check with
 * auditYColumnFees() first if unsure, since that's read-only and won't consume quota needed for
 * this to succeed.
 */
function recoverAllZeroTransportationFees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows  = lastRow - BF_START_ROW + 1;
  var feeRange = sheet.getRange(BF_START_ROW, BF_COL_FEE, numRows, 1);
  var feeVals  = feeRange.getValues();

  var cleared = 0;
  for (var i = 0; i < numRows; i++) {
    var v = feeVals[i][0];
    if (v === 0 || v === '0') {
      feeVals[i][0] = '';
      cleared++;
    }
  }

  if (!cleared) {
    Logger.log('recoverAllZeroTransportationFees: no zero-fee cells found — nothing to do.');
    return;
  }

  feeRange.setValues(feeVals); // single batched write, not one setValue() call per row
  Logger.log('recoverAllZeroTransportationFees: cleared ' + cleared + ' zero-fee cell(s) in one batch, running backfillMCFFees()...');
  backfillMCFFees();
}

/**
 * READ-ONLY audit: computes the true Finances API fee for every row in col Y that's currently
 * blank or a literal 0, WITHOUT writing anything back to col Y. Writes a side-by-side comparison
 * (row, orderId, sentDate, current Y value, computed true fee) to a separate "Y_Fee_Audit" sheet
 * tab instead, so this can be reviewed before touching the original data at all.
 *
 * Uses the same batched 60-day-window Finances API approach as backfillMCFFees() (a handful of
 * calls covering the whole date range, not one call per row) to minimize additional load while
 * SP-API's quota may still be recovering from earlier heavy use today. Handles 429 with the same
 * sleep-and-retry-once-per-window pattern; a window that still fails is logged and skipped rather
 * than blocking the rest of the audit.
 *
 * Run manually from the Apps Script editor.
 */
function auditYColumnFees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows   = lastRow - BF_START_ROW + 1;
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();
  var sentDates = sheet.getRange(BF_START_ROW, BF_COL_SENT,   numRows, 1).getValues();
  var regions   = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var currentY  = sheet.getRange(BF_START_ROW, BF_COL_FEE,    numRows, 1).getValues();

  var targets = [];
  var hasJP   = false;
  var minDate = null;

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var yVal = currentY[i][0];
    var isZeroOrBlank = (yVal === 0 || yVal === '0' || yVal === '' || yVal === null || yVal === undefined);
    if (!isZeroOrBlank) continue;

    var sentDate = String(sentDates[i][0] || '').trim();
    var isJP     = String(regions[i][0]   || '').trim().toUpperCase() === 'JP';
    if (isJP) hasJP = true;

    if (sentDate) {
      var d = new Date(sentDate);
      if (!minDate || d < minDate) minDate = d;
    }

    targets.push({ row: BF_START_ROW + i, orderId: orderId, sentDate: sentDate, isJP: isJP, currentY: yVal });
  }

  if (!targets.length) {
    Logger.log('auditYColumnFees: no zero/blank rows found — nothing to audit.');
    return;
  }

  var now = new Date(Date.now() - 5 * 60 * 1000);
  if (!minDate) minDate = new Date(now.getTime() - 180 * 24 * 3600 * 1000);

  Logger.log('auditYColumnFees: %s target rows, window %s → now', targets.length, minDate.toISOString().slice(0, 10));

  function fetchFeeMap(ep) {
    var feeMap      = {};
    var windowStart = new Date(minDate);

    while (windowStart < now) {
      var windowEnd = new Date(windowStart);
      windowEnd.setDate(windowEnd.getDate() + 60);
      if (windowEnd > now) windowEnd = now;
      if (windowStart >= windowEnd) break;

      try {
        var chunk = _buildFeeMapForWindow(ep, windowStart, windowEnd);
        Object.keys(chunk).forEach(function(k) { feeMap[k] = chunk[k]; });
        Logger.log('  [%s] %s – %s: %s orders found', ep,
          windowStart.toISOString().slice(0, 10), windowEnd.toISOString().slice(0, 10), Object.keys(chunk).length);
      } catch (e) {
        if (_isRateLimit429(e)) {
          Logger.log('  [%s] 429 — sleeping 15 s, retrying window', ep);
          Utilities.sleep(15000);
          try {
            var chunk2 = _buildFeeMapForWindow(ep, windowStart, windowEnd);
            Object.keys(chunk2).forEach(function(k) { feeMap[k] = chunk2[k]; });
          } catch (e2) {
            Logger.log('  [%s] retry also failed — skipping window: %s', ep, e2.message);
          }
        } else {
          Logger.log('  [%s] window error (%s): %s', ep, windowStart.toISOString().slice(0, 10), e.message);
        }
      }

      windowStart = new Date(windowEnd);
      windowStart.setDate(windowStart.getDate() + 1);
      Utilities.sleep(500);
    }

    return feeMap;
  }

  var euMap = fetchFeeMap('EU');
  var feMap = hasJP ? fetchFeeMap('FE') : {};

  var results = [['Row', 'OrderID', 'SentDate', 'Current Y value', 'Computed true fee', 'Source', 'Note']];
  var matched = 0, stillUnresolved = 0;

  targets.forEach(function(t) {
    var primaryMap   = t.isJP ? feMap : euMap;
    var secondaryMap = t.isJP ? euMap : feMap;
    var primaryEp    = t.isJP ? 'FE' : 'EU';
    var secondaryEp  = t.isJP ? 'EU' : 'FE';

    var fee, src;
    if (primaryMap[t.orderId] !== undefined)        { fee = primaryMap[t.orderId];   src = primaryEp; }
    else if (secondaryMap[t.orderId] !== undefined) { fee = secondaryMap[t.orderId]; src = secondaryEp; }
    else                                             { fee = ''; src = ''; }

    if (fee !== '') matched++; else stillUnresolved++;

    results.push([
      t.row, t.orderId, t.sentDate, t.currentY,
      fee, src,
      fee !== '' ? '' : 'not found in window — may be outside range, unsettled, or needs displayableOrderId fallback'
    ]);
  });

  var auditSheet = ss.getSheetByName('Y_Fee_Audit');
  if (!auditSheet) auditSheet = ss.insertSheet('Y_Fee_Audit');
  auditSheet.clearContents();
  auditSheet.getRange(1, 1, results.length, results[0].length).setValues(results);

  Logger.log('auditYColumnFees done — %s matched, %s still unresolved. See "Y_Fee_Audit" sheet. Col Y was NOT modified.',
    matched, stillUnresolved);
}

/**
 * Writes MCF fulfillment fees as static values into BF_COL_FEE (col Y).
 *
 * Batch approach: fetches all financial events in 60-day windows (EU + FE endpoints)
 * then matches every pending order ID against the result map — far fewer API calls
 * than the previous per-order approach.
 *
 * - Skips rows where fee is already filled (never overwrites a value).
 * - Rows marked RETRY or ERR are retried on the next run.
 * Run manually or set a daily time-based trigger.
 */
function backfillMCFFees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows   = lastRow - BF_START_ROW + 1;
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();
  var sentDates = sheet.getRange(BF_START_ROW, BF_COL_SENT,   numRows, 1).getValues();
  var regions   = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var existing  = sheet.getRange(BF_START_ROW, BF_COL_FEE,    numRows, 1).getValues();

  // Collect rows that still need a fee
  var pending = [];
  var hasJP   = false;
  var minDate = null;

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var curStr = String(existing[i][0] === null || existing[i][0] === undefined ? '' : existing[i][0]).trim();
    // Skip rows that already have a valid numeric fee
    if (curStr !== '' && curStr !== 'RETRY' && !_isErrorValue(curStr)) continue;

    var sentDate = String(sentDates[i][0] || '').trim();
    var isJP     = String(regions[i][0]   || '').trim().toUpperCase() === 'JP';
    if (isJP) hasJP = true;

    if (sentDate) {
      var d = new Date(sentDate);
      if (!minDate || d < minDate) minDate = d;
    }

    pending.push({ i: i, orderId: orderId, sentDate: sentDate, isJP: isJP });
  }

  if (!pending.length) {
    Logger.log('backfillMCFFees: nothing to process');
    return;
  }

  var now = new Date(Date.now() - 5 * 60 * 1000); // 5-min buffer for clock drift
  if (!minDate) minDate = new Date(now.getTime() - 180 * 24 * 3600 * 1000); // fallback: 180 days

  Logger.log('backfillMCFFees: %s rows pending, window %s → now', pending.length, minDate.toISOString().slice(0, 10));

  // Fetch all financial events in 60-day windows for one endpoint
  function fetchFeeMap(ep) {
    var feeMap      = {};
    var windowStart = new Date(minDate);

    while (windowStart < now) {
      var windowEnd = new Date(windowStart);
      windowEnd.setDate(windowEnd.getDate() + 60);
      if (windowEnd > now) windowEnd = now;
      if (windowStart >= windowEnd) break;

      try {
        var chunk = _buildFeeMapForWindow(ep, windowStart, windowEnd);
        var keys  = Object.keys(chunk);
        keys.forEach(function(k) { feeMap[k] = chunk[k]; });
        Logger.log('  [%s] %s – %s: %s orders found', ep,
          windowStart.toISOString().slice(0, 10), windowEnd.toISOString().slice(0, 10), keys.length);
      } catch (e) {
        if (_isRateLimit429(e)) {
          Logger.log('  [%s] 429 — sleeping 15 s, retrying window', ep);
          Utilities.sleep(15000);
          try {
            var chunk2 = _buildFeeMapForWindow(ep, windowStart, windowEnd);
            Object.keys(chunk2).forEach(function(k) { feeMap[k] = chunk2[k]; });
          } catch (e2) {
            Logger.log('  [%s] retry also failed: %s', ep, e2.message);
          }
        } else {
          Logger.log('  [%s] window error (%s): %s', ep, windowStart.toISOString().slice(0, 10), e.message);
        }
      }

      windowStart = new Date(windowEnd);
      windowStart.setDate(windowStart.getDate() + 1);
      Utilities.sleep(500); // stay under SP-API rate limit between windows
    }

    return feeMap;
  }

  var euMap = fetchFeeMap('EU');
  var feMap = hasJP ? fetchFeeMap('FE') : {};

  // Write fees — track unmatched rows for displayableOrderId fallback
  var written = 0, notSettled = 0;
  var unfilledRows = [];

  pending.forEach(function(r) {
    var primaryMap   = r.isJP ? feMap : euMap;
    var secondaryMap = r.isJP ? euMap : feMap;
    var fee = primaryMap[r.orderId]   !== undefined ? primaryMap[r.orderId]
            : secondaryMap[r.orderId] !== undefined ? secondaryMap[r.orderId]
            : null;

    if (fee !== null) {
      sheet.getRange(BF_START_ROW + r.i, BF_COL_FEE).setValue(fee);
      Logger.log('Row %s (%s): fee = %s', BF_START_ROW + r.i, r.orderId, fee);
      written++;
    } else {
      unfilledRows.push(r);
    }
  });

  // Fallback: some MCF orders settle in Finances API under displayableOrderId
  // (e.g. when the MCF order is linked to an Amazon marketplace order).
  // Call getFulfillmentOrderRaw per unmatched row to resolve the alternate ID.
  if (unfilledRows.length) {
    Logger.log('backfillMCFFees: %s rows unmatched — trying displayableOrderId fallback', unfilledRows.length);
    unfilledRows.forEach(function(r) {
      var resolved = false;
      try {
        var ep       = r.isJP ? 'FE' : 'EU';
        var foResult = getFulfillmentOrderRaw(r.orderId, ep);
        var dispId   = ((foResult.fulfillmentOrder || {}).displayableOrderId || '').trim();
        if (dispId && dispId !== r.orderId) {
          var primaryMap   = r.isJP ? feMap : euMap;
          var secondaryMap = r.isJP ? euMap : feMap;
          var fee2 = primaryMap[dispId]   !== undefined ? primaryMap[dispId]
                   : secondaryMap[dispId] !== undefined ? secondaryMap[dispId]
                   : null;
          if (fee2 !== null) {
            sheet.getRange(BF_START_ROW + r.i, BF_COL_FEE).setValue(fee2);
            Logger.log('Row %s (%s → %s): fee = %s via displayableOrderId', BF_START_ROW + r.i, r.orderId, dispId, fee2);
            written++;
            resolved = true;
          }
        }
      } catch (e) {
        Logger.log('Row %s (%s): fallback error: %s', BF_START_ROW + r.i, r.orderId, e.message || e);
      }
      if (!resolved) {
        Logger.log('Row %s (%s): not yet settled', BF_START_ROW + r.i, r.orderId);
        notSettled++;
      }
      Utilities.sleep(400);
    });
  }

  Logger.log('backfillMCFFees done — written: %s, not settled: %s', written, notSettled);
}

/**
 * Scoped variant of backfillMCFFees() — only processes rows sent within the last `days` days
 * (default 90).
 *
 * Fee source: GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2 settlement reports via
 * _buildSettlementFeeMap() (see that function's doc comment in sp-api.js for why — short version:
 * confirmed live 2026-07-30 that Amazon's Finances API has no fee data at all for these
 * self-created MCF shipments, but the settlement report's `merchant-order-id` column preserves
 * the exact GCX sheet id, so this is a **direct** match — no numeric alias or
 * displayableOrderId resolution needed, and no per-row SP-API calls at all (unlike the old
 * approach). The settlement report listing endpoint itself refuses `createdSince` older than
 * ~90 days, which is exactly why this function (and not the unbounded backfillMCFFees()) is the
 * one that can actually work — `days` beyond ~89 is clamped down to fit.
 *
 * Run manually from the Apps Script editor, or via the 30-min time-driven trigger.
 */
function backfillMCFFeesRecent(days) {
  // Time-driven triggers call the handler with a trigger event object as the first
  // argument (not undefined) — `days || 90` let that object through as a truthy value,
  // so `days * 24 * 3600 * 1000` became NaN, cutoff became an Invalid Date, and
  // cutoff.toISOString() below threw immediately on every scheduled run (100% failure,
  // confirmed in Executions log: "RangeError: Invalid time value at
  // backfillMCFFeesRecent(autoFill:669:28)"), before ever reaching SP-API.
  days = (typeof days === 'number' && isFinite(days) && days > 0) ? Math.min(days, 89) : 89;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var now      = new Date(Date.now() - 5 * 60 * 1000);
  var cutoff   = new Date(now.getTime() - days * 24 * 3600 * 1000);

  var numRows   = lastRow - BF_START_ROW + 1;
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();
  var sentDates = sheet.getRange(BF_START_ROW, BF_COL_SENT,   numRows, 1).getValues();
  var regions   = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var existing  = sheet.getRange(BF_START_ROW, BF_COL_FEE,    numRows, 1).getValues();
  var notes     = sheet.getRange(BF_START_ROW, BF_COL_FEE,    numRows, 1).getNotes();
  // Col R (Tracking Number) — a real value here means the order has genuinely shipped (AMZTK
  // formula resolved), so a getFulfillmentPreview() estimate fallback makes sense. A blank R
  // means the order hasn't shipped yet — no fallback, nothing to estimate against.
  var tracking  = sheet.getRange(BF_START_ROW, RETRY_R_COL,   numRows, 1).getDisplayValues();

  var pending = [];
  var hasJP   = false;

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var curStr = String(existing[i][0] === null || existing[i][0] === undefined ? '' : existing[i][0]).trim();
    // A cell holding an estimate placeholder (fillPreviewEstimatesForRange) is NOT "already
    // filled" for this function's purposes — it should be overwritten as soon as a real
    // settlement fee is found, same as a blank cell.
    var isEstimate = _isFeeEstimateNote(notes[i][0]);
    if (curStr !== '' && curStr !== 'RETRY' && !_isErrorValue(curStr) && !isEstimate) continue;

    var sentDateStr = String(sentDates[i][0] || '').trim();
    if (!sentDateStr) continue; // no sentDate — can't confirm it's within the recent window, skip here
    var d = new Date(sentDateStr);
    if (d < cutoff) continue; // older than the window — leave for a full backfillMCFFees() run

    var isJP = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    if (isJP) hasJP = true;

    var rStr = String(tracking[i][0] || '').trim();
    var hasTracking = rStr !== '' && !_isErrorValue(rStr);

    pending.push({ i: i, orderId: orderId, sentDate: sentDateStr, isJP: isJP, hasTracking: hasTracking, isEstimate: isEstimate });
  }

  if (!pending.length) {
    Logger.log('backfillMCFFeesRecent: nothing pending within the last %s days.', days);
    return;
  }

  Logger.log('backfillMCFFeesRecent: %s rows pending, window %s → now (last %s days)',
    pending.length, cutoff.toISOString().slice(0, 10), days);

  var euMap = _buildSettlementFeeMap('EU', cutoff);
  var feMap = hasJP ? _buildSettlementFeeMap('FE', cutoff) : {};

  var written = 0, notSettled = 0;
  var needsEstimate = [];

  pending.forEach(function(r) {
    var map = r.isJP ? feMap : euMap;
    var fee = map[r.orderId];

    if (fee !== undefined) {
      // .setNote('') clears any estimate-placeholder note — this is now the real settled fee.
      sheet.getRange(BF_START_ROW + r.i, BF_COL_FEE).setValue(fee).setNote('');
      Logger.log('Row %s (%s): fee = %s', BF_START_ROW + r.i, r.orderId, fee);
      written++;
    } else if (r.hasTracking) {
      // Not settled yet, but the order has genuinely shipped (real tracking number in col R) —
      // queue for a getFulfillmentPreview() estimate fallback below instead of leaving blank.
      needsEstimate.push(r);
    } else {
      notSettled++; // hasn't even shipped yet — nothing to estimate against
    }
  });

  // Estimate fallback (added 2026-08-13): for rows that have shipped but haven't settled yet,
  // write a getFulfillmentPreview() ESTIMATE so col Y is never left blank once a real tracking
  // number exists — same mechanism fillPreviewEstimatesForRange() used manually, now automatic
  // every 30 min. Tagged with FEE_ESTIMATE_NOTE_PREFIX so the "already filled" check above treats
  // it as still-pending and overwrites it with the real fee on a later run once settlement data
  // actually arrives — no separate cleanup step needed.
  //
  // Capped and paced separately from the settlement-report scan above (own time budget + circuit
  // breaker) so a burst of newly-shipped rows can't blow past GAS's execution limit on top of
  // whatever the settlement scan already used.
  var estimated = 0;
  var maxEstimatesPerRun = 30;
  var estimateBudgetMs   = 90 * 1000;
  var estStart  = Date.now();
  var consec429 = 0;

  // Bug fixed 2026-08-14: rows that already carry an estimate stay in `pending` every run (so a
  // real fee can overwrite them) and were being re-estimated ahead of rows that had NEVER been
  // estimated, in plain row order. With more already-estimated rows than the 30/run cap, the same
  // low-numbered rows consumed the entire budget every single run — confirmed live: 24h after the
  // first estimate pass, backfillMCFFeesRecent was still only ever re-touching rows 2782-2823,
  // never advancing to newer orders (Aug 12-14 sends) despite running every 30 min without error.
  // Fix: never-estimated rows go first; already-estimated rows (a refresh, not a first fill) only
  // get whatever budget is left over.
  var neverEstimated = needsEstimate.filter(function(r) { return !r.isEstimate; });
  var alreadyEstimated = needsEstimate.filter(function(r) { return r.isEstimate; });
  var toEstimate = neverEstimated.concat(alreadyEstimated).slice(0, maxEstimatesPerRun);

  for (var e = 0; e < toEstimate.length; e++) {
    if (Date.now() - estStart > estimateBudgetMs) {
      Logger.log('backfillMCFFeesRecent: estimate-fallback time budget reached — %s/%s attempted', e, toEstimate.length);
      break;
    }
    if (consec429 >= 5) {
      Logger.log('backfillMCFFeesRecent: %s consecutive 429s on estimate fallback — stopping early', consec429);
      break;
    }

    var r2 = toEstimate[e];
    var endpoints = r2.isJP ? ['FE', 'EU'] : ['EU', 'FE'];
    var fee2 = '', err2 = null;

    for (var k = 0; k < endpoints.length; k++) {
      try {
        fee2 = _fetchMcfFeePreview(r2.orderId, endpoints[k]);
        err2 = null;
        consec429 = 0;
        break;
      } catch (errK) {
        err2 = errK;
        if (_isRetryableRegionMismatchError(errK) || _isNoOrderInfoError(errK)) continue;
        if (_isRateLimit429(errK)) consec429++;
        break;
      }
    }

    if (!err2 && fee2 !== '' && fee2 !== null && fee2 !== undefined) {
      var noteStamp = FEE_ESTIMATE_NOTE_PREFIX + ' — written ' + new Date().toISOString() +
        '. NOT the actual charged fee; will be auto-replaced by backfillMCFFeesRecent() once real settlement data arrives.';
      sheet.getRange(BF_START_ROW + r2.i, BF_COL_FEE).setValue(fee2).setNote(noteStamp);
      Logger.log('Row %s (%s): estimate fallback = %s', BF_START_ROW + r2.i, r2.orderId, fee2);
      estimated++;
    } else {
      if (err2) Logger.log('Row %s (%s): estimate fallback error — %s', BF_START_ROW + r2.i, r2.orderId, err2.message || err2);
      notSettled++;
    }
    Utilities.sleep(500);
  }
  notSettled += (needsEstimate.length - toEstimate.length); // deferred to next run, not lost

  Logger.log('backfillMCFFeesRecent done — written: %s, estimated: %s, not settled: %s (older rows outside the %s-day window untouched)',
    written, estimated, notSettled, days);
}

/**
 * Fills col Y blanks in [startRow, endRow] with a getFulfillmentPreview() ESTIMATE — a fresh
 * shipping-quote using today's rates, NOT the actual fee Amazon charged for that historical
 * shipment. Only use this when the real source (settlement reports, via backfillMCFFeesRecent)
 * has genuinely nothing yet — confirmed live 2026-08-11: for these self-created MCF orders, the
 * Finances API never carries this fee at all (see README "Col Y fee source" section), and
 * settlement reports only have it once Amazon closes the relevant settlement period.
 *
 * Each written cell gets a note tagged with FEE_ESTIMATE_NOTE_PREFIX so backfillMCFFeesRecent()
 * knows to overwrite it with the real value once settlement data for that order actually posts —
 * it is not treated as "already filled" the way a normal written value is.
 *
 * Run manually from the Apps Script editor: fillPreviewEstimatesForRange(2782, 2909)
 */
function fillPreviewEstimatesForRange(startRow, endRow) {
  if (!startRow || !endRow || endRow < startRow) throw new Error('Provide a valid startRow/endRow, e.g. fillPreviewEstimatesForRange(2782, 2909)');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var numRows  = endRow - startRow + 1;
  var orderIds = sheet.getRange(startRow, BF_COL_ORDER,  numRows, 1).getValues();
  var regions  = sheet.getRange(startRow, BF_COL_REGION, numRows, 1).getValues();
  var existing = sheet.getRange(startRow, BF_COL_FEE,    numRows, 1).getValues();

  var pending = [];
  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;
    var curStr = String(existing[i][0] === null || existing[i][0] === undefined ? '' : existing[i][0]).trim();
    if (curStr !== '' && curStr !== 'RETRY' && !_isErrorValue(curStr)) continue; // already has a value
    var isJP = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    pending.push({ i: i, orderId: orderId, isJP: isJP });
  }

  if (!pending.length) {
    Logger.log('fillPreviewEstimatesForRange: nothing blank in %s-%s', startRow, endRow);
    return;
  }
  Logger.log('fillPreviewEstimatesForRange: %s blank row(s) in %s-%s', pending.length, startRow, endRow);

  var startTime    = Date.now();
  var timeBudgetMs = 4.5 * 60 * 1000;
  var maxConsec429 = 5;
  var consec429    = 0;
  var noteStamp    = FEE_ESTIMATE_NOTE_PREFIX + ' — written ' + new Date().toISOString() +
    '. NOT the actual charged fee; will be auto-replaced by backfillMCFFeesRecent() once real settlement data arrives.';

  var filled = 0, skipped = 0, errored = 0;

  for (var p = 0; p < pending.length; p++) {
    if (Date.now() - startTime > timeBudgetMs) {
      Logger.log('fillPreviewEstimatesForRange: time budget reached — stopped at %s/%s', p, pending.length);
      break;
    }
    if (consec429 >= maxConsec429) {
      Logger.log('fillPreviewEstimatesForRange: %s consecutive 429s — stopping early (%s/%s attempted)', consec429, p, pending.length);
      break;
    }

    var r = pending[p];
    var endpoints = r.isJP ? ['FE', 'EU'] : ['EU', 'FE'];
    var fee = '';
    var lastErr = null;

    for (var e = 0; e < endpoints.length; e++) {
      try {
        fee = _fetchMcfFeePreview(r.orderId, endpoints[e]);
        lastErr = null;
        consec429 = 0;
        break;
      } catch (err) {
        lastErr = err;
        if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
        if (_isRateLimit429(err)) { consec429++; break; }
        break; // other errors: don't try the fallback endpoint
      }
    }

    if (lastErr) {
      Logger.log('Row %s (%s): error — %s', startRow + r.i, r.orderId, lastErr.message || lastErr);
      errored++;
    } else if (fee === '' || fee === null || fee === undefined) {
      Logger.log('Row %s (%s): no preview available (cancelled / missing items?)', startRow + r.i, r.orderId);
      skipped++;
    } else {
      sheet.getRange(startRow + r.i, BF_COL_FEE).setValue(fee).setNote(noteStamp);
      Logger.log('Row %s (%s): estimate = %s', startRow + r.i, r.orderId, fee);
      filled++;
    }

    Utilities.sleep(500); // pace preview calls (2 SP-API calls per row: getFulfillmentOrderRaw + preview)
  }

  Logger.log('fillPreviewEstimatesForRange done — filled: %s, no-preview: %s, errored: %s (out of %s pending)',
    filled, skipped, errored, pending.length);
}

/**
 * STEP 1 — run this first to undo the bad freeze.
 * Scans for HYPERLINK formulas whose "tracking number" looks like a price/margin value
 * (see _looksLikePriceNotTracking()) and restores the original
 * =IF(OR(B…,Q…),"",IF(B…<>"JP", HYPERLINK(…AMZTK…), HYPERLINK(…AMZTK_JP…))) formula.
 *
 * FIXED 2026-07-31: this used to have its own inline check (`parseFloat(m[2]) >= 500`) instead of
 * calling the shared _looksLikePriceNotTracking(). That threshold only caught EUR/GBP-style
 * decimal prices and completely missed JPY-style col Z margin values (bare integers, no decimal,
 * routinely in the thousands — e.g. "2799") — a real corruption incident on JP rows went
 * undetected by this exact function until caught manually. Now uses the same, more reliable
 * shared heuristic everywhere so the two checks can't silently drift apart again.
 */
function unfreezeAmztkFormulas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < BF_START_ROW) return;

  var numRows     = lastRow - BF_START_ROW + 1;
  var allFormulas = sheet.getRange(BF_START_ROW, 1, numRows, lastCol).getFormulas();
  // Matches =HYPERLINK("https://www.swiship.XX/track?id=2799","2799")
  var feePattern  = /^=HYPERLINK\("https:\/\/www\.swiship\.(de|jp)\/track\?id=(\d+\.?\d*)","(\d+\.?\d*)"\)$/i;

  var restored = 0;
  for (var i = 0; i < numRows; i++) {
    var rowNum = BF_START_ROW + i;
    var bRef = 'B' + rowNum;
    var qRef = 'Q' + rowNum;
    for (var c = 0; c < lastCol; c++) {
      var f = allFormulas[i][c] || '';
      var m = f.match(feePattern);
      if (!m) continue;
      if (!_looksLikePriceNotTracking(m[2])) continue;
      var orig =
        '=IF(OR(' + bRef + '="",' + qRef + '=""),"",IF(' + bRef + '<>"JP",' +
        'HYPERLINK("https://www.swiship.de/track?id="&AMZTK(' + qRef + '),AMZTK(' + qRef + ')),' +
        'HYPERLINK("https://www.swiship.jp/track?id="&AMZTK_JP(' + qRef + '),AMZTK_JP(' + qRef + '))))';
      sheet.getRange(rowNum, c + 1).setFormula(orig);
      restored++;
    }
  }
  Logger.log('unfreezeAmztkFormulas done — restored: ' + restored + ' cells');
  SpreadsheetApp.getUi().alert('수식 복원 완료: ' + restored + '개 셀');
}

/**
 * STEP 2 — run after unfreezeAmztkFormulas() (and after EU LWA is fixed).
 * Fetches real tracking numbers from SP-API only (never reads col Z as a source,
 * since col Z contains fees in this sheet).
 *
 * - JP rows: fetched via FE endpoint immediately.
 * - EU rows: skipped until EU LWA refresh token is re-authorised.
 * - Replaces each AMZTK formula cell with =HYPERLINK("swiship.XX/…","TN").
 * - Safe to run multiple times — skips cells that already have a non-AMZTK formula.
 */
function freezeAmztkFormulas() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  _warmLwaTokens();

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < BF_START_ROW) { Logger.log('No data rows.'); return; }

  var numRows = lastRow - BF_START_ROW + 1;

  var allFormulas = sheet.getRange(BF_START_ROW, 1, numRows, lastCol).getFormulas();
  var regions     = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var orderIds    = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();

  var frozen = 0, pending = 0, skippedEU = 0;

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var amztkCols = [];
    for (var c = 0; c < lastCol; c++) {
      if ((allFormulas[i][c] || '').toUpperCase().indexOf('AMZTK') !== -1) {
        amztkCols.push(c + 1);
      }
    }
    if (!amztkCols.length) continue;

    var isJP = String(regions[i][0] || '').trim().toUpperCase() === 'JP';

    // EU rows: skip until LWA is re-authorised.
    if (!isJP) {
      Logger.log('Row ' + (BF_START_ROW + i) + ': EU — skipping (LWA broken)');
      skippedEU++;
      continue;
    }

    // JP rows: fetch from FE endpoint, with outer 429 retry (45 s backoff, up to 3×).
    var tn = '';
    var retries429 = 0;
    var fetchOk = false;
    while (!fetchOk) {
      try {
        var tracks = _tracksWithFallbacks(orderId, isJP ? ['FE', 'EU'] : ['EU', 'FE']);
        tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
          ? tracks[0].trackingNumber.trim() : '';
        Utilities.sleep(400);
        fetchOk = true;
      } catch (e) {
        if (_isRateLimit429(e) && retries429 < 3) {
          retries429++;
          Logger.log('Row ' + (BF_START_ROW + i) + ': 429 — waiting 45 s (retry ' + retries429 + '/3)');
          Utilities.sleep(45000);
        } else {
          Logger.log('Row ' + (BF_START_ROW + i) + ': fetch error — ' + e.message);
          pending++;
          break;
        }
      }
    }
    if (!fetchOk) continue;

    if (!tn) { pending++; continue; }

    var url           = 'https://www.swiship.jp/track?id=' + tn;
    var staticFormula = '=HYPERLINK("' + url + '","' + tn + '")';
    for (var ci = 0; ci < amztkCols.length; ci++) {
      sheet.getRange(BF_START_ROW + i, amztkCols[ci]).setFormula(staticFormula);
    }
    Logger.log('Row ' + (BF_START_ROW + i) + ': frozen → ' + tn);
    frozen++;
  }

  Logger.log(
    'freezeAmztkFormulas done — frozen: ' + frozen +
    ', EU skipped: ' + skippedEU +
    ', pending (no TN yet): ' + pending
  );
  SpreadsheetApp.getUi().alert(
    '완료\n\n' +
    '✅ 고정됨 (JP): ' + frozen + '행\n' +
    '⏭ EU 스킵 (LWA 재인증 필요): ' + skippedEU + '행\n' +
    '⏳ 아직 운송장 없음: ' + pending + '행'
  );
}

function onEdit_mcf(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  const row = editedCell.getRow();
  const col = editedCell.getColumn();
  const value = String(editedCell.getValue()).trim().toUpperCase();

  // Only run in MCF sheet
  if (sheet.getName() !== 'MCF 발송 로그') return;
  if (row < 4) return;  // ignore header rows

  /**********************************************
   * 1) STOCK CHECK (AB column = 28)
   **********************************************/
  if (col === 28 && value === "STOCK") {
    runStockCheckOnly(sheet, row);
    return;
  }

  /**********************************************
   * 2) FULL MCF RUN (W column = 23)
   **********************************************/
  if (col === 23 && value === "RUN") {
    processMCFRow(sheet, row);
    return;
  }

  /**********************************************
   * 3) ORIGINAL LOGIC
   **********************************************/

  // I → M (col 9 → 13) insert date once
  if (col === 9) {
    const targetCell = sheet.getRange(row, 13);
    if (!targetCell.getValue()) {
      const formattedDate = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
      targetCell.setValue(formattedDate);
    }
  }

  // N → S (col 14 → 19) set Pending
  if (col === 14) {
    const statusCell = sheet.getRange(row, 19);
    const nValue = editedCell.getValue();
    if (nValue !== "" && !statusCell.getValue()) {
      statusCell.setValue("Pending");
    }
  }

  // U → P + S (col 21 → 16 & 19)
  if (col === 21) {
    const uValue = editedCell.getValue();
    const pCell = sheet.getRange(row, 16);
    const sCell = sheet.getRange(row, 19);

    if (uValue !== "") {
      if (!pCell.getValue()) {
        const formattedDate = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
        pCell.setValue(formattedDate);
      }
      sCell.setValue("MCF");
    }
  }

  // F column → update H
  if (col === 6) { 
    updateMcfStockForRow(sheet, row);
  }

  /**********************************************
   * 4) NEW RULE — Y column triggers T + W
   *    Y = col 25
   *    T = col 20 (date)
   *    W = col 23 ("MCF")
   **********************************************/
  if (col === 25) {
    const yValue = editedCell.getValue().toString().trim();

    if (yValue !== "") {
      const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

      // T column (20)
      const tCell = sheet.getRange(row, 20);
      if (!tCell.getValue()) {
        tCell.setValue(today);
      }

      // W column (23)
      const wCell = sheet.getRange(row, 23);
      wCell.setValue("MCF");
    }

    return;
  }
}

/**
 * Converts col R (Tracking Number) cells from a LIVE =AMZTK()/=AMZTK_JP() formula to a static
 * =HYPERLINK() referencing the already-resolved tracking number in col Z, for any row where
 * BF_COL_RESULT (Z) is already filled.
 *
 * Why this exists (confirmed live, 2026-07-30): 2,489 of 3,755 rows in col R still had a live
 * AMZTK()/AMZTK_JP() formula. Every time Google Sheets does a broad recalculation (opening the
 * file, an edit elsewhere, etc.) it fires a burst of dozens-to-hundreds of these custom-function
 * calls simultaneously — confirmed in the Executions log, ~50+ AMZTK/AMZTK_JP calls within the
 * same 3-second window — which floods the FBA Outbound API (GetFulfillmentOrder) and produces
 * fresh "EU ERR: SP-API error 429" text in col R almost immediately, faster than the hourly
 * retryR429Errors() trigger (capped at 40 rows/run, and only scanning row 1108+) can keep up
 * with. Meanwhile 2,817 rows already have a resolved tracking number sitting unused in col Z
 * (from backfillTrackingNumbers()) — col R was never actually wired to read it.
 *
 * This function costs zero SP-API calls — it's a pure spreadsheet read/write — so it's safe to
 * run any time. Only rewrites cells whose current formula contains "AMZTK" AND whose Z value is
 * non-empty; every other cell (already-frozen HYPERLINK, blank, or some other static value) is
 * left completely untouched, to avoid ever clobbering a cell that isn't a live AMZTK formula.
 * Batches contiguous qualifying rows into single setFormulas() calls for speed.
 *
 * BUG FIXED 2026-07-30 (same day, caught by user within minutes): the first version of this
 * function only checked that col Z was non-empty before freezing it into col R as "the tracking
 * number" — it never validated that the value actually looked like one. For a batch of rows, col
 * Z held a plain price (e.g. "17.99") instead of a tracking number, and got frozen straight into
 * col R, displaying prices where tracking numbers should be. Reverted live via the existing
 * unfreezeAmztkFormulas() (already built for exactly this "tracking number looks like a price"
 * shape) and fixed here with the same validation that function already uses: a bare number under
 * 500 is treated as a price, never a tracking number, and is now skipped rather than frozen —
 * that row's live AMZTK formula is left alone so it keeps trying to resolve the real value.
 *
 * Run manually from the Apps Script editor.
 */
// FIXED 2026-07-31: the original "< 500" threshold only caught EUR/GBP-style decimal prices
// (e.g. "17.99"). It completely missed JPY-style col Z margin values, which are bare integers
// with NO decimal point and are routinely in the thousands (e.g. "2799", "4712") — well over 500.
// Confirmed live: JP rows corrupted by the same freezeTrackingColumnR() bug slipped through this
// exact check undetected, in both the original incident and the first "clean" verification pass.
// Real tracking numbers observed on this sheet are either alphanumeric with a country prefix
// ("UK4618438146", "JJD000390016584418318") or a bare digit string of 12+ digits
// ("371434845460", "00340434685209460376") — never a decimal, never under 8 digits. That's a much
// safer signal than any currency-shaped magnitude threshold.
function _looksLikePriceNotTracking(v) {
  if (!/^\d+\.?\d*$/.test(v)) return false;   // has letters → real tracking-number shape, not a price
  if (v.indexOf('.') >= 0) return true;        // any decimal point at all → price, never a tracking number
  return v.length < 8;                         // short bare integer → margin/price, not a 12+ digit tracking number
}

// DISABLED 2026-07-31 — see backfillTrackingNumbers()'s doc comment. Col Z (BF_COL_RESULT) is a
// live profit-margin ARRAYFORMULA, not a resolved-tracking-number cache, so freezing its value
// into col R produces fake tracking numbers that are actually prices. This is what corrupted col R
// earlier today; the _looksLikePriceNotTracking() guard below only caught values under 500, which
// is not a real fix — it's disabled outright until col R's 429 problem gets a real target column
// (or a different approach) instead of Z.
function freezeTrackingColumnR() {
  throw new Error('freezeTrackingColumnR() is disabled: col Z (BF_COL_RESULT) is a live profit-margin ' +
    'ARRAYFORMULA, not a tracking-number cache. See the function doc comment before re-enabling.');
  // eslint-disable-next-line no-unreachable
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows   = lastRow - BF_START_ROW + 1;
  var rFormulas = sheet.getRange(BF_START_ROW, RETRY_R_COL,  numRows, 1).getFormulas();
  var zValues   = sheet.getRange(BF_START_ROW, BF_COL_RESULT, numRows, 1).getValues();
  var regions   = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();

  var skippedPriceLike = 0;

  function qualifies(i) {
    var f = rFormulas[i][0] || '';
    var z = String(zValues[i][0] || '').trim();
    var orderId = String(orderIds[i][0] || '').trim();
    if (f.toUpperCase().indexOf('AMZTK') < 0 || z === '' || orderId === '') return false;
    if (_looksLikePriceNotTracking(z)) { skippedPriceLike++; return false; }
    return true;
  }

  function formulaFor(i) {
    var z = String(zValues[i][0]).trim();
    var isJP = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    var domain = isJP ? 'jp' : 'de';
    return '=HYPERLINK("https://www.swiship.' + domain + '/track?id=' + z + '","' + z + '")';
  }

  var frozen = 0, runStart = -1;
  for (var i = 0; i <= numRows; i++) {
    var q = i < numRows && qualifies(i);
    if (q && runStart < 0) {
      runStart = i;
    } else if (!q && runStart >= 0) {
      var runLen = i - runStart;
      var batch = [];
      for (var j = runStart; j < i; j++) batch.push([formulaFor(j)]);
      sheet.getRange(BF_START_ROW + runStart, RETRY_R_COL, runLen, 1).setFormulas(batch);
      frozen += runLen;
      runStart = -1;
    }
  }

  Logger.log('freezeTrackingColumnR done — %s cell(s) converted from live AMZTK formula to static HYPERLINK (zero SP-API calls). %s row(s) skipped because col Z held a price-like value, not a tracking number.',
    frozen, skippedPriceLike);
}

// DISABLED 2026-07-31 — both functions this calls are disabled (see their doc comments); the
// underlying premise (col Z as a tracking-number cache) is false. The trigger for this function
// was already deleted live during the col R corruption cleanup; this stub just makes sure a stray
// manual run or a re-added trigger fails loudly instead of writing bad data again.
function dailyTrackingMaintenance() {
  throw new Error('dailyTrackingMaintenance() is disabled: it composes backfillTrackingNumbers() and ' +
    'freezeTrackingColumnR(), both disabled because col Z is not a tracking-number cache.');
}
