/***** ========= BACKFILL CONFIG ========= *****/
// Adjust these if your sheet layout changes.
var BF_SHEET_NAME  = 'MCF 발송 로그';
var BF_START_ROW   = 4;    // first data row (skip headers)
var BF_COL_REGION  = 2;    // B — "JP" triggers FE-first, anything else EU-first
var BF_COL_ORDER   = 17;   // Q — sellerFulfillmentOrderId
var BF_COL_SENT    = 16;   // P — MCF sent date (yyyy-mm-dd)
var BF_COL_RESULT  = 26;   // Z — static tracking number
                            //     replace =AMZTK(Q…) formula with =IF(Z…="","",HYPERLINK(…Z…,Z…))
var BF_COL_FEE     = 25;   // Y — Transportation Fee (€, ¥, £) written by backfillMCFFees()

/**
 * Writes tracking numbers as static values into BF_COL_RESULT.
 * - Skips rows that already have a valid (non-error) tracking number.
 * - Retries rows whose result cell contains an error string ("EU ERR:…", "ERR:…", etc.).
 * - On 429 / transient error, writes the error back to the cell so the next run retries it.
 *
 * Run manually or set a daily time-based trigger on this function.
 */
function backfillTrackingNumbers() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

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

    try {
      var tracks = _tracksWithFallbacks(orderId, endpoints);
      var tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
        ? tracks[0].trackingNumber.trim()
        : '';

      if (tn) {
        sheet.getRange(BF_START_ROW + i, BF_COL_RESULT).setValue(tn);
        Logger.log('Row ' + (BF_START_ROW + i) + ': wrote ' + tn);
      }
    } catch (e) {
      var errMsg = (isJP ? 'JP' : 'EU') + ' ERR: ' + (e.message || e);
      // Write error back so the next backfill run retries this row.
      sheet.getRange(BF_START_ROW + i, BF_COL_RESULT).setValue(errMsg);
      Logger.log('Row ' + (BF_START_ROW + i) + ': ' + errMsg);
    }

    Utilities.sleep(400); // stay under SP-API rate limit
  }
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
