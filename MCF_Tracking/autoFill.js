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
 * - Skips rows where fee is already filled (never overwrites a value).
 * - Uses P col (sent date) as PostedAfter to skip the fulfillment order lookup.
 * - On 429, writes a retry marker so the next run picks it up again.
 * Run manually or set a daily time-based trigger.
 */
function backfillMCFFees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows  = lastRow - BF_START_ROW + 1;
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER,  numRows, 1).getValues();
  var sentDates = sheet.getRange(BF_START_ROW, BF_COL_SENT,   numRows, 1).getValues();
  var regions   = sheet.getRange(BF_START_ROW, BF_COL_REGION, numRows, 1).getValues();
  var existing  = sheet.getRange(BF_START_ROW, BF_COL_FEE,    numRows, 1).getValues();

  var written = 0, skipped = 0, pending = 0;

  for (var i = 0; i < numRows; i++) {
    var orderId  = String(orderIds[i][0]  || '').trim();
    var sentDate = String(sentDates[i][0] || '').trim();
    if (!orderId) continue;

    var current = existing[i][0];
    // Already has a numeric fee — never overwrite.
    if (current !== '' && current !== null && !_isErrorValue(String(current))) {
      skipped++;
      continue;
    }

    var isJP      = String(regions[i][0] || '').trim().toUpperCase() === 'JP';
    var endpoints = isJP ? ['FE', 'EU'] : ['EU', 'FE'];
    var fee = null;
    var lastErr = null;

    for (var e = 0; e < endpoints.length; e++) {
      try {
        fee = _fetchMcfFeeFinancesApi(orderId, endpoints[e], sentDate || null);
        lastErr = null;
        break;
      } catch (err) {
        lastErr = err;
        if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
        break;
      }
    }

    var cell = sheet.getRange(BF_START_ROW + i, BF_COL_FEE);

    if (lastErr) {
      if (_isRateLimit429(lastErr)) {
        // Write retry marker — next run will retry this row.
        cell.setValue('RETRY');
        Logger.log('Row ' + (BF_START_ROW + i) + ': 429 — marked RETRY');
      } else {
        cell.setValue('ERR: ' + (lastErr.message || lastErr));
        Logger.log('Row ' + (BF_START_ROW + i) + ': ERR — ' + (lastErr.message || lastErr));
      }
      pending++;
    } else if (fee !== '' && fee !== null) {
      cell.setValue(fee);
      Logger.log('Row ' + (BF_START_ROW + i) + ': fee = ' + fee);
      written++;
    } else {
      // Not yet settled — leave blank, next run retries.
      Logger.log('Row ' + (BF_START_ROW + i) + ': not yet settled, skipping');
      pending++;
    }

    Utilities.sleep(400); // stay under SP-API rate limit
  }

  Logger.log('backfillMCFFees done — written: ' + written + ', pending/unsettled: ' + pending + ', already filled: ' + skipped);
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
