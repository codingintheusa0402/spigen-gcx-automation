const CONFIG = {
  sourceId:    '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc',
  sourceSheet: '26년 전체문의',
  destId:      '1qxwUjuV3-_0HRS1Bsb3Fsua0n8N6r6GzNnqiv9wRU10',
  destSheet:   'Sheet1',
  chunkSize:   1000   // rows per read-write cycle
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 Mirror')
    .addItem('Mirror now', 'mirrorSheet')
    .addToUi();
}

function mirrorSheet() {
  const src = SpreadsheetApp.openById(CONFIG.sourceId).getSheetByName(CONFIG.sourceSheet);
  const dst = SpreadsheetApp.openById(CONFIG.destId).getSheetByName(CONFIG.destSheet);

  if (!src) throw new Error('Source sheet not found: ' + CONFIG.sourceSheet);
  if (!dst) throw new Error('Dest sheet not found: ' + CONFIG.destSheet);

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();

  dst.clearContents();

  if (lastRow === 0 || lastCol === 0) {
    Logger.log('Source is empty — destination cleared.');
    return;
  }

  // Expand destination dimensions if needed
  if (dst.getMaxRows() < lastRow) {
    dst.insertRowsAfter(dst.getMaxRows(), lastRow - dst.getMaxRows());
  }
  if (dst.getMaxColumns() < lastCol) {
    dst.insertColumnsAfter(dst.getMaxColumns(), lastCol - dst.getMaxColumns());
  }

  // Read and write in chunks to stay within GAS memory/time limits
  for (let row = 1; row <= lastRow; row += CONFIG.chunkSize) {
    const numRows = Math.min(CONFIG.chunkSize, lastRow - row + 1);
    const data = src.getRange(row, 1, numRows, lastCol).getValues();
    dst.getRange(row, 1, numRows, lastCol).setValues(data);
    SpreadsheetApp.flush();
  }

  Logger.log('Mirror complete: ' + lastRow + ' rows × ' + lastCol + ' cols → ' + CONFIG.destSheet);
}
