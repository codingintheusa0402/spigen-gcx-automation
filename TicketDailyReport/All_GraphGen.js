function appendZendeskDailyStatus() {
  const sheetId = '10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo';
  const date = getKoreanFormattedDate();

  const ss = SpreadsheetApp.openById(sheetId);
  const srcSheet = ss.getSheetByName('Zendesk_Daily');
  const destSheet = ss.getSheetByName('All_Graph');

  if (!srcSheet || !destSheet) {
    Logger.log("Either 'Zendesk_Daily' or 'All_Graph' sheet not found.");
    return;
  }

  const statusCol = srcSheet.getRange('B2:B' + srcSheet.getLastRow()).getValues();

  let newCount = 0, openCount = 0, pendingCount = 0;

  statusCol.forEach(row => {
    const status = (row[0] || '').toString().toLowerCase();
    if (status === 'new') newCount++;
    else if (status === 'open') openCount++;
    else if (status === 'pending') pendingCount++;
  });

  const lastRow = destSheet.getLastRow() + 1;
  destSheet.getRange(lastRow, 1, 1, 4).setValues([[date, newCount, openCount, pendingCount]]);

  srcSheet.clearContents(); // Clean after write
}

function getKoreanFormattedDate() {
  const timeZone = 'Asia/Seoul';
  const date = new Date();
  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const formattedDate = Utilities.formatDate(date, timeZone, 'MM/dd');
  const weekday = dayNames[date.getDay()];
  return `${formattedDate} (${weekday})`;
}

function collapseOldRowsIfNeeded() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_Graph');
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const visibleValues = [];

  for (let row = startRow; row <= lastRow; row++) {
    if (!sheet.isRowHiddenByUser(row)) {
      const value = sheet.getRange(row, 1).getValue(); // Col A
      if (value !== "") {
        visibleValues.push(row);
      }
    }
  }

  if (visibleValues.length >= 20) {
    const rowsToCollapse = visibleValues.slice(0, 5);
    rowsToCollapse.forEach(row => {
      sheet.hideRows(row);
    });
    Logger.log(`🔒 Collapsed rows: ${rowsToCollapse.join(', ')}`);
  } else {
    Logger.log(`✅ No need to collapse. Visible rows: ${visibleValues.length}`);
  }
}

