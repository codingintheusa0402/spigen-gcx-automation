function setVoucherDropdown_LazadaLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lazada log");
  const range = sheet.getRange("Q2:Q9000"); // Adjust range if needed

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "Provide 100% voucher",
      "Provide 50% voucher",
      "Provide 10% voucher"
    ], true) // true = show dropdown
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
}

function debug_setVoucherDropdown_LazadaLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lazada log");
  if (!sheet) {
    Logger.log("Sheet not found. Check sheet name.");
    return;
  }

  const range = sheet.getRange("Q2:Q9000");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "Provide 100% voucher",
      "Provide 50% voucher",
      "Provide 10% voucher"
    ], true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
  Logger.log("Dropdown successfully applied to Q2:Q9000");
}

