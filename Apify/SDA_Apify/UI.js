/**********************************************************
 * UI MENU (SAFE FOR onOpen SIMPLE TRIGGER)
 **********************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Apify')
    .addSubMenu(
      ui.createMenu('Product')
        .addItem('Run Product (auto polling)', 'menuRunProduct')
        .addSeparator()
        .addItem('Cancel Product Polling', 'cancelProductPolling')
    )
    .addToUi();
}


/**********************************************************
 * Product menu handler → uses RECURRING polling
 **********************************************************/
function menuRunProduct() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const pendingRunId = props.getProperty('PRODUCT_LAST_RUN_ID');
  if (pendingRunId) {
    ui.alert(
      'Already running',
      'A Product run is already pending.\n\n' +
      'Run ID:\n' + pendingRunId + '\n\n' +
      'Recurring polling is active and will write the result automatically.',
      ui.ButtonSet.OK
    );

    // Ensure poll trigger still exists (defensive)
    try {
      _scheduleRecurringProductPoll_();
    } catch (e) {
      Logger.log('Failed to ensure recurring poller: ' + e);
    }
    return;
  }

  const pollEvery = Math.max(1, Number(CONFIG.pollIntervalMinutes || 1));

  const choice = ui.alert(
    'Start Apify Product run?',
    'This will:\n' +
      '• Start the Product Apify Task using saved input\n' +
      '• Create a recurring poll (every ' + pollEvery + ' minute[s])\n' +
      '• Write results automatically when the run succeeds\n\n' +
      'Do you want to continue?',
    ui.ButtonSet.OK_CANCEL
  );

  if (choice !== ui.Button.OK) return;

  try {
    runProductNowAndPollRecurring();

    ui.alert(
      'Scheduled',
      'Product run has started.\n\n' +
        'Polling is active and the sheet will be updated automatically\n' +
        'once the run finishes.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'Error starting Product run',
      String(e && e.message ? e.message : e),
      ui.ButtonSet.OK
    );
  }
}
