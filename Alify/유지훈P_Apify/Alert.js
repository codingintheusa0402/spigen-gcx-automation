const SPREADSHEET_ID = '1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw';
const TARGET_SHEET_NAME = 'US';
const TARGET_COLUMN = 11; // K
const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAFjxOPoY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=qrL3TAfjDsSFJaxWlToGZZZ1Rqp4wQ7N1ycJC6_pPwM';

/**
 * Installable onEdit trigger
 * Sends a Google Chat alert when Column K is edited in the US sheet
 */
function onEditAlertToGoogleChat(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();
    const spreadsheet = sheet.getParent();

    if (spreadsheet.getId() !== SPREADSHEET_ID) return;
    if (sheet.getName() !== TARGET_SHEET_NAME) return;
    if (range.getColumn() !== TARGET_COLUMN) return;
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

    const a1 = range.getA1Notation();
    const newValue = range.getDisplayValue();
    const oldValue = e.oldValue || '';

    const sheetUrl =
      `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=${sheet.getSheetId()}`;

    const message =
      `Old value: ${oldValue}\n` +
      `New value: ${newValue}\n` +
      `Sheet: ${sheet.getName()}\n` +
      `Cell: ${a1}\n` +
      `Link: ${sheetUrl}`;

    sendGoogleChatMessage(message);

  } catch (error) {
    sendGoogleChatMessage(`Error: ${error.message}`);
  }
}

/**
 * Sends a plain text message to Google Chat
 */
function sendGoogleChatMessage(text) {
  const response = UrlFetchApp.fetch(CHAT_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true
  });

  Logger.log(`Webhook response code: ${response.getResponseCode()}`);
  Logger.log(`Webhook response body: ${response.getContentText()}`);
}

/**
 * Run once to create the installable onEdit trigger
 */
function createEditTrigger() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditAlertToGoogleChat') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEditAlertToGoogleChat')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('Installable onEdit trigger created successfully.');
}

/**
 * Test the webhook directly
 */
function testSendChat() {
  sendGoogleChatMessage('Test message from Apps Script: webhook is working.');
}

/**
 * Check existing triggers
 */
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger, i) => {
    Logger.log(
      `#${i + 1} Function=${trigger.getHandlerFunction()}, EventType=${trigger.getEventType()}`
    );
  });
}