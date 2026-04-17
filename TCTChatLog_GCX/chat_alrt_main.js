// const WEBHOOK_LAZADA = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=j8ga57jWt9s0Kqxx8Ha0GnXbuK0srzxBiBJXY0Kchxw'; // 지우P //
// const WEBHOOK_SHOPEE = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=j8ga57jWt9s0Kqxx8Ha0GnXbuK0srzxBiBJXY0Kchxw'; // 지우P //
const WEBHOOK_LAZADA = 'https://chat.googleapis.com/v1/spaces/AAQAOhOAbz0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=q2Ahf2sNgETGt9I9mcmTYUg2vLoEOOqoeCv-aTCTkHI'; // 영신P //
const WEBHOOK_SHOPEE = 'https://chat.googleapis.com/v1/spaces/AAQAOhOAbz0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=q2Ahf2sNgETGt9I9mcmTYUg2vLoEOOqoeCv-aTCTkHI'; // 영신P //
const WEBHOOK_TEST = 'https://chat.googleapis.com/v1/spaces/AAQAdqYt1ro/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=fbZXlPuUEA18sVE7DKblDSYN-C33UuLMoLKLpFdw4-Y'; 

// ──────────────── Triggered When Editing Sheet ──────────────── //
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Ensure the edit is in the correct column and row
  if (col !== 1 || row <= 4) {
    Logger.log(`Edit ignored. Column: ${col}, Row: ${row}`);
    return;
  }

  const status = (sheet.getRange(row, 1).getValue() || '').toString().trim();      // A
  const rawOrderId = (sheet.getRange(row, 2).getValue() || '').toString().trim();  // B
  const platform = (sheet.getRange(row, 4).getValue() || '').toString().trim();    // D
  const cleanOrderId = rawOrderId.replace(/^\//, '') || `row${row}`;
  const uniqueKey = `${sheetName}_${cleanOrderId}_${platform}_${status}`;

  Logger.log(`Status: ${status}, UniqueKey: ${uniqueKey}`);

  const props = PropertiesService.getScriptProperties();
  const sentIds = JSON.parse(props.getProperty("SENT_IDS") || "[]");

  if (status !== 'Esc T2' || sentIds.includes(uniqueKey)) {
    Logger.log('Message not sent: Status not "Esc T2" or already sent.');
    return;
  }

  // Collect rows with 'Esc T2' status in column A
  const escT2Rows = getEscT2Rows(sheet);

  Logger.log(`Esc T2 Rows: ${escT2Rows}`);

  if (escT2Rows.length === 0) {
    Logger.log('No rows with Esc T2 found.');
    return;
  }

  // Send the message with buttons based on the platform
  if (sheetName === 'Lazada log' && platform.includes('Lazada')) {
    sendChatMessageWithButtons(escT2Rows, WEBHOOK_LAZADA, '@Lim', sheetName);
  } else if (sheetName === 'Shopee log' && platform.includes('Shopee')) {
    sendChatMessageWithButtons(escT2Rows, WEBHOOK_SHOPEE, '@Lim', sheetName);
  }

  // Store the sent ID to avoid duplicates
  sentIds.push(uniqueKey);
  props.setProperty("SENT_IDS", JSON.stringify(sentIds));
}

// ──────────────── Get All Rows with "Esc T2" Status ──────────────── //
function getEscT2Rows(sheet) {
  const columnA = sheet.getRange('A5:A').getValues();  // Get all values from column A (starting from row 5)
  const rowsWithEscT2 = [];

  columnA.forEach((row, index) => {
    const status = (row[0] || '').toString().trim();  // Trim spaces and handle empty values
    Logger.log(`Checking row ${index + 5}: Status = "${status}"`);  // Log the status being checked
    
    if (status === 'Esc T2') {
      const rowNumber = index + 5;  // Adjust for the starting row (A5)
      rowsWithEscT2.push(rowNumber);
    }
  });

  Logger.log(`Rows with Esc T2: ${rowsWithEscT2}`);
  return rowsWithEscT2;
}

// ──────────────── Send Chat Message with Buttons for "Esc T2" Rows ──────────────── //
function sendChatMessageWithButtons(rows, webhook, mention, sheetName) {
  if (rows.length === 0) {
    Logger.log('No rows with Esc T2, no message sent.');
    return;
  }

  // Create the buttons for each row with 'Esc T2'
  const buttons = rows.map((row) => ({
    textButton: {
      text: `Open Claim on ${row}`,
      onClick: { openLink: { url: `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}/edit#gid=${SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetId()}&range=A${row}` } }
    }
  }));

  Logger.log(`Buttons created: ${JSON.stringify(buttons)}`);

  // Message content with sheetName
  const message = `TCT ${sheetName} Esc T2 Created:`;

  // Payload structure
  const payload = {
    cards: [
      {
        sections: [
          {
            widgets: [
              {
                textParagraph: {
                  text: `${mention} ${message}`,  // Use the formatted message
                }
              },
              {
                buttons: buttons  // Add the buttons for each Esc T2 row
              }
            ]
          }
        ]
      }
    ]
  };

  Logger.log(`Payload being sent: ${JSON.stringify(payload)}`);

  // Send the request to the webhook
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhook, options);
  Logger.log(response.getContentText()); // Log the response from the webhook for debugging
}

// ──────────────── Reset All Esc T2 Alerts ──────────────── //
function resetSentAlerts() {
  PropertiesService.getScriptProperties().deleteProperty("SENT_IDS");
}
