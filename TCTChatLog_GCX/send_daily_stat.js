// ──────────────── Daily Summary Trigger (Esc T2 Count) ──────────────── //
function sendChatMessageWithSheetLinks(summaryText, webhookUrl, sheetLinks) {
  // Build buttons from sheetLinks
  const buttons = [];

  if (sheetLinks && Object.keys(sheetLinks).length > 0) {
    Object.keys(sheetLinks).forEach(name => {
      const url = sheetLinks[name];
      if (!url) return;

      buttons.push({
        text: name,
        onClick: {
          openLink: {
            url: url,
          },
        },
      });
    });
  }

  // cardsV2 payload
  const payload = {
    cardsV2: [
      {
        cardId: "tct-daily-summary",
        card: {
          header: {
            title: "TCT Chat Log_GCX 마감보고",
          },
          sections: [
            {
              widgets: [
                {
                  textParagraph: {
                    // Convert newlines to <br> for Chat
                    text: summaryText.replace(/\n/g, "<br>"),
                  },
                },
              ],
            },
            // Only add button section if we have buttons
            ...(buttons.length
              ? [
                  {
                    header: "시트 링크",
                    widgets: [
                      {
                        buttonList: {
                          buttons: buttons,
                        },
                      },
                    ],
                  },
                ]
              : []),
          ],
        },
      },
    ],
  };

  const params = {
    method: "post",
    contentType: "application/json; charset=UTF-8",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  UrlFetchApp.fetch(webhookUrl, params);
}

function sendDailyEscT2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Lazada log", "Shopee log"];
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Header is on the card; keep body focused on date + counts
  let summaryText = `날짜: ${today}\n\n`;
  const sheetLinks = {};

  sheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 4) {
      // No data rows
      summaryText += `• ${name}: 0 건\n`;
      sheetLinks[name] = `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
      return;
    }

    const escT2Count = sheet
      .getRange(5, 1, lastRow - 4, 1)
      .getValues()
      .filter(row => (row[0] || "").toString().trim() === "Esc T2").length;

    summaryText += `• ${name}: ${escT2Count} 건\n`;
    sheetLinks[name] = `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
  });

  sendChatMessageWithSheetLinks(summaryText, WEBHOOK_TEST, sheetLinks);
}


// ──────────────── Manual Test Message ──────────────── //
function testSendSampleMessageToKevin() {
  const testMessage =
    "@Kevin Test message: Lazada alert triggered manually.\n\n" +
    "Order ID: 123456789\n" +
    "Platform: Lazada SG\n" +
    "Category: Return";

  const testRowUrl = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit#gid=YOUR_GID&range=99:99";

  // Assumes sendChatMessage(webhook, mentionText, userName, userId, rowUrl) is defined elsewhere
  sendChatMessage(testMessage, WEBHOOK_LAZADA, "@Kevin", "users/101864629530502407495", testRowUrl);
}
