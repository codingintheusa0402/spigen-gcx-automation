function MCFReporter() {
  // var webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=b4zApCmKNq1pPBDmemgVv1Y8xoXm4h_w_eKccjtqCiI'; // Private 
  var webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQAdqYt1ro/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=SL-5NCPrO2MdkzC_6deniWr5y5n7C7aB4k6k2sVWGO8'; // GCX T2 ESC. Ticket 
  
  var sheetId = '1g6a-S7eeA1oY19aTEFhTNAyp2A5nLqLNkPRqOqriWfc';
  var sheetName = 'MCF 발송 로그';
  var sheetGid = '1608794212';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  var missingRows = [];
  for (var i = 1; i < data.length; i++) {
    var rCol = data[i][17];
    var sCol = data[i][18];
    var vCol = data[i][21];
    if (rCol && !sCol) {
      missingRows.push({ rowNum: i + 1, 담당자: vCol ? vCol : "(담당자 없음)" });
    }
  }

  var count = missingRows.length;
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var summaryText = "Tracking No. 미기입: " + count + "개  (" + today + ")";

  var rowButtons = [];
  for (var j = 0; j < missingRows.length; j++) {
    var row = missingRows[j];
    rowButtons.push({
      "text": row.rowNum + "행 이동",
      "onClick": {
        "openLink": {
          "url": "https://docs.google.com/spreadsheets/d/" + sheetId +
                 "/edit?gid=" + sheetGid + "&range=" + row.rowNum + ":" + row.rowNum
        }
      }
    });
  }

  var payload = {
    "cardsV2": [
      {
        "cardId": "tracking-alert",
        "card": {
          "header": {
            "title": "MCF Daily Report",
            "subtitle": summaryText
          },
          "sections": [
            {
              "widgets": [
                {
                  "buttonList": {
                    "buttons": rowButtons
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var attempts = 0;
  var maxAttempts = 5;
  var sent = false;

  while (!sent && attempts < maxAttempts) {
    attempts++;
    try {
      var res = UrlFetchApp.fetch(webhookUrl, options);
      var code = res.getResponseCode();
      Logger.log("Attempt " + attempts + ": Code " + code + ", Body: " + res.getContentText());
      if (code >= 200 && code < 300) {
        sent = true;
        Logger.log("Message sent successfully.");
      } else {
        Utilities.sleep(2000);
      }
    } catch (err) {
      Logger.log("Attempt " + attempts + " failed: " + err);
      Utilities.sleep(2000);
    }
  }

  if (!sent) {
    Logger.log("Failed to send message after " + maxAttempts + " attempts.");

    
  }
}
function authorizeSpApi() {
  // Replace with your real marketplace id (EU example)
  var token = getLwaAccessToken('EU');
  Logger.log("Token acquired: " + token);
}