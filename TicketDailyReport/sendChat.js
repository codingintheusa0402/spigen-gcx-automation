function all_GraphChartToGoogleChat() {
  const sheetId = '10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo';
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('All_Graph');
  const charts = sheet.getCharts();

  if (charts.length === 0) {
    Logger.log("No charts found.");
    return;
  }

  const chartBlob = charts[0].getAs('image/png');
  const base64Image = Utilities.base64Encode(chartBlob.getBytes());
  const imageUrl = uploadToFreeImageHost(base64Image);

  if (!imageUrl) {
    Logger.log("Failed to upload image.");
    return;
  }

  const payload = {
    cardsV2: [
      {
        card: {
          header: {
            title: `[오전 보고]`,
            subtitle: "Zendesk 전체"
          },
          sections: [
            {
              widgets: [
                {
                  image: {
                    imageUrl: imageUrl,
                    altText: "Zendesk Daily Chart"
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  Logger.log(response.getContentText());
}

function uploadToFreeImageHost(base64Image) {
  const url = 'https://freeimage.host/api/1/upload';
  const payload = {
    key: '6d207e02198a847aa98d0a2a901485a5',
    action: 'upload',
    source: base64Image,
    format: 'json'
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result?.status_code === 200 ? result.image.display_url : null;
  } catch (e) {
    Logger.log("Exception during upload: " + e.message);
    return null;
  }
}
