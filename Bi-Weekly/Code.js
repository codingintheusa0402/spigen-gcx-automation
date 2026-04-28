function onOpen() {
  SlidesApp.getUi()
    .createMenu('Slide Updater')
    .addItem('Update Slide Text', 'updateSlideTextBoxes')
    .addToUi();
}

function updateSlideTextBoxes() {
  const sheetId = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc';
  const sheetName = '26년 전체문의';

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const lastRow = sheet.getLastRow();
  const rowCount = Math.max(lastRow - 1, 0);

  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const reasonCol = getColumnIndexByHeader(sheet, '인입사유');

  const categoryValues = sheet.getRange(2, categoryCol, rowCount, 1).getDisplayValues().flat();
  const reasonValues = sheet.getRange(2, reasonCol, rowCount, 1).getDisplayValues().flat();

  const filteredReasons = [];

  for (let i = 0; i < rowCount; i++) {
    const category = String(categoryValues[i]).trim();
    const reason = String(reasonValues[i]).trim();

    if (category === '4. Product Issue' && reason) {
      filteredReasons.push(reason);
    }
  }

  const reasonCounts = {};

  filteredReasons.forEach(function(reason) {
    reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
  });

  const sortedReasons = Object.entries(reasonCounts).sort(function(a, b) {
    return b[1] - a[1];
  });

  const presentation = SlidesApp.getActivePresentation();

  const replacements = {
    '{{TOTAL_INQUIRIES}}': rowCount.toLocaleString()
  };

  for (let i = 1; i <= 5; i++) {
    const item = sortedReasons[i - 1];

    replacements[`{{Defect_Reason_${i}}}`] = item ? item[0] : '';
    replacements[`{{Defect_Reason_${i}_Count}}`] = item ? item[1].toLocaleString() : '';
  }

  const keywordPlaceholders = extractKeywordPlaceholders(presentation, 'Defect_Reason_');

  keywordPlaceholders.forEach(function(placeholder) {
    const keyword = placeholder
      .replace('{{Defect_Reason_', '')
      .replace('}}', '')
      .trim();

    if (/^\d+(_Count)?$/i.test(keyword)) return;

    const count = filteredReasons.filter(function(reason) {
      return reason.indexOf(keyword) !== -1;
    }).length;

    replacements[placeholder] = count.toLocaleString();
  });

  // Top-product placeholders (Title / Count / Legend) for each chart card
  const topProducts = buildTopProductsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const p = topProducts[n - 1];
    replacements['{{Defect_Model_Chart_Title_' + n + '}}']  = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_' + n + '}}']  = p ? p.total.toLocaleString() : '';
    replacements['{{Defect_Model_Chart_Legend_' + n + '}}'] = p ? buildLegendText(p) : '';
  }

  Object.keys(replacements).forEach(function(key) {
    presentation.replaceAllText(key, replacements[key]);
  });

  updateDefectModelCharts(presentation, topProducts);

  refreshLinkedCharts(presentation);

  presentation.saveAndClose();
}

// Extracts top-3 defect products from the sheet. Returns array of:
//   { productName, total, reasons: [[name, count], ...] (top 3), other: remainderCount }
function buildTopProductsData(sheet, rowCount) {
  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const productCol  = getColumnIndexByHeader(sheet, 'Product Name');
  const reasonCol   = getColumnIndexByHeader(sheet, '인입사유');

  const categories = sheet.getRange(2, categoryCol, rowCount, 1).getDisplayValues().flat();
  const products   = sheet.getRange(2, productCol,  rowCount, 1).getDisplayValues().flat();
  const reasons    = sheet.getRange(2, reasonCol,   rowCount, 1).getDisplayValues().flat();

  const productMap = {};
  for (let i = 0; i < rowCount; i++) {
    const category = String(categories[i]).trim();
    const product  = String(products[i]).trim();
    const reason   = String(reasons[i]).trim();
    if (category !== '4. Product Issue' || !product || !reason) continue;
    if (!productMap[product]) productMap[product] = { total: 0, reasons: {} };
    productMap[product].total++;
    productMap[product].reasons[reason] = (productMap[product].reasons[reason] || 0) + 1;
  }

  return Object.entries(productMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, 3)
    .map(function(entry) {
      const productName = entry[0];
      const total       = entry[1].total;
      const topReasons  = Object.entries(entry[1].reasons)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topReasons.reduce(function(s, r) { return s + r[1]; }, 0);
      return { productName: productName, total: total, reasons: topReasons, other: total - topTotal };
    });
}

// Formats legend text for one chart card.
// Each line: "<reason>\t<count>" — use a right-aligned tab stop in the Slides text box
// for the counts to stick to the right edge.
function buildLegendText(item) {
  const lines = item.reasons.map(function(r) { return r[0] + '\t' + r[1]; });
  if (item.other > 0) lines.push('그 외\t' + item.other);
  return lines.join('\n');
}

function updateDefectModelCharts(presentation, topProducts) {
  removeOldAutoCharts(presentation);

  topProducts.forEach(function(p, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(
      presentation,
      '{{Defect_Model_Chart_' + rank + '}}',
      p,
      'AUTO_Defect_Model_Chart_' + rank
    );
  });
}

function insertChartAtPlaceholder(presentation, placeholder, chartData, title) {
  const anchors = findPlaceholderShapes(presentation, placeholder);

  if (anchors.length === 0) return;

  const anchor = anchors[0];
  const slide = anchor.slide;
  const shape = anchor.shape;

  const left = shape.getLeft();
  const top = shape.getTop();
  const width = shape.getWidth();
  const height = shape.getHeight();

  shape.getText().setText('');

  const blob = buildDefectModelChartBlob(chartData, title);
  const image = slide.insertImage(blob, left, top, width, height);
  image.setTitle(title);
}

function buildDefectModelChartBlob(data, title) {
  const labels = data.reasons.map(function(r) { return r[0] || '기타'; });
  const values = data.reasons.map(function(r) { return r[1]; });
  if (data.other > 0) {
    labels.push('그 외');
    values.push(data.other);
  }

  // Arc-only chart — no text, no legend (handled by separate {{}} placeholders on the slide).
  const config = {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: ['#d336f4', '#1554ff', '#19c7f3', '#8790b5'],
        borderWidth: 0
      }]
    },
    options: {
      rotation: 270,      // 9 o'clock → sweeps clockwise through 12 → 3 = ∩
      circumference: 180,
      cutout: '90%',
      layout: { padding: 4 },
      plugins: {
        legend: { display: false },
        datalabels: { display: false }
      }
    }
  };

  const payload = JSON.stringify({
    version: '3',
    backgroundColor: '#11162d',
    width: 500,
    height: 400,
    chart: config
  });

  const response = UrlFetchApp.fetch('https://quickchart.io/chart', {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('QuickChart ' + response.getResponseCode() + ': ' + response.getContentText().slice(0, 300));
    return buildChartsFallback_(data, title);
  }

  return response.getBlob().setName(title + '.png');
}

// Fallback used when QuickChart is unreachable — full donut via built-in Charts service.
function buildChartsFallback_(data, title) {
  const rows = data.reasons.map(function(r) { return [r[0] || '기타', r[1]]; });
  if (data.other > 0) rows.push(['그 외', data.other]);

  const dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Reason')
    .addColumn(Charts.ColumnType.NUMBER, 'Count');
  rows.forEach(function(r) { dt.addRow(r); });

  return Charts.newPieChart()
    .setDataTable(dt.build())
    .setTitle(data.productName + '  |  ' + data.total + '건')
    .setDimensions(500, 400)
    .setColors(['#d336f4', '#1554ff', '#19c7f3', '#8790b5'])
    .setOption('pieHole', 0.65)
    .setOption('pieSliceText', 'value')
    .setOption('backgroundColor', '#11162d')
    .setOption('titleTextStyle', { color: '#ffffff', fontSize: 14 })
    .setOption('legend', { textStyle: { color: '#c3c9e6' } })
    .build()
    .getBlob()
    .setName(title + '.png');
}

function refreshLinkedCharts(presentation) {
  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHEETS_CHART) {
        element.asSheetsChart().refresh();
      }
    });
  });
}

function findPlaceholderShapes(presentation, placeholder) {
  const results = [];

  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

      const shape = element.asShape();
      const text = shape.getText().asString();

      if (text.indexOf(placeholder) !== -1) {
        results.push({
          slide: slide,
          shape: shape
        });
      }
    });
  });

  return results;
}

function removeOldAutoCharts(presentation) {
  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      const title = element.getTitle ? element.getTitle() : '';

      if (title && title.indexOf('AUTO_Defect_Model_Chart_') === 0) {
        element.remove();
      }
    });
  });
}

function getColumnIndexByHeader(sheet, headerName) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(function(header) {
      return String(header).trim();
    });

  const index = headers.indexOf(headerName);

  if (index === -1) {
    throw new Error('Header not found: ' + headerName);
  }

  return index + 1;
}

function extractKeywordPlaceholders(presentation, prefix) {
  const placeholders = new Set();
  const pattern = new RegExp('{{' + prefix + '[^}]+}}', 'g');

  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

      const shape = element.asShape();
      const text = shape.getText().asString();
      const matches = text.match(pattern);

      if (matches) {
        matches.forEach(function(match) {
          placeholders.add(match);
        });
      }
    });
  });

  return Array.from(placeholders);
}

