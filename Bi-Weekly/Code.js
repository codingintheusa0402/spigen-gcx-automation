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

  Object.keys(replacements).forEach(function(key) {
    presentation.replaceAllText(key, replacements[key]);
  });

  updateDefectModelCharts(presentation, sheet, rowCount);

  refreshLinkedCharts(presentation);

  presentation.saveAndClose();
}

function updateDefectModelCharts(presentation, sheet, rowCount) {
  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const productCol = getColumnIndexByHeader(sheet, 'Product Name');
  const reasonCol = getColumnIndexByHeader(sheet, '인입사유');

  const categories = sheet.getRange(2, categoryCol, rowCount, 1).getDisplayValues().flat();
  const products = sheet.getRange(2, productCol, rowCount, 1).getDisplayValues().flat();
  const reasons = sheet.getRange(2, reasonCol, rowCount, 1).getDisplayValues().flat();

  const productMap = {};

  for (let i = 0; i < rowCount; i++) {
    const category = String(categories[i]).trim();
    const product = String(products[i]).trim();
    const reason = String(reasons[i]).trim();

    if (category !== '4. Product Issue') continue;
    if (!product || !reason) continue;

    if (!productMap[product]) {
      productMap[product] = {
        total: 0,
        reasons: {}
      };
    }

    productMap[product].total++;
    productMap[product].reasons[reason] = (productMap[product].reasons[reason] || 0) + 1;
  }

  const topProducts = Object.entries(productMap)
    .sort(function(a, b) {
      return b[1].total - a[1].total;
    })
    .slice(0, 3);

  removeOldAutoCharts(presentation);

  topProducts.forEach(function(item, index) {
    const rank = index + 1;
    const productName = item[0];
    const total = item[1].total;
    const reasonCounts = item[1].reasons;

    const topReasons = Object.entries(reasonCounts)
      .sort(function(a, b) {
        return b[1] - a[1];
      })
      .slice(0, 3);

    const topReasonTotal = topReasons.reduce(function(sum, reasonItem) {
      return sum + reasonItem[1];
    }, 0);

    const otherCount = total - topReasonTotal;

    const chartData = {
      productName: productName,
      total: total,
      reasons: topReasons,
      other: otherCount
    };

    insertChartAtPlaceholder(
      presentation,
      `{{Defect_Model_Chart_${rank}}}`,
      chartData,
      `AUTO_Defect_Model_Chart_${rank}`
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
  const rows = data.reasons.map(function(r) { return [r[0] || '기타', r[1]]; });
  if (data.other > 0) rows.push(['그 외', data.other]);

  const dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Reason')
    .addColumn(Charts.ColumnType.NUMBER, 'Count');
  rows.forEach(function(r) { dt.addRow(r); });

  const chart = Charts.newPieChart()
    .setDataTable(dt.build())
    .setTitle(data.productName + '  |  ' + data.total + '건')
    .setDimensions(620, 360)
    .setColors(['#d336f4', '#1554ff', '#19c7f3', '#8790b5'])
    .setOption('pieHole', 0.5)
    .setOption('backgroundColor', '#11162d')
    .setOption('titleTextStyle', { color: '#ffffff', fontSize: 16 })
    .setOption('legend', { textStyle: { color: '#c3c9e6' } })
    .build();

  return chart.getBlob().setName(title + '.png');
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

function escapeSvg(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}