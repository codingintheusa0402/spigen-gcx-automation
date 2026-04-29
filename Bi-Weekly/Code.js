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

  // Top-product placeholders (Title / Count / Legend / Legend_Value) for each chart card
  const topProducts = buildTopProductsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const p = topProducts[n - 1];
    replacements['{{Defect_Model_Chart_Title_' + n + '}}']        = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_' + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{Defect_Model_Chart_Legend_' + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{Defect_Model_Chart_Legend_Value_' + n + '}}'] = p ? buildLegendValues(p) : '';
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

// Returns reason names only, one per line (pair with buildLegendValues for left/right text boxes).
function buildLegendText(item) {
  const lines = item.reasons.map(function(r) { return r[0]; });
  if (item.other > 0) lines.push('그 외');
  return lines.join('\n');
}

// Returns counts only, one per line — mirrors buildLegendText line-for-line.
function buildLegendValues(item) {
  const lines = item.reasons.map(function(r) { return String(r[1]); });
  if (item.other > 0) lines.push(String(item.other));
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

  // Half-donut trick (no external API needed):
  // Add a spacer slice equal to the total of all visible slices.
  // Total becomes 2× visible sum → each half = 180° exactly.
  // The spacer is colored to match the background, making it invisible.
  // pieStartAngle: -90 positions visible data at 9 o'clock → sweeps through
  // 12 o'clock to 3 o'clock = ∩ upward arch.
  const visibleSum = values.reduce(function(a, b) { return a + b; }, 0);
  labels.push('');        // no label for spacer
  values.push(visibleSum);

  const dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Label')
    .addColumn(Charts.ColumnType.NUMBER, 'Value');
  for (let i = 0; i < labels.length; i++) {
    dt.addRow([labels[i], values[i]]);
  }

  const spacerOpt = {};
  spacerOpt[labels.length - 1] = { color: '#11162d' };  // spacer = background color

  // Canvas: 297×228 logical → rendered at 3× (891×684) for high resolution.
  // Half-donut arc: 150×75 logical → 450×225 at 3×.
  // Full circle chart area = 450×450 (diameter), centered horizontally in 891px canvas.
  //   left  = (891 - 450) / 2 = 220.5 → 221
  //   top   = 15  (small top gap)
  //   width = 450, height = 450
  // Visible arc (top 225px of the 450px circle) sits in the top ~33% of the canvas,
  // matching the 150×75 proportion. The invisible spacer half falls below y=240,
  // still within the 684px canvas but colored background so it disappears.
  return Charts.newPieChart()
    .setDataTable(dt.build())
    .setDimensions(891, 684)
    .setColors(['#d336f4', '#1554ff', '#19c7f3', '#8790b5'])
    .setOption('pieHole', 0.9)
    .setOption('pieStartAngle', -90)
    .setOption('slices', spacerOpt)
    .setOption('pieSliceBorderColor', '#11162d')
    .setOption('backgroundColor', '#11162d')
    .setOption('chartArea', { left: 221, top: 15, width: 450, height: 450 })
    .setOption('pieSliceText', 'none')
    .setOption('legend', { position: 'none' })
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

