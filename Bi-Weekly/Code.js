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

  // {Defect_Model_}: group by Product Name → top-3 products, each showing top-3 reasons
  const topProducts = buildTopProductsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const p = topProducts[n - 1];
    replacements['{{Defect_Model_Chart_Title_' + n + '}}']        = p ? p.productName : '';
    replacements['{{Defect_Model_Chart_Count_' + n + '}}']        = p ? p.total.toLocaleString() + '건' : '';
    replacements['{{Defect_Model_Chart_Legend_' + n + '}}']       = p ? buildLegendText(p) : '';
    replacements['{{Defect_Model_Chart_Legend_Value_' + n + '}}'] = p ? buildLegendValues(p) : '';
  }

  // {Model_Defect_}: group by 인입사유 → top-3 reasons, each showing top-3 Product Names
  const topReasons = buildTopReasonsData(sheet, rowCount);
  for (let n = 1; n <= 3; n++) {
    const r = topReasons[n - 1];
    replacements['{{Model_Defect_Chart_Title_' + n + '}}']        = r ? r.reasonName : '';
    replacements['{{Model_Defect_Chart_Count_' + n + '}}']        = r ? r.total.toLocaleString() + '건' : '';
    replacements['{{Model_Defect_Chart_Legend_' + n + '}}']       = r ? buildModelLegendText(r) : '';
    replacements['{{Model_Defect_Chart_Legend_Value_' + n + '}}'] = r ? buildModelLegendValues(r) : '';
  }

  Object.keys(replacements).forEach(function(key) {
    presentation.replaceAllText(key, replacements[key]);
  });

  updateDefectModelCharts(presentation, topProducts);
  updateModelDefectCharts(presentation, topReasons);

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

// Extracts top-3 인입사유 from the sheet. Returns array of:
//   { reasonName, total, models: [[name, count], ...] (top 3 Product Names), other: remainderCount }
function buildTopReasonsData(sheet, rowCount) {
  const categoryCol = getColumnIndexByHeader(sheet, 'Category');
  const productCol  = getColumnIndexByHeader(sheet, 'Product Name');
  const reasonCol   = getColumnIndexByHeader(sheet, '인입사유');

  const categories = sheet.getRange(2, categoryCol, rowCount, 1).getDisplayValues().flat();
  const products   = sheet.getRange(2, productCol,  rowCount, 1).getDisplayValues().flat();
  const reasons    = sheet.getRange(2, reasonCol,   rowCount, 1).getDisplayValues().flat();

  const reasonMap = {};
  for (let i = 0; i < rowCount; i++) {
    const category = String(categories[i]).trim();
    const product  = String(products[i]).trim();
    const reason   = String(reasons[i]).trim();
    if (category !== '4. Product Issue' || !product || !reason) continue;
    if (!reasonMap[reason]) reasonMap[reason] = { total: 0, models: {} };
    reasonMap[reason].total++;
    reasonMap[reason].models[product] = (reasonMap[reason].models[product] || 0) + 1;
  }

  return Object.entries(reasonMap)
    .sort(function(a, b) { return b[1].total - a[1].total; })
    .slice(0, 3)
    .map(function(entry) {
      const reasonName = entry[0];
      const total      = entry[1].total;
      const topModels  = Object.entries(entry[1].models)
        .sort(function(a, b) { return b[1] - a[1]; })
        .slice(0, 3);
      const topTotal = topModels.reduce(function(s, m) { return s + m[1]; }, 0);
      return { reasonName: reasonName, total: total, models: topModels, other: total - topTotal };
    });
}

// Product names only, one per line (for {{Model_Defect_Chart_Legend_N}}).
function buildModelLegendText(item) {
  const lines = item.models.map(function(m) { return m[0]; });
  if (item.other > 0) lines.push('그 외');
  return lines.join('\n');
}

// Counts only, one per line — mirrors buildModelLegendText line-for-line.
function buildModelLegendValues(item) {
  const lines = item.models.map(function(m) { return String(m[1]); });
  if (item.other > 0) lines.push(String(item.other));
  return lines.join('\n');
}

// Inserts {{Model_Defect_Chart_N}} arc images (top-3 models per reason).
function updateModelDefectCharts(presentation, topReasons) {
  topReasons.forEach(function(r, index) {
    const rank = index + 1;
    insertChartAtPlaceholder(
      presentation,
      '{{Model_Defect_Chart_' + rank + '}}',
      { reasons: r.models, other: r.other },
      'AUTO_Model_Defect_Chart_' + rank
    );
  });
}

function updateDefectModelCharts(presentation, topProducts) {
  // Do NOT call removeOldAutoCharts() here — preservation logic lives inside
  // insertChartAtPlaceholder: if the {{}} text box is gone the image is kept as-is.
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

  // Placeholder already replaced by an image on a prior run → preserve it.
  if (anchors.length === 0) return;

  // Remove only this slot's previous auto-chart (if any) before inserting the new one.
  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(el) {
      if ((el.getTitle ? el.getTitle() : '') === title) el.remove();
    });
  });

  const anchor = anchors[0];
  const slide = anchor.slide;
  const shape = anchor.shape;

  const left = shape.getLeft();
  const top = shape.getTop();
  const width = shape.getWidth();
  const height = shape.getHeight();

  shape.getText().setText('');  // clear placeholder text → marks slot as "placed"

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

  // Canvas: 440×340 px.
  // Half-donut arc: 220×110 (2:1 semicircle), centered horizontally.
  //   Full circle chart area = 220×220.
  //   left = (440 - 220) / 2 = 110
  //   top  = 45  (pushed down a bit so the arc sits lower in the card)
  // Visible arc = top half of the 220px circle → y: 45 to 155 (110px tall).
  // Invisible spacer half → y: 155 to 265, still within canvas, background color.
  return Charts.newPieChart()
    .setDataTable(dt.build())
    .setDimensions(440, 340)
    .setColors(['#d336f4', '#1554ff', '#19c7f3', '#8790b5'])
    .setOption('pieHole', 0.9)
    .setOption('pieStartAngle', -90)
    .setOption('slices', spacerOpt)
    .setOption('pieSliceBorderColor', '#11162d')
    .setOption('backgroundColor', '#11162d')
    .setOption('chartArea', { left: 110, top: 45, width: 220, height: 220 })
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

