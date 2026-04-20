/***** ========= CUSTOM MENU ========= *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SP-API Dashboard')
    .addItem('Setup Sheets',            'setupSheets')
    .addSeparator()
    .addItem('Refresh Marketplaces',    'refreshMarketplaces')
    .addItem('Refresh Orders (all)',    'refreshOrders')
    .addItem('Refresh Sales Metrics',   'refreshSalesMetrics')
    .addItem('Refresh Feedback',        'refreshFeedback')
    .addItem('Refresh Inventory',       'refreshInventory')
    .addToUi();
}

/***** ========= CONFIG HELPERS ========= *****/
function _cfg(key, fallback) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName('Config');
  if (!sheet) return fallback;
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1]).trim() || fallback;
  }
  return fallback;
}

function _today() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

function _writeSheet(name, rows) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  if (!rows.length) return sheet;
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  return sheet;
}

/***** ========= SETUP ========= *****/
function setupSheets() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config') || ss.insertSheet('Config', 0);
  cfg.clearContents();

  var cfgData = [
    ['Setting',            'Value',                'Notes'],
    ['MARKETPLACE_ID',     'A1F83G8C2ARO7P',       'Run =SPMARKETPLACES() on any cell to see all your active marketplace IDs'],
    ['START_DATE',         '2025-01-01',            'Default start date for Refresh functions (YYYY-MM-DD)'],
    ['END_DATE',           _today(),                'Default end date'],
    ['SALES_GRANULARITY',  'Day',                   'Day / Week / Month / Year / Total / Hour'],
    ['FULFILLMENT_NETWORK','All',                   'All / AFN (FBA only) / MFN (seller-fulfilled only)'],
    ['MAX_PAGES',          '10',                    'Max pagination pages for Refresh Orders (100 orders per page)']
  ];
  cfg.getRange(1, 1, cfgData.length, 3).setValues(cfgData);
  cfg.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
  cfg.setFrozenRows(1);
  cfg.setColumnWidth(1, 180);
  cfg.setColumnWidth(2, 160);
  cfg.setColumnWidth(3, 420);

  ['Marketplaces', 'Orders', 'Order Items', 'Sales Metrics', 'Feedback', 'Inventory'].forEach(function(n) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });

  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    '1. Edit the Config sheet settings.\n' +
    '2. Set your SP-API credentials in Extensions → Apps Script → Project Settings → Script Properties.\n' +
    '3. Use the SP-API Dashboard menu to refresh each data sheet.'
  );
}

/***** ========= REFRESH: MARKETPLACES ========= *****/
function refreshMarketplaces() {
  try {
    var rows = [['MarketplaceId', 'Name', 'Country', 'Currency', 'Language', 'Active', 'SuspendedListings']];
    var seen = {};
    ['EU', 'FE'].forEach(function(ep) {
      try {
        var res  = spapiFetchWithRetry('GET', '/sellers/v1/marketplaceParticipations', { endpoint: ep }, 3, 5000);
        (res.payload || []).forEach(function(p) {
          var mkt = p.marketplace || {}, part = p.participation || {};
          if (!mkt.id || seen[mkt.id]) return;
          seen[mkt.id] = true;
          rows.push([mkt.id, mkt.name || '', mkt.countryCode || '', mkt.defaultCurrencyCode || '',
                     mkt.defaultLanguageCode || '', part.isParticipating ? 'Yes' : 'No',
                     part.hasSuspendedListings ? 'Yes' : 'No']);
        });
      } catch(e) {}
    });
    _writeSheet('Marketplaces', rows);
    SpreadsheetApp.getUi().alert('Marketplaces refreshed: ' + (rows.length - 1) + ' marketplaces.');
  } catch(e) { SpreadsheetApp.getUi().alert('Error: ' + e.message); }
}

/***** ========= REFRESH: ORDERS ========= *****/
function refreshOrders() {
  var mktId     = _cfg('MARKETPLACE_ID',  'A1F83G8C2ARO7P');
  var startDate = _cfg('START_DATE',      '2025-01-01');
  var endDate   = _cfg('END_DATE',        _today());
  var maxPages  = parseInt(_cfg('MAX_PAGES', '10')) || 10;

  try {
    var baseQs = 'MarketplaceIds=' + encodeURIComponent(mktId) +
                 '&CreatedAfter='  + encodeURIComponent(_toIso(startDate, false)) +
                 '&CreatedBefore=' + encodeURIComponent(_toIso(endDate, true)) +
                 '&MaxResultsPerPage=100';

    var allOrders = [], nextToken = null, page = 0;
    do {
      var qs  = baseQs + (nextToken ? '&NextToken=' + encodeURIComponent(nextToken) : '');
      var res = spapiFetchWithRetry('GET', '/orders/v0/orders', { queryString: qs, endpoint: mktId }, 3, 5000);
      var pl  = res.payload || res;
      allOrders  = allOrders.concat(pl.Orders || []);
      nextToken  = pl.NextToken || null;
      if (nextToken) Utilities.sleep(1000);
      page++;
    } while (nextToken && page < maxPages);

    var rows = [['OrderId', 'PurchaseDate', 'Status', 'FulfillmentChannel', 'SalesChannel',
                 'ItemsShipped', 'ItemsUnshipped', 'Total', 'Currency', 'BuyerEmail(masked)']];
    allOrders.forEach(function(o) {
      rows.push([
        o.AmazonOrderId || '', o.PurchaseDate || '', o.OrderStatus || '',
        o.FulfillmentChannel || '', o.SalesChannel || '',
        o.NumberOfItemsShipped || 0, o.NumberOfItemsUnshipped || 0,
        (o.OrderTotal || {}).Amount || '', (o.OrderTotal || {}).CurrencyCode || '',
        (o.BuyerInfo  || {}).BuyerEmail || ''
      ]);
    });

    _writeSheet('Orders', rows);
    SpreadsheetApp.getUi().alert('Orders refreshed: ' + allOrders.length + ' orders' +
      (nextToken ? ' (more pages exist — increase MAX_PAGES in Config).' : '.'));
  } catch(e) { SpreadsheetApp.getUi().alert('Error: ' + e.message); }
}

/***** ========= REFRESH: SALES METRICS ========= *****/
function refreshSalesMetrics() {
  var mktId       = _cfg('MARKETPLACE_ID',     'A1F83G8C2ARO7P');
  var startDate   = _cfg('START_DATE',         '2025-01-01');
  var endDate     = _cfg('END_DATE',           _today());
  var granularity = _cfg('SALES_GRANULARITY',  'Day');
  var network     = _cfg('FULFILLMENT_NETWORK','All');

  try {
    var interval = _toIso(startDate, false) + '--' + _toIso(endDate, true);
    var qs = 'marketplaceIds='    + encodeURIComponent(mktId) +
             '&interval='         + encodeURIComponent(interval) +
             '&granularity='      + encodeURIComponent(granularity);
    if (network !== 'All') qs += '&fulfillmentNetwork=' + encodeURIComponent(network);

    var res     = spapiFetchWithRetry('GET', '/sales/v1/orderMetrics', { queryString: qs, endpoint: mktId }, 3, 5000);
    var metrics = res.payload || [];

    var rows = [['Interval', 'Units', 'OrderItems', 'Orders', 'AvgUnitPrice', 'TotalSales', 'Currency']];
    metrics.forEach(function(m) {
      rows.push([
        m.interval || '', m.unitCount || 0, m.orderItemCount || 0, m.orderCount || 0,
        (m.averageUnitPrice || {}).amount || '', (m.totalSales || {}).amount || '',
        (m.totalSales || {}).currencyCode || ''
      ]);
    });

    _writeSheet('Sales Metrics', rows);
    SpreadsheetApp.getUi().alert('Sales Metrics refreshed: ' + metrics.length + ' rows.');
  } catch(e) { SpreadsheetApp.getUi().alert('Error: ' + e.message); }
}

/***** ========= REFRESH: CUSTOMER FEEDBACK ========= *****/
function refreshFeedback() {
  var mktId     = _cfg('MARKETPLACE_ID', 'A1F83G8C2ARO7P');
  var startDate = _cfg('START_DATE',     '2025-01-01');
  var endDate   = _cfg('END_DATE',       _today());

  try {
    var baseQs = 'marketplaceId=' + encodeURIComponent(mktId) + '&count=100' +
                 '&createdAfter='  + encodeURIComponent(_toIso(startDate, false)) +
                 '&createdBefore=' + encodeURIComponent(_toIso(endDate, true));

    var all = [], pageToken = null, page = 0, maxPages = 20;
    do {
      var qs  = baseQs + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      var res = spapiFetchWithRetry('GET', '/customer-feedback/2024-06-01/feedbacks', { queryString: qs, endpoint: mktId }, 3, 5000);
      all       = all.concat(res.feedbacks || []);
      pageToken = res.nextPageToken || null;
      if (pageToken) Utilities.sleep(500);
      page++;
    } while (pageToken && page < maxPages);

    var rows = [['FeedbackId', 'OrderId', 'Date', 'Rating', 'Comments', 'SellerResponse']];
    all.forEach(function(f) {
      rows.push([f.feedbackId || '', f.orderId || '', f.createdTime || '',
                 f.rating || '', f.comments || '', (f.response || {}).text || '']);
    });

    _writeSheet('Feedback', rows);
    SpreadsheetApp.getUi().alert('Feedback refreshed: ' + all.length + ' entries.');
  } catch(e) { SpreadsheetApp.getUi().alert('Error: ' + e.message); }
}

/***** ========= REFRESH: INVENTORY ========= *****/
function refreshInventory() {
  var mktId = _cfg('MARKETPLACE_ID', 'A1F83G8C2ARO7P');

  try {
    var qs = 'granularityType=Marketplace' +
             '&granularityId='  + encodeURIComponent(mktId) +
             '&marketplaceIds=' + encodeURIComponent(mktId) +
             '&details=true';

    var all = [], nextToken = null, page = 0, maxPages = 20;
    do {
      var qsFull = qs + (nextToken ? '&nextToken=' + encodeURIComponent(nextToken) : '');
      var res    = spapiFetchWithRetry('GET', '/fba/inventory/v1/summaries', { queryString: qsFull, endpoint: mktId }, 3, 5000);
      var pl     = res.payload || res;
      all        = all.concat(pl.inventorySummaries || []);
      nextToken  = (pl.pagination || {}).nextToken || null;
      if (nextToken) Utilities.sleep(500);
      page++;
    } while (nextToken && page < maxPages);

    var rows = [['ASIN', 'FnSKU', 'SellerSKU', 'Condition', 'Available', 'Reserved', 'Unfulfillable', 'InboundWorking', 'Total']];
    all.forEach(function(s) {
      var det = s.inventoryDetails || {};
      var rsv = det.reservedQuantity      || {};
      var unf = det.unfulfillableQuantity || {};
      rows.push([
        s.asin || '', s.fnSku || '', s.sellerSku || '', s.condition || '',
        det.fulfillableQuantity || 0,
        rsv.totalReservedQuantity || 0,
        unf.totalUnfulfillableQuantity || 0,
        det.inboundWorkingQuantity || 0,
        s.totalQuantity || 0
      ]);
    });

    _writeSheet('Inventory', rows);
    SpreadsheetApp.getUi().alert('Inventory refreshed: ' + all.length + ' SKUs.');
  } catch(e) { SpreadsheetApp.getUi().alert('Error: ' + e.message); }
}

/***** ========= DATE HELPER (shared with formulas.js) ========= *****/
function _toIso(d, endOfDay) {
  var dt = (d instanceof Date) ? new Date(d.getTime()) : new Date(String(d).trim());
  if (isNaN(dt.getTime())) throw new Error('Invalid date: ' + d);
  endOfDay ? dt.setHours(23, 59, 59, 999) : dt.setHours(0, 0, 0, 0);
  return dt.toISOString();
}
