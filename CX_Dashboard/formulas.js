/***** ========= FORMULA HELPERS ========= *****/

// Accepts Date object (from cell reference) or "YYYY-MM-DD" string.
function _toIso(d, endOfDay) {
  var dt = (d instanceof Date) ? new Date(d.getTime()) : new Date(String(d).trim());
  if (isNaN(dt.getTime())) throw new Error('Invalid date: ' + d);
  endOfDay ? dt.setHours(23, 59, 59, 999) : dt.setHours(0, 0, 0, 0);
  return dt.toISOString();
}

// Safely cache a 2D array. Silently skips if value exceeds CacheService 100KB limit.
function _cacheSet(key, rows, ttl) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(rows), ttl); } catch(e) {}
}

function _cacheGet(key) {
  var v = CacheService.getScriptCache().get(key);
  if (!v) return null;
  try { return JSON.parse(v); } catch(e) { return null; }
}

/***** ========= MARKETPLACE REFERENCE ========= *****/

/**
 * Returns all Amazon marketplaces this selling account participates in.
 * Use the MarketplaceId column as input for all other SP formulas.
 * @customfunction
 * @return {Array} MarketplaceId | Name | Country | Currency | Language | Active | SuspendedListings
 */
function SPMARKETPLACES() {
  var cached = _cacheGet('SPMKT');
  if (cached) return cached;

  var rows = [['MarketplaceId', 'Name', 'Country', 'Currency', 'Language', 'Active', 'SuspendedListings']];
  var seen = {};

  ['EU', 'FE'].forEach(function(ep) {
    try {
      var res  = spapiFetchWithRetry('GET', '/sellers/v1/marketplaceParticipations', { endpoint: ep }, 3, 5000);
      var list = res.payload || [];
      list.forEach(function(p) {
        var mkt  = p.marketplace   || {};
        var part = p.participation || {};
        if (!mkt.id || seen[mkt.id]) return;
        seen[mkt.id] = true;
        rows.push([mkt.id, mkt.name || '', mkt.countryCode || '', mkt.defaultCurrencyCode || '',
                   mkt.defaultLanguageCode || '', part.isParticipating ? 'Yes' : 'No',
                   part.hasSuspendedListings ? 'Yes' : 'No']);
      });
    } catch(e) {}
  });

  _cacheSet('SPMKT', rows, 3600);
  return rows.length > 1 ? rows : [['No data — check SP-API credentials in Script Properties']];
}

/***** ========= ORDERS ========= *****/

/**
 * Returns Amazon orders for a marketplace and date range (max 100 rows — use Refresh Orders menu for full export).
 * @customfunction
 * @param {string} marketplaceId Marketplace ID. Run =SPMARKETPLACES() to see your active IDs.
 * @param {string} startDate Start date (YYYY-MM-DD or cell reference)
 * @param {string} endDate End date (YYYY-MM-DD or cell reference)
 * @param {string} [status] Optional: filter by status — Pending, Unshipped, PartiallyShipped, Shipped, Canceled, Unfulfillable
 * @return {Array} OrderId | PurchaseDate | Status | FulfillmentChannel | SalesChannel | ItemsShipped | ItemsUnshipped | Total | Currency | BuyerEmail(masked)
 */
function SPORDERS(marketplaceId, startDate, endDate, status) {
  if (!marketplaceId || !startDate || !endDate)
    return [['Required: marketplaceId, startDate, endDate']];

  var cKey = 'SPORDERS_' + [marketplaceId, startDate, endDate, status || ''].join('_');
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  try {
    var qs = 'MarketplaceIds=' + encodeURIComponent(marketplaceId) +
             '&CreatedAfter='  + encodeURIComponent(_toIso(startDate, false)) +
             '&CreatedBefore=' + encodeURIComponent(_toIso(endDate, true)) +
             '&MaxResultsPerPage=100';
    if (status) qs += '&OrderStatuses=' + encodeURIComponent(status);

    var res    = spapiFetchWithRetry('GET', '/orders/v0/orders', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
    var orders = (res.payload || res).Orders || [];

    var rows = [['OrderId', 'PurchaseDate', 'Status', 'FulfillmentChannel', 'SalesChannel',
                 'ItemsShipped', 'ItemsUnshipped', 'Total', 'Currency', 'BuyerEmail(masked)']];
    orders.forEach(function(o) {
      rows.push([
        o.AmazonOrderId || '',
        o.PurchaseDate  || '',
        o.OrderStatus   || '',
        o.FulfillmentChannel || '',
        o.SalesChannel  || '',
        o.NumberOfItemsShipped   || 0,
        o.NumberOfItemsUnshipped || 0,
        (o.OrderTotal || {}).Amount       || '',
        (o.OrderTotal || {}).CurrencyCode || '',
        (o.BuyerInfo  || {}).BuyerEmail   || ''
      ]);
    });

    _cacheSet(cKey, rows, 300);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}

/***** ========= ORDER ITEMS ========= *****/

/**
 * Returns the line items for a specific Amazon order.
 * @customfunction
 * @param {string} orderId Amazon order ID (e.g. 026-1234567-1234567 or MCF sellerFulfillmentOrderId)
 * @param {string} [endpoint] Endpoint override: "EU" or "FE". Auto-detected if omitted.
 * @return {Array} ASIN | SellerSKU | Title | QtyOrdered | QtyShipped | ItemPrice | ItemTax | Currency | Condition | PromotionDiscount
 */
function SPORDERITEMS(orderId, endpoint) {
  if (!orderId) return [['orderId is required']];

  var cKey = 'SPOI_' + orderId;
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  var endpoints = endpoint ? [endpoint] : ['EU', 'FE'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var res   = spapiFetchWithRetry('GET', '/orders/v0/orders/' + encodeURIComponent(orderId) + '/orderItems',
                    { endpoint: endpoints[i] }, 3, 5000);
      var items = (res.payload || res).OrderItems || [];

      var rows = [['ASIN', 'SellerSKU', 'Title', 'QtyOrdered', 'QtyShipped',
                   'ItemPrice', 'ItemTax', 'Currency', 'Condition', 'PromotionDiscount']];
      items.forEach(function(item) {
        rows.push([
          item.ASIN || '', item.SellerSKU || '', item.Title || '',
          item.QuantityOrdered || 0, item.QuantityShipped || 0,
          (item.ItemPrice          || {}).Amount || '',
          (item.ItemTax            || {}).Amount || '',
          (item.ItemPrice          || {}).CurrencyCode || '',
          item.ConditionId || '',
          (item.PromotionDiscount  || {}).Amount || ''
        ]);
      });

      _cacheSet(cKey, rows, 600);
      return rows;
    } catch(err) {
      lastErr = err;
      var m = err.message || '';
      if (m.indexOf('SP-API error 400') >= 0 || m.indexOf('SP-API error 404') >= 0) continue;
      break;
    }
  }
  return [['ERR: ' + (lastErr ? (lastErr.message || lastErr) : 'Not found')]];
}

/***** ========= SALES METRICS ========= *****/

/**
 * Returns aggregate sales metrics from the Sales API.
 * @customfunction
 * @param {string} marketplaceId Marketplace ID
 * @param {string} startDate Start date (YYYY-MM-DD)
 * @param {string} endDate End date (YYYY-MM-DD)
 * @param {string} [granularity] Day (default) | Week | Month | Year | Total | Hour
 * @param {string} [fulfillmentNetwork] All (default) | AFN (FBA only) | MFN (seller-fulfilled only)
 * @return {Array} Interval | Units | OrderItems | Orders | AvgUnitPrice | TotalSales | Currency
 */
function SPSALES(marketplaceId, startDate, endDate, granularity, fulfillmentNetwork) {
  if (!marketplaceId || !startDate || !endDate)
    return [['Required: marketplaceId, startDate, endDate']];

  granularity       = granularity       || 'Day';
  fulfillmentNetwork = fulfillmentNetwork || 'All';

  var cKey = 'SPSALES_' + [marketplaceId, startDate, endDate, granularity, fulfillmentNetwork].join('_');
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  try {
    var interval = _toIso(startDate, false) + '--' + _toIso(endDate, true);
    var qs = 'marketplaceIds='      + encodeURIComponent(marketplaceId) +
             '&interval='           + encodeURIComponent(interval) +
             '&granularity='        + encodeURIComponent(granularity);
    if (fulfillmentNetwork !== 'All')
      qs += '&fulfillmentNetwork=' + encodeURIComponent(fulfillmentNetwork);

    var res     = spapiFetchWithRetry('GET', '/sales/v1/orderMetrics', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
    var metrics = res.payload || [];

    var rows = [['Interval', 'Units', 'OrderItems', 'Orders', 'AvgUnitPrice', 'TotalSales', 'Currency']];
    metrics.forEach(function(m) {
      rows.push([
        m.interval || '',
        m.unitCount       || 0,
        m.orderItemCount  || 0,
        m.orderCount      || 0,
        (m.averageUnitPrice || {}).amount       || '',
        (m.totalSales       || {}).amount       || '',
        (m.totalSales       || {}).currencyCode || ''
      ]);
    });

    _cacheSet(cKey, rows, 600);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}

/***** ========= CUSTOMER FEEDBACK ========= *****/

/**
 * Returns customer feedback (seller ratings) from the Customer Feedback API.
 * Requires the Buyer Communication SP-API role.
 * @customfunction
 * @param {string} marketplaceId Marketplace ID
 * @param {string} [startDate] Start date YYYY-MM-DD (optional)
 * @param {string} [endDate] End date YYYY-MM-DD (optional)
 * @return {Array} FeedbackId | OrderId | Date | Rating | Comments | SellerResponse
 */
function SPFEEDBACK(marketplaceId, startDate, endDate) {
  if (!marketplaceId) return [['marketplaceId is required']];

  var cKey = 'SPFB_' + [marketplaceId, startDate || '', endDate || ''].join('_');
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  try {
    var qs = 'marketplaceId=' + encodeURIComponent(marketplaceId) + '&count=100';
    if (startDate) qs += '&createdAfter='  + encodeURIComponent(_toIso(startDate, false));
    if (endDate)   qs += '&createdBefore=' + encodeURIComponent(_toIso(endDate, true));

    var res       = spapiFetchWithRetry('GET', '/customer-feedback/2024-06-01/feedbacks', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
    var feedbacks = res.feedbacks || [];

    var rows = [['FeedbackId', 'OrderId', 'Date', 'Rating', 'Comments', 'SellerResponse']];
    feedbacks.forEach(function(f) {
      rows.push([
        f.feedbackId   || '',
        f.orderId      || '',
        f.createdTime  || '',
        f.rating       || '',
        f.comments     || '',
        (f.response    || {}).text || ''
      ]);
    });

    _cacheSet(cKey, rows, 300);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}

/***** ========= FBA INVENTORY ========= *****/

/**
 * Returns FBA inventory summary for a marketplace.
 * @customfunction
 * @param {string} marketplaceId Marketplace ID
 * @return {Array} ASIN | FnSKU | SellerSKU | Condition | Available | Reserved | Unfulfillable | InboundWorking | Total
 */
function SPINVENTORY(marketplaceId) {
  if (!marketplaceId) return [['marketplaceId is required']];

  var cKey = 'SPINV_' + marketplaceId;
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  try {
    var qs = 'granularityType=Marketplace' +
             '&granularityId='   + encodeURIComponent(marketplaceId) +
             '&marketplaceIds='  + encodeURIComponent(marketplaceId) +
             '&details=true';

    var res        = spapiFetchWithRetry('GET', '/fba/inventory/v1/summaries', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
    var summaries  = (res.payload || res).inventorySummaries || [];

    var rows = [['ASIN', 'FnSKU', 'SellerSKU', 'Condition', 'Available', 'Reserved', 'Unfulfillable', 'InboundWorking', 'Total']];
    summaries.forEach(function(s) {
      var det = s.inventoryDetails || {};
      var res = det.reservedQuantity    || {};
      var unf = det.unfulfillableQuantity || {};
      rows.push([
        s.asin || '', s.fnSku || '', s.sellerSku || '', s.condition || '',
        det.fulfillableQuantity   || 0,
        res.totalReservedQuantity || 0,
        unf.totalUnfulfillableQuantity || 0,
        det.inboundWorkingQuantity || 0,
        s.totalQuantity || 0
      ]);
    });

    _cacheSet(cKey, rows, 600);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}

/***** ========= FBA INVENTORY BY ASIN ========= *****/

/**
 * Returns FBA inventory aggregated by ASIN (sums across all SKUs per ASIN).
 * @customfunction
 * @param {string} marketplaceId Marketplace ID (e.g. A1PA6795UKMFR9)
 * @param {string|Array} [asins] Optional: single ASIN, comma-separated string, or a vertical/horizontal cell range of ASINs. Omit to return all.
 * @param {boolean} [includeZero] Pass TRUE to include ASINs with zero stock (default: FALSE)
 * @return {Array} ASIN | Available | Reserved | Unfulfillable | InboundWorking | Total | SKUs
 */
function SPINVENTORY_ASIN(marketplaceId, asins, includeZero) {
  if (!marketplaceId) return [['marketplaceId is required']];

  // Normalise asins → a Set-like object (keys = uppercase ASIN), or null for "all"
  var asinSet = null;
  if (asins) {
    var flat = [];
    if (Array.isArray(asins)) {
      var arr2d = Array.isArray(asins[0]) ? asins : asins.map(function(v) { return [v]; });
      arr2d.forEach(function(row) {
        row.forEach(function(v) { if (v) flat.push(String(v).trim().toUpperCase()); });
      });
    } else if (String(asins).trim()) {
      flat = String(asins).split(/[\s,]+/).filter(Boolean).map(function(v) { return v.toUpperCase(); });
    }
    if (flat.length) {
      asinSet = {};
      flat.forEach(function(a) { asinSet[a] = true; });
    }
  }

  // Build a short cache fingerprint that won't exceed CacheService key limits
  var asinFp = asinSet
    ? (function() {
        var k = Object.keys(asinSet).sort();
        return k.length + '_' + k[0] + '_' + k[k.length - 1];
      })()
    : 'ALL';
  var cKey = 'SPINV_v4_' + marketplaceId + '_' + (includeZero ? '1' : '0') + '_' + asinFp;
  var cached = _cacheGet(cKey);
  if (cached) return cached;

  try {
    // Paginate through all inventory summaries
    var all = [], nextToken = null, maxPages = 20, page = 0;
    do {
      var qs = 'granularityType=Marketplace' +
               '&granularityId='  + encodeURIComponent(marketplaceId) +
               '&marketplaceIds=' + encodeURIComponent(marketplaceId) +
               '&details=true';
      if (nextToken) qs += '&nextToken=' + encodeURIComponent(nextToken);

      var res = spapiFetchWithRetry('GET', '/fba/inventory/v1/summaries', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
      var pl  = res.payload || res;
      all     = all.concat(pl.inventorySummaries || []);
      nextToken = (pl.pagination || {}).nextToken || null;
      if (nextToken) Utilities.sleep(300);
      page++;
    } while (nextToken && page < maxPages);

    // Aggregate per ASIN
    var map = {};
    all.forEach(function(s) {
      var asin = (s.asin || '').toUpperCase();
      if (!asin) return;
      if (asinSet && !asinSet[asin]) return; // skip if not in requested list

      var det  = s.inventoryDetails        || {};
      var rsv  = det.reservedQuantity      || {};
      var unf  = det.unfulfillableQuantity || {};

      var avail = (typeof det.fulfillableQuantity === 'object')
        ? (det.fulfillableQuantity || {}).quantity || 0
        : det.fulfillableQuantity || 0;

      var inbound = (typeof det.inboundWorkingQuantity === 'object')
        ? (det.inboundWorkingQuantity || {}).quantity || 0
        : det.inboundWorkingQuantity || 0;

      if (!map[asin]) map[asin] = { available: 0, reserved: 0, unfulfillable: 0, inbound: 0, total: 0, skus: 0 };
      var a = map[asin];
      a.available     += avail;
      a.reserved      += rsv.totalReservedQuantity      || 0;
      a.unfulfillable += unf.totalUnfulfillableQuantity || 0;
      a.inbound       += inbound;
      a.total         += s.totalQuantity                || 0;
      a.skus          += 1;
    });

    var rows = [['ASIN', 'Available', 'Reserved', 'Unfulfillable', 'InboundWorking', 'Total', 'SKUs']];

    // Preserve input order when specific ASINs were requested; sort otherwise
    var outputKeys = asinSet ? Object.keys(asinSet).sort() : Object.keys(map).sort();
    outputKeys.forEach(function(asin) {
      var a = map[asin];
      if (!a) {
        // ASIN was requested but not found in inventory at all
        if (asinSet) rows.push([asin, 'NOT FOUND', '', '', '', '', '']);
        return;
      }
      if (!includeZero && a.total === 0) return;
      rows.push([asin, a.available, a.reserved, a.unfulfillable, a.inbound, a.total, a.skus]);
    });

    if (rows.length === 1) rows.push(['No inventory found', '', '', '', '', '', '']);

    _cacheSet(cKey, rows, 1800);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}

/***** ========= FBA INVENTORY DEBUG ========= *****/

/**
 * Returns first-page raw FBA inventory from the API for diagnostic purposes.
 * Use this to confirm the API is returning data and inspect actual field shapes.
 * @customfunction
 * @param {string} marketplaceId Marketplace ID
 * @return {Array} Raw first-page inventory + pagination info
 */
function SPINVENTORY_DEBUG(marketplaceId) {
  if (!marketplaceId) return [['marketplaceId is required']];
  try {
    var qs = 'granularityType=Marketplace' +
             '&granularityId='  + encodeURIComponent(marketplaceId) +
             '&marketplaceIds=' + encodeURIComponent(marketplaceId) +
             '&details=true';
    var res   = spapiFetchWithRetry('GET', '/fba/inventory/v1/summaries', { queryString: qs, endpoint: marketplaceId }, 3, 5000);
    var pl    = res.payload || res;
    var items = pl.inventorySummaries || [];
    var hasNext = !!((pl.pagination || {}).nextToken);

    var rows = [['#', 'ASIN', 'SellerSKU', 'FnSKU', 'Condition', 'TotalQty', 'fulfillableQty_raw', 'inboundWorking_raw']];
    items.slice(0, 50).forEach(function(s, i) {
      var det = s.inventoryDetails || {};
      rows.push([
        i + 1,
        s.asin || '',
        s.sellerSku || '',
        s.fnSku || '',
        s.condition || '',
        s.totalQuantity != null ? s.totalQuantity : '',
        JSON.stringify(det.fulfillableQuantity != null ? det.fulfillableQuantity : null),
        JSON.stringify(det.inboundWorkingQuantity != null ? det.inboundWorkingQuantity : null)
      ]);
    });
    rows.push(['---', 'Items on page: ' + items.length, 'More pages: ' + (hasNext ? 'YES' : 'NO'), '', '', '', '', '']);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
}
