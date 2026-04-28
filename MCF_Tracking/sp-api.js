/***** ========= CONFIG + AUTH HELPERS ========= *****/
function _prop(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v == null || v === '') ? fallback : v;
}

function _nowIsoBasic() {
  var d = new Date();
  var y = d.getUTCFullYear();
  var m = ('0' + (d.getUTCMonth() + 1)).slice(-2);
  var day = ('0' + d.getUTCDate()).slice(-2);
  var hh = ('0' + d.getUTCHours()).slice(-2);
  var mm = ('0' + d.getUTCMinutes()).slice(-2);
  var ss = ('0' + d.getUTCSeconds()).slice(-2);
  return { amzDate: y + m + day + 'T' + hh + mm + ss + 'Z', shortDate: '' + y + m + day };
}

function _toHex(bytes) {
  return bytes.map(function (b) {
    var s = (b & 0xff).toString(16);
    return s.length === 1 ? '0' + s : s;
  }).join('');
}

function _sha256Hex(msg) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, msg, Utilities.Charset.UTF_8);
  return _toHex(digest);
}

// If you assumeRole to get temporary creds, set AWS_SESSION_TOKEN in Script Properties.
function _sessionToken() {
  return _prop('AWS_SESSION_TOKEN', '');
}

/***** ========= LWA (Login With Amazon) ========= *****/
function _resolveLwaProfileKey(endpointKey) {
  // Return 'JP' when endpoint implies FE/JP; otherwise 'EU'.
  var k = (endpointKey || '').toString().toUpperCase();
  var feMkt = ['A1VC38T7YXB528', 'A39IBJ37TRP1C6', 'A19VAU5U5O7RUS']; // JP, AU, SG marketplaceIds
  if (k === 'JP' || k === 'FE' || feMkt.indexOf(k) >= 0) return 'JP';
  return 'EU';
}

function getLwaAccessToken(endpointKey) {
  var prof = _resolveLwaProfileKey(endpointKey);

  // Cache the token for 55 min (LWA tokens expire in 60 min).
  // Concurrent formula cells reuse the same token instead of each fetching a new one.
  var cacheKey = 'LWA_TOKEN_' + prof;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return cached;

  var clientId, clientSecret, refreshToken;
  if (prof === 'JP') {
    clientId     = _prop('LWA_CLIENT_ID_JP',     _prop('LWA_CLIENT_ID'));
    clientSecret = _prop('LWA_CLIENT_SECRET_JP', _prop('LWA_CLIENT_SECRET'));
    refreshToken = _prop('LWA_REFRESH_TOKEN_JP', _prop('LWA_REFRESH_TOKEN'));
  } else {
    clientId     = _prop('LWA_CLIENT_ID');
    clientSecret = _prop('LWA_CLIENT_SECRET');
    refreshToken = _prop('LWA_REFRESH_TOKEN');
  }
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Missing LWA credentials for profile ' + prof);
  }

  var resp = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret
    },
    muteHttpExceptions: true
  });

  var text = resp.getContentText() || '{}';
  var body = JSON.parse(text);
  if (resp.getResponseCode() >= 300 || !body.access_token) {
    throw new Error('LWA token fetch failed: ' + resp.getResponseCode() + ' ' + text);
  }

  cache.put(cacheKey, body.access_token, 3300); // 55 min TTL
  return body.access_token;
}

/***** ========= AWS SigV4 ========= *****/
function signSpApiRequest(method, host, path, queryString, body, region) {
  var accessKey = _prop('AWS_ACCESS_KEY_ID');
  var secretKey = _prop('AWS_SECRET_ACCESS_KEY');
  if (!accessKey || !secretKey) throw new Error('Missing AWS keys in Script properties.');

  var service = 'execute-api';
  var ts = _nowIsoBasic();
  var amzDate = ts.amzDate, shortDate = ts.shortDate;
  var sessionToken = _sessionToken(); // may be empty

  var canonicalUri = path;
  var canonicalQueryString = queryString || '';
  var payload = body || '';
  var payloadHash = _sha256Hex(payload);

  var canonicalHeaders =
    'host:' + host + '\n' +
    'x-amz-date:' + amzDate + '\n' +
    (sessionToken ? ('x-amz-security-token:' + sessionToken + '\n') : '');

  var signedHeaders = sessionToken ? 'host;x-amz-date;x-amz-security-token' : 'host;x-amz-date';

  var canonicalRequest = [
    method,
    canonicalUri,
    canonicalQueryString,
    canonicalHeaders,
    signedHeaders,
    payloadHash
  ].join('\n');

  var algorithm = 'AWS4-HMAC-SHA256';
  var credentialScope = shortDate + '/' + region + '/' + service + '/aws4_request';
  var stringToSign = [
    algorithm,
    amzDate,
    credentialScope,
    _sha256Hex(canonicalRequest)
  ].join('\n');

  var enc = function (s) { return Utilities.newBlob(s).getBytes(); };
  var kSecret  = enc('AWS4' + secretKey);
  var kDate    = Utilities.computeHmacSha256Signature(enc(shortDate), kSecret);
  var kRegion  = Utilities.computeHmacSha256Signature(enc(region),    kDate);
  var kService = Utilities.computeHmacSha256Signature(enc(service),   kRegion);
  var kSigning = Utilities.computeHmacSha256Signature(enc('aws4_request'), kService);
  var sigBytes = Utilities.computeHmacSha256Signature(enc(stringToSign), kSigning);
  var signature = _toHex(sigBytes);

  var authorizationHeader = algorithm + ' Credential=' + accessKey + '/' + credentialScope +
    ', SignedHeaders=' + signedHeaders + ', Signature=' + signature;

  return { amzDate: amzDate, authorizationHeader: authorizationHeader, sessionToken: sessionToken };
}

/***** ========= ENDPOINT RESOLVER (EU / FE) ========= *****/
function _getEndpoint(groupOrMarketplace) {
  var g = String(groupOrMarketplace || '').toUpperCase();
  var feMkt = ['A1VC38T7YXB528', 'A39IBJ37TRP1C6', 'A19VAU5U5O7RUS']; // JP, AU, SG marketplaceIds
  if (feMkt.indexOf(g) >= 0) g = 'FE';
  if (['JP', 'AU', 'SG'].indexOf(g) >= 0) g = 'FE';
  if (!g) g = 'EU';

  if (g === 'EU') {
    return {
      host: _prop('SPAPI_HOST_EU', _prop('SPAPI_HOST', 'sellingpartnerapi-eu.amazon.com')),
      region: _prop('SPAPI_REGION_EU', _prop('SPAPI_REGION', 'eu-west-1')),
      group: 'EU'
    };
  }
  // FE (Japan/AU/SG)
  return {
    host: _prop('SPAPI_HOST_FE', 'sellingpartnerapi-fe.amazon.com'),
    region: _prop('SPAPI_REGION_FE', 'us-west-2'),
    group: 'FE'
  };
}

/***** ========= CORE FETCH ========= *****/
function spapiFetch(method, path, opts) {
  opts = opts || {};
  var ep = _getEndpoint(opts.endpoint);
  var host = ep.host;
  var region = ep.region;

  var queryString = opts.queryString || '';
  var body = opts.body || '';

  var token = getLwaAccessToken(opts.endpoint);

  var sig = signSpApiRequest(method, host, path, queryString, body, region);
  var url = 'https://' + host + path + (queryString ? ('?' + queryString) : '');

  var headers = {
    'x-amz-date': sig.amzDate,
    'x-amz-access-token': token,
    'Authorization': sig.authorizationHeader,
    'Content-Type': 'application/json'
  };
  if (sig.sessionToken) headers['x-amz-security-token'] = sig.sessionToken;

  var fetchOpts = {
    method: method,
    headers: headers,
    muteHttpExceptions: true
  };
  if (method !== 'GET' && method !== 'DELETE' && body) fetchOpts.payload = body;

  var resp = UrlFetchApp.fetch(url, fetchOpts);
  var text = resp.getContentText();
  var code = resp.getResponseCode();
  if (code >= 300) throw new Error('SP-API error ' + code + ': ' + text);
  return JSON.parse(text || '{}');
}

/***** ========= RETRY ON 429 / BANDWIDTH HELPERS ========= *****/
function _isRateLimit429(err) {
  var msg = (err && err.message) ? err.message : String(err);
  // Covers both "SP-API error 429" and JSON body { "code": "QuotaExceeded" }
  return msg.indexOf('SP-API error 429') >= 0 || msg.indexOf('"code":"QuotaExceeded"') >= 0;
}

function _isBandwidthError(err) {
  var msg = (err && err.message) ? err.message : String(err);
  return msg.indexOf('Bandwidth quota exceeded') >= 0;
}

/**
 * Wrapper around spapiFetch that retries on 429 (QuotaExceeded) and
 * transient GAS bandwidth quota errors (15 s wait).
 */
function spapiFetchWithRetry(method, path, opts, attempts, waitMs) {
  attempts = (attempts == null) ? 3 : attempts;
  waitMs   = (waitMs == null)   ? 5000 : waitMs;

  var lastErr = null;
  for (var i = 0; i < attempts; i++) {
    try {
      return spapiFetch(method, path, opts);
    } catch (err) {
      lastErr = err;
      if (i < attempts - 1) {
        if (_isRateLimit429(err))  { Utilities.sleep(waitMs);  continue; }
        if (_isBandwidthError(err)) { Utilities.sleep(15000); continue; }
      }
      throw err;
    }
  }
  if (lastErr) throw lastErr;
  return null;
}

/***** ========= FBA OUTBOUND HELPERS (with 429 retry) ========= *****/
// maxAttempts: default 3 (server-side batch); pass 1 for formula cells to avoid 5s sleep × retries → timeout
function getFulfillmentOrderRaw(sellerFulfillmentOrderId, endpoint, maxAttempts) {
  var path = '/fba/outbound/2020-07-01/fulfillmentOrders/' + encodeURIComponent(sellerFulfillmentOrderId);
  var res = spapiFetchWithRetry('GET', path, { endpoint: endpoint }, maxAttempts != null ? maxAttempts : 3, 5000);
  return res.payload || res;
}

function getPackageTrackingDetails(packageNumber, endpoint) {
  var qs = 'packageNumber=' + encodeURIComponent(String(packageNumber));
  var path = '/fba/outbound/2020-07-01/tracking';
  // 3 attempts, 5s apart
  var res = spapiFetchWithRetry('GET', path, { queryString: qs, endpoint: endpoint }, 3, 5000);
  return res.payload || res;
}

function getTrackingBySellerFulfillmentOrderId(sfoId, endpoint) {
  if (!sfoId) throw new Error('sellerFulfillmentOrderId is required');
  var out = [];
  var fo = getFulfillmentOrderRaw(sfoId, endpoint);
  var shipments = (fo && fo.fulfillmentShipments) || [];
  for (var i = 0; i < shipments.length; i++) {
    var sh = shipments[i];
    var pkgs = (sh.fulfillmentShipmentPackage && sh.fulfillmentShipmentPackage.length)
      ? sh.fulfillmentShipmentPackage
      : (sh.packages || []);
    for (var j = 0; j < pkgs.length; j++) {
      var p = pkgs[j];
      var tn = p.trackingNumber || p.trackingId || p.amazonFulfillmentTrackingNumber || '';
      var carrier = p.carrierCode || p.carrierName || '';
      var pkgNo = p.packageNumber != null ? String(p.packageNumber) : '';
      var shipmentId = sh.amazonShipmentId || sh.shipmentId || '';
      if (!tn && pkgNo) {
        try {
          var det = getPackageTrackingDetails(pkgNo, endpoint);
          tn = det.trackingNumber || tn;
          carrier = det.carrierCode || carrier;
          Utilities.sleep(120); // small delay between per-package lookups
        } catch (e) {}
      }
      out.push({ trackingNumber: (tn || '').trim(), carrier: carrier || '', shipmentId: shipmentId, packageNumber: pkgNo });
    }
  }
  return out;
}

/***** ========= RETRY HELPERS (region fallbacks) ========= *****/
function _isRetryableRegionMismatchError(err) {
  var msg = (err && err.message) ? err.message : String(err);
  return (
    msg.indexOf('SP-API error 400') >= 0 &&
    (msg.indexOf('InvalidInput') >= 0 ||
     msg.indexOf('GetOrderByMerchantOrderIdRequest') >= 0 ||
     msg.indexOf('Unable to get order info') >= 0)
  );
}

function _tracksWithFallbacks(orderId, endpoints) {
  var lastErr = null;
  for (var i = 0; i < endpoints.length; i++) {
    var ep = endpoints[i];
    try {
      var tracks = getTrackingBySellerFulfillmentOrderId(String(orderId), ep);
      if (tracks && tracks.length > 0) return tracks;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err)) continue;
      throw err;
    }
  }
  if (lastErr) throw lastErr;
  return [];
}

function _isUnauthorizedError(err) {
  var msg = (err && err.message) ? err.message : String(err);
  return msg.indexOf('SP-API error 403') >= 0 &&
         msg.indexOf('"code":"Unauthorized"') >= 0 &&
         msg.indexOf('Access to requested resource is denied') >= 0;
}

// 400 InvalidInput "Unable to get order info" → treat as "no data" (blank cell)
function _isNoOrderInfoError(err) {
  var msg = (err && err.message) ? err.message : String(err);
  return msg.indexOf('SP-API error 400') >= 0 &&
         (msg.indexOf('"code":"InvalidInput"') >= 0 || msg.indexOf('InvalidInput') >= 0) &&
         (msg.indexOf('Unable to get order info') >= 0 || msg.indexOf('GetOrderByMerchantOrderIdRequest') >= 0);
}

// Returns true for error strings written back to cells ("EU ERR: ...", "JP ERR: ...", "ERR: ...")
// Used by backfillTrackingNumbers to detect cells that need a retry.
function _isErrorValue(v) {
  return /^(EU |JP )?ERR:/i.test(String(v || '').trim());
}

/***** ========= SHEET FUNCTIONS ========= *****/
function AMZTK(orderId) {
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'AMZTK_' + orderId;
  var cached = cache.get(key);

  if (cached !== null) return cached;

  try {
    var tracks = _tracksWithFallbacks(String(orderId), ['EU', 'FE']);
    var tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
      ? tracks[0].trackingNumber
      : '';

    // Found tracking number → stable, cache 6h. Still searching → retry in 10min.
    cache.put(key, tn, tn ? 21600 : 600);
    return tn;

  } catch (err) {
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) return '';
    // 429 / transient errors: do NOT cache — let the next recalculation retry.
    return 'EU ERR: ' + (err && err.message ? err.message : err);
  }
}

function AMZTK_JP(orderId) {
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'AMZTK_JP_' + orderId;
  var cached = cache.get(key);

  if (cached !== null) return cached;

  try {
    var tracks = _tracksWithFallbacks(String(orderId), ['FE', 'EU']);
    var tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
      ? tracks[0].trackingNumber
      : '';

    // Found tracking number → stable, cache 6h. Still searching → retry in 10min.
    cache.put(key, tn, tn ? 21600 : 600);
    return tn;

  } catch (err) {
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) return '';
    // 429 / transient errors: do NOT cache — let the next recalculation retry.
    return 'JP ERR: ' + (err && err.message ? err.message : err);
  }
}


/******************************************************
 *   MCF STOCK LOOKUP (FBA Inventory)
 *   Input: asin + region group ("EU" / "FE")
 *   Output: Available stock (integer)
 ******************************************************/
function getMcfStockByAsin(asin, marketplaceId) {
  if (!asin) return 0;

  var body = JSON.stringify({
    marketplaceIds: [marketplaceId],
    granularityType: "Marketplace",
    granularityId: marketplaceId,
    details: true,
    asin: asin
  });

  var path = "/fba/inventory/v1/summaries";

  try {
    var res = spapiFetchWithRetry("POST", path, {
      body: body,
      endpoint: marketplaceId   // IMPORTANT FIX
    });

    var summaries = res?.payload?.inventorySummaries || [];
    if (!summaries.length) return 0;

    var item = summaries.find(s => s.asin === asin);
    if (!item) return 0;

    return item.inventoryDetails?.available?.quantity || 0;

  } catch (e) {
    Logger.log("MCF STOCK ERROR: " + JSON.stringify(e));
    Logger.log("STACK: " + (e.stack || e));
    return "ERR";
  }
}

/***** ========= MCF FEE LOOKUP ========= *****/

/**
 * Returns the actual settled MCF fulfillment fee for an EU/UK/DE/FR/IT/ES order.
 * Uses the targeted Finances API: getFulfillmentOrderRaw → displayableOrderId →
 * listFinancialEventsByOrderId. Two fast calls, no date-range scan, no retry sleep.
 * Returns blank until settled — retries automatically on next recalculation (90 s on 429).
 *
 * Usage:
 *   =MCFFee(Q35)        — orderId only
 *   =MCFFee(P35, Q35)   — 2-arg form accepted for backwards compat; sentDate (P) is ignored,
 *                          orderId is taken from Q
 *
 * Required roles: Amazon Fulfillment + Finance and Accounting.
 *
 * @customfunction
 * @param {string} arg1  orderId when called with 1 arg; ignored sentDate when called with 2 args
 * @param {string} [arg2] orderId (col Q) when called with 2 args
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee(arg1, arg2) {
  // MCFFee(Q)    → orderId only
  // MCFFee(P, Q) → sentDate (P col, raw GAS value) + orderId (Q col)
  var orderId, sentDate;
  if (arg2 !== undefined && arg2 !== null && String(arg2).trim() !== '') {
    sentDate = arg1;               // raw Date object or string from sheet cell
    orderId  = String(arg2).trim();
  } else {
    orderId  = String(arg1 || '').trim();
    sentDate = null;
  }
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE2_' + orderId;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['EU', 'FE'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = _fetchMcfFeeFinancesApi(orderId, endpoints[i], sentDate);
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), fee === '' ? 600 : 21600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 21600); return ''; }
      if (_isRateLimit429(err))      { cache.put(key, '__EMPTY__', 90);    return ''; }
      throw err;
    }
  }

  if (lastErr) {
    if (_isUnauthorizedError(lastErr)) { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isNoOrderInfoError(lastErr))  { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isRateLimit429(lastErr))      { cache.put(key, '__EMPTY__', 90);    return ''; }
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}

/**
 * Returns the actual settled MCF fulfillment fee for a Japan / AU / SG order.
 * Same as MCFFee but tries the FE (Far East) endpoint first.
 *
 * Usage: =MCFFee_JP(Q35) or =MCFFee_JP(P35, Q35)
 *
 * @customfunction
 * @param {string} arg1  orderId when called with 1 arg; ignored sentDate when called with 2 args
 * @param {string} [arg2] orderId (col Q) when called with 2 args
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee_JP(arg1, arg2) {
  var orderId, sentDate;
  if (arg2 !== undefined && arg2 !== null && String(arg2).trim() !== '') {
    sentDate = arg1;
    orderId  = String(arg2).trim();
  } else {
    orderId  = String(arg1 || '').trim();
    sentDate = null;
  }
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE2_JP_' + orderId;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['FE', 'EU'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = _fetchMcfFeeFinancesApi(orderId, endpoints[i], sentDate);
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), fee === '' ? 600 : 21600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 21600); return ''; }
      if (_isRateLimit429(err))      { cache.put(key, '__EMPTY__', 90);    return ''; }
      throw err;
    }
  }

  if (lastErr) {
    if (_isUnauthorizedError(lastErr)) { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isNoOrderInfoError(lastErr))  { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isRateLimit429(lastErr))      { cache.put(key, '__EMPTY__', 90);    return ''; }
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}


/**
 * Safely converts a value from a sheet cell (may be a Date object, a numeric serial, or a
 * yyyy-mm-dd string) into a JS Date. Returns null when absent or invalid.
 * GAS custom functions receive date-formatted cells as Date objects, NOT strings.
 */
function _toSafeDate(val) {
  if (val == null || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : new Date(val.getTime());
  // Numeric: Google Sheets date serial (days since 1899-12-30)
  var n = Number(val);
  if (!isNaN(n) && n > 0) {
    var d = new Date((n - 25569) * 86400000); // serial → Unix ms (UTC)
    return isNaN(d.getTime()) ? null : d;
  }
  // String: ISO yyyy-mm-dd first, then generic parse
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3]));
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

/**
 * Targeted Finances API lookup by Amazon order ID (displayableOrderId).
 * GET /finances/v0/orders/{orderId}/financialEvents
 * Returns only that order's ShipmentEvents — tiny response, no pagination needed.
 * maxAttempts: pass 1 for formula cells (no retry sleep → no timeout risk).
 */
function _listFinancialEventsByOrderId(amazonOrderId, ep, maxAttempts) {
  var path = '/finances/v0/orders/' + encodeURIComponent(amazonOrderId) + '/financialEvents';
  var res = spapiFetchWithRetry('GET', path, { endpoint: ep }, maxAttempts != null ? maxAttempts : 3, 5000);
  var payload = res.payload || res;
  return (payload.FinancialEvents || {}).ShipmentEventList || [];
}

/**
 * Sums MCF fees across all shipment events (OrderFeeList + ShipmentFeeList + ItemFeeList).
 * Use when all events in the array already belong to one order (targeted endpoint result).
 */
function _sumAllMcfFees(shipments) {
  var total = 0;
  for (var i = 0; i < shipments.length; i++) {
    var ev = shipments[i];
    (ev.OrderFeeList    || []).forEach(function(f) { if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0); });
    (ev.ShipmentFeeList || []).forEach(function(f) { if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0); });
    (ev.ShipmentItemList || []).forEach(function(item) {
      (item.ItemFeeList || []).forEach(function(f) { if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0); });
    });
  }
  return total !== 0 ? Math.abs(total) : '';
}

/**
 * Finances API — actual settled MCF fee, formula-safe.
 *
 * Two-strategy approach, all calls with 1 attempt (no retry sleep → no 30s timeout):
 *
 * Strategy A — targeted (for orders linked to an Amazon marketplace order):
 *   getFulfillmentOrderRaw → displayableOrderId (Amazon 3-7-7) →
 *   listFinancialEventsByOrderId → sum fees
 *
 * Strategy B — date-range scan (for seller-created MCF orders with GCX orderId):
 *   Finances API date-range (sentDate..sentDate+60d, 2 pages, 1 attempt) →
 *   match SellerOrderId = orderId → sum fees
 *
 * Why 2 strategies: seller-created MCF orders store SellerOrderId = GCX orderId
 * in the Finances API (confirmed: backfillMCFFees() matches this way).
 * The targeted endpoint only works for Amazon marketplace–format IDs.
 *
 * On 429: error propagates up to MCFFee → caches '' for 90 s → auto-retries.
 */
function _fetchMcfFeeFinancesApi(orderId, ep, sentDate) {
  // Strategy A: targeted via displayableOrderId (fast: 2 calls, ~1-2s)
  try {
    var foResult = getFulfillmentOrderRaw(orderId, ep, 1);
    var displayableId = ((foResult.fulfillmentOrder || {}).displayableOrderId || '').trim();
    if (/^\w{3}-\d{7}-\d{7}$/.test(displayableId)) {
      var targeted = _listFinancialEventsByOrderId(displayableId, ep, 1);
      var feeA = _sumAllMcfFees(targeted);
      if (feeA !== '') return feeA;
    }
  } catch (e) { /* FO lookup failed (e.g. wrong endpoint) — try date-range */ }

  // Strategy B: date-range scan, match by SellerOrderId (2 pages, 1 attempt, no retry sleep)
  var postedAfter, postedBefore;
  var parsedSentDate = _toSafeDate(sentDate);
  if (parsedSentDate) {
    postedAfter  = parsedSentDate;
    postedBefore = new Date(parsedSentDate.getTime());
    postedBefore.setDate(postedBefore.getDate() + 60);
  } else {
    var _ref = new Date(Date.now() - 5 * 60 * 1000);
    postedAfter  = new Date(_ref.getTime() - 30 * 24 * 3600 * 1000); // last 30 days
    postedBefore = new Date(_ref);
  }
  var _now = new Date(Date.now() - 5 * 60 * 1000);
  if (postedBefore > _now) postedBefore = _now;

  var shipments = _collectShipmentEvents(ep, postedAfter, postedBefore, 2, 1); // 2 pages, 1 attempt
  return _sumMcfFeeFromShipments(shipments, orderId);
}

// Matches FBA / fulfillment fee type names used in Finances API ShipmentEvent
function _isMcfFeeType(feeType) {
  if (!feeType) return false;
  var t = String(feeType).toUpperCase();
  return t.indexOf('FBA') >= 0 || t.indexOf('FULFILLMENT') >= 0;
}

/**
 * Fetches ALL ShipmentEvent financial events for a single date window on one endpoint.
 * Returns a plain object mapping SellerOrderId → fee (absolute value).
 * Used by backfillMCFFees() to build a bulk fee map instead of one call per order.
 */
function _buildFeeMapForWindow(ep, postedAfter, postedBefore) {
  var feeMap    = {};
  var nextToken = null;
  var maxPages  = 20;
  var page      = 0;

  do {
    var qs = 'PostedAfter='   + encodeURIComponent(postedAfter.toISOString()) +
             '&PostedBefore=' + encodeURIComponent(postedBefore.toISOString()) +
             '&MaxResultsPerPage=100';
    if (nextToken) qs += '&NextToken=' + encodeURIComponent(nextToken);

    var res      = spapiFetchWithRetry('GET', '/finances/v0/financialEvents', { queryString: qs, endpoint: ep }, 3, 5000);
    var payload  = res.payload || res;
    nextToken    = payload.NextToken || null;
    var shipments = (payload.FinancialEvents || {}).ShipmentEventList || [];

    for (var i = 0; i < shipments.length; i++) {
      var ev  = shipments[i];
      var sid = String(ev.SellerOrderId || '').trim();
      if (!sid) continue;

      var total = 0;
      (ev.OrderFeeList || []).forEach(function(f) {
        if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentFeeList || []).forEach(function(f) {
        if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentItemList || []).forEach(function(item) {
        (item.ItemFeeList || []).forEach(function(f) {
          if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
        });
      });

      if (total !== 0) feeMap[sid] = Math.abs(total);
    }

    page++;
  } while (nextToken && page < maxPages);

  return feeMap;
}

/**
 * Paginates ShipmentEventList for one endpoint + date window.
 * Returns a flat array of all ShipmentEvent objects (up to maxPages × 100).
 * maxAttempts: default 3 (batch); pass 1 for formula cells (no retry sleep → no timeout).
 */
function _collectShipmentEvents(ep, postedAfter, postedBefore, maxPages, maxAttempts) {
  var all = [], nextToken = null, page = 0;
  maxPages    = maxPages    || 5;
  maxAttempts = maxAttempts != null ? maxAttempts : 3;
  do {
    var qs = 'PostedAfter='   + encodeURIComponent(postedAfter.toISOString()) +
             '&PostedBefore=' + encodeURIComponent(postedBefore.toISOString()) +
             '&MaxResultsPerPage=100';
    if (nextToken) qs += '&NextToken=' + encodeURIComponent(nextToken);
    var res      = spapiFetchWithRetry('GET', '/finances/v0/financialEvents', { queryString: qs, endpoint: ep }, maxAttempts, 5000);
    var payload  = res.payload || res;
    nextToken    = payload.NextToken || null;
    var batch    = (payload.FinancialEvents || {}).ShipmentEventList || [];
    all          = all.concat(batch);
    page++;
  } while (nextToken && page < maxPages);
  return all;
}

/**
 * Finds the first ShipmentEvent matching targetOrderId (by SellerOrderId or AmazonOrderId)
 * and sums its MCF fee lines including OrderFeeList (MCF order-level fees per SP-API docs).
 * Returns Math.abs(total) on match, '' if not found.
 */
function _sumMcfFeeFromShipments(shipments, targetOrderId) {
  var target = String(targetOrderId).trim().toUpperCase();
  for (var i = 0; i < shipments.length; i++) {
    var ev = shipments[i];
    var sid = String(ev.SellerOrderId  || '').trim().toUpperCase();
    var aid = String(ev.AmazonOrderId  || '').trim().toUpperCase();
    if (sid !== target && aid !== target) continue;
    var total = 0;
    // OrderFeeList: order-level fees — specifically applicable to MCF orders (SP-API docs)
    (ev.OrderFeeList || []).forEach(function(f) {
      if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
    });
    (ev.ShipmentFeeList || []).forEach(function(f) {
      if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
    });
    (ev.ShipmentItemList || []).forEach(function(item) {
      (item.ItemFeeList || []).forEach(function(f) {
        if (_isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
    });
    return Math.abs(total);
  }
  return '';
}

/**
 * Diagnoses why =MCFFee(P,Q) returns blank for a given order.
 * Fast: 1-attempt calls only (completes in ~2-3 s, no timeout).
 *
 * Runs 3 checks and returns results as a table:
 *   Step A  – getFulfillmentOrderRaw → displayableOrderId
 *   Step B  – listFinancialEventsByOrderId(displayableOrderId) → targeted events
 *   Step C  – listFinancialEvents date-range page 1 → first 100 events, match check
 *
 * Usage: =MCFFeeDebug(Q35, P35)
 *
 * @customfunction
 * @param {string} orderId  sellerFulfillmentOrderId (col Q)
 * @param {string} sentDate sent date (col P, yyyy-mm-dd or date cell)
 * @return {Array} diagnostic rows
 */
function MCFFeeDebug(orderId, sentDate) {
  if (!orderId) return [['orderId is required']];
  orderId = String(orderId).trim();

  var rows = [['Step', 'Key', 'Value']];

  try {
    // ── Step A: getFulfillmentOrderRaw ──────────────────────────────────
    var displayableId = '', foStatus = '';
    try {
      var foResult = getFulfillmentOrderRaw(orderId, 'EU', 1);
      var fo = foResult.fulfillmentOrder || {};
      displayableId = (fo.displayableOrderId || '').trim();
      foStatus = 'OK';
      rows.push(['A: FBA Outbound (EU)', 'displayableOrderId', displayableId || '(empty)']);
      rows.push(['A: FBA Outbound (EU)', 'receivedDate',       fo.receivedDate || '(empty)']);
      rows.push(['A: FBA Outbound (EU)', 'marketplaceId',      fo.marketplaceId || '(empty)']);
    } catch (e) {
      foStatus = 'ERR: ' + (e.message || e);
      rows.push(['A: FBA Outbound (EU)', 'error', foStatus]);
    }

    // ── Step B: listFinancialEventsByOrderId (targeted) ─────────────────
    var isAmazonFmt = /^\w{3}-\d{7}-\d{7}$/.test(displayableId);
    rows.push(['B: Targeted endpoint', 'displayableId Amazon-format?', isAmazonFmt ? 'YES → querying' : 'NO → skipped']);
    if (isAmazonFmt) {
      try {
        var targeted = _listFinancialEventsByOrderId(displayableId, 'EU', 1);
        rows.push(['B: Targeted endpoint', 'ShipmentEvents returned', String(targeted.length)]);
        targeted.slice(0, 5).forEach(function(ev, i) {
          var sid = ev.SellerOrderId || '';
          var aid = ev.AmazonOrderId || '';
          var fee = _sumAllMcfFees([ev]);
          rows.push(['B: event ' + i, 'SellerOrderId / AmazonOrderId', sid + ' / ' + aid]);
          rows.push(['B: event ' + i, 'MCF fee sum', String(fee)]);
        });
      } catch (e) {
        rows.push(['B: Targeted endpoint', 'error', String(e.message || e)]);
      }
    }

    // ── Step C: date-range scan page 1 ─────────────────────────────────
    var parsedSentDate = _toSafeDate(sentDate);
    var postedAfter, postedBefore;
    if (parsedSentDate) {
      postedAfter  = parsedSentDate;
      postedBefore = new Date(parsedSentDate.getTime());
      postedBefore.setDate(postedBefore.getDate() + 60);
    } else {
      var _ref = new Date(Date.now() - 5 * 60 * 1000);
      postedAfter  = new Date(_ref.getTime() - 30 * 24 * 3600 * 1000);
      postedBefore = new Date(_ref);
    }
    var _now = new Date(Date.now() - 5 * 60 * 1000);
    if (postedBefore > _now) postedBefore = _now;

    rows.push(['C: Date-range scan', 'window', postedAfter.toISOString().slice(0,10) + ' → ' + postedBefore.toISOString().slice(0,10)]);

    try {
      var page1 = _collectShipmentEvents('EU', postedAfter, postedBefore, 1, 1); // 1 page, 1 attempt
      rows.push(['C: Date-range scan', 'events on page 1', String(page1.length)]);

      var matchFound = false;
      page1.forEach(function(ev) {
        var sid = (ev.SellerOrderId || '').trim().toUpperCase();
        var aid = (ev.AmazonOrderId || '').trim().toUpperCase();
        var target = orderId.toUpperCase();
        if (sid === target || aid === target) {
          matchFound = true;
          rows.push(['C: MATCH FOUND', 'SellerOrderId', ev.SellerOrderId || '']);
          rows.push(['C: MATCH FOUND', 'AmazonOrderId', ev.AmazonOrderId || '']);
          rows.push(['C: MATCH FOUND', 'MCF fee sum',   String(_sumMcfFeeFromShipments([ev], orderId))]);
        }
      });

      if (!matchFound) {
        rows.push(['C: NO MATCH on page 1', 'sample SellerOrderIds (first 5)',
          page1.slice(0, 5).map(function(e) { return e.SellerOrderId || ''; }).join(' | ')]);
      }
    } catch (e) {
      rows.push(['C: Date-range scan', 'error', String(e.message || e)]);
    }

    return rows;
  } catch (e) {
    return [['ERR', '', String(e.message || e)]];
  }
}

/**
 * SERVER-SIDE diagnostic — run from GAS editor, NOT as a sheet formula.
 * No 30s timeout. Results appear in View → Logs.
 *
 * HOW TO USE:
 *   1. Open Apps Script editor
 *   2. Paste a real order ID into the variable below
 *   3. Select this function and click Run
 *   4. Check View → Logs
 */
function debugMcfFeeManual() {
  var ORDER_ID  = 'PASTE-YOUR-ORDER-ID-HERE';   // ← replace with Q col value, e.g. GCX-IT-250404-24
  var SENT_DATE = '2025-04-04';                  // ← replace with P col value (yyyy-mm-dd)
  var EP        = 'EU';                          // 'EU' or 'FE' (use 'FE' for JP orders)

  Logger.log('=== debugMcfFeeManual ===');
  Logger.log('orderId=%s  sentDate=%s  ep=%s', ORDER_ID, SENT_DATE, EP);

  // ── Step 0: LWA token ──────────────────────────────────────────────
  try {
    var tok = getLwaAccessToken(EP);
    Logger.log('[0] LWA token OK (length=%s)', tok.length);
  } catch (e) {
    Logger.log('[0] LWA token FAILED: %s', e.message || e);
    return;
  }

  // ── Step A: getFulfillmentOrderRaw ─────────────────────────────────
  var displayableId = '';
  try {
    var foResult = getFulfillmentOrderRaw(ORDER_ID, EP, 1);
    var fo = foResult.fulfillmentOrder || {};
    displayableId = (fo.displayableOrderId || '').trim();
    Logger.log('[A] FBA Outbound OK');
    Logger.log('    displayableOrderId : %s', displayableId || '(empty)');
    Logger.log('    receivedDate       : %s', fo.receivedDate || '(empty)');
    Logger.log('    marketplaceId      : %s', fo.marketplaceId || '(empty)');
    Logger.log('    sellerFulfillmentOrderId : %s', fo.sellerFulfillmentOrderId || '(empty)');
  } catch (e) {
    Logger.log('[A] FBA Outbound FAILED: %s', e.message || e);
  }

  // ── Step B: targeted Finances endpoint (only if Amazon 3-7-7 format) ──
  var isAmazonFmt = /^\w{3}-\d{7}-\d{7}$/.test(displayableId);
  Logger.log('[B] displayableId Amazon-format? %s', isAmazonFmt ? 'YES' : 'NO → skipping targeted call');
  if (isAmazonFmt) {
    try {
      var targeted = _listFinancialEventsByOrderId(displayableId, EP, 1);
      Logger.log('[B] ShipmentEvents returned: %s', targeted.length);
      targeted.slice(0, 3).forEach(function(ev, i) {
        Logger.log('    [B.%s] SellerOrderId=%s  AmazonOrderId=%s  feeSum=%s',
          i, ev.SellerOrderId || '', ev.AmazonOrderId || '', _sumAllMcfFees([ev]));
      });
    } catch (e) {
      Logger.log('[B] targeted call FAILED: %s', e.message || e);
    }
  }

  // ── Step C: date-range scan (full 3 pages, 3 attempts — server-side has 6 min timeout) ──
  var postedAfter  = new Date(SENT_DATE + 'T00:00:00Z');
  var postedBefore = new Date(postedAfter.getTime() + 60 * 24 * 3600 * 1000); // +60 days
  var now = new Date(Date.now() - 5 * 60 * 1000);
  if (postedBefore > now) postedBefore = now;

  Logger.log('[C] date-range scan %s → %s', postedAfter.toISOString().slice(0,10), postedBefore.toISOString().slice(0,10));
  try {
    var events = _collectShipmentEvents(EP, postedAfter, postedBefore, 3, 3); // 3 pages, 3 attempts
    Logger.log('[C] total ShipmentEvents fetched: %s', events.length);

    var matchFound = false;
    events.forEach(function(ev) {
      var sid = (ev.SellerOrderId || '').trim();
      var aid = (ev.AmazonOrderId || '').trim();
      if (sid.toUpperCase() === ORDER_ID.toUpperCase() || aid.toUpperCase() === ORDER_ID.toUpperCase()) {
        matchFound = true;
        Logger.log('[C] *** MATCH *** SellerOrderId=%s  AmazonOrderId=%s  feeSum=%s',
          sid, aid, _sumMcfFeeFromShipments([ev], ORDER_ID));
      }
    });

    if (!matchFound) {
      Logger.log('[C] No match found. Sample SellerOrderIds (first 10):');
      events.slice(0, 10).forEach(function(ev, i) {
        Logger.log('    [%s] %s / %s', i, ev.SellerOrderId || '', ev.AmazonOrderId || '');
      });
    }
  } catch (e) {
    Logger.log('[C] date-range scan FAILED: %s', e.message || e);
  }

  Logger.log('=== done ===');
}

/**
 * getFulfillmentPreview method — estimated MCF fee (instant, may differ from actual).
 * Currency: GBP for UK orders, EUR for other EU orders.
 */
function _fetchMcfFeePreview(orderId, ep) {
  var result = getFulfillmentOrderRaw(orderId, ep);
  var fo     = result.fulfillmentOrder || {};
  var items  = result.fulfillmentOrderItems || [];
  if (!items.length || !fo.destinationAddress) return '';

  var previewItems = items.map(function(item, idx) {
    return {
      sellerSku: item.sellerSku,
      sellerFulfillmentOrderItemId: 'prev_' + idx,
      quantity: item.quantity
    };
  });

  var body = JSON.stringify({
    marketplaceId: fo.marketplaceId || '',
    address: fo.destinationAddress,
    items: previewItems,
    shippingSpeedCategories: ['Expedited']
  });

  var res = spapiFetchWithRetry(
    'POST',
    '/fba/outbound/2020-07-01/fulfillmentOrders/preview',
    { body: body, endpoint: ep },
    3, 5000
  );

  var previews = ((res.payload || res).fulfillmentPreviews) || [];
  var preview  = null;
  for (var i = 0; i < previews.length; i++) {
    if (previews[i].shippingSpeedCategory === 'Expedited') { preview = previews[i]; break; }
  }
  if (!preview || !preview.estimatedFees || !preview.estimatedFees.length) return '';

  var total = 0;
  for (var j = 0; j < preview.estimatedFees.length; j++) {
    total += parseFloat(preview.estimatedFees[j].amount.value || 0);
  }
  return total;
}
