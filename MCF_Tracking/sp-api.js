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
function getFulfillmentOrderRaw(sellerFulfillmentOrderId, endpoint) {
  var path = '/fba/outbound/2020-07-01/fulfillmentOrders/' + encodeURIComponent(sellerFulfillmentOrderId);
  // 3 attempts, 5s apart
  var res = spapiFetchWithRetry('GET', path, { endpoint: endpoint }, 3, 5000);
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
 * Uses the Finances API (actual charged amount, available days after shipment).
 * Returns blank until settled — retries automatically on the next sheet recalculation.
 * Currency matches the marketplace: GBP for UK, EUR for other EU marketplaces.
 *
 * Usage:
 *   =MCFFee(Q35)        — orderId only, searches last 180 days
 *   =MCFFee(P35, Q35)   — sentDate (col P, yyyy-mm-dd) + orderId (col Q), faster & less bandwidth
 *
 * Required roles: Amazon Fulfillment + Finance and Accounting.
 *
 * @customfunction
 * @param {string} arg1  orderId when called with 1 arg; sentDate (yyyy-mm-dd, col P) when called with 2 args
 * @param {string} [arg2] orderId (col Q) when called with 2 args
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee(arg1, arg2) {
  // MCFFee(orderId)           → 1-arg: orderId only, uses 180-day window
  // MCFFee(sentDate, orderId) → 2-arg: sentDate (P col) + orderId (Q col)
  // sentDate is kept as the raw GAS value (may be a Date object) — _toSafeDate handles both
  var orderId, sentDate;
  if (arg2 !== undefined && arg2 !== null && String(arg2).trim() !== '') {
    sentDate = arg1;                // raw — Date object or string from sheet cell
    orderId  = String(arg2).trim();
  } else {
    orderId  = String(arg1 || '').trim();
    sentDate = null;
  }
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var sentDateStr = sentDate ? (sentDate instanceof Date
    ? Utilities.formatDate(sentDate, 'UTC', 'yyyy-MM-dd')
    : String(sentDate).trim()) : '';
  var key = 'MCFFEE_' + orderId + (sentDateStr ? '_' + sentDateStr : '');
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
 * Usage:
 *   =MCFFee_JP(Q35)        — orderId only
 *   =MCFFee_JP(P35, Q35)   — sentDate (col P) + orderId (col Q)
 *
 * @customfunction
 * @param {string} arg1  orderId when called with 1 arg; sentDate (yyyy-mm-dd, col P) when called with 2 args
 * @param {string} [arg2] orderId (col Q) when called with 2 args
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee_JP(arg1, arg2) {
  var orderId, sentDate;
  if (arg2 !== undefined && arg2 !== null && String(arg2).trim() !== '') {
    sentDate = arg1;                // raw — Date object or string from sheet cell
    orderId  = String(arg2).trim();
  } else {
    orderId  = String(arg1 || '').trim();
    sentDate = null;
  }
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var sentDateStr = sentDate ? (sentDate instanceof Date
    ? Utilities.formatDate(sentDate, 'UTC', 'yyyy-MM-dd')
    : String(sentDate).trim()) : '';
  var key = 'MCFFEE_JP_' + orderId + (sentDateStr ? '_' + sentDateStr : '');
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
 * Targeted Finances API lookup by Amazon order ID (displayableOrderId).
 * GET /finances/v0/orders/{orderId}/financialEvents
 * Much more bandwidth-efficient than a date-range scan — returns only that order's events.
 * Returns array of ShipmentEvent objects (may be empty if not yet settled).
 */
function _listFinancialEventsByOrderId(amazonOrderId, ep) {
  var path = '/finances/v0/orders/' + encodeURIComponent(amazonOrderId) + '/financialEvents';
  var res = spapiFetchWithRetry('GET', path, { endpoint: ep }, 3, 5000);
  var payload = res.payload || res;
  return (payload.FinancialEvents || {}).ShipmentEventList || [];
}

/**
 * Sums MCF fees (ShipmentFeeList + OrderFeeList + ItemFeeList) across all shipment events.
 * Used when all events in the array already belong to the target order (targeted endpoint).
 * Returns Math.abs(total) or '' if total is 0 / array is empty.
 */
function _sumAllMcfFees(shipments) {
  var total = 0;
  for (var i = 0; i < shipments.length; i++) {
    var ev = shipments[i];
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
  }
  return total !== 0 ? Math.abs(total) : '';
}

/**
 * Finances API method — actual settled MCF fee.
 *
 * Strategy (ordered by bandwidth cost, lowest first):
 * 1. Call getFulfillmentOrderRaw to get displayableOrderId + receivedDate (one small API call).
 * 2. If displayableOrderId is Amazon 3-7-7 format (e.g. S02-XXXXXXX-XXXXXXX), use the targeted
 *    listFinancialEventsByOrderId endpoint — returns only that order's events, very small response.
 * 3. Fall back to date-range scan (sentDate±60d, or receivedDate±60d, or last 180d) and match
 *    by both SellerOrderId and displayableOrderId.
 *
 * Fees are stored as negative in Finances API; returns Math.abs(total).
 * Returns '' if the order has not yet settled.
 */

/**
 * Safely converts a value from a sheet cell (may be a Date object, a number serial, or a
 * yyyy-mm-dd string) to a JavaScript Date. Returns null when the value is absent or invalid.
 * GAS custom functions receive date cells as Date objects, NOT strings — so we must handle both.
 */
function _toSafeDate(val) {
  if (val == null || val === '') return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : new Date(val.getTime());
  // Numeric: Google Sheets date serial (days since 1899-12-30)
  var n = Number(val);
  if (!isNaN(n) && n > 0) {
    var d = new Date((n - 25569) * 86400000); // convert serial → Unix ms (UTC)
    return isNaN(d.getTime()) ? null : d;
  }
  // String: try ISO format first (yyyy-mm-dd), then generic parse
  var s = String(val).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Date.UTC(+iso[1], +iso[2] - 1, +iso[3]));
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function _fetchMcfFeeFinancesApi(orderId, ep, sentDate) {
  // Step 1: Resolve fulfillment order metadata (displayableOrderId, receivedDate).
  var foResult = null, displayableId = '', receivedDate = '';
  try {
    foResult = getFulfillmentOrderRaw(orderId, ep);
    var fo = foResult.fulfillmentOrder || {};
    displayableId = (fo.displayableOrderId || '').trim();
    receivedDate  = (fo.receivedDate       || '').trim();
  } catch (e) { /* FO not found on this endpoint — proceed without metadata */ }

  // Step 2: Targeted lookup via listFinancialEventsByOrderId (requires Amazon 3-7-7 format ID).
  // Amazon order IDs look like "S02-1234567-1234567" or "114-1234567-1234567".
  var isAmazonFmt = /^\w{3}-\d{7}-\d{7}$/.test(displayableId);
  if (isAmazonFmt) {
    try {
      var targeted = _listFinancialEventsByOrderId(displayableId, ep);
      if (targeted.length) {
        var fee0 = _sumAllMcfFees(targeted);
        if (fee0 !== '') return fee0;
      }
    } catch (e) { /* targeted lookup failed — fall through to date-range */ }
  }

  // Step 3: Date-range scan fallback.
  // _toSafeDate handles Date objects passed by GAS from date-formatted cells (P col).
  var postedAfter, postedBefore;
  var parsedSentDate = _toSafeDate(sentDate);
  if (parsedSentDate) {
    postedAfter  = parsedSentDate;
    postedBefore = new Date(parsedSentDate.getTime());
    postedBefore.setDate(postedBefore.getDate() + 60);
  } else if (receivedDate) {
    postedAfter  = new Date(receivedDate); // ISO string from API — always valid
    postedBefore = new Date(postedAfter.getTime());
    postedBefore.setDate(postedBefore.getDate() + 60);
  } else {
    var _ref = new Date(Date.now() - 5 * 60 * 1000);
    postedAfter  = new Date(_ref.getTime() - 180 * 24 * 3600 * 1000);
    postedBefore = new Date(_ref);
  }

  var _now = new Date(Date.now() - 5 * 60 * 1000);
  if (postedBefore > _now) postedBefore = _now;

  var shipments = _collectShipmentEvents(ep, postedAfter, postedBefore, 5);

  // Match by GCX order ID first
  var fee = _sumMcfFeeFromShipments(shipments, orderId);
  if (fee !== '') return fee;

  // Fallback match by displayableOrderId (MCF orders linked to Amazon marketplace orders
  // settle in Finances API under the Amazon order ID, not the GCX order ID)
  if (displayableId && displayableId !== orderId) {
    var fee2 = _sumMcfFeeFromShipments(shipments, displayableId);
    if (fee2 !== '') return fee2;
  }

  return ''; // order not yet settled — caller caches as __EMPTY__ for 10min and retries
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
 */
function _collectShipmentEvents(ep, postedAfter, postedBefore, maxPages) {
  var all = [], nextToken = null, page = 0;
  maxPages = maxPages || 5;
  do {
    var qs = 'PostedAfter='   + encodeURIComponent(postedAfter.toISOString()) +
             '&PostedBefore=' + encodeURIComponent(postedBefore.toISOString()) +
             '&MaxResultsPerPage=100';
    if (nextToken) qs += '&NextToken=' + encodeURIComponent(nextToken);
    var res      = spapiFetchWithRetry('GET', '/finances/v0/financialEvents', { queryString: qs, endpoint: ep }, 3, 5000);
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
 * Debug: shows SellerOrderIds found in Finances API around this order's date window.
 * Also resolves displayableOrderId so you can see if the fee settled under a different ID.
 * @customfunction
 * @param {string} orderId The sellerFulfillmentOrderId to debug
 * @param {string} [sentDate] Optional yyyy-mm-dd sent date from col P.
 * @return {Array} SellerOrderId_in_API | Input_orderId | Exact_match | FeeTypes | Total
 */
function MCFFeeDebug(orderId, sentDate) {
  if (!orderId) return [['orderId is required']];
  try {
    var postedAfter, postedBefore, dateSource;

    var parsedSent = _toSafeDate(sentDate);
    if (parsedSent) {
      postedAfter  = parsedSent;
      postedBefore = new Date(parsedSent.getTime());
      postedBefore.setDate(postedBefore.getDate() + 90);
      dateSource = 'sentDate (P col): ' + parsedSent.toISOString().slice(0, 10);
    } else {
      var result = getFulfillmentOrderRaw(String(orderId), 'EU');
      var fo = result.fulfillmentOrder || {};
      if (!fo.receivedDate) return [['Order found but no receivedDate — check order ID']];
      postedAfter  = new Date(fo.receivedDate);
      postedBefore = new Date(fo.receivedDate);
      postedBefore.setDate(postedBefore.getDate() + 90);
      dateSource = 'receivedDate (API): ' + fo.receivedDate;
    }

    var now = new Date(Date.now() - 5 * 60 * 1000);
    if (postedBefore > now) postedBefore = now;

    var shipments = _collectShipmentEvents('EU', postedAfter, postedBefore, 5);

    // Resolve displayableOrderId for fallback matching info
    var displayableId = '';
    try {
      var foResult  = getFulfillmentOrderRaw(String(orderId), 'EU');
      var candidate = ((foResult.fulfillmentOrder || {}).displayableOrderId || '').trim();
      if (candidate && candidate !== String(orderId).trim()) displayableId = candidate;
    } catch (e) { /* ignore */ }

    var rows = [['SellerOrderId_in_API', 'Input_orderId', 'Exact_match', 'FeeTypes', 'Total']];

    if (!shipments.length) {
      rows.push(['(no ShipmentEvents in window)', orderId, '', '', '']);
      rows.push([dateSource, 'window end: ' + postedBefore.toISOString(), '', '', '']);
      return rows;
    }

    shipments.forEach(function(ev) {
      var sid = String(ev.SellerOrderId || '');
      var aid = String(ev.AmazonOrderId || '');
      var displayKey = sid + (aid && aid !== sid ? ' / AmazonOrderId:' + aid : '');
      var match = sid.trim().toUpperCase() === String(orderId).trim().toUpperCase()         ? 'YES'
                : aid.trim().toUpperCase() === String(orderId).trim().toUpperCase()         ? 'YES (AmazonOrderId)'
                : (displayableId && sid.trim().toUpperCase() === displayableId.toUpperCase()) ? 'YES (displayableOrderId)'
                : (displayableId && aid.trim().toUpperCase() === displayableId.toUpperCase()) ? 'YES (AmazonOrderId=displayable)'
                : 'no';
      var feeTypes = [], total = 0;
      (ev.OrderFeeList || []).forEach(function(f) {
        feeTypes.push('ORDER:' + f.FeeType);
        total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentFeeList || []).forEach(function(f) {
        feeTypes.push(f.FeeType);
        total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentItemList || []).forEach(function(item) {
        (item.ItemFeeList || []).forEach(function(f) {
          feeTypes.push(f.FeeType);
          total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
        });
      });
      rows.push([displayKey, orderId, match, feeTypes.join(', '), Math.abs(total)]);
    });

    rows.push([dateSource, 'window end: ' + postedBefore.toISOString(), 'Total events: ' + shipments.length, '', '']);
    if (displayableId) rows.push(['displayableOrderId fallback', displayableId, '', '', '']);
    return rows;
  } catch(e) { return [['ERR: ' + (e.message || e)]]; }
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
