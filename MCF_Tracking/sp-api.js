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

  // Cache LWA tokens for 50 min (tokens last 60 min) to avoid one HTTP call per spapiFetch.
  var cache = CacheService.getScriptCache();
  var cacheKey = 'LWA_TOKEN_' + prof;
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
  cache.put(cacheKey, body.access_token, 3000); // 50 min
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

/***** ========= RETRY ON 429 HELPERS ========= *****/
function _isRateLimit429(err) {
  var msg = (err && err.message) ? err.message : String(err);
  // Covers both "SP-API error 429" and JSON body { "code": "QuotaExceeded" }
  return msg.indexOf('SP-API error 429') >= 0 || msg.indexOf('"code":"QuotaExceeded"') >= 0;
}

/**
 * Wrapper around spapiFetch that retries when 429 (QuotaExceeded).
 * @param {string} method
 * @param {string} path
 * @param {Object} opts - same as spapiFetch opts (queryString, body, endpoint)
 * @param {number} [attempts=3] - total attempts including first try
 * @param {number} [waitMs=5000] - wait between retries (ms)
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
      if (_isRateLimit429(err) && i < attempts - 1) {
        Utilities.sleep(waitMs);
        continue;
      }
      throw err; // non-429 or out of attempts
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
 * Returns the MCF fulfillment fee for an existing order.
 *
 * method "FinancesAPI": queries the Finances API for the actual settled fee.
 * Returns blank until the order settles (usually a few days after shipment) — retries automatically.
 * Currency matches the marketplace: GBP for UK orders, EUR for other EU orders.
 *
 * method "getFulfillmentPreview": calls getFulfillmentPreview for an instant estimate.
 * Available immediately but may differ from the actual charged amount.
 *
 * Tries EU endpoint first, then FE (Japan/AU/SG) as fallback.
 * Required roles: Amazon Fulfillment (both methods) + Finance and Accounting (FinancesAPI only).
 *
 * @customfunction
 * @param {"FinancesAPI"|"getFulfillmentPreview"} method "FinancesAPI" = actual settled fee (GBP/EUR, available days after shipment). "getFulfillmentPreview" = instant estimate (may differ from actual).
 * @param {string} orderId The sellerFulfillmentOrderId of the MCF order (e.g. value in col Q).
 * @param {string} [sentDate] Optional yyyy-mm-dd sent date from col P. Skips the fulfillment order lookup when provided.
 * @return {number} Fee amount in the order's marketplace currency (GBP for UK, EUR for EU).
 */
function MCFFee(orderIdOrMethod, sentDateOrOrderId, legacySentDate) {
  // Supports both calling conventions:
  //   New (simple):  =MCFFee(Q14)          or  =MCFFee(Q14, P14)
  //   Old (verbose): =MCFFee("FinancesAPI", Q14, P14)
  // Always uses FinancesAPI. getFulfillmentPreview still works via old convention.
  var METHODS = ['FinancesAPI', 'getFulfillmentPreview'];
  var method, orderId, sentDate;
  if (METHODS.indexOf(String(orderIdOrMethod || '').trim()) >= 0) {
    method   = String(orderIdOrMethod).trim();
    orderId  = sentDateOrOrderId;
    sentDate = legacySentDate;
  } else {
    method   = 'FinancesAPI';
    orderId  = orderIdOrMethod;
    sentDate = sentDateOrOrderId;
  }
  if (!orderId) return '';
  var dateKey = sentDate ? '_' + String(sentDate).trim() : '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE_' + method + '_' + String(orderId) + dateKey;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  try {
    var fee = '';
    // maxPages: 3 when sentDate given (tight 60-day window), 2 without (wide 180-day window).
    // Keeps each formula call well under GAS's 30-second custom-function limit.
    var pages = sentDate ? 3 : 2;
    if (method === 'FinancesAPI') {
      fee = _fetchMcfFeeFinancesApi(String(orderId), 'EU', sentDate, pages);
      if (fee === '') fee = _fetchMcfFeeFinancesApi(String(orderId), 'FE', sentDate, pages);
    } else {
      try { fee = _fetchMcfFeePreview(String(orderId), 'EU'); } catch(e) {}
      if (fee === '') { try { fee = _fetchMcfFeePreview(String(orderId), 'FE'); } catch(e) {} }
    }
    cache.put(key, fee !== '' ? String(fee) : '__EMPTY__', fee !== '' ? 21600 : 600);
    return fee !== '' ? parseFloat(fee) : '';
  } catch (e) {
    if (_isRateLimit429(e)) { cache.put(key, '__EMPTY__', 90); }
    return '';
  }
}

/**
 * Returns the MCF fulfillment fee for a Japan / AU / SG order.
 * Same as MCFFee but tries the FE (Far East) endpoint first.
 *
 * @customfunction
 * @param {"FinancesAPI"|"getFulfillmentPreview"} method "FinancesAPI" = actual settled fee (available days after shipment). "getFulfillmentPreview" = instant estimate (may differ from actual).
 * @param {string} orderId The sellerFulfillmentOrderId of the MCF order.
 * @param {string} [sentDate] Optional yyyy-mm-dd sent date from col P. Skips the fulfillment order lookup when provided.
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee_JP(orderIdOrMethod, sentDateOrOrderId, legacySentDate) {
  var METHODS = ['FinancesAPI', 'getFulfillmentPreview'];
  var method, orderId, sentDate;
  if (METHODS.indexOf(String(orderIdOrMethod || '').trim()) >= 0) {
    method   = String(orderIdOrMethod).trim();
    orderId  = sentDateOrOrderId;
    sentDate = legacySentDate;
  } else {
    method   = 'FinancesAPI';
    orderId  = orderIdOrMethod;
    sentDate = sentDateOrOrderId;
  }
  if (!orderId) return '';
  var dateKey = sentDate ? '_' + String(sentDate).trim() : '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE_JP_' + method + '_' + String(orderId) + dateKey;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  try {
    var fee = '';
    var pages = sentDate ? 3 : 2;
    if (method === 'FinancesAPI') {
      fee = _fetchMcfFeeFinancesApi(String(orderId), 'FE', sentDate, pages);
      if (fee === '') fee = _fetchMcfFeeFinancesApi(String(orderId), 'EU', sentDate, pages);
    } else {
      try { fee = _fetchMcfFeePreview(String(orderId), 'FE'); } catch(e) {}
      if (fee === '') { try { fee = _fetchMcfFeePreview(String(orderId), 'EU'); } catch(e) {} }
    }
    cache.put(key, fee !== '' ? String(fee) : '__EMPTY__', fee !== '' ? 21600 : 600);
    return fee !== '' ? parseFloat(fee) : '';
  } catch (e) {
    if (_isRateLimit429(e)) { cache.put(key, '__EMPTY__', 90); }
    return '';
  }
}

/**
 * Finances API method — actual settled MCF fee.
 * Searches ShipmentEventList for SellerOrderId === orderId and sums FBA/fulfillment fees.
 * Fees are stored as negative values in the Finances API; returns Math.abs(total).
 * Returns '' if the order has not yet settled.
 */
/**
 * Safely parses a date value that may come from a Sheets cell.
 * Handles: Date objects, ISO strings (yyyy-mm-dd), YYMMDD strings ("250325" → 2025-03-25),
 * and other date-like strings.  Avoids V8's misparse of 6-digit strings as a year.
 */
function _parseCellDate(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val);
  var s = String(val).trim();
  // YYMMDD: exactly 6 digits → prepend "20" to get yyyy-mm-dd
  if (/^\d{6}$/.test(s)) {
    return new Date('20' + s.slice(0, 2) + '-' + s.slice(2, 4) + '-' + s.slice(4, 6));
  }
  return new Date(s);
}

function _fetchMcfFeeFinancesApi(orderId, ep, sentDate, maxPages) {
  var postedAfter, postedBefore;
  // skipFallback: true when called from formula (no sentDate) — avoids extra getFulfillmentOrderRaw call
  var skipDisplayableFallback = false;

  if (sentDate) {
    // Use the caller-supplied sent date (P col) — skip getFulfillmentOrderRaw entirely
    postedAfter  = _parseCellDate(sentDate);
    postedBefore = new Date(postedAfter);
    postedBefore.setDate(postedBefore.getDate() + 60);
  } else {
    // No sentDate: search the last 180 days without calling getFulfillmentOrderRaw.
    // getFulfillmentOrderRaw is slow (~2-3 s per call); skipping it prevents formula
    // cells from timing out when many run concurrently (GAS 30 s limit).
    var _now2 = new Date();
    postedAfter  = new Date(_now2.getTime() - 180 * 24 * 3600 * 1000);
    postedBefore = new Date(_now2.getTime() - 5 * 60 * 1000);
    skipDisplayableFallback = true; // skip the extra API call in the fallback below
  }

  var _now = new Date(Date.now() - 5 * 60 * 1000); // 5-min buffer for GAS-Amazon clock drift
  if (postedBefore > _now) postedBefore = _now; // cap — API rejects dates in the future

  // Collect all shipment events once — reused for primary search and displayableOrderId fallback
  var shipments = _collectShipmentEvents(ep, postedAfter, postedBefore, maxPages || 5);

  // Primary: match by sellerFulfillmentOrderId
  var fee = _sumMcfFeeFromShipments(shipments, orderId);
  if (fee !== '') return fee;

  // GCX alias: Q col stores N, Finances API may record as N+1 (see _gcxNumAlias)
  var gcxAlias = _gcxNumAlias(String(orderId), 1);
  if (gcxAlias) {
    fee = _sumMcfFeeFromShipments(shipments, gcxAlias);
    if (fee !== '') return fee;
  }

  // Fallback: some MCF orders settle in the Finances API under displayableOrderId
  // Only run when sentDate was provided (skipDisplayableFallback = false) — calling
  // getFulfillmentOrderRaw here without sentDate would cause the formula timeout again.
  if (!skipDisplayableFallback) {
    try {
      var foResult = getFulfillmentOrderRaw(orderId, ep);
      var foCache  = foResult.fulfillmentOrder || {};
      var displayableId = (foCache.displayableOrderId || '').trim();
      if (displayableId && displayableId !== orderId) {
        var fee2 = _sumMcfFeeFromShipments(shipments, displayableId);
        if (fee2 !== '') return fee2;
      }
    } catch (e) { /* fallback failed — order not yet settled */ }
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
 * Returns a GCX order ID with its trailing number shifted by delta.
 * Works on IDs matching GCX-XX-YYMMDD-N (single dash, non-negative number).
 * Returns null for double-dash or out-of-range IDs.
 *
 * Why: The Q column auto-generates IDs one step before the order is submitted to
 * Amazon's FBA Outbound API. Amazon records the submitted ID (N) as SellerOrderId
 * in the Finances API, while the sheet stores N-1. So we index each feeMap entry
 * under both the API ID (N) and the Q-column ID (N-1) for the lookup to succeed.
 */
function _gcxNumAlias(id, delta) {
  var m = /^(GCX-[A-Z]+-\d{6}-)(\d+)$/.exec(String(id));
  if (!m) return null;
  var newNum = parseInt(m[2], 10) + delta;
  if (newNum < 0) return null;
  var ns  = String(newNum);
  var pad = m[2].length;
  while (ns.length < pad) ns = '0' + ns;
  return m[1] + ns;
}

/**
 * Fetches ALL ShipmentEvent financial events for a single date window on one endpoint.
 * Returns a plain object mapping SellerOrderId → fee (absolute value).
 * Used by backfillMCFFees() to build a bulk fee map instead of one call per order.
 */
function _buildFeeMapForWindow(ep, postedAfter, postedBefore, maxPages) {
  var feeMap    = {};
  var nextToken = null;
  maxPages = maxPages || 20;
  var page  = 0;

  do {
    var qs = 'PostedAfter='   + encodeURIComponent(postedAfter.toISOString()) +
             '&PostedBefore=' + encodeURIComponent(postedBefore.toISOString()) +
             '&MaxResultsPerPage=50';  // 50 (not 100) to halve response size and stay under GAS bandwidth quota
    if (nextToken) qs += '&NextToken=' + encodeURIComponent(nextToken);

    var res       = spapiFetchWithRetry('GET', '/finances/v0/financialEvents', { queryString: qs, endpoint: ep }, 3, 5000);
    var payload   = res.payload || res;
    nextToken     = payload.NextToken || null;
    var shipments = (payload.FinancialEvents || {}).ShipmentEventList || [];

    for (var i = 0; i < shipments.length; i++) {
      var ev  = shipments[i];
      var sid = String(ev.SellerOrderId || '').trim();
      if (!sid) continue;

      // For GCX MCF orders: sum ALL fees regardless of type (fee type names differ from marketplace).
      // For other orders: apply the standard MCF-fee-type filter.
      var isGcx = sid.toUpperCase().indexOf('GCX') === 0;
      var total = 0;

      (ev.ShipmentFeeList || []).forEach(function(f) {
        if (isGcx || _isMcfFeeType(f.FeeType))
          total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentItemList || []).forEach(function(item) {
        (item.ItemFeeList || []).forEach(function(f) {
          if (isGcx || _isMcfFeeType(f.FeeType))
            total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
        });
      });

      if (total !== 0) {
        var abs = Math.abs(total);
        feeMap[sid] = abs;
        // Index under Q-column ID (N-1) so the backfill lookup succeeds even when
        // the sheet stores one less than the sellerFulfillmentOrderId Amazon recorded.
        if (isGcx) {
          var alias = _gcxNumAlias(sid, -1);
          if (alias && !feeMap[alias]) feeMap[alias] = abs;
        }
      }
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
             '&MaxResultsPerPage=50';
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
 * Finds the first ShipmentEvent matching targetOrderId and sums its MCF fee lines.
 * Returns Math.abs(total) on match, '' if not found.
 */
function _sumMcfFeeFromShipments(shipments, targetOrderId) {
  var target = String(targetOrderId).trim();
  var isGcx  = target.toUpperCase().indexOf('GCX') === 0;
  for (var i = 0; i < shipments.length; i++) {
    var ev = shipments[i];
    if (String(ev.SellerOrderId || '').trim() !== target) continue;
    var total = 0;
    (ev.ShipmentFeeList || []).forEach(function(f) {
      // GCX MCF orders: sum ALL fee types (same logic as _buildFeeMapForWindow)
      if (isGcx || _isMcfFeeType(f.FeeType))
        total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
    });
    (ev.ShipmentItemList || []).forEach(function(item) {
      (item.ItemFeeList || []).forEach(function(f) {
        if (isGcx || _isMcfFeeType(f.FeeType))
          total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
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

    if (sentDate) {
      postedAfter  = _parseCellDate(sentDate);
      postedBefore = new Date(postedAfter);
      postedBefore.setDate(postedBefore.getDate() + 90);
      dateSource = 'sentDate (P col): ' + String(sentDate).trim();
    } else {
      // Try FBA Outbound API to get sent date. Old/archived orders return 400 here —
      // in that case, pass sentDate (col P) as the 2nd arg: =MCFFeeDebug(Q3, P3)
      var foErr = null;
      var result;
      try { result = getFulfillmentOrderRaw(String(orderId), 'EU'); } catch (e) { foErr = e; }
      if (foErr || !result) {
        return [['FBA Outbound API cannot find this order — pass sentDate as 2nd arg: =MCFFeeDebug(' + String(orderId) + ', P_col)'],
                ['Error: ' + (foErr ? (foErr.message || foErr) : 'no result')]];
      }
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
      var match = sid.trim() === String(orderId).trim()         ? 'YES'
                : (displayableId && sid.trim() === displayableId) ? 'YES (displayableOrderId)'
                : 'no';
      var feeTypes = [], total = 0;
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
      rows.push([sid, orderId, match, feeTypes.join(', '), Math.abs(total)]);
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

/**
 * Bulk-populates MCFFee() cells on the active sheet.
 * Run from the Apps Script editor (Run ▶ backfillMCFFees) instead of waiting for
 * 280+ formula cells to evaluate simultaneously — that floods GAS's execution queue.
 *
 * What it does:
 *   1. Auto-detects the column containing MCFFee() formulas.
 *   2. Groups unfilled rows by monthly Finances API window (one bulk fetch per month).
 *   3. Writes static fee values to the cells (replaces the formula — fees never change).
 *   4. Also stores results in CacheService so any remaining formula cells return instantly.
 *
 * After running, the 280 cells will show static values and never recalculate again.
 * For newly added rows, run backfillMCFFees() again — it skips rows that already have a value.
 */
function backfillMCFFees_legacy() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) { Logger.log('No data rows.'); return; }

  var allFormulas = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  var allValues   = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ── Detect fee column (first column with MCFFee( in rows 2–11) ──────────────
  var feeColIdx = -1;
  outerLoop:
  for (var c = 0; c < lastCol; c++) {
    for (var r = 1; r < Math.min(11, lastRow); r++) {
      if ((allFormulas[r][c] || '').toUpperCase().indexOf('MCFFEE(') >= 0) {
        feeColIdx = c; break outerLoop;
      }
    }
  }
  if (feeColIdx < 0) { Logger.log('MCFFee() formula column not found.'); return; }

  // Columns P=15(idx) and Q=16(idx) are hardcoded to match the sheet layout.
  var sentDateColIdx = 15; // P
  var orderIdColIdx  = 16; // Q

  Logger.log('feeCol=' + _bfColLetter(feeColIdx + 1) +
             '  orderIdCol=Q  sentDateCol=P  sheet=' + sheet.getName());

  // ── Collect rows that still need a fee ──────────────────────────────────────
  var orders = [];
  for (var r = 1; r < lastRow; r++) {
    var orderId  = String(allValues[r][orderIdColIdx] || '').trim();
    if (!orderId) continue;
    var existing = allValues[r][feeColIdx];
    if (typeof existing === 'number' && existing > 0) continue;
    orders.push({ row: r + 1, orderId: orderId, sentDate: allValues[r][sentDateColIdx] });
  }
  Logger.log('Rows to fill: ' + orders.length);
  if (!orders.length) { Logger.log('Nothing to do.'); return; }

  // ── Group by YYYY-MM of sentDate ─────────────────────────────────────────────
  var byMonth = {}, noDate = [];
  orders.forEach(function(o) {
    var d = _parseCellDate(o.sentDate);
    if (!d || isNaN(d.getTime())) { noDate.push(o); return; }
    var mk = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!byMonth[mk]) byMonth[mk] = [];
    byMonth[mk].push(o);
  });
  if (noDate.length) Logger.log('Skipped (no sentDate): ' + noDate.length);

  // ── Fetch fee map per monthly window, write results ──────────────────────────
  var cache   = CacheService.getScriptCache();
  var fetched = 0, notSettled = 0;
  var monthKeys = Object.keys(byMonth).sort();

  for (var mi = 0; mi < monthKeys.length; mi++) {
    var mk    = monthKeys[mi];
    var parts = mk.split('-');
    var yr    = parseInt(parts[0]), mo = parseInt(parts[1]);

    // Window: month start → month end + 45 days (catches late settlements)
    var after  = new Date(yr, mo - 1, 1);
    var before = new Date(yr, mo,     1);
    before.setDate(before.getDate() + 45);
    var now = new Date(Date.now() - 2 * 60 * 1000); // 2-min buffer: API rejects future dates
    if (before > now) before = now;

    Logger.log('Month ' + (mi + 1) + '/' + monthKeys.length + ': ' + mk +
               ' (' + byMonth[mk].length + ' orders)');

    // Build a fee map for this window across both endpoints (maxPages=20 covers ~2000 events/month)
    var feeMap = {};
    ['EU', 'FE'].forEach(function(ep) {
      try {
        var m = _buildFeeMapForWindow(ep, after, before, 20);
        var gcxFound = Object.keys(m).filter(function(k) { return k.toUpperCase().indexOf('GCX') === 0; });
        if (gcxFound.length) Logger.log('  [' + ep + '] GCX entries in feeMap (' + gcxFound.length + '): ' + gcxFound.slice(0, 10).join(', ') + (gcxFound.length > 10 ? ' …' : ''));
        Object.keys(m).forEach(function(k) { if (!feeMap[k]) feeMap[k] = m[k]; });
      } catch (e) { Logger.log('  ' + ep + ' error: ' + e.message); }
    });

    // Write matches back to sheet and cache
    byMonth[mk].forEach(function(o) {
      var fee = feeMap[o.orderId];
      if (fee != null) {
        sheet.getRange(o.row, feeColIdx + 1).setValue(fee);
        var ck = 'MCFFEE_FinancesAPI_' + o.orderId +
                 (o.sentDate ? '_' + String(o.sentDate).trim() : '');
        cache.put(ck, String(fee), 21600);
        fetched++;
      } else {
        notSettled++;
      }
    });

    SpreadsheetApp.flush();
    if (mi < monthKeys.length - 1) Utilities.sleep(500);
  }

  Logger.log('backfillMCFFees done — written: ' + fetched + ', not yet settled: ' + notSettled);
}

function _bfColLetter(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * Diagnostic: logs what the Finances API actually returns for the first N unfilled rows.
 * Run from Apps Script editor to diagnose why backfillMCFFees writes 0 fees.
 * Checks:
 *   1. Raw ShipmentEvent count + sample SellerOrderIds from the API
 *   2. Whether any SellerOrderId matches the Q-col order ID
 *   3. What FeeTypes exist on matched events
 *   4. Whether ServiceFeeEventList (alternative MCF billing) has matching entries
 */
function diagnoseMCFFee() {
  var SAMPLE_ROWS  = 5;   // how many rows to inspect
  var sheet        = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow      = sheet.getLastRow();
  var lastCol      = sheet.getLastColumn();
  var allValues    = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Log first 4 rows entirely so we can see the full header structure
  Logger.log('=== ROWS 1–4 (all non-empty cells) ===');
  for (var ri = 0; ri < Math.min(4, allValues.length); ri++) {
    allValues[ri].forEach(function(v, ci) {
      if (v !== '' && v !== null && v !== undefined) {
        Logger.log('  row' + (ri+1) + ' col ' + _bfColLetter(ci+1) + ': ' + String(v).substring(0, 60));
      }
    });
  }

  // Scan rows 4–10 for Amazon-format order IDs (NNN-NNNNNNN-NNNNNNN)
  Logger.log('=== Cols with Amazon order ID format (rows 4–10) ===');
  for (var ri2 = 3; ri2 < Math.min(10, allValues.length); ri2++) {
    allValues[ri2].forEach(function(v, ci) {
      if (/^\d{3}-\d{7}-\d{7}$/.test(String(v).trim())) {
        Logger.log('  row' + (ri2+1) + ' col ' + _bfColLetter(ci+1) + ': ' + v);
      }
    });
  }

  // Try listFulfillmentOrders to see if Amazon has the sellerFulfillmentOrderId↔displayableOrderId mapping
  Logger.log('=== listFulfillmentOrders sample (EU, from 2025-03-01) ===');
  try {
    var lfRes = spapiFetchWithRetry('GET', '/fba/outbound/2020-07-01/fulfillmentOrders',
      { queryString: 'queryStartDate=' + encodeURIComponent('2025-03-01T00:00:00Z'), endpoint: 'EU' }, 2, 3000);
    var lfPayload = lfRes.payload || lfRes;
    var lfOrders  = lfPayload.fulfillmentOrders || [];
    Logger.log('  count returned: ' + lfOrders.length);
    lfOrders.slice(0, 5).forEach(function(o) {
      Logger.log('  sellerFulfillmentOrderId=' + o.sellerFulfillmentOrderId +
                 '  displayableOrderId=' + o.displayableOrderId);
    });
  } catch (e) {
    Logger.log('  listFulfillmentOrders error: ' + e.message);
  }

  var sentDateColIdx = 15; // P
  var orderIdColIdx  = 16; // Q

  // Collect first SAMPLE_ROWS with both orderId and sentDate
  var samples = [];
  for (var r = 1; r < lastRow && samples.length < SAMPLE_ROWS; r++) {
    var orderId  = String(allValues[r][orderIdColIdx] || '').trim();
    var sentDate = allValues[r][sentDateColIdx];
    if (orderId && sentDate) samples.push({ row: r + 1, orderId: orderId, sentDate: sentDate });
  }

  Logger.log('=== diagnoseMCFFee: inspecting ' + samples.length + ' rows ===');

  samples.forEach(function(s) {
    Logger.log('\n--- Row ' + s.row + ' | orderId=' + s.orderId + ' | sentDate=' + s.sentDate + ' ---');

    var after = _parseCellDate(s.sentDate);
    if (!after || isNaN(after.getTime())) { Logger.log('  SKIP: bad sentDate'); return; }
    var before = new Date(after); before.setDate(before.getDate() + 60);
    var cap = new Date(Date.now() - 5 * 60 * 1000);
    if (before > cap) before = cap;

    Logger.log('  window: ' + after.toISOString() + ' → ' + before.toISOString());

    ['EU', 'FE'].forEach(function(ep) {
      try {
        // ── ShipmentEventList ───────────────────────────────────────────────
        var qs = 'PostedAfter='   + encodeURIComponent(after.toISOString()) +
                 '&PostedBefore=' + encodeURIComponent(before.toISOString()) +
                 '&MaxResultsPerPage=100';
        var res      = spapiFetchWithRetry('GET', '/finances/v0/financialEvents',
                         { queryString: qs, endpoint: ep }, 2, 3000);
        var payload  = res.payload || res;
        var shipments = (payload.FinancialEvents || {}).ShipmentEventList || [];

        Logger.log('  [' + ep + '] ShipmentEventList count: ' + shipments.length);
        if (shipments.length) {
          // Log first 5 SellerOrderIds so we can see the ID format
          var sids = shipments.slice(0, 5).map(function(e) { return e.SellerOrderId || '(none)'; });
          Logger.log('  [' + ep + '] Sample SellerOrderIds: ' + sids.join(' | '));

          // Log ALL GCX-prefixed SellerOrderIds found (our MCF format)
          var gcxIds = shipments.filter(function(e) {
            return String(e.SellerOrderId || '').toUpperCase().indexOf('GCX') === 0;
          }).map(function(e) { return e.SellerOrderId; });
          if (gcxIds.length) Logger.log('  [' + ep + '] GCX-format SellerOrderIds found: ' + gcxIds.join(' | '));

          // Check if our orderId appears (exact)
          var match = shipments.filter(function(e) {
            return String(e.SellerOrderId || '').trim() === s.orderId;
          });
          if (match.length) {
            var feeTypes = (match[0].ShipmentFeeList || []).map(function(f) { return f.FeeType; });
            var itemFeeTypes = [];
            (match[0].ShipmentItemList || []).forEach(function(item) {
              (item.ItemFeeList || []).forEach(function(f) { itemFeeTypes.push(f.FeeType); });
            });
            Logger.log('  [' + ep + '] MATCH FOUND — ShipmentFeeTypes: ' + feeTypes.join(', ') +
                       '  ItemFeeTypes: ' + itemFeeTypes.join(', '));
          } else {
            Logger.log('  [' + ep + '] No match for "' + s.orderId + '"');
          }
        }

        // ── ServiceFeeEventList (alternative MCF billing path) ──────────────
        var svcFees = (payload.FinancialEvents || {}).ServiceFeeEventList || [];
        Logger.log('  [' + ep + '] ServiceFeeEventList count: ' + svcFees.length);
        if (svcFees.length) {
          var svcSample = svcFees.slice(0, 3).map(function(e) {
            return (e.SellerOrderId || e.AmazonOrderId || e.FeeReason || '?');
          });
          Logger.log('  [' + ep + '] Sample ServiceFee keys: ' + svcSample.join(' | '));
        }

        // ── List ALL event list keys present in this response ───────────────
        var fe = payload.FinancialEvents || {};
        var presentKeys = Object.keys(fe).filter(function(k) {
          return Array.isArray(fe[k]) && fe[k].length > 0;
        });
        Logger.log('  [' + ep + '] Non-empty event lists: ' + (presentKeys.join(', ') || '(none)'));

      } catch (e) {
        Logger.log('  [' + ep + '] ERROR: ' + e.message);
      }
    });
  });

  Logger.log('\n=== diagnoseMCFFee done ===');
}

/**
 * Broad scan: queries Finances API over a 3-month window and searches EVERY
 * event list for any string value that contains "GCX" (our MCF order prefix).
 * Also logs counts for all non-empty event lists.
 *
 * Run from the Apps Script editor to find which event list MCF fees appear in.
 * Adjust SCAN_START / SCAN_END if you want a different window.
 */
function scanFinancesForMCF() {
  // Split into 90-day chunks to stay within the 180-day API limit.
  var windows = [
    { start: '2025-01-01T00:00:00Z', end: '2025-04-01T00:00:00Z' },
    { start: '2025-04-01T00:00:00Z', end: '2025-07-01T00:00:00Z' }
  ];

  ['EU', 'FE'].forEach(function(ep) {
    var listCounts = {};
    var gcxHits    = [];
    var totalPages = 0;

    windows.forEach(function(win) {
      Logger.log('\n══════ ' + ep + ' | ' + win.start + ' → ' + win.end + ' ══════');

      var nextToken = null, page = 0, maxPages = 30;
      do {
        var qs = 'PostedAfter='   + encodeURIComponent(win.start) +
                 '&PostedBefore=' + encodeURIComponent(win.end)   +
                 '&MaxResultsPerPage=100';
        if (nextToken) qs += '&NextToken=' + encodeURIComponent(nextToken);

        try {
          var res     = spapiFetchWithRetry('GET', '/finances/v0/financialEvents',
                          { queryString: qs, endpoint: ep }, 3, 5000);
          var payload = res.payload || res;
          nextToken   = payload.NextToken || null;
          var fe      = payload.FinancialEvents || {};

          Object.keys(fe).forEach(function(listName) {
            var events = fe[listName];
            if (!Array.isArray(events) || !events.length) return;
            listCounts[listName] = (listCounts[listName] || 0) + events.length;
            events.forEach(function(ev) { _deepScanForGcx(ev, listName, gcxHits); });
          });

          page++;
        } catch (e) {
          Logger.log('  page ' + page + ' error: ' + e.message);
          break;
        }
      } while (nextToken && page < maxPages);

      Logger.log('  window pages fetched: ' + page);
      totalPages += page;
    });

    // Per-endpoint summary
    Logger.log('\n── ' + ep + ' SUMMARY ──');
    Logger.log('  Total pages: ' + totalPages);
    Logger.log('  Non-empty event lists:');
    Object.keys(listCounts).sort().forEach(function(k) {
      Logger.log('    ' + k + ': ' + listCounts[k] + ' events');
    });
    if (gcxHits.length) {
      Logger.log('  GCX hits (' + gcxHits.length + '):');
      gcxHits.slice(0, 30).forEach(function(h) {
        Logger.log('    [' + h.list + '] ' + h.key + ' = ' + h.value);
      });
    } else {
      Logger.log('  *** NO GCX hits in any event list ***');
    }
  });

  Logger.log('\n=== scanFinancesForMCF done ===');
}

/** Recursively walks an object and records any string containing "GCX". */
function _deepScanForGcx(obj, listName, hits) {
  if (!obj || typeof obj !== 'object') return;
  Object.keys(obj).forEach(function(k) {
    var v = obj[k];
    if (typeof v === 'string' && v.toUpperCase().indexOf('GCX') >= 0) {
      hits.push({ list: listName, key: k, value: v.substring(0, 80) });
    } else if (Array.isArray(v)) {
      v.forEach(function(item) { _deepScanForGcx(item, listName, hits); });
    } else if (v && typeof v === 'object') {
      _deepScanForGcx(v, listName, hits);
    }
  });
}

/**
 * Logs the first 20 non-empty Q-column (col 17) values exactly as stored,
 * plus their P-column (col 16) sentDate. Run to see the exact Q-col format
 * so we can compare it to SellerOrderId values in the Finances API.
 */
function sampleQcol() {
  var sheet    = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow  = sheet.getLastRow();
  var startRow = 4; // data starts at row 4 (row 3 is header)
  var maxRows  = Math.min(lastRow - startRow + 1, 200);
  var data     = sheet.getRange(startRow, 1, maxRows, 17).getValues(); // A:Q

  Logger.log('=== sampleQcol — first 20 non-empty Q (col 17) values ===');
  var shown = 0;
  for (var i = 0; i < data.length && shown < 20; i++) {
    var q = data[i][16]; // Q col index 16
    var p = data[i][15]; // P col index 15
    if (q === '' || q === null || q === undefined) continue;
    Logger.log('  row ' + (startRow + i) + '  P="' + p + '"  Q="' + String(q) + '"');
    shown++;
  }
  Logger.log('=== done (' + shown + ' rows shown) ===');
}
