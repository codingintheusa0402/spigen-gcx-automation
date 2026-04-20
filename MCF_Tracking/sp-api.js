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

/***** ========= SHEET FUNCTIONS ========= *****/
function AMZTK(orderId) {
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'AMZTK_' + orderId;
  var cached = cache.get(key);

  if (cached !== null) return cached;  // return cached value immediately

  try {
    var tracks = _tracksWithFallbacks(String(orderId), ['EU', 'FE']);
    var tn = (tracks && tracks.length && (tracks[0].trackingNumber || '').trim())
      ? tracks[0].trackingNumber
      : '';

    cache.put(key, tn, 600); // cache 10 minutes
    return tn;

  } catch (err) {
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) return '';
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

    cache.put(key, tn, 600); // 10 minutes
    return tn;

  } catch (err) {
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) return '';
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
 * Custom formula: =FBAFee(orderId)
 * Returns total Expedited MCF fulfillment fee for an existing order.
 * Tries EU first, then FE — same fallback pattern as AMZTK.
 */
function FBAFee(orderId) {
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'FBAFEE_' + String(orderId);
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['EU', 'FE'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = _fetchMcfFee(String(orderId), endpoints[i]);
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), 600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 600); return ''; }
      throw err;
    }
  }

  if (lastErr) {
    if (_isNoOrderInfoError(lastErr) || _isUnauthorizedError(lastErr)) {
      cache.put(key, '__EMPTY__', 600);
      return '';
    }
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}

/**
 * Custom formula: =FBAFee_JP(orderId)
 * Same as FBAFee but tries FE first (Japan/AU/SG orders).
 */
function FBAFee_JP(orderId) {
  if (!orderId) return '';

  var cache = CacheService.getScriptCache();
  var key = 'FBAFEE_JP_' + String(orderId);
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['FE', 'EU'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = _fetchMcfFee(String(orderId), endpoints[i]);
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), 600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 600); return ''; }
      throw err;
    }
  }

  if (lastErr) {
    if (_isNoOrderInfoError(lastErr) || _isUnauthorizedError(lastErr)) {
      cache.put(key, '__EMPTY__', 600);
      return '';
    }
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}

function _fetchMcfFee(orderId, ep) {
  // Step 1: pull existing order for address + items
  var result = getFulfillmentOrderRaw(orderId, ep);
  var fo     = result.fulfillmentOrder || {};
  var items  = result.fulfillmentOrderItems || [];
  if (!items.length || !fo.destinationAddress) return '';

  // Step 2: build getFulfillmentPreview request
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

  // Step 3: call getFulfillmentPreview
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

  // Step 4: sum all fee components (value is a string in SP-API response)
  var total = 0;
  for (var j = 0; j < preview.estimatedFees.length; j++) {
    total += parseFloat(preview.estimatedFees[j].amount.value || 0);
  }
  return total;
}
