
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
// In-process token store — survives for the lifetime of one script execution.
// Prevents repeated LWA fetches (and Amazon rate-limit 401s) within a single run.
var _lwaTokenInProcess = {};

function _warmLwaTokens() {
  getLwaAccessToken('EU');
  getLwaAccessToken('JP');
}

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
  // 1. In-process store (reliable within one execution)
  if (_lwaTokenInProcess[prof]) return _lwaTokenInProcess[prof];

  // 2. ScriptCache (shared across executions, best-effort)
  var cached = cache.get(cacheKey);
  if (cached) { _lwaTokenInProcess[prof] = cached; return cached; }

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
  _lwaTokenInProcess[prof] = body.access_token; // in-process store
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
  // Trim: a stray trailing \r/newline (e.g. pasted into the cell) makes Amazon reject the URL with a bare
  // HTML 400 ("<title>x</title>") instead of a JSON error — looked like a JP credential problem, wasn't.
  var path = '/fba/outbound/2020-07-01/fulfillmentOrders/' + encodeURIComponent(String(sellerFulfillmentOrderId).trim());
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
     msg.indexOf('Unable to get order info') >= 0 ||
     msg.indexOf('Received 500 response') >= 0)  // EU FBA Outbound internal error
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
  orderId = String(orderId || '').trim();
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
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) {
      cache.put(key, '', 600); // no tracking yet — retry in 10 min
      return '';
    }
    // 429 / transient errors: do NOT cache — let the next recalculation retry.
    return 'EU ERR: ' + (err && err.message ? err.message : err);
  }
}

function AMZTK_JP(orderId) {
  orderId = String(orderId || '').trim();
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
    if (_isUnauthorizedError(err) || _isNoOrderInfoError(err)) {
      cache.put(key, '', 600); // no tracking yet — retry in 10 min
      return '';
    }
    // 429 / transient errors: do NOT cache — let the next recalculation retry.
    return 'JP ERR: ' + (err && err.message ? err.message : err);
  }
}

/**
 * Step-by-step AMZTK debug. Run from editor: set ORDER_ID below, then Run → debugAMZTK.
 * Also clears the cache for that order so you see a live API result, not a cached blank.
 * View → Logs after running.
 */
function debugAMZTK() {
  var ORDER_ID = 'GCX-FR-260529-2323'; // ← change to the order you want to test

  var cache = CacheService.getScriptCache();

  // 1. Show what's currently in cache
  var cachedEU = cache.get('AMZTK_' + ORDER_ID);
  var cachedJP = cache.get('AMZTK_JP_' + ORDER_ID);
  Logger.log('Cache AMZTK_'    + ORDER_ID + ' = ' + JSON.stringify(cachedEU) + ' (null = no entry)');
  Logger.log('Cache AMZTK_JP_' + ORDER_ID + ' = ' + JSON.stringify(cachedJP));

  // 2. Clear cache so we get a fresh API call
  cache.removeAll(['AMZTK_' + ORDER_ID, 'AMZTK_JP_' + ORDER_ID]);
  Logger.log('Cache cleared for ' + ORDER_ID);

  // 3. Try EU endpoint directly
  Logger.log('--- EU endpoint ---');
  try {
    var foEU = getFulfillmentOrderRaw(ORDER_ID, 'EU');
    var shipmentsEU = (foEU && foEU.fulfillmentShipments) || [];
    Logger.log('getFulfillmentOrderRaw EU: OK — ' + shipmentsEU.length + ' shipment(s)');
    shipmentsEU.forEach(function(sh, i) {
      var pkgs = (sh.fulfillmentShipmentPackage && sh.fulfillmentShipmentPackage.length)
        ? sh.fulfillmentShipmentPackage : (sh.packages || []);
      Logger.log('  Shipment ' + i + ': status=' + sh.fulfillmentShipmentStatus + ', ' + pkgs.length + ' package(s)');
      pkgs.forEach(function(p, j) {
        Logger.log('    Package ' + j + ': trackingNumber=' + (p.trackingNumber || '(empty)') +
                   ', packageNumber=' + (p.packageNumber || '(none)'));
      });
    });
    if (!shipmentsEU.length) Logger.log('  → fulfillmentShipments is EMPTY (order not yet shipped)');
  } catch (e) {
    Logger.log('getFulfillmentOrderRaw EU: ERROR — ' + e.message);
  }

  // 4. Try FE endpoint directly
  Logger.log('--- FE endpoint ---');
  try {
    var foFE = getFulfillmentOrderRaw(ORDER_ID, 'FE');
    var shipmentsFE = (foFE && foFE.fulfillmentShipments) || [];
    Logger.log('getFulfillmentOrderRaw FE: OK — ' + shipmentsFE.length + ' shipment(s)');
  } catch (e) {
    Logger.log('getFulfillmentOrderRaw FE: ERROR — ' + e.message);
  }

  // 5. Call AMZTK() itself and show final return value
  Logger.log('--- AMZTK() result ---');
  var result = AMZTK(ORDER_ID);
  Logger.log('AMZTK(' + ORDER_ID + ') returned: ' + JSON.stringify(result));
  Logger.log('New cache value: ' + JSON.stringify(cache.get('AMZTK_' + ORDER_ID)));
}

/**
 * Clears MCFFee()/MCFFee_JP() CacheService entries for every order in the sheet.
 *
 * CacheService persists independently of script deployments — pushing a code fix via clasp
 * does NOT invalidate it. A cell that cached a wrong result (e.g. a literal 0 from the
 * pre-fix GCX fee-type filtering bug) keeps serving that stale wrong value for up to its 6h
 * TTL even after the underlying bug is fixed, because MCFFee() checks the cache before ever
 * reaching the (now-correct) computation logic. Run this once after fixing any MCFFee-related
 * bug so cells actually recompute instead of replaying stale results.
 *
 * Rebuilds cache keys using the exact same method the live formula would (reads orderId/
 * sentDate straight from the sheet, same Date-object stringification Apps Script uses when
 * passing a date-formatted cell to a custom function), so this reliably targets the same keys
 * =MCFFee(Q,P)-style formulas actually use.
 */
function clearMcfFeeCache() {
  var cache = CacheService.getScriptCache();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(BF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + BF_SHEET_NAME);

  var lastRow = sheet.getLastRow();
  if (lastRow < BF_START_ROW) return;

  var numRows   = lastRow - BF_START_ROW + 1;
  var orderIds  = sheet.getRange(BF_START_ROW, BF_COL_ORDER, numRows, 1).getValues();
  var sentDates = sheet.getRange(BF_START_ROW, BF_COL_SENT,  numRows, 1).getValues();

  var METHODS = ['FinancesAPI', 'getFulfillmentPreview'];
  var keys = [];

  for (var i = 0; i < numRows; i++) {
    var orderId = String(orderIds[i][0] || '').trim();
    if (!orderId) continue;

    var rawSentDate = sentDates[i][0]; // Date object or string, exactly as MCFFee() would receive it
    var dateKey = rawSentDate ? '_' + String(rawSentDate).trim() : '';

    METHODS.forEach(function(method) {
      keys.push('MCFFEE_'    + method + '_' + orderId + dateKey);
      keys.push('MCFFEE_'    + method + '_' + orderId);       // in case it was called without sentDate
      keys.push('MCFFEE_JP_' + method + '_' + orderId + dateKey);
      keys.push('MCFFEE_JP_' + method + '_' + orderId);
    });
  }

  var BATCH = 100;
  for (var s = 0; s < keys.length; s += BATCH) cache.removeAll(keys.slice(s, s + BATCH));
  Logger.log('clearMcfFeeCache: cleared up to ' + keys.length + ' MCFFee cache entries.');
}

/**
 * Clears AMZTK cache for all orders in the sheet so cells re-fetch live.
 * Run from editor after fixing any underlying issue.
 */
function clearAmztkCache() {
  var cache = CacheService.getScriptCache();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MCF 발송 로그');
  var keys  = ['LWA_TOKEN_EU', 'LWA_TOKEN_JP'];

  if (sheet) {
    var lastRow  = sheet.getLastRow();
    var startRow = 4;
    if (lastRow >= startRow) {
      var orderIds = sheet.getRange(startRow, 17, lastRow - startRow + 1, 1).getValues();
      for (var i = 0; i < orderIds.length; i++) {
        var id = String(orderIds[i][0] || '').trim();
        if (id) { keys.push('AMZTK_' + id); keys.push('AMZTK_JP_' + id); }
      }
    }
  }

  var BATCH = 100;
  for (var s = 0; s < keys.length; s += BATCH) cache.removeAll(keys.slice(s, s + BATCH));
  Logger.log('Cleared ' + (keys.length - 2) + ' order cache entries + LWA tokens.');
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

// Distinguishes a sentDate argument from an orderId argument by shape, not position —
// Sheets passes date-formatted cells as Date objects; a manually-typed yyyy-mm-dd string
// also matches. Order IDs (e.g. "GCX-UK-260721-2725") never match either.
function _looksLikeDate(v) {
  if (v instanceof Date) return true;
  return /^\d{4}-\d{2}-\d{2}/.test(String(v || '').trim());
}

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
 * Accepts two calling conventions (auto-detected from the first argument):
 *   Simple:  =MCFFee(Q14)  or  =MCFFee(Q14, P14)        — orderId first, FinancesAPI default
 *   Verbose: =MCFFee("FinancesAPI", Q14, P14)            — method first
 *
 * @customfunction
 * @param {string} orderIdOrMethod The sellerFulfillmentOrderId (simple form), or "FinancesAPI"/"getFulfillmentPreview" (verbose form).
 * @param {string} [sentDateOrOrderId] Sent date yyyy-mm-dd (simple form), or the orderId (verbose form).
 * @param {string} [legacySentDate] Sent date yyyy-mm-dd — verbose form only.
 * @return {number} Fee amount in the order's marketplace currency (GBP for UK, EUR for EU).
 */
function MCFFee(orderIdOrMethod, sentDateOrOrderId, legacySentDate) {
  // Supports every calling convention that has existed live in col Y across the sheet's
  // history — this project's README/commits have documented both =MCFFee(Q14, P14) and
  // =MCFFee(P14, Q14) at different points, so different row blocks may use either order:
  //   Simple:  =MCFFee(Q14)  or  =MCFFee(Q14, P14)  or  =MCFFee(P14, Q14)  — FinancesAPI default
  //   Verbose: =MCFFee("FinancesAPI", Q14, P14)                            — method first
  var METHODS = ['FinancesAPI', 'getFulfillmentPreview'];
  var method, orderId, sentDate;
  if (METHODS.indexOf(String(orderIdOrMethod || '').trim()) >= 0) {
    method   = String(orderIdOrMethod).trim();
    orderId  = sentDateOrOrderId;
    sentDate = legacySentDate;
  } else {
    method = 'FinancesAPI';
    // Detect orderId vs sentDate by shape, not position — don't trust argument order.
    if (_looksLikeDate(orderIdOrMethod) && !_looksLikeDate(sentDateOrOrderId)) {
      sentDate = orderIdOrMethod;
      orderId  = sentDateOrOrderId;
    } else {
      orderId  = orderIdOrMethod;
      sentDate = sentDateOrOrderId;
    }
  }
  if (!orderId) return '';
  var dateKey = sentDate ? '_' + String(sentDate).trim() : '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE_' + method + '_' + String(orderId) + dateKey;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['EU', 'FE'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = (method === 'FinancesAPI')
        ? _fetchMcfFeeFinancesApi(String(orderId), endpoints[i], sentDate)
        : _fetchMcfFeePreview(String(orderId), endpoints[i]);
      // Fee found → stable, cache 6h.  Not yet settled / no preview → retry in 10min.
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), fee === '' ? 600 : 21600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 21600); return ''; }
      if (_isRateLimit429(err))      { cache.put(key, '__EMPTY__', 90);    return ''; } // retry after 90s
      throw err;
    }
  }

  if (lastErr) {
    if (_isUnauthorizedError(lastErr)) { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isNoOrderInfoError(lastErr))  { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isRateLimit429(lastErr))      { cache.put(key, '__EMPTY__', 90);    return ''; } // retry after 90s
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}

/**
 * Returns the MCF fulfillment fee for a Japan / AU / SG order.
 * Same as MCFFee but tries the FE (Far East) endpoint first. Same dual calling convention.
 *
 * @customfunction
 * @param {string} orderIdOrMethod The sellerFulfillmentOrderId (simple form), or "FinancesAPI"/"getFulfillmentPreview" (verbose form).
 * @param {string} [sentDateOrOrderId] Sent date yyyy-mm-dd (simple form), or the orderId (verbose form).
 * @param {string} [legacySentDate] Sent date yyyy-mm-dd — verbose form only.
 * @return {number} Fee amount in the order's marketplace currency.
 */
function MCFFee_JP(orderIdOrMethod, sentDateOrOrderId, legacySentDate) {
  // Same calling conventions as MCFFee() (including shape-based orderId/sentDate detection
  // for the simple form) — see comment there.
  var METHODS = ['FinancesAPI', 'getFulfillmentPreview'];
  var method, orderId, sentDate;
  if (METHODS.indexOf(String(orderIdOrMethod || '').trim()) >= 0) {
    method   = String(orderIdOrMethod).trim();
    orderId  = sentDateOrOrderId;
    sentDate = legacySentDate;
  } else {
    method = 'FinancesAPI';
    if (_looksLikeDate(orderIdOrMethod) && !_looksLikeDate(sentDateOrOrderId)) {
      sentDate = orderIdOrMethod;
      orderId  = sentDateOrOrderId;
    } else {
      orderId  = orderIdOrMethod;
      sentDate = sentDateOrOrderId;
    }
  }
  if (!orderId) return '';
  var dateKey = sentDate ? '_' + String(sentDate).trim() : '';

  var cache = CacheService.getScriptCache();
  var key = 'MCFFEE_JP_' + method + '_' + String(orderId) + dateKey;
  var cached = cache.get(key);
  if (cached !== null) return cached === '__EMPTY__' ? '' : parseFloat(cached);

  var endpoints = ['FE', 'EU'];
  var lastErr = null;

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var fee = (method === 'FinancesAPI')
        ? _fetchMcfFeeFinancesApi(String(orderId), endpoints[i], sentDate)
        : _fetchMcfFeePreview(String(orderId), endpoints[i]);
      cache.put(key, fee === '' ? '__EMPTY__' : String(fee), fee === '' ? 600 : 21600);
      return fee;
    } catch (err) {
      lastErr = err;
      if (_isRetryableRegionMismatchError(err) || _isNoOrderInfoError(err)) continue;
      if (_isUnauthorizedError(err)) { cache.put(key, '__EMPTY__', 21600); return ''; }
      if (_isRateLimit429(err))      { cache.put(key, '__EMPTY__', 90);    return ''; } // retry after 90s
      throw err;
    }
  }

  if (lastErr) {
    if (_isUnauthorizedError(lastErr)) { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isNoOrderInfoError(lastErr))  { cache.put(key, '__EMPTY__', 21600); return ''; }
    if (_isRateLimit429(lastErr))      { cache.put(key, '__EMPTY__', 90);    return ''; } // retry after 90s
    return 'ERR: ' + (lastErr.message || lastErr);
  }
  return '';
}

/**
 * Finances API method — actual settled MCF fee.
 * Searches ShipmentEventList for SellerOrderId === orderId and sums FBA/fulfillment fees.
 * Fees are stored as negative values in the Finances API; returns Math.abs(total).
 * Returns '' if the order has not yet settled.
 */
function _fetchMcfFeeFinancesApi(orderId, ep, sentDate) {
  var postedAfter, postedBefore;
  var foCache = null; // cache getFulfillmentOrderRaw result to avoid a double call in fallback

  if (sentDate) {
    // Use the caller-supplied sent date (P col) — skip getFulfillmentOrderRaw entirely
    postedAfter  = new Date(String(sentDate).trim());
    postedBefore = new Date(postedAfter);
    postedBefore.setDate(postedBefore.getDate() + 60);
  } else {
    // No sentDate supplied — default to last 180 days instead of calling getFulfillmentOrderRaw.
    // getFulfillmentOrderRaw costs one FBA Outbound API call per formula cell; with many cells
    // running concurrently that exhausts GAS's daily URL-fetch bandwidth quota.
    var _ref = new Date(Date.now() - 5 * 60 * 1000);
    postedAfter  = new Date(_ref.getTime() - 180 * 24 * 3600 * 1000);
    postedBefore = new Date(_ref);
  }

  var _now = new Date(Date.now() - 5 * 60 * 1000); // 5-min buffer for GAS-Amazon clock drift
  if (postedBefore > _now) postedBefore = _now; // cap — API rejects dates in the future

  // Collect all shipment events once — reused for primary search and displayableOrderId fallback
  var shipments = _collectShipmentEvents(ep, postedAfter, postedBefore, 5);

  // Primary: match by sellerFulfillmentOrderId
  var fee = _sumMcfFeeFromShipments(shipments, orderId);
  if (fee !== '') return fee;

  // GCX alias: the Q column auto-generates its ID one step before the order is submitted
  // to Amazon's FBA Outbound API, so Amazon records SellerOrderId = N while the sheet
  // stores N-1. Try N (see _gcxNumAlias).
  var gcxAlias = _gcxNumAlias(String(orderId), 1);
  if (gcxAlias) {
    fee = _sumMcfFeeFromShipments(shipments, gcxAlias);
    if (fee !== '') return fee;
  }

  // Fallback: some MCF orders settle in the Finances API under displayableOrderId
  // (e.g. when fulfilling a linked Amazon marketplace order).
  // Only run when sentDate was provided — foCache is only populated in that path.
  // Without sentDate, getFulfillmentOrderRaw was already skipped to save bandwidth.
  if (foCache !== null) {
    try {
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
 * Works on IDs matching GCX-XX-YYMMDD-N. Returns null for non-matching or out-of-range IDs.
 *
 * Why: the Q column auto-generates its ID one step before the order is actually submitted to
 * Amazon's FBA Outbound API. Amazon records the submitted ID (N) as SellerOrderId in the
 * Finances API, while the sheet stores N-1 — so a direct match against the Q-column value
 * fails for essentially every GCX order. _fetchMcfFeeFinancesApi tries N (delta=+1) after the
 * direct match fails; _buildFeeMapForWindow indexes each entry under N-1 too (delta=-1) so
 * backfillMCFFees()'s lookup (keyed by the Q-column value) succeeds directly.
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

      // GCX-prefixed orders: sum ALL fee types — their fee lines aren't tagged FBA/FULFILLMENT
      // (see _sumMcfFeeFromShipments). Non-GCX orders keep the standard filter.
      var isGcx = sid.toUpperCase().indexOf('GCX') === 0;
      var total = 0;
      (ev.ShipmentFeeList || []).forEach(function(f) {
        if (isGcx || _isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
      (ev.ShipmentItemList || []).forEach(function(item) {
        (item.ItemFeeList || []).forEach(function(f) {
          if (isGcx || _isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
        });
      });

      if (total !== 0) {
        var abs = Math.abs(total);
        feeMap[sid] = abs;
        // Index under the Q-column ID (N-1) too, so backfillMCFFees()'s lookup — keyed by
        // the Q-column value — succeeds directly without a separate alias-retry step.
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

// Representative marketplaceId per endpoint group for settlement-report listing.
// EU settlement reports are unified across all EU marketplaces regardless of which
// EU marketplaceId is passed here (confirmed live: a DE-scoped list call returned GCX-FR/GCX-UK
// orders too) — same for FE (JP covers JP/AU/SG).
function _settlementMarketplaceId(ep) {
  return ep === 'FE' ? 'A1VC38T7YXB528' /* JP */ : 'A1PA6795UKMFR9' /* DE */;
}

// European settlement TSVs use a locale decimal comma ("7,50"); NA/other locales use a dot.
// Replacing "," with "." is safe here since these are per-line-item amounts, never
// thousands-grouped.
function _parseSettlementAmount(s) {
  var n = parseFloat(String(s || '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// Persists which settlement reportIds have already been scanned, as { reportId: createdTime },
// so successive backfillMCFFeesRecent() runs can prioritize reports they haven't seen yet instead
// of always re-scanning the same newest N (see _buildSettlementFeeMap).
function _getScannedSettlementReportIds(ep) {
  var raw = _prop('SETTLEMENT_SCANNED_' + ep, '');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

// Saves the scanned-report map, pruning entries whose report has already aged out of the 90-day
// listing window (no longer useful to remember, keeps the stored property small).
function _saveScannedSettlementReportIds(ep, map) {
  var cutoff = Date.now() - 90 * 24 * 3600 * 1000;
  var pruned = {};
  Object.keys(map).forEach(function(id) {
    var t = new Date(map[id]).getTime();
    if (t >= cutoff) pruned[id] = map[id];
  });
  PropertiesService.getScriptProperties().setProperty('SETTLEMENT_SCANNED_' + ep, JSON.stringify(pruned));
}

/**
 * Forces every settlement report in the 90-day window to be re-scanned on the next
 * backfillMCFFeesRecent() run, for both EU and FE. Only needed if a row was reset back to
 * pending (e.g. via retryZeroTransportationFees()) whose fee actually posted in a report that's
 * already marked scanned — normally _buildSettlementFeeMap()'s scanned-tracking should just be
 * left alone. Run manually from the Apps Script editor.
 */
function clearSettlementScanCache() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('SETTLEMENT_SCANNED_EU');
  props.deleteProperty('SETTLEMENT_SCANNED_FE');
  Logger.log('clearSettlementScanCache: cleared — all settlement reports in the 90-day window will be re-scanned on the next run.');
}

/***** ========= PERSISTENT FEE CACHE (fixes the "steady state = permanently stuck" bug) ========= *****/
// FIXED 2026-07-31: _buildSettlementFeeMap() used to return a fee map built ONLY from reports
// scanned in the CURRENT run. Once every report in the 90-day window had been marked "scanned"
// (which happens naturally after enough runs — confirmed live: 109/109 already scanned), the map
// came back empty on every subsequent run, and backfillMCFFeesRecent() permanently reported every
// remaining pending row as "not settled" — even for orders whose fee had already been extracted
// in a past run and then silently discarded, because only the report-id bookkeeping was persisted,
// never the fee data itself. This sheet-backed cache persists every (merchant-order-id → fee)
// pair the moment it's extracted, independent of whether it matched a currently-pending row, so
// it accumulates across runs instead of resetting to empty once scanning catches up.
var FEE_CACHE_SHEET_NAME = '_SettlementFeeCache';
var FEE_CACHE_HEADER = ['merchant-order-id', 'fee', 'endpoint', 'reportId', 'cachedAt'];

function _getFeeCacheSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FEE_CACHE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FEE_CACHE_SHEET_NAME);
    sheet.getRange(1, 1, 1, FEE_CACHE_HEADER.length).setValues([FEE_CACHE_HEADER]);
    sheet.hideSheet();
  }
  return sheet;
}

// Loads the whole cache into { [merchant-order-id]: fee }.
function _loadFeeCache() {
  var sheet = _getFeeCacheSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  data.forEach(function(r) {
    var id = String(r[0] || '').trim();
    if (id) map[id] = Number(r[1]) || 0;
  });
  return map;
}

// Appends entries not already present in the cache (never overwrites an existing row — settlement
// data is immutable once posted, so the first value seen for an order id is authoritative).
// Returns how many new rows were appended.
function _appendToFeeCache(ep, entries, reportIdByOrder) {
  var sheet = _getFeeCacheSheet();
  var existing = _loadFeeCache();
  var rows = [];
  var now = new Date();
  Object.keys(entries).forEach(function(orderId) {
    if (existing.hasOwnProperty(orderId)) return;
    rows.push([orderId, entries[orderId], ep, (reportIdByOrder && reportIdByOrder[orderId]) || '', now]);
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, FEE_CACHE_HEADER.length).setValues(rows);
  }
  return rows.length;
}

/**
 * ONE-TIME CLEANUP — run once from the editor after deploying the "GCX-" filter above.
 * The very first post-fix scan (before the filter existed) wrote ~188,000 rows into
 * _SettlementFeeCache, the overwhelming majority of them ordinary Amazon retail orders this sheet
 * will never look up. Rewrites the sheet keeping only GCX- prefixed rows. Safe to run any time —
 * read-then-overwrite, no data this sheet actually needs is at risk.
 */
function pruneFeeCacheToGcxOnly() {
  var sheet = _getFeeCacheSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('pruneFeeCacheToGcxOnly: cache is empty, nothing to prune.');
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, FEE_CACHE_HEADER.length).getValues();
  var kept = data.filter(function(r) { return String(r[0] || '').indexOf('GCX-') === 0; });

  sheet.getRange(2, 1, lastRow - 1, FEE_CACHE_HEADER.length).clearContent();
  if (kept.length) {
    sheet.getRange(2, 1, kept.length, FEE_CACHE_HEADER.length).setValues(kept);
  }
  Logger.log('pruneFeeCacheToGcxOnly: %s rows before, %s GCX- rows kept, %s non-GCX rows removed.',
    data.length, kept.length, data.length - kept.length);
}

/**
 * Builds a merchant-order-id → fee map from GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2 reports.
 *
 * Why this exists (confirmed live, 2026-07-30): Amazon's Finances API
 * (/finances/v0/financialEvents and /finances/v0/orders/{id}/financialEvents) has NO fee data
 * at all for these self-created ("Non-Amazon <country>" marketplace) MCF shipments — their
 * ShipmentEventList entries use a synthetic "S02-..." order id, never the GCX sheet id, and a
 * targeted single-order lookup by the GCX id returns a structurally valid but completely empty
 * result. The **Settlement Report**, however, preserves the sheet's exact GCX id in its
 * `merchant-order-id` column, alongside a proper `ItemFees` / `FBAPerUnitFulfillmentFee` line —
 * e.g. confirmed live: merchant-order-id=GCX-FR-260716-2695, amount-type=ItemFees,
 * amount-description=FBAPerUnitFulfillmentFee, amount=-7,50. So this is now the primary (only
 * working) source for col Y fees, keyed by a **direct** match against the sheet's Q-column value
 * — no numeric alias or displayableOrderId resolution needed.
 *
 * Hard constraint: the reports-listing endpoint rejects `createdSince` older than ~90 days
 * ("RequestedFromDate ... is more than 90 days old"), so this can only ever cover the last ~89
 * days — which happens to match backfillMCFFeesRecent()'s own scoping. Older pending rows are
 * simply outside what SP-API allows querying at all via this mechanism.
 *
 * @param ep 'EU' or 'FE'.
 * @param sinceDate Date — clamped internally to at most 89 days ago.
 * @return plain object { [merchant-order-id]: absoluteFeeAmount }
 */
function _buildSettlementFeeMap(ep, sinceDate) {
  var feeMap = {};
  var reportIdByOrder = {};
  var now = new Date(Date.now() - 5 * 60 * 1000);
  var earliestAllowed = new Date(now.getTime() - 89 * 24 * 3600 * 1000);
  var since = (sinceDate && sinceDate > earliestAllowed) ? sinceDate : earliestAllowed;

  var marketplaceId = _settlementMarketplaceId(ep);
  var reports = [];
  var nextToken = null;
  var page = 0;

  do {
    // The reports-listing endpoint rejects nextToken combined with any other query param
    // ("NextToken cannot be specified with other input parameters") — confirmed live, since
    // this account settles often enough to exceed one page within the 89-day window.
    var qs = nextToken
      ? 'nextToken=' + encodeURIComponent(nextToken)
      : 'reportTypes=GET_V2_SETTLEMENT_REPORT_DATA_FLAT_FILE_V2' +
        '&processingStatuses=DONE' +
        '&marketplaceIds=' + marketplaceId +
        '&createdSince=' + encodeURIComponent(since.toISOString()) +
        '&createdUntil=' + encodeURIComponent(now.toISOString()) +
        '&pageSize=100';

    var res = spapiFetchWithRetry('GET', '/reports/2021-06-30/reports', { queryString: qs, endpoint: ep }, 2, 15000);
    var payload = res.payload || res;
    reports = reports.concat(payload.reports || []);
    nextToken = payload.nextToken || null;
    page++;
  } while (nextToken && page < 10);

  // This account can settle very frequently (100+ reports observed within the 89-day window,
  // some with 200k+ lines) — cap how many get downloaded+parsed per run so this can't run away
  // past GAS's execution time limit.
  //
  // Bug fixed 2026-07-30: originally always took the N most-recent reports every run. Since
  // brand-new reports keep appearing, any report ranked below the cap NEVER gets scanned before
  // it ages out of the 89-day window — silently losing that data forever, not just deferring it.
  // Fix: persist which reportIds have already been scanned (Script Properties, survives across
  // runs) and prioritize UNSCANNED reports, oldest first — oldest because those are closest to
  // aging out of the listing window and need to be captured before that happens. Already-scanned
  // reports need never be re-read (settlement reports are immutable once DONE).
  var scannedIds = _getScannedSettlementReportIds(ep);
  var unscanned = reports.filter(function(r) { return !scannedIds[r.reportId]; });
  unscanned.sort(function(a, b) { return new Date(a.createdTime) - new Date(b.createdTime); }); // oldest first
  var maxReports = 40;
  var toScan = unscanned.slice(0, maxReports);
  Logger.log('_buildSettlementFeeMap[%s]: %s report(s) in window since %s, %s already scanned, %s new this run (oldest-unscanned-first)',
    ep, reports.length, since.toISOString().slice(0,10), reports.length - unscanned.length, toScan.length);

  var startTime = Date.now();
  var timeBudgetMs = 4.5 * 60 * 1000; // bail before GAS's execution limit, not after
  var scanned = 0;
  var newlyScanned = {};
  // FIXED 2026-08-01: this loop had no circuit breaker — under sustained account-wide
  // throttling (confirmed live: runs where every single attempt 429s) it would burn the entire
  // 4.5-minute budget retrying doomed reports one by one, making zero progress each time.
  // retryR429Errors() (col R) already bails out after 5 consecutive 429s for the same reason;
  // mirror that here instead of wasting the whole run when the API is clearly still saturated.
  var maxConsec429 = 5;
  var consec429 = 0;

  for (var ri = 0; ri < toScan.length; ri++) {
    if (Date.now() - startTime > timeBudgetMs) {
      Logger.log('_buildSettlementFeeMap[%s]: time budget reached — scanned %s/%s, rest deferred to next run', ep, scanned, toScan.length);
      break;
    }
    if (consec429 >= maxConsec429) {
      Logger.log('_buildSettlementFeeMap[%s]: %s consecutive 429s — quota clearly still exhausted, stopping early (scanned %s/%s, rest deferred to next run)',
        ep, consec429, scanned, toScan.length);
      break;
    }
    var r = toScan[ri];
    scanned++;
    try {
      var docRes = spapiFetchWithRetry('GET', '/reports/2021-06-30/documents/' + encodeURIComponent(r.reportDocumentId), { endpoint: ep }, 2, 15000);
      var doc = docRes.payload || docRes;
      var resp = UrlFetchApp.fetch(doc.url, { muteHttpExceptions: true });
      var text;
      if (doc.compressionAlgorithm === 'GZIP') {
        text = Utilities.ungzip(Utilities.newBlob(resp.getBlob().getBytes(), 'application/gzip')).getDataAsString('UTF-8');
      } else {
        text = resp.getContentText('UTF-8');
      }

      var lines = text.split('\n');
      if (!lines.length) continue;
      var header = lines[0].split('\t');
      var moiIdx = header.indexOf('merchant-order-id');
      var typeIdx = header.indexOf('amount-type');
      var amtIdx = header.indexOf('amount');
      if (moiIdx < 0 || typeIdx < 0 || amtIdx < 0) {
        Logger.log('  report %s: unexpected header, skipping', r.reportId);
        continue;
      }

      for (var i = 1; i < lines.length; i++) {
        if (!lines[i]) continue;
        var f = lines[i].split('\t');
        var moi = (f[moiIdx] || '').trim();
        // A settlement report covers the ENTIRE seller account, not just this sheet's
        // self-created MCF orders — confirmed live: 19 reports alone produced 124,647 distinct
        // order ids, the overwhelming majority of which are ordinary Amazon retail orders
        // ("111-1234567-1234567" style) this sheet will never look up. Filtering to this sheet's
        // own ID scheme keeps the persistent cache from growing into a multi-hundred-thousand-row
        // sheet full of data nothing ever reads.
        if (!moi || moi.indexOf('GCX-') !== 0 || f[typeIdx] !== 'ItemFees') continue;
        var amt = Math.abs(_parseSettlementAmount(f[amtIdx]));
        if (amt === 0) continue;
        feeMap[moi] = (feeMap[moi] || 0) + amt;
        if (!reportIdByOrder[moi]) reportIdByOrder[moi] = r.reportId;
      }
      newlyScanned[r.reportId] = r.createdTime; // only mark done on full success — a failed report retries next run
      consec429 = 0;
    } catch (e) {
      // spapiFetchWithRetry above already retries once on 429 (15s wait) — if it still failed,
      // this report is simply skipped and picked up on a later run.
      Logger.log('  report %s: error — %s', r.reportId, e.message || e);
      if (_isRateLimit429(e)) consec429++;
    }
    Utilities.sleep(800); // pace document downloads — confirmed live this endpoint 429s in bursts
  }

  var mergedScanned = scannedIds;
  Object.keys(newlyScanned).forEach(function(id) { mergedScanned[id] = newlyScanned[id]; });
  _saveScannedSettlementReportIds(ep, mergedScanned);

  // Persist every fee extracted this run into the durable cache immediately — regardless of
  // whether it matches a row that happens to be pending in THIS run's sheet snapshot. This is
  // what makes the fix durable: once a report has been scanned, its data is never lost, even
  // after the report itself is never looked at again.
  var newlyCached = _appendToFeeCache(ep, feeMap, reportIdByOrder);

  // Return the FULL persistent cache merged with this run's fresh finds, not just this run's
  // finds — so a caller always sees every fee ever extracted, including from reports scanned in
  // past runs whose data would otherwise have been silently discarded once scanning caught up.
  var cached = _loadFeeCache();
  var merged = {};
  Object.keys(cached).forEach(function(id) { merged[id] = cached[id]; });
  Object.keys(feeMap).forEach(function(id) { merged[id] = feeMap[id]; });

  Logger.log('_buildSettlementFeeMap[%s]: %s order id(s) resolved this run (%s newly cached), %s report(s) newly marked scanned, %s total order id(s) available (persistent cache + this run)',
    ep, Object.keys(feeMap).length, newlyCached, Object.keys(newlyScanned).length, Object.keys(merged).length);
  return merged;
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
 * Sums MCF fee lines across ALL ShipmentEvents matching targetOrderId (an order can settle
 * across more than one event — e.g. split shipments — so stopping at the first match can
 * silently miss the event that actually carries the fee).
 *
 * GCX-prefixed orders (all orders in this sheet) sum every fee line regardless of FeeType —
 * their MCF fee lines are posted under type names that don't contain "FBA"/"FULFILLMENT",
 * so the standard _isMcfFeeType filter silently drops real fee data to 0 for them. Non-GCX
 * orders keep the standard filter.
 *
 * Returns '' (not '0') when a match was found but no event ever summed to a nonzero total —
 * treated as not-yet-settled rather than risking a misleading $0 in the sheet's profit/loss
 * columns; run MCFFeeDebug(orderId, sentDate) to inspect the raw events if it stays '' for
 * a long-settled order.
 */
function _sumMcfFeeFromShipments(shipments, targetOrderId) {
  var target = String(targetOrderId).trim();
  var isGcx  = target.toUpperCase().indexOf('GCX') === 0;
  var best   = 0;

  for (var i = 0; i < shipments.length; i++) {
    var ev = shipments[i];
    if (String(ev.SellerOrderId || '').trim() !== target) continue;
    var total = 0;
    (ev.ShipmentFeeList || []).forEach(function(f) {
      if (isGcx || _isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
    });
    (ev.ShipmentItemList || []).forEach(function(item) {
      (item.ItemFeeList || []).forEach(function(f) {
        if (isGcx || _isMcfFeeType(f.FeeType)) total += parseFloat((f.FeeAmount || {}).CurrencyAmount || 0);
      });
    });
    if (total !== 0) best = Math.abs(total);
  }

  return best !== 0 ? best : '';
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
      postedAfter  = new Date(String(sentDate).trim());
      postedBefore = new Date(postedAfter);
      postedBefore.setDate(postedBefore.getDate() + 90);
      dateSource = 'sentDate (P col): ' + String(sentDate).trim();
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
 * Wider-window variant of MCFFeeDebug() — searches up to `windowDays` (default 270) from
 * sentDate instead of the fixed 90-day window, and also checks the +1 GCX alias
 * (_gcxNumAlias) so it's clear whether widening the window, the alias, or neither would help.
 * Only returns matching rows (exact / alias / displayableOrderId) plus a summary of how many
 * events were scanned — the raw event list can run into the thousands (mostly unrelated
 * marketplace orders) and isn't useful for this question.
 *
 * @customfunction
 * @param {string} orderId The sellerFulfillmentOrderId to debug (Q col value).
 * @param {string} sentDate yyyy-mm-dd sent date from col P.
 * @param {number} [windowDays] How many days from sentDate to scan. Default 270 (~9 months).
 * @return {Array} SellerOrderId_in_API | Match type | FeeTypes | Total
 */
function MCFFeeDebugWide(orderId, sentDate, windowDays) {
  if (!orderId) return [['orderId is required']];
  if (!sentDate) return [['sentDate is required']];
  windowDays = windowDays || 270;

  try {
    var postedAfter  = new Date(String(sentDate).trim());
    var postedBefore = new Date(postedAfter);
    postedBefore.setDate(postedBefore.getDate() + windowDays);

    var now = new Date(Date.now() - 5 * 60 * 1000);
    if (postedBefore > now) postedBefore = now;

    var shipments = _collectShipmentEvents('EU', postedAfter, postedBefore, 15); // up to ~1500 events

    var alias = _gcxNumAlias(String(orderId), 1);

    var displayableId = '';
    try {
      var foResult  = getFulfillmentOrderRaw(String(orderId), 'EU');
      var candidate = ((foResult.fulfillmentOrder || {}).displayableOrderId || '').trim();
      if (candidate && candidate !== String(orderId).trim()) displayableId = candidate;
    } catch (e) { /* ignore */ }

    var rows  = [['SellerOrderId_in_API', 'Match type', 'FeeTypes', 'Total']];
    var found = 0;

    shipments.forEach(function(ev) {
      var sid = String(ev.SellerOrderId || '').trim();
      var matchType = null;
      if (sid === String(orderId).trim())        matchType = 'exact (Q col)';
      else if (alias && sid === alias)           matchType = 'alias (N+1)';
      else if (displayableId && sid === displayableId) matchType = 'displayableOrderId';
      if (!matchType) return;

      found++;
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
      rows.push([sid, matchType, feeTypes.join(', '), Math.abs(total)]);
    });

    if (!found) {
      rows.push(['(no match — scanned ' + shipments.length + ' events over ' + windowDays + ' days)', '', '', '']);
    }
    rows.push([
      'window: ' + postedAfter.toISOString().slice(0, 10) + ' → ' + postedBefore.toISOString().slice(0, 10),
      'alias tried: ' + (alias || 'n/a'),
      'displayableId: ' + (displayableId || 'n/a'),
      ''
    ]);

    return rows;
  } catch (e) { return [['ERR: ' + (e.message || e)]]; }
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
