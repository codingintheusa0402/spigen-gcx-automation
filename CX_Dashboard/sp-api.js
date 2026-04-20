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
  return bytes.map(function(b) {
    var s = (b & 0xff).toString(16);
    return s.length === 1 ? '0' + s : s;
  }).join('');
}

function _sha256Hex(msg) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, msg, Utilities.Charset.UTF_8);
  return _toHex(digest);
}

function _sessionToken() {
  return _prop('AWS_SESSION_TOKEN', '');
}

/***** ========= LWA ========= *****/
function _resolveLwaProfileKey(endpointKey) {
  var k = (endpointKey || '').toString().toUpperCase();
  var feMkt = ['A1VC38T7YXB528', 'A39IBJ37TRP1C6', 'A19VAU5U5O7RUS'];
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
  if (!clientId || !clientSecret || !refreshToken)
    throw new Error('Missing LWA credentials for profile ' + prof);

  var resp = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    payload: { grant_type: 'refresh_token', refresh_token: refreshToken, client_id: clientId, client_secret: clientSecret },
    muteHttpExceptions: true
  });
  var body = JSON.parse(resp.getContentText() || '{}');
  if (resp.getResponseCode() >= 300 || !body.access_token)
    throw new Error('LWA token fetch failed: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  return body.access_token;
}

/***** ========= AWS SigV4 ========= *****/
function signSpApiRequest(method, host, path, queryString, body, region) {
  var accessKey = _prop('AWS_ACCESS_KEY_ID');
  var secretKey = _prop('AWS_SECRET_ACCESS_KEY');
  if (!accessKey || !secretKey) throw new Error('Missing AWS keys in Script Properties.');

  var service = 'execute-api';
  var ts = _nowIsoBasic();
  var amzDate = ts.amzDate, shortDate = ts.shortDate;
  var sessionToken = _sessionToken();
  var payloadHash = _sha256Hex(body || '');

  var canonicalHeaders =
    'host:' + host + '\n' +
    'x-amz-date:' + amzDate + '\n' +
    (sessionToken ? 'x-amz-security-token:' + sessionToken + '\n' : '');
  var signedHeaders = sessionToken ? 'host;x-amz-date;x-amz-security-token' : 'host;x-amz-date';

  var canonicalRequest = [method, path, queryString || '', canonicalHeaders, signedHeaders, payloadHash].join('\n');
  var credentialScope  = shortDate + '/' + region + '/' + service + '/aws4_request';
  var stringToSign     = ['AWS4-HMAC-SHA256', amzDate, credentialScope, _sha256Hex(canonicalRequest)].join('\n');

  var enc      = function(s) { return Utilities.newBlob(s).getBytes(); };
  var kSigning = Utilities.computeHmacSha256Signature(enc('aws4_request'),
    Utilities.computeHmacSha256Signature(enc(service),
      Utilities.computeHmacSha256Signature(enc(region),
        Utilities.computeHmacSha256Signature(enc(shortDate), enc('AWS4' + secretKey)))));
  var signature = _toHex(Utilities.computeHmacSha256Signature(enc(stringToSign), kSigning));

  return {
    amzDate: amzDate,
    sessionToken: sessionToken,
    authorizationHeader: 'AWS4-HMAC-SHA256 Credential=' + accessKey + '/' + credentialScope +
      ', SignedHeaders=' + signedHeaders + ', Signature=' + signature
  };
}

/***** ========= ENDPOINT RESOLVER ========= *****/
// EU marketplaces → sellingpartnerapi-eu  |  FE (JP/AU/SG) → sellingpartnerapi-fe
function _getEndpoint(groupOrMarketplace) {
  var g = String(groupOrMarketplace || '').toUpperCase();
  var feMkt = ['A1VC38T7YXB528', 'A39IBJ37TRP1C6', 'A19VAU5U5O7RUS'];
  if (feMkt.indexOf(g) >= 0 || ['JP', 'AU', 'SG', 'FE'].indexOf(g) >= 0) {
    return { host: _prop('SPAPI_HOST_FE', 'sellingpartnerapi-fe.amazon.com'), region: _prop('SPAPI_REGION_FE', 'us-west-2'), group: 'FE' };
  }
  return { host: _prop('SPAPI_HOST_EU', 'sellingpartnerapi-eu.amazon.com'), region: _prop('SPAPI_REGION_EU', 'eu-west-1'), group: 'EU' };
}

/***** ========= CORE FETCH ========= *****/
function spapiFetch(method, path, opts) {
  opts = opts || {};
  var ep   = _getEndpoint(opts.endpoint);
  var body = opts.body || '';
  var qs   = opts.queryString || '';
  var sig  = signSpApiRequest(method, ep.host, path, qs, body, ep.region);
  var url  = 'https://' + ep.host + path + (qs ? '?' + qs : '');

  var headers = {
    'x-amz-date': sig.amzDate,
    'x-amz-access-token': getLwaAccessToken(opts.endpoint),
    'Authorization': sig.authorizationHeader,
    'Content-Type': 'application/json'
  };
  if (sig.sessionToken) headers['x-amz-security-token'] = sig.sessionToken;

  var fetchOpts = { method: method, headers: headers, muteHttpExceptions: true };
  if (method !== 'GET' && method !== 'DELETE' && body) fetchOpts.payload = body;

  var resp = UrlFetchApp.fetch(url, fetchOpts);
  var code = resp.getResponseCode();
  var text = resp.getContentText();
  if (code >= 300) throw new Error('SP-API error ' + code + ': ' + text);
  return JSON.parse(text || '{}');
}

function spapiFetchWithRetry(method, path, opts, attempts, waitMs) {
  attempts = attempts == null ? 3 : attempts;
  waitMs   = waitMs   == null ? 5000 : waitMs;
  var lastErr = null;
  for (var i = 0; i < attempts; i++) {
    try { return spapiFetch(method, path, opts); }
    catch (err) {
      lastErr = err;
      var msg = err.message || String(err);
      if ((msg.indexOf('SP-API error 429') >= 0 || msg.indexOf('QuotaExceeded') >= 0) && i < attempts - 1) {
        Utilities.sleep(waitMs); continue;
      }
      throw err;
    }
  }
  if (lastErr) throw lastErr;
}

/***** ========= ERROR HELPERS ========= *****/
function _isUnauthorizedError(err) {
  var m = (err && err.message) ? err.message : String(err);
  return m.indexOf('SP-API error 403') >= 0 && m.indexOf('"code":"Unauthorized"') >= 0;
}

function _isNotFoundError(err) {
  var m = (err && err.message) ? err.message : String(err);
  return m.indexOf('SP-API error 404') >= 0;
}
