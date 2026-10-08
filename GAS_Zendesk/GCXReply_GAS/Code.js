// GCX Reply — Apps Script Web App (v2.7.3)
// Endpoint: ?orderId=XXX  |  ?asin=XXX  |  ?orderId=XXX&asin=XXX
// Deploy as: Execute as Me, Access: Anyone (or Anyone anonymous)
// Script Properties required: AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY,
//   LWA_CLIENT_ID, LWA_CLIENT_SECRET, LWA_REFRESH_TOKEN,         ← EU + NA
//   LWA_CLIENT_ID_JP, LWA_CLIENT_SECRET_JP, LWA_REFRESH_TOKEN_JP ← Japan (FE)
//   LWA_CLIENT_ID_IN, LWA_CLIENT_SECRET_IN, LWA_REFRESH_TOKEN_IN ← India

// Script Properties are themselves a network-backed GAS service call, not an
// in-memory read — memoize per execution so building N signed SP-API requests
// in one doGet() invocation only pays this cost once, not N times.
let _scriptPropsCache_ = null;
function scriptProps_() {
  if (!_scriptPropsCache_) _scriptPropsCache_ = PropertiesService.getScriptProperties().getProperties();
  return _scriptPropsCache_;
}

const SHEET_ID    = '1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo';
const SHEET_NAME  = 'Data';
const PRODUCT_COLS  = ['SKU','모델명','브랜드','제조사명','기종명','색상명','대분류','생산업체','원산지정보'];
const MARKET_SS_ID  = '172fDVw4tu-hgbpV5FShWj4_SAMxeB54-v5BUlVgJUoA';
const MARKET_SHEETS = ['DE', 'NL', 'SE', 'ES', 'UK', 'FR', 'IT', 'JP', 'IN', 'SG'];

// ── AI 인입사유 (DR) ───────────────────────────────────────────────────────────
const DEFECT_SS_ID      = '1jC2k5Cssj4nYI6R1gMlqtpyVTeDUIc-hTirdY3Ij8mk';
const DEFECT_SHEET_NAME = 'GCX 인입사유';
const DR_CACHE_VERSION  = 'DR_v24_';
const DR_CACHE_TTL      = 21600;
const GEMINI_MODELS_DR  = ['gemini-2.5-flash-lite', 'gemini-2.5-flash'];

const REGIONS = [
  { endpoint: 'https://sellingpartnerapi-eu.amazon.com', region: 'eu-west-1', cred: 'main' },
  { endpoint: 'https://sellingpartnerapi-fe.amazon.com', region: 'us-west-2', cred: 'jp'   },
  { endpoint: 'https://sellingpartnerapi-na.amazon.com', region: 'us-east-1', cred: 'main' },
  // India uses the EU endpoint but a separate Seller Central account (own refresh token)
  { endpoint: 'https://sellingpartnerapi-eu.amazon.com', region: 'eu-west-1', cred: 'in'   },
];

// SalesChannel suffix → SP-API Marketplace ID (order matters: longest suffix first)
const MARKETPLACE_MAP = [
  ['.com.sg', 'A19VAU5U5O7RUS'],
  ['.com.au', 'A39IBJ37TRP1C6'],
  ['.com.mx', 'A1AM78C64UM0Y8'],
  ['.com.tr', 'A33AVAJ2PDY3EV'],
  ['.co.uk',  'A1F83G8C2ARO7P'],
  ['.co.jp',  'A1VC38T7YXB528'],
  ['.de',     'A1PA6795UKMFR9'],
  ['.fr',     'A13V1IB3VIYZZH'],
  ['.it',     'APJ6JRA9NG5V4'],
  ['.es',     'A1RKKUPIHCS9HS'],
  ['.nl',     'A1805IZSGTT6HS'],
  ['.pl',     'AZ1PBY3F3E3AE'],
  ['.se',     'A2NODRKZP88ZB9'],
  ['.be',     'AMEN7PMS3EDWL'],
  ['.in',     'A21TJRUUN4KGV'],
  ['.ca',     'A2EUQ1WTGCTBG2'],
  ['.tr',     'A33AVAJ2PDY3EV'],
  ['.com',    'ATVPDKIKX0DER'],
];

function marketplaceId_(salesChannel) {
  if (!salesChannel) return null;
  const s = salesChannel.toLowerCase();
  const match = MARKETPLACE_MAP.find(([suffix]) => s.includes(suffix));
  return match ? match[1] : null;
}

// ── Entry point ───────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const p       = (e && e.parameter) || {};
    const orderId = p.orderId;
    const asin    = p.asin;

    if (p.action === 'inferReason') {
      const review   = p.review   || '';
      const category = p.category || '';
      if (!review) return respond({ error: 'Provide review parameter' });
      return respond({ reason: inferReason_(review, category) });
    }

    if (p.action === 'abmRelayPending') {
      return respond({ pending: getAbmRelayPending_(p.clientVersion) });
    }

    if (p.action === 'abmRelayStatus') {
      return respond({ status: getAbmRelayStatus_(p.ticketId) });
    }

    if (p.action === 'abmRelayAll') {
      return respond({ rows: getAbmRelayAll_(Number(p.limit) || 50) });
    }

    if (!orderId && !asin) {
      return respond({ error: 'Provide orderId and/or asin parameter' });
    }

    const result = {};

    if (orderId) {
      if (!/^\d{3}-\d{7}-\d{7}$/.test(orderId)) {
        return respond({ error: 'Invalid order ID format' });
      }
      const orderData = fetchOrderData_(orderId);
      Object.assign(result, orderData);

      // Auto-lookup ASIN from items if not passed explicitly
      const itemAsin = !asin && orderData.items && orderData.items[0]
        ? orderData.items[0].ASIN : null;
      if (itemAsin) Object.assign(result, lookupProductFull_(itemAsin));
    }

    if (asin) Object.assign(result, lookupProductFull_(asin));

    return respond(result);
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ABM relay status log ──────────────────────────────────────────────────────
// Durable, cross-browser-session record of whether each ABM ticket's Public
// reply was actually relayed to the Amazon buyer via Seller Central — a toast
// in the agent's browser at send-time isn't enough on its own: it's gone the
// moment the agent navigates away, and doesn't help if the whole relay attempt
// silently died (tab closed mid-flight, etc.).
//
// Upserted by `RelayKey` (`${ticketId}_${startTimeMs}`, generated once per
// relay invocation client-side), NOT by bare TicketId — a single ticket can
// get multiple agent replies over time, each triggering its own separate
// relay. An earlier version keyed by TicketId alone, so a second reply's log
// entry silently overwrote the first reply's row, making it impossible to
// tell which of two same-ticket messages actually failed (confirmed live:
// exactly this ambiguity on tickets #1000153623/#1000153609 — two replies
// each, only one `success` row survived per ticket with no way to tell which
// send it corresponded to). `MessageText` is kept so a later retry sweep (any
// agent's browser, not necessarily the one that first tried) can resend
// without needing to re-open the original ticket.
const ABM_LOG_SHEET_NAME = 'ABM_Relay_Log';
// CommentId ties a row to the specific Zendesk comment it relays — needed by
// the reconciliation job (reconcileAbmRelays_ in ABM_TicketMerge) so an
// auto-resent reply can re-fetch that exact comment's attachments, and so its
// dedup can tell one agent reply apart from another on the same ticket.
const ABM_LOG_HEADERS = ['Timestamp', 'RelayKey', 'TicketId', 'CommentId', 'CaseId', 'Marketplace', 'Status', 'Attempts', 'LastError', 'MessageText'];

function getAbmLogSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ABM_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ABM_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, ABM_LOG_HEADERS.length).setValues([ABM_LOG_HEADERS]);
    return sheet;
  }
  // Self-repairing migration: insert any expected column that doesn't exist
  // yet, at its target index, preserving existing rows/data. Handles both the
  // original pre-RelayKey (8-col) sheet and the pre-CommentId (9-col) one —
  // without this, column lookups would silently read the wrong cell against
  // existing rows. Only this code writes the sheet, so header order is
  // otherwise stable.
  ABM_LOG_HEADERS.forEach((h, idx) => {
    const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    if (existing.indexOf(h) === -1) sheet.insertColumnBefore(idx + 1);
  });
  sheet.getRange(1, 1, 1, ABM_LOG_HEADERS.length).setValues([ABM_LOG_HEADERS]);
  return sheet;
}

function upsertAbmRelayLog_(entry) {
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  const keyCol = ABM_LOG_HEADERS.indexOf('RelayKey');
  // Fallback so a pre-migration row (old TicketId-only key, or a client still
  // running a pre-RelayKey version) doesn't get duplicated on its next update.
  const relayKey = entry.relayKey || String(entry.ticketId);
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]) === relayKey) { rowIdx = i + 1; break; }
  }
  const row = [
    new Date(), relayKey, entry.ticketId, entry.commentId || '', entry.caseId || '', entry.marketplace || '',
    entry.status || '', entry.attempts || 1, entry.lastError || '', entry.messageText || ''
  ];
  if (rowIdx === -1) sheet.appendRow(row);
  else sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
}

// A "pending" row is written the instant a relay starts (see
// GCX Reply.user.js relayAbmReply_), before any send attempt — so a tab
// closed/reloaded mid-flight still leaves a recoverable record. But that
// means a fresh "pending" row is often just a completely normal send still
// in progress a few seconds ago, not an abandoned one. Only surface it to
// the cross-session sweep once it's old enough that the originating tab's
// own retry loop (up to 3 attempts, each with network round trips + a
// deliberate backoff — worst case comfortably under a minute) must already
// be done one way or another. Without this, a second tab's sweep could pick
// up a still-in-flight send and fire a genuine duplicate message to the buyer.
const ABM_PENDING_STALE_MS = 90 * 1000;
// A 'sending' row is one a browser has atomically claimed (claimAbmRelay) and
// is actively delivering. Hidden from the pending list so no other browser
// grabs the same reply — until this long passes, at which point we assume the
// claiming browser died mid-send (tab closed, crash) and let another retry.
const ABM_SENDING_STALE_MS = 5 * 60 * 1000;

// Minimum userscript version that can correctly relay a queued reply's
// attachments (fetch the specific comment's files + upload to SC). Reconciliation
// queues rows carrying a CommentId; a pre-3.3.12 tab's sweep would send those
// TEXT-ONLY and silently drop the attachments (exactly what dropped Wolfgang's
// 3 invoice PDFs). So CommentId-bearing rows are withheld from any client that
// doesn't declare it's at least this version — they wait for a capable browser.
const ABM_ATTACH_MIN_VERSION = '3.3.12';

function versionAtLeast_(v, min) {
  const a = String(v || '').split('.').map(Number);
  const b = String(min).split('.').map(Number);
  for (let i = 0; i < b.length; i++) {
    const x = a[i] || 0, y = b[i] || 0;
    if (x !== y) return x > y;
  }
  return true;
}

function getAbmRelayPending_(clientVersion) {
  const capable = versionAtLeast_(clientVersion, ABM_ATTACH_MIN_VERSION);
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idx = name => headers.indexOf(name);
  const now = Date.now();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[idx('Status')];
    if (!row[idx('TicketId')] || status === 'success') continue;
    const ts = new Date(row[idx('Timestamp')]).getTime();
    // in-flight browser send — leave it alone briefly
    if (status === 'pending' && (!ts || (now - ts) < ABM_PENDING_STALE_MS)) continue;
    // claimed by another browser and still plausibly delivering — skip until stale
    if (status === 'sending' && (!ts || (now - ts) < ABM_SENDING_STALE_MS)) continue;
    // attachment-capable send required — hide from too-old clients so they can't
    // text-only-send it and drop the files.
    if (row[idx('CommentId')] && !capable) continue;
    out.push({
      relayKey:    row[idx('RelayKey')],
      ticketId:    row[idx('TicketId')],
      commentId:   row[idx('CommentId')],
      caseId:      row[idx('CaseId')],
      marketplace: row[idx('Marketplace')],
      status:      status,
      attempts:    row[idx('Attempts')],
      lastError:   row[idx('LastError')],
      timestamp:   row[idx('Timestamp')],
      messageText: row[idx('MessageText')],
    });
  }
  return out;
}

// Atomic claim so only ONE browser sends a given queued reply — critical now
// that reconciliation can auto-queue real-customer replies that any agent's
// tab might pick up. Uses LockService to serialize the read-modify-write:
// grants the claim (flips the row to 'sending' with a fresh timestamp) only if
// the row is still pending/failed/stale-sending. Returns { claimed: bool }.
function claimAbmRelay_(relayKey) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { claimed: false, reason: 'lock_timeout' }; }
  try {
    const sheet = getAbmLogSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idx = name => headers.indexOf(name);
    const now = Date.now();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx('RelayKey')]) !== String(relayKey)) continue;
      const status = data[i][idx('Status')];
      const ts = new Date(data[i][idx('Timestamp')]).getTime();
      const claimable = status === 'pending' || status === 'failed' ||
        (status === 'sending' && (!ts || (now - ts) >= ABM_SENDING_STALE_MS));
      if (!claimable) return { claimed: false, reason: 'already_' + status };
      sheet.getRange(i + 1, idx('Status') + 1).setValue('sending');
      sheet.getRange(i + 1, idx('Timestamp') + 1).setValue(new Date());
      SpreadsheetApp.flush();
      return { claimed: true };
    }
    return { claimed: false, reason: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

// Guards the INITIAL send (relayAbmReply_ client-side), a gap claimAbmRelay_
// above never covered — that one only stops two browsers double-sending an
// ALREADY-QUEUED retry row, but nothing stopped the live send itself from
// firing twice for one real reply in the first place. Confirmed live via
// ABM_Relay_Log: 28+ tickets since 2026-07-28 show the identical message
// successfully relayed twice within seconds to minutes of each other, both
// from the live-send path — e.g. ticket #1000157277, 3 seconds apart, only
// one real Zendesk comment ever existed. `key` is ticketId+contentHash
// (computed client-side), not just ticketId, so a genuinely different
// second reply on the same ticket is never blocked by this — only an exact
// repeat within the TTL window is.
//
// CORRECTED 2026-08-19: this used to be a plain CacheService get-then-put
// with no lock, on the stated assumption that it "only needs to beat a
// network race measured in seconds." Disproven live — ABM_Relay_Log rows
// 632/633 (ticket #1000157657) show the identical duplicate-send pattern
// with RelayKeys 785ms apart, both status=success. A get-then-put has NO
// atomicity at any gap size: two calls can both pass cache.get() before
// either has finished cache.put(), and GAS's own scheduling/queuing under
// load (this account has had plenty this session) can easily produce
// sub-second overlaps. Wrapped in LockService for a real atomic
// check-and-set. Preserves the existing fail-OPEN philosophy (a skipped
// send is worse than a rare duplicate) for the one new failure mode this
// introduces: if the lock itself can't be acquired in time, still allow
// the send rather than silently blocking a real reply.
const ABM_SEND_CLAIM_TTL_SEC_ = 300; // 5min — comfortable margin over the observed few-second-to-~90s double-fire gaps
function claimAbmSend_(key) {
  if (!key) return { claimed: false, reason: 'missing_key' };
  const cache = CacheService.getScriptCache();
  const cacheKey = 'abm_send_claim_' + key;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    Logger.log(`claimAbmSend_(${key}): lock timeout, failing open — ${e}`);
    return { claimed: true, reason: 'lock_timeout_failopen' };
  }
  try {
    if (cache.get(cacheKey)) return { claimed: false, reason: 'already_claimed' };
    cache.put(cacheKey, '1', ABM_SEND_CLAIM_TTL_SEC_);
    return { claimed: true };
  } finally {
    lock.releaseLock();
  }
}

// A ticket can have multiple relay rows (one per reply) — returns ALL of
// them, not just one, so nothing is hidden the way a single-row-per-ticket
// lookup previously did.
function getAbmRelayStatus_(ticketId) {
  if (!ticketId) return null;
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const idx = name => headers.indexOf(name);
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx('TicketId')]) === String(ticketId)) {
      rows.push({
        relayKey:    data[i][idx('RelayKey')],
        commentId:   data[i][idx('CommentId')],
        status:      data[i][idx('Status')],
        attempts:    data[i][idx('Attempts')],
        lastError:   data[i][idx('LastError')],
        timestamp:   data[i][idx('Timestamp')],
        messageText: data[i][idx('MessageText')],
      });
    }
  }
  return rows.length ? rows : null;
}

// Full recent history (delivered + undelivered), newest first — backs the
// agent-facing status panel in GCX Reply, which lists both groups and lets
// agents edit a row's status (mark delivered / re-queue).
function getAbmRelayAll_(limit) {
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idx = name => headers.indexOf(name);
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[idx('TicketId')]) continue;
    out.push({
      relayKey:    row[idx('RelayKey')],
      ticketId:    row[idx('TicketId')],
      commentId:   row[idx('CommentId')],
      caseId:      row[idx('CaseId')],
      marketplace: row[idx('Marketplace')],
      status:      row[idx('Status')],
      attempts:    row[idx('Attempts')],
      lastError:   row[idx('LastError')],
      timestamp:   row[idx('Timestamp')],
      messageText: row[idx('MessageText')],
    });
  }
  out.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return out.slice(0, limit || 50);
}

// Agent-initiated status edit from the panel. Marking 'success' takes a row
// out of the retry sweep (agent confirmed/handled it manually); marking
// 'failed' re-queues a delivered-looking row the agent knows didn't actually
// arrive. Matched by RelayKey; bare-TicketId fallback covers pre-RelayKey rows.
function setAbmRelayStatus_(relayKey, ticketId, status) {
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = name => headers.indexOf(name);
  for (let i = 1; i < data.length; i++) {
    const rk = String(data[i][idx('RelayKey')] || '');
    const match = relayKey ? rk === String(relayKey)
                           : String(data[i][idx('TicketId')]) === String(ticketId);
    if (match) {
      sheet.getRange(i + 1, idx('Status') + 1).setValue(status);
      sheet.getRange(i + 1, idx('LastError') + 1).setValue('manually set by agent');
      return true;
    }
  }
  return false;
}

function deleteAbmRelayRows_(relayKey, ticketId) {
  const sheet = getAbmLogSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idx = name => headers.indexOf(name);
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const rk  = String(data[i][idx('RelayKey')] || '');
    const tid = String(data[i][idx('TicketId')] || '');
    if ((relayKey && rk === String(relayKey)) || (ticketId && tid === String(ticketId))) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}

// POST-only: message text can be long enough to make a GET query string unwieldy.
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'logAbmRelay') {
      upsertAbmRelayLog_(body);
      return respond({ ok: true });
    }
    if (body.action === 'setAbmRelayStatus') {
      return respond({ ok: setAbmRelayStatus_(body.relayKey, body.ticketId, body.status) });
    }
    if (body.action === 'deleteAbmRelayRow') {
      return respond({ deleted: deleteAbmRelayRows_(body.relayKey, body.ticketId) });
    }
    if (body.action === 'claimAbmRelay') {
      return respond(claimAbmRelay_(body.relayKey));
    }
    if (body.action === 'claimAbmSend') {
      return respond(claimAbmSend_(body.key));
    }
    if (body.action === 'setWorkerConfig') {
      return respond(setWorkerConfig_(body));
    }
    return respond({ error: 'unknown action' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── Keep-warm: prevents GAS cold starts during business hours ─────────────────
// Run setupKeepWarmTrigger() ONCE from the GAS editor to install it.
function keepWarm() {
  CacheService.getScriptCache().get('ping');
  refreshProductIndexIfStale_();
}

function setupKeepWarmTrigger() {
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'keepWarm');
  if (existing.length) { Logger.log('keepWarm trigger already exists (' + existing.length + ')'); return; }
  ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(5).create();
  Logger.log('keepWarm trigger created — fires every 5 minutes');
}

// ── JP auth diagnostic (run manually from GAS editor) ────────────────────────
// Clears cache, tests getLwaToken_('jp'), then hits a simple SP-API endpoint.
// Check Execution log for results.
function diagJP() {
  CacheService.getScriptCache().remove('lwa_jp');
  Logger.log('Cache cleared.');

  const props = scriptProps_();
  Logger.log('CLIENT_ID_JP present:      ' + !!props['LWA_CLIENT_ID_JP']);
  Logger.log('CLIENT_SECRET_JP present:  ' + !!props['LWA_CLIENT_SECRET_JP']);
  Logger.log('CLIENT_SECRET_JP prefix:   ' + (props['LWA_CLIENT_SECRET_JP'] || '').slice(0, 30));
  Logger.log('REFRESH_TOKEN_JP present:  ' + !!props['LWA_REFRESH_TOKEN_JP']);
  Logger.log('REFRESH_TOKEN_JP prefix:   ' + (props['LWA_REFRESH_TOKEN_JP'] || '').slice(0, 30));

  let token;
  try {
    token = getLwaToken_('jp');
    Logger.log('LWA token: OK  (' + token.slice(0, 30) + '...)');
  } catch(e) {
    Logger.log('LWA token: FAILED → ' + e.message);
    return;
  }

  // /sellers/v1/marketplaceParticipations needs no order ID — good auth test
  const r = spApiGet_('https://sellingpartnerapi-fe.amazon.com', 'us-west-2', 'jp',
                       '/sellers/v1/marketplaceParticipations');
  Logger.log('SP-API status: ' + r.status);
  Logger.log('SP-API body:   ' + r.body.slice(0, 600));
}

// ── LWA token (cached 4 min) ──────────────────────────────────────────────────
function getLwaToken_(cred) {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'lwa_' + cred;
  const hit      = cache.get(cacheKey);
  if (hit) return hit;

  const props  = scriptProps_();
  const sfx    = cred === 'jp' ? '_JP' : cred === 'in' ? '_IN' : '';
  const resp   = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    payload: {
      grant_type:    'refresh_token',
      refresh_token: props['LWA_REFRESH_TOKEN' + sfx],
      client_id:     props['LWA_CLIENT_ID'     + sfx],
      client_secret: props['LWA_CLIENT_SECRET' + sfx],
    },
    muteHttpExceptions: true,
  });

  const d = JSON.parse(resp.getContentText());
  if (!d.access_token) throw new Error('LWA failed: ' + resp.getContentText());
  // LWA access tokens live 3600 s. This used to cache for at most 180 s, so
  // every 3 min per credential set an agent's request paid an extra LWA
  // round-trip. Keep a 5-min safety margin; a 403 'expired' still evicts it.
  cache.put(cacheKey, d.access_token, Math.max(60, Math.min(d.expires_in - 300, 21600)));
  return d.access_token;
}

// ── AWS SigV4 ─────────────────────────────────────────────────────────────────
function sha256Hex_(msg) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, msg)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function hmac_(key, msg) {
  const msgBytes = Utilities.newBlob(msg).getBytes();
  const keyBytes = typeof key === 'string' ? Utilities.newBlob(key).getBytes() : key;
  return Utilities.computeHmacSha256Signature(msgBytes, keyBytes);
}

function hmacHex_(key, msg) {
  return hmac_(key, msg).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function signingKey_(secret, dateStamp, region) {
  const kDate    = hmac_('AWS4' + secret, dateStamp);
  const kRegion  = hmac_(kDate,           region);
  const kService = hmac_(kRegion,         'execute-api');
  return hmac_(kService, 'aws4_request');
}

// Returns a signed fetch-options object (url + headers) without calling UrlFetchApp.
// Use directly with UrlFetchApp.fetchAll() to parallelise multiple SP-API requests.
function spApiBuildFetch_(endpoint, region, cred, fullPath, tokenOverride) {
  const props     = scriptProps_();
  const accessKey = props['AWS_ACCESS_KEY_ID'];
  const secretKey = props['AWS_SECRET_ACCESS_KEY'];
  const token     = tokenOverride || getLwaToken_(cred);
  const host      = endpoint.replace('https://', '');

  const now       = new Date();
  const amzDate   = Utilities.formatDate(now, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp = Utilities.formatDate(now, 'UTC', 'yyyyMMdd');

  // Split path from query string — SigV4 canonical request requires them separately
  const qIdx      = fullPath.indexOf('?');
  const uriPath   = qIdx >= 0 ? fullPath.slice(0, qIdx) : fullPath;
  const rawQuery  = qIdx >= 0 ? fullPath.slice(qIdx + 1) : '';
  const canonQuery = rawQuery
    ? rawQuery.split('&').map(pair => {
        const eq = pair.indexOf('=');
        const k  = eq >= 0 ? pair.slice(0, eq) : pair;
        const v  = eq >= 0 ? pair.slice(eq + 1) : '';
        return encodeURIComponent(decodeURIComponent(k)) + '=' + encodeURIComponent(decodeURIComponent(v));
      }).sort().join('&')
    : '';

  // host must be signed but UrlFetchApp rejects it as a custom header — GAS sets it automatically
  const signHdrs = { 'host': host, 'x-amz-access-token': token, 'x-amz-date': amzDate };
  const keys = Object.keys(signHdrs).sort();
  const canonHdrs  = keys.map(k => k + ':' + signHdrs[k]).join('\n') + '\n';
  const signedHdrs = keys.join(';');

  const canonReq = ['GET', uriPath, canonQuery, canonHdrs, signedHdrs, sha256Hex_('')].join('\n');
  const scope    = `${dateStamp}/${region}/execute-api/aws4_request`;
  const sts      = ['AWS4-HMAC-SHA256', amzDate, scope, sha256Hex_(canonReq)].join('\n');
  const sig      = hmacHex_(signingKey_(secretKey, dateStamp, region), sts);
  const auth     = `AWS4-HMAC-SHA256 Credential=${accessKey}/${scope}, SignedHeaders=${signedHdrs}, Signature=${sig}`;

  return {
    url:                endpoint + fullPath,
    method:             'get',
    headers:            { 'x-amz-access-token': token, 'x-amz-date': amzDate, 'Authorization': auth },
    muteHttpExceptions: true,
  };
}

function spApiGet_(endpoint, region, cred, fullPath, tokenOverride) {
  const req = spApiBuildFetch_(endpoint, region, cred, fullPath, tokenOverride);
  const res = UrlFetchApp.fetch(req.url, req);
  return { status: res.getResponseCode(), body: res.getContentText() };
}

// ── SP-API POST (for Tokens API) ──────────────────────────────────────────────
function spApiPost_(endpoint, region, cred, path, body) {
  const props     = scriptProps_();
  const accessKey = props['AWS_ACCESS_KEY_ID'];
  const secretKey = props['AWS_SECRET_ACCESS_KEY'];
  const token     = getLwaToken_(cred);
  const host      = endpoint.replace('https://', '');

  const now       = new Date();
  const amzDate   = Utilities.formatDate(now, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp = Utilities.formatDate(now, 'UTC', 'yyyyMMdd');

  const bodyStr    = JSON.stringify(body);
  const payloadHash = sha256Hex_(bodyStr);

  const signHdrs  = { 'content-type': 'application/json', 'host': host, 'x-amz-access-token': token, 'x-amz-date': amzDate };
  const keys      = Object.keys(signHdrs).sort();
  const canonHdrs  = keys.map(k => k + ':' + signHdrs[k]).join('\n') + '\n';
  const signedHdrs = keys.join(';');

  const canonReq = ['POST', path, '', canonHdrs, signedHdrs, payloadHash].join('\n');
  const scope    = `${dateStamp}/${region}/execute-api/aws4_request`;
  const sts      = ['AWS4-HMAC-SHA256', amzDate, scope, sha256Hex_(canonReq)].join('\n');
  const sig      = hmacHex_(signingKey_(secretKey, dateStamp, region), sts);
  const auth     = `AWS4-HMAC-SHA256 Credential=${accessKey}/${scope}, SignedHeaders=${signedHdrs}, Signature=${sig}`;

  const res = UrlFetchApp.fetch(endpoint + path, {
    method:             'post',
    headers:            { 'Content-Type': 'application/json', 'x-amz-access-token': token, 'x-amz-date': amzDate, 'Authorization': auth },
    payload:            bodyStr,
    muteHttpExceptions: true,
  });
  return { status: res.getResponseCode(), body: res.getContentText() };
}

// ── Restricted Data Token for getOrderItems ───────────────────────────────────
function getRdt_(endpoint, region, cred, orderId) {
  const r = spApiPost_(endpoint, region, cred, '/tokens/2021-03-01/restrictedDataToken', {
    restrictedResources: [
      { method: 'GET', path: `/orders/v0/orders/${orderId}/items` },
      { method: 'GET', path: `/orders/v0/orders/${orderId}/buyerInfo`, dataElements: ['buyerInfo'] },
    ],
  });
  if (r.status !== 200) return { token: null, status: r.status, error: r.body };
  try {
    const t = JSON.parse(r.body).restrictedDataToken || null;
    return { token: t, status: r.status, error: null };
  } catch { return { token: null, status: r.status, error: r.body }; }
}

// ── Buyer purchase + refund stats (last 2 years, up to 500 orders) ───────────
// Returns { totalPurchases, totalRefunds }
// totalRefunds = orders with a Finances API RefundEventList entry (post-shipment
// refunds are NOT reflected in OrderStatus — 'Canceled' is pre-fulfillment only).
// Cached 5 min per buyer email — shared across agents on the same ticket.
function fetchBuyerPurchaseStats_(endpoint, region, cred, salesChannel, buyerEmail) {
  const mpId = marketplaceId_(salesChannel);
  if (!mpId || !buyerEmail) return null;

  const cache    = CacheService.getScriptCache();
  const cacheKey = 'bstat_' + Utilities.base64Encode(buyerEmail).replace(/[+/=]/g, '').slice(0, 50);
  const hit      = cache.get(cacheKey);
  if (hit) { try { return JSON.parse(hit); } catch(_) {} }

  const createdAfter = new Date(Date.now() - 2 * 365.25 * 24 * 3600 * 1000).toISOString().slice(0, 19) + 'Z';
  let totalPurchases = 0;
  let nextToken      = null;
  let page           = 0;
  const orderIds     = [];

  do {
    const path = nextToken
      ? `/orders/v0/orders?NextToken=${encodeURIComponent(nextToken)}`
      : `/orders/v0/orders?MarketplaceIds=${encodeURIComponent(mpId)}&BuyerEmail=${encodeURIComponent(buyerEmail)}&CreatedAfter=${encodeURIComponent(createdAfter)}&MaxResultsPerPage=100`;
    const r = spApiGet_(endpoint, region, cred, path);
    if (r.status !== 200) break;
    try {
      const d      = JSON.parse(r.body);
      const orders = d.payload?.Orders || [];
      totalPurchases += orders.length;
      orders.forEach(o => { if (o.AmazonOrderId) orderIds.push(o.AmazonOrderId); });
      nextToken       = d.payload?.NextToken || null;
    } catch { break; }
    page++;
  } while (nextToken && page < 5);

  const totalRefunds = fetchBuyerRefundCount_(endpoint, region, cred, orderIds);
  const result = { totalPurchases, totalRefunds };
  try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(_) {}
  return result;
}

// ── Count refunded orders via Finances API ────────────────────────────────────
// Uses GET /finances/v0/orders/{id}/financialEvents, batched 30 at a time via
// UrlFetchApp.fetchAll() (respects Finances API burst limit of 30).
// Per-order results cached 1 hr — refund status for Shipped orders is immutable.
// Returns 0 gracefully if Finances API role is not enabled (403).
function fetchBuyerRefundCount_(endpoint, region, cred, orderIds) {
  if (!orderIds || !orderIds.length) return 0;
  const cache = CacheService.getScriptCache();
  const BATCH = 30;
  let count = 0;
  const uncached = [];

  for (const id of orderIds) {
    const hit = cache.get('fin_' + id);
    if (hit !== null) { if (hit === '1') count++; }
    else uncached.push(id);
  }

  for (let i = 0; i < uncached.length; i += BATCH) {
    const batch = uncached.slice(i, i + BATCH);
    const requests = batch.map(id =>
      spApiBuildFetch_(endpoint, region, cred, `/finances/v0/orders/${id}/financialEvents`));
    let results;
    try { results = UrlFetchApp.fetchAll(requests); }
    catch(_) { break; }

    let financeUnavailable = false;
    for (let j = 0; j < results.length; j++) {
      const id   = batch[j];
      const code = results[j].getResponseCode();
      if (code === 403) { financeUnavailable = true; break; }
      if (code === 200) {
        try {
          const ev = JSON.parse(results[j].getContentText()).payload?.FinancialEvents;
          const refunded = (ev?.RefundEventList?.length > 0) ||
                           (ev?.GuaranteeClaimEventList?.length > 0) ||
                           (ev?.ChargebackEventList?.length > 0);
          cache.put('fin_' + id, refunded ? '1' : '0', 3600);
          if (refunded) count++;
        } catch(_) {}
      }
    }
    if (financeUnavailable) break;
    if (i + BATCH < uncached.length) Utilities.sleep(2000);
  }

  return count;
}

// ── Fetch order + items + address + buyer ─────────────────────────────────────
// Results are cached 90 s per order ID so concurrent agents on the same ticket
// share a single SP-API round-trip instead of each making 5-6 calls.
function fetchOrderData_(orderId) {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'ord2_' + orderId;
  const hit      = cache.get(cacheKey);
  if (hit) { try { return JSON.parse(hit); } catch(_) {} }

  const result = fetchOrderDataFresh_(orderId);
  try { cache.put(cacheKey, JSON.stringify(result), 90); } catch(_) {}
  return result;
}

// Tries each SP-API region one at a time (EU -> FE/JP -> NA -> IN) until one
// has the order. Previously tried parallelizing this via UrlFetchApp.fetchAll()
// on the theory that non-EU orders were paying for wasted sequential attempts
// — reverted after real-world testing showed it made things SLOWER overall,
// not faster: it unconditionally built all 4 signed requests (and their LWA/
// props lookups) up front every time, whereas the sequential version almost
// always succeeds on the very first (EU) attempt for this account's order mix,
// so parallelizing traded a rare slow path for a much more common slower one.
// Keep this sequential; scriptProps_() memoization below is the safe win.
// Order-ID prefix → REGIONS index learned from past hits (7 days): a JP/NA/IN
// order no longer pays a wasted EU attempt first. Unknown prefix or a miss on
// the hinted region falls through to the original EU→FE→NA→IN order.
function findOrderRegion_(orderId) {
  const regionErrors = [];
  const cache   = CacheService.getScriptCache();
  const hintKey = 'oreg_' + orderId.slice(0, 3);
  const hint    = Number(cache.get(hintKey));
  const order_  = REGIONS.map((r, i) => i);
  if (hint > 0 && hint < REGIONS.length) { order_.splice(hint, 1); order_.unshift(hint); }
  for (const ri of order_) {
    const { endpoint, region, cred } = REGIONS[ri];
    let r;
    try { r = spApiGet_(endpoint, region, cred, `/orders/v0/orders/${orderId}`); }
    catch (e) { regionErrors.push(`${cred}:LWA(${e.message})`); continue; }

    if (r.status === 403 && r.body.includes('expired')) {
      CacheService.getScriptCache().remove('lwa_' + cred);
      try { r = spApiGet_(endpoint, region, cred, `/orders/v0/orders/${orderId}`); }
      catch (e) { regionErrors.push(`${cred}:LWA-retry(${e.message})`); continue; }
    }

    if (r.status !== 200) {
      const detail = (r.status === 403 && r.body.includes('expired')) ? '(auth-revoked)' : '';
      regionErrors.push(`${cred}:${r.status}${detail}`);
      continue;
    }

    const order = JSON.parse(r.body).payload || {};
    if (!order.AmazonOrderId) {
      const preview = r.body.substring(0, 120).replace(/\s+/g, ' ');
      regionErrors.push(`${cred}:200-noId(${preview})`);
      continue;
    }

    try { cache.put(hintKey, String(ri), 604800); } catch (_) {}
    return { endpoint, region, cred, order };
  }
  throw new Error('Order not found — ' + regionErrors.join(' | '));
}

function fetchOrderDataFresh_(orderId) {
  const { endpoint, region, cred, order } = findOrderRegion_(orderId);

  // This app has no buyerInfo data role and no order-items role, so the RDT
  // request always 400s and items always 403. Once that outcome is observed it
  // is cached per credential set for 6 h and those two calls are skipped
  // (identical fields returned). buyerInfo is ALWAYS still called: without an
  // RDT it still returns BuyerName for many orders (Auto-Fill's Customer Full
  // Name) — v2.7.2 wrongly skipped it too. If Amazon ever grants the roles,
  // the next uncached run picks it up.
  const cache   = CacheService.getScriptCache();
  const noRdtKey = 'nordt_v2_' + cred;
  let known = null;
  try { known = JSON.parse(cache.get(noRdtKey) || 'null'); } catch (_) {}

  let rdtResult, itemsR, addrR, buyerR;
  if (known) {
    rdtResult = { token: null, status: known.rdtStatus, error: known.rdtError };
    itemsR = { status: known.itemsStatus, body: known.itemsError };
    [addrR, buyerR] = UrlFetchApp.fetchAll([
      spApiBuildFetch_(endpoint, region, cred, `/orders/v0/orders/${orderId}/address`),
      spApiBuildFetch_(endpoint, region, cred, `/orders/v0/orders/${orderId}/buyerInfo`),
    ]).map(res => ({ status: res.getResponseCode(), body: res.getContentText() }));
  } else {
    rdtResult = getRdt_(endpoint, region, cred, orderId);
    const rdtToken  = rdtResult.token || undefined;

    // Fire items + address + buyerInfo in parallel — saves ~600 ms vs sequential
    [itemsR, addrR, buyerR] = UrlFetchApp.fetchAll([
      spApiBuildFetch_(endpoint, region, cred, `/orders/v0/orders/${orderId}/items`,     rdtToken),
      spApiBuildFetch_(endpoint, region, cred, `/orders/v0/orders/${orderId}/address`),
      spApiBuildFetch_(endpoint, region, cred, `/orders/v0/orders/${orderId}/buyerInfo`, rdtToken),
    ]).map(res => ({ status: res.getResponseCode(), body: res.getContentText() }));

    if (!rdtResult.token && rdtResult.status === 400 && itemsR.status === 403) {
      try {
        cache.put(noRdtKey, JSON.stringify({
          rdtStatus: rdtResult.status, rdtError: rdtResult.error,
          itemsStatus: itemsR.status, itemsError: itemsR.body,
        }), 21600);
      } catch (_) {}
    }
  }

  const buyer = buyerR.status === 200 ? JSON.parse(buyerR.body).payload || {} : {};
  const stats = fetchBuyerPurchaseStats_(endpoint, region, cred, order.SalesChannel, buyer.BuyerEmail || null);

  return {
    order,
    items:          itemsR.status === 200 ? JSON.parse(itemsR.body).payload?.OrderItems || [] : [],
    itemsStatus:    itemsR.status,
    itemsError:     itemsR.body,
    rdtStatus:      rdtResult.status,
    rdtError:       rdtResult.error,
    address:        addrR.status === 200 ? JSON.parse(addrR.body).payload?.ShippingAddress || {} : {},
    buyer,
    orderCount:     stats ? stats.totalPurchases : null,
    totalPurchases: stats ? stats.totalPurchases : null,
    totalRefunds:   stats ? stats.totalRefunds   : null,
    region,
  };
}


// ── ASIN marketplace availability (market spreadsheet) ───────────────────────
// Returns array of country codes (e.g. ['DE','UK']) where the ASIN is selling.
// A row is counted only if it contains the ASIN AND no cell contains '단종'.
function colToLetter_(col) {
  let letter = '';
  for (let c = col + 1; c > 0; ) {
    const rem = (c - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    c = Math.floor((c - 1) / 26);
  }
  return letter;
}

// ── Sheet row caching helpers ─────────────────────────────────────────────────
// GAS ScriptCache has a 100 KB per-entry limit. We only write if the serialised
// JSON fits in 95 KB; otherwise we fall through to a live read every time.
const SHEET1_CACHE_KEY = 'prod_sheet1_v1';
const SHEET2_CACHE_KEY = 'prod_sheet2_v1';
const SHEET_ROW_TTL    = 3600;
const MKT_INDEX_CACHE_KEY = 'mkt_idx_v1';
const MKT_INDEX_TTL    = 14400;

function loadSheetRows_(ssId, sheetName, cacheKey) {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const json = JSON.stringify(rows);
  if (json.length < 95000) cache.put(cacheKey, json, SHEET_ROW_TTL);
  return rows;
}

// Builds a full ASIN → [{name, gid, cell}] index across all market sheets and
// caches it for 4 hours. Returns null if the index is too large to cache (rare).
function getMktIndex_() {
  const cache = CacheService.getScriptCache();
  const hit   = cache.get(MKT_INDEX_CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const ss    = SpreadsheetApp.openById(MARKET_SS_ID);
  const index = {};
  for (const sheetName of MARKET_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const gid  = sheet.getSheetId();
    const rows = sheet.getDataRange().getValues();
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].map(c => String(c));
      for (let col = 0; col < cells.length; col++) {
        const cell = cells[col];
        if (/^B[A-Z0-9]{9}$/.test(cell)) {
          if (!cells.some(c => c.includes('단종'))) {
            if (!index[cell]) index[cell] = [];
            index[cell].push({ name: sheetName, gid, cell: colToLetter_(col) + (r + 1) });
          }
          break;
        }
      }
    }
  }
  const json = JSON.stringify(index);
  if (json.length < 95000) cache.put(MKT_INDEX_CACHE_KEY, json, MKT_INDEX_TTL);
  return index;
}

function checkMarketplaces_(asin) {
  const idx = getMktIndex_();
  return idx[asin] || [];
}

// ── Google Sheet ASIN lookup ──────────────────────────────────────────────────
function lookupAsin_(asin) {
  const data    = loadSheetRows_(SHEET_ID, SHEET_NAME, SHEET1_CACHE_KEY);
  if (!data.length) return null;
  const headers = data[0];
  const asinIdx = headers.indexOf('ASIN');
  if (asinIdx < 0) throw new Error('ASIN column not found in sheet');

  const match = data.slice(1).find(row => String(row[asinIdx]) === asin);
  if (!match) return null;

  const result = {};
  PRODUCT_COLS.forEach(col => {
    const i = headers.indexOf(col);
    if (i >= 0) result[col] = match[i] !== undefined ? String(match[i]) : '';
  });
  return result;
}

// ── Market spreadsheet — Data sheet ASIN lookup (source 2) ───────────────────
function lookupAsin2_(asin) {
  const data = loadSheetRows_(MARKET_SS_ID, 'Data', SHEET2_CACHE_KEY);
  if (!data.length) return null;
  const headers = data[0];
  const asinIdx = headers.indexOf('ASIN');
  if (asinIdx < 0) return null;

  const match = data.slice(1).find(row => String(row[asinIdx]) === asin);
  if (!match) return null;

  const result = {};
  PRODUCT_COLS.forEach(col => {
    const i = headers.indexOf(col);
    if (i >= 0) result[col] = match[i] !== undefined ? String(match[i]) : '';
  });
  return result;
}

// ── Market country sheets — partial lookup (col A = 기종명, col B = 모델명) ───
// Returns a partial product object; other fields will be filled by Amazon page.
function lookupAsinFromMarket_(asin) {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'mkt_partial_' + asin;
  const hit      = cache.get(cacheKey);
  if (hit !== null) return hit === '__null__' ? null : JSON.parse(hit);

  const ss = SpreadsheetApp.openById(MARKET_SS_ID);
  for (const sheetName of MARKET_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data = sheet.getDataRange().getValues();
    for (const rowData of data) {
      const cells = rowData.map(c => String(c));
      if (cells.some(c => c === asin)) {
        const partial = {
          'SKU': '', '모델명': String(rowData[1] || ''), '브랜드': '',
          '제조사명': '', '기종명': String(rowData[0] || ''), '색상명': '',
          '대분류': '', '생산업체': '', '원산지정보': '',
        };
        cache.put(cacheKey, JSON.stringify(partial), 3600);
        return partial;
      }
    }
  }

  cache.put(cacheKey, '__null__', 3600);
  return null;
}

// ── Full lookup chain: sheet1 → sheet2 Data → market country sheets ──────────
// Always checks both sheet1 and sheet2 for allSources; uses priority for product.
// Results are cached per-ASIN for 30 minutes to avoid repeated full-sheet reads.
function lookupAsinAll_(asin) {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'asin_all_' + asin;
  const hit      = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);

  const sheet1 = lookupAsin_(asin);
  const sheet2 = lookupAsin2_(asin);

  let product = sheet1 || sheet2;
  let productSource = sheet1 ? 'sheet1' : sheet2 ? 'sheet2' : null;

  if (!product) {
    product = lookupAsinFromMarket_(asin);
    if (product) productSource = 'market';
  }

  const result = { product, productSource, allSources: { sheet1: sheet1 || null, sheet2: sheet2 || null } };
  try { cache.put(cacheKey, JSON.stringify(result), 1800); } catch (_) {}
  return result;
}

// ── Precomputed product index ─────────────────────────────────────────────────
// Every ?asin= request used to rebuild the market index from all 10 country
// sheets (~3 MB, 25k ASIN cells) because the index (~1 MB) never fit the 95 KB
// cache-entry limit — ~7-8 s per lookup, ~16 s for an ASIN in no sheet. Same
// for ASIN Master (1.5 MB). This builds ONE pass over all sources into per-ASIN
// records, split into PIDX_BUCKETS hash buckets (~20 KB each) so each lookup is
// a single cache read. Same output as lookupAsinAll_ + checkMarketplaces_:
//   s1/s2 = sheet1/sheet2 product values in PRODUCT_COLS order (null = col absent)
//   p     = market partial [기종명, 모델명] (first sheet/row containing the ASIN)
//   m     = marketplaces [[sheetName, gid, cell], ...]
// A full build reads ~4.5 MB of sheets and takes ~20-40 s, so it NEVER runs
// inside an agent's request: refreshProductIndex (15-min trigger, see
// setupProductIndexTrigger) and keepWarm rebuild it in the background.
// The product sheets only change weekly-to-monthly, so the 15-min trigger
// just checks the age and rebuilds every PIDX_REFRESH_MS (4 h — inside the
// 6 h ScriptCache TTL, so the index never expires). After editing a product
// sheet, run forceRefreshProductIndex to publish the change immediately.
// Requests only read it. Missing / older than PIDX_MAX_AGE_MS (trigger not
// running) / oversized bucket / non-B ASIN → null → caller uses the original
// live-read path, so a request is never slower than before this change.
const PIDX_PREFIX      = 'pidx_v1_';
const PIDX_META_KEY    = PIDX_PREFIX + 'meta';
const PIDX_BUCKETS     = 128;
const PIDX_TTL         = 21600;
const PIDX_MAX_AGE_MS  = 6 * 60 * 60 * 1000;
const PIDX_REFRESH_MS  = 4 * 60 * 60 * 1000; // rebuild once older than this (trigger checks every 15 min)
const ASIN_CELL_RE     = /^B[A-Z0-9]{9}$/;

function pidxBucketKey_(asin) {
  let h = 0;
  for (let i = 0; i < asin.length; i++) h = (h * 31 + asin.charCodeAt(i)) | 0;
  return PIDX_PREFIX + (((h % PIDX_BUCKETS) + PIDX_BUCKETS) % PIDX_BUCKETS);
}

function pidxProductRows_(ssId, sheetName) {
  const sheet = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (!data.length) return map;
  const headers = data[0];
  const asinIdx = headers.indexOf('ASIN');
  if (asinIdx < 0) return null;
  const colIdx = PRODUCT_COLS.map(c => headers.indexOf(c));
  for (let r = 1; r < data.length; r++) {
    const a = String(data[r][asinIdx]);
    if (map[a]) continue; // first match wins, same as .find()
    map[a] = colIdx.map(i => i < 0 ? null : (data[r][i] !== undefined ? String(data[r][i]) : ''));
  }
  return map;
}

// Returns { bucketKey: bucketObject } on success, null on failure.
function buildProductIndex_() {
  const s1 = pidxProductRows_(SHEET_ID, SHEET_NAME);
  if (!s1) return null; // ASIN column missing → let live path raise its usual error
  const s2 = pidxProductRows_(MARKET_SS_ID, 'Data') || {};

  const mkt = {}, partial = {};
  const ss = SpreadsheetApp.openById(MARKET_SS_ID);
  for (const sheetName of MARKET_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const gid  = sheet.getSheetId();
    const rows = sheet.getDataRange().getValues();
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].map(c => String(c));
      // Marketplaces: first ASIN-shaped cell of the row, row skipped if 단종 (getMktIndex_)
      for (let col = 0; col < cells.length; col++) {
        if (ASIN_CELL_RE.test(cells[col])) {
          if (!cells.some(c => c.includes('단종'))) {
            (mkt[cells[col]] = mkt[cells[col]] || []).push([sheetName, gid, colToLetter_(col) + (r + 1)]);
          }
          break;
        }
      }
      // Partial: first sheet/row containing the ASIN in any cell (lookupAsinFromMarket_)
      for (const c of cells) {
        if (c.length === 10 && c[0] === 'B' && !partial[c] && ASIN_CELL_RE.test(c)) partial[c] = [String(rows[r][0] || ''), String(rows[r][1] || '')];
      }
    }
  }

  const buckets = {};
  const all = new Set([...Object.keys(s1), ...Object.keys(s2), ...Object.keys(mkt), ...Object.keys(partial)]);
  all.forEach(a => {
    if (!ASIN_CELL_RE.test(a)) return;
    const rec = {};
    if (s1[a]) rec.s1 = s1[a];
    if (s2[a]) rec.s2 = s2[a];
    if (partial[a]) rec.p = partial[a];
    if (mkt[a]) rec.m = mkt[a];
    const k = pidxBucketKey_(a);
    (buckets[k] = buckets[k] || {})[a] = rec;
  });
  for (let i = 0; i < PIDX_BUCKETS; i++) buckets[PIDX_PREFIX + i] = buckets[PIDX_PREFIX + i] || {};

  const toPut = {};
  for (const k of Object.keys(buckets)) {
    const json = JSON.stringify(buckets[k]);
    if (Utilities.newBlob(json).getBytes().length > 95000) {
      Logger.log('buildProductIndex_: bucket ' + k + ' too large — index not stored');
      return null;
    }
    toPut[k] = json;
  }
  const builtAt = Date.now();
  toPut[PIDX_META_KEY] = JSON.stringify({ builtAt, asins: all.size });
  CacheService.getScriptCache().putAll(toPut, PIDX_TTL);
  const forWorker = Object.assign({}, toPut);
  delete forWorker[PIDX_META_KEY];
  pushIndexToWorker_(forWorker, builtAt);
  return buckets;
}

// ── GCX Reply Worker (Cloudflare) — product-index mirror ─────────────────────
// The panel reads product data from the Worker first (GCXReply_Worker/), which
// serves it from this same index. GAS stays the source of truth: after every
// rebuild the buckets are POSTed to the Worker. WORKER_URL / WORKER_PUSH_KEY
// live in Script Properties (set once via the setWorkerConfig doPost action —
// never in this file, which is in a public repo). Failures only log: the
// Worker falls back to asking GAS when its copy is missing or >6 h old.
function pushIndexToWorker_(bucketJsonByKey, builtAt) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('WORKER_URL'), key = props.getProperty('WORKER_PUSH_KEY');
  if (!url || !key) return false;
  try {
    const r = UrlFetchApp.fetch(url.replace(/\/$/, '') + '/admin/pidx', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-push-key': key },
      payload: JSON.stringify({ builtAt, buckets: bucketJsonByKey }),
    });
    if (r.getResponseCode() !== 200) Logger.log('pushIndexToWorker_: HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    return r.getResponseCode() === 200;
  } catch (e) { Logger.log('pushIndexToWorker_: ' + e.message); return false; }
}

// Re-push the index currently in ScriptCache (no sheet reads) — used right
// after the Worker is configured so it doesn't wait for the next rebuild.
function pushCachedIndexToWorker_() {
  const cache = CacheService.getScriptCache();
  const keys = [];
  for (let i = 0; i < PIDX_BUCKETS; i++) keys.push(PIDX_PREFIX + i);
  const got = cache.getAll(keys.concat([PIDX_META_KEY]));
  if (!got[PIDX_META_KEY] || keys.some(k => got[k] == null)) return false;
  const buckets = {};
  keys.forEach(k => { buckets[k] = got[k]; });
  return pushIndexToWorker_(buckets, JSON.parse(got[PIDX_META_KEY]).builtAt);
}

// doPost { action:'setWorkerConfig', workerUrl, pushKey, proof } where
// proof = hex(sha256(AWS_SECRET_ACCESS_KEY + '|' + pushKey + '|' + workerUrl)).
// Only someone who already holds this project's SP-API secret can compute it,
// so the endpoint can stay public without exposing anything new.
function setWorkerConfig_(body) {
  const secret = scriptProps_()['AWS_SECRET_ACCESS_KEY'] || '';
  const expected = sha256Hex_(secret + '|' + body.pushKey + '|' + body.workerUrl);
  if (!secret || !body.pushKey || !body.workerUrl || body.proof !== expected) return { ok: false, error: 'bad proof' };
  PropertiesService.getScriptProperties().setProperties({ WORKER_URL: body.workerUrl, WORKER_PUSH_KEY: body.pushKey });
  return { ok: true, pushed: pushCachedIndexToWorker_() };
}

// Keeps overlapping trigger runs (keepWarm + refreshProductIndex) from both
// rebuilding. Deliberately NOT LockService.getScriptLock(): a rebuild takes
// 20-40 s, and holding the script lock that long made claimAbmRelay_/
// claimAbmSend_ (waitLock 5-10 s) fail with lock_timeout during rebuilds.
// A short-lived cache flag is enough here — a rare double rebuild is harmless.
function rebuildProductIndexLocked_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(PIDX_PREFIX + 'building')) return null;
  cache.put(PIDX_PREFIX + 'building', '1', 120);
  try { return buildProductIndex_(); }
  finally { cache.remove(PIDX_PREFIX + 'building'); }
}

function refreshProductIndexIfStale_() {
  try {
    const meta = CacheService.getScriptCache().get(PIDX_META_KEY);
    const age  = meta ? Date.now() - JSON.parse(meta).builtAt : Infinity;
    if (age > PIDX_REFRESH_MS) rebuildProductIndexLocked_();
  } catch (e) { Logger.log('refreshProductIndexIfStale_: ' + e.message); }
}

// Trigger entry point: rebuilds only when the index is older than PIDX_REFRESH_MS.
function refreshProductIndex() {
  refreshProductIndexIfStale_();
}

// Run manually after editing ASIN Master / market sheets — rebuilds right now.
function forceRefreshProductIndex() {
  const t0 = Date.now();
  const b  = rebuildProductIndexLocked_();
  Logger.log(b ? `Product index rebuilt in ${Date.now() - t0} ms` : 'Product index NOT rebuilt (lock busy or oversized bucket)');
}

// Run ONCE from the GAS editor: installs the 15-min refreshProductIndex trigger
// and builds the index immediately.
function setupProductIndexTrigger() {
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshProductIndex');
  if (!existing.length) ScriptApp.newTrigger('refreshProductIndex').timeBased().everyMinutes(15).create();
  Logger.log('refreshProductIndex trigger ' + (existing.length ? 'already exists' : 'created — fires every 15 minutes'));
  forceRefreshProductIndex();
}

function pidxExpand_(vals) {
  if (!vals) return null;
  const o = {};
  PRODUCT_COLS.forEach((c, i) => { if (vals[i] !== null) o[c] = vals[i]; });
  return o;
}

// Returns the doGet product fields from the index, or null if unavailable.
function lookupFromIndex_(asin) {
  if (!ASIN_CELL_RE.test(asin)) return null;
  const cache = CacheService.getScriptCache();
  const k     = pidxBucketKey_(asin);
  const got   = cache.getAll([PIDX_META_KEY, k]);
  if (!got[PIDX_META_KEY]) return null;
  if (!got[k]) { cache.remove(PIDX_META_KEY); return null; } // bucket evicted early → next 15-min trigger rebuilds
  if (Date.now() - JSON.parse(got[PIDX_META_KEY]).builtAt > PIDX_MAX_AGE_MS) return null;
  const bucket = JSON.parse(got[k]);
  if (!bucket) return null;

  const rec    = bucket[asin] || {};
  const sheet1 = pidxExpand_(rec.s1);
  const sheet2 = pidxExpand_(rec.s2);
  let product = sheet1 || sheet2;
  let productSource = sheet1 ? 'sheet1' : sheet2 ? 'sheet2' : null;
  if (!product && rec.p) {
    product = {
      'SKU': '', '모델명': rec.p[1], '브랜드': '',
      '제조사명': '', '기종명': rec.p[0], '색상명': '',
      '대분류': '', '생산업체': '', '원산지정보': '',
    };
    productSource = 'market';
  }
  return {
    product, productSource,
    allSources: { sheet1: sheet1 || null, sheet2: sheet2 || null },
    marketplaces: (rec.m || []).map(([name, gid, cell]) => ({ name, gid, cell })),
  };
}

// Index first; original live-read chain as fallback.
function lookupProductFull_(asin) {
  try {
    const hit = lookupFromIndex_(asin);
    if (hit) return hit;
  } catch (e) { Logger.log('lookupFromIndex_ failed, using live path: ' + e.message); }
  const lu = lookupAsinAll_(asin);
  let marketplaces;
  try { marketplaces = checkMarketplaces_(asin); } catch { marketplaces = []; }
  return { product: lu.product, productSource: lu.productSource, allSources: lu.allSources, marketplaces };
}

// ── AI 인입사유 functions ──────────────────────────────────────────────────────
function inferReason_(text, category) {
  text     = String(text     || '').trim().toLowerCase();
  category = String(category || '').trim();
  if (!text) return '';

  const cacheKey = DR_CACHE_VERSION + Utilities.base64Encode(text + '|' + category).slice(0, 100);
  const cache    = CacheService.getScriptCache();
  const hit      = cache.get(cacheKey);
  if (hit) { Logger.log('inferReason cache hit: ' + hit); return hit; }

  const { rawList, list, enrichedList } = loadDefectDataDR_(category);
  Logger.log(`loadDefectData: rawList.length=${rawList.length}, category="${category}"`);
  if (!rawList.length) return '';

  let output = '';
  for (const model of GEMINI_MODELS_DR) {
    const r = callGeminiDR_(text, enrichedList, model);
    Logger.log(`Gemini [${model}]: "${r}"`);
    if (r && r.trim()) { output = r.replace(/["'\n\r]/g, '').trim(); break; }
  }
  Logger.log(`gemini output: "${output}"`);
  if (!output) return '';

  const norm = drNorm_(output);
  const idx  = list.findIndex(v => v === norm);
  if (idx !== -1) { cache.put(cacheKey, rawList[idx], DR_CACHE_TTL); return rawList[idx]; }

  const li = list.findIndex(v => norm.includes(v) || v.includes(norm));
  if (li  !== -1) { cache.put(cacheKey, rawList[li],  DR_CACHE_TTL); return rawList[li]; }

  Logger.log(`no match found for norm="${norm}", list=${JSON.stringify(list.slice(0,5))}`);
  return '';
}

function loadDefectDataDR_(category) {
  const cache     = CacheService.getScriptCache();
  const rowsKey   = 'dr_defect_list_v1';
  let rows;
  const rowsHit   = cache.get(rowsKey);
  if (rowsHit) {
    rows = JSON.parse(rowsHit);
  } else {
    const sh = SpreadsheetApp.openById(DEFECT_SS_ID).getSheetByName(DEFECT_SHEET_NAME);
    if (!sh) { Logger.log('loadDefectData: sheet not found'); return { rawList: [], list: [], enrichedList: [] }; }
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { rawList: [], list: [], enrichedList: [] };
    rows = sh.getRange(2, 1, lastRow - 1, 3).getValues();
    const json = JSON.stringify(rows);
    if (json.length < 95000) cache.put(rowsKey, json, 21600);
  }

  let filtered = rows.filter(r => (!category || String(r[0]).trim() === category) && r[1]);
  if (!filtered.length && category) {
    Logger.log(`loadDefectData: no rows for category="${category}", using all`);
    filtered = rows.filter(r => r[1]);
  }

  const rawList      = filtered.map(r => String(r[1]).trim());
  const list         = rawList.map(drNorm_);
  const enrichedList = filtered.map(r => {
    const label = String(r[1]).trim();
    const desc  = String(r[2] || '').trim().slice(0, 80);
    return desc ? `${label}: ${desc}` : label;
  });
  return { rawList, list, enrichedList };
}

function callGeminiDR_(text, enrichedList, model) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) { Logger.log('callGeminiDR: GEMINI_API_KEY not set'); return ''; }
    const prompt =
      `You are classifying a customer service ticket into ONE predefined category.\n\nThe input may contain:\n- The customer's original message (English, Spanish, French, German, or other languages)\n- A Korean summary written by the support team (after "[CS요약]")\n\nUse ALL available context to determine the most accurate category.\n\nCategories:\n${enrichedList.join('\n')}\n\nRules:\n- Return ONLY the exact label text from the list above\n- No explanation, punctuation, quotes, or markdown\n\nInput:\n${text}`;
    const fetchOpts = {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0, maxOutputTokens: 150 },
      }),
    };
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    let res = UrlFetchApp.fetch(url, fetchOpts);
    if (res.getResponseCode() === 503) {
      Utilities.sleep(2500);
      res = UrlFetchApp.fetch(url, fetchOpts);
    }
    if (res.getResponseCode() !== 200) {
      Logger.log(`callGeminiDR [${model}] HTTP ${res.getResponseCode()}: ${res.getContentText().slice(0, 200)}`);
      return '';
    }
    const json  = JSON.parse(res.getContentText());
    const parts = (json.candidates?.[0]?.content?.parts || []).filter(p => !p.thought);
    const out   = parts.map(p => p.text || '').join('').trim();
    if (!out) Logger.log(`callGeminiDR [${model}] 200 but empty — raw: ${res.getContentText().slice(0, 200)}`);
    return out;
  } catch (e) { Logger.log('callGeminiDR error: ' + e.message); return ''; }
}

function testInferReason() {
  Logger.log('=== testInferReason ===');
  Logger.log(JSON.stringify({ reason: inferReason_('this case is too thick and heavy', '') }));
}

function drKeyword_(text) {
  if (/heavy|bulky/.test(text))              return '두꺼움';
  if (text.includes('yellow'))               return '황변';
  if (text.includes('button'))               return '버튼불량';
  if (/attach|difficult/.test(text))         return '부착어려움';
  if (/scratch|scratched/.test(text))        return '스크래치';
  return '';
}

function drNorm_(text) {
  return String(text || '').toLowerCase().replace(/\s+/g, '').replace(/[()]/g, '').replace(/[^a-z0-9가-힣]/gi, '').trim();
}

function updateFeedbackSheet() {
  const ss    = SpreadsheetApp.openById('1v0cTiDvFM060e3pSqLHMhAiIBEsP-6gMDH_Q5eM7Yjo');
  const sheet = ss.getSheetByName('GCX Reply 피드백');
  const F = 6, J = 10; // GCX 피드백 col, 빌드 col

  // [sheetRow, feedbackText, buildValue_or_null]
  // buildValue null = keep existing (already filled by team)
  const UPDATES = [
    [8,
      '【Before】 토크나이저가 숫자 "7"을 인접 단어와 병합하여 Jaccard 매칭 오류 → "Galaxy Z Fold" 선택\n【After v2.7.1】 camelCase·숫자 경계 분리 토크나이저 개선 → "Galaxy Z Fold 7" 정확 선택\n✅ 테스트 티켓: #1000148979 (ASIN B0F1B7GKC9)',
      null],
    [9,
      '【Before】 동일 원인 → "Galaxy Z Flip" 선택\n【After v2.7.1】 동일 수정 사항 적용 → "Galaxy Z Flip 7" 정확 선택\n✅ 테스트 티켓: #1000149667 (ASIN B0F1C7LBTW)',
      'v2.7.1'],
    [10,
      '【Before】 제품 정보 시트 모델명이 "Enzo Aramid T"로 누락 → 불일치 발생\n【After v2.7.7】 제품 정보 시트 ACS09959 모델명 → "Enzo Aramid T MagFit"으로 직접 수정 (코드 아닌 시트 데이터 수정)\n✅ 테스트 티켓: #1000149002 (ASIN B0FD22YWNK)',
      'v2.7.7'],
    [11,
      '【Before】 제품 정보 시트 모델명이 "Glas.tR EZ Fit"으로만 등록 → "Glas.tR EZ Fit FC"로 오매칭\n【After v2.7.7】 시트 AGL07929 모델명 → "Glas.tR EZ Fit Privacy"으로 수정\n✅ 테스트 티켓: #1000149064 (ASIN B0D84JFDCB)',
      'v2.7.7'],
    [12,
      '【Before】 제품 정보 시트 모델명이 "Glas.tR EZ Fit"으로만 등록 → "Glas.tR EZ Fit FC"로 오매칭\n【After v2.7.7】 시트 AGL07928 모델명 → "Glas.tR EZ Fit Slim"으로 수정\n✅ 테스트 티켓: #1000149249, #1000149638 (ASIN B0D84YX465)',
      'v2.7.7'],
    [13,
      '【Before】 패널 기본 위치가 top: 56px + 저장된 위치 최솟값 4px로 인해 Chrome UI(탭바·북마크바) 아래 패널 헤더가 가려짐\n【After v2.7.7】 CSS 기본값 top: 72px, 위치 저장 최솟값 72px로 상향 → 헤더가 항상 화면 안에 표시\n✅ 다수 티켓에서 드래그 가능 확인',
      'v2.7.7'],
    [14,
      '【Before】 SP-API SellerSKU 원시값(바코드 8809613760408)을 문의SKU 필드에 입력\n【After v2.7.3】 제품 정보 시트 SKU(ACS02957) 최우선 적용 / 8자리 이상 숫자는 바코드로 간주하여 자동 제외\n✅ 테스트 티켓: #1000149643 (ASIN B08ZH4WD9D)',
      'v2.7.3'],
    [15,
      '【Before】 UTC 기준 날짜 표시 → 인도(IST, UTC+5:30) 자정 이후 주문 시 날짜가 하루 당겨 표시됨\n【After v2.7.0】 PURCHASE_TZ_ 맵 추가 → IN/JP/SG/AU/KR 현지 타임존 기준 날짜 변환\n✅ 테스트 티켓: #1000149654, #1000149668',
      null],
    [16,
      '【Before】 Zendesk SPA 재사용 React 인스턴스로 이전 티켓 데이터 잔류 → 잘못된 디바이스 표시\n【After v2.7.3】 clearAllZdFields_()로 티켓 이동 시 ZD 필드 전체 초기화\n✅ 테스트 티켓: #1000149692 (ASIN B0D84YX465)',
      null],
    [17,
      '【Before】 동일 원인(React state 교차 오염) → 이전 티켓의 Product Name 표시\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149692 (ASIN B0D84YX465)',
      null],
    [18,
      '【Before】 React state 교차 오염 → 이전 티켓 ASIN(B0C8HB6XNK)이 현재 티켓에 표시\n【After v2.7.3】 티켓 이동 시 ASIN 포함 전체 필드 초기화\n✅ 테스트 티켓: #1000149692 (ASIN B0D84YX465)',
      'v2.7.3'],
    [19,
      '【Before】 React state 교차 오염 → 이전 티켓의 SKU/카테고리 값("3in1 - C to C cable")이 문의SKU 필드에 입력\n【After v2.7.3】 티켓 이동 시 전체 초기화 + 시트 SKU 우선 적용\n✅ 테스트 티켓: #1000149675 (ASIN B0C5D9734Q)',
      'v2.7.3'],
    [20,
      '【Before】 동일 원인(React state 교차 오염) → 이전 티켓의 SKU(PA2306)가 문의SKU 필드에 입력\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149385 (ASIN B0CSSTCLCD)',
      'v2.7.3'],
    [21,
      '【Before】 Zendesk SPA에서 티켓 이동 시 React 컴포넌트 재사용으로 이전 ZD 필드값 잔류 → Solved/Pending 처리 시 다른 고객 정보 전송\n【After v2.7.3】 Auto-Fill 확인 시에만 clearAllZdFields_() 호출 → 미사용 티켓 필드는 보존\n✅ 테스트 티켓: #1000149584',
      null],
    [22,
      '【Before】 동일 원인 → Customer Full Name이 다른 고객명으로 자동 변경\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149494',
      'v2.7.3'],
    [23,
      '【Before】 동일 원인 → Order ID가 다른 티켓 Order ID로 자동 변경\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149631',
      null],
    [24,
      '【Before】 제품 정보 시트 대분류가 "거치대/스탠드"로만 등록 → SDA로 매핑됨\n【After v2.7.7】 시트 AMP09837 대분류 → "차량용 거치대/스탠드"으로 수정 → NewBiz 정확 선택\n✅ 테스트 티켓: #1000149702 (ASIN B0FFFPGK6R)',
      'v2.7.7'],
    [25,
      '【Before】 제품 정보 시트 모델명이 "Glas.tR EZ Fit"으로 등록 → FC로 오매칭\n【After v2.7.7】 시트 AGL10819 모델명 → "Glas.tR EZ Fit Anti-Glare"으로 수정\n✅ 테스트 티켓: #1000149707 (ASIN B0FPJ6RYFN)',
      'v2.7.7'],
    [26,
      '【Before】 SP-API items 권한 없음(403)으로 SellerSKU 미수신 → rawSellerSku = "" → PAN 판별 불가 → FBA로 잘못 선택\n【After v2.7.6】 ASIN 확인 후에도 Seller Central orders-api 병렬 조회 → SellerSKU(ACS06557PAN) 확보 → PAN 감지 → PAN EU 선택\n✅ 테스트 티켓: #1000149739 (ASIN B0C5S85TD2)',
      'v2.7.6'],
    [27,
      '【Before】 Invoice 티켓 여부를 인식하지 못해 Device 필드가 제품 기종명으로 채워지거나 비어있음\n【After v2.7.7】 티켓 subject에 "invoice" 포함 시 Device dropdown에서 "INVOICE - Spigen (Cases)" 옵션 자동 선택\n✅ 테스트 티켓: #1000149788',
      'v2.7.7'],
    [28,
      '【Before】 UTC 기준 날짜 표시로 현지 타임존과 불일치\n【After v2.7.0】 PURCHASE_TZ_ 맵 적용 → 현지 기준 날짜 정확 표시\n✅ 테스트 티켓: #1000149883',
      null],
    [29,
      '【Before】 동일 원인(React state 교차 오염) → Solved 처리 시 다른 티켓(#1000149697) Order ID로 변경\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149698',
      null],
    [30,
      '【Before】 동일 원인 → Pending 처리 시 다른 티켓 Order ID·고객명으로 변경\n【After v2.7.3】 동일 수정 사항 적용\n✅ 테스트 티켓: #1000149714',
      null],
    [31,
      '【Before】 v2.7.3 clearAllZdFields_()가 모든 티켓 이동 시 무조건 호출 → Zendesk가 저장한 필드값까지 초기화됨\n【After v2.7.7】 Auto-Fill 확인 후에만 _gcrFilledThisTicket = true 설정 → 해당 플래그 true일 때만 필드 초기화. Auto-Fill 미사용 티켓은 필드값 보존\n✅ 임의 티켓에서 Auto-Fill 없이 이동 후 복귀 시 필드 정상 유지 확인',
      'v2.7.7'],
    [32,
      '【Before】 동일 원인(v2.7.3 clearAllZdFields_() 잔류) → Auto-Fill 없이 다른 티켓 이동 후 복귀 시 Customer Full Name / Purchase Date / Order ID가 다른 티켓 값으로 변경됨\n【After v2.7.7】 _gcrFilledThisTicket 플래그 도입 → Auto-Fill 확인 후에만 필드 초기화. 미사용 티켓은 필드값 보존\n✅ 테스트 티켓: #1000150001',
      null],
    [33,
      '【Before】 SC SellerSKU "PE2213IN 35w"(공백 포함 모델명)이 기존 숫자 바코드 필터를 통과 → 문의SKU 필드에 그대로 입력됨\n【After v2.8.2】 공백 포함 SellerSKU 제외 추가\n【After v2.8.3】 Spigen SKU 패턴(^[A-Z]{3}\\d{5}) 양성 일치 방식으로 변경 → "PE2213IN" 등 공백 없는 비표준 코드도 차단\n✅ 테스트 티켓: #1000150015 (ASIN B0CG8QTWP2)',
      'v2.8.3'],
    [34,
      '【Before】 동일 원인 → SC SellerSKU "PE2212IN 65w"가 문의SKU 필드에 입력됨\n【After v2.8.3】 동일 수정 사항 적용 (Spigen SKU 패턴 검증)\n✅ 테스트 티켓: #1000150108, #1000150207 (ASIN B0DKSNXLCT)',
      'v2.8.3'],
    [35,
      '【Before】 제품 정보 시트 AGL07930 모델명이 "Glas.tR EZ Fit"으로만 등록 → "Glas.tR EZ Fit FC"로 오매칭됨\n【After v2.8.2】 시트 AGL07930 모델명 → "Glas.tR EZ Fit Anti Reflection"으로 수정 (fixProductSheetData 실행)\n✅ 테스트 티켓: #1000150171 (ASIN B0D84Z9693)',
      'v2.8.2'],
    [36,
      '【Before】 SC orders-api SellerSKU 미확보(세션 만료) 또는 ACS10046에 PAN 미포함 → PAN EU 미감지 → FBA 선택\n【현황】 v2.7.6 병렬 SC 조회 적용 중 — SC 세션 유지 시 자동 감지 가능. SellerSKU 직접 확인 후 추가 조치 예정\n📋 티켓: #1000150172 (ASIN B0FD22YW2J)',
      null],
    [37,
      '【현황 - GCX Test: Fail】 소스 데이터 불일치\nASIN_Master(먼데이보드): "Glas.tR EZ Fit" / TCK 국가별 제품 리스트(아마존 직판): "Glas.tR Slim (2P)" / 젠데스크: "Glas.tR EZ Fit Slim"\n【시도】 v2.8.2 fixProductSheetData AGL07928 모델명 "Glas.tR EZ Fit Slim" 재적용 → 단, 소스 데이터 근본 불일치로 재발 가능성 있음\n📋 테스트 티켓: #1000150312 (ASIN B0D84YX465)',
      'v2.7.7'],
    [38,
      '【Before】 동일 원인 → SC SellerSKU "PE2304IN 45w"가 문의SKU 필드에 입력됨\n【After v2.8.3】 동일 수정 사항 적용 (Spigen SKU 패턴 검증)\n✅ 테스트 티켓: #1000150368 (ASIN B0DQ14CVX1)',
      'v2.8.3'],
    [39,
      '【Before】 인도(Amazon.in) SP-API LWA 토큰 만료 시 예외 미처리 → 주문 조회 전체 실패\n【After GAS 2026-06-15】 지역별 LWA 예외 개별 catch, 403+만료 시 캐시 자동 삭제 후 1회 재시도\n📋 테스트 티켓: #1000150413',
      'v2.8.4'],
    [40,
      '【Before】 SPA 티켓 이동 시 history.pushState 직후 clearAllZdFields_() 호출 → DOM이 아직 이전 티켓을 렌더링 중일 경우 해당 티켓 ZD 필드값이 지워짐 → 다른 Zendesk 탭으로 이동 후 복귀 시 Purchase Date 변경, Order ID/ASIN/SKU 공란 발생\n【After v2.19.1】 clearAllZdFields_()를 resetPanel()에서 Auto-Fill 확인 직전으로 이동 → DOM 타이밍과 무관하게 정확한 티켓 필드만 초기화\n✅ 테스트 티켓: #1000150854',
      'v2.19.1'],
    [41,
      '【Before】 제품 정보 시트 AGL10123 모델명 불일치 → "Glas.tR EZ Fit Pro"로 오매칭\n【After v2.19.1】 시트 AGL10123 모델명 → "Glas.tR Optik Pro"으로 수정\n✅ 테스트 티켓: #1000150922 (ASIN B0FFVDVB4S)',
      'v2.19.1'],
    [42,
      '【Before】 제품 정보 시트 ACS07624 모델명 불일치 → "Ultra Hybrid OneTap Metal Ring MagFit"으로 오매칭\n【After v2.19.1】 시트 ACS07624 모델명 → "Ultra Hybrid MagFit"으로 수정\n✅ 테스트 티켓: #1000150934 (ASIN B0CV71SSMQ)',
      'v2.19.1'],
    [43,
      '【Before】 제품 정보 시트 AMP08885 대분류가 "거치대/스탠드" 형태로 등록 → SDA로 매핑됨\n【After v2.19.1】 시트 AMP08885 대분류 → "차량용 거치대/스탠드"으로 수정 → NewBiz 정확 선택\n✅ 테스트 티켓: #1000151691 (ASIN B0DJPZ8RHG)',
      'v2.19.1'],
    [44,
      '【Before】 html2canvas(1.4MB) @require가 잔류(Liquid Glass 제거 후 dead code) / 매 Auto-Fill마다 ZD field-opts API 3회 + ticket JSON 2회 중복 호출 / SC 세션 만료 시 무음 실패 / SC 구매이력 업데이트 시 전체 패널 re-render / GAS 시트 읽기 매 요청마다 cold read\n【After v2.20.0】 html2canvas @require 제거, fetchTicketJson_(60s TTL) + fetchZdFieldOptsCached(sessionStorage 30min TTL) 캐싱 도입, SC 세션 만료 배너 표시, SC stats 인플레이스 DOM 패치(#sp-stat-badge/#sp-stat-link-wrap), loadSheetRows_(ScriptCache 1h) + getMktIndex_(4h) + loadDefectDataDR_ allRows(6h) 캐싱\n✅ 2026-06-30 적용',
      'v2.20.0'],
    [45,
      '【Before】 환불 수 = OrderStatus === "Canceled" 주문 수 → 배송 완료 후 환불된 주문은 OrderStatus가 Shipped 유지, 반영 안 됨 → 환불 0건 표시\n【After GAS v2.6.2 / TM v2.20.1】 Finances API GET /finances/v0/orders/{id}/financialEvents 병렬 호출(UrlFetchApp.fetchAll, 30건 배치) → RefundEventList 비어있지 않으면 환불로 집계. 오더당 결과 1hr 캐시, 403 시 graceful 폴백. SC fallback totalRefunds: 0 하드코딩 → null로 수정(기존 SP-API 값 유지)\n✅ 2026-07-06 적용',
      'v2.20.1'],
  ];

  UPDATES.forEach(([row, feedback, build]) => {
    sheet.getRange(row, F).setValue(feedback);
    if (build) sheet.getRange(row, J).setValue(build);
  });

  sheet.getRange(35, 2).setValue('Product name'); // B35 was "Purchase Date" — incorrect category

  // GCX Test (H=8) results for rows 32-39
  const H_COL = 8;
  const GCX_TEST = [
    [32, 'Pass'],   // v2.7.7 _gcrFilledThisTicket flag covers field-change issue
    [33, 'Pass'],   // v2.8.3 Spigen SKU pattern rejects "PE2213IN 35w"
    [34, 'Pass'],   // v2.8.3 Spigen SKU pattern rejects "PE2212IN 65w"
    [35, 'Pass'],   // v2.8.2 AGL07930 → "Glas.tR EZ Fit Anti Reflection"
    [36, 'Fail'],   // 📋 ongoing — SC session dependent, 추가 조치 예정
    // row 37: Fail already set by team — do not overwrite
    [38, 'Pass'],   // v2.8.3 Spigen SKU pattern rejects "PE2304IN 45w"
    [39, 'Pass'],   // GAS 2026-06-15 India LWA per-region catch + retry applied
    [40, 'Pass'],   // v2.19.1 clearAllZdFields_() moved to Auto-Fill callback
    [41, 'Pass'],   // v2.19.1 AGL10123 → "Glas.tR Optik Pro"
    [42, 'Pass'],   // v2.19.1 ACS07624 → "Ultra Hybrid MagFit"
    [43, 'Pass'],   // v2.19.1 AMP08885 대분류 → "차량용 거치대/스탠드"
  ];
  GCX_TEST.forEach(([row, val]) => sheet.getRange(row, H_COL).setValue(val));

  Logger.log('Done — updated ' + UPDATES.length + ' rows in GCX Reply 피드백');
}

function fixProductSheetData() {
  const ss    = SpreadsheetApp.openById('1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo');
  const sheet = ss.getSheetByName('Data');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const SKU_COL    = headers.indexOf('SKU');
  const MODEL_COL  = headers.indexOf('모델명');
  const DAEBUN_COL = headers.indexOf('대분류');

  // SKU → [field, newValue]
  const FIXES = {
    'ACS09959': ['모델명',  'Enzo Aramid T MagFit'],
    'AGL07929': ['모델명',  'Glas.tR EZ Fit Privacy'],
    'AGL07928': ['모델명',  'Glas.tR EZ Fit Slim'],
    'AGL10819': ['모델명',  'Glas.tR EZ Fit Anti-Glare'],
    'AGL07930': ['모델명',  'Glas.tR EZ Fit Anti Reflection'],
    'AMP09837': ['대분류',  '차량용 거치대/스탠드'],
    'AGL10123': ['모델명',  'Glas.tR Optik Pro'],
    'ACS07624': ['모델명',  'Ultra Hybrid MagFit'],
    'AMP08885': ['대분류',  '차량용 거치대/스탠드'],
  };

  const colMap = { '모델명': MODEL_COL, '대분류': DAEBUN_COL };
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const sku = data[i][SKU_COL];
    if (FIXES[sku]) {
      const [field, val] = FIXES[sku];
      const col = colMap[field];
      if (col >= 0) {
        sheet.getRange(i + 1, col + 1).setValue(val);
        Logger.log(`Updated row ${i+1}: ${sku} ${field} → "${val}"`);
        count++;
      }
    }
  }
  Logger.log(`Done — fixed ${count} rows in product sheet`);

  // Clear per-ASIN cache for the corrected SKUs so agents see updated data immediately.
  // Find ASINs that correspond to the fixed SKUs, then remove their cache entries.
  const asinsToInvalidate = [];
  for (let i = 1; i < data.length; i++) {
    const sku = data[i][SKU_COL];
    if (FIXES[sku]) {
      const asinCol = headers.indexOf('ASIN');
      if (asinCol >= 0 && data[i][asinCol]) asinsToInvalidate.push(String(data[i][asinCol]));
    }
  }
  if (asinsToInvalidate.length) {
    const cache = CacheService.getScriptCache();
    asinsToInvalidate.forEach(a => cache.remove('asin_all_' + a));
    cache.remove(PIDX_META_KEY); // live path until the next refreshProductIndex run (≤15 min)
    Logger.log('Cache invalidated for ASINs: ' + asinsToInvalidate.join(', '));
  }
}
