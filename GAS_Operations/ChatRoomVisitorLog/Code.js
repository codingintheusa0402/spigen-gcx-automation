/**
 * ChatRoomVisitorLog — records who visits the GCX Google Chat rooms.
 *
 * Google Chat exposes no "who opened the room" read receipts for other people,
 * so a "visit" is recorded from the two signals the Chat API does expose:
 *
 *   BASELINE — first run for a space: every current member is written once
 *   JOINED   — a person became a member of the space
 *   LEFT     — a person is no longer a member of the space
 *   POSTED   — a person sent a message in the space (optional, on by default)
 *
 * Runs on a time-based trigger (default every 10 min). Each run, per space:
 *   1. lists the current human members and diffs them against the roster kept on
 *      the `Members` sheet → BASELINE / JOINED / LEFT rows
 *   2. lists messages created after the per-space checkpoint → POSTED rows
 *   3. appends everything to `Visitor_Log`, rewrites `Members`, and moves the
 *      checkpoint forward (Script Properties, one per space)
 *
 * Auth: plain user OAuth — the account that authorizes the script must be a
 * member of every space in SPACE_IDS. Uses the Chat advanced service (`Chat`) and,
 * best-effort, the People advanced service (`People`) to turn `users/<id>` into a
 * name + email (results are cached on the `Users` sheet).
 *
 * Script Properties (Project Settings → Script Properties):
 *   SPACE_IDS              required  comma-separated, e.g. "spaces/AAAAabc123,spaces/AAAAdef456"
 *                                    (run listMySpaces() to print the IDs you can see)
 *   LOG_SPREADSHEET_ID     optional  target spreadsheet; created + stored on first run if empty
 *   LOG_MESSAGES           optional  "false" to skip POSTED rows (default "true")
 *   LOG_MESSAGE_TEXT       optional  "false" to log only the message ID, not a text snippet
 *   INITIAL_LOOKBACK_HOURS optional  how far back to read messages on the first run (default 24)
 *   POLL_MINUTES           optional  trigger interval for setupChatRoomVisitorLogTriggers() (1/5/10/15/30, default 10)
 *
 * Entry points:
 *   logChatRoomVisitors()             — the trigger target (also runnable by hand)
 *   listMySpaces()                    — prints space IDs + names the authorizing user belongs to
 *   setupChatRoomVisitorLogTriggers() — (re)creates the recurring trigger
 *   test_dryRunChatRoomVisitors()     — computes everything and logs it, writes nothing
 *   clearMessageCheckpoints()         — forget per-space message checkpoints (re-reads INITIAL_LOOKBACK_HOURS)
 */

var VL_SHEET_LOG     = 'Visitor_Log';
var VL_SHEET_MEMBERS = 'Members';
var VL_SHEET_USERS   = 'Users';

var VL_LOG_HEADER     = ['Timestamp', 'Space', 'Space ID', 'Event', 'User', 'Email', 'User ID', 'Detail'];
var VL_MEMBERS_HEADER = ['Space', 'Space ID', 'User', 'Email', 'User ID', 'Role', 'Joined (Chat)', 'First Seen', 'Last Seen'];
var VL_USERS_HEADER   = ['User ID', 'Name', 'Email', 'Resolved At'];

var VL_DEFAULT_SPREADSHEET_TITLE = 'GCX Chatroom Visitor Log';
var VL_MAX_MESSAGES_PER_RUN      = 500;   // per space per run; the checkpoint carries the rest to the next run
var VL_SNIPPET_LEN               = 80;
var VL_PAGE_SIZE                 = 1000;

/* ======================= entry points ======================= */

function logChatRoomVisitors() {
  runVisitorLog_(false);
}

function test_dryRunChatRoomVisitors() {
  runVisitorLog_(true);
}

/**
 * Prints every space the authorizing account is a member of, so SPACE_IDS can
 * be filled in. Output: <space id> | <type> | <display name>
 */
function listMySpaces() {
  var pageToken = null;
  var count = 0;
  do {
    var res = Chat.Spaces.list({ pageSize: VL_PAGE_SIZE, pageToken: pageToken }) || {};
    (res.spaces || []).forEach(function (s) {
      count++;
      Logger.log('%s | %s | %s', s.name, s.spaceType || s.type || '', s.displayName || '(no name — DM/group)');
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  Logger.log('%s space(s). Put the ones to track into Script Properties → SPACE_IDS (comma-separated).', count);
}

function setupChatRoomVisitorLogTriggers() {
  var fn = 'logChatRoomVisitors';
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
  });
  var minutes = parseInt(vlProp_('POLL_MINUTES', '10'), 10);
  if ([1, 5, 10, 15, 30].indexOf(minutes) === -1) {
    throw new Error('POLL_MINUTES must be one of 1, 5, 10, 15, 30 (got ' + minutes + ')');
  }
  ScriptApp.newTrigger(fn).timeBased().everyMinutes(minutes).create();
  Logger.log('Trigger created: %s every %s min', fn, minutes);
}

function clearMessageCheckpoints() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var n = 0;
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('MSG_CKPT_') === 0) { props.deleteProperty(k); n++; }
  });
  Logger.log('Cleared %s checkpoint(s).', n);
}

/* ======================= main run ======================= */

function runVisitorLog_(dryRun) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('Another run is still in progress — skipping.');
    return;
  }
  try {
    var spaceIds = vlSpaceIds_();
    var ss = vlOpenSpreadsheet_(dryRun);
    var sheets = ss ? vlEnsureSheets_(ss) : null;

    var users = new VlUserDirectory_(sheets ? sheets.users : null);
    var prevMembers = sheets ? vlReadMembers_(sheets.members) : {};   // spaceId -> { userId -> row }
    var now = new Date();

    var logRows = [];
    var nextMembers = {};      // spaceId -> [row]
    var checkpoints = {};      // spaceId -> ISO string

    spaceIds.forEach(function (spaceId) {
      var spaceName = vlSpaceDisplayName_(spaceId);
      var current = vlListHumanMembers_(spaceId);           // userId -> membership
      var prev = prevMembers[spaceId] || {};
      var isBaseline = Object.keys(prev).length === 0;

      // --- membership diff ---
      var rows = [];
      Object.keys(current).forEach(function (userId) {
        var m = current[userId];
        var who = users.resolve(userId, m.member && m.member.displayName);
        var joinedAt = m.createTime ? new Date(m.createTime) : '';
        var p = prev[userId];
        if (!p) {
          logRows.push([isBaseline ? now : (joinedAt || now), spaceName, spaceId,
                        isBaseline ? 'BASELINE' : 'JOINED', who.name, who.email, userId,
                        m.role || '']);
        }
        rows.push([spaceName, spaceId, who.name, who.email, userId, m.role || '',
                   joinedAt, p ? p.firstSeen : now, now]);
      });
      Object.keys(prev).forEach(function (userId) {
        if (current[userId]) return;
        var p = prev[userId];
        logRows.push([now, spaceName, spaceId, 'LEFT', p.name, p.email, userId,
                      p.firstSeen ? 'first seen ' + vlFmt_(p.firstSeen) : '']);
      });
      nextMembers[spaceId] = rows;

      // --- message activity ---
      if (vlProp_('LOG_MESSAGES', 'true').toLowerCase() !== 'false') {
        var r = vlListNewMessages_(spaceId);
        r.messages.forEach(function (msg) {
          var sender = msg.sender || {};
          if (sender.type && sender.type !== 'HUMAN') return;
          var who = users.resolve(sender.name || '', sender.displayName);
          logRows.push([msg.createTime ? new Date(msg.createTime) : now, spaceName, spaceId,
                        'POSTED', who.name, who.email, sender.name || '', vlMessageDetail_(msg)]);
        });
        if (r.checkpoint) checkpoints[spaceId] = r.checkpoint;
      }
    });

    logRows.sort(function (a, b) { return a[0] - b[0]; });

    if (dryRun) {
      Logger.log('DRY RUN — %s log row(s) would be appended:', logRows.length);
      logRows.forEach(function (r) { Logger.log(r.map(function (c) { return c instanceof Date ? vlFmt_(c) : c; }).join(' | ')); });
      Object.keys(checkpoints).forEach(function (s) { Logger.log('checkpoint %s → %s', s, checkpoints[s]); });
      return;
    }

    if (logRows.length) {
      sheets.log.getRange(sheets.log.getLastRow() + 1, 1, logRows.length, VL_LOG_HEADER.length).setValues(logRows);
    }
    vlWriteMembers_(sheets.members, prevMembers, nextMembers);
    users.flush();

    var props = PropertiesService.getScriptProperties();
    Object.keys(checkpoints).forEach(function (s) { props.setProperty('MSG_CKPT_' + s, checkpoints[s]); });

    Logger.log('Done — %s log row(s) appended across %s space(s). Sheet: %s', logRows.length, spaceIds.length, ss.getUrl());
  } finally {
    lock.releaseLock();
  }
}

/* ======================= Chat API ======================= */

function vlSpaceDisplayName_(spaceId) {
  try {
    var s = Chat.Spaces.get(spaceId);
    return s.displayName || spaceId;
  } catch (e) {
    Logger.log('Spaces.get(%s) failed: %s', spaceId, e);
    return spaceId;
  }
}

/** Returns { 'users/123': membership } for joined human members of the space. */
function vlListHumanMembers_(spaceId) {
  var out = {};
  var pageToken = null;
  do {
    var res = Chat.Spaces.Members.list(spaceId, { pageSize: VL_PAGE_SIZE, pageToken: pageToken }) || {};
    (res.memberships || []).forEach(function (m) {
      var u = m.member;
      if (!u || !u.name) return;
      if (u.type && u.type !== 'HUMAN') return;          // skip bots / apps
      if (m.state && m.state !== 'JOINED') return;        // skip INVITED etc.
      out[u.name] = m;
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

/**
 * Messages created after the space's checkpoint (oldest first), capped at
 * VL_MAX_MESSAGES_PER_RUN. Returns { messages, checkpoint } where checkpoint is
 * the createTime of the last message read (or the previous checkpoint).
 */
function vlListNewMessages_(spaceId) {
  var props = PropertiesService.getScriptProperties();
  var since = props.getProperty('MSG_CKPT_' + spaceId);
  if (!since) {
    var hours = parseFloat(vlProp_('INITIAL_LOOKBACK_HOURS', '24')) || 24;
    since = new Date(Date.now() - hours * 3600 * 1000).toISOString();
  }
  var messages = [];
  var last = since;
  var pageToken = null;
  do {
    var res = Chat.Spaces.Messages.list(spaceId, {
      pageSize: Math.min(VL_PAGE_SIZE, VL_MAX_MESSAGES_PER_RUN - messages.length),
      pageToken: pageToken,
      filter: 'createTime > "' + since + '"',
      orderBy: 'createTime ASC'
    }) || {};
    (res.messages || []).forEach(function (m) {
      messages.push(m);
      if (m.createTime && m.createTime > last) last = m.createTime;
    });
    pageToken = res.nextPageToken;
  } while (pageToken && messages.length < VL_MAX_MESSAGES_PER_RUN);
  if (pageToken) Logger.log('%s: more than %s new messages — the rest is picked up next run.', spaceId, VL_MAX_MESSAGES_PER_RUN);
  return { messages: messages, checkpoint: last };
}

function vlMessageDetail_(msg) {
  var id = (msg.name || '').split('/').pop();
  if (vlProp_('LOG_MESSAGE_TEXT', 'true').toLowerCase() === 'false') return id;
  var text = String(msg.text || msg.argumentText || '').replace(/\s+/g, ' ').trim();
  if (!text && msg.attachment && msg.attachment.length) text = '[attachment]';
  if (text.length > VL_SNIPPET_LEN) text = text.slice(0, VL_SNIPPET_LEN) + '…';
  return text ? id + ' · ' + text : id;
}

/* ======================= user directory (name/email) ======================= */

/**
 * users/<id> → { name, email }. Order of preference:
 *   1. cache on the `Users` sheet
 *   2. displayName the Chat API sent along
 *   3. People API (people/<id>, domain profile) — needs directory.readonly; best-effort
 *   4. the raw users/<id>
 */
function VlUserDirectory_(sheet) {
  this.sheet = sheet;
  this.cache = {};
  this.dirty = [];
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, VL_USERS_HEADER.length).getValues().forEach(function (r) {
      if (r[0]) this.cache[String(r[0])] = { name: String(r[1] || ''), email: String(r[2] || '') };
    }, this);
  }
}

VlUserDirectory_.prototype.resolve = function (userId, displayName) {
  if (!userId) return { name: displayName || '', email: '' };
  var hit = this.cache[userId];
  if (hit && hit.email) return hit;                     // fully resolved
  var name = displayName || (hit && hit.name) || '';
  var email = '';
  var people = vlLookupPerson_(userId);
  if (people) { name = name || people.name; email = people.email; }
  if (!name) name = userId;
  var out = { name: name, email: email };
  if (!hit || hit.name !== name || hit.email !== email) {
    this.cache[userId] = out;
    this.dirty.push([userId, name, email, new Date()]);
  }
  return out;
};

VlUserDirectory_.prototype.flush = function () {
  if (!this.sheet || !this.dirty.length) return;
  // rewrite the cache sheet from the map so updated entries don't duplicate
  var rows = Object.keys(this.cache).sort().map(function (id) {
    return [id, this.cache[id].name, this.cache[id].email, new Date()];
  }, this);
  var s = this.sheet;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, VL_USERS_HEADER.length).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, VL_USERS_HEADER.length).setValues(rows);
  this.dirty = [];
};

function vlLookupPerson_(userId) {
  var id = String(userId).replace(/^users\//, '');
  if (!/^\d+$/.test(id)) return null;                   // e.g. users/<email> — nothing to look up
  if (typeof People === 'undefined') return null;
  try {
    var p = People.People.get('people/' + id, {
      personFields: 'names,emailAddresses',
      sources: ['READ_SOURCE_TYPE_PROFILE', 'READ_SOURCE_TYPE_DOMAIN_PROFILE']
    });
    var name = (p.names && p.names[0] && p.names[0].displayName) || '';
    var email = (p.emailAddresses && p.emailAddresses[0] && p.emailAddresses[0].value) || '';
    return (name || email) ? { name: name, email: email } : null;
  } catch (e) {
    Logger.log('People lookup failed for %s: %s', userId, e);
    return null;
  }
}

/* ======================= sheets ======================= */

function vlOpenSpreadsheet_(dryRun) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('LOG_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (dryRun) { Logger.log('No LOG_SPREADSHEET_ID yet — dry run continues without a sheet.'); return null; }
  var ss = SpreadsheetApp.create(VL_DEFAULT_SPREADSHEET_TITLE);
  props.setProperty('LOG_SPREADSHEET_ID', ss.getId());
  Logger.log('Created log spreadsheet: %s', ss.getUrl());
  return ss;
}

function vlEnsureSheets_(ss) {
  var log = vlEnsureSheet_(ss, VL_SHEET_LOG, VL_LOG_HEADER);
  var members = vlEnsureSheet_(ss, VL_SHEET_MEMBERS, VL_MEMBERS_HEADER);
  var users = vlEnsureSheet_(ss, VL_SHEET_USERS, VL_USERS_HEADER);
  var stray = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (stray && ss.getSheets().length > 3 && stray.getLastRow() === 0) ss.deleteSheet(stray);
  return { log: log, members: members, users: users };
}

function vlEnsureSheet_(ss, name, header) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

/** Members sheet → { spaceId: { userId: {name, email, role, joinedAt, firstSeen, lastSeen, raw} } } */
function vlReadMembers_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, VL_MEMBERS_HEADER.length).getValues().forEach(function (r) {
    var spaceId = String(r[1] || ''), userId = String(r[4] || '');
    if (!spaceId || !userId) return;
    if (!out[spaceId]) out[spaceId] = {};
    out[spaceId][userId] = {
      name: String(r[2] || ''), email: String(r[3] || ''), role: String(r[5] || ''),
      joinedAt: r[6], firstSeen: r[7] instanceof Date ? r[7] : (r[7] ? new Date(r[7]) : ''), lastSeen: r[8], raw: r
    };
  });
  return out;
}

/** Rewrites the Members sheet: rows for processed spaces are replaced, other spaces' rows are kept. */
function vlWriteMembers_(sheet, prevMembers, nextMembers) {
  var rows = [];
  Object.keys(prevMembers).forEach(function (spaceId) {
    if (nextMembers.hasOwnProperty(spaceId)) return;
    Object.keys(prevMembers[spaceId]).forEach(function (u) { rows.push(prevMembers[spaceId][u].raw); });
  });
  Object.keys(nextMembers).forEach(function (spaceId) {
    nextMembers[spaceId].forEach(function (r) { rows.push(r); });
  });
  rows.sort(function (a, b) { return (a[0] + a[2]).localeCompare(b[0] + b[2]); });
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, VL_MEMBERS_HEADER.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, VL_MEMBERS_HEADER.length).setValues(rows);
}

/* ======================= helpers ======================= */

function vlProp_(key, dflt) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === undefined || v === '') ? dflt : String(v).trim();
}

function vlSpaceIds_() {
  var ids = vlProp_('SPACE_IDS', '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!ids.length) throw new Error('Script Property SPACE_IDS is empty. Run listMySpaces() and set it to e.g. "spaces/AAAAabc123,spaces/AAAAdef456".');
  ids.forEach(function (id) {
    if (!/^spaces\/[A-Za-z0-9_-]+$/.test(id)) throw new Error('Bad space id "' + id + '" — expected "spaces/<id>".');
  });
  return ids;
}

function vlFmt_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
