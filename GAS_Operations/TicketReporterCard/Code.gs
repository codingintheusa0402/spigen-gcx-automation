/**
 * Ticket Reporter — interactive Google Chat app.
 *
 * Flow: the ticket-reporter monitor (Chrome-automation session) writes a row into the
 * TicketQueue sheet for a new Pending + 4.Product Issue ticket, then sends a short
 * trigger message ("티켓 1000161577" or similar) into the target space using the
 * signed-in human's own Chrome session (NOT this app's identity — this app never needs
 * to post proactively, so it never needs the chat.bot scope / domain-wide delegation
 * that blocked ../BadReview_ChatReport/chat_app's sibling project).
 *
 * onMessage reads the ticket id out of that trigger text, looks up the queued row (or
 * falls back to an empty context if not queued), and replies with the interactive card:
 * one reference dropdown (all 43 canned phrases) + one freeform note textbox the user
 * actually types the internal note into. Picking a dropdown item just appends that phrase
 * as a new line onto whatever is already in the note box (see onRefChange) — the dropdown
 * never "is" the note, it's just a fast way to insert a line into it. Submitting posts the
 * note box's current text to Zendesk with comment.public = false — NEVER true. That is a
 * hard rule; see postInternalNote_.
 */

/* ===================== Chat envelope helpers (add-on mode) ===================== */

function chatCreate_(message) {
  return { hostAppDataAction: { chatDataAction: { createMessageAction: { message: message } } } };
}
function chatUpdate_(message) {
  return { hostAppDataAction: { chatDataAction: { updateMessageAction: { message: message } } } };
}

/* ===================== Chat event handlers ===================== */

function onAddedToSpace(event) {
  return chatCreate_({
    text: 'Ticket Reporter 앱입니다. 모니터가 새 티켓을 큐에 넣으면 "티켓 <번호>"라고 보내 카드를 불러오세요.'
  });
}

function onRemovedFromSpace(event) {}

function onMessage(event) {
  var text = messageText_(event);
  var trimmed = String(text || '').trim();

  // `/revision <feedback>` — a note about how future ticket reports should be written
  // (사용자 지시 2026-09-16). Never touches Zendesk; just logged to the Feedback sheet tab
  // for the ticket-reporter session to pick up and apply as a permanent rule change. Checked
  // before the ticket-number regex below since pasted feedback often quotes a full report
  // (which contains ticket numbers of its own).
  if (/^\/revision\b/i.test(trimmed)) {
    return handleRevisionFeedback_(trimmed.replace(/^\/revision\s*/i, ''), event);
  }

  // `/manual` — dumps the full ticket-reporter SKILL.md text into the chat (사용자 지시
  // 2026-09-17). The file itself lives only on the local machine running the monitor
  // session, so it's mirrored into the Manual sheet tab (see MANUAL_TAB) every time
  // SKILL.md changes; this just reads that cell back out. Checked before the ticket-number
  // regex for the same reason /revision is.
  if (/^\/manual\b/i.test(trimmed)) {
    return handleManualRequest_(event);
  }

  // `/btw <question>` — general-purpose Q&A escape hatch, NOT the ticket-report flow.
  // Anything after `/btw` is answered directly (e.g. "/btw 이번 주 PI 티켓 경향 어때?") instead
  // of being parsed as a ticket trigger/report. Checked before the ticket-number regex for the
  // same reason /revision and /manual are — a free-form question very often contains a digit
  // sequence of its own (order numbers, dates, ticket numbers quoted as examples) that must not
  // get misread as "open ticket card for this id". See handleBtwQuestion_.
  if (/^\/btw\b/i.test(trimmed)) {
    return handleBtwQuestion_(trimmed.replace(/^\/btw\s*/i, ''), event);
  }

  // Thread-reply direct-post path: a reply inside a thread the monitor already mapped to a
  // ticket (see lookupTicketByThread_) is treated as the internal-note content itself — no
  // card round-trip needed. This is how a human replies under an auto-sent static report and
  // has it land on Zendesk directly. Checked BEFORE the ticket-number regex below, same
  // reasoning as /revision and /manual above — and for good reason: a reply inside a mapped
  // thread often legitimately cites another ticket's number (e.g. "#1000164335 건과 동일
  // 이슈" as a precedent reference), which must NOT redirect the note onto that cited ticket
  // instead of the thread's own. (Bug fixed 2026-10-06 — originally this thread-mapping
  // check only ran when the reply text contained NO digit sequence at all, so a reply to the
  // #1000164768 thread that merely mentioned precedent ticket #1000164335 got misrouted and
  // posted as an internal note onto #1000164335 instead of #1000164768. Caught via /revision
  // feedback; see auto-memory for the full diagnosis.)
  var threadName = threadName_(event);
  var mappedId = threadName ? lookupTicketByThread_(threadName) : null;
  if (mappedId) {
    // TCT-log-sourced tickets (Lazada/Shopee "Esc T2" rows, 사용자 지시 2026-09-18) have a
    // non-numeric id like "260915BQH9J0" — a plain Zendesk ticket id is always all-digit.
    // These write back into the TCT log sheet instead of calling the Zendesk API.
    if (!/^\d+$/.test(mappedId)) {
      return updateTctLogRow_(mappedId, trimmed, event);
    }
    return postThreadReplyAsNote_(mappedId, trimmed, event);
  }

  var m = trimmed.match(/(\d{6,})/);
  if (!m) {
    // No mapping found and no ticket number in the text — log the thread so a backfill can
    // register it later (e.g. a report sent before send.py's --ticket-id auto-derive fix,
    // 사용자 지시 2026-09-16). Without this, an old thread's mapping can never be recovered
    // since the app has no Chat API read access to work out which ticket the thread belongs
    // to on its own.
    logUnmappedThread_(threadName, event);
    return chatCreate_({ text: '티켓 번호를 찾을 수 없습니다. 예: "티켓 1000161577"' });
  }
  // Short trigger ("티켓 1000161577") just opens the card. A longer pasted message
  // (a full TCK 전달 전 보고 리포트) is rendered above the same dropdowns/button
  // so confirmer can read the report and act on it in one card.
  var isShortTrigger = /^티켓\s*\d{6,}$/.test(trimmed);
  return chatCreate_({ cardsV2: [buildCard_(m[1], isShortTrigger ? null : trimmed, null)] });
}

function messageText_(event) {
  var msg = (event && event.message) ||
            (event && event.chat && event.chat.messagePayload && event.chat.messagePayload.message) || {};
  return msg.argumentText || msg.text || '';
}

function threadName_(event) {
  var msg = (event && event.message) ||
            (event && event.chat && event.chat.messagePayload && event.chat.messagePayload.message) || {};
  return (msg.thread && msg.thread.name) || '';
}

/** Classic-mode shim; add-on mode calls the named function in onClick.action.function directly. */
function onCardClick(event) {
  var fn = (event.common && event.common.invokedFunction) ||
           (event.action && event.action.actionMethodName) || '';
  if (fn === 'submitNote') return submitNote(event);
  if (fn === 'onRefChange') return onRefChange(event);
  return chatCreate_({ text: '알 수 없는 동작입니다.' });
}

/**
 * Fired by the reference dropdown's onChangeAction. Reads whatever is currently typed in
 * the note textbox plus the newly picked phrase, appends the phrase as a new line, and
 * rebuilds the card with that as the note box's new value. The dropdown itself is always
 * rebuilt with nothing selected (see buildCard_) so picking the SAME phrase again still
 * counts as a change and appends it again.
 */
function onRefChange(event) {
  var inputs = formInputs_(event);
  var params = actionParams_(event);
  var ticketId = params.ticketId;
  var reportText = params.reportText || null;

  var picked = {};
  var chosen = extractString_(inputs, 'ref');
  var currentNote = extractString_(inputs, 'noteText');
  picked.noteText = chosen ? (currentNote ? currentNote + '\n' + chosen : chosen) : currentNote;
  picked.confirmer = extractString_(inputs, 'confirmer');

  return chatUpdate_({ cardsV2: [buildCard_(ticketId, reportText, picked)] });
}

/* ===================== Card building ===================== */

/**
 * @param {string} ticketId
 * @param {?string} reportText  full pasted TCK report body, shown above the card body when present
 * @param {?Object} picked      { noteText, confirmer } to preserve across an onRefChange
 *                              rebuild; null on the first render (empty note, KJW default).
 */
function buildCard_(ticketId, reportText, picked) {
  picked = picked || {};
  var row = lookupQueueRow_(ticketId);
  var widgets = [];

  if (reportText) {
    widgets.push({ textParagraph: { text: reportHtml_(reportText) } });
  }

  if (row) {
    widgets.push({ decoratedText: {
      topLabel: 'Zendesk 티켓', text: '<b>#' + ticketId + '</b> — ' + escapeHtml_(row.subject || ''),
      bottomLabel: row.country ? (row.country + ' · ' + (row.category || '')) : '',
      onClick: { openLink: { url: zendeskTicketUrl_(ticketId) } }
    }});
  } else {
    widgets.push({ decoratedText: {
      topLabel: 'Zendesk 티켓', text: '<b>#' + ticketId + '</b> (큐에 없음 — 수동 입력)',
      onClick: { openLink: { url: zendeskTicketUrl_(ticketId) } }
    }});
  }

  var changeParams = [
    { key: 'ticketId', value: String(ticketId) },
    { key: 'reportText', value: reportText || '' }
  ];

  // Reference dropdown never keeps a selected item — every pick is rendered as a fresh
  // "change" so picking the same phrase twice in a row still fires onRefChange twice.
  var refItems = [{ text: '(문구 선택 시 아래 노트에 자동 추가)', value: '', selected: true }]
    .concat(REFERENCE_PHRASES.map(function (text) { return { text: text, value: text, selected: false }; }));
  widgets.push({ selectionInput: {
    name: 'ref', label: '①', type: 'DROPDOWN', items: refItems,
    onChangeAction: { function: 'onRefChange', parameters: changeParams }
  }});

  widgets.push({ textInput: {
    name: 'noteText', label: '내부 노트 내용 (직접 입력 가능 · 드롭다운 선택 시 자동 추가)',
    type: 'MULTIPLE_LINE', value: picked.noteText || ''
  } });

  var confirmerCurrent = picked.confirmer || CONFIRMERS[0];
  var confirmerItems = CONFIRMERS.map(function (c) {
    return { text: c, value: c, selected: c === confirmerCurrent };
  });
  widgets.push({ selectionInput: { name: 'confirmer', label: '[GCX ___ 컨펌]', type: 'DROPDOWN', items: confirmerItems } });

  widgets.push({ buttonList: { buttons: [
    { text: 'Zendesk 내부 노트로 전송', type: 'FILLED', onClick: { action: {
        function: 'submitNote',
        parameters: [
          { key: 'ticketId', value: String(ticketId) },
          { key: 'reportText', value: reportText || '' }
        ]
    } } }
  ]}});

  return {
    cardId: 'ticket-reporter-' + ticketId,
    card: {
      header: { title: '티켓 #' + ticketId + ' — 처리 요청 사항', subtitle: '노트 작성 후 전송 · 항상 내부 노트로만 등록됩니다' },
      sections: [{ widgets: widgets }]
    }
  };
}

/* ===================== Submit → compose note → post to Zendesk (internal only) ===================== */

function submitNote(event) {
  var inputs = formInputs_(event);
  var params = actionParams_(event);
  var ticketId = params.ticketId;
  var reportText = params.reportText || null;
  if (!ticketId) return chatUpdate_({ text: '티켓 번호를 찾을 수 없습니다.' });

  var noteText = extractString_(inputs, 'noteText').trim();
  var confirmer = extractString_(inputs, 'confirmer') || CONFIRMERS[0];

  if (!noteText) {
    return chatUpdate_({
      cardsV2: [buildCard_(ticketId, reportText, { noteText: noteText, confirmer: confirmer })],
      text: '⚠️ 내부 노트 내용을 입력하거나 드롭다운에서 문구를 선택해 주세요.'
    });
  }

  var body = composeNote_(noteText, confirmer);
  var result = postInternalNote_(ticketId, body);

  if (!result.ok) {
    return chatUpdate_({ text: '❌ #' + ticketId + ' 내부 노트 전송 실패 (' + result.status + '): ' + result.error });
  }

  logSubmission_(ticketId, body, confirmer);

  return chatUpdate_({
    cardsV2: [{
      cardId: 'ticket-reporter-done-' + ticketId,
      card: {
        header: { title: '✅ #' + ticketId + ' 내부 노트 전송 완료', subtitle: '[GCX ' + confirmer + ' 컨펌]' },
        sections: [{ widgets: [
          { textParagraph: { text: escapeHtml_(body).replace(/\n/g, '<br>') } },
          { buttonList: { buttons: [{ text: '티켓 열기', onClick: { openLink: { url: zendeskTicketUrl_(ticketId) } } }] } }
        ]}]
      }
    }]
  });
}

/**
 * Thread-reply direct-post path (no interactive card): a human replied in the thread under
 * an auto-sent static report, mentioning this app, with the actual note text and no ticket
 * number (see onMessage). Post it straight to Zendesk as the internal note AND reopen the
 * ticket to Open (사용자 지시 2026-09-16 — this path represents a confirmer's finished
 * decision, so the ticket should move out of Pending/On-hold back into the active queue).
 * Confirmer code is resolved from the sender's email (confirmerForUser_) — the message
 * text is the note only, not a form, so there's no confirmer dropdown to read from here.
 */
function postThreadReplyAsNote_(ticketId, noteText, event) {
  if (!noteText) {
    return chatCreate_({ text: '⚠️ 노트 내용이 비어 있습니다.' });
  }
  var confirmer = confirmerForUser_(event);
  var body = composeNote_(noteText, confirmer);
  var result = postInternalNoteAndReopen_(ticketId, body);

  if (!result.ok) {
    return chatCreate_({ text: '❌ #' + ticketId + ' 내부 노트 전송 실패 (' + result.status + '): ' + result.error });
  }

  logSubmission_(ticketId, body, confirmer);

  return chatCreate_({
    cardsV2: [{
      cardId: 'ticket-reporter-thread-done-' + ticketId + '-' + new Date().getTime(),
      card: {
        header: { title: '내부 노트 전송 완료', subtitle: '[GCX ' + confirmer + ' 컨펌]' },
        sections: [{ widgets: [
          { textParagraph: { text: escapeHtml_(body).replace(/\n/g, '<br>') } },
          { buttonList: { buttons: [{ text: '티켓 열기', onClick: { openLink: { url: zendeskTicketUrl_(ticketId) } } }] } }
        ]}]
      }
    }]
  });
}

/** "처리 요청 사항" \n\n <노트 내용> \n\n "[GCX <confirmer> 컨펌]" — exact user-specified format. */
function composeNote_(noteText, confirmer) {
  return [
    '처리 요청 사항',
    '',
    noteText,
    '',
    '[GCX ' + confirmer + ' 컨펌]'
  ].join('\n');
}

/**
 * Posts `body` as a Zendesk INTERNAL note (comment.public = false). HARD RULE: never
 * flip this to true — internal notes only, this is never sent to the customer directly.
 */
function postInternalNote_(ticketId, body) {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('ZENDESK_EMAIL');
  var token = props.getProperty('ZENDESK_API_TOKEN');
  if (!email || !token) return { ok: false, status: 0, error: 'ZENDESK_EMAIL/ZENDESK_API_TOKEN not set in Script Properties' };

  var url = 'https://' + ZENDESK_SUBDOMAIN + '.zendesk.com/api/v2/tickets/' + encodeURIComponent(ticketId) + '.json';
  var auth = Utilities.base64Encode(email + '/token:' + token);
  var payload = { ticket: { comment: { body: body, public: false } } };

  var resp = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + auth },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code >= 200 && code < 300) return { ok: true, status: code };
  return { ok: false, status: code, error: resp.getContentText().slice(0, 300) };
}

function zendeskTicketUrl_(ticketId) {
  return 'https://' + ZENDESK_SUBDOMAIN + '.zendesk.com/agent/tickets/' + ticketId;
}

/**
 * Same as postInternalNote_ but also reopens the ticket to `open` in the same PUT call —
 * used only by the thread-reply direct-post path (postThreadReplyAsNote_), where the
 * confirmer's reply is a finished decision that should move the ticket out of
 * Pending/On-hold. postInternalNote_ (interactive-card submit) never changes status.
 */
function postInternalNoteAndReopen_(ticketId, body) {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('ZENDESK_EMAIL');
  var token = props.getProperty('ZENDESK_API_TOKEN');
  if (!email || !token) return { ok: false, status: 0, error: 'ZENDESK_EMAIL/ZENDESK_API_TOKEN not set in Script Properties' };

  var url = 'https://' + ZENDESK_SUBDOMAIN + '.zendesk.com/api/v2/tickets/' + encodeURIComponent(ticketId) + '.json';
  var auth = Utilities.base64Encode(email + '/token:' + token);
  var payload = { ticket: { comment: { body: body, public: false }, status: 'open' } };

  var resp = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + auth },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code >= 200 && code < 300) return { ok: true, status: code };
  return { ok: false, status: code, error: resp.getContentText().slice(0, 300) };
}

/**
 * Resolves the confirmer code for the thread-reply direct-post path from the Chat event's
 * sender email (CONFIRMER_BY_EMAIL in Config.gs), falling back to CONFIRMERS[0] (KJW) for
 * anyone not explicitly mapped.
 */
function confirmerForUser_(event) {
  var email = (event && event.user && event.user.email) || '';
  return CONFIRMER_BY_EMAIL[email.toLowerCase()] || CONFIRMERS[0];
}

/* ===================== Lazada/Shopee TCT log sheet (thread-reply hand-off) ===================== */

/**
 * Thread-reply counterpart to postThreadReplyAsNote_, for tickets sourced from the
 * Lazada/Shopee TCT log sheet instead of Zendesk (사용자 지시 2026-09-18). There is no
 * Zendesk ticket to post an internal note to here — instead the reply is written straight
 * into that row's Voucher/Advice/GCX STATUS/Status columns, which is what sends the row
 * back down to TCT (Tier 1).
 *
 * Unlike the Zendesk path, this does NOT wrap the memo in "처리 요청 사항 / [GCX 컨펌]" —
 * the raw memo text (minus a leading "/voucher N%" token, if present) goes straight into
 * the Advice/Internal Memo column, verbatim, per explicit user instruction.
 *
 * "/voucher <100|50|10>[%] <나머지 메모>" → Voucher column = "Provide <N>% voucher" (must
 * match the sheet's strict dropdown string exactly) + memo column = the remaining text.
 * No "/voucher" token → Voucher column left untouched, memo column = the full text as-is.
 * Either way, GCX STATUS✅ → "Advice given" and Status → "Esc T1  " (sic — the dropdown's
 * own value has two trailing spaces; TCT_LOG_STATUS_ESC_T1 preserves that exactly).
 */
function updateTctLogRow_(ticketId, noteText, event) {
  if (!noteText) {
    return chatCreate_({ text: '⚠️ 메모 내용이 비어 있습니다.' });
  }
  var loc = findTctLogRow_(ticketId);
  if (!loc) {
    return chatCreate_({ text: '❌ TCT 로그 시트에서 티켓 ' + ticketId + '을(를) 찾을 수 없습니다.' });
  }

  var voucherMatch = noteText.match(/\/voucher\s*(\d+)\s*%?\s*/i);
  var voucherLabel = null;
  var memo = noteText;
  if (voucherMatch) {
    var pct = voucherMatch[1];
    if (pct === '100' || pct === '50' || pct === '10') {
      voucherLabel = 'Provide ' + pct + '% voucher';
    }
    memo = (noteText.slice(0, voucherMatch.index) + noteText.slice(voucherMatch.index + voucherMatch[0].length)).trim();
  }

  var sheet = SpreadsheetApp.openById(TCT_LOG_SHEET_ID).getSheetByName(loc.tab);
  if (voucherLabel) sheet.getRange(loc.row, TCT_LOG_COL.VOUCHER).setValue(voucherLabel);
  sheet.getRange(loc.row, TCT_LOG_COL.MEMO).setValue(memo);
  sheet.getRange(loc.row, TCT_LOG_COL.GCX_STATUS).setValue(TCT_LOG_GCX_STATUS_ADVICE_GIVEN);
  sheet.getRange(loc.row, TCT_LOG_COL.STATUS).setValue(TCT_LOG_STATUS_ESC_T1);

  var confirmer = confirmerForUser_(event);
  logSubmission_(ticketId, memo, confirmer);

  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + TCT_LOG_SHEET_ID + '/edit#gid=' +
      (loc.tab === 'Shopee log' ? '1376766342' : '43582188') + '&range=A' + loc.row;
  return chatCreate_({
    cardsV2: [{
      cardId: 'ticket-reporter-tct-done-' + ticketId + '-' + new Date().getTime(),
      card: {
        header: { title: 'TCT 로그 기록 완료', subtitle: loc.tab + ' · ' + (voucherLabel || 'Voucher 미지정') },
        sections: [{ widgets: [
          { textParagraph: { text: escapeHtml_(memo).replace(/\n/g, '<br>') } },
          { buttonList: { buttons: [{ text: '시트에서 보기', onClick: { openLink: { url: sheetUrl } } }] } }
        ]}]
      }
    }]
  });
}

/** Searches both TCT_LOG_TABS for a row whose Ticket ID column matches. */
function findTctLogRow_(ticketId) {
  var ss = SpreadsheetApp.openById(TCT_LOG_SHEET_ID);
  for (var t = 0; t < TCT_LOG_TABS.length; t++) {
    var sheet = ss.getSheetByName(TCT_LOG_TABS[t]);
    if (!sheet) continue;
    var vals = sheet.getRange(1, TCT_LOG_COL.TICKET_ID, sheet.getLastRow(), 1).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (String(vals[r][0]).trim() === String(ticketId).trim()) {
        return { tab: TCT_LOG_TABS[t], row: r + 1 };
      }
    }
  }
  return null;
}

/* ===================== TicketQueue sheet (monitor → card hand-off) ===================== */

/** Plain name so it shows in the editor's Run-function dropdown (GAS hides `_`-suffixed names). */
function setupOnce() { setupOnce_(); }

/** Run once from the Apps Script editor to provision the queue sheet. */
function setupOnce_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(QUEUE_SHEET_PROP)) {
    Logger.log('Queue sheet already set: ' + props.getProperty(QUEUE_SHEET_PROP));
    return;
  }
  var ss = SpreadsheetApp.create('Ticket Reporter — Interactive Card Queue');
  var sheet = ss.getSheets()[0];
  sheet.setName(QUEUE_TAB);
  sheet.appendRow(['ticketId', 'subject', 'country', 'category', 'queuedAt', 'submittedAt', 'confirmer', 'threadId']);
  props.setProperty(QUEUE_SHEET_PROP, ss.getId());
  Logger.log('Created queue sheet: ' + ss.getUrl());
}

function lookupQueueRow_(ticketId) {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return null;
  var vals = SpreadsheetApp.openById(id).getSheetByName(QUEUE_TAB).getDataRange().getValues();
  var H = vals[0];
  var iId = H.indexOf('ticketId'), iSub = H.indexOf('subject'), iCty = H.indexOf('country'), iCat = H.indexOf('category');
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][iId]) === String(ticketId)) {
      return { subject: vals[r][iSub], country: vals[r][iCty], category: vals[r][iCat] };
    }
  }
  return null;
}

/**
 * Reverse lookup: which ticket does this Chat thread belong to? Populated by send.py right
 * after it posts a static report webhook message — it writes {ticketId, threadId} into the
 * same queue sheet. Lets onMessage resolve a plain thread reply (no ticket number in the
 * text) back to a ticket without needing Chat API read access (chat.bot is not
 * user-consentable in Apps Script's OAuth flow, so the app can't call spaces.messages.list
 * to read the thread's own history — this sheet-based mapping avoids needing that).
 */
function lookupTicketByThread_(threadName) {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id || !threadName) return null;
  var vals = SpreadsheetApp.openById(id).getSheetByName(QUEUE_TAB).getDataRange().getValues();
  var H = vals[0];
  var iId = H.indexOf('ticketId'), iThread = H.indexOf('threadId');
  if (iThread === -1) return null;
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][iThread]) === String(threadName)) return String(vals[r][iId]);
  }
  return null;
}

/**
 * `/revision <feedback>` handler (사용자 지시 2026-09-16). This is a feedback channel, not a
 * Zendesk action: it never touches a ticket. The feedback text (often a full pasted report
 * plus a note about what should have been different) is appended to the Feedback sheet tab
 * with `appliedAt` left blank. The ticket-reporter Claude session checks that tab at the
 * start of every monitor tick, reads any unapplied rows, updates SKILL.md's writing rules
 * accordingly, then stamps `appliedAt` so the same feedback isn't re-applied.
 */
function handleRevisionFeedback_(feedbackText, event) {
  feedbackText = String(feedbackText || '').trim();
  if (!feedbackText) {
    return chatCreate_({ text: '⚠️ /revision 뒤에 피드백 내용을 입력해 주세요.' });
  }
  var email = (event && event.user && event.user.email) || '';
  logRevisionFeedback_(feedbackText, email);
  return chatCreate_({
    cardsV2: [{
      cardId: 'ticket-reporter-revision-' + new Date().getTime(),
      card: {
        header: { title: '피드백 접수 완료' },
        sections: [{ widgets: [
          { textParagraph: { text: '다음 리포트부터 반영됩니다.<br><br>' + escapeHtml_(feedbackText).replace(/\n/g, '<br>') } }
        ]}]
      }
    }]
  });
}

/**
 * Logs a thread that mentioned the app with no ticket number and no TicketQueue mapping
 * (사용자 지시 2026-09-16) — lets a human backfill TicketQueue's threadId for that ticket
 * later, since the app has no other way to discover which ticket an old thread belongs to.
 */
function logUnmappedThread_(threadName, event) {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return;
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(UNMAPPED_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(UNMAPPED_TAB);
    sheet.appendRow(['ts', 'threadId', 'senderEmail', 'text', 'resolvedTicketId']);
  }
  var email = (event && event.user && event.user.email) || '';
  var text = messageText_(event);
  sheet.appendRow([new Date(), threadName || '', email, text, '']);
}

/**
 * `/manual` handler (사용자 지시 2026-09-17, 2026-09-18 링크 방식으로 변경).
 *
 * 처음엔 SKILL.md 전체를 textParagraph 여러 개로 쪼개 카드에 직접 채워 넣었으나, SKILL.md가
 * /revision 누적으로 계속 커지면서(24KB+) 카드 전체 payload가 Chat 쪽 크기 제한에 걸린 것으로
 * 보인다 — Apps Script 실행 로그는 "Completed"(에러 없음)로 남는데 정작 Chat에는 메시지가
 * 전혀 표시되지 않는 증상 확인(2026-09-18). Apps Script 쪽에서는 이 실패가 보이지 않아
 * 디버그가 어려우므로, 이제는 짧은 미리보기 텍스트 + Manual 시트 탭으로 바로 이동하는 링크
 * 버튼을 반환한다 — 페이로드 크기 문제 자체를 원천적으로 피한다.
 */
function handleManualRequest_(event) {
  var content = getManualText_();
  if (!content) {
    return chatCreate_({ text: '⚠️ SKILL.md 매뉴얼 텍스트가 아직 Manual 시트에 등록되지 않았습니다.' });
  }
  var preview = content.length > 1500 ? content.slice(0, 1500) + '\n…' : content;
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + id + '/edit#gid=' + MANUAL_TAB_GID;
  return chatCreate_({
    cardsV2: [{
      cardId: 'ticket-reporter-manual-' + new Date().getTime(),
      card: {
        header: { title: 'Ticket Reporter — SKILL.md', subtitle: '전체 ' + content.length + '자 (미리보기 1,500자)' },
        sections: [{ widgets: [
          { textParagraph: { text: escapeHtml_(preview) } },
          { buttonList: { buttons: [{ text: '전체 매뉴얼 시트에서 보기', onClick: { openLink: { url: sheetUrl } } }] } }
        ]}]
      }
    }]
  });
}

/**
 * `/btw <question>` handler — free-form Q&A, deliberately separate from the ticket-report
 * flow above. Answers via the Claude API (ANTHROPIC_API_KEY Script Property; never
 * hardcoded — see feedback_no_hardcoded_secrets_in_docs memory).
 *
 * This app intentionally has no chat.bot scope / domain-wide delegation (see file header),
 * so it cannot push more than one reply per invocation through the normal card-action
 * response channel, and that one reply can't be "updated" mid-flight the way a card can be
 * rebuilt in response to a click. To still satisfy "if the answer takes a long time,
 * periodically say which step it's on", progress pings are posted PROACTIVELY through this
 * same room's own incoming webhook (PROGRESS_WEBHOOK_URL Script Property — the exact
 * mechanism send.py already uses to post ticket reports) as each stage of the pipeline
 * completes, while the actual synchronous action response at the end carries the final
 * answer. If PROGRESS_WEBHOOK_URL isn't set, postProgress_ no-ops silently — /btw still
 * answers, just without the interim pings.
 */
function handleBtwQuestion_(question, event) {
  question = String(question || '').trim();
  if (!question) {
    return chatCreate_({ text: '⚠️ /btw 뒤에 질문을 입력해 주세요. 예: "/btw 이번 주 접수된 PI 티켓 경향이 어때?"' });
  }

  var startedAt = new Date().getTime();
  postProgress_('🤔 질문 확인 중… (' + truncate_(question, 120) + ')');
  postProgress_('📡 답변 생성 중…');

  var answer;
  try {
    answer = askClaude_(question);
  } catch (err) {
    postProgress_('❌ 답변 생성 실패: ' + err.message);
    return chatCreate_({ text: '❌ /btw 답변 생성 실패: ' + err.message });
  }

  var elapsedSec = Math.round((new Date().getTime() - startedAt) / 1000);
  if (elapsedSec >= 10) {
    postProgress_('✅ 답변 생성 완료 (' + elapsedSec + '초 소요)');
  }

  return chatCreate_({
    cardsV2: [{
      cardId: 'ticket-reporter-btw-' + new Date().getTime(),
      card: {
        header: { title: '/btw 답변', subtitle: elapsedSec + '초 소요' },
        sections: [{ widgets: [
          { textParagraph: { text: escapeHtml_(answer).replace(/\n/g, '<br>') } }
        ]}]
      }
    }]
  });
}

/**
 * Proactively posts a short plain-text status line into this app's room via its incoming
 * webhook (Script Property PROGRESS_WEBHOOK_URL) — NOT the synchronous card-action response
 * (that channel only accepts a single reply per invocation; see handleBtwQuestion_). Any
 * failure here is swallowed — a progress ping is best-effort and must never break /btw's
 * actual answer.
 */
function postProgress_(text) {
  var url = PropertiesService.getScriptProperties().getProperty('PROGRESS_WEBHOOK_URL');
  if (!url) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
  } catch (err) {
    // best-effort only
  }
}

/** Minimal single-turn Claude call for /btw. ANTHROPIC_API_KEY lives in Script Properties. */
function askClaude_(question) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set in Script Properties');

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-sonnet-5',
      max_tokens: 1024,
      messages: [{ role: 'user', content: question }]
    }),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Claude API ' + code + ': ' + resp.getContentText().slice(0, 300));
  }
  var body = JSON.parse(resp.getContentText());
  var text = (body.content && body.content[0] && body.content[0].text) || '(응답 없음)';
  return text;
}

function truncate_(s, n) {
  s = String(s || '');
  return s.length > n ? s.slice(0, n) + '…' : s;
}

function getManualText_() {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return '';
  var sheet = SpreadsheetApp.openById(id).getSheetByName(MANUAL_TAB);
  if (!sheet) return '';
  return String(sheet.getRange('A1').getValue() || '');
}

function logRevisionFeedback_(feedbackText, email) {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return;
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(FEEDBACK_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(FEEDBACK_TAB);
    sheet.appendRow(['ts', 'submitterEmail', 'feedbackText', 'appliedAt']);
  }
  sheet.appendRow([new Date(), email, feedbackText, '']);
}

function logSubmission_(ticketId, body, confirmer) {
  var id = PropertiesService.getScriptProperties().getProperty(QUEUE_SHEET_PROP);
  if (!id) return;
  var sheet = SpreadsheetApp.openById(id).getSheetByName(QUEUE_TAB);
  var vals = sheet.getDataRange().getValues();
  var H = vals[0];
  var iId = H.indexOf('ticketId'), iSubAt = H.indexOf('submittedAt'), iConf = H.indexOf('confirmer');
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][iId]) === String(ticketId)) {
      sheet.getRange(r + 1, iSubAt + 1).setValue(new Date());
      sheet.getRange(r + 1, iConf + 1).setValue(confirmer);
      return;
    }
  }
}

/* ===================== Form input parsing (same shape as BadReview_ChatReport/chat_app) ===================== */

function formInputs_(event) {
  return (event.common && event.common.formInputs) ||
         (event.commonEventObject && event.commonEventObject.formInputs) ||
         event.formInputs || {};
}

function inputField_(inputs, name) {
  var v = inputs && inputs[name];
  if (!v) return null;
  return v[''] || v;
}

function extractString_(inputs, name) {
  var f = inputField_(inputs, name);
  if (!f) return '';
  var si = f.stringInputs;
  if (si && si.value && si.value.length) return String(si.value[0]);
  return '';
}

function actionParams_(event) {
  var list = (event.common && event.common.parameters) ||
             (event.commonEventObject && event.commonEventObject.parameters) ||
             (event.action && event.action.parameters) || {};
  if (Array.isArray(list)) {
    var o = {};
    list.forEach(function (p) { o[p.key] = p.value; });
    return o;
  }
  return list;
}

function escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Renders a pasted report body as card text: escape, then turn newlines into <br>. */
function reportHtml_(s) {
  return escapeHtml_(s).replace(/\n/g, '<br>');
}
