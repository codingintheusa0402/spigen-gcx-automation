/********************************
 * PurchaseDate_Sync — Zendesk → monday.com "Purchase Date" bridge
 *
 * WHY THIS EXISTS
 * The native monday↔Zendesk integration on the Case+CP boards (Galaxy Z8
 * 18421346787, Pixel 11 18425190666) cannot map Zendesk CUSTOM fields to
 * board columns — its date
 * dropdown only offers the system fields (Created at / Due at / Updated at).
 * "Purchase Date" (Zendesk custom field 360019586172) therefore never
 * reaches the board's Purchase Date column.
 *
 * HOW IT WORKS  (rewritten 2026-09-08 — scheduled batch)
 * Previously this ran as a Zendesk webhook (doPost) that fired ~5-6x/min and
 * walked the ENTIRE board up to 4x on every single call — the top consumer
 * of the account's monday.com API budget. It now runs as a scheduled batch:
 *
 *   scheduledPurchaseDateSync()  — installed by setupPurchaseDateTriggers()
 *   to run every SYNC_EVERY_MINUTES minutes (15 → ~96 runs/day):
 *     1. ONE walk of EACH board in MONDAY_BOARD_IDS (500 items/page) to
 *        collect every item that has a linked Zendesk ticket, plus its
 *        current Purchase Date cell. (2026-09-11: Pixel 11 board added —
 *        both boards share the same column ids.)
 *     2. Fetch those tickets from Zendesk in bulk (show_many, 100 ids/call).
 *     3. Write the Purchase Date to an item ONLY when it differs from what's
 *        already on the board.
 *   A script lock makes an overlapping tick a no-op, so a long run can never
 *   pile up. The board is small (~960 items = 2 pages), so a run costs ~2
 *   monday reads + only-changed writes → ~200-300 monday calls/day total
 *   (96 runs), versus the ~40,000/day the old webhook was burning.
 *   MONDAY_CALLS_MAX_PER_RUN is a hard ceiling: mondayGql_ throws past it so
 *   no future change can quietly turn this back into a runaway.
 *
 * The old webhook path is disabled — doPost() is now a no-op. After
 * deploying this version, also deactivate the Zendesk trigger + webhook that
 * used to POST here (Zendesk Admin → Business rules → Triggers, and Apps and
 * integrations → Webhooks).
 *
 * backfillPurchaseDates() is kept as a manual alias for scheduledPurchaseDateSync().
 ********************************/

const ZENDESK_EMAIL = 'kjw@spigen.com';
const ZENDESK_TOKEN = 'QhM2AiBYwTZTSb04Qjor918PHtttxp8xAzCFfFsg';
const ZENDESK_SUBDOMAIN = 'spigenhelp';
const ZD_PURCHASE_DATE_FIELD = 360019586172; // custom ticket field "Purchase Date" (type: date)

const MONDAY_TOKEN = 'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjU0ODE3MjIzOSwiYWFpIjoxMSwidWlkIjozMTE0NDEyMSwiaWFkIjoiMjAyNS0wOC0wOFQwNToyMTozNS40ODdaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTExNjU5NTcsInJnbiI6InVzZTEifQ._Z9iAbMMY9bvJnCG3jFwdUIHMaw8aihN2pcRNnkFUVM';
// Every Case+CP board to keep in sync. All share the same column ids below
// (cloned from one template). Add a board here when a new series launches.
const MONDAY_BOARD_IDS = [
  18421346787, // 📌Galaxy Z8 Case+CP
  18425190666, // 📌Pixel 11 Case+CP   (added 2026-09-11)
];
const MONDAY_DATE_COL = 'date_mm59ejfp';          // "Purchase Date" (date)
const MONDAY_TICKET_COL = 'integration_mm0fzmv0'; // "Zendesk Ticket" (integration)

// Shared secret the old Zendesk webhook sent as the `?secret=` query param.
// Kept only so doPost can still recognise stray webhook traffic; unused otherwise.
const WEBHOOK_SECRET = '8GD3uY_vYU5N9GJlD0T1y1b9jJylrPnv21QmeqBSsKU';

// How often the batch sync runs. 15 min → 96 runs/day (~100/day).
// GAS everyMinutes() only accepts 1, 5, 10, 15 or 30.
const SYNC_EVERY_MINUTES = 15;

const BOARD_PAGE_LIMIT = 500;    // monday items_page max
const ZD_SHOW_MANY_CHUNK = 100;  // Zendesk tickets/show_many max ids per call
const WRITE_SLEEP_MS = 120;      // small gap between monday writes

// Hard ceiling on monday API calls per single run. A healthy run is ~2 reads
// plus a handful of writes. 50 * 96 runs/day = 4,800/day absolute worst case
// (< 5,000 by design); real usage is ~200-300/day. If a bulk date edit needs
// more than ~48 writes in one 15-min window the run stops at the cap and the
// remainder is picked up by the next run (it re-walks and still sees the diff),
// so nothing is lost — and a genuine bug/loop bails loud instead of repeating
// the 2026-09 runaway.
const MONDAY_CALLS_MAX_PER_RUN = 50;
let _mondayCallCount = 0; // reset at the top of each run

// ── Zendesk helpers ───────────────────────────────────────────────────────────

function zdAuthHeader_() {
  return 'Basic ' + Utilities.base64Encode(`${ZENDESK_EMAIL}/token:${ZENDESK_TOKEN}`);
}

function zdGetTicket_(ticketId) {
  const resp = UrlFetchApp.fetch(
    `https://${ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/${ticketId}.json`,
    { headers: { Authorization: zdAuthHeader_() }, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Zendesk GET ticket ${ticketId} -> ${resp.getResponseCode()}: ${resp.getContentText()}`);
  }
  return JSON.parse(resp.getContentText()).ticket;
}

function zdPurchaseDate_(ticket) {
  const f = (ticket.custom_fields || []).find(cf => cf.id === ZD_PURCHASE_DATE_FIELD);
  return f && f.value ? String(f.value) : null; // raw value is YYYY-MM-DD
}

// Bulk fetch: ticketId (Number) -> Purchase Date string 'YYYY-MM-DD' | null.
// Uses tickets/show_many (100 ids/request) instead of one GET per ticket.
function zdGetPurchaseDatesForTickets_(ticketIds) {
  const out = {};
  const ids = Array.from(new Set(ticketIds.map(Number).filter(Boolean)));
  for (let i = 0; i < ids.length; i += ZD_SHOW_MANY_CHUNK) {
    const chunk = ids.slice(i, i + ZD_SHOW_MANY_CHUNK);
    const resp = UrlFetchApp.fetch(
      `https://${ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/tickets/show_many.json?ids=${chunk.join(',')}`,
      { headers: { Authorization: zdAuthHeader_() }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) {
      throw new Error(`Zendesk show_many -> ${resp.getResponseCode()}: ${resp.getContentText().slice(0, 500)}`);
    }
    (JSON.parse(resp.getContentText()).tickets || []).forEach(t => {
      out[Number(t.id)] = zdPurchaseDate_(t);
    });
    Utilities.sleep(200);
  }
  return out;
}

// ── monday helpers ────────────────────────────────────────────────────────────

const MONDAY_BUDGET_ERR = 'MONDAY_CALLS_MAX_PER_RUN';

function mondayGql_(query) {
  if (++_mondayCallCount > MONDAY_CALLS_MAX_PER_RUN) {
    throw new Error(`${MONDAY_BUDGET_ERR} (${MONDAY_CALLS_MAX_PER_RUN}) reached — stopping this run; remainder resumes next tick.`);
  }
  const resp = UrlFetchApp.fetch('https://api.monday.com/v2', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: MONDAY_TOKEN },
    payload: JSON.stringify({ query }),
    muteHttpExceptions: true,
  });
  const body = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() !== 200 || body.errors) {
    throw new Error('monday API: ' + resp.getContentText());
  }
  return body.data;
}

// ONE walk of every board in MONDAY_BOARD_IDS. Returns
// [{ boardId, itemId, ticketId (Number), dateText }] for every item whose
// "Zendesk Ticket" integration column holds an entity_id.
function collectLinkedItems_() {
  const linked = [];
  MONDAY_BOARD_IDS.forEach(boardId => collectLinkedItemsFromBoard_(boardId, linked));
  return linked;
}

function collectLinkedItemsFromBoard_(boardId, linked) {
  let cursor = null;
  do {
    const pageArgs = cursor
      ? `limit: ${BOARD_PAGE_LIMIT}, cursor: "${cursor}"`
      : `limit: ${BOARD_PAGE_LIMIT}`;
    const data = mondayGql_(`query {
      boards(ids: [${boardId}]) {
        items_page(${pageArgs}) {
          cursor
          items {
            id
            column_values(ids: ["${MONDAY_TICKET_COL}", "${MONDAY_DATE_COL}"]) { id value text }
          }
        }
      }
    }`);
    const page = data.boards[0].items_page;
    for (const item of page.items) {
      const tickCol = item.column_values.find(c => c.id === MONDAY_TICKET_COL);
      if (!tickCol || !tickCol.value) continue;
      let entityId;
      try { entityId = JSON.parse(tickCol.value).entity_id; } catch (e) { continue; }
      if (!entityId) continue;
      const dateCol = item.column_values.find(c => c.id === MONDAY_DATE_COL);
      linked.push({
        boardId: boardId,
        itemId: item.id,
        ticketId: Number(entityId),
        dateText: (dateCol && dateCol.text) || '',
      });
    }
    cursor = page.cursor;
  } while (cursor);
}

function setItemPurchaseDate_(boardId, itemId, isoDate) {
  const colVals = JSON.stringify({ [MONDAY_DATE_COL]: { date: isoDate } })
    .replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  mondayGql_(`mutation {
    change_multiple_column_values(board_id: ${boardId}, item_id: ${itemId}, column_values: "${colVals}") { id }
  }`);
}

// ── Scheduled batch sync (installed by setupPurchaseDateTriggers) ─────────────

function scheduledPurchaseDateSync() {
  const startedAt = new Date();

  // A frequent timer must never let two runs overlap and double-walk the board.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('scheduledPurchaseDateSync: another run holds the lock — skipping this tick.');
    return;
  }
  try {
    _runPurchaseDateSync_(startedAt);
  } finally {
    lock.releaseLock();
  }
}

function _runPurchaseDateSync_(startedAt) {
  _mondayCallCount = 0;

  const linked = collectLinkedItems_();
  Logger.log(`Board walk (${MONDAY_BOARD_IDS.length} board(s)): ${linked.length} item(s) linked to a Zendesk ticket (${_mondayCallCount} monday read call(s)).`);
  if (!linked.length) return;

  const dateByTicket = zdGetPurchaseDatesForTickets_(linked.map(x => x.ticketId));

  let updated = 0, unchanged = 0, noDate = 0, failed = 0, cappedAt = 0;
  for (const it of linked) {
    const isoDate = dateByTicket[it.ticketId];
    if (!isoDate) { noDate++; continue; }
    if (isoDate === it.dateText) { unchanged++; continue; }
    try {
      setItemPurchaseDate_(it.boardId, it.itemId, isoDate);
      updated++;
      Logger.log(`ticket #${it.ticketId} → board ${it.boardId} item ${it.itemId}: Purchase Date ${it.dateText || '(empty)'} → ${isoDate}`);
    } catch (err) {
      if (String(err).indexOf(MONDAY_BUDGET_ERR) !== -1) {
        cappedAt = updated;
        Logger.log(`Hit MONDAY_CALLS_MAX_PER_RUN after ${updated} write(s) — leaving the rest for the next run.`);
        break;
      }
      failed++;
      Logger.log(`item ${it.itemId} (ticket #${it.ticketId}) → ${isoDate} FAILED: ${err}`);
    }
    Utilities.sleep(WRITE_SLEEP_MS);
  }

  Logger.log(
    `scheduledPurchaseDateSync done in ${((new Date() - startedAt) / 1000).toFixed(1)}s — ` +
    `updated ${updated}, unchanged ${unchanged}, ticket has no purchase date ${noDate}, failed ${failed}` +
    (cappedAt ? `, CAPPED at ${MONDAY_CALLS_MAX_PER_RUN} monday calls` : '') +
    `. Total monday calls this run: ${_mondayCallCount}.`
  );
}

// Run once from the GAS editor (needs the script.scriptapp OAuth consent, so
// it 403s if invoked headlessly). Replaces any existing scheduledPurchaseDateSync
// trigger with a single every-SYNC_EVERY_MINUTES timer. Safe to re-run.
function setupPurchaseDateTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'scheduledPurchaseDateSync' || fn === 'backfillPurchaseDates') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scheduledPurchaseDateSync')
    .timeBased().everyMinutes(SYNC_EVERY_MINUTES).create();
  Logger.log(`Installed 1 trigger for scheduledPurchaseDateSync every ${SYNC_EVERY_MINUTES} min (~${Math.round(1440 / SYNC_EVERY_MINUTES)} runs/day).`);
}

// Kept for compatibility with older docs/muscle memory. The scheduled sync
// supersedes the old "fill empty only" backfill — it also refreshes items
// whose Purchase Date changed on the ticket after it was first written.
function backfillPurchaseDates() {
  scheduledPurchaseDateSync();
}

// ── Web App entry point — DISABLED 2026-09-08 ────────────────────────────────
// This script now syncs on a schedule (scheduledPurchaseDateSync). The Zendesk
// webhook that used to POST here should be deactivated; until it is, respond
// cheaply and do NO monday/Zendesk work so it can't run up the API bill again.
function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      disabled: true,
      note: 'PurchaseDate_Sync runs on a twice-daily schedule now; this webhook is a no-op. Deactivate the Zendesk trigger/webhook.',
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Smoke test (run from the GAS editor) ──────────────────────────────────────

// Reports what the batch would do for one ticket (does one board walk).
function testSyncOneTicket() {
  const TICKET_ID = 1000153779; // Jane's test ticket (item "1010101010")
  const linked = collectLinkedItems_().filter(x => x.ticketId === TICKET_ID);
  const dates = zdGetPurchaseDatesForTickets_([TICKET_ID]);
  Logger.log(JSON.stringify({
    ticketId: TICKET_ID,
    zendeskPurchaseDate: dates[TICKET_ID] || null,
    linkedBoardItems: linked,
  }, null, 2));
}
