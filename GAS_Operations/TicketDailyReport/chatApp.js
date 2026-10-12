/********************************************************************
 * Google Chat app — date-picker report
 *
 * Lets a user pick an "Update 날짜" from a dropdown and see that day's
 * K_시트 pending-ticket table (read from the K_시트_history sheet that
 * archiveKSheetHistory() fills every run).
 *
 * REQUIRES (one-time, done in Google Cloud console — clasp cannot do this):
 *   1. Enable the "Google Chat API" on this script's GCP project.
 *   2. Chat API > Configuration:
 *        - App name / avatar / description
 *        - Functionality: "Receive 1:1 messages" + "Join spaces / group conversations"
 *        - Connection settings: "Apps Script" -> Deployment ID of this project
 *        - Slash command:  /report   (any numeric command ID)
 *   3. Visibility: make it available to yourself / the GCX group / the org.
 *   4. Add the app to the space, then type  /report .
 *
 * Entry points Apps Script Chat apps call automatically:
 *   onMessage        - a DM, @mention, or slash command
 *   onAddToSpace     - app added to a space
 *   loadDayReport    - card button click (onClick.action.function = "loadDayReport")
 ********************************************************************/

const CHAT_HISTORY_MAX_DATES = 30;
const CHAT_CARD_MAX_ROWS = 40;

function onMessage(event) {
  // Slash command or plain message -> show the picker with the newest date
  try {
    return buildPickerMessage_(null);
  } catch (e) {
    return { text: '⚠️ onMessage 오류: ' + (e && e.stack ? e.stack : e) };
  }
}

function onAddToSpace(event) {
  return { text: '`/report` 를 입력하면 날짜별(Update 날짜) Pending 티켓 리포트를 볼 수 있습니다.' };
}

function onRemoveFromSpace(event) {
  // nothing to clean up
}

// Card button handler — Apps Script invokes the function named in
// onClick.action.function. Kept tolerant of the different event shapes.
function loadDayReport(event) {
  try {
    const key = chatFormValue_(event, 'reportDate');
    const msg = buildPickerMessage_(key);
    msg.actionResponse = { type: 'UPDATE_MESSAGE' };
    return msg;
  } catch (e) {
    return { text: '⚠️ loadDayReport 오류: ' + (e && e.stack ? e.stack : e) };
  }
}

/**
 * Run this once from the Apps Script editor (Run > test_chatPicker) to trigger
 * the OAuth authorization prompt (Sheets access). Approve it, then /report works.
 * Also logs the card JSON so you can eyeball it.
 */
function test_chatPicker() {
  const msg = buildPickerMessage_(null);
  Logger.log(JSON.stringify(msg, null, 2));
  return msg;
}

/* ---------- internals ---------- */

function chatFormValue_(event, name) {
  try {
    const common = event && event.common;
    if (common && common.formInputs && common.formInputs[name]) {
      const fi = common.formInputs[name];
      if (fi.stringInputs && fi.stringInputs.value && fi.stringInputs.value.length) {
        return String(fi.stringInputs.value[0]);
      }
    }
    if (common && common.parameters && common.parameters[name]) return String(common.parameters[name]);
  } catch (e) { /* fall through */ }
  return '';
}

// Read K_시트_history -> { order: [key,...] newest first, labelByKey, rowsByKey: {key: [[seq..reason],...]} }
function readKSheetHistory_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hist = ss.getSheetByName(KSHEET_HISTORY_NAME);
  if (!hist || hist.getLastRow() < 2) return { order: [], labelByKey: {}, rowsByKey: {} };

  const data = hist.getRange(2, 1, hist.getLastRow() - 1, 10).getValues(); // A..J
  const labelByKey = {}, rowsByKey = {}, seen = [];
  data.forEach(d => {
    const key = String(d[0]).trim();
    if (!key) return;
    if (!rowsByKey[key]) { rowsByKey[key] = []; seen.push(key); }
    labelByKey[key] = String(d[1] || key);
    rowsByKey[key].push(d.slice(2)); // seq, country, brand, category, qty, owner, device, reason
  });

  const order = seen.sort().reverse().slice(0, CHAT_HISTORY_MAX_DATES); // yyyy-MM-dd sorts lexically
  return { order: order, labelByKey: labelByKey, rowsByKey: rowsByKey };
}

function buildPickerMessage_(selectedKey) {
  const h = readKSheetHistory_();
  if (!h.order.length) {
    return { text: 'K_시트_history 에 아직 저장된 리포트가 없습니다. (매 평일 실행마다 한 줄씩 쌓입니다.)' };
  }

  const key = (selectedKey && h.rowsByKey[selectedKey]) ? selectedKey : h.order[0];
  const label = h.labelByKey[key] || key;

  const items = h.order.map(k => ({ text: h.labelByKey[k] || k, value: k, selected: k === key }));

  const allRows = h.rowsByKey[key] || [];

  return {
    cardsV2: [{
      cardId: 'kSheetPicker',
      card: {
        header: {
          title: label + ' · Pending 티켓',
          subtitle: 'Update 날짜별 리포트',
          imageUrl: 'https://img.icons8.com/color/512/zendesk.png',
          imageType: 'SQUARE',
          imageAltText: 'Zendesk'
        },
        sections: [
          { widgets: [
            { selectionInput: { name: 'reportDate', label: 'Update 날짜', type: 'DROPDOWN', items: items } },
            { buttonList: { buttons: [
              { text: '리포트 보기', type: 'FILLED', icon: { knownIcon: 'DESCRIPTION' },
                onClick: { action: { function: 'loadDayReport' } } }
            ] } },
            { divider: {} }
          ]},
          // Same layout as the daily card (no PIC — not tagged for T2 Escalated since 2026-10-12)
          ...buildKSheetCardSections_(allRows, { maxRows: CHAT_CARD_MAX_ROWS, withButtons: false })
        ]
      }
    }]
  };
}
