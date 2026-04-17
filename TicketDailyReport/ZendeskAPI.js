/********************************
 * SETTINGS
 ********************************/
const ZENDESK_EMAIL = 'kjw@spigen.com';
const ZENDESK_TOKEN = 'QhM2AiBYwTZTSb04Qjor918PHtttxp8xAzCFfFsg';
const ZENDESK_SUBDOMAIN = 'spigenhelp';
const SPREADSHEET_ID = '10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo';

/********************************
 * COMMON HELPERS
 ********************************/
function getZendeskTicketsByView(viewId) {
  const headers = {
    "Authorization": "Basic " + Utilities.base64Encode(`${ZENDESK_EMAIL}/token:${ZENDESK_TOKEN}`)
  };
  const url = `https://${ZENDESK_SUBDOMAIN}.zendesk.com/api/v2/views/${viewId}/tickets.json`;
  const response = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
  const json = JSON.parse(response.getContentText());
  return json.tickets || [];
}

function removeDuplicatesByTicketID(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const seen = new Set();
  const deduped = [data[0]]; // Keep headers
  for (let i = 1; i < data.length; i++) {
    const ticketId = data[i][0];
    if (!seen.has(ticketId)) {
      seen.add(ticketId);
      deduped.push(data[i]);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, deduped.length, deduped[0].length).setValues(deduped);
}

/********************************
 * 1. Fetch Multiple Views → Zendesk_Daily
 ********************************/
function fetchZendeskViewToSheet() {
  const viewIds = [
    '360102672972',
    '360121546552',
    '360121545992',
    '360108214671',
    '28990066416793',
    '360096790151',
    '19940259463705',
    '360103290632',
    '37502606662809'
  ];

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Zendesk_Daily');
  sheet.clearContents();

  const headersRow = [
    'Ticket ID', 'Status', 'Subject', 'Priority',
    'Created At', 'Updated At', 'Assignee ID', 'Requester ID', 'Tags'
  ];
  sheet.getRange(1, 1, 1, headersRow.length).setValues([headersRow]);

  let allRows = [];

  for (const viewId of viewIds) {
    const tickets = getZendeskTicketsByView(viewId);
    if (tickets.length === 0) {
      Logger.log(`No tickets found in view ${viewId}`);
      continue;
    }
    const rows = tickets.map(t => [
      t.id || '',
      t.status || '',
      t.subject || '',
      t.priority || '',
      t.created_at || '',
      t.updated_at || '',
      t.assignee_id || '',
      t.requester_id || '',
      (t.tags || []).join(', ')
    ]);
    allRows = allRows.concat(rows);
  }

  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, headersRow.length).setValues(allRows);
    removeDuplicatesByTicketID(sheet);
  } else {
    Logger.log("No tickets found in any of the views.");
  }
}

/********************************
 * 2. Fetch Ksheet View → K_시트
 ********************************/
function fetchZendeskViewToKsheet() {
  const viewId = '49523632520985';
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('K_시트');

  // Clear B~G from row 5 to bottom
  const maxRows = sheet.getMaxRows();
  if (maxRows > 4) {
    sheet.getRange(5, 2, maxRows - 4, 6).clearContent();
  }

  const tickets = getZendeskTicketsByView(viewId);
  if (tickets.length === 0) {
    Logger.log("No tickets found in view.");
    updateB3DateBanner(sheet);
    return;
  }

  // Format helpers
  const formatBrandCode = raw => ({
    'spigen_case_': 'Spigen(CASE)',
    'spigen_steinheil_': 'Spigen(Steinheil)',
    'spigen_odm_': 'Spigen(ODM)',
    'spigen_pacc._': 'Spigen(PAcc.)',
    'spigen_new_biz_': 'Spigen(New Biz)',
    'n/a': 'n/a'
  }[raw] || raw);

  const formatCategoryCode = raw => ({
    '1._invoice': '1. Invoice',
    '2._delivery': '2. Delivery',
    '3._exchange': '3. Exchange',
    '4._issue': '4. Product Issue',
    '5._fbm': '5. FBM',
    '6._product_inquiry': '6. Product Inquiry',
    '7._other_inquiry': '7. Other Inquiry',
    '8._문의_사항_파악_불가': '8. 문의 사항 파악 불가'
  }[raw] || raw);

  const getPIC = country => {
    const groupA = ['DE', 'FR', 'IT', 'UK', 'ES'];

    // Base date: Monday, 2025-08-11 00:00 KST
    const baseDateKST = new Date('2025-08-11T00:00:00+09:00');

    // Current date/time in KST
    const nowKST = new Date(
      new Date().toLocaleString("en-US", { timeZone: "Asia/Seoul" })
    );

    // Weeks since base date
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    const weekIndex = Math.floor((nowKST - baseDateKST) / msPerWeek);

    // Even week index → KJW, Odd week index → LYS for Group A
    const assignKJWThisWeek = (weekIndex % 2 === 0);

    return groupA.includes(country)
      ? (assignKJWThisWeek ? 'KJW' : 'LYS')
      : (assignKJWThisWeek ? 'LYS' : 'KJW');
  };

  // Process rows
  const rawRows = tickets
    .filter(t => t.status && t.status.toLowerCase() === 'pending')
    .map(t => {
      const cf = t.custom_fields || [];
      const brandRaw = cf.find(f => f.id === 5495572594201)?.value || '';
      const countryRaw = cf.find(f => f.id === 4513936822297)?.value || '';
      const categoryRaw = cf.find(f => f.id === 900006613446)?.value || '';
      return [
        (countryRaw || '').toUpperCase(),
        formatBrandCode(brandRaw),
        formatCategoryCode(categoryRaw)
      ];
    });

  // Group by country|brand|category
  const grouped = {};
  rawRows.forEach(([country, brand, category]) => {
    if (!country) return;
    const key = `${country}|${brand}|${category}`;
    grouped[key] = (grouped[key] || 0) + 1;
  });

  // Final rows
  const finalRows = Object.entries(grouped).map(([key, count]) => {
    const [country, brand, category] = key.split('|');
    return [country, brand, category, count, getPIC(country)];
  });

  if (finalRows.length > 0) {
    sheet.getRange(5, 3, finalRows.length, 5).setValues(finalRows);
    const seq = Array.from({ length: finalRows.length }, (_, i) => [i + 1]);
    sheet.getRange(5, 2, finalRows.length, 1).setValues(seq);
  }

  updateB3DateBanner(sheet);
}

/********************************
 * Banner Date Helper
 ********************************/
function updateB3DateBanner(sheet) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd");
  const dayName = now.toLocaleDateString('en-US', { weekday: 'short' });
  sheet.getRange('B3').setValue(`${dateStr} (${dayName}) T2 업무시작 Pending Ticket 수`);
}
