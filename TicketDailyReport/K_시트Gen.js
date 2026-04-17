function kSheetToChat() {
  const spreadsheetId = '10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo';
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('K_시트');

  // Title (B3) and total (G3)
  const titleText = String(sheet.getRange('B3').getValue() || '');
  const manualTotal = sheet.getRange('G3').getValue();

  // Headers + data from B4:G
  const all = sheet.getRange('B:G').getValues();
  const headers = all[3]; // B4..G4
  let last = all.length;
  while (last > 5 && all[last - 1].every(v => v === '' || v === null)) last--;
  const rows = all.slice(4, last); // B5..G(last)
  if (!rows.length) return;

  // Build compact row strings + compute totals
  const MAX_ROWS = 25;
  const clip = rows.slice(0, MAX_ROWS);

  let totalQty = 0;
  let lysTotal = 0;
  let kjwTotal = 0;

  const listWidgets = [];

  // Header line (as topLabel)
  listWidgets.push({
    decoratedText: {
      topLabel: headers.join(' | '),
      text: ''
    }
  });

  for (const r of clip) {
    const no = safeStr(r[0]);
    const country = safeStr(r[1]);
    const brand = safeStr(r[2]);
    const category = safeStr(r[3]);
    const qty = toInt(r[4]);
    const owner = safeStr(r[5]);

    totalQty += qty;
    const ownerUpper = owner.toUpperCase();
    if (ownerUpper === 'LYS') lysTotal += qty;
    if (ownerUpper === 'KJW') kjwTotal += qty;

    const iso = normalizeIso(country);
    const flag = toFlagEmoji(iso);
    const line = `${no}. ${flag ? flag + ' ' : ''}${iso || country} | ${brand} | ${category} | ${qty} | ${owner}`;
    listWidgets.push({ decoratedText: { text: line } });
  }

  // If G3 is present, prefer it; otherwise use computed sum
  const headlineTotal = (manualTotal !== '' && manualTotal !== null) ? Number(manualTotal) : totalQty;

  // Totals line (All, LYS, KJW)
  const totalsLine = `All: ${headlineTotal} | LYS: ${lysTotal} | KJW: ${kjwTotal}`;

  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}`;
  const payload = {
    text: titleText || 'K_시트 Pending Ticket Snapshot',
    cardsV2: [
      {
        cardId: 'k_sheet_compact',
        card: {
        header: {
          title: titleText || 'K_시트 Pending Ticket 수',
          subtitle: '담당자별 Pending 티켓 현황',
          imageUrl: 'https://img.icons8.com/color/512/zendesk.png',
          imageType: 'SQUARE',
          imageAltText: 'Zendesk'
        },
          sections: [
            {
              widgets: [
                {
                  decoratedText: {
                    topLabel: 'Ticket Totals',
                    text: totalsLine
                  }
                },
                {
                              buttonList: {
              buttons: [
                {
                  text: 'Start Zendesk',
                  onClick: {
                    openLink: {
                      url: 'https://spigenhelp.zendesk.com/agent/filters/360103290632'
                    }
                  }
                }
              ]
            },
                },
                { divider: {} }
              ]
            },
            { widgets: listWidgets },
            ...(rows.length > MAX_ROWS
              ? [{
                  widgets: [{
                    decoratedText: {
                      topLabel: 'Note',
                      text: `Showing first ${MAX_ROWS} rows — open sheet for all.`
                    }
                  }]
                }]
              : [])
          ]
        }
      }
    ]
  };

  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

/* ===== helpers ===== */
function safeStr(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function toInt(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
function normalizeIso(code) {
  const raw = safeStr(code).toUpperCase();
  if (!raw) return '';
  return raw === 'UK' ? 'GB' : raw;
}
function toFlagEmoji(iso2) {
  if (!iso2 || iso2.length !== 2) return '';
  const A = 0x41; const REGIONAL = 0x1F1E6;
  const c1 = iso2.charCodeAt(0), c2 = iso2.charCodeAt(1);
  if (c1 < A || c1 > 0x5A || c2 < A || c2 > 0x5A) return '';
  return String.fromCodePoint(REGIONAL + (c1 - A), REGIONAL + (c2 - A));
}
