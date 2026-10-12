function kSheetToChat() {
  const spreadsheetId = '10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo';
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('K_시트');

  // Title (B3) and total (G3)
  const titleText = String(sheet.getRange('B3').getValue() || '');
  const manualTotal = sheet.getRange('G3').getValue();

  // Headers + data from B4:I  (B no, C country, D brand, E category, F qty, G owner, H device, I 인입사유)
  const all = sheet.getRange('B:I').getValues();
  let last = all.length;
  while (last > 5 && all[last - 1].every(v => v === '' || v === null)) last--;
  const rows = all.slice(4, last); // B5..I(last)
  if (!rows.length) return;

  // If G3 is present, prefer it; otherwise the card sums the rows
  const headlineTotal = (manualTotal !== '' && manualTotal !== null) ? Number(manualTotal) : null;

  const payload = {
    text: titleText || 'K_시트 Pending Ticket Snapshot',
    cardsV2: [{ cardId: 'k_sheet_compact', card: buildKSheetCard_(titleText, rows, headlineTotal) }]
  };

  const res = UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

/* ===== cardsV2 builder (shared with chatApp.js) ===== */

// No person-in-charge tagging for T2 Escalated tickets (decided 2026-10-12) — the card
// shows no PIC / owner subtotal. K_시트 column G is still written but no longer displayed.
const KSHEET_ZENDESK_URL = 'https://spigenhelp.zendesk.com/agent/filters/360103290632';
const KSHEET_SHEET_URL = 'https://docs.google.com/spreadsheets/d/10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo/edit#gid=1528385180';
const KSHEET_CARD_MAX_ROWS = 60; // 1 widget per ticket row; cards drop sections past 100 widgets

// Category order + accent colour/icon (unknown categories go last, grey)
const KSHEET_CATEGORY_STYLE = {
  'Product Issue':   { order: 1, color: '#D93025', icon: 'report' },
  'Product Inquiry': { order: 2, color: '#1A73E8', icon: 'help' }
};
const KSHEET_CATEGORY_DEFAULT = { order: 9, color: '#E37400', icon: 'label' };

function kSheetCategoryStyle_(category) {
  return KSHEET_CATEGORY_STYLE[category] || KSHEET_CATEGORY_DEFAULT;
}

// Full card for the daily webhook post.
function buildKSheetCard_(titleText, rawRows, headlineTotal) {
  const m = (titleText || '').match(/^(.*?)\s*Pending Ticket 수\s*$/);
  return {
    header: {
      title: m ? m[1] + ' Pending 티켓' : (titleText || 'K_시트 Pending 티켓'),
      subtitle: 'Zendesk T2 Escalated · 카테고리별 Pending 현황',
      imageUrl: 'https://img.icons8.com/color/512/zendesk.png',
      imageType: 'SQUARE',
      imageAltText: 'Zendesk'
    },
    sections: buildKSheetCardSections_(rawRows, {
      headlineTotal: headlineTotal,
      maxRows: KSHEET_CARD_MAX_ROWS,
      withButtons: true
    })
  };
}

// Sections: summary (total, per-category, per-country, buttons) + one section per category
// listing its tickets + a truncation note. opts: { headlineTotal, maxRows, withButtons }
function buildKSheetCardSections_(rawRows, opts) {
  opts = opts || {};
  const maxRows = opts.maxRows || KSHEET_CARD_MAX_ROWS;
  const items = rawRows.map(cleanKSheetDisplayRow_);

  const rowSum = items.reduce((s, it) => s + it.qty, 0);
  const total = (opts.headlineTotal !== null && opts.headlineTotal !== undefined && !isNaN(opts.headlineTotal))
    ? opts.headlineTotal : rowSum;

  // Per-category / per-country sums over ALL rows (not just the displayed ones)
  const catSum = {}, ccSum = {};
  items.forEach(it => {
    catSum[it.category] = (catSum[it.category] || 0) + it.qty;
    ccSum[it.cc] = (ccSum[it.cc] || 0) + it.qty;
  });
  const cats = Object.keys(catSum).sort((a, b) =>
    (kSheetCategoryStyle_(a).order - kSheetCategoryStyle_(b).order) || (catSum[b] - catSum[a]));

  const catWidgets = cats.map(c => {
    const st = kSheetCategoryStyle_(c);
    return { decoratedText: {
      startIcon: { materialIcon: { name: st.icon } },
      topLabel: c,
      text: '<b><font color="' + st.color + '">' + catSum[c] + '</font></b> 건'
    } };
  });

  const ccLine = Object.keys(ccSum)
    .sort((a, b) => (ccSum[b] - ccSum[a]) || a.localeCompare(b))
    .map(cc => flagCc_(cc) + ' <b>' + ccSum[cc] + '</b>')
    .join(' &nbsp;·&nbsp; ');

  const summaryWidgets = [
    { columns: { columnItems: [
      { horizontalSizeStyle: 'FILL_MINIMUM_SPACE', verticalAlignment: 'CENTER', widgets: [
        { decoratedText: {
          startIcon: { materialIcon: { name: 'pending_actions' } },
          topLabel: 'Total Pending',
          text: '<b><font color="#1A73E8">' + total + '</font></b> 건'
        } }
      ] },
      { horizontalSizeStyle: 'FILL_AVAILABLE_SPACE', verticalAlignment: 'CENTER', widgets: catWidgets }
    ] } },
    { decoratedText: {
      startIcon: { materialIcon: { name: 'public' } },
      topLabel: '국가별',
      text: ccLine || '-',
      wrapText: true
    } }
  ];
  if (opts.withButtons) {
    summaryWidgets.push({ buttonList: { buttons: [
      { text: 'Start Zendesk', type: 'FILLED',
        icon: { materialIcon: { name: 'open_in_new' } },
        onClick: { openLink: { url: KSHEET_ZENDESK_URL } } },
      { text: 'K_시트 열기', type: 'OUTLINED',
        icon: { materialIcon: { name: 'table_view' } },
        onClick: { openLink: { url: KSHEET_SHEET_URL } } }
    ] } });
  }

  const sections = [{ widgets: summaryWidgets }];

  // Ticket rows grouped by category, sheet order kept inside each group
  const shown = items.slice(0, maxRows);
  cats.forEach(c => {
    const st = kSheetCategoryStyle_(c);
    const rows = shown.filter(it => it.category === c);
    if (!rows.length) return;
    sections.push({
      header: '<font color="' + st.color + '"><b>' + htmlEscape_(c) + '</b></font> · ' + catSum[c] + '건',
      widgets: rows.map(it => ({ decoratedText: {
        topLabel: '#' + it.no + '  ' + flagCc_(it.cc) + ' · ' + it.brand + ' · Qty ' + it.qty,
        text: '<b>' + htmlEscape_(it.device) + '</b><br><font color="' + st.color + '">' +
              htmlEscape_(it.reason) + '</font>',
        wrapText: true
      } }))
    });
  });

  if (items.length > maxRows) {
    sections.push({ widgets: [{ decoratedText: {
      startIcon: { materialIcon: { name: 'info' } },
      topLabel: 'Note',
      text: '상위 ' + maxRows + '행만 표시 (전체 ' + items.length + '행) — 전체는 K_시트를 확인하세요.',
      wrapText: true
    } }] });
  }
  return sections;
}

/* ===== helpers ===== */

// Turn one raw K_시트 / K_시트_history row [seq, country, brand, category, qty, owner, device, reason]
// into { no, cc, brand, category, device, reason, qty } with the card-only cleanups
// (Brand strip+UPPER, Category drop "N. ", 인입사유 drop "(XXX)_"). Owner is not displayed.
function cleanKSheetDisplayRow_(raw) {
  const no = safeStr(raw[0]).replace(/\.$/, '');
  const cc = normalizeIso(safeStr(raw[1])) || safeStr(raw[1]) || '-';
  const brand = safeStr(raw[2])
    .replace(/spigen_/g, '')
    .replace(/Spigen/g, '')
    .replace(/[()]/g, '')
    .replace(/^_+|_+$/g, '')
    .trim()
    .toUpperCase();
  const category = safeStr(raw[3]).replace(/^\d+\.\s*/, '');
  const device = safeStr(raw[6]);
  const reason = safeStr(raw[7]).replace(/^\([^)]*\)_/, '');
  return {
    no: no,
    cc: cc,
    brand: brand || '-',
    category: category || '기타',
    device: device || '-',
    reason: reason || '-',
    qty: toInt(raw[4])
  };
}

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

// "🇩🇪 DE" (just "DE" when the code has no flag)
function flagCc_(cc) {
  const f = toFlagEmoji(cc);
  return f ? f + ' ' + cc : cc;
}

function htmlEscape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
