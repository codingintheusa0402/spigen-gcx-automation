// ── Trigger end date for this project ────────────────────────────────────────
const MASTER_END_DATE = '2026-05-08';

// ── Test webhook (Private ❤️) ─────────────────────────────────────────────────
const STATUS_WEBHOOK = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=T_rTrPKTYq6biglb8kRL3GOVfQg3AAOH-JPKELutbAY';

// ── All monitored projects ────────────────────────────────────────────────────
// Update end dates here whenever createTriggers() is re-run in each project.
const TRIGGER_PROJECTS = [
  { name: 'MasterTrigger',     endDate: '2026-05-08', time: '06:00 KST (Mon–Fri)' },
  { name: 'TicketDailyReport', endDate: '2026-04-26', time: '09:00 KST (Mon–Fri)' },
  { name: 'TCT Chat Log_GCX',  endDate: '2026-04-26', time: '17:30 KST (Mon–Fri, Thu 15:30)' },
];

// ── Count remaining weekdays from tomorrow (KST) through endDateStr ───────────
function _countRemainingWeekdays_(endDateStr) {
  const kst = new Date(Date.now() + 9 * 3600 * 1000);
  const d = new Date(kst);
  d.setUTCDate(d.getUTCDate() + 1);
  d.setUTCHours(0, 0, 0, 0);
  const end = new Date(endDateStr + 'T23:59:59+09:00');
  let count = 0;
  while (d <= end) {
    const day = d.getDay();
    if (day >= 1 && day <= 5) count++;
    d.setDate(d.getDate() + 1);
  }
  return count;
}

// ── ASCII bar: filled = remaining / maxDays scaled to barWidth chars ──────────
function _makeBar_(remaining, maxDays, barWidth) {
  const filled = Math.round(Math.min(remaining, maxDays) / maxDays * barWidth);
  return '█'.repeat(filled) + '░'.repeat(barWidth - filled);
}

// ── Build & send combined trigger-countdown bar chart to Google Chat ──────────
function sendAllTriggerStatus() {
  const kst      = new Date(Date.now() + 9 * 3600 * 1000);
  const todayStr = Utilities.formatDate(kst, 'UTC', 'yyyy-MM-dd (EEE)');

  const BAR_MAX = 20; // 20 weekdays ≈ 4 weeks = full bar
  const BAR_W   = 15;

  const rows = TRIGGER_PROJECTS.map(p => {
    const rem  = _countRemainingWeekdays_(p.endDate);
    const bar  = _makeBar_(rem, BAR_MAX, BAR_W);
    const icon = rem === 0 ? '🔴' : rem <= 5 ? '⚠️ ' : rem <= 10 ? '🔔 ' : '✅ ';
    return { name: p.name, rem, bar, icon, endDate: p.endDate };
  });

  // Pad names to same width for alignment inside code block
  const padLen    = Math.max(...rows.map(r => r.name.length));
  const chartLines = rows.map(r => {
    const name = r.name.padEnd(padLen);
    const cnt  = String(r.rem).padStart(2);
    return `${r.icon} ${name}  ${r.bar}  ${cnt}d  (→ ${r.endDate})`;
  });

  const urgentRows = rows.filter(r => r.rem <= 5 && r.rem > 0);
  const urgentNote = urgentRows.length
    ? '\n⚠️ *' + urgentRows.map(r => `${r.name}: ${r.rem}d left`).join(', ') + ' — extend soon!*'
    : '';

  const text = [
    `*📊 Trigger Countdown — ${todayStr} KST*`,
    '```',
    ...chartLines,
    '```',
    urgentNote,
  ].filter(l => l !== '').join('\n');

  const resp = UrlFetchApp.fetch(STATUS_WEBHOOK, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({ text }),
    muteHttpExceptions: true,
  });
  Logger.log('Trigger status sent (HTTP %s)', resp.getResponseCode());
  rows.forEach(r => Logger.log('  %s: %s days remaining', r.name, r.rem));
}

// ── Recreate all masterDailyJob triggers up to MASTER_END_DATE ────────────────
function createTriggers() {
  const endDate = new Date(MASTER_END_DATE + 'T23:59:59+09:00');
  let date = new Date();
  date.setHours(4, 0, 0, 0); // 06:00 AM KST (UTC+9 → 04:00 UTC)

  // Delete ALL existing triggers in this project
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  let created = 0;
  const now = new Date();

  while (date <= endDate) {
    const day = date.getDay(); // 0 = Sun, 6 = Sat
    if (day >= 1 && day <= 5 && date > now) {
      ScriptApp.newTrigger('masterDailyJob')
        .timeBased()
        .at(new Date(date))
        .create();
      Logger.log('Trigger set for: %s', date);
      created++;
    }
    date.setDate(date.getDate() + 1);
    date.setHours(4, 0, 0, 0);
  }

  Logger.log('Created %s weekday 06:00 KST triggers until %s', created, MASTER_END_DATE);
}

// ── Daily entry point ─────────────────────────────────────────────────────────
function masterDailyJob() {
  runAllScrapers();
  sendAllTriggerStatus();
}
