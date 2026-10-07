// Apps Script triggers for the Schedules menu. There's no API for another project's triggers with
// the scopes we have, so GCX Mesh keeps its own signed-in Google session (partition persist:google)
// and reads Google's "My Triggers" page (script.google.com/home/triggers) — every trigger across all
// your projects, with last run, event, function and error rate. Editing opens that project's own
// Triggers page in an app window (Google's editor); due-date constants in code are edited by
// gasConstants.js (file edit → clasp push → git).
const { BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

const PART = 'persist:google';
const DASH = 'https://script.google.com/home/triggers';
let cachePath = null, scraping = null;

// runs inside the My Triggers page: walk every page of the table
const SCRAPE = `(async () => {
  const out = [], seen = new Set(), wait = ms => new Promise(r => setTimeout(r, ms));
  for (let i = 0; i < 40 && !document.querySelector('tr[data-script-id]'); i++) await wait(250);
  for (let page = 0; page < 40; page++) {
    const rows = [...document.querySelectorAll('tr[data-script-id]')];
    for (const r of rows) {
      const c = [...r.querySelectorAll('td')].map(td => td.innerText.replace(/\\s+/g, ' ').trim());
      const key = (r.getAttribute('jsdata') || '') + '|' + c.join('|'); if (seen.has(key)) continue; seen.add(key);
      out.push({ scriptId: r.getAttribute('data-script-id'), project: c[0], last: c[1], deployment: c[2], event: c[3], fn: c[4], errorRate: c[5] });
    }
    const nx = document.querySelector('[aria-label="Go to next page"]');
    if (!nx || nx.getAttribute('aria-disabled') === 'true' || nx.disabled) break;
    const first = rows[0] && rows[0].innerText; nx.click();
    for (let i = 0; i < 25; i++) { await wait(300); const f = document.querySelector('tr[data-script-id]'); if (f && f.innerText !== first) break; }
  }
  return out;
})()`;

function init(userData) { cachePath = path.join(userData, 'gas-triggers.json'); }
function cached() { try { return JSON.parse(fs.readFileSync(cachePath, 'utf8')); } catch { return null; } }

function win(show, url) {
  const w = new BrowserWindow({ width: 1180, height: 820, show, title: 'Apps Script — GCX Mesh', webPreferences: { partition: PART, contextIsolation: true } });
  w.loadURL(url); return w;
}
const signedOut = u => /accounts\.google\.com|ServiceLogin|signin/i.test(u);

// read the live list (hidden window). If Google asks to sign in, show the window and wait for it.
function refresh() {
  if (scraping) return scraping;
  scraping = new Promise(resolve => {
    const w = win(false, DASH); let done = false;
    const finish = r => { if (done) return; done = true; try { w.destroy(); } catch { } scraping = null; resolve(r); };
    const tryScrape = async () => {
      const u = w.webContents.getURL();
      if (signedOut(u)) { w.setTitle('Sign in to Google (kjw@spigen.com) — GCX Mesh reads your Apps Script triggers'); w.show(); return; }
      if (!/script\.google\.com\/home\/triggers/.test(u)) { w.loadURL(DASH); return; }
      try {
        const list = await w.webContents.executeJavaScript(SCRAPE);
        const r = { ok: true, at: Date.now(), triggers: list };
        try { fs.writeFileSync(cachePath, JSON.stringify(r)); } catch { }
        finish(r);
      } catch (e) { finish({ ok: false, err: String(e.message || e), ...(cached() || {}) }); }
    };
    w.webContents.on('did-finish-load', () => setTimeout(tryScrape, 900));
    w.on('closed', () => finish({ ok: false, err: 'sign-in window closed', ...(cached() || {}) }));
    setTimeout(() => finish({ ok: false, err: 'timed out reading triggers', ...(cached() || {}) }), 180000);
  });
  return scraping;
}

// Google's own trigger editor for one project, inside the app; re-read the list when it closes
function edit(scriptId) {
  return new Promise(resolve => {
    const w = win(true, `https://script.google.com/home/projects/${encodeURIComponent(scriptId)}/triggers`);
    w.on('closed', () => resolve(true));
  });
}
function openUrl(url) { win(true, url); return true; }

module.exports = { init, cached, refresh, edit, openUrl };
