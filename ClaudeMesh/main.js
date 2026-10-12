const { app, BrowserWindow, ipcMain, dialog, Notification, shell, clipboard } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { execFile } = require('child_process');
const pty = require('node-pty');
const { Collector } = require('./collector');
const { EventEmitter } = require('events');
const { createPhoneServer } = require('./phone-server');
const { createCore, cleanEnv, loginShell } = require('./core');
const ptyBus = new EventEmitter(); ptyBus.setMaxListeners(50);   // pty output fan-out (desktop + phone terminals)
let lastSnap = null, phone = null;

const HOME = os.homedir();
// renamed Claude Mesh → GCX Mesh: keep using the original data folder so pairing, caches and saved orbits survive
app.setPath('userData', process.env.MESH_USERDATA || path.join(app.getPath('appData'), 'Claude Mesh'));   // MESH_USERDATA: isolated test runs
const collector = new Collector();
const ptys = new Map();          // id -> { proc, title, kind, cwd }
let win = null;
let nextId = 1;

function createWindow() {
  win = new BrowserWindow({
    width: 1680, height: 1020, minWidth: 340, minHeight: 240,
    backgroundColor: '#05060a', title: 'GCX Mesh',
    titleBarStyle: 'hiddenInset', trafficLightPosition: { x: 14, y: 14 },
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, backgroundThrottling: true },
  });
  if (process.env.MESH_SIZE) { const [w, h] = process.env.MESH_SIZE.split('x').map(Number); win.setSize(w, h); }   // test runs
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.on('closed', () => { win = null; });
  // tell the renderer when nobody can see the window so it stops drawing (telemetry keeps flowing)
  const vis = v => () => win && win.webContents.send('win:visible', v);
  win.on('hide', vis(false)); win.on('minimize', vis(false)); win.on('show', vis(true)); win.on('restore', vis(true));
  app.on('hide', vis(false)); app.on('show', vis(true));
  if (process.env.MESH_DEBUG) win.webContents.on('console-message', (e, ...a) => console.log('[renderer]', JSON.stringify(e.message ?? a)));
}

// ---------------- terminals (node-pty) ----------------
// cleanEnv / loginShell come from core.js (shared with the server program)

ipcMain.handle('pty:spawn', (_e, opts) => {
  const id = 't' + nextId++;
  const cwd = opts.cwd && fs.existsSync(opts.cwd) ? opts.cwd : HOME;
  const sh = loginShell();
  let args = ['-l'];
  let title = opts.title || 'shell';
  if (opts.kind === 'claude') {
    let flags = (opts.flags || '').trim();
    // every Claude session started from GCX Mesh runs with bypass-permissions + Remote Control
    // (not for subcommands like `attach <id>`, which take no session flags)
    if (!/^(attach|agents|logs|stop|rm|respawn)\b/.test(flags)) {
      if (!/--dangerously-skip-permissions|--permission-mode/.test(flags)) flags = '--dangerously-skip-permissions ' + flags;
      if (!/--remote-control\b|(^|\s)--rc\b/.test(flags)) flags = '--remote-control ' + flags;
      flags = flags.trim();
    }
    // exec so the claude pid == pty pid; -i loads the user's PATH (~/.local/bin) like Terminal does
    args = ['-l', '-i', '-c', `exec claude ${flags}`];
    title = opts.title || 'claude';
  }
  const proc = pty.spawn(sh, args, {
    name: 'xterm-256color', cols: opts.cols || 120, rows: opts.rows || 32, cwd,
    env: { ...cleanEnv(), TERM: 'xterm-256color', COLORTERM: 'truecolor', TERM_PROGRAM: 'ClaudeMesh' },
  });
  const entry = { proc, title, kind: opts.kind, cwd, buf: '' };
  ptys.set(id, entry);
  proc.onData(d => {
    win && win.webContents.send('pty:data', id, d);
    entry.buf = (entry.buf + d).slice(-1500000);                 // backlog for a phone that opens this terminal later
    ptyBus.emit('data', id, d);
  });
  proc.onExit(({ exitCode }) => {
    ptys.delete(id);
    win && win.webContents.send('pty:exit', id, exitCode);
    ptyBus.emit('exit', id);
  });
  return { id, pid: proc.pid, title, cwd };
});
ipcMain.on('pty:write', (_e, id, data) => { const p = ptys.get(id); if (p) p.proc.write(data); });
ipcMain.on('pty:resize', (_e, id, c, r) => { const p = ptys.get(id); if (p && c > 0 && r > 0) try { p.proc.resize(c, r); } catch { } });
ipcMain.on('pty:kill', (_e, id) => { const p = ptys.get(id); if (p) try { p.proc.kill(); } catch { } });

// ---------------- control of sessions running in Terminal.app ----------------
function osa(lines) {
  return new Promise(res => execFile('/usr/bin/osascript', lines.flatMap(l => ['-e', l]),
    (err, so, se) => res({ ok: !err, out: (so || '').trim(), err: (se || '').trim() })));
}
const asStr = s => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';

function terminalTabScript(tty, body) {
  return ['tell application "Terminal"',
    'repeat with w in windows', 'repeat with t in tabs of w',
    `if tty of t is ${asStr(tty)} then`, ...body, 'return "ok"', 'end if',
    'end repeat', 'end repeat', 'return "notfound"', 'end tell'];
}

// Typing into the external session: Terminal "do script" types the text then CR (= Enter).
ipcMain.handle('ext:send', async (_e, tty, text) => {
  const clean = String(text).replace(/\r?\n/g, ' ');
  return osa(terminalTabScript(tty, [`do script ${asStr(clean)} in t`]));
});
ipcMain.handle('ext:interrupt', async (_e, tty) =>
  osa(terminalTabScript(tty, ['do script (ASCII character 27) in t'])));
ipcMain.handle('ext:focus', async (_e, tty) =>
  osa(terminalTabScript(tty, ['set selected of t to true', 'set index of w to 1', 'activate'])));

// Write to an in-app session (owned pty) by id
function ctlSend(ptyId, text, submit) {
  const p = ptys.get(ptyId); if (!p) return { ok: false };
  p.proc.write(String(text));
  if (submit) setTimeout(() => p.proc.write('\r'), 60);
  return { ok: true };
}
ipcMain.handle('ctl:send', (_e, ptyId, text, submit) => ctlSend(ptyId, text, submit));

// by session id (used by the phone): in-app → pty, Terminal.app → AppleScript
const liveSession = sid => lastSnap && lastSnap.sessions.find(s => s.sid === sid);
async function sendToSession(sid, text) {
  const s = liveSession(sid); if (!s) return { ok: false, err: 'session is not running' };
  if (s.owner) return ctlSend(s.owner, text, true);
  if (s.tty && s.kind !== 'bg') return osa(terminalTabScript(s.tty, [`do script ${asStr(String(text).replace(/\r?\n/g, ' '))} in t`]));
  return { ok: false, err: 'background session — view only' };
}
async function interruptSession(sid) {
  const s = liveSession(sid); if (!s) return { ok: false };
  if (s.owner) { const p = ptys.get(s.owner); if (p) p.proc.write('\x1b'); return { ok: true }; }
  if (s.tty) return osa(terminalTabScript(s.tty, ['do script (ASCII character 27) in t']));
  return { ok: false };
}

ipcMain.handle('dlg:dir', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'], defaultPath: HOME });
  return r.canceled ? null : r.filePaths[0];
});
ipcMain.handle('notify', (_e, title, body) => { new Notification({ title, body, silent: false }).show(); });
ipcMain.handle('open:path', (_e, p) => shell.openPath(p));
ipcMain.handle('home', () => HOME);

// ---------------- shared with the server program (core.js): commands, history, details, agents, usage ----------------
// Plan usage: the same request Claude Code's /usage makes, with the login Claude Code keeps in the keychain.
// The token is read per request and never stored or logged.
function keychainToken() {
  return new Promise(res => execFile('/usr/bin/security', ['find-generic-password', '-s', 'Claude Code-credentials', '-w'], { timeout: 10000 }, (err, so) => {
    try { res(err ? null : JSON.parse(so).claudeAiOauth.accessToken); } catch { res(null); }
  }));
}
const core = createCore({ userData: app.getPath('userData'), getToken: keychainToken });
const { commandsFor, sessionDetails, renameTranscript, convertToAgent, refreshAgents, refreshUsage } = core;
const scanHistory = () => core.listHistory();
ipcMain.handle('cmds:list', (_e, cwd) => commandsFor(cwd));
ipcMain.handle('hist:list', () => scanHistory());
ipcMain.handle('hist:rename', (_e, sid, name) => renameTranscript(sid, name));
ipcMain.handle('sess:details', (_e, sid) => sessionDetails(sid));
// stop a running session (one outside the app — in-app terminals are closed by the renderer)
ipcMain.handle('session:stop', async (_e, pid) => {
  pid = +pid; if (!pid || pid === process.pid) return { ok: false, err: 'bad pid' };
  try { process.kill(pid, 'SIGTERM'); } catch (e) { return { ok: e.code === 'ESRCH', err: e.message }; }
  for (let i = 0; i < 30; i++) { await new Promise(r => setTimeout(r, 100)); try { process.kill(pid, 0); } catch { return { ok: true }; } }
  try { process.kill(pid, 'SIGKILL'); } catch { } return { ok: true };
});

// ---------------- schedules (Mac launchd + server cron) — schedules.js ----------------
const schedules = require('./schedules');
const syncToServer = () => new Promise(res => { const f = path.join(HOME, '.claude', 'skills', 'run-on-server', 'sync_to_server.sh');
  if (!fs.existsSync(f)) return res(); execFile('/bin/bash', [f], { timeout: 180000 }, () => res()); });
ipcMain.handle('sched:list', () => schedules.list().catch(e => ({ error: String(e.message || e) })));
ipcMain.handle('sched:set', (_e, id, s) => schedules.setSchedule(id, s));
ipcMain.handle('sched:enable', (_e, id, on) => schedules.setEnabled(id, on));
ipcMain.handle('sched:move', (_e, id, to) => schedules.move(id, to, syncToServer));
ipcMain.handle('sched:run', (_e, id) => schedules.runNow(id));
ipcMain.handle('sched:infra', () => schedules.infra().catch(e => ({ reachable: false, items: [], err: String(e) })));
ipcMain.handle('sched:infraEnable', (_e, id, on) => schedules.setInfraEnabled(id, on));
// Apps Script triggers (gas.js): live list via the app's own Google session; editing in Google's editor
const gas = require('./gas');
app.whenReady().then(() => gas.init(app.getPath('userData')));
ipcMain.handle('gas:cached', () => gas.cached());
ipcMain.handle('gas:refresh', () => gas.refresh());
ipcMain.handle('gas:edit', (_e, scriptId) => gas.edit(scriptId));
ipcMain.handle('gas:open', (_e, url) => gas.openUrl(url));
const gasConstants = require('./gasConstants');
ipcMain.handle('gasDue:list', () => gasConstants.list());
ipcMain.handle('gasDue:set', (_e, id, i, iso) => gasConstants.set(id, i, iso));

// ---------------- background agents (`claude agents`, `claude --bg`) — core.js ----------------
ipcMain.handle('agent:convert', (_e, sid, cwd, task) => convertToAgent(sid, cwd, task));
ipcMain.handle('agent:stop', async (_e, id) => { const r = await core.claudeCli(['stop', id]); refreshAgents(true); return r; });
ipcMain.handle('agent:logs', (_e, id) => core.claudeCli(['logs', id]));
ipcMain.handle('agents:list', () => refreshAgents(true));

if (process.env.MESH_DEBUG) ipcMain.handle('debug:shot', async (_e, f) => { const img = await win.webContents.capturePage(); fs.writeFileSync(f, img.toPNG()); return true; });

// ---------------- clipboard: files copied in Finder, images (screenshots), text ----------------
const PASTE_DIR = path.join(os.tmpdir(), 'claude-mesh-paste');
const unxml = v => v.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
// Electron 44 clipboard is async (W3C-style ClipboardItem). Finder copies expose
// NSFilenamesPboardType (all files) and text/uri-list; screenshots expose image/png.
ipcMain.handle('clip:read', async () => {
  let items = [];
  try { items = await clipboard.read(); } catch { }
  for (const it of items) {
    const t = it.types;
    const fnames = t.find(x => x.includes('NSFilenamesPboardType'));
    if (fnames) {
      const xml = await (await it.getType(fnames)).text();
      const paths = [...xml.matchAll(/<string>([^<]+)<\/string>/g)].map(m => unxml(m[1]));
      if (paths.length) return { kind: 'files', paths };
    }
    if (t.includes('text/uri-list')) {
      const list = (await (await it.getType('text/uri-list')).text()).split(/\r?\n/).filter(u => u.startsWith('file://'));
      if (list.length) return { kind: 'files', paths: list.map(u => decodeURIComponent(new URL(u).pathname)) };
    }
    const imgType = t.find(x => /^image\/(png|jpeg|gif|webp|tiff)$/.test(x));
    if (imgType) {
      const buf = Buffer.from(await (await it.getType(imgType)).arrayBuffer());
      fs.mkdirSync(PASTE_DIR, { recursive: true });
      const f = path.join(PASTE_DIR, `pasted-${new Date().toISOString().replace(/[:.]/g, '-')}.${imgType.split('/')[1].replace('jpeg', 'jpg')}`);
      fs.writeFileSync(f, buf);
      return { kind: 'files', paths: [f], image: true };
    }
  }
  let text = ''; try { text = await clipboard.readText(); } catch { }
  return { kind: 'text', text };
});

// ---------------- plan usage limits (what /usage shows) — core.js ----------------
ipcMain.handle('usage:refresh', () => refreshUsage(true));

// ---------------- phone access (phone-server.js) ----------------
function initPhone() {
  phone = createPhoneServer({
    root: __dirname, userData: app.getPath('userData'), ptys, ptyBus,
    getSnap: () => lastSnap,
    getHistory: () => scanHistory(),
    details: sid => sessionDetails(sid),
    commandsFor,
    sendToSession, interruptSession,
    resumeOnMac: sid => { if (!win) return { ok: false }; win.webContents.send('phone:resume', sid); return { ok: true }; },
    rename: async (sid, name) => {
      const s = liveSession(sid);
      if (s && (s.owner || (s.tty && s.kind !== 'bg'))) { const r = await sendToSession(sid, '/rename ' + name); if (r && r.ok !== false && r.out !== 'notfound') return { ok: true }; }
      return renameTranscript(sid, name);
    },
    convert: async (sid, task) => {
      const s = liveSession(sid), h = core.histEntry(sid);
      if (s && s.owner) { const p = ptys.get(s.owner); if (p) try { p.proc.kill(); } catch { } await new Promise(r => setTimeout(r, 1500)); }
      return convertToAgent(sid, (s && s.cwd) || (h && h.cwd) || HOME, task);
    },
  });
}
ipcMain.handle('phone:status', () => phone.status());
ipcMain.handle('phone:set', (_e, v) => { phone.setEnabled(v); return phone.status(); });
ipcMain.handle('phone:regen', () => { phone.regenerate(); return phone.status(); });
ipcMain.handle('phone:qr', (_e, url) => require('qrcode').toDataURL(url, { margin: 1, width: 360, color: { dark: '#0a0826', light: '#ffffff' } }));

// ---------------- reopen last sessions ----------------
// While running, the in-app Claude sessions are recorded to userData/last-sessions.json. On the next
// launch the window asks whether to reopen them; recording resumes only after that answer, so a fresh
// (empty) app never overwrites the list before you've been asked.
const LAST_FILE = () => path.join(app.getPath('userData'), 'last-sessions.json');
let restorePending = null, restoreAnswered = false, quitting = false, lastSaved = '';
function readLastSessions() { try { const d = JSON.parse(fs.readFileSync(LAST_FILE(), 'utf8')); return Array.isArray(d.sessions) ? d : null; } catch { return null; } }
function recordSessions(snap) {
  if (quitting || !restoreAnswered) return;
  const list = snap.sessions.filter(s => s.owner).map(s => ({ sid: s.sid, cwd: s.cwd, name: s.name, model: s.model || '' }));
  const key = JSON.stringify(list); if (key === lastSaved) return; lastSaved = key;
  try { fs.writeFileSync(LAST_FILE(), JSON.stringify({ at: Date.now(), sessions: list })); } catch { }
}
ipcMain.handle('restore:get', () => { if (restorePending === null) restorePending = readLastSessions() || { sessions: [] }; return restorePending; });
ipcMain.handle('restore:done', () => { restoreAnswered = true; return true; });
app.on('before-quit', () => { quitting = true; });

// ---------------- telemetry loop ----------------
async function loop() {
  try {
    await collector.tick();
    if (win) {
      const ptyMap = new Map([...ptys].map(([id, p]) => [id, p.proc.pid]));
      const snap = collector.snapshot(ptyMap);
      snap.agents = core.agents(); refreshAgents();
      snap.usage = core.usage(); refreshUsage();
      snap.ptys = [...ptys].map(([id, p]) => ({ id, pid: p.proc.pid, title: p.title, kind: p.kind, cwd: p.cwd }));
      win.webContents.send('telemetry', snap);
      lastSnap = snap; if (phone) phone.broadcast(snap); recordSessions(snap);
    }
  } catch (e) { console.error(e); }
  setTimeout(loop, 1000);
}

app.whenReady().then(() => {
  app.setName('GCX Mesh');
  createWindow();
  initPhone();
  loop();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', () => {
  quitting = true;                                   // keep last-sessions.json as it was at quit
  for (const p of ptys.values()) try { p.proc.kill(); } catch { }
  app.quit();
});

// Debug: MESH_SHOT=/path.png captures the window after it settles.
if (process.env.MESH_EVAL) app.whenReady().then(() => setTimeout(() => win && win.webContents.executeJavaScript(process.env.MESH_EVAL).catch(e => console.log('eval err', e)), 3000));
if (process.env.MESH_SHOT) app.whenReady().then(() => setTimeout(async () => {
  if (!win) return;
  const img = await win.webContents.capturePage();
  fs.writeFileSync(process.env.MESH_SHOT, img.toPNG());
  if (process.env.MESH_SHOT_QUIT) app.quit();
}, +(process.env.MESH_SHOT_DELAY || 7000)));
