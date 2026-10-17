const { app, BrowserWindow, ipcMain, dialog, Notification, shell, clipboard, screen } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { execFile, execFileSync } = require('child_process');
const pty = require('node-pty');
const { Collector } = require('./collector');
const { EventEmitter } = require('events');
const { createPhoneServer } = require('./phone-server');
const { createPeer } = require('./peer');             // the GCX server: one mesh on the iPhone whichever machine serves it
const { createCore, cleanEnv, loginShell } = require('./core');
const ptyBus = new EventEmitter(); ptyBus.setMaxListeners(50);   // pty output fan-out (desktop + phone terminals)
const panes = require('./tmux-panes')({ ptyBus });           // Claude sessions in tmux panes (the mini keeps them there)
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
  else keepOnScreen(loadBounds());
  win.on('moved', saveBounds); win.on('resized', saveBounds);
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.on('closed', () => { win = null; });
  // tell the renderer when nobody can see the window so it stops drawing (telemetry keeps flowing)
  const vis = v => () => win && win.webContents.send('win:visible', v);
  win.on('hide', vis(false)); win.on('minimize', vis(false)); win.on('show', vis(true)); win.on('restore', vis(true));
  app.on('hide', vis(false)); app.on('show', vis(true));
  if (process.env.MESH_DEBUG) win.webContents.on('console-message', (e, ...a) => console.log('[renderer]', JSON.stringify(e.message ?? a)));
}

// the window comes back where it was, and never hangs off the edge of the screen — also after a resolution or
// display change (e.g. the mini's screen switched to 2560×1440 while the window was bigger or further right)
const BOUNDS_F = () => path.join(app.getPath('userData'), 'window-bounds.json');
function loadBounds() { try { return JSON.parse(fs.readFileSync(BOUNDS_F(), 'utf8')); } catch { return null; } }
function saveBounds() { if (win && !win.isFullScreen() && !win.isMaximized()) try { fs.writeFileSync(BOUNDS_F(), JSON.stringify(win.getBounds())); } catch { } }
function keepOnScreen(want) {
  if (!win || win.isFullScreen()) return;
  const b = want || win.getBounds(), wa = screen.getDisplayMatching(b).workArea;
  const width = Math.min(b.width, wa.width), height = Math.min(b.height, wa.height);
  const x = Math.min(Math.max(b.x, wa.x), wa.x + wa.width - width), y = Math.min(Math.max(b.y, wa.y), wa.y + wa.height - height);
  const cur = win.getBounds();
  if (cur.x !== x || cur.y !== y || cur.width !== width || cur.height !== height) win.setBounds({ x, y, width, height });
}
app.whenReady().then(() => {
  const fix = () => setTimeout(() => keepOnScreen(), 300);
  screen.on('display-metrics-changed', fix); screen.on('display-removed', fix); screen.on('display-added', fix);
});

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
ipcMain.on('pty:write', (_e, id, data) => {
  if (peerLink && peerLink.ownsTerm(id)) return void peerLink.termWrite(id, data).catch(() => { });   // a mirrored terminal of the peer
  if (panes.has(id)) return panes.write(id, data);                                                     // a tmux pane on this machine
  const p = ptys.get(id); if (p) p.proc.write(data);
});
// a clean copy of an in-app terminal's screen, serialized by the window's xterm (the raw backlog replays
// redraws made at other sizes and garbles); null if the window can't answer in time
const snapWaits = new Map(); let snapSeq = 0;
ipcMain.on('term:snap:reply', (_e, rq, data) => { const w = snapWaits.get(rq); if (w) { snapWaits.delete(rq); w(data); } });
function termSnapshot(id) {
  if (panes.has(id)) return Promise.resolve(panes.snapshot(id));   // tmux knows its pane's screen itself
  if (!win) return Promise.resolve(null);
  return new Promise(res => {
    const rq = ++snapSeq; snapWaits.set(rq, res); win.webContents.send('term:snap', rq, id);
    setTimeout(() => { if (snapWaits.delete(rq)) res(null); }, 1500);
  });
}
// Mirrors of the peer machine's in-app terminals: its /api/term SSE stream re-emitted as pty:data / pty:exit
const peerTerms = new Map();   // id (peer:…) -> upstream response
ipcMain.on('peerterm:open', async (_e, id) => {
  if (panes.has(id)) return openPaneView(id);
  if (!peerLink || peerTerms.has(id)) return;
  peerTerms.set(id, null);
  let up;
  try { up = await peerLink.openTerm(id); } catch { peerTerms.delete(id); return; }
  if (!peerTerms.has(id)) return up.destroy();          // closed while connecting
  if (up.statusCode !== 200) { up.resume(); peerTerms.delete(id); return; }
  peerTerms.set(id, up);
  let buf = '';
  up.setEncoding('utf8');
  up.on('data', chunk => {
    buf += chunk; let i;
    while ((i = buf.indexOf('\n\n')) >= 0) {
      const block = buf.slice(0, i); buf = buf.slice(i + 2);
      const ev = (block.match(/^event: (.*)$/m) || [])[1], dm = block.match(/^data: (.*)$/m);
      if (!ev || !dm || !win) continue;
      let d; try { d = JSON.parse(dm[1]); } catch { continue; }
      if (ev === 'data') win.webContents.send('pty:data', id, d);
      else if (ev === 'init' || ev === 'size') win.webContents.send('peerterm:' + ev, id, d);
      else if (ev === 'exit') win.webContents.send('pty:exit', id, 0);
    }
  });
  const gone = () => { if (peerTerms.get(id) === up) { peerTerms.delete(id); if (win) win.webContents.send('peerterm:closed', id); } };
  up.on('end', gone); up.on('error', gone); up.on('close', gone);
});
// this window's view of one of our own tmux panes: the same events a peer mirror gets, straight from ptyBus
function openPaneView(id) {
  if (peerTerms.has(id) || !win) return;
  let held = [];
  const on = (pid, d) => { if (pid === id && win) held ? held.push(d) : win.webContents.send('pty:data', id, d); };
  const sz = (pid, cols, rows) => { if (pid === id && win) win.webContents.send('peerterm:size', id, { cols, rows }); };
  const off = pid => { if (pid === id && win) win.webContents.send('pty:exit', id, 0); };
  ptyBus.on('data', on); ptyBus.on('resize', sz); ptyBus.on('exit', off);
  const view = { destroy() { ptyBus.off('data', on); ptyBus.off('resize', sz); ptyBus.off('exit', off); panes.close(id); } };
  peerTerms.set(id, view);
  panes.open(id).then(() => {
    const p = panes.proc(id); if (!p || !win || peerTerms.get(id) !== view) return;
    win.webContents.send('peerterm:init', id, { cols: p.proc.cols, rows: p.proc.rows, backlog: panes.snapshot(id) });
    for (const d of held) win.webContents.send('pty:data', id, d); held = null;
  });
}
ipcMain.on('peerterm:close', (_e, id) => { const up = peerTerms.get(id); peerTerms.delete(id); if (up) up.destroy(); });
ipcMain.on('pty:resize', (_e, id, c, r) => {
  const p = ptys.get(id);
  if (p && c > 0 && r > 0 && (c !== p.proc.cols || r !== p.proc.rows)) try { p.proc.resize(c, r); ptyBus.emit('resize', id, c, r); } catch { }
});
// Whichever machine is being used sets the size of a shared terminal (like tmux's "latest" window-size): the
// peer's mirror asks for its pane size, the pty is resized so Claude Code rewraps, and our window scales its copy
// to that size until it is used here again (then it fits itself and takes the size back).
function termResize(id, c, r) {
  if (panes.has(id)) {
    if (!(c >= 20 && c <= 500 && r >= 5 && r <= 200) || !panes.resize(id, c, r)) return { ok: false };
    ptyBus.emit('resize', id, c, r); return { ok: true };
  }
  const p = ptys.get(id);
  if (!p || !(c >= 20 && c <= 500 && r >= 5 && r <= 200)) return { ok: false };
  if (c !== p.proc.cols || r !== p.proc.rows) try { p.proc.resize(c, r); } catch { return { ok: false }; }
  ptyBus.emit('resize', id, c, r);   // also when unchanged: the asking mirror waits for this as confirmation
  if (win) win.webContents.send('pty:remoteSize', id, c, r);
  return { ok: true };
}
ipcMain.on('peerterm:resize', (_e, id, c, r) => { if (panes.has(id)) return void termResize(id, c, r); if (peerLink) peerLink.termResize(id, c, r).catch(() => { }); });
ipcMain.on('pty:kill', (_e, id) => { const p = ptys.get(id); if (p) try { p.proc.kill(); } catch { } });

// ---------------- control of sessions running in Terminal.app ----------------
function osa(lines) {
  return new Promise(res => execFile('/usr/bin/osascript', lines.flatMap(l => ['-e', l]),
    (err, so, se) => res({ ok: !err, out: (so || '').trim(), err: (se || '').trim() })));
}
const asStr = s => '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';

// Sessions hosted in a tmux pane (the mini keeps every Claude session in tmux gcx:*) have a tty but no
// Terminal.app tab: type into the pane with send-keys instead. Looked up by tty on every call (panes come and go).
const TMUX = ['/opt/homebrew/bin/tmux', '/usr/local/bin/tmux'].find(p => fs.existsSync(p)) || 'tmux';
// launched from the Dock there is no locale, and tmux then mangles tabs and non-ASCII text (Korean prompts) into '_'
const TMUX_ENV = { ...process.env, LANG: 'en_US.UTF-8', LC_CTYPE: 'en_US.UTF-8' };
function tmuxPaneForTty(tty) {
  if (!tty) return null;
  try {
    for (const l of execFileSync(TMUX, ['list-panes', '-a', '-F', '#{pane_id}|#{pane_tty}'], { timeout: 5000, env: TMUX_ENV }).toString().split('\n')) {
      const [id, t] = l.split('|'); if (t === tty) return id;
    }
  } catch { }
  return null;
}
function tmuxKeys(pane, ...args) {
  try { execFileSync(TMUX, ['send-keys', '-t', pane, ...args], { timeout: 5000, env: TMUX_ENV }); return { ok: true, out: 'tmux' }; }
  catch (e) { return { ok: false, err: String(e.message || e) }; }
}
// External session by tty: tmux pane if it has one, else the Terminal.app tab (osascript 'notfound' when neither).
function extSend(tty, text) {
  const clean = String(text).replace(/\r?\n/g, ' ');
  const pane = tmuxPaneForTty(tty);
  if (!pane) return osa(terminalTabScript(tty, [`do script ${asStr(clean)} in t`]));
  const r = tmuxKeys(pane, '-l', clean);
  if (r.ok) setTimeout(() => tmuxKeys(pane, 'Enter'), 60);          // Enter after the text settles, like ctlSend
  return r;
}
function extInterrupt(tty) {
  const pane = tmuxPaneForTty(tty);
  return pane ? tmuxKeys(pane, 'Escape') : osa(terminalTabScript(tty, ['do script (ASCII character 27) in t']));
}
function extFocus(tty) {
  if (tmuxPaneForTty(tty)) return { ok: true, out: 'tmux' };          // nothing to bring forward for a tmux pane
  return osa(terminalTabScript(tty, ['set selected of t to true', 'set index of w to 1', 'activate']));
}

function terminalTabScript(tty, body) {
  return ['tell application "Terminal"',
    'repeat with w in windows', 'repeat with t in tabs of w',
    `if tty of t is ${asStr(tty)} then`, ...body, 'return "ok"', 'end if',
    'end repeat', 'end repeat', 'return "notfound"', 'end tell'];
}

// Typing into the external session: Terminal "do script" types the text then CR (= Enter).
ipcMain.handle('ext:send', async (_e, tty, text) => extSend(tty, text));
ipcMain.handle('ext:interrupt', async (_e, tty) => extInterrupt(tty));
ipcMain.handle('ext:focus', async (_e, tty) => extFocus(tty));

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
  if (s.tty && s.kind !== 'bg') return extSend(s.tty, text);
  return { ok: false, err: 'background session — view only' };
}
async function interruptSession(sid) {
  const s = liveSession(sid); if (!s) return { ok: false };
  if (s.owner) { const p = ptys.get(s.owner); if (p) p.proc.write('\x1b'); return { ok: true }; }
  if (s.tty) return extInterrupt(s.tty);
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
// the history list shows both machines: our transcripts, plus the peer's that haven't been synced here yet
// (tagged with its name; resuming one starts it over there, where its transcript is)
ipcMain.handle('hist:list', async () => {
  const local = await scanHistory();
  if (!peerLink || !peerLink.online()) return local;
  const theirs = await peerLink.forward('GET', '/api/history').catch(() => null);
  if (!Array.isArray(theirs)) return local;
  const mine = new Set(local.map(h => h.sid)), label = peerLink.label();
  const extra = theirs.filter(h => h && h.sid && !mine.has(h.sid)).map(h => ({ ...h, file: null, onPeer: true, hostLabel: label }));
  return [...local, ...extra].sort((a, b) => b.mtime - a.mtime);
});
ipcMain.handle('peer:resume', (_e, sid) => peerLink ? peerLink.forward('POST', '/api/resume', { sid }).catch(e => ({ ok: false, err: String(e.message || e) })) : { ok: false, err: 'no peer' });
ipcMain.handle('hist:rename', (_e, sid, name) => renameTranscript(sid, name));
ipcMain.handle('sess:details', (_e, sid) => !liveSession(sid) && peerLink && peerLink.owns(sid)
  ? peerLink.forward('GET', '/api/details', undefined, { sid }).catch(e => ({ ok: false, err: String(e.message || e) })) : sessionDetails(sid));
// a session running on the peer machine: send / interrupt go through its phone server
// files pasted or dropped into a mirror of the peer's terminal: copy them over, hand back the peer's paths
// (a path on this machine means nothing to a session running on the other one). Folders/huge files keep their path.
ipcMain.handle('peer:upload', async (_e, paths) => {
  const out = []; let err = null;
  for (const f of paths || []) {
    try {
      const st = fs.statSync(f);
      if (!peerLink || !st.isFile() || st.size > 25e6) { out.push(f); continue; }
      const r = await peerLink.upload(path.basename(f), fs.readFileSync(f).toString('base64'));
      if (r && r.path) out.push(r.path); else { err = 'not supported there — update/restart GCX Mesh on ' + peerLink.label(); break; }
    } catch (e) { err = String(e.message || e); break; }
  }
  return err ? { err } : { paths: out };
});
ipcMain.handle('peer:send', (_e, sid, text) => peerLink ? peerLink.forward('POST', '/api/send', { sid, text }).catch(e => ({ ok: false, err: String(e.message || e) })) : { ok: false, err: 'no peer' });
ipcMain.handle('peer:interrupt', (_e, sid) => peerLink ? peerLink.forward('POST', '/api/interrupt', { sid }).catch(e => ({ ok: false, err: String(e.message || e) })) : { ok: false, err: 'no peer' });
// stop a running session (one outside the app — in-app terminals are closed by the renderer)
ipcMain.handle('session:stop', async (_e, pid) => {
  pid = +pid; if (!pid || pid === process.pid) return { ok: false, err: 'bad pid' };
  if (!(lastSnap && lastSnap.sessions.some(s => s.pid === pid))) return { ok: false, err: 'not a session on this machine' };   // never a peer's pid
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
ipcMain.on('clip:write', (_e, text) => { if (typeof text === 'string' && text.length <= 5e6) clipboard.writeText(text); });
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
let peerLink = null;
function initPhone() {
  peerLink = createPeer({ userData: app.getPath('userData'), selfKind: 'mac' }); peerLink.start();
  phone = createPhoneServer({
    root: __dirname, userData: app.getPath('userData'), ptys: { get: id => ptys.get(id) || panes.proc(id) }, ptyBus, peer: peerLink, termSnapshot, termResize,
    termOpen: id => panes.has(id) ? panes.open(id) : null, termClose: id => { if (panes.has(id)) panes.close(id); },
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
app.on('before-quit', () => { quitting = true; panes.stopAll(); });

// ---------------- this machine's memory (what Activity Monitor calls "Memory Used") ----------------
// app memory (anonymous − purgeable) + wired + compressed, from vm_stat; pressure 1 normal / 2 warning / 4 critical
let hostMem = null;
function readMem() {
  require('child_process').execFile('/usr/bin/vm_stat', (err, out) => {
    if (err) return;
    const ps = +(out.match(/page size of (\d+)/) || [, 16384])[1];
    const pg = k => +((out.match(new RegExp(k + ':\\s+(\\d+)')) || [, 0])[1]);
    const used = (pg('Anonymous pages') - pg('Pages purgeable') + pg('Pages wired down') + pg('Pages occupied by compressor')) * ps;
    require('child_process').execFile('/usr/sbin/sysctl', ['-n', 'kern.memorystatus_vm_pressure_level'], (e2, lv) => {
      hostMem = { used, total: os.totalmem(), pressure: +String(lv || '1').trim() || 1 };
    });
  });
}
readMem(); setInterval(readMem, 5000);

// ---------------- telemetry loop ----------------
async function loop() {
  try {
    await collector.tick();
    if (win) {
      const ptyMap = new Map([...ptys].map(([id, p]) => [id, p.proc.pid]));
      const snap = collector.snapshot(ptyMap);
      snap.agents = core.agents(); refreshAgents();
      snap.usage = core.usage(); refreshUsage();
      snap.mem = hostMem;
      snap.ptys = [...ptys].map(([id, p]) => ({ id, pid: p.proc.pid, title: p.title, kind: p.kind, cwd: p.cwd }));
      panes.refresh();
      const byTty = new Map(panes.list().map(p => [p.tty, p]));
      for (const s of snap.sessions) {
        const p = !s.owner && s.tty && byTty.get(s.tty); if (!p) continue;
        s.pane = p.id;
        snap.ptys.push({ id: p.id, pid: p.pid, title: s.name, kind: 'tmux', cwd: s.cwd });
      }
      // the window shows BOTH machines (the peer's sessions and terminals merged in); phone + peer keep the local snap
      const shown = peerLink ? peerLink.merge(snap) : snap;   // includes the peer's terminals (peer:… ids, onPeer)
      win.webContents.send('telemetry', shown);
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
