const { app, BrowserWindow, ipcMain, dialog, Notification, shell, clipboard } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { execFile } = require('child_process');
const pty = require('node-pty');
const { Collector } = require('./collector');
const { EventEmitter } = require('events');
const { createPhoneServer } = require('./phone-server');
const ptyBus = new EventEmitter(); ptyBus.setMaxListeners(50);   // pty output fan-out (desktop + phone terminals)
let lastSnap = null, phone = null;

const HOME = os.homedir();
const collector = new Collector();
const ptys = new Map();          // id -> { proc, title, kind, cwd }
let win = null;
let nextId = 1;

function createWindow() {
  win = new BrowserWindow({
    width: 1680, height: 1020, minWidth: 1100, minHeight: 700,
    backgroundColor: '#05060a', title: 'Claude Mesh',
    titleBarStyle: 'hiddenInset', trafficLightPosition: { x: 14, y: 14 },
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, backgroundThrottling: true },
  });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.on('closed', () => { win = null; });
  // tell the renderer when nobody can see the window so it stops drawing (telemetry keeps flowing)
  const vis = v => () => win && win.webContents.send('win:visible', v);
  win.on('hide', vis(false)); win.on('minimize', vis(false)); win.on('show', vis(true)); win.on('restore', vis(true));
  app.on('hide', vis(false)); app.on('show', vis(true));
  if (process.env.MESH_DEBUG) win.webContents.on('console-message', (e, ...a) => console.log('[renderer]', JSON.stringify(e.message ?? a)));
}

// ---------------- terminals (node-pty) ----------------
function cleanEnv() {
  const e = { ...process.env };
  for (const k of Object.keys(e)) if (/^(CLAUDECODE|CLAUDE_CODE_|ELECTRON_|MESH_)/.test(k)) delete e[k];
  return e;
}
function loginShell() { return process.env.SHELL || '/bin/zsh'; }

ipcMain.handle('pty:spawn', (_e, opts) => {
  const id = 't' + nextId++;
  const cwd = opts.cwd && fs.existsSync(opts.cwd) ? opts.cwd : HOME;
  const sh = loginShell();
  let args = ['-l'];
  let title = opts.title || 'shell';
  if (opts.kind === 'claude') {
    const flags = (opts.flags || '').trim();
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
    entry.buf = (entry.buf + d).slice(-200000);                 // backlog for a phone that opens this terminal later
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

// ---------------- custom slash commands: ~/.claude/{commands,skills} (+ the session's project .claude) ----------------
function readDesc(file) {
  try {
    const head = fs.readFileSync(file, 'utf8').slice(0, 1500);
    const fm = /^---\n([\s\S]*?)\n---/.exec(head);
    if (fm) {
      const name = (/^name:\s*(.+)$/m.exec(fm[1]) || [])[1];
      let desc = (/^description:\s*(?:>-?|\|)?\s*\n?\s*([^\n]+)/m.exec(fm[1]) || [])[1] || '';
      return { name: name && name.trim().replace(/^["']|["']$/g, ''), desc: desc.trim().replace(/^["']|["']$/g, '') };
    }
    return { desc: (head.split('\n').find(l => l.trim() && !l.startsWith('#')) || '').trim() };
  } catch { return {}; }
}
function scanCommands(base, scope) {
  const out = [];
  try { for (const f of fs.readdirSync(path.join(base, 'commands'))) if (f.endsWith('.md')) { const d = readDesc(path.join(base, 'commands', f)); out.push({ cmd: '/' + f.slice(0, -3), desc: d.desc || 'Custom command', group: scope }); } } catch { }
  try {
    for (const e of fs.readdirSync(path.join(base, 'skills'), { withFileTypes: true })) {
      if (e.isDirectory()) { const f = path.join(base, 'skills', e.name, 'SKILL.md'); if (fs.existsSync(f)) { const d = readDesc(f); out.push({ cmd: '/' + (d.name || e.name), desc: d.desc || 'Skill', group: scope }); } }
      else if (e.name.endsWith('.md')) { const d = readDesc(path.join(base, 'skills', e.name)); out.push({ cmd: '/' + (d.name || e.name.slice(0, -3)), desc: d.desc || 'Skill', group: scope }); }
    }
  } catch { }
  return out;
}
ipcMain.handle('cmds:list', (_e, cwd) => commandsFor(cwd));
function commandsFor(cwd) {
  const list = scanCommands(path.join(HOME, '.claude'), 'Your skills & commands');
  if (cwd && cwd !== HOME) list.push(...scanCommands(path.join(cwd, '.claude'), 'Project'));
  const seen = new Set(); return list.filter(c => !seen.has(c.cmd) && seen.add(c.cmd));
}
// ---------------- session history (what `claude --resume` lists) ----------------
const PROJ_DIR = path.join(HOME, '.claude', 'projects');
const histCache = new Map();      // file -> { mtimeMs, size, meta }
const readSlice = (fd, pos, len) => { const b = Buffer.alloc(len); const n = fs.readSync(fd, b, 0, len, pos); return b.toString('utf8', 0, n); };
const promptText = m => {
  const c = m && m.content; let t = typeof c === 'string' ? c : Array.isArray(c) ? (c.find(x => x.type === 'text') || {}).text || '' : '';
  t = t.trim(); if (!t || /^<(command|local-command|system-reminder|bash-)/.test(t) || t.startsWith('Caveat:')) return '';
  return t.replace(/\s+/g, ' ').slice(0, 200);
};
function histMeta(file, st) {
  const meta = { sid: path.basename(file, '.jsonl'), file, mtime: st.mtimeMs, size: st.size, cwd: '', first: '', last: '', aiTitle: '', title: '', entry: '', branch: '' };
  let fd; try { fd = fs.openSync(file, 'r'); } catch { return null; }
  try {
    let head = '';
    for (let n = 262144; ; n *= 4) {          // grow the head until the first prompt + cwd are in it (big pastes)
      head = readSlice(fd, 0, Math.min(st.size, n));
      if (n >= st.size || n >= 16777216 || (/"cwd":/.test(head) && /"type":"user"[^\n]*"origin":\{"kind":"human"/.test(head))) break;
    }
    for (const line of head.split('\n')) {
      if (!line.includes('"type":"user"') && !line.includes('"cwd"')) continue;
      let j; try { j = JSON.parse(line); } catch { continue; }
      if (!meta.cwd && j.cwd) meta.cwd = j.cwd;
      if (!meta.entry && j.entrypoint) meta.entry = j.entrypoint;
      if (!meta.branch && j.gitBranch) meta.branch = j.gitBranch;
      if (!meta.first && j.type === 'user' && !j.isMeta && !j.isSidechain) meta.first = promptText(j.message);
      if (meta.cwd && meta.first && meta.entry) break;
    }
    const tl = Math.min(st.size, 524288), tail = readSlice(fd, st.size - tl, tl).split('\n');
    for (let i = tail.length - 1; i >= 0 && !(meta.last && meta.aiTitle); i--) {
      const l = tail[i];
      if (!meta.last && l.includes('"type":"last-prompt"')) try { meta.last = (JSON.parse(l).lastPrompt || '').replace(/\s+/g, ' ').slice(0, 200); } catch { }
      if (!meta.aiTitle && l.includes('"type":"ai-title"')) try { meta.aiTitle = JSON.parse(l).aiTitle || ''; } catch { }
    }
  } finally { fs.closeSync(fd); }
  return meta.first || meta.last ? meta : null;
}
// titles can sit anywhere in a big transcript → read each changed file once and take the LAST
// custom-title / ai-title record (Buffer.lastIndexOf is far faster than grep on ~1GB of jsonl)
const T_CUSTOM = Buffer.from('{"type":"custom-title"'), T_AI = Buffer.from('{"type":"ai-title"');
async function scanTitles(file, meta, from = 0) {
  let buf;
  try {
    if (!from) buf = await fs.promises.readFile(file);
    else { const fh = await fs.promises.open(file, 'r'); try { buf = Buffer.alloc(meta.size - from); await fh.read(buf, 0, buf.length, from); } finally { await fh.close(); } }
  } catch { return; }
  const pick = tag => { const i = buf.lastIndexOf(tag); if (i < 0) return null; let e = buf.indexOf(10, i); if (e < 0) e = buf.length; try { return JSON.parse(buf.toString('utf8', i, e)); } catch { return null; } };
  const c = pick(T_CUSTOM), a = pick(T_AI);
  if (c && c.customTitle != null) meta.title = c.customTitle;
  if (a && a.aiTitle) meta.aiTitle = a.aiTitle;
}
const HIST_CACHE_FILE = () => path.join(app.getPath('userData'), 'history-cache.json');
let histLoaded = false, histScan = null;
ipcMain.handle('hist:list', () => histScan || (histScan = scanHistory().finally(() => { histScan = null; })));
async function scanHistory() {
  if (!histLoaded) {                 // warm start from the previous run's cache
    histLoaded = true;
    try { for (const [f, c] of Object.entries(JSON.parse(fs.readFileSync(HIST_CACHE_FILE(), 'utf8')))) histCache.set(f, c); } catch { }
  }
  const files = [];
  try { for (const d of fs.readdirSync(PROJ_DIR)) { const dir = path.join(PROJ_DIR, d); try { for (const f of fs.readdirSync(dir)) if (f.endsWith('.jsonl')) files.push(path.join(dir, f)); } catch { } } } catch { }
  let changed = 0;
  for (const f of files) {
    let st; try { st = fs.statSync(f); } catch { continue; }
    const c = histCache.get(f);
    if (c && c.mtimeMs === st.mtimeMs && c.size === st.size) continue;
    const meta = histMeta(f, st);
    // a growing transcript only needs its new bytes scanned; earlier titles carry over
    const grew = c && c.meta && meta && st.size > c.size;
    if (grew) { meta.title = c.meta.title; meta.aiTitle = meta.aiTitle || c.meta.aiTitle; }
    if (meta) await scanTitles(f, meta, grew ? c.size : 0);
    histCache.set(f, { mtimeMs: st.mtimeMs, size: st.size, meta }); changed++;
  }
  const live = new Set(files);
  for (const f of histCache.keys()) if (!live.has(f)) { histCache.delete(f); changed++; }
  if (changed) try { fs.writeFileSync(HIST_CACHE_FILE(), JSON.stringify(Object.fromEntries(histCache))); } catch { }
  return [...histCache.values()].map(c => c.meta).filter(Boolean).sort((a, b) => b.mtime - a.mtime);
}
// Rename = the same record `/rename` writes, appended to the transcript (shows in the /resume picker)
ipcMain.handle('hist:rename', (_e, sid, name) => renameTranscript(sid, name));
function renameTranscript(sid, name) {
  const c = [...histCache.values()].find(c => c.meta && c.meta.sid === sid);
  let file = c && c.meta.file;
  if (!file) try { for (const d of fs.readdirSync(PROJ_DIR)) { const f = path.join(PROJ_DIR, d, sid + '.jsonl'); if (fs.existsSync(f)) { file = f; break; } } } catch { }
  if (!file) return { ok: false, err: 'transcript not found' };
  fs.appendFileSync(file, JSON.stringify({ type: 'custom-title', customTitle: String(name), sessionId: sid }) + '\n');
  if (c) c.meta.title = String(name);
  return { ok: true };
}

// ---------------- session details (press-and-hold on a session bubble) ----------------
// Skills / slash commands / tools a session used, its latest compaction summary (its "context md"),
// first & last prompt, and the CLAUDE.md instructions it runs under. Parsed incrementally per file.
const detailCache = new Map();          // file -> { size, skills, cmds, tools, summary, lastText }
function findTranscript(sid) {
  const c = [...histCache.values()].find(c => c.meta && c.meta.sid === sid);
  if (c) return c.meta.file;
  try { for (const d of fs.readdirSync(PROJ_DIR)) { const f = path.join(PROJ_DIR, d, sid + '.jsonl'); if (fs.existsSync(f)) return f; } } catch { }
  return null;
}
const bump = (o, k) => { if (k) o[k] = (o[k] || 0) + 1; };
async function sessionDetails(sid) {
  const file = findTranscript(sid); if (!file) return { ok: false, err: 'transcript not found' };
  const st = fs.statSync(file);
  let d = detailCache.get(file);
  if (!d || st.size < d.size) d = { size: 0, skills: {}, cmds: {}, tools: {}, summary: '', summaryAt: '', lastText: '', rest: '' };
  if (st.size > d.size) {                                   // read only the new bytes
    const fh = await fs.promises.open(file, 'r');
    try {
      const CH = 8 << 20;
      for (let pos = d.size; pos < st.size; pos += CH) {
        const len = Math.min(CH, st.size - pos), b = Buffer.alloc(len); await fh.read(b, 0, len, pos);
        const lines = (d.rest + b.toString('utf8')).split('\n'); d.rest = lines.pop();
        for (const line of lines) {
          if (line.includes('"name":"Skill"')) for (const m of line.matchAll(/"name":"Skill","input":\{"skill":"([^"]+)"/g)) bump(d.skills, m[1]);
          if (line.includes('<command-name>')) for (const m of line.matchAll(/<command-name>\/?([^<]+)<\/command-name>/g)) bump(d.cmds, '/' + m[1].trim());
          if (line.includes('"type":"tool_use"')) for (const m of line.matchAll(/"type":"tool_use","id":"[^"]*","name":"([^"]+)"/g)) if (m[1] !== 'Skill') bump(d.tools, m[1]);
          if (line.includes('"isCompactSummary":true')) {
            try { const j = JSON.parse(line), c = j.message && j.message.content; const t = typeof c === 'string' ? c : Array.isArray(c) ? c.map(x => x.text || '').join('\n') : '';
              d.summary = t.replace(/^This session is being continued[^\n]*\n+/, '').replace(/^The summary below covers[^\n]*\n+/, '').trim(); d.summaryAt = j.timestamp || ''; } catch { }
          }
          if (line.includes('"type":"assistant"') && line.includes('"type":"text"')) {
            try { const j = JSON.parse(line), t = (j.message.content || []).filter(x => x.type === 'text').map(x => x.text).join('\n').trim(); if (t.length > 40) d.lastText = t; } catch { }
          }
        }
      }
    } finally { await fh.close(); }
    d.size = st.size; detailCache.set(file, d);
  }
  const meta = ([...histCache.values()].find(c => c.meta && c.meta.sid === sid) || {}).meta || histMeta(file, st) || {};
  const cwd = meta.cwd || '';
  const readMd = f => { try { const t = fs.readFileSync(f, 'utf8'); return { path: f.replace(HOME, '~'), text: t.length > 6000 ? t.slice(0, 6000) + '\n…' : t }; } catch { return null; } };
  const instructions = [cwd && cwd !== HOME && readMd(path.join(cwd, 'CLAUDE.md')), cwd && readMd(path.join(cwd, '.claude', 'CLAUDE.md')), readMd(path.join(HOME, 'CLAUDE.md')), readMd(path.join(HOME, '.claude', 'CLAUDE.md'))]
    .filter(Boolean).filter((v, i, a) => a.findIndex(x => x.path === v.path) === i);
  const top = (o, n) => Object.entries(o).sort((a, b) => b[1] - a[1]).slice(0, n);
  return { ok: true, sid, cwd, title: meta.title || meta.aiTitle || '', first: meta.first || '', last: meta.last || '',
    skills: top(d.skills, 30), cmds: top(d.cmds, 20), tools: top(d.tools, 12),
    summary: d.summary.slice(0, 12000), summaryAt: d.summaryAt, lastText: d.lastText.slice(0, 1500), instructions };
}
ipcMain.handle('sess:details', (_e, sid) => sessionDetails(sid));

// ---------------- schedules (Mac launchd + server cron) — schedules.js ----------------
const schedules = require('./schedules');
const syncToServer = () => new Promise(res => { const f = path.join(HOME, '.claude', 'skills', 'run-on-server', 'sync_to_server.sh');
  if (!fs.existsSync(f)) return res(); execFile('/bin/bash', [f], { timeout: 180000 }, () => res()); });
ipcMain.handle('sched:list', () => schedules.list().catch(e => ({ error: String(e.message || e) })));
ipcMain.handle('sched:set', (_e, id, s) => schedules.setSchedule(id, s));
ipcMain.handle('sched:enable', (_e, id, on) => schedules.setEnabled(id, on));
ipcMain.handle('sched:move', (_e, id, to) => schedules.move(id, to, syncToServer));
ipcMain.handle('sched:run', (_e, id) => schedules.runNow(id));

// ---------------- background agents (`claude agents`, `claude --bg`) ----------------
let claudeBin = null;            // resolved once through a login shell (PATH from the user's profile)
function resolveClaude() {
  if (claudeBin) return Promise.resolve(claudeBin);
  return new Promise(res => execFile(loginShell(), ['-l', '-i', '-c', 'command -v claude'], { env: cleanEnv(), timeout: 15000 }, (_e, so) => {
    const p = String(so || '').trim().split('\n').pop(); claudeBin = p && fs.existsSync(p) ? p : path.join(HOME, '.local', 'bin', 'claude'); res(claudeBin);
  }));
}
async function claudeCli(args, cwd) {
  const bin = await resolveClaude();
  return new Promise(res => execFile(bin, args, { cwd: cwd && fs.existsSync(cwd) ? cwd : HOME, env: { ...cleanEnv(), PATH: path.dirname(bin) + ':' + (process.env.PATH || '/usr/bin:/bin') }, timeout: 60000, maxBuffer: 16e6 },
    (err, so, se) => res({ ok: !err, out: String(so || '').trim(), err: String(se || (err && err.message) || '').trim() })));
}
let agentsCache = [], agentsAt = 0, agentsBusy = false;
async function refreshAgents(force) {
  if (agentsBusy || (!force && Date.now() - agentsAt < 8000)) return agentsCache;
  agentsBusy = true;
  try {
    const r = await claudeCli(['agents', '--json']);
    const i = r.out.indexOf('['); if (i >= 0) agentsCache = JSON.parse(r.out.slice(i)).filter(a => a.kind === 'background');
    agentsAt = Date.now();
  } catch { } finally { agentsBusy = false; }
  return agentsCache;
}
// Convert a session into a background agent: resume it under the same id with --bg
ipcMain.handle('agent:convert', (_e, sid, cwd, task) => convertToAgent(sid, cwd, task));
async function convertToAgent(sid, cwd, task) {
  const args = ['--bg', '--dangerously-skip-permissions', '--resume', sid];
  if (task && task.trim()) args.push(task.trim());
  const r = await claudeCli(args, cwd);
  refreshAgents(true);
  return r;
}
ipcMain.handle('agent:stop', async (_e, id) => { const r = await claudeCli(['stop', id]); refreshAgents(true); return r; });
ipcMain.handle('agent:logs', (_e, id) => claudeCli(['logs', id]));
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

// ---------------- plan usage limits (what /usage shows) ----------------
// Same request Claude Code's /usage makes, with the login Claude Code keeps in the keychain.
// The token is read per request and never stored or logged. Falls back to the last status-line payload.
let usage = null, usageAt = 0, usageBusy = false;
function keychainToken() {
  return new Promise(res => execFile('/usr/bin/security', ['find-generic-password', '-s', 'Claude Code-credentials', '-w'], { timeout: 10000 }, (err, so) => {
    try { res(err ? null : JSON.parse(so).claudeAiOauth.accessToken); } catch { res(null); }
  }));
}
async function refreshUsage(force) {
  if (usageBusy || (!force && Date.now() - usageAt < 180000)) return;
  usageBusy = true; usageAt = Date.now();
  try {
    const tok = await keychainToken();
    if (tok) {
      const r = await fetch('https://api.anthropic.com/api/oauth/usage', { headers: { Authorization: 'Bearer ' + tok, 'anthropic-beta': 'oauth-2025-04-20' }, signal: AbortSignal.timeout(15000) });
      if (r.ok) {
        const d = await r.json();
        const limits = (d.limits || []).map(l => ({
          label: l.kind === 'session' ? 'Session' : l.scope && l.scope.model ? l.scope.model.display_name : l.kind === 'weekly_all' ? 'All models' : (l.kind || '').replace(/_/g, ' '),
          kind: l.kind, group: l.group, used: +l.percent || 0, resetsAt: Date.parse(l.resets_at) || 0, severity: l.severity, active: !!l.is_active,
        }));
        if (limits.length) { usage = { limits, at: Date.now(), src: 'api', monthly: monthlyFrom(limits) }; return; }
      }
    }
    // fallback: rate_limits from the newest status-line payload (session + week only)
    const j = JSON.parse(fs.readFileSync(path.join(HOME, '.claude', 'state', 'mesh-statusline.json'), 'utf8')).rate_limits || {};
    const L = []; if (j.five_hour) L.push({ label: 'Session', kind: 'session', group: 'session', used: j.five_hour.used_percentage, resetsAt: j.five_hour.resets_at * 1000 });
    if (j.seven_day) L.push({ label: 'All models', kind: 'weekly_all', group: 'weekly', used: j.seven_day.used_percentage, resetsAt: j.seven_day.resets_at * 1000 });
    if (L.length) usage = { limits: L, at: Date.now(), src: 'statusline', monthly: monthlyFrom(L) };
  } catch { } finally { usageBusy = false; }
}
// There is no monthly plan limit, so "monthly" = this calendar month's share of the weekly
// allowance: each week's peak usage %, weighted by how many of its days fall in this month,
// over the number of weeks in the month. Week peaks are remembered across restarts.
function monthlyFrom(limits) {
  const wk = limits.find(l => l.kind === 'weekly_all'); if (!wk || !wk.resetsAt) return null;
  const f = path.join(app.getPath('userData'), 'weekly-usage.json');
  let hist = {}; try { hist = JSON.parse(fs.readFileSync(f, 'utf8')); } catch { }
  const key = String(Math.round(wk.resetsAt / 3600e3));            // one entry per weekly window
  hist[key] = Math.max(hist[key] || 0, wk.used);
  try { fs.writeFileSync(f, JSON.stringify(hist)); } catch { }
  const now = new Date(), m0 = new Date(now.getFullYear(), now.getMonth(), 1).getTime(), m1 = new Date(now.getFullYear(), now.getMonth() + 1, 1).getTime();
  let sum = 0;
  for (const [k, used] of Object.entries(hist)) {
    const end = +k * 3600e3, start = end - 7 * 864e5, overlap = Math.max(0, Math.min(end, m1) - Math.max(start, m0));
    sum += used * overlap / (7 * 864e5);
  }
  return Math.min(1, sum / 100 / ((m1 - m0) / (7 * 864e5)));
}
ipcMain.handle('usage:refresh', async () => { await refreshUsage(true); return usage; });

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
      const s = liveSession(sid), h = [...histCache.values()].map(c => c.meta).find(m => m && m.sid === sid);
      if (s && s.owner) { const p = ptys.get(s.owner); if (p) try { p.proc.kill(); } catch { } await new Promise(r => setTimeout(r, 1500)); }
      return convertToAgent(sid, (s && s.cwd) || (h && h.cwd) || HOME, task);
    },
  });
}
ipcMain.handle('phone:status', () => phone.status());
ipcMain.handle('phone:set', (_e, v) => { phone.setEnabled(v); return phone.status(); });
ipcMain.handle('phone:regen', () => { phone.regenerate(); return phone.status(); });
ipcMain.handle('phone:qr', (_e, url) => require('qrcode').toDataURL(url, { margin: 1, width: 360, color: { dark: '#0a0826', light: '#ffffff' } }));

// ---------------- telemetry loop ----------------
async function loop() {
  try {
    await collector.tick();
    if (win) {
      const ptyMap = new Map([...ptys].map(([id, p]) => [id, p.proc.pid]));
      const snap = collector.snapshot(ptyMap);
      snap.agents = agentsCache; refreshAgents();
      snap.usage = usage; refreshUsage();
      snap.ptys = [...ptys].map(([id, p]) => ({ id, pid: p.proc.pid, title: p.title, kind: p.kind, cwd: p.cwd }));
      win.webContents.send('telemetry', snap);
      lastSnap = snap; if (phone) phone.broadcast(snap);
    }
  } catch (e) { console.error(e); }
  setTimeout(loop, 1000);
}

app.whenReady().then(() => {
  app.setName('Claude Mesh');
  createWindow();
  initPhone();
  loop();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', () => {
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
