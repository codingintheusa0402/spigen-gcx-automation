// GCX Mesh for iPhone — a home-screen web app served by the Mac app (phone-server.js).
// Same mesh engine as the desktop (mesh/mother/bubbles/organism.js); live data arrives over SSE,
// controls go back as small JSON POSTs. Pairing key: ?t=… in the URL, remembered on the phone.
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt$ = v => '$' + (v >= 100 ? v.toFixed(0) : (v || 0).toFixed(2));
const ago = s => s < 60 ? Math.round(s) + 's' : s < 3600 ? Math.round(s / 60) + 'm' : s < 86400 ? (s / 3600).toFixed(1) + 'h' : Math.round(s / 86400) + 'd';

// ---------- pairing ----------
const qs = new URLSearchParams(location.search);
let TOKEN = qs.get('t') || localStorage.getItem('mesh.token') || '';
if (qs.get('t')) localStorage.setItem('mesh.token', qs.get('t'));
$('#manifest').href = '/manifest.webmanifest?t=' + encodeURIComponent(TOKEN);
const api = (p, body) => fetch(p + (p.includes('?') ? '&' : '?') + 't=' + encodeURIComponent(TOKEN), body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : {})
  .then(r => r.json().catch(() => ({ ok: r.ok })));

function toast(m) { const t = $('#toast'); t.textContent = m; t.classList.add('show'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('show'), 2400); }
function offline(msg) { const o = $('#offline'); if (msg) { o.innerHTML = msg; o.classList.remove('hidden'); } else o.classList.add('hidden'); }

// ---------- mesh (cap pixel density at 2× to keep the phone cool) ----------
const realDPR = window.devicePixelRatio || 1;
Object.defineProperty(window, 'devicePixelRatio', { get: () => Math.min(2, realDPR) });
let snap = { sessions: [], models: [], agents: [] }, selected = null, openTab = 'mesh';
const viz = new Mesh($('#mesh'), { onSelect: sid => sid && openSession(sid), onHub: () => openBroadcast() });
viz.wheelZoom = e => e.preventDefault();
if (innerWidth < 700) { viz.cam.ts = viz.cam.s = .66; viz.clampCam(); viz.cam.tx = viz.cam.ttx; viz.cam.ty = viz.cam.tty; }
setTimeout(() => { $('#hint').style.opacity = 0; }, 6000);

// touch: one finger pans (or drags a session to a new orbit), two fingers pinch-zoom, a tap selects
const cv = $('#mesh');
let T = null, holdT = null;
const pt = t => ({ clientX: t.clientX, clientY: t.clientY });
cv.addEventListener('touchstart', e => {
  e.preventDefault(); viz.lastInput = performance.now();
  const c = viz.cam;
  if (e.touches.length === 1) {
    const p = pt(e.touches[0]), h = viz.hit(p);
    T = { mode: 'one', x0: p.clientX, y0: p.clientY, tx: c.ttx, ty: c.tty, hit: h, moved: false };
    clearTimeout(holdT);
    if (h && h !== 'hub') holdT = setTimeout(() => { if (T && !T.moved && T.hit === h) { T.held = true; navigator.vibrate && navigator.vibrate(10); openDetails(h); } }, 520);
  } else if (e.touches.length === 2) {
    if (T && T.mode === 'drag') viz.dragEnd();
    const [a, b] = [pt(e.touches[0]), pt(e.touches[1])], r = cv.getBoundingClientRect();
    T = { mode: 'pinch', d0: Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY), s0: c.ts, mx: (a.clientX + b.clientX) / 2 - r.left, my: (a.clientY + b.clientY) / 2 - r.top, tx: c.ttx, ty: c.tty, moved: true };
  }
}, { passive: false });
cv.addEventListener('touchmove', e => {
  e.preventDefault(); viz.lastInput = performance.now(); if (!T) return;
  const c = viz.cam;
  if (T.mode === 'pinch' && e.touches.length === 2) {
    const [a, b] = [pt(e.touches[0]), pt(e.touches[1])], d = Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
    const s = Math.min(2, Math.max(.65, T.s0 * d / T.d0));
    c.ttx = T.mx - (T.mx - T.tx) * s / T.s0; c.tty = T.my - (T.my - T.ty) * s / T.s0; c.ts = s;
    viz.clampCam(); c.s = c.ts; c.tx = c.ttx; c.ty = c.tty; return;
  }
  const p = pt(e.touches[0]);
  if (T.held) return;
  if (!T.moved && Math.hypot(p.clientX - T.x0, p.clientY - T.y0) > 8) { clearTimeout(holdT);
    T.moved = true;
    if (T.hit && T.hit !== 'hub') { T.mode = 'drag'; viz.dragStart({ clientX: T.x0, clientY: T.y0 }, T.hit); }
  }
  if (!T.moved) return;
  if (T.mode === 'drag') viz.dragMove(p);
  else { c.ttx = T.tx + p.clientX - T.x0; c.tty = T.ty + p.clientY - T.y0; viz.clampCam(); c.tx = c.ttx; c.ty = c.tty; }
}, { passive: false });
cv.addEventListener('touchend', e => {
  e.preventDefault(); clearTimeout(holdT); if (!T) return;
  if (T.held) { if (e.touches.length === 0) T = null; return; }
  if (T.mode === 'drag') { viz.dragEnd(); toast('Orbit saved on this phone'); }
  else if (!T.moved && e.touches.length === 0) {
    if (T.hit === 'hub') openBroadcast(); else if (T.hit) openSession(T.hit);
  }
  if (e.touches.length === 0) T = null;
}, { passive: false });

// ---------- live data ----------
let es = null, lastMsg = 0;
function connect() {
  if (!TOKEN) return offline('Not paired yet.<br>On your Mac open <b>GCX Mesh → 📱</b> and scan the QR code.');
  es && es.close();
  es = new EventSource('/api/events?t=' + encodeURIComponent(TOKEN));
  es.addEventListener('telemetry', ev => {
    lastMsg = Date.now(); offline(); $('#conn').classList.add('on');
    snap = JSON.parse(ev.data); viz.update(snap); render();
  });
  es.onerror = () => { $('#conn').classList.remove('on'); };
}
setInterval(() => {
  if (Date.now() - lastMsg > 8000) {
    $('#conn').classList.remove('on');
    if (lastMsg) offline('Can’t reach your Mac.<br><span style="font-size:13px">Is GCX Mesh open and awake, with phone access on? Away from home, turn on Tailscale on this phone.</span>');
    if (!es || es.readyState === 2) connect();
  }
}, 3000);
fetch('/api/history?t=' + encodeURIComponent(TOKEN)).then(r => { if (r.status === 401) offline('This pairing key is no longer valid.<br>Scan the QR code in <b>GCX Mesh → 📱</b> again.'); else connect(); }).catch(() => connect());
document.addEventListener('visibilitychange', () => { if (!document.hidden && (!es || es.readyState !== 1)) connect(); });

function render() {
  $('#today').textContent = fmt$(snap.todayUsd || 0) + ' today';
  if (openTab === 'sessions') renderSessions();
  if (selected && !$('#sheet').classList.contains('hidden') && sheetKind === 'session') renderSession();
}

// ---------- sheets ----------
let sheetKind = null;
function sheet(kind, html, full) {
  sheetKind = kind; $('#sheetBody').innerHTML = html;
  $('#sheet').classList.toggle('full', !!full); $('#sheet').classList.remove('hidden');
  const x = $('#sheetBody .x'); if (x) x.onclick = closeSheet;
}
function closeSheet() { $('#sheet').classList.add('hidden'); sheetKind = null; selected = null; viz.selected = null; setTab('mesh', true); }
// swipe the sheet down to close
{ let y0 = null; const sh = $('#sheet');
  sh.addEventListener('touchstart', e => { y0 = $('#sheetBody').scrollTop <= 0 ? e.touches[0].clientY : null; }, { passive: true });
  sh.addEventListener('touchend', e => { if (y0 != null && e.changedTouches[0].clientY - y0 > 90 && sheetKind !== 'term') closeSheet(); y0 = null; }, { passive: true }); }

function setTab(t, quiet) {
  openTab = t;
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  if (quiet) return;
  if (t === 'mesh') closeSheet();
  if (t === 'sessions') renderSessions(true);
  if (t === 'history') openHistory();
  if (t === 'term' && termId) openTerminal(termId);
}
document.querySelectorAll('#tabs button').forEach(b => b.onclick = () => setTab(b.dataset.tab));

// ---------- sessions list ----------
function renderSessions(open) {
  const S = snap.sessions, A = snap.agents || [];
  const html = `<div class="sh-h"><h2>Sessions · ${S.length}</h2><button class="x">×</button></div>` +
    (S.map(s => `<div class="item" data-sid="${s.sid}"><span class="dot ${s.health}"></span><div class="t"><div class="n">${esc(s.name)}</div>
      <div class="m">${esc((s.model || '—').replace('claude-', ''))} · ${fmt$(s.todayUsd)} · ${s.health === 'working' ? Math.round(s.tps) + ' tok/s' : s.health}${s.owner ? ' · in-app' : ''}</div></div>›</div>`).join('') || '<div class="m" style="padding:14px 4px;color:var(--dim)">No Claude Code sessions running.</div>') +
    (A.length ? `<h3>Agents</h3>` + A.map(a => `<div class="item"><span class="dot ${a.state === 'running' ? 'working' : 'waiting'}"></span><div class="t"><div class="n">${esc(a.name || a.id)}</div><div class="m">${esc(a.state || '')}</div></div></div>`).join('') : '');
  if (open || sheetKind === 'sessions') { sheet('sessions', html); document.querySelectorAll('#sheetBody .item[data-sid]').forEach(el => el.onclick = () => openSession(el.dataset.sid)); }
}

// ---------- one session ----------
const QUICK = ['/compact', '/context', '/cost', '/model', '/doctor', '/skills', '/todos', '/status'];
const cur = () => snap.sessions.find(s => s.sid === selected);
function openSession(sid) { selected = sid; viz.selected = sid; openTab = 'sessions'; sheetKind = 'session'; renderSession(true); }
function renderSession(first) {
  const s = cur();
  if (!s) { if (first) sheet('session', `<div class="sh-h"><h2>Session ended</h2><button class="x">×</button></div>`); return; }
  const can = !!(s.owner || (s.tty && s.kind !== 'bg')), col = FAMILY_COLOR[s.family] || '#9fb4d8', pct = s.ctxMax ? s.ctx / s.ctxMax * 100 : 0;
  const stats = `<div class="meta"><span style="color:${col}">${esc(s.model || 'no model yet')}</span><span style="color:${HEALTH_COLOR[s.health]}">${esc(s.health)}</span><span>${s.owner ? 'in-app' : s.kind === 'bg' ? 'background' : 'Terminal.app'}</span><span>${esc((s.cwd || '').replace(/^\/Users\/[^/]+/, '~'))}</span></div>
    <div class="kpis"><div class="kpi"><div class="l">Speed</div><div class="v">${Math.round(s.tps)}<small> t/s</small></div></div>
      <div class="kpi"><div class="l">Today</div><div class="v">${fmt$(s.todayUsd)}</div></div><div class="kpi"><div class="l">Session</div><div class="v">${fmt$(s.totalUsd)}</div></div></div>
    <div class="meta" style="justify-content:space-between;margin:0"><span>Context</span><span>${Math.round(pct)}%</span></div>
    <div class="bar"><div style="width:${Math.min(100, pct)}%;background:${pct > 85 ? 'var(--err)' : pct > 60 ? 'var(--warn)' : col}"></div></div>`;
  const log = `<h3>Activity</h3><div class="log">${s.log.slice(-10).reverse().map(l => `<div class="${l.kind}">${l.kind === 'tool' ? '⚙ ' : l.kind === 'you' ? '› ' : ''}${esc(l.text)}</div>`).join('') || '<div>—</div>'}</div>`;
  if (!first && sheetKind === 'session' && $('#sStats')) { $('#sStats').innerHTML = stats; $('#sLog').innerHTML = log; $('#sName').textContent = s.name; return; }
  sheet('session', `<div class="sh-h"><span class="dot ${s.health}"></span><h2 id="sName">${esc(s.name)}</h2><button class="x">×</button></div>
    <div id="sStats">${stats}</div>
    ${can ? `<textarea id="sText" rows="3" placeholder="Message ${esc(s.name)}…"></textarea>
    <div class="row"><button class="primary" id="sSend">Send</button><button id="sInt">Interrupt</button>${s.owner ? '<button id="sTerm">Open terminal</button>' : ''}</div>
    <div class="chips">${QUICK.map(c => `<button data-c="${c}">${c}</button>`).join('')}<button id="sMore">More…</button></div>` : '<div class="meta">Background session — view only.</div>'}
    <div class="row"><button id="sRen">✎ Rename</button><button id="sAg">⚙ Convert to agent</button></div>
    <div id="sLog">${log}</div>`);
  if (can) {
    $('#sSend').onclick = async () => { const v = $('#sText').value.trim(); if (!v) return; const r = await api('/api/send', { sid: s.sid, text: v }); if (r.ok === false || r.out === 'notfound') return toast('Send failed' + (r.err ? ': ' + r.err : '')); $('#sText').value = ''; toast('Sent'); };
    $('#sInt').onclick = async () => { await api('/api/interrupt', { sid: s.sid }); toast('Interrupt sent'); };
    document.querySelectorAll('#sheetBody [data-c]').forEach(b => b.onclick = () => runCmd(s, b.dataset.c));
    $('#sMore').onclick = () => openCommands(s);
    const tb = $('#sTerm'); if (tb) tb.onclick = () => openTerminal(s.owner);
  }
  $('#sRen').onclick = async () => { const n = prompt('Rename session', s.name); if (n && n.trim()) { const r = await api('/api/rename', { sid: s.sid, name: n.trim() }); toast(r.ok === false ? 'Rename failed' : 'Renamed'); } };
  $('#sAg').onclick = async () => {
    if (!confirm(`Continue "${s.name}" as a background agent on your Mac?` + (s.owner ? '\n(its in-app terminal will close)' : s.tty ? '\n(a copy starts; the Terminal one keeps running)' : ''))) return;
    const task = prompt('Task for the agent (optional)', '') || '';
    toast('Starting agent…'); const r = await api('/api/convert', { sid: s.sid, task }); toast(r.ok ? 'Agent started' : 'Convert failed');
  };
}
async function runCmd(s, cmd) {
  if (/^\/(clear|exit)$/.test(cmd) && !confirm(`Run ${cmd} in ${s.name}? This can't be undone.`)) return;
  const r = await api('/api/send', { sid: s.sid, text: cmd });
  toast(r.ok === false || r.out === 'notfound' ? 'Failed' : `${cmd} → ${s.name}`);
}
const ALL_BUILTIN = ['/compact', '/clear', '/context', '/cost', '/usage', '/status', '/todos', '/doctor', '/help', '/model', '/skills', '/agents', '/mcp', '/memory', '/config', '/permissions', '/hooks', '/init', '/review', '/release-notes', '/export', '/rewind', '/resume', '/remote-control', '/exit'];
async function openCommands(s) {
  const mine = await api('/api/commands?cwd=' + encodeURIComponent(s.cwd || ''));
  const list = ALL_BUILTIN.map(c => ({ cmd: c, desc: '' })).concat(Array.isArray(mine) ? mine : []);
  sheet('cmds', `<div class="sh-h"><h2>Commands</h2><button class="x">×</button></div><input id="cq" placeholder="Filter…">
    <div id="cl"></div>`, true);
  const draw = () => { const q = $('#cq').value.toLowerCase().replace(/^\//, ''); $('#cl').innerHTML = list.filter(c => !q || c.cmd.includes(q) || (c.desc || '').toLowerCase().includes(q))
    .map(c => `<div class="item" data-c="${esc(c.cmd)}"><div class="t"><div class="n">${esc(c.cmd)}</div>${c.desc ? `<div class="m">${esc(c.desc)}</div>` : ''}</div></div>`).join('');
    document.querySelectorAll('#cl .item').forEach(el => el.onclick = () => { runCmd(s, el.dataset.c); openSession(s.sid); }); };
  $('#cq').oninput = draw; draw();
}

// ---------- broadcast (tap the mother star) ----------
function openBroadcast() {
  const ctl = snap.sessions.filter(s => s.owner || (s.tty && s.kind !== 'bg'));
  sheet('bc', `<div class="sh-h"><h2>Broadcast</h2><button class="x">×</button></div>
    <div class="meta">Send one message to every session you can control (${ctl.length}).</div>
    <textarea id="bText" rows="4" placeholder="Message all sessions…"></textarea><div class="row"><button class="primary" id="bGo">Send to ${ctl.length}</button></div>`);
  $('#bGo').onclick = async () => {
    const v = $('#bText').value.trim(); if (!v || !ctl.length || !confirm(`Send to ${ctl.length} sessions?`)) return;
    await Promise.all(ctl.map(s => api('/api/send', { sid: s.sid, text: v }))); toast('Broadcast sent'); closeSheet();
  };
}

// ---------- all sessions (resume on the Mac) ----------
let hist = [];
async function openHistory() {
  sheet('history', `<div class="sh-h"><h2>All sessions</h2><button class="x">×</button></div><input id="hq" placeholder="Search…"><div id="hl"><div class="meta" style="padding:12px 0">Loading…</div></div>`, true);
  hist = await api('/api/history'); if (!Array.isArray(hist)) hist = [];
  const live = new Set(snap.sessions.map(s => s.sid)), title = h => h.title || h.aiTitle || h.first || h.last || h.sid.slice(0, 8);
  const draw = () => {
    const q = $('#hq').value.toLowerCase();
    $('#hl').innerHTML = hist.filter(h => h.entry !== 'sdk-cli' && (!q || [h.title, h.aiTitle, h.first, h.cwd].some(v => v && v.toLowerCase().includes(q)))).slice(0, 120)
      .map(h => `<div class="item" data-h="${h.sid}"><div class="t"><div class="n">${esc(title(h))}</div><div class="m">${esc((h.cwd || '').split('/').pop() || '~')} · ${ago((Date.now() - h.mtime) / 1000)}</div></div>${live.has(h.sid) ? '<span class="tag live">live</span>' : ''}</div>`).join('');
    document.querySelectorAll('#hl .item').forEach(el => el.onclick = async () => {
      const sid = el.dataset.h;
      if (live.has(sid)) return openSession(sid);
      if (!confirm(`Resume "${title(hist.find(h => h.sid === sid))}" on your Mac?`)) return;
      const r = await api('/api/resume', { sid }); if (r.ok === false) return toast('Mac didn’t respond');
      toast('Resuming on your Mac…');
      const t0 = Date.now(), iv = setInterval(() => { if (snap.sessions.some(s => s.sid === sid)) { clearInterval(iv); openSession(sid); } else if (Date.now() - t0 > 25000) clearInterval(iv); }, 700);
    });
  };
  $('#hq').oninput = draw; draw();
}

// ---------- terminal of an in-app session ----------
let termId = null, term = null, tes = null;
function openTerminal(id) {
  termId = id; $('#termTab').classList.remove('hidden'); setTab('term', true);
  sheet('term', `<div id="termWrap"><div class="sh-h"><h2>Terminal</h2><button class="x">×</button></div><div id="xt"></div>
    <div class="keys">${[['Esc', '\x1b'], ['Tab', '\t'], ['↑', '\x1b[A'], ['↓', '\x1b[B'], ['←', '\x1b[D'], ['→', '\x1b[C'], ['⏎', '\r'], ['^C', '\x03'], ['⇧Tab', '\x1b[Z']].map(([l, k]) => `<button data-k="${encodeURIComponent(k)}">${l}</button>`).join('')}</div>
    <div class="tin"><input id="tIn" placeholder="Type, then Send (adds Enter)" autocapitalize="off" autocorrect="off"><button class="primary" id="tGo">Send</button></div></div>`, true);
  tes && tes.close(); term && term.dispose();
  const W = $('#xt').clientWidth - 8;
  term = new Terminal({ fontFamily: 'Menlo, monospace', fontSize: 9, convertEol: false, disableStdin: false, scrollback: 5000,
    theme: { background: '#06051a', foreground: '#e6e9f2', cursor: '#a78bfa' } });
  term.open($('#xt'));
  const write = d => api('/api/term/' + id, { data: d });
  term.onData(write);
  tes = new EventSource('/api/term/' + id + '?t=' + encodeURIComponent(TOKEN));
  tes.addEventListener('init', ev => {
    const i = JSON.parse(ev.data), fs = Math.max(7.5, Math.min(11, W / (i.cols * .6)));   // fit the Mac's columns; if still too wide, scroll sideways
    term.options.fontSize = fs; term.resize(i.cols, i.rows); term.reset(); term.write(i.backlog);
  });
  tes.addEventListener('data', ev => term.write(JSON.parse(ev.data)));
  tes.addEventListener('exit', () => term.write('\r\n\x1b[90m[process exited]\x1b[0m\r\n'));
  document.querySelectorAll('.keys button').forEach(b => b.onclick = () => write(decodeURIComponent(b.dataset.k)));
  const go = () => { const v = $('#tIn').value; write(v); setTimeout(() => write('\r'), 60); $('#tIn').value = ''; };
  $('#tGo').onclick = go; $('#tIn').onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); go(); } };
}

// ---------- press-and-hold a session: what it does, skills, commands, tools, CLAUDE.md ----------
async function openDetails(sid) {
  const s = snap.sessions.find(x => x.sid === sid);
  sheet('details', `<div class="sh-h"><h2>${esc(s ? s.name : 'Session')}</h2><button class="x">×</button></div><div class="meta">Reading the session…</div>`, true);
  const d = await api('/api/details?sid=' + encodeURIComponent(sid));
  if (sheetKind !== 'details') return;
  if (!d || !d.ok) { $('#sheetBody .meta').textContent = (d && d.err) || 'No details'; return; }
  const chips = list => list.length ? `<div class="chips">${list.map(([k, n]) => `<span class="tag" style="font-size:12.5px;padding:4px 9px">${esc(k)} <span style="opacity:.6">${n}</span></span>`).join('')}</div>` : '<div class="meta">none</div>';
  const what = d.summary ? mdToHtml(d.summary) + '<div class="meta">From the latest /compact summary</div>'
    : `${d.first ? `<p><b>Started with:</b> ${esc(d.first)}</p>` : ''}${d.last && d.last !== d.first ? `<p><b>Latest ask:</b> ${esc(d.last)}</p>` : ''}${d.lastText ? '<p><b>Latest reply:</b></p>' + mdToHtml(d.lastText) : ''}`;
  $('#sheetBody').innerHTML = `<div class="sh-h"><h2>${esc(s ? s.name : d.title || 'Session')}</h2><button class="x">×</button></div>
    <div class="meta">${esc((d.cwd || '').replace(/^\/Users\/[^/]+/, '~'))}</div>
    <h3>Skills used · ${d.skills.length}</h3>${chips(d.skills)}
    <h3>Slash commands · ${d.cmds.length}</h3>${chips(d.cmds)}
    <h3>What this session does</h3><div class="md" style="max-height:40vh;overflow-y:auto">${what}</div>
    <h3>Tools · ${d.tools.length}</h3>${chips(d.tools)}
    ${d.instructions.map(f => `<details><summary style="color:var(--dim);font-size:12px;margin:12px 0 4px">Instructions · ${esc(f.path)}</summary><div class="md">${mdToHtml(f.text)}</div></details>`).join('')}
    <div class="row" style="margin-top:14px"><button class="primary" id="dOpen">Open session</button></div>`;
  $('#sheetBody .x').onclick = closeSheet; $('#dOpen').onclick = () => openSession(sid);
}
