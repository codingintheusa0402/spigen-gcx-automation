const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt$ = v => v >= 100 ? '$' + v.toFixed(0) : '$' + (v || 0).toFixed(2);
const fmtN = v => v >= 1e9 ? (v / 1e9).toFixed(2) + 'B' : v >= 1e6 ? (v / 1e6).toFixed(1) + 'M' : v >= 1e3 ? (v / 1e3).toFixed(1) + 'k' : String(Math.round(v || 0));
const ago = s => s < 60 ? Math.round(s) + 's' : s < 3600 ? Math.round(s / 60) + 'm' : (s / 3600).toFixed(1) + 'h';
const hhmm = t => t ? new Date(t).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }) : '';
const store = { get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { } } };

// live-dot glow: rows are re-rendered every tick, so pin each pulse to the shared document clock (no restart flicker)
document.addEventListener('animationstart', e => {
  if (e.animationName === 'liveGlow') for (const a of e.target.getAnimations()) if (a.animationName === 'liveGlow') a.startTime = 0;
});

let snap = { sessions: [], models: [], ptys: [] };
let selected = null;
let bell = store.get('bell', true);
const prevHealth = new Map();
// sessions that finished (or stopped to ask you something) and you haven't looked at yet → red title, soft blink
const doneSids = new Set(store.get('doneSids', []));
const saveDone = () => store.set('doneSids', [...doneSids]);
const doneFade = new Map();   // sid → when it was looked at; fades out over FADE_MS instead of vanishing
const FADE_MS = 900, BLINK_MS = 2400;
// lists re-render every second, so the animations are pinned to the clock (negative delay) instead of restarting
function doneCls(sid) {
  if (doneSids.has(sid)) return ` done" style="animation-delay:-${(Date.now() % BLINK_MS) / 1000}s`;
  const f = doneFade.get(sid); if (!f) return '';
  const el = Date.now() - f.t; if (el >= FADE_MS) { doneFade.delete(sid); return ''; }
  return ` done-out" style="--a:${f.a.toFixed(3)};animation-delay:-${el / 1000}s`;
}
function seen(sid) { if (sid && doneSids.delete(sid)) { saveDone(); { const t = Date.now(); doneFade.set(sid, { t, a: .5 - .5 * Math.cos(2 * Math.PI * (t % BLINK_MS) / BLINK_MS) }); /* start from where the glow is now */ } renderList(); typeof renderHistory === 'function' && renderHistory();
  setTimeout(() => { renderList(); typeof renderHistory === 'function' && renderHistory(); }, FADE_MS + 50); } }
const terms = new Map();   // ptyId -> {term, fit, el, title, kind}
let activeTerm = null;

function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(t._h); t._h = setTimeout(() => t.classList.remove('show'), 2600); }

// ---------------- mesh ----------------
const viz = new Mesh($('#mesh'), { onSelect: sid => select(sid), onHub: () => openBroadcast() });
window.api.onVisible(v => { viz.offscreen = !v; });

// ---------------- telemetry ----------------
window.api.onTelemetry(s => {
  snap = s;
  if (s.ready) $('#loading').style.display = 'none';
  viz.update(s);
  renderHeader(); renderList(); renderToday(); renderPanel(); syncPeerTerms(); syncTermTitles(); notifyTransitions(); renderRail(); renderBoard();
});

// memory of each machine, as Activity Monitor's "Memory Used"; the bar turns amber/red with macOS memory pressure
function memPill(label, m) {
  if (!m || !m.total) return '';
  const gb = b => (b / 2 ** 30).toFixed(1), pct = Math.min(100, Math.round(m.used / m.total * 100));
  const lvl = m.pressure >= 4 ? 'crit' : m.pressure >= 2 ? 'warn' : 'ok';
  return `<span class="hpill mem ${lvl}" title="${esc(label)}: ${gb(m.used)} of ${gb(m.total)} GB memory used${lvl === 'ok' ? '' : ' · memory pressure ' + (lvl === 'crit' ? 'critical' : 'high')}">${esc(label)} <i>${gb(m.used)}/${Math.round(m.total / 2 ** 30)} GB</i><span class="membar"><span style="width:${pct}%"></span></span></span>`;
}
function renderHeader() {
  const order = ['fable', 'mythos', 'opus', 'sonnet', 'haiku', 'other'];
  const ms = [...snap.models].sort((a, b) => order.indexOf(a.family) - order.indexOf(b.family));
  $('#models').innerHTML = ms.map(m => `<div class="mchip" title="${m.sessions} live session(s) · ${fmtN(m.todayOut)} output tokens today · burn ${fmt$(m.liveUsdHr)}/h (last 5 min)">
    <span class="dot" style="background:${FAMILY_COLOR[m.family]};box-shadow:0 0 8px ${FAMILY_COLOR[m.family]}"></span>
    <b>${m.family[0].toUpperCase() + m.family.slice(1)}</b><span>${fmt$(m.todayUsd)}</span>
    <span class="sub">${Math.round(m.tps)} tok/s · ${fmt$(m.liveUsdHr)}/h</span></div>`).join('');
  const c = k => snap.sessions.filter(s => s.health === k).length;
  const errs = snap.sessions.reduce((a, s) => a + s.errors1h, 0);
  $('#health').innerHTML = `<span class="hpill">working <i>${c('working')}</i></span>
    <span class="hpill">idle <i>${c('idle')}</i></span>
    ${c('waiting') + c('quiet') ? `<span class="hpill" style="border-color:#6b5a1a">attention <i>${c('waiting') + c('quiet')}</i></span>` : ''}
    <span class="hpill" ${errs ? 'style="border-color:#6b2a2a"' : ''}>API errors 1h <i>${errs}</i></span>
    ${snap.peer && snap.peer.online && snap.peer.build && snap.build && snap.peer.build !== snap.build
      ? `<span class="hpill" style="border-color:#6b5a1a" title="This Mac runs build ${esc(snap.build)}, ${esc(snap.peer.label)} runs ${esc(snap.peer.build)}. Install the same GCX Mesh on both and restart it.">versions differ</span>` : ''}
    ${memPill(snap.peer ? 'This Mac' : 'RAM', snap.mem)}${snap.peer && snap.peer.mem ? memPill(snap.peer.label, snap.peer.mem) : ''}`;
}

function renderList() {
  $('#sessCount').textContent = `· ${snap.sessions.length}`;
  $('#list').innerHTML = snap.sessions.map(s => `<div class="srow ${s.sid === selected ? 'sel' : ''}${doneCls(s.sid)}" data-sid="${s.sid}">
    <div class="n"><span class="hdot h-${s.health}"></span><span class="t">${esc(s.name)}</span>${s.onPeer || s.hostLabel ? peerTag(s) : s.owner ? '<span class="tag">app</span>' : s.kind === 'bg' ? '<span class="tag">bg</span>' : ''}<span class="more" data-more="${s.sid}" title="Commands (or right-click)">⋯</span></div>
    <div class="m"><span style="color:${FAMILY_COLOR[s.family]}">${esc((s.model || '—').replace('claude-', ''))}</span><span>${fmt$(s.todayUsd)}</span><span>${s.health === 'working' ? Math.round(s.tps) + ' tok/s' : ago(s.sinceWrite)}</span></div>
  </div>`).join('');
  for (const el of document.querySelectorAll('.srow')) el.onclick = () => select(el.dataset.sid);
}

// Simple-mode stand-in for the mesh: one static card per session; repainted only when its text changes
let lastBoard = '';
function renderBoard() {
  if (!isSimple()) return;
  const html = snap.sessions.map(s => `<div class="bcard ${s.sid === selected ? 'sel' : ''}" data-sid="${s.sid}">
    <div class="n"><span class="hdot h-${s.health}"></span><span class="t">${esc(s.name)}</span>${peerTag(s)}</div>
    <div class="m"><span style="color:${FAMILY_COLOR[s.family]}">${esc((s.model || '—').replace('claude-', ''))}</span><span>${fmt$(s.todayUsd)}</span><span>${s.health === 'working' ? 'working' : s.health}</span></div>
  </div>`).join('') || '<div class="bempty">No live sessions.</div>';
  if (html === lastBoard) return;
  lastBoard = html; $('#board').innerHTML = html;
}
$('#board').addEventListener('click', e => { const c = e.target.closest('.bcard'); if (c) select(c.dataset.sid); });
$('#board').addEventListener('contextmenu', e => { const c = e.target.closest('.bcard'); if (c) { e.preventDefault(); openCmdMenu(c.dataset.sid, e.clientX, e.clientY); } });

function renderToday() {
  const tot = snap.todayUsd || 0;
  const burn = snap.models.reduce((a, m) => a + m.liveUsdHr, 0);
  $('#today').innerHTML = `<div class="big">${fmt$(tot)}</div><div style="margin-bottom:10px">burn rate ${fmt$(burn)}/h · all transcripts</div>` +
    snap.models.filter(m => m.todayUsd > 0).sort((a, b) => b.todayUsd - a.todayUsd).map(m => `
    <div style="display:flex;justify-content:space-between"><span>${m.family}</span><span>${fmt$(m.todayUsd)}</span></div>
    <div class="bar"><div style="width:${tot ? m.todayUsd / tot * 100 : 0}%;background:${FAMILY_COLOR[m.family]}"></div></div>`).join('');
}

// ---------------- detail panel ----------------
const termOf = s => s && (s.owner || s.pane);   // the session's terminal tab: in-app pty or its tmux pane
function select(sid) {
  seen(sid);
  const ms = sid && snap.sessions.find(x => x.sid === sid);
  if (isMini() && termOf(ms) && terms.has(termOf(ms))) { selected = sid; viz.selected = sid; $('#panel').classList.add('hidden'); miniShow('term'); activate(termOf(ms)); renderList(); return; }
  selected = sid; viz.selected = sid;
  $('#panel').classList.toggle('hidden', !sid);
  $('#panelBody').innerHTML = sid ? '<div id="pStatic"></div><div id="pCtl"></div>' : '';
  if (sid) { renderCtl(); renderPanel(); }
  renderList(); renderBoard();
  // an in-app session: bring its terminal tab to the front in the dock (Split / Grid views)
  const s = sid && snap.sessions.find(x => x.sid === sid);
  if (s && termOf(s) && terms.has(termOf(s)) && termOf(s) !== activeTerm && document.body.dataset.view !== 'mesh') activate(termOf(s));
  requestAnimationFrame(() => viz.resize());
}

function cur() { return snap.sessions.find(s => s.sid === selected); }

// one path for every session: in-app terminal, Terminal.app tab, or a session on the other machine (peer)
function sessSend(s, text) {
  if (s.onPeer) return window.api.peer.send(s.sid, text);
  return s.owner ? window.api.ctl.send(s.owner, text, true) : window.api.ext.send(s.tty, text);
}
function sessInterrupt(s) {
  if (s.onPeer) return window.api.peer.interrupt(s.sid);
  return s.owner ? window.api.pty.write(s.owner, '\x1b') : window.api.ext.interrupt(s.tty);
}
const peerTag = s => s.onPeer ? `<span class="tag peer" title="Runs on ${esc(s.hostLabel)}">${esc(s.hostLabel)}${s.offline ? ' · offline' : ''}</span>`
  : s.hostLabel ? `<span class="tag peer" title="Runs on this machine (${s.owner ? 'in-app' : s.kind === 'bg' ? 'background' : 'Terminal'})">${esc(s.hostLabel)}</span>` : '';

function renderPanel() {
  if (!selected) return;
  const s = cur(), box = $('#pStatic'); if (!box) return;
  if (!s) { box.innerHTML = `<div class="ph"><div class="row1"><h2>Session ended</h2><button class="close" id="pClose">×</button></div></div>`; $('#pClose').onclick = () => select(null); return; }
  const logEl = box.querySelector('.log');
  const stick = !logEl || logEl.scrollTop + logEl.clientHeight >= logEl.scrollHeight - 8;
  const ctxPct = s.ctxMax ? s.ctx / s.ctxMax * 100 : 0;
  const col = FAMILY_COLOR[s.family];
  const models = Object.entries(s.byModel).sort((a, b) => b[1].usd - a[1].usd);
  const maxUsd = Math.max(...models.map(m => m[1].usd), 0.0001);
  box.innerHTML = `
  <div class="ph"><div class="row1"><span class="hdot h-${s.health}"></span><h2 title="${esc(s.name)}">${esc(s.name)}</h2><button class="close" id="pClose">×</button></div>
    <div class="meta"><span style="color:${col}">${esc(s.model || 'no model yet')}</span><span>${s.health}${s.status && s.status !== 'busy' && s.status !== 'idle' ? ' (' + esc(s.status) + ')' : ''}</span>
    <span>pid ${s.pid}</span><span>${s.tty ? s.tty.replace('/dev/', '') : 'no tty'}</span><span>${s.onPeer ? 'on ' + esc(s.hostLabel) : s.owner ? 'in-app' : s.kind === 'bg' ? 'background' : 'Terminal.app'}</span>
    ${s.remote ? '<span>remote-control on</span>' : ''}<span>v${esc(s.version)}</span><span title="${esc(s.cwd)}">${esc(s.cwd.replace(/^\/Users\/[^/]+/, '~'))}</span></div></div>
  <div class="sec"><h3>Live</h3><div class="kpis">
    <div class="kpi"><div class="l">Speed now</div><div class="v">${Math.round(s.tps)} <small>tok/s</small></div></div>
    <div class="kpi"><div class="l">Last reply</div><div class="v">${Math.round(s.lastSpeed)} <small>tok/s</small></div></div>
    <div class="kpi"><div class="l">Burn</div><div class="v">${fmt$(s.usdHr)}<small>/h</small></div></div>
    <div class="kpi"><div class="l">Today</div><div class="v">${fmt$(s.todayUsd)}</div></div>
    <div class="kpi"><div class="l">Session total</div><div class="v">${fmt$(s.totalUsd)}</div></div>
    <div class="kpi"><div class="l">Errors 5m / 1h</div><div class="v" style="color:${s.errors5m ? 'var(--err)' : 'inherit'}">${s.errors5m} <small>/ ${s.errors1h}</small></div></div>
    <div class="kpi"><div class="l">Sub-agents</div><div class="v">${s.subagents}</div></div>
    <div class="kpi"><div class="l">CPU · Mem</div><div class="v">${Math.round(s.cpu)}<small>% · ${s.rssMb}MB</small></div></div>
    <div class="kpi"><div class="l">Uptime · last write</div><div class="v" style="font-size:13px">${esc(s.etime)} <small>· ${ago(s.sinceWrite)}</small></div></div>
  </div>
  <div style="margin-top:10px;font-size:12px;color:var(--dim);display:flex;justify-content:space-between"><span>Context</span><span>${fmtN(s.ctx)} / ${fmtN(s.ctxMax)} · ${ctxPct.toFixed(0)}%</span></div>
  <div class="bar"><div style="width:${Math.min(100, ctxPct)}%;background:${ctxPct > 85 ? 'var(--err)' : ctxPct > 60 ? 'var(--warn)' : col}"></div></div>
  <div style="font-size:12px;color:var(--dim);margin-top:4px">tok/s · last 5 min</div><canvas class="spark" id="spark"></canvas>
  ${s.lastError ? `<div class="note" style="color:#fca5a5">last API error: ${esc(s.lastError)}</div>` : ''}
  </div>
  <div class="sec"><h3>Tokens (session)</h3><div class="tokgrid">
    <span>Input</span><span>${fmtN(s.tokIn)}</span><span>Output</span><span>${fmtN(s.tokOut)}</span>
    <span>Cache read</span><span>${fmtN(s.tokCacheRead)}</span><span>Cache write</span><span>${fmtN(s.tokCacheWrite)}</span><span>Turns</span><span>${s.turns}</span></div>
    ${models.length > 1 ? '<h3 style="margin-top:12px">Spend by model</h3>' + models.map(([m, v]) => `<div style="display:flex;justify-content:space-between;font-size:12px"><span>${esc(m.replace('claude-', ''))}</span><span>${fmt$(v.usd)}</span></div><div class="bar"><div style="width:${v.usd / maxUsd * 100}%;background:${FAMILY_COLOR[(/fable|mythos|opus|sonnet|haiku/.exec(m) || ['other'])[0]]}"></div></div>`).join('') : ''}
  </div>
  <div class="sec"><h3>Activity</h3><div class="log">${s.log.map(l => `<div class="${l.kind}"><time>${hhmm(l.t)}</time>${l.kind === 'tool' ? '⚙ ' : l.kind === 'you' ? '› ' : ''}${esc(l.text)}</div>`).join('') || '<div>No activity yet</div>'}</div></div>`;
  $('#pClose').onclick = () => select(null);
  const lg = box.querySelector('.log'); if (stick) lg.scrollTop = lg.scrollHeight; else if (logEl) lg.scrollTop = logEl.scrollTop;
  drawSpark($('#spark'), s.spark, col, s.sparkPeak);
}

function drawSpark(c, data, col, peak) {
  const d = devicePixelRatio, w = c.clientWidth, h = c.clientHeight;
  c.width = w * d; c.height = h * d;
  const x = c.getContext('2d'); x.scale(d, d);
  // steady scale: a 'nice' ceiling over the last 15 min's peak, so past spikes keep their height
  const top = Math.max(10, peak || 0, ...data), mag = Math.pow(10, Math.floor(Math.log10(top))), max = [1, 2, 2.5, 5, 10].map(k => k * mag).find(v => v >= top);
  x.strokeStyle = 'rgba(255,255,255,.06)'; x.beginPath(); x.moveTo(0, h - .5); x.lineTo(w, h - .5); x.stroke();
  const pts = data.map((v, i) => [i / (data.length - 1) * w, h - 4 - v / max * (h - 10)]);
  const g = x.createLinearGradient(0, 0, 0, h); g.addColorStop(0, col + '66'); g.addColorStop(1, col + '00');
  x.fillStyle = g; x.beginPath(); x.moveTo(0, h); pts.forEach(p => x.lineTo(...p)); x.lineTo(w, h); x.fill();
  x.strokeStyle = col; x.lineWidth = 1.5; x.beginPath(); pts.forEach((p, i) => i ? x.lineTo(...p) : x.moveTo(...p)); x.stroke();
  x.fillStyle = '#8a93ab'; x.font = '10px -apple-system'; x.fillText(Math.round(max) + ' tok/s', 4, 11);
}

// Controls stay mounted while stats re-render, so typing isn't interrupted.
function renderCtl() {
  const s = cur(); const box = $('#pCtl'); if (!box || !s) return;
  const peer = !!s.onPeer, live = peer ? !s.offline && (s.owner || (s.tty && s.kind !== 'bg')) : false;
  const owned = !peer && !!s.owner, ext = !peer && !owned && s.tty && s.kind !== 'bg';
  box.innerHTML = `<div class="sec ctl"><h3>Control</h3>
    <textarea id="cText" placeholder="${owned || ext || live ? 'Type a prompt… drop files or ⌘V images to attach (⌘↩ to send)' : 'This session has no terminal the app can reach'}"></textarea>
    <div class="btns">
      <button class="primary" id="cSend">Send</button>
      <button id="cEsc" title="Send Esc (interrupt the current turn)">Interrupt</button>
      <button id="cCompact">/compact</button>
      ${peer ? (s.owner && terms.has(s.owner) ? '<button id="cFocus">Show terminal</button>' : '') : `<button id="cFocus">${owned || s.pane ? 'Show terminal' : 'Focus in Terminal'}</button>`}
      ${owned ? '<button class="danger" id="cKill">Kill</button>' : ''}
    </div>
    <div class="note">${peer ? (live ? `Runs on ${esc(s.hostLabel)}: prompts are passed to it over Tailscale.` : `${esc(s.hostLabel)} is offline or this session can’t be reached.`) : owned ? 'Runs inside GCX Mesh — keystrokes go straight to its terminal.' :
      ext && s.pane ? 'Runs in a tmux pane on this Mac (keeps running when GCX Mesh restarts). Prompts are typed into it with tmux send-keys.' :
      ext ? 'Lives in a Terminal.app window. Prompts are typed into that tab via AppleScript (first use asks for Automation permission). Multi-line text is sent as one line.' :
      'Background job — view only.'}</div></div>`;
  const send = async (text) => {
    const s = cur(); if (!s || !text) return;
    let r;
    if (s.onPeer || s.owner || s.tty) r = await sessSend(s, text);
    if (r && (r.ok === false || r.out === 'notfound')) toast(r.out === 'notfound' ? 'No Terminal.app tab found for ' + s.tty : 'Send failed: ' + (r.err || ''));
    else { viz.command(s.sid); toast('Sent to ' + s.name); }
  };
  $('#cSend').onclick = () => { const v = $('#cText').value.trim(); if (v) { send(v); $('#cText').value = ''; } };
  const ta = $('#cText');
  const insertAt = t => { const a = ta.selectionStart, b = ta.selectionEnd; ta.value = ta.value.slice(0, a) + t + ta.value.slice(b); ta.selectionStart = ta.selectionEnd = a + t.length; ta.focus(); };
  ta.addEventListener('paste', async e => {
    const types = [...(e.clipboardData?.types || [])];
    if (types.includes('text/plain') && !types.includes('Files')) return;   // plain text: default paste
    e.preventDefault(); const r = await clipboardInsert(); if (r.text) insertAt(r.text);
  });
  wireDrop(ta, insertAt);
  $('#cText').onkeydown = e => { if (e.key === 'Enter' && e.metaKey) { e.preventDefault(); $('#cSend').click(); } };
  $('#cCompact').onclick = () => send('/compact');
  $('#cEsc').onclick = async () => {
    const s = cur(); if (!s) return;
    if (s.onPeer || s.owner || s.tty) await sessInterrupt(s);
    viz.command(s.sid); toast('Interrupt sent');
  };
  if (peer && $('#cFocus')) $('#cFocus').onclick = () => { if (document.body.dataset.view === 'mesh') setView('split'); activate(s.owner); };
  if (!peer) $('#cFocus').onclick = async () => {
    const s = cur(); if (!s) return;
    if (termOf(s) && terms.has(termOf(s))) { if (document.body.dataset.view === 'mesh') setView('split'); activate(termOf(s)); }
    else if (s.tty) { const r = await window.api.ext.focus(s.tty); if (r.out === 'notfound') toast('No Terminal.app tab for ' + s.tty); }
  };
  if (owned) $('#cKill').onclick = () => { window.api.pty.kill(s.owner); toast('Killed'); };
  if (!owned && !ext && !live) for (const id of ['#cSend', '#cEsc', '#cCompact', '#cFocus', '#cText']) if ($(id)) $(id).disabled = true;
}

// ---------------- broadcast (click the core) ----------------
function openBroadcast() {
  const targets = snap.sessions.filter(s => !s.offline && (s.owner || (s.tty && s.kind !== 'bg')));
  modal(`<h2>Broadcast to sessions</h2>
    <div class="blist">${targets.map(s => `<label class="chk"><input type="checkbox" data-sid="${s.sid}" checked><span class="hdot h-${s.health}"></span>${esc(s.name)} <span style="color:var(--dim)">${s.health}</span></label>`).join('') || '<div class="note">No controllable sessions.</div>'}</div>
    <label>Message</label><textarea id="bText" style="width:100%;min-height:80px" placeholder="e.g. /compact  ·  give me a one-line status"></textarea>
    <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary" id="bSend">Send to selected</button></div>`);
  $('#bSend').onclick = async () => {
    const text = $('#bText').value.trim(); if (!text) return;
    const sids = [...document.querySelectorAll('.blist input:checked')].map(i => i.dataset.sid);
    for (const sid of sids) {
      const s = snap.sessions.find(x => x.sid === sid); if (!s) continue;
      await sessSend(s, text);
      viz.command(sid);
    }
    closeModal(); toast(`Broadcast to ${sids.length} session(s)`);
  };
}

// ---------------- new session ----------------
async function openNew() {
  const home = await window.api.home();
  const last = store.get('newOpts', { kind: 'claude', cwd: home, flags: '--dangerously-skip-permissions --remote-control', model: '' });
  modal(`<h2>New session</h2>
    <div class="choice"><button data-k="claude">Claude Code</button><button data-k="shell">Shell</button></div>
    <label>Working directory</label><div class="row"><input type="text" id="nCwd" value="${esc(last.cwd)}"><button id="nBrowse" style="flex:none">Browse…</button></div>
    <div id="nClaude">
      <label>Model</label><select id="nModel" style="width:100%">
        ${[['', 'Default'], ['opus', 'Opus'], ['sonnet', 'Sonnet'], ['haiku', 'Haiku'], ['claude-fable-5-1', 'Fable 5.1']].map(([v, l]) => `<option value="${v}" ${v === last.model ? 'selected' : ''}>${l}</option>`).join('')}
      </select>
      <label>Flags</label><input type="text" id="nFlags" value="${esc(last.flags)}">
      <label class="chk"><input type="checkbox" id="nResume"> Resume a previous session (--resume picker)</label>
    </div>
    <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary" id="nGo">Start</button></div>`);
  let kind = last.kind;
  const setKind = k => { kind = k; document.querySelectorAll('.choice button').forEach(b => b.classList.toggle('on', b.dataset.k === k)); $('#nClaude').style.display = k === 'claude' ? '' : 'none'; };
  document.querySelectorAll('.choice button').forEach(b => b.onclick = () => setKind(b.dataset.k)); setKind(kind);
  $('#nBrowse').onclick = async () => { const d = await window.api.pickDir(); if (d) $('#nCwd').value = d; };
  $('#nGo').onclick = async () => {
    const o = { kind, cwd: $('#nCwd').value.trim() || home, flags: $('#nFlags').value.trim(), model: $('#nModel').value };
    store.set('newOpts', o);
    let flags = o.flags + (o.model ? ` --model ${o.model}` : '') + ($('#nResume').checked ? ' --resume' : '');
    closeModal();
    await newTerm({ kind, cwd: o.cwd, flags, title: kind === 'claude' ? 'claude · starting' : o.cwd.split('/').pop() || '~' });
  };
}

function modal(html) { $('#modalCard').classList.remove('wide'); $('#modalCard').innerHTML = html; $('#modal').classList.remove('hidden'); const c = $('#mCancel'); if (c) c.onclick = closeModal; }
function closeModal() { $('#modal').classList.add('hidden'); }
$('#modal').onclick = e => { if (e.target.id === 'modal') closeModal(); };

// ---------------- paste / drop of files & images ----------------
// Paths are backslash-escaped exactly like Terminal.app does on drag-and-drop, so Claude Code
// turns image paths into [Image #n] attachments and other files into @-mentionable paths.
const shq = p => p.replace(/([\s\\'"()&;|<>$`!*?#~{}\[\]])/g, '\\$1');
const droppedPaths = e => [...(e.dataTransfer?.files || [])].map(f => window.api.pathForFile(f)).filter(Boolean);
async function clipboardInsert(peer) {
  const c = await window.api.clipRead();
  if (c.kind === 'files') {
    if (peer) {                                                  // the session runs on the other machine: send the files there
      const u = await window.api.peer.upload(c.paths);
      if (u.err) { toast("Couldn't send the file: " + u.err); return { text: '', files: true }; }
      c.paths = u.paths;
    }
    if (c.image) toast('Image pasted as ' + c.paths[0].split('/').pop()); return { text: c.paths.map(shq).join(' ') + ' ', files: true };
  }
  return { text: c.text || '', files: false };
}
function wireDrop(el, insert) {
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('dropping'); });
  el.addEventListener('dragleave', () => el.classList.remove('dropping'));
  el.addEventListener('drop', e => {
    e.preventDefault(); el.classList.remove('dropping');
    const ps = droppedPaths(e); if (ps.length) insert(ps.map(shq).join(' ') + ' ');
  });
}

// ---------------- terminals ----------------
// OSC 52 = a program asking the terminal to set the clipboard. Claude Code copies its own mouse selections this
// way ("sent N chars via OSC 52"). Only the window being used takes it: a session shared with the other machine
// prints the same sequence into both windows, and copying on one Mac must not overwrite the other's clipboard.
function wireOsc52(term) {
  term.parser.registerOscHandler(52, data => {
    const b64 = data.slice(data.indexOf(';') + 1);
    if (!b64 || b64 === '?' || !document.hasFocus()) return true;
    try { window.api.clipWrite(new TextDecoder().decode(Uint8Array.from(atob(b64), c => c.charCodeAt(0)))); } catch { }
    return true;
  });
}
const termOpts = () => ({
  fontFamily: '"SF Mono", Menlo, monospace', fontSize: 12.5, lineHeight: 1.15, cursorBlink: !isSimple(), allowProposedApi: true,
  macOptionIsMeta: true, scrollback: 20000,
  theme: { background: '#06051a', foreground: '#e6e9f2', cursor: '#a78bfa', selectionBackground: '#3b4a7a88',
    black: '#1b1f2a', brightBlack: '#555e78', blue: '#7aa2f7', brightBlue: '#9ab8ff', magenta: '#bb9af7', cyan: '#2dd4bf', green: '#34d399', red: '#f87171', yellow: '#fbbf24' },
});

// Live mirrors of the other machine's in-app terminals (peer:… ids). They keep the host's size — the
// text is scaled to fit instead — and keystrokes go back to the host. Closing a mirror only detaches it.
const peerClosed = new Set();
function syncPeerTerms() {
  const want = new Map((snap.ptys || []).filter(p => p.onPeer || p.kind === 'tmux').map(p => [p.id, p]));   // the peer's terminals + our tmux panes
  for (const [id, p] of want) {
    if (peerClosed.has(id)) continue;
    let t = terms.get(id);
    if (!t) t = mirrorTerm(p);
    if (!t.attached) { t.attached = true; window.api.peerTerm.open(id); }
  }
  for (const [id, t] of terms) if (t.peer && !want.has(id)) { window.api.peerTerm.close(id); t.term.dispose(); t.el.remove(); terms.delete(id); if (activeTerm === id) activeTerm = null; renderTabs(); layoutGrid(); }
  for (const id of peerClosed) if (!want.has(id)) peerClosed.delete(id);
  $('#noTerms').style.display = terms.size ? 'none' : '';
}
function mirrorTerm(p) {
  const el = document.createElement('div'); el.className = 'term peer';
  el.innerHTML = `<div class="tlabel"><span class="hdot h-idle"></span><span class="ttl"></span></div><div class="xt"></div>`;
  $('#terms').appendChild(el);
  const term = new Terminal(termOpts());
  const fit = new FitAddon.FitAddon(); term.loadAddon(fit);
  term.open(el.querySelector('.xt'));
  wireOsc52(term);
  // only real keystrokes go back: xterm's automatic replies to terminal queries (device attributes, cursor
  // position, focus, colour reports) are already answered by the host's own terminal, and would land as typed text
  const isReport = d => /^\x1b\[[?>=]?[\d;]*[cnRtIOy]$/.test(d) || /^\x1b\[[IO]$/.test(d) || d.startsWith('\x1b]') || d.startsWith('\x1bP');
  term.onData(d => { if (!isReport(d)) window.api.pty.write(p.id, d); });
  el.addEventListener('paste', async e => { e.preventDefault(); e.stopPropagation(); const r = await clipboardInsert(true); if (r.text) term.paste(r.text); }, true);
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('dropping'); });
  el.addEventListener('dragleave', () => el.classList.remove('dropping'));
  el.addEventListener('drop', async e => {
    e.preventDefault(); el.classList.remove('dropping');
    const ps = droppedPaths(e); if (!ps.length) return;
    const u = await window.api.peer.upload(ps);
    if (u.err) return toast("Couldn't send the file: " + u.err);
    term.paste(u.paths.map(shq).join(' ') + ' '); term.focus();
  });
  term.textarea && term.textarea.addEventListener('focus', () => { activeTerm = p.id; markFocus(); });
  const t = { term, fit, el, title: p.title, kind: p.kind, peer: true, local: !p.onPeer, cols: 120, rows: 30, driving: false };
  terms.set(p.id, t);
  if (!activeTerm) activeTerm = p.id;
  renderTabs(); layoutGrid();
  return t;
}
// A terminal shared with the other machine has one size. Whichever window was used last sets it (keys, clicks,
// scrolling, focus): its panes fit normally and the session rewraps; the other side keeps that grid and scales
// the text to fit, so nothing re-wraps wrong there.
const BASE_FONT = termOpts().fontSize;
function fitScaled(t) {
  // always the given grid, even while the tab is hidden: output written at another width wraps wrong for good
  if (t.term.cols !== t.cols || t.term.rows !== t.rows) t.term.resize(t.cols, t.rows);
  const box = t.el.querySelector('.xt'), w = box.clientWidth - 12, h = box.clientHeight - 6; if (w <= 0 || h <= 0) return;
  const f0 = t.term.options.fontSize, cell = t.term._core?._renderService?.dimensions?.css?.cell;   // measured cell size at f0
  const cw = cell && cell.width ? cell.width / f0 : .602, ch = cell && cell.height ? cell.height / f0 : 1.15 * 1.18;
  const fs = Math.max(7, Math.min(BASE_FONT, w / (t.cols * cw), h / (t.rows * ch)));
  if (Math.abs(f0 - fs) > .05) t.term.options.fontSize = fs;
}
const termVisible = t => document.body.dataset.view === 'grid' || t.el.classList.contains('on');
// this window's own size for a terminal; a mirror also asks the host to resize the session to it
function fitOwn(id, t) {
  if (t.term.options.fontSize !== BASE_FONT) t.term.options.fontSize = BASE_FONT;
  if (!t.peer) return t.fit.fit();
  const box = t.el.querySelector('.xt'); if (box.clientWidth < 200 || box.clientHeight < 80) return;   // not laid out yet
  const d = t.fit.proposeDimensions(); if (!d || !(d.cols >= 20) || !(d.rows >= 5)) return;
  t.cols = Math.max(20, d.cols); t.rows = Math.max(5, d.rows);
  if (t.term.cols !== t.cols || t.term.rows !== t.rows) t.term.resize(t.cols, t.rows);
  if (t.sentCols !== t.cols || t.sentRows !== t.rows) {
    t.sentCols = t.cols; t.sentRows = t.rows; window.api.peerTerm.resize(id, t.cols, t.rows);
    // no confirmation (host offline, or a build that can't resize): show the host's real size again
    clearTimeout(t.confirmT);
    t.confirmT = setTimeout(() => { if (t.driving && (t.hostCols !== t.cols || t.hostRows !== t.rows)) { t.driving = false; t.noClaimUntil = Date.now() + 60000; t.sentCols = t.sentRows = 0; t.cols = t.hostCols || t.cols; t.rows = t.hostRows || t.rows; fitScaled(t); } }, 2000);
  }
}
function fitTerm(id, t) { (t.peer ? t.driving : !t.remote) ? fitOwn(id, t) : fitScaled(t); }
let lastUse = 0;
const peerCanResize = () => !!(snap.peer && snap.peer.online && (snap.peer.caps || []).includes('resize'));
function windowUsed(e) {
  if (!e.isTrusted || document.body.dataset.size === 'mini') return;   // real input only; the compact window never takes the size
  // never while copying: ⌘-shortcuts and bare modifier keys don't count, a click counts only once released without
  // selecting, and nothing is resized while text is selected (a resize clears the selection and Claude Code redraws)
  if (e.type === 'keydown' && (e.metaKey || ['Meta', 'Shift', 'Alt', 'Control', 'CapsLock'].includes(e.key))) return;
  for (const [, t] of terms) if (t.term.hasSelection()) return;
  if (e.type === 'mouseup' && (getSelection() || '').toString()) return;
  if (Date.now() - lastUse < 700) return; lastUse = Date.now();
  for (const [id, t] of terms) {
    if (!termVisible(t)) continue;
    // a host build that can't resize (or didn't confirm a minute ago) is left alone: asking again would flip the size back and forth
    if (t.peer && !t.driving && (t.local || peerCanResize()) && !(t.noClaimUntil > Date.now())) { t.driving = true; t.sentCols = t.sentRows = 0; fitOwn(id, t); }
    else if (!t.peer && t.remote) { t.remote = false; fitOwn(id, t); }
  }
}
// (not window focus: bringing the app to the front — e.g. `open` over ssh — is not using it)
['keydown', 'mouseup', 'wheel'].forEach(ev => addEventListener(ev, windowUsed, true));
// the other machine took the size of one of our terminals
window.api.peerTerm.onRemoteSize((id, c, r) => { const t = terms.get(id); if (!t || t.peer) return; t.remote = true; t.cols = c; t.rows = r; fitScaled(t); });
// the exact current screen (+ recent scrollback) of one of our terminals, for a phone or the peer opening it
window.api.termSnap.onRequest((rq, id) => {
  const t = terms.get(id);
  if (!t || !t.ser) return window.api.termSnap.reply(rq, null);
  // wait until everything written before this request is parsed — output after it is streamed separately
  t.term.write('', () => { let out = null; try { out = t.ser.serialize({ scrollback: 1500 }); } catch { } window.api.termSnap.reply(rq, out); });
});
window.api.peerTerm.onInit((id, i) => {
  const t = terms.get(id); if (!t) return;
  t.cols = t.hostCols = i.cols; t.rows = t.hostRows = i.rows; t.term.reset(); t.term.options.fontSize = BASE_FONT; fitScaled(t); t.term.write(i.backlog || '');
  if (t.driving) { t.sentCols = t.sentRows = 0; fitOwn(id, t); }   // reconnected while this window had the size: ask again
});
window.api.peerTerm.onSize((id, i) => {
  const t = terms.get(id); if (!t) return;
  t.hostCols = i.cols; t.hostRows = i.rows;
  if (t.driving && i.cols === t.cols && i.rows === t.rows) return;   // our own request, confirmed
  t.driving = false; t.cols = i.cols; t.rows = i.rows; fitScaled(t);   // the host (or a phone) set another size
});
window.api.peerTerm.onClosed(id => { const t = terms.get(id); if (t) t.attached = false; });   // reconnects on the next telemetry tick
async function newTerm(opts) {
  if (document.body.dataset.view === 'mesh') setView('split');
  const info = await window.api.pty.spawn({ ...opts, cols: 120, rows: 30 });
  const el = document.createElement('div'); el.className = 'term';
  el.innerHTML = `<div class="tlabel"><span class="hdot h-idle"></span><span class="ttl"></span></div><div class="xt"></div>`;
  $('#terms').appendChild(el);
  const term = new Terminal(termOpts());
  const fit = new FitAddon.FitAddon(); term.loadAddon(fit);
  const ser = new SerializeAddon.SerializeAddon(); term.loadAddon(ser);
  term.open(el.querySelector('.xt'));
  wireOsc52(term);
  term.onData(d => window.api.pty.write(info.id, d));
  // ⌘V: Finder-copied files → their paths, screenshot/image → saved PNG path, else plain text.
  // (Ctrl+V still goes straight to Claude Code, which reads image clipboards itself.)
  el.addEventListener('paste', async e => {
    e.preventDefault(); e.stopPropagation();
    const r = await clipboardInsert(); if (r.text) term.paste(r.text);
  }, true);
  wireDrop(el, text => { term.paste(text); term.focus(); });
  term.onResize(({ cols, rows }) => window.api.pty.resize(info.id, cols, rows));
  term.textarea && term.textarea.addEventListener('focus', () => { activeTerm = info.id; markFocus(); });
  terms.set(info.id, { term, fit, ser, el, title: info.title, kind: opts.kind });
  $('#noTerms').style.display = 'none';
  activate(info.id); if (isMini()) miniShow('term');
  return info.id;
}

window.api.pty.onData((id, d) => { const t = terms.get(id); if (t) t.term.write(d); });
window.api.pty.onExit((id) => {
  const t = terms.get(id); if (!t) return;
  t.term.write('\r\n\x1b[90m[process exited — close this tab with ×]\x1b[0m\r\n');
  t.exited = true; renderTabs();
});

function closeTerm(id) {
  const t = terms.get(id); if (!t) return;
  if (t.peer) { peerClosed.add(id); window.api.peerTerm.close(id); }   // a mirror: detach only, the session keeps running on the peer
  else if (!t.exited) window.api.pty.kill(id);
  t.term.dispose(); t.el.remove(); terms.delete(id);
  if (activeTerm === id) activeTerm = [...terms.keys()].pop() || null;
  $('#noTerms').style.display = terms.size ? 'none' : '';
  if (activeTerm) activate(activeTerm); else renderTabs();
  if (isMini() && !terms.size) miniShow('mesh'); else renderRail();
  layoutGrid();
}

function activate(id) {
  activeTerm = id;
  { const s = snap.sessions.find(x => termOf(x) === id); if (s) seen(s.sid); }
  for (const [k, t] of terms) t.el.classList.toggle('on', k === id);
  renderTabs(); markFocus();
  const tab = document.querySelector(`#tabs .tab[data-id="${id}"]`);   // bring the tab into view in a long tab row
  if (tab) tab.scrollIntoView({ inline: 'nearest', block: 'nearest', behavior: isSimple() ? 'auto' : 'smooth' });
  requestAnimationFrame(() => { fitAll(); const t = terms.get(id); if (t) t.term.focus(); });
}
function markFocus() { for (const [k, t] of terms) t.el.classList.toggle('focus', k === activeTerm); }

function renderTabs() {
  $('#tabs').innerHTML = [...terms].map(([id, t]) => `<div class="tab ${id === activeTerm ? 'on' : ''}" data-id="${id}">
    <span class="hdot h-${t.health || (t.exited ? 'error' : 'idle')}"></span><span class="t">${esc(t.title)}</span>${t.local ? '<span class="tag">tmux</span>' : t.peer && snap.peer ? `<span class="tag peer">${esc(snap.peer.label)}</span>` : ''}<span class="x" data-x="${id}">×</span></div>`).join('') +
    `<div class="tab" id="tabNew" title="New session (⌘T)">＋</div>`;
  for (const el of document.querySelectorAll('.tab[data-id]')) el.onclick = e => { if (e.target.dataset.x) closeTerm(e.target.dataset.x); else activate(el.dataset.id); };
  $('#tabNew').onclick = openNew;
  for (const [, t] of terms) { t.el.querySelector('.ttl').textContent = t.title; t.el.querySelector('.hdot').className = 'hdot h-' + (t.health || 'idle'); }
}

function syncTermTitles() {
  let changed = false;
  for (const [id, t] of terms) {
    const s = snap.sessions.find(x => termOf(x) === id);
    const title = s ? s.name : t.kind === 'claude' ? t.title : t.title;
    const health = s ? s.health : t.exited ? 'error' : 'idle';
    if (title !== t.title || health !== t.health) { t.title = title; t.health = health; changed = true; }
  }
  if (changed) renderTabs();
}

function fitAll() {
  for (const [id, t] of terms) {
    if (document.body.dataset.view !== 'grid' && !t.el.classList.contains('on')) continue;
    try { fitTerm(id, t); } catch { }
  }
}
function layoutGrid() {
  const n = Math.max(1, terms.size);
  const cols = n <= 1 ? 1 : n <= 4 ? 2 : 3;
  $('#terms').style.setProperty('--cols', cols);
  requestAnimationFrame(fitAll);
}
new ResizeObserver(() => fitAll()).observe($('#terms'));

// ---------------- views, dock resize, keys ----------------
function setView(v) {
  document.body.dataset.view = v; store.set('view', v);
  document.querySelectorAll('#viewSeg button').forEach(b => b.classList.toggle('on', b.dataset.view === v));
  layoutGrid(); requestAnimationFrame(() => { viz.resize(); fitAll(); });
}
document.querySelectorAll('#viewSeg button').forEach(b => b.onclick = () => setView(b.dataset.view));
setView(store.get('view', 'split'));

(() => {
  const dock = $('#dock'), h = $('#dockHandle');
  const saved = store.get('dockH', null); if (saved) dock.style.height = saved + 'px';
  h.onmousedown = e => {
    const y0 = e.clientY, h0 = dock.offsetHeight;
    const mm = ev => { const nh = Math.max(80, Math.min(window.innerHeight - 200, h0 - (ev.clientY - y0))); dock.style.height = nh + 'px'; viz.resize(); fitAll(); };
    const mu = () => { removeEventListener('mousemove', mm); removeEventListener('mouseup', mu); store.set('dockH', dock.offsetHeight); };
    addEventListener('mousemove', mm); addEventListener('mouseup', mu);
  };
})();

$('#newBtn').onclick = openNew;
// Simple mode: no canvas animation, glow or motion. The mesh becomes a static session board, so a
// Screen Sharing / VNC client (e.g. from the iPhone) only receives pixels that actually changed.
function isSimple() { return document.body.classList.contains('simple'); }
function setSimple(on) {
  document.body.classList.toggle('simple', on); store.set('simple', on);
  $('#simpleBtn').classList.toggle('on', on);
  viz.paused = on;
  for (const [, t] of terms) t.term.options.cursorBlink = !on;
  lastBoard = ''; renderBoard();
  requestAnimationFrame(() => { viz.resize(); fitAll(); });
}
$('#simpleBtn').onclick = () => { setSimple(!isSimple()); toast(isSimple() ? 'Simple mode on: no animation' : 'Simple mode off'); };
setSimple(store.get('simple', false));
$('#bell').classList.toggle('on', bell);
$('#bell').onclick = () => { bell = !bell; store.set('bell', bell); $('#bell').classList.toggle('on', bell); toast(bell ? 'Notifications on' : 'Notifications off'); };
addEventListener('keydown', e => {
  if (!e.metaKey) { if (e.key === 'Escape' && !$('#modal').classList.contains('hidden')) closeModal(); return; }
  if (e.key === '1') { e.preventDefault(); setView('mesh'); }
  else if (e.key === '2') { e.preventDefault(); setView('split'); }
  else if (e.key === '3') { e.preventDefault(); setView('grid'); }
  else if (e.key === 't') { e.preventDefault(); openNew(); }
  else if (e.key === 'b') { e.preventDefault(); openBroadcast(); }
  else if (e.key === 'p') { e.preventDefault(); $('#simpleBtn').click(); }
  else if (/^[4-9]$/.test(e.key)) { const id = [...terms.keys()][+e.key - 4]; if (id) { e.preventDefault(); activate(id); } }
});

// ---------------- notifications ----------------
function notifyTransitions() {
  const live = new Set(snap.sessions.map(s => s.sid));
  for (const sid of doneSids) if (!live.has(sid)) { doneSids.delete(sid); saveDone(); }
  for (const s of snap.sessions) {
    const p = prevHealth.get(s.sid);
    if (p && p !== s.health) {
      const busy = h => h === 'working' || h === 'quiet';
      const watching = document.hasFocus() && (s.sid === selected || (s.owner && s.owner === activeTerm && document.body.dataset.view !== 'mesh'));
      if (busy(p) && (s.health === 'idle' || s.health === 'waiting' || s.health === 'error') && !watching) { doneSids.add(s.sid); saveDone(); renderList(); }
      else if (busy(s.health) && doneSids.has(s.sid)) seen(s.sid);   // picked up again elsewhere
    }
    if (p && p !== s.health && bell) {
      if (p === 'working' && s.health === 'idle') window.api.notify(`✓ ${s.name}`, `Finished · ${fmt$(s.todayUsd)} today`);
      if (s.health === 'error') window.api.notify(`⚠ ${s.name}`, s.lastError || 'API error');
      if (s.health === 'waiting') window.api.notify(`… ${s.name}`, 'Needs your input');
    }
    prevHealth.set(s.sid, s.health);
  }
}

// ---------------- resizable sidebar (drag its right edge; double-click to reset) ----------------
{
  const side = $('#side'), h = $('#sideHandle'), MIN = 200, MAX = 560, DEF = 250;
  const apply = w => { side.style.width = Math.max(MIN, Math.min(MAX, w)) + 'px'; };
  apply(store.get('sideW', DEF));
  h.addEventListener('mousedown', e => {
    e.preventDefault(); const x0 = e.clientX, w0 = side.getBoundingClientRect().width;
    document.body.classList.add('side-resizing');
    const mv = ev => { apply(w0 + ev.clientX - x0); viz.resize(); fitAll && fitAll(); };
    const up = () => { removeEventListener('mousemove', mv); removeEventListener('mouseup', up); document.body.classList.remove('side-resizing');
      store.set('sideW', side.getBoundingClientRect().width); viz.resize(); };
    addEventListener('mousemove', mv); addEventListener('mouseup', up);
  });
  h.addEventListener('dblclick', () => { apply(DEF); store.set('sideW', DEF); viz.resize(); });
}

// ---------------- window size modes: full → compact → mini ----------------
// compact (narrow/short window): the sidebar becomes a slide-over drawer (☰ / ⌘0), the detail panel floats,
//   header drops the model chips and health pills.
// mini (small widget-sized window): a slim bar with plan usage left, the mesh with a smaller star, and a
//   session strip along the bottom — tap a session to fill the window with its terminal, ◎ to go back.
const isMini = () => document.body.dataset.size === 'mini';
let miniPane = store.get('miniPane', 'mesh');
function sizeMode() {
  const w = innerWidth, h = innerHeight;
  const m = w < 700 || h < 470 ? 'mini' : w < 1180 || h < 700 ? 'compact' : 'full';
  if (document.body.dataset.size !== m) {
    document.body.dataset.size = m;
    viz.compact = m === 'mini';
    document.body.classList.remove('drawer');
    if (m === 'mini') { if (miniPane === 'term' && !terms.size) miniPane = 'mesh'; document.body.dataset.mini = miniPane; }
    renderRail();
  }
  requestAnimationFrame(() => { viz.resize(); fitAll(); });
}
function miniShow(p) {
  miniPane = p === 'term' && terms.size ? 'term' : 'mesh'; store.set('miniPane', miniPane);
  document.body.dataset.mini = miniPane; renderRail();
  requestAnimationFrame(() => { viz.resize(); fitAll(); const t = terms.get(activeTerm); if (miniPane === 'term' && t) t.term.focus(); });
}
function renderRail() {
  const u = (snap.usage && snap.usage.limits) || [];
  $('#miniUsage').innerHTML = u.slice(0, 3).map(l => { const left = Math.max(0, 100 - l.used), col = limitColor(left);
    return `<span class="mu" title="${esc(l.label)} · resets in ${fmtReset(l.resetsAt)}"><b style="color:${col}">${Math.round(left)}%</b> ${esc(l.label === 'All models' ? 'week' : l.label === 'Session' ? '5h' : l.label.split(' ')[0].toLowerCase())}<i><em style="width:${left}%;background:${col}"></em></i></span>`; }).join('')
    + `<span class="mu">${fmt$(snap.todayUsd || 0)}</span>`;
  if (!isMini()) return;
  const S = [...snap.sessions].sort((a, b) => (b.owner ? 1 : 0) - (a.owner ? 1 : 0) || (b.health === 'working') - (a.health === 'working'));
  const onTerm = miniPane === 'term';
  const curSid = onTerm ? (snap.sessions.find(x => x.owner === activeTerm) || {}).sid : selected;
  const chips = S.map(s => `<button class="rchip ${s.sid === curSid ? 'on' : ''} ${s.owner ? 'app' : ''}" data-sid="${s.sid}" title="${esc(s.name)} · ${esc(s.health)} · ${fmt$(s.todayUsd)} today${s.owner ? '' : ' · not in this app — opens its controls'}">
      <span class="hdot h-${s.health}"></span><span class="t">${esc(s.name)}</span>${s.health === 'working' ? `<span class="r">${Math.round(s.tps)}</span>` : ''}</button>`).join('');
  const shells = [...terms].filter(([id]) => !snap.sessions.some(s => s.owner === id))
    .map(([id, t]) => `<button class="rchip app ${onTerm && id === activeTerm ? 'on' : ''}" data-pty="${id}"><span class="hdot h-${t.exited ? 'error' : 'idle'}"></span><span class="t">${esc(t.title)}</span></button>`).join('');
  $('#rail').innerHTML = `<button class="rchip home ${onTerm ? '' : 'on'}" id="railMesh" title="Mesh">◎</button>${shells}${chips}<button class="rchip add" id="railNew" title="New session (⌘T)">＋</button>`;
  const on = $('#rail .rchip.on:not(.home)'); if (on && curSid !== renderRail.last) { renderRail.last = curSid; on.scrollIntoView({ inline: 'nearest', block: 'nearest' }); }
  $('#railMesh').onclick = () => miniShow('mesh');
  $('#railNew').onclick = openNew;
  document.querySelectorAll('#rail [data-sid]').forEach(b => b.onclick = () => { const s = snap.sessions.find(x => x.sid === b.dataset.sid); if (s && s.owner && terms.has(s.owner)) select(s.sid); else { miniShow('mesh'); select(b.dataset.sid); } });
  document.querySelectorAll('#rail [data-pty]').forEach(b => b.onclick = () => { miniShow('term'); activate(b.dataset.pty); });
}
$('#drawerBtn').onclick = () => document.body.classList.toggle('drawer');
$('#mesh').addEventListener('pointerdown', () => document.body.classList.remove('drawer'));
addEventListener('keydown', e => { if (e.metaKey && e.key === '0') { e.preventDefault(); $('#drawerBtn').click(); } else if (e.key === 'Escape') document.body.classList.remove('drawer'); });
addEventListener('resize', sizeMode);
sizeMode();

// coming back to the window with a finished session's terminal already in front counts as looking at it
addEventListener('focus', () => { if (document.body.dataset.view === 'mesh') return; const s = snap.sessions.find(x => x.owner && x.owner === activeTerm); if (s) seen(s.sid); });
