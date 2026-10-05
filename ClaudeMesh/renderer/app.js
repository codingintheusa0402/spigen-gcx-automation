const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt$ = v => v >= 100 ? '$' + v.toFixed(0) : '$' + (v || 0).toFixed(2);
const fmtN = v => v >= 1e9 ? (v / 1e9).toFixed(2) + 'B' : v >= 1e6 ? (v / 1e6).toFixed(1) + 'M' : v >= 1e3 ? (v / 1e3).toFixed(1) + 'k' : String(Math.round(v || 0));
const ago = s => s < 60 ? Math.round(s) + 's' : s < 3600 ? Math.round(s / 60) + 'm' : (s / 3600).toFixed(1) + 'h';
const hhmm = t => t ? new Date(t).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }) : '';
const store = { get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { } } };

let snap = { sessions: [], models: [], ptys: [] };
let selected = null;
let bell = store.get('bell', true);
const prevHealth = new Map();
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
  renderHeader(); renderList(); renderToday(); renderPanel(); syncTermTitles(); notifyTransitions();
});

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
    <span class="hpill" ${errs ? 'style="border-color:#6b2a2a"' : ''}>API errors 1h <i>${errs}</i></span>`;
}

function renderList() {
  $('#sessCount').textContent = `· ${snap.sessions.length}`;
  $('#list').innerHTML = snap.sessions.map(s => `<div class="srow ${s.sid === selected ? 'sel' : ''}" data-sid="${s.sid}">
    <div class="n"><span class="hdot h-${s.health}"></span><span class="t">${esc(s.name)}</span>${s.owner ? '<span class="tag">app</span>' : s.kind === 'bg' ? '<span class="tag">bg</span>' : ''}<span class="more" data-more="${s.sid}" title="Commands (or right-click)">⋯</span></div>
    <div class="m"><span style="color:${FAMILY_COLOR[s.family]}">${esc((s.model || '—').replace('claude-', ''))}</span><span>${fmt$(s.todayUsd)}</span><span>${s.health === 'working' ? Math.round(s.tps) + ' tok/s' : ago(s.sinceWrite)}</span></div>
  </div>`).join('');
  for (const el of document.querySelectorAll('.srow')) el.onclick = () => select(el.dataset.sid);
}

function renderToday() {
  const tot = snap.todayUsd || 0;
  const burn = snap.models.reduce((a, m) => a + m.liveUsdHr, 0);
  $('#today').innerHTML = `<div class="big">${fmt$(tot)}</div><div style="margin-bottom:10px">burn rate ${fmt$(burn)}/h · all transcripts</div>` +
    snap.models.filter(m => m.todayUsd > 0).sort((a, b) => b.todayUsd - a.todayUsd).map(m => `
    <div style="display:flex;justify-content:space-between"><span>${m.family}</span><span>${fmt$(m.todayUsd)}</span></div>
    <div class="bar"><div style="width:${tot ? m.todayUsd / tot * 100 : 0}%;background:${FAMILY_COLOR[m.family]}"></div></div>`).join('');
}

// ---------------- detail panel ----------------
function select(sid) {
  selected = sid; viz.selected = sid;
  $('#panel').classList.toggle('hidden', !sid);
  $('#panelBody').innerHTML = sid ? '<div id="pStatic"></div><div id="pCtl"></div>' : '';
  if (sid) { renderCtl(); renderPanel(); }
  renderList();
  requestAnimationFrame(() => viz.resize());
}

function cur() { return snap.sessions.find(s => s.sid === selected); }

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
    <span>pid ${s.pid}</span><span>${s.tty ? s.tty.replace('/dev/', '') : 'no tty'}</span><span>${s.owner ? 'in-app' : s.kind === 'bg' ? 'background' : 'Terminal.app'}</span>
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
  drawSpark($('#spark'), s.spark, col);
}

function drawSpark(c, data, col) {
  const d = devicePixelRatio, w = c.clientWidth, h = c.clientHeight;
  c.width = w * d; c.height = h * d;
  const x = c.getContext('2d'); x.scale(d, d);
  const max = Math.max(10, ...data);
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
  const owned = !!s.owner, ext = !owned && s.tty && s.kind !== 'bg';
  box.innerHTML = `<div class="sec ctl"><h3>Control</h3>
    <textarea id="cText" placeholder="${owned || ext ? 'Type a prompt… drop files or ⌘V images to attach (⌘↩ to send)' : 'This session has no terminal the app can reach'}"></textarea>
    <div class="btns">
      <button class="primary" id="cSend">Send</button>
      <button id="cEsc" title="Send Esc (interrupt the current turn)">Interrupt</button>
      <button id="cCompact">/compact</button>
      <button id="cFocus">${owned ? 'Show terminal' : 'Focus in Terminal'}</button>
      ${owned ? '<button class="danger" id="cKill">Kill</button>' : ''}
    </div>
    <div class="note">${owned ? 'Runs inside Claude Mesh — keystrokes go straight to its terminal.' :
      ext ? 'Lives in a Terminal.app window. Prompts are typed into that tab via AppleScript (first use asks for Automation permission). Multi-line text is sent as one line.' :
      'Background job — view only.'}</div></div>`;
  const send = async (text) => {
    const s = cur(); if (!s || !text) return;
    let r;
    if (s.owner) r = await window.api.ctl.send(s.owner, text, true);
    else if (s.tty) r = await window.api.ext.send(s.tty, text);
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
    if (s.owner) window.api.pty.write(s.owner, '\x1b'); else if (s.tty) await window.api.ext.interrupt(s.tty);
    viz.command(s.sid); toast('Interrupt sent');
  };
  $('#cFocus').onclick = async () => {
    const s = cur(); if (!s) return;
    if (s.owner) { if (document.body.dataset.view === 'mesh') setView('split'); activate(s.owner); }
    else if (s.tty) { const r = await window.api.ext.focus(s.tty); if (r.out === 'notfound') toast('No Terminal.app tab for ' + s.tty); }
  };
  if (owned) $('#cKill').onclick = () => { window.api.pty.kill(s.owner); toast('Killed'); };
  if (!owned && !ext) for (const id of ['#cSend', '#cEsc', '#cCompact', '#cFocus', '#cText']) $(id).disabled = true;
}

// ---------------- broadcast (click the core) ----------------
function openBroadcast() {
  const targets = snap.sessions.filter(s => s.owner || (s.tty && s.kind !== 'bg'));
  modal(`<h2>Broadcast to sessions</h2>
    <div class="blist">${targets.map(s => `<label class="chk"><input type="checkbox" data-sid="${s.sid}" checked><span class="hdot h-${s.health}"></span>${esc(s.name)} <span style="color:var(--dim)">${s.health}</span></label>`).join('') || '<div class="note">No controllable sessions.</div>'}</div>
    <label>Message</label><textarea id="bText" style="width:100%;min-height:80px" placeholder="e.g. /compact  ·  give me a one-line status"></textarea>
    <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary" id="bSend">Send to selected</button></div>`);
  $('#bSend').onclick = async () => {
    const text = $('#bText').value.trim(); if (!text) return;
    const sids = [...document.querySelectorAll('.blist input:checked')].map(i => i.dataset.sid);
    for (const sid of sids) {
      const s = snap.sessions.find(x => x.sid === sid); if (!s) continue;
      if (s.owner) await window.api.ctl.send(s.owner, text, true); else await window.api.ext.send(s.tty, text);
      viz.command(sid);
    }
    closeModal(); toast(`Broadcast to ${sids.length} session(s)`);
  };
}

// ---------------- new session ----------------
async function openNew() {
  const home = await window.api.home();
  const last = store.get('newOpts', { kind: 'claude', cwd: home, flags: '--dangerously-skip-permissions', model: '' });
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

function modal(html) { $('#modalCard').innerHTML = html; $('#modal').classList.remove('hidden'); const c = $('#mCancel'); if (c) c.onclick = closeModal; }
function closeModal() { $('#modal').classList.add('hidden'); }
$('#modal').onclick = e => { if (e.target.id === 'modal') closeModal(); };

// ---------------- paste / drop of files & images ----------------
// Paths are backslash-escaped exactly like Terminal.app does on drag-and-drop, so Claude Code
// turns image paths into [Image #n] attachments and other files into @-mentionable paths.
const shq = p => p.replace(/([\s\\'"()&;|<>$`!*?#~{}\[\]])/g, '\\$1');
const droppedPaths = e => [...(e.dataTransfer?.files || [])].map(f => window.api.pathForFile(f)).filter(Boolean);
async function clipboardInsert() {
  const c = await window.api.clipRead();
  if (c.kind === 'files') { if (c.image) toast('Image pasted as ' + c.paths[0].split('/').pop()); return { text: c.paths.map(shq).join(' ') + ' ', files: true }; }
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
async function newTerm(opts) {
  if (document.body.dataset.view === 'mesh') setView('split');
  const info = await window.api.pty.spawn({ ...opts, cols: 120, rows: 30 });
  const el = document.createElement('div'); el.className = 'term';
  el.innerHTML = `<div class="tlabel"><span class="hdot h-idle"></span><span class="ttl"></span></div><div class="xt"></div>`;
  $('#terms').appendChild(el);
  const term = new Terminal({ 
    fontFamily: '"SF Mono", Menlo, monospace', fontSize: 12.5, lineHeight: 1.15, cursorBlink: true, allowProposedApi: true,
    macOptionIsMeta: true, scrollback: 20000,
    theme: { background: '#06051a', foreground: '#e6e9f2', cursor: '#a78bfa', selectionBackground: '#3b4a7a88',
      black: '#1b1f2a', brightBlack: '#555e78', blue: '#7aa2f7', brightBlue: '#9ab8ff', magenta: '#bb9af7', cyan: '#2dd4bf', green: '#34d399', red: '#f87171', yellow: '#fbbf24' },
  });
  const fit = new FitAddon.FitAddon(); term.loadAddon(fit);
  term.open(el.querySelector('.xt'));
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
  terms.set(info.id, { term, fit, el, title: info.title, kind: opts.kind });
  $('#noTerms').style.display = 'none';
  activate(info.id);
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
  if (!t.exited) window.api.pty.kill(id);
  t.term.dispose(); t.el.remove(); terms.delete(id);
  if (activeTerm === id) activeTerm = [...terms.keys()].pop() || null;
  $('#noTerms').style.display = terms.size ? 'none' : '';
  if (activeTerm) activate(activeTerm); else renderTabs();
  layoutGrid();
}

function activate(id) {
  activeTerm = id;
  for (const [k, t] of terms) t.el.classList.toggle('on', k === id);
  renderTabs(); markFocus();
  requestAnimationFrame(() => { fitAll(); const t = terms.get(id); if (t) t.term.focus(); });
}
function markFocus() { for (const [k, t] of terms) t.el.classList.toggle('focus', k === activeTerm); }

function renderTabs() {
  $('#tabs').innerHTML = [...terms].map(([id, t]) => `<div class="tab ${id === activeTerm ? 'on' : ''}" data-id="${id}">
    <span class="hdot h-${t.health || (t.exited ? 'error' : 'idle')}"></span><span class="t">${esc(t.title)}</span><span class="x" data-x="${id}">×</span></div>`).join('') +
    `<div class="tab" id="tabNew" title="New session (⌘T)">＋</div>`;
  for (const el of document.querySelectorAll('.tab[data-id]')) el.onclick = e => { if (e.target.dataset.x) closeTerm(e.target.dataset.x); else activate(el.dataset.id); };
  $('#tabNew').onclick = openNew;
  for (const [, t] of terms) { t.el.querySelector('.ttl').textContent = t.title; t.el.querySelector('.hdot').className = 'hdot h-' + (t.health || 'idle'); }
}

function syncTermTitles() {
  let changed = false;
  for (const [id, t] of terms) {
    const s = snap.sessions.find(x => x.owner === id);
    const title = s ? s.name : t.kind === 'claude' ? t.title : t.title;
    const health = s ? s.health : t.exited ? 'error' : 'idle';
    if (title !== t.title || health !== t.health) { t.title = title; t.health = health; changed = true; }
  }
  if (changed) renderTabs();
}

function fitAll() {
  for (const [, t] of terms) {
    if (document.body.dataset.view !== 'grid' && !t.el.classList.contains('on')) continue;
    try { t.fit.fit(); } catch { }
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
function setPixel(on) {
  document.body.classList.toggle('pixel', on); store.set('pixel', on);
  $('#pixelBtn').classList.toggle('on', on);
  viz.setPixel(on);
  for (const [, t] of terms) t.term.options.fontFamily = on ? 'Monaco, Menlo, monospace' : '"SF Mono", Menlo, monospace';
  requestAnimationFrame(fitAll);
}
$('#pixelBtn').onclick = () => { setPixel(!document.body.classList.contains('pixel')); toast(document.body.classList.contains('pixel') ? 'Pixel mode on — lightweight' : 'Pixel mode off'); };
setPixel(store.get('pixel', false));
$('#bell').classList.toggle('on', bell);
$('#bell').onclick = () => { bell = !bell; store.set('bell', bell); $('#bell').classList.toggle('on', bell); toast(bell ? 'Notifications on' : 'Notifications off'); };
addEventListener('keydown', e => {
  if (!e.metaKey) { if (e.key === 'Escape' && !$('#modal').classList.contains('hidden')) closeModal(); return; }
  if (e.key === '1') { e.preventDefault(); setView('mesh'); }
  else if (e.key === '2') { e.preventDefault(); setView('split'); }
  else if (e.key === '3') { e.preventDefault(); setView('grid'); }
  else if (e.key === 't') { e.preventDefault(); openNew(); }
  else if (e.key === 'b') { e.preventDefault(); openBroadcast(); }
  else if (e.key === 'p') { e.preventDefault(); $('#pixelBtn').click(); }
  else if (/^[4-9]$/.test(e.key)) { const id = [...terms.keys()][+e.key - 4]; if (id) { e.preventDefault(); activate(id); } }
});

// ---------------- notifications ----------------
function notifyTransitions() {
  for (const s of snap.sessions) {
    const p = prevHealth.get(s.sid);
    if (p && p !== s.health && bell) {
      if (p === 'working' && s.health === 'idle') window.api.notify(`✓ ${s.name}`, `Finished · ${fmt$(s.todayUsd)} today`);
      if (s.health === 'error') window.api.notify(`⚠ ${s.name}`, s.lastError || 'API error');
      if (s.health === 'waiting') window.api.notify(`… ${s.name}`, 'Needs your input');
    }
    prevHealth.set(s.sid, s.health);
  }
}
