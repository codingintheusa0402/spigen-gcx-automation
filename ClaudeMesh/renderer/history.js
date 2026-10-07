// All sessions (the same set `claude --resume` lists), background agents, rename and
// convert-to-agent. Clicking a past session resumes it in an in-app terminal straight away,
// with --dangerously-skip-permissions (never stops to ask) and --remote-control (reachable from phone/web).
const RESUME_FLAGS = '--dangerously-skip-permissions --remote-control';
let hist = [], histShown = 60;
const projName = cwd => !cwd ? '?' : cwd === HOME_DIR ? '~' : cwd.split('/').filter(Boolean).pop();
let HOME_DIR = ''; window.api.home().then(h => { HOME_DIR = h; });
const histTitle = h => h.title || h.aiTitle || h.first || h.last || h.sid.slice(0, 8);
const agoT = t => ago((Date.now() - t) / 1000);

async function loadHistory() {
  try { hist = await window.api.history(); } catch { }
  renderHistory();
}
function renderHistory() {
  const q = $('#hSearch').value.trim().toLowerCase(), auto = $('#hAuto').checked;
  const liveSids = new Set(snap.sessions.map(s => s.sid)), agentSids = new Set((snap.agents || []).map(a => a.sessionId));
  const rows = hist.filter(h => (auto || h.entry !== 'sdk-cli') &&
    (!q || [h.title, h.aiTitle, h.first, h.last, h.cwd, h.branch, h.sid].some(v => v && v.toLowerCase().includes(q))));
  $('#hCount').textContent = '· ' + rows.length;
  $('#hList').innerHTML = rows.slice(0, histShown).map(h => {
    const live = liveSids.has(h.sid), ag = agentSids.has(h.sid);
    return `<div class="hrow ${live ? 'live' : ''}" data-hsid="${h.sid}" title="${esc(histTitle(h))}\n${esc(h.cwd)}${h.last ? '\nlast: ' + esc(h.last) : ''}">
      <div class="n"><span class="t">${esc(histTitle(h))}</span>${live ? '<span class="tag">live</span>' : ag ? '<span class="tag">agent</span>' : ''}<span class="more" data-hmore="${h.sid}">⋯</span></div>
      <div class="m">${esc(projName(h.cwd))}${h.branch && h.branch !== 'HEAD' ? ' · ' + esc(h.branch) : ''} · ${agoT(h.mtime)} · ${h.size > 1e6 ? (h.size / 1e6).toFixed(1) + 'MB' : Math.round(h.size / 1e3) + 'kB'}</div></div>`;
  }).join('') + (rows.length > histShown ? `<div class="hmore" id="hMore">Show ${Math.min(100, rows.length - histShown)} more…</div>` : '');
  const m = $('#hMore'); if (m) m.onclick = () => { histShown += 100; renderHistory(); };
}
$('#hSearch').oninput = () => { histShown = 60; renderHistory(); };
$('#hAuto').onchange = () => { store.set('hAuto', $('#hAuto').checked); renderHistory(); };
$('#hAuto').checked = store.get('hAuto', false);

// resume: live → select it; agent → attach; otherwise open a terminal running --resume
async function resumeSession(sid) {
  const live = snap.sessions.find(s => s.sid === sid);
  if (live) { select(sid); if (live.owner) { if (document.body.dataset.view === 'mesh') setView('split'); activate(live.owner); } return; }
  const ag = (snap.agents || []).find(a => a.sessionId === sid);
  if (ag) return attachAgent(ag);
  const h = hist.find(x => x.sid === sid); if (!h) return;
  const id = await newTerm({ kind: 'claude', cwd: h.cwd, flags: `${RESUME_FLAGS} --resume ${sid}`, title: histTitle(h).slice(0, 40) });
  autoTrust.set(id, { until: Date.now() + 20000, buf: '' });
  toast('Resuming ' + histTitle(h).slice(0, 50));
}
// You already worked in this folder (that's where the session came from), so the one-time
// "Do you trust this folder?" screen is answered for you: ↓ to "Yes, I trust this folder", Enter.
const autoTrust = new Map();
window.api.pty.onData((id, d) => {
  const a = autoTrust.get(id); if (!a) return;
  if (Date.now() > a.until) return autoTrust.delete(id);
  a.buf = (a.buf + d).slice(-4000);
  if (/Yes,Itrustthisfolder/.test(a.buf.replace(/\x1b\[[0-9;?]*[A-Za-z]/g, '').replace(/\s/g, ''))) {   // Ink draws spaces as cursor moves
    autoTrust.delete(id);
    setTimeout(() => { window.api.pty.write(id, '\x1b[B'); setTimeout(() => window.api.pty.write(id, '\r'), 700); }, 700);
  }
});
function attachAgent(a) {
  return newTerm({ kind: 'claude', cwd: a.cwd, flags: `attach ${a.id}`, title: '⚙ ' + (agentName(a)) });
}

$('#hList').addEventListener('click', e => {
  const m = e.target.closest('[data-hmore]');
  if (m) { e.stopPropagation(); const r = m.getBoundingClientRect(); return openSessMenu(m.dataset.hmore, r.right + 6, r.top); }
  const row = e.target.closest('.hrow'); if (row) resumeSession(row.dataset.hsid);
});
$('#hList').addEventListener('contextmenu', e => { const row = e.target.closest('.hrow'); if (row) { e.preventDefault(); openSessMenu(row.dataset.hsid, e.clientX, e.clientY); } });

// ---------------- small action menu for history / agent rows ----------------
function openSessMenu(sid, x, y, agent) {
  closeCmdMenu();
  const h = hist.find(q => q.sid === sid), live = snap.sessions.find(s => s.sid === sid);
  const ag = agent || (snap.agents || []).find(a => a.sessionId === sid);
  const items = ag ? [
    ['Attach (open in terminal)', () => attachAgent(ag)],
    ['Show logs', () => showAgentLogs(ag)],
    ['Rename…', () => renameSession(ag.sessionId, ag.name)],
    ['Stop agent', () => stopAgent(ag), true],
  ] : [
    [live ? 'Go to live session' : 'Resume here', () => resumeSession(sid)],
    ['Rename…', () => renameSession(sid, h ? histTitle(h) : live && live.name)],
    ['Convert to background agent…', () => convertToAgent(sid)],
  ];
  const el = document.createElement('div'); el.id = 'cmdMenu';
  el.innerHTML = `<div class="cm-h">${esc(ag ? (ag.name || ag.id) : h ? histTitle(h) : live ? live.name : sid.slice(0, 8))}</div><div class="cm-list">` +
    items.map(([l, , d], i) => `<div class="cm-i ${d ? 'danger' : ''}" data-i="${i}"><b>${esc(l)}</b></div>`).join('') + '</div>';
  document.body.appendChild(el); cmdMenu = { el };
  const r = el.getBoundingClientRect();
  el.style.left = Math.min(x, innerWidth - r.width - 10) + 'px'; el.style.top = Math.min(y, innerHeight - r.height - 10) + 'px';
  el.querySelectorAll('.cm-i').forEach(n => n.onclick = () => { closeCmdMenu(); items[+n.dataset.i][1](); });
}

// ---------------- rename ----------------
function renameSession(sid, current) {
  modal(`<h2>Rename session</h2><label>Name</label><input type="text" id="rnName" value="${esc(current || '')}">
    <div class="note" style="color:var(--dim);font-size:12px;margin-top:8px">Shows in the /resume picker and here.</div>
    <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary" id="rnGo">Rename</button></div>`);
  const inp = $('#rnName'); inp.focus(); inp.select();
  const go = async () => {
    const name = inp.value.trim().replace(/\s+/g, ' '); if (!name) return;
    closeModal();
    const live = snap.sessions.find(s => s.sid === sid);
    let r;
    if (live && canControl(live)) {          // a running session renames itself, so its prompt box/title update too
      r = live.owner ? await window.api.ctl.send(live.owner, '/rename ' + name, true) : await window.api.ext.send(live.tty, '/rename ' + name);
      if (r && (r.ok === false || r.out === 'notfound')) r = await window.api.rename(sid, name);
    } else r = await window.api.rename(sid, name);
    if (r && r.ok === false) return toast('Rename failed: ' + (r.err || ''));
    const h = hist.find(x => x.sid === sid); if (h) h.title = name;
    renderHistory(); toast('Renamed → ' + name);
  };
  $('#rnGo').onclick = go; inp.onkeydown = e => { if (e.key === 'Enter') go(); };
}

// ---------------- convert to background agent ----------------
function convertToAgent(sid) {
  const live = snap.sessions.find(s => s.sid === sid), h = hist.find(x => x.sid === sid);
  const name = live ? live.name : h ? histTitle(h) : sid.slice(0, 8), cwd = live ? live.cwd : h && h.cwd;
  const note = live && live.owner ? 'It is open in an in-app terminal — that terminal will be closed and the session continues in the background under the same id.'
    : live ? 'It is still running in Terminal.app, so Claude starts a background <b>copy</b> (the Terminal one keeps going). Exit it first to move it instead.'
    : 'The session continues in the background under the same id, with permissions skipped. It shows under <b>Agents</b>; attach any time.';
  modal(`<h2>Convert to agent</h2>
    <div style="color:var(--dim);line-height:1.6;font-size:13px"><b style="color:var(--text)">${esc(name)}</b><br>${note}</div>
    <label>Task for the agent (optional — leave empty to just park it in the background)</label>
    <textarea id="cvTask" placeholder="e.g. keep going with the remaining items and report when done"></textarea>
    <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary" id="cvGo">Convert</button></div>`);
  $('#cvTask').focus();
  $('#cvGo').onclick = async () => {
    const task = $('#cvTask').value; closeModal();
    if (live && live.owner) { closeTerm(live.owner); await new Promise(r => setTimeout(r, 1500)); }
    toast('Starting agent…');
    const r = await window.api.agents.convert(sid, cwd, task);
    if (!r.ok) return modal(`<h2>Convert failed</h2><pre class="logs">${esc(r.err || r.out)}</pre><div class="btns" style="justify-content:flex-end"><button id="mCancel">Close</button></div>`);
    toast((r.out.split('\n').pop() || 'Agent started').slice(0, 120));
    snap.agents = await window.api.agents.list(); renderAgents();
  };
}

// ---------------- agents list ----------------
const agentName = a => a.name || (hist.find(h => h.sid === a.sessionId) && histTitle(hist.find(h => h.sid === a.sessionId))) || a.id;
function renderAgents() {
  const A = snap.agents || [];
  $('#agentsBox').style.display = A.length ? '' : 'none';
  $('#agCount').textContent = '· ' + A.length;
  $('#agList').innerHTML = A.map(a => `<div class="arow" data-aid="${a.id}" title="click to attach · ${esc(a.cwd)}">
    <div class="n"><span class="hdot ast-${esc(a.state || a.status || '')}"></span><span class="t">${esc(agentName(a))}</span><span class="more" data-amore="${a.id}">⋯</span></div>
    <div class="m">${esc(a.state || a.status || '')} · ${esc(projName(a.cwd))} · ${a.startedAt ? agoT(a.startedAt) : ''}</div></div>`).join('');
}
$('#agList').addEventListener('click', e => {
  const A = snap.agents || [], m = e.target.closest('[data-amore]');
  if (m) { e.stopPropagation(); const a = A.find(x => x.id === m.dataset.amore), r = m.getBoundingClientRect(); return openSessMenu(a.sessionId, r.right + 6, r.top, a); }
  const row = e.target.closest('.arow'); if (row) attachAgent(A.find(x => x.id === row.dataset.aid));
});
$('#agList').addEventListener('contextmenu', e => { const row = e.target.closest('.arow'); if (row) { e.preventDefault(); const a = (snap.agents || []).find(x => x.id === row.dataset.aid); openSessMenu(a.sessionId, e.clientX, e.clientY, a); } });
async function showAgentLogs(a) {
  modal(`<h2>${esc(agentName(a))} · logs</h2><pre class="logs">loading…</pre><div class="btns" style="justify-content:flex-end"><button id="mCancel">Close</button></div>`);
  const r = await window.api.agents.logs(a.id); const pre = document.querySelector('.card pre.logs'); if (pre) { pre.textContent = r.out || r.err || '(no output)'; pre.scrollTop = pre.scrollHeight; }
}
async function stopAgent(a) {
  if (!await confirmCmd({ name: agentName(a) }, { cmd: 'Stop agent', desc: 'Stop the background agent' })) return;
  const r = await window.api.agents.stop(a.id); toast(r.ok ? 'Stopped ' + (agentName(a)) : 'Stop failed');
  snap.agents = await window.api.agents.list(); renderAgents();
}

// keep in sync with telemetry; history rescans every 20s (cached per file, only changed ones re-read)
// re-render only when something visible changed (live set, agents), not on every 1 s tick
let lastHistKey = '', lastAgKey = '';
window.api.onTelemetry(() => {
  const ak = JSON.stringify(snap.agents || []); if (ak !== lastAgKey) { lastAgKey = ak; renderAgents(); }
  const hk = snap.sessions.map(s => s.sid).join() + '|' + (snap.agents || []).map(a => a.sessionId).join() + '|' + hist.length + '|' + Math.floor(Date.now() / 60000);
  if (hist.length && hk !== lastHistKey && !$('#hList').matches(':hover')) { lastHistKey = hk; renderHistory(); }
});
loadHistory(); setInterval(loadHistory, 20000);

// ---------------- iPhone access (phone-server.js) ----------------
// The phone asks the Mac to resume a session → do exactly what a click in "All sessions" does.
window.api.phone.onResume(async sid => { if (!hist.length) await loadHistory(); resumeSession(sid); });

async function openPhone() {
  const st = await window.api.phone.status();
  const qrs = await Promise.all(st.urls.map(u => window.api.phone.qr(u.url)));
  modal(`<h2>GCX Mesh on iPhone</h2>
    <label class="chk" style="margin:4px 0 10px"><input type="checkbox" id="phOn" ${st.enabled ? 'checked' : ''}> Allow phone access <span style="color:var(--dim)">(off = nothing listens)</span></label>
    ${st.enabled ? (st.urls.length ? `<div class="ph-qrs">${st.urls.map((u, i) => `<div class="ph-qr"><img src="${qrs[i]}" alt=""><b>${esc(u.kind)}</b><code>${esc(u.url.replace(/\?t=.*/, ''))}</code><button data-copy="${esc(u.url)}">Copy link</button></div>`).join('')}</div>
      <div style="color:var(--dim);font-size:12px;line-height:1.6;margin-top:10px">1. Scan with the iPhone camera → it opens in Safari.<br>2. Share → <b style="color:var(--text)">Add to Home Screen</b> — or scan it inside the native GCX Mesh iPhone app.<br>
      Tailscale works from anywhere (phone needs Tailscale on). The QR holds a private pairing key — don't share it. ${st.clients ? `<br><b style="color:var(--ok)">${st.clients} phone(s) connected</b>` : ''}</div>` : '<div style="color:var(--dim)">No network address found — connect to Wi-Fi or Tailscale.</div>')
      : '<div style="color:var(--dim);font-size:13px">Turn on to show the pairing QR code.</div>'}
    <div class="btns" style="justify-content:space-between;margin-top:14px">${st.enabled ? '<button id="phRegen">New pairing key</button>' : '<span></span>'}<button id="mCancel">Close</button></div>`);
  document.querySelectorAll('[data-copy]').forEach(b => b.onclick = () => { navigator.clipboard.writeText(b.dataset.copy); toast('Pairing link copied — paste it in the iPhone app (or AirDrop it)'); });
  $('#phOn').onchange = async e => { await window.api.phone.set(e.target.checked); openPhone(); };
  const rg = $('#phRegen'); if (rg) rg.onclick = async () => { await window.api.phone.regen(); toast('New key — paired phones must scan again'); openPhone(); };
}
$('#phoneBtn').onclick = openPhone;
