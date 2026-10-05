// Slash-command menu for any session: right-click a session row or its mote, click the ⋯ on a
// row, or use the Commands section / "/" autocomplete in the session panel. Built-in Claude Code
// commands plus the user's own skills & commands (~/.claude and the session's project .claude).
// ui: opens an interactive screen or prints output in the terminal → bring that terminal forward.
const BUILTIN_CMDS = [
  { cmd: '/compact', desc: 'Summarize the conversation to free context', group: 'Session' },
  { cmd: '/clear', desc: 'Clear the conversation history', group: 'Session', danger: true },
  { cmd: '/resume', desc: 'Resume a previous conversation', group: 'Session', ui: true },
  { cmd: '/rewind', desc: 'Rewind conversation / code to an earlier point', group: 'Session', ui: true },
  { cmd: '/export', desc: 'Export the conversation', group: 'Session', ui: true },
  { cmd: '/remote-control', desc: 'Continue this session from phone / web', group: 'Session', ui: true },
  { cmd: '/exit', desc: 'Quit this Claude Code session', group: 'Session', danger: true },
  { cmd: '/context', desc: 'Show what is using the context window', group: 'Inspect', ui: true },
  { cmd: '/cost', desc: 'Token usage & cost of this session', group: 'Inspect', ui: true },
  { cmd: '/usage', desc: 'Plan usage limits', group: 'Inspect', ui: true },
  { cmd: '/status', desc: 'Version, model, account, connectivity', group: 'Inspect', ui: true },
  { cmd: '/todos', desc: 'Current todo list', group: 'Inspect', ui: true },
  { cmd: '/doctor', desc: 'Diagnose the Claude Code install', group: 'Inspect', ui: true },
  { cmd: '/help', desc: 'List available commands', group: 'Inspect', ui: true },
  { cmd: '/model', desc: 'Switch model', group: 'Configure', ui: true },
  { cmd: '/skills', desc: 'Browse available skills', group: 'Configure', ui: true },
  { cmd: '/agents', desc: 'Manage sub-agents', group: 'Configure', ui: true },
  { cmd: '/mcp', desc: 'MCP servers & tools', group: 'Configure', ui: true },
  { cmd: '/memory', desc: 'Edit memory files', group: 'Configure', ui: true },
  { cmd: '/config', desc: 'Open settings', group: 'Configure', ui: true },
  { cmd: '/permissions', desc: 'Allow / deny tool rules', group: 'Configure', ui: true },
  { cmd: '/hooks', desc: 'Manage hooks', group: 'Configure', ui: true },
  { cmd: '/plugin', desc: 'Manage plugins', group: 'Configure', ui: true },
  { cmd: '/statusline', desc: 'Set up the status line', group: 'Configure' },
  { cmd: '/add-dir', desc: 'Add a working directory…', group: 'Configure', arg: true },
  { cmd: '/init', desc: 'Create a CLAUDE.md for this project', group: 'Project' },
  { cmd: '/review', desc: 'Review a pull request', group: 'Project' },
  { cmd: '/release-notes', desc: 'What changed in this version', group: 'Project', ui: true },
];
const SESSION_ACTIONS = [
  { cmd: 'Rename…', desc: 'Give this session a name (shows in /resume)', group: 'Manage', action: s => renameSession(s.sid, s.name) },
  { cmd: 'Convert to agent…', desc: 'Continue it as a background agent', group: 'Manage', action: s => convertToAgent(s.sid) },
];
const QUICK_CMDS = ['/compact', '/context', '/cost', '/doctor', '/skills', '/model'];
const customCache = new Map();   // cwd -> [cmds]

async function commandsFor(s) {
  const key = s.cwd || '';
  if (!customCache.has(key)) customCache.set(key, await window.api.listCommands(s.cwd).catch(() => []));
  return SESSION_ACTIONS.concat(BUILTIN_CMDS, customCache.get(key));
}
const canControl = s => s && (s.owner || (s.tty && s.kind !== 'bg'));

async function runCommand(s, item) {
  if (item.action) return item.action(s);
  if (!canControl(s)) return toast('Background session — view only');
  if (item.arg) { select(s.sid); requestAnimationFrame(() => { const ta = $('#cText'); if (ta) { ta.value = item.cmd + ' '; ta.focus(); } }); return; }
  if (item.danger && !await confirmCmd(s, item)) return;
  let r;
  if (s.owner) r = await window.api.ctl.send(s.owner, item.cmd, true);
  else r = await window.api.ext.send(s.tty, item.cmd);
  if (r && (r.ok === false || r.out === 'notfound')) return toast(r.out === 'notfound' ? 'No Terminal.app tab for ' + s.tty : 'Send failed');
  viz.command(s.sid);
  if (item.ui) {                       // show the terminal where the command's screen/output appears
    if (s.owner) { if (document.body.dataset.view === 'mesh') setView('split'); activate(s.owner); }
    else await window.api.ext.focus(s.tty);
  }
  toast(`${item.cmd} → ${s.name}`);
}

function confirmCmd(s, item) {
  return new Promise(res => {
    modal(`<h2>${esc(item.cmd)}</h2><div style="color:var(--dim);line-height:1.6">${esc(item.desc)} in <b style="color:var(--text)">${esc(s.name)}</b>. This can't be undone.</div>
      <div class="btns" style="justify-content:flex-end"><button id="mCancel">Cancel</button><button class="primary danger-solid" id="cfOk">Run ${esc(item.cmd)}</button></div>`);
    $('#mCancel').onclick = () => { closeModal(); res(false); };
    $('#cfOk').onclick = () => { closeModal(); res(true); };
  });
}

// ---------------- floating menu ----------------
let cmdMenu = null;
async function openCmdMenu(sid, x, y) {
  closeCmdMenu();
  const s = snap.sessions.find(q => q.sid === sid); if (!s) return;
  const all = await commandsFor(s);
  const el = document.createElement('div'); el.id = 'cmdMenu';
  el.innerHTML = `<div class="cm-h"><span class="hdot h-${s.health}"></span>${esc(s.name)}</div>
    <input id="cmSearch" placeholder="${canControl(s) ? 'Type to filter…  ↑↓ ↩' : 'Background session — view only'}" ${canControl(s) ? '' : 'disabled'}>
    <div class="cm-list"></div>`;
  document.body.appendChild(el); cmdMenu = { el, s, all, idx: 0, items: [] };
  const render = () => {
    const q = $('#cmSearch').value.trim().toLowerCase().replace(/^\//, '');
    const items = all.filter(c => !q || c.cmd.toLowerCase().includes(q) || c.desc.toLowerCase().includes(q));
    cmdMenu.items = items; cmdMenu.idx = Math.min(cmdMenu.idx, Math.max(0, items.length - 1));
    let g = '', html = '';
    items.forEach((c, i) => {
      if (c.group !== g) { g = c.group; html += `<div class="cm-g">${esc(g)}</div>`; }
      html += `<div class="cm-i ${i === cmdMenu.idx ? 'on' : ''} ${c.danger ? 'danger' : ''}" data-i="${i}"><b>${esc(c.cmd)}</b><span>${esc(c.desc)}</span>${c.ui ? '<em>opens terminal</em>' : ''}</div>`;
    });
    el.querySelector('.cm-list').innerHTML = html || '<div class="cm-g">No match</div>';
    el.querySelectorAll('.cm-i').forEach(n => { n.onclick = () => { const it = cmdMenu.items[+n.dataset.i]; closeCmdMenu(); runCommand(s, it); }; n.onmouseenter = () => { cmdMenu.idx = +n.dataset.i; el.querySelectorAll('.cm-i').forEach(m => m.classList.toggle('on', m === n)); }; });
    el.querySelector('.cm-i.on')?.scrollIntoView({ block: 'nearest' });
  };
  render();
  const r = el.getBoundingClientRect();
  el.style.left = Math.min(x, innerWidth - r.width - 10) + 'px'; el.style.top = Math.min(y, innerHeight - r.height - 10) + 'px';
  const inp = $('#cmSearch'); inp.focus();
  inp.oninput = () => { cmdMenu.idx = 0; render(); };
  inp.onkeydown = e => {
    if (e.key === 'ArrowDown') { e.preventDefault(); cmdMenu.idx = Math.min(cmdMenu.items.length - 1, cmdMenu.idx + 1); render(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); cmdMenu.idx = Math.max(0, cmdMenu.idx - 1); render(); }
    else if (e.key === 'Enter') { e.preventDefault(); const it = cmdMenu.items[cmdMenu.idx]; if (it) { closeCmdMenu(); runCommand(s, it); } }
    else if (e.key === 'Escape') closeCmdMenu();
  };
}
function closeCmdMenu() { if (cmdMenu) { cmdMenu.el.remove(); cmdMenu = null; } }
addEventListener('mousedown', e => { if (cmdMenu && !cmdMenu.el.contains(e.target)) closeCmdMenu(); }, true);

// right-click on rows (delegated, the list re-renders every second) and on motes
$('#list').addEventListener('contextmenu', e => { const row = e.target.closest('.srow'); if (row) { e.preventDefault(); openCmdMenu(row.dataset.sid, e.clientX, e.clientY); } });
$('#list').addEventListener('click', e => { const m = e.target.closest('[data-more]'); if (m) { e.stopPropagation(); const r = m.getBoundingClientRect(); openCmdMenu(m.dataset.more, r.right + 6, r.top); } }, true);
$('#mesh').addEventListener('contextmenu', e => { const h = viz.hit(e); if (h && h !== 'hub') { e.preventDefault(); openCmdMenu(h, e.clientX, e.clientY); } });

// ---------------- panel: quick command chips + "/" autocomplete in the prompt box ----------------
function enhanceCtl() {
  const s = cur(), box = $('#pCtl'); if (!box || !s || box.querySelector('.cmd-chips')) return;
  const sec = box.querySelector('.sec');
  const chips = document.createElement('div'); chips.className = 'cmd-chips';
  chips.innerHTML = `<h3 style="margin-top:12px">Commands</h3><div class="btns">${QUICK_CMDS.map(c => `<button data-c="${c}">${c}</button>`).join('')}<button id="cAll">All commands…</button></div>
    <div class="btns" style="margin-top:6px"><button id="cRename">✎ Rename</button><button id="cAgent">⚙ Convert to agent</button></div>`;
  sec.appendChild(chips);
  chips.querySelectorAll('[data-c]').forEach(b => b.onclick = () => runCommand(cur(), BUILTIN_CMDS.find(c => c.cmd === b.dataset.c)));
  $('#cAll').onclick = e => { const r = e.target.getBoundingClientRect(); openCmdMenu(s.sid, r.left, r.bottom + 6); };
  if (!canControl(s)) chips.querySelectorAll('[data-c],#cAll').forEach(b => b.disabled = true);
  $('#cRename').onclick = () => renameSession(s.sid, s.name);
  $('#cAgent').onclick = () => convertToAgent(s.sid);
  const old = $('#cCompact'); if (old) old.remove();          // folded into the chips
  // "/" autocomplete
  const ta = $('#cText'), ac = document.createElement('div'); ac.className = 'cmd-ac'; ta.after(ac);
  let acItems = [], acIdx = 0;
  const hide = () => { ac.style.display = 'none'; acItems = []; };
  const show = async () => {
    const v = ta.value; if (!/^\/\S*$/.test(v)) return hide();
    const all = await commandsFor(s), q = v.slice(1).toLowerCase();
    acItems = all.filter(c => !c.action && c.cmd.slice(1).toLowerCase().startsWith(q)).slice(0, 8); acIdx = 0;
    if (!acItems.length) return hide();
    ac.innerHTML = acItems.map((c, i) => `<div class="${i === acIdx ? 'on' : ''}" data-i="${i}"><b>${esc(c.cmd)}</b> <span>${esc(c.desc)}</span></div>`).join('');
    ac.style.display = 'block';
    ac.querySelectorAll('div').forEach(d => d.onmousedown = ev => { ev.preventDefault(); ta.value = acItems[+d.dataset.i].cmd; hide(); ta.focus(); });
  };
  ta.addEventListener('input', show);
  ta.addEventListener('blur', () => setTimeout(hide, 100));
  ta.addEventListener('keydown', e => {
    if (!acItems.length) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); acIdx = (acIdx + (e.key === 'ArrowDown' ? 1 : acItems.length - 1)) % acItems.length; ac.querySelectorAll('div').forEach((d, i) => d.classList.toggle('on', i === acIdx)); }
    else if (e.key === 'Tab' || (e.key === 'Enter' && !e.metaKey)) { e.preventDefault(); ta.value = acItems[acIdx].cmd; hide(); }
    else if (e.key === 'Escape') hide();
  }, true);
}
new MutationObserver(() => enhanceCtl()).observe($('#panelBody'), { childList: true, subtree: true });
