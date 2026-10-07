// Schedules: one view of the GCX scheduled jobs on the Mac (launchd) and the 24/7 server (cron
// on gcx-server, over Tailscale SSH). The catalog (name, what it does, skills, command, state
// that must travel with it) is ServerBootstrap/jobs.json in the GCX repo.
// Schedules are {every: minutes} or {at: ["HH:MM"…], days: [1..7 Mon=1] | null (every day)}.
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile, spawn } = require('child_process');

const HOME = os.homedir();
const REPO = path.join(HOME, 'Desktop', 'GCX');
const CATALOG = path.join(REPO, 'ServerBootstrap', 'jobs.json');
const AGENTS = path.join(HOME, 'Library', 'LaunchAgents');
const PARKED = path.join(AGENTS, 'disabled-moved-to-server');
const SERVER = 'kevinkim@gcx-server', SRV_HOME = '/home/kevinkim', SRV_REPO = SRV_HOME + '/Desktop/GCX';
const UID = process.getuid ? process.getuid() : 501;
const label = id => 'com.spigen.gcx.' + id;

const run = (cmd, args, input, timeout = 30000) => new Promise(res => {
  const p = spawn(cmd, args); let out = '', err = '';
  const t = setTimeout(() => { try { p.kill(); } catch { } }, timeout);
  p.stdout.on('data', d => out += d); p.stderr.on('data', d => err += d);
  p.on('close', code => { clearTimeout(t); res({ ok: code === 0, out, err: err.trim(), code }); });
  p.on('error', e => { clearTimeout(t); res({ ok: false, out: '', err: e.message }); });
  if (input != null) p.stdin.end(input); else p.stdin.end();
});
const ssh = (cmd, input, timeout) => run('/usr/bin/ssh', ['-o', 'BatchMode=yes', '-o', 'ConnectTimeout=8', SERVER, cmd], input, timeout);
const catalog = () => JSON.parse(fs.readFileSync(CATALOG, 'utf8'));
const pad = n => String(n).padStart(2, '0');
const sortSched = s => s.at ? { at: [...new Set(s.at)].sort(), days: s.days && s.days.length && s.days.length < 7 ? [...new Set(s.days)].sort() : null } : { every: Math.max(1, Math.round(s.every)) };

// ---------------- launchd (Mac) ----------------
async function readPlist(file) {
  const r = await run('/usr/bin/plutil', ['-convert', 'json', '-o', '-', file]);
  return r.ok ? JSON.parse(r.out) : null;
}
function schedFromPlist(p) {
  if (p.StartInterval) return { every: Math.round(p.StartInterval / 60) };
  const cal = [].concat(p.StartCalendarInterval || []);
  if (!cal.length) return null;
  const at = [...new Set(cal.map(c => pad(c.Hour ?? 0) + ':' + pad(c.Minute ?? 0)))];
  const wd = cal.filter(c => c.Weekday != null).map(c => c.Weekday === 0 ? 7 : c.Weekday);
  return sortSched({ at, days: wd.length ? wd : null });
}
function plistFor(job, s) {
  const p = { Label: label(job.id), ProgramArguments: ['/opt/homebrew/bin/python3', ...job.cmd.replace(/^python3\s+/, '').split(/\s+/).map((a, i) => i === 0 ? path.join(REPO, job.dir, a) : a)],
    RunAtLoad: !!job.runAtLoad, StandardOutPath: path.join(REPO, job.log + '.out.log'), StandardErrorPath: path.join(REPO, job.log + '.err.log') };
  if (s.every) p.StartInterval = s.every * 60;
  else p.StartCalendarInterval = s.at.flatMap(t => { const [h, m] = t.split(':').map(Number); return (s.days || [null]).map(d => d == null ? { Hour: h, Minute: m } : { Weekday: d % 7, Hour: h, Minute: m }); });
  return p;
}
async function writePlist(file, obj) {
  const tmp = path.join(os.tmpdir(), 'mesh-plist-' + Date.now() + '.json');
  fs.writeFileSync(tmp, JSON.stringify(obj));
  const r = await run('/usr/bin/plutil', ['-convert', 'xml1', '-o', file, tmp]); fs.unlinkSync(tmp); return r;
}
const bootout = id => run('/bin/launchctl', ['bootout', `gui/${UID}/${label(id)}`]);
const bootstrap = file => run('/bin/launchctl', ['bootstrap', `gui/${UID}`, file]);

async function macState(job, loaded) {
  const live = path.join(AGENTS, label(job.id) + '.plist'), parked = path.join(PARKED, label(job.id) + '.plist');
  const file = fs.existsSync(live) ? live : fs.existsSync(parked) ? parked : null;
  const p = file && await readPlist(file);
  let last = null; try { last = fs.statSync(path.join(REPO, job.log + '.out.log')).mtimeMs; } catch { }
  return { installed: file === live, parked: file === parked, loaded: loaded.has(label(job.id)), schedule: p ? schedFromPlist(p) : null, last };
}
async function macSet(job, s, enabled) {
  const live = path.join(AGENTS, label(job.id) + '.plist'), parked = path.join(PARKED, label(job.id) + '.plist');
  await bootout(job.id);
  if (!enabled) {                                   // park it (keeps its schedule for a later move back)
    fs.mkdirSync(PARKED, { recursive: true });
    await writePlist(parked, plistFor(job, s)); try { fs.unlinkSync(live); } catch { }
    return { ok: true };
  }
  try { fs.mkdirSync(path.join(REPO, path.dirname(job.log)), { recursive: true }); } catch { }
  const w = await writePlist(live, plistFor(job, s)); if (!w.ok) return w;
  try { fs.unlinkSync(parked); } catch { }
  return bootstrap(live);
}

// ---------------- cron (server) ----------------
// managed lines end with "# mesh:<id>"; a disabled job is kept as "#off <line>" so its schedule survives
const cronLine = (job, s) => {
  const cmd = `cd $G/${job.dir} && ${job.cmd} >> $G/${job.log}.out.log 2>> $G/${job.log}.err.log  # mesh:${job.id}`;
  if (s.every) {
    const e = s.every;
    if (e < 60) return [`*/${e} * * * *  ${cmd}`];
    if (e % 60 === 0 && e < 1440) return [`0 */${e / 60} * * *  ${cmd}`];
    return [`0 0 * * *  ${cmd}`];
  }
  const dow = s.days ? s.days.map(d => d % 7).sort().join(',') : '*';
  return s.at.map(t => { const [h, m] = t.split(':').map(Number); return `${m} ${h} * * ${dow}  ${cmd}`; });
};
const expand = f => f === '*' ? null : f.split(',').flatMap(p => { const [a, b] = p.split('-').map(Number); return b != null ? Array.from({ length: b - a + 1 }, (_, i) => a + i) : [a]; });
function schedFromCron(lines) {
  const at = [], days = new Set(); let every = null, allDays = false;
  for (const l of lines) {
    const f = l.replace(/^#off\s+/, '').trim().split(/\s+/);
    if (/^\*\/\d+$/.test(f[0]) && f[1] === '*') { every = +f[0].slice(2); continue; }
    if (f[0] === '0' && /^\*\/\d+$/.test(f[1])) { every = 60 * +f[1].slice(2); continue; }
    for (const h of expand(f[1]) || [0]) for (const m of expand(f[0]) || [0]) at.push(pad(h) + ':' + pad(m));
    const d = expand(f[4]); if (!d) allDays = true; else d.forEach(x => days.add(x === 0 ? 7 : x));
  }
  if (every) return { every };
  return at.length ? sortSched({ at, days: allDays ? null : [...days] }) : null;
}
const belongs = (job, l) => l.includes('# mesh:' + job.id) ||
  (!l.includes('# mesh:') && l.includes(job.dir) && new RegExp(job.cmd.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(\\s|$|\\s*>>)').test(l) &&
    !(job.cmd.split(' ').length === 2 && /--(retry-if-held|catchup)/.test(l)));   // the bare broadcaster vs its --flag siblings

async function serverRead(jobs) {
  const logs = jobs.map(j => `${SRV_REPO}/${j.log}.out.log`).join(' ');
  const r = await ssh(`crontab -l 2>/dev/null; echo '@@MESH@@'; stat -c '%Y %n' ${logs} 2>/dev/null; echo '@@MESH@@'; tmux list-windows -t gcx -F '#{window_name}' 2>/dev/null`, null, 20000);
  if (!r.ok && !r.out) return { reachable: false, err: r.err || 'server unreachable' };
  const [cron, stats, wins] = r.out.split('@@MESH@@\n');
  const mt = {}; for (const l of (stats || '').split('\n')) { const m = /^(\d+) (.+)$/.exec(l); if (m) mt[m[2]] = +m[1] * 1000; }
  return { reachable: true, cron: cron || '', mtimes: mt, windows: (wins || '').split('\n').filter(Boolean) };
}
function serverState(job, srv) {
  const lines = srv.cron.split('\n').filter(l => l.trim() && !/^\s*#(?!off)/.test(l) && belongs(job, l));
  const on = lines.filter(l => !l.startsWith('#off')), off = lines.filter(l => l.startsWith('#off'));
  return { installed: on.length > 0, parked: !on.length && off.length > 0, schedule: schedFromCron(on.length ? on : off), last: srv.mtimes[`${SRV_REPO}/${job.log}.out.log`] || null };
}
async function serverSet(job, s, enabled) {
  const r = await ssh('crontab -l 2>/dev/null'); const cur = r.out || '';
  const keep = cur.split('\n').filter(l => !(l.trim() && !/^\s*#(?!off)/.test(l) && belongs(job, l)) && !l.includes('# mesh:' + job.id));
  while (keep.length && !keep[keep.length - 1].trim()) keep.pop();
  if (!keep.some(l => /^G=/.test(l))) keep.unshift('SHELL=/bin/bash', `PATH=${SRV_HOME}/.local/bin:/usr/local/bin:/usr/bin:/bin`, `G=${SRV_REPO}`);
  const add = cronLine(job, s).map(l => enabled ? l : '#off ' + l);
  const next = keep.concat(add).join('\n') + '\n';
  await ssh(`mkdir -p ${SRV_REPO}/${path.posix.dirname(job.log)}`);
  return ssh('crontab -', next);
}

// ---------------- moving a job between machines (state travels with it) ----------------
const statePaths = job => (job.state || []).map(p => p.startsWith('~/') ? p.slice(2) : 'Desktop/GCX/' + p);
async function copyState(job, to) {
  const rel = statePaths(job); if (!rel.length) return { ok: true };
  if (to === 'server') {
    const exist = rel.filter(r => fs.existsSync(path.join(HOME, r))); if (!exist.length) return { ok: true };
    return new Promise(res => {
      const tar = spawn('/usr/bin/tar', ['czf', '-', '-C', HOME, ...exist], { env: { ...process.env, COPYFILE_DISABLE: '1' } });
      const s = spawn('/usr/bin/ssh', ['-o', 'BatchMode=yes', SERVER, 'tar xzf - -C ~ 2>/dev/null']);
      tar.stdout.pipe(s.stdin); s.on('close', c => res({ ok: c === 0 })); s.on('error', e => res({ ok: false, err: e.message }));
    });
  }
  return new Promise(res => {
    const s = spawn('/usr/bin/ssh', ['-o', 'BatchMode=yes', SERVER, `cd ~ && tar czf - ${rel.map(r => `'${r}'`).join(' ')} 2>/dev/null`]);
    const tar = spawn('/usr/bin/tar', ['xzf', '-', '-C', HOME]);
    s.stdout.pipe(tar.stdin); tar.on('close', c => res({ ok: c === 0 })); tar.on('error', e => res({ ok: false, err: e.message }));
  });
}

// ---------------- public API ----------------
async function list() {
  const cat = catalog(), jobs = cat.jobs;
  const lc = await run('/bin/launchctl', ['list']), loaded = new Set((lc.out || '').split('\n').map(l => l.split('\t')[2]).filter(Boolean));
  const srv = await serverRead(jobs);
  const out = [];
  for (const j of jobs) {
    const mac = await macState(j, loaded), server = srv.reachable ? serverState(j, srv) : null;
    const where = mac.installed && server && server.installed ? 'both' : mac.installed ? 'mac' : server && server.installed ? 'server' : 'off';
    const home = where === 'both' || where === 'mac' ? 'mac' : where === 'server' ? 'server' : (mac.parked && !(server && server.parked)) ? 'mac' : 'server';
    const schedule = (where === 'server' ? server.schedule : where === 'mac' || where === 'both' ? mac.schedule : (home === 'mac' ? mac.schedule : server && server.schedule)) || j.default;
    out.push({ ...j, where, home, enabled: where !== 'off', schedule, last: where === 'server' ? server.last : mac.last, mac, server });
  }
  const sessions = (cat.sessions || []).map(s => ({ ...s, running: srv.reachable ? srv.windows.includes(s.tmux.split(':')[1]) : null }));
  return { jobs: out, sessions, cloud: cat.cloud || [], server: { reachable: srv.reachable, err: srv.err } };
}
const find = id => catalog().jobs.find(j => j.id === id);

async function setSchedule(id, s) {
  const job = find(id); if (!job) return { ok: false, err: 'unknown job' };
  s = sortSched(s); const st = (await list()).jobs.find(j => j.id === id);
  if (st.where === 'mac' || st.where === 'both') { const r = await macSet(job, s, true); if (!r.ok) return r; }
  if (st.where === 'server' || st.where === 'both') { const r = await serverSet(job, s, true); if (!r.ok) return r; }
  if (st.where === 'off') { const r = st.home === 'mac' ? await macSet(job, s, false) : await serverSet(job, s, false); if (!r.ok) return r; }
  return { ok: true };
}
async function setEnabled(id, on) {
  const job = find(id); if (!job) return { ok: false, err: 'unknown job' };
  const st = (await list()).jobs.find(j => j.id === id);
  return st.home === 'mac' ? macSet(job, st.schedule, on) : serverSet(job, st.schedule, on);
}
// enable on the target first (state copied before), then switch the source off — never both "off"
async function move(id, to, syncCode) {
  const job = find(id); if (!job) return { ok: false, err: 'unknown job' };
  if (job.movable === false) return { ok: false, err: job.pinnedReason || 'this job can’t move' };
  const st = (await list()).jobs.find(j => j.id === id);
  if (to === 'server' && !(await serverRead([job])).reachable) return { ok: false, err: 'server unreachable' };
  if (to === 'server' && syncCode) await syncCode();
  // switch the source off first, then copy the newest state, then enable the target → no double run
  if (to === 'server') await macSet(job, st.schedule, false); else await serverSet(job, st.schedule, false);
  const c = await copyState(job, to); if (!c.ok) return { ok: false, err: 'state copy failed — the job is OFF on both; retry the move' };
  const r = to === 'server' ? await serverSet(job, st.schedule, true) : await macSet(job, st.schedule, true);
  return r.ok ? { ok: true } : { ok: false, err: r.err || 'enable failed' };
}
async function runNow(id) {
  const job = find(id); if (!job) return { ok: false, err: 'unknown job' };
  const st = (await list()).jobs.find(j => j.id === id), where = st.where === 'off' ? st.home : st.where === 'both' ? 'mac' : st.where;
  if (where === 'mac') return run('/bin/launchctl', ['kickstart', `gui/${UID}/${label(id)}`]).then(r => r.ok ? r :
    run('/bin/bash', ['-lc', `cd '${path.join(REPO, job.dir)}' && ${job.cmd.replace(/^python3/, '/opt/homebrew/bin/python3')} >> '${path.join(REPO, job.log)}.out.log' 2>> '${path.join(REPO, job.log)}.err.log' &`]));
  return ssh(`cd ${SRV_REPO}/${job.dir} && export PATH=${SRV_HOME}/.local/bin:/usr/local/bin:/usr/bin:/bin && nohup ${job.cmd} >> ${SRV_REPO}/${job.log}.out.log 2>> ${SRV_REPO}/${job.log}.err.log & echo started`);
}
// ---------------- silent background infrastructure (runs with no Claude window open) ----------------
// Windows Task Scheduler on the server laptop (claude-server) and the WSL @reboot cron line.
const WIN = ['user@claude-server', '-i', path.join(HOME, '.ssh', 'id_ed25519_gcx_server')];
const WIN_TASKS = {
  'GCX-Live': 'At Windows logon: starts WSL Ubuntu and attaches the shared tmux session `gcx` on the laptop screen (ticket monitor, logs).',
  'WSL-KeepAlive': 'At logon: keeps WSL Ubuntu (gcx-server) running so cron jobs and Claude sessions never stop.',
  'GCX-SetupLog': 'At logon: shows the setup/admin log window (C:\\GCX-Setup\\setup.log) on the laptop screen.',
  'ClaudeServer-Elevate': 'On demand: helper used by remote admin scripts to run elevated commands.',
};
const ps = cmd => { const enc = Buffer.from(cmd, 'utf16le').toString('base64');
  return run('/usr/bin/ssh', ['-o', 'BatchMode=yes', '-o', 'ConnectTimeout=8', '-i', WIN[2], WIN[0], `powershell -NoProfile -NonInteractive -EncodedCommand ${enc}`], null, 30000); };
async function infra() {
  const names = Object.keys(WIN_TASKS);
  const r = await ps(`$ProgressPreference='SilentlyContinue'; foreach ($n in @(${names.map(n => `'${n}'`).join(',')})) { $t = Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue; if ($t) { $i = Get-ScheduledTaskInfo -TaskName $n; $tr = ($t.Triggers | % { $_.CimClass.CimClassName -replace 'MSFT_Task','' -replace 'Trigger','' }) -join ','; Write-Output ("@@" + $n + "|" + $t.State + "|" + $tr + "|" + $i.LastRunTime.ToString('o') + "|" + $i.NextRunTime) } }`);
  const items = [];
  for (const line of (r.out || '').split(/\r?\n/)) {
    const m = /^@@([^|]+)\|([^|]*)\|([^|]*)\|([^|]*)\|(.*)$/.exec(line.trim()); if (!m) continue;
    items.push({ id: 'win:' + m[1], kind: 'windows-task', where: 'server-windows', name: m[1], what: WIN_TASKS[m[1]] || '', enabled: m[2] !== 'Disabled', state: m[2],
      when: m[3] === 'Logon' ? 'At Windows logon' : m[3] === 'Boot' ? 'At boot' : m[3] || 'On demand', last: Date.parse(m[4]) || null });
  }
  const c = await ssh('crontab -l 2>/dev/null | grep -E "@reboot" ', null, 15000);
  for (const l of (c.out || '').split('\n').filter(Boolean)) items.push({ id: 'cron:reboot', kind: 'cron-reboot', where: 'server', name: 'tmux gcx at boot (WSL)', what: 'At WSL start: runs ~/tmux_start.sh, which creates the shared tmux session gcx with the ticket-monitor Claude session and log windows.', enabled: !l.startsWith('#off'), when: 'At boot', line: l.replace(/^#off\s+/, '') });
  return { reachable: r.ok || !!items.length, items, err: r.ok ? null : r.err };
}
async function setInfraEnabled(id, on) {
  if (id.startsWith('win:')) { const n = id.slice(4).replace(/'/g, ''); return ps(`${on ? 'Enable' : 'Disable'}-ScheduledTask -TaskName '${n}' | Out-Null; 'ok'`); }
  if (id === 'cron:reboot') {
    const r = await ssh('crontab -l 2>/dev/null'); const lines = (r.out || '').split('\n').map(l => /@reboot/.test(l) ? (on ? l.replace(/^#off\s+/, '') : (l.startsWith('#off') ? l : '#off ' + l)) : l);
    return ssh('crontab -', lines.join('\n').replace(/\n*$/, '\n'));
  }
  return { ok: false, err: 'unknown item' };
}
module.exports = { list, setSchedule, setEnabled, move, runNow, infra, setInfraEnabled };
