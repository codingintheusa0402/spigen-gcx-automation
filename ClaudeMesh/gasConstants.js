// Due dates / end dates that live as constants in Apps Script code: listed in the Schedules menu and
// editable there. Edit = rewrite the literal in the source file → `clasp push --force` → commit + push.
// The list of editable constants is GCX ServerBootstrap/gas_due_dates.json (filled from the trigger scan).
const fs = require('fs'); const os = require('os'); const path = require('path'); const { execFile } = require('child_process');
const REPO = path.join(os.homedir(), 'Desktop', 'GCX'), CAT = path.join(REPO, 'ServerBootstrap', 'gas_due_dates.json');
const sh = (cmd, args, cwd) => new Promise(res => execFile(cmd, args, { cwd, timeout: 120000, env: { ...process.env, PATH: '/opt/homebrew/bin:/usr/local/bin:' + process.env.PATH } },
  (e, so, se) => res({ ok: !e, out: String(so || ''), err: String(se || (e && e.message) || '') })));
function load() { try { return JSON.parse(fs.readFileSync(CAT, 'utf8')); } catch { return { projects: [] }; } }
// fill each due date's current value from the source (catalog holds the regex to find it)
function list() {
  const c = load();
  for (const p of c.projects || []) for (const d of p.dueDates || []) {
    try { const src = fs.readFileSync(path.join(REPO, d.file.split(':')[0]), 'utf8'), m = new RegExp(d.match).exec(src); d.value = m ? toISO(m[1]) : d.value || ''; d.raw = m ? m[1] : null; } catch { }
  }
  return c;
}
function toISO(v) {                                  // '2026-11-18', '2026/11/18', 'new Date(2026, 10, 18)' args, '26/11/18'
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/.exec(v); if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = /^(\d{4}),\s*(\d{1,2}),\s*(\d{1,2})/.exec(v); if (m) return `${m[1]}-${String(+m[2] + 1).padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = /^(\d{2})[/.](\d{1,2})[/.](\d{1,2})$/.exec(v); if (m) return `20${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return v;
}
function fromISO(iso, raw) {                         // write back in the same style as the original literal
  const [y, mo, d] = iso.split('-');
  if (/^\d{4},/.test(raw)) return `${y}, ${+mo - 1}, ${+d}`;
  if (/^\d{4}\//.test(raw)) return `${y}/${mo}/${d}`;
  if (/^\d{4}\./.test(raw)) return `${y}.${mo}.${d}`;
  if (/^\d{2}[/.]/.test(raw)) return `${y.slice(2)}${raw[2]}${mo}${raw[2]}${d}`;
  return `${y}-${mo}-${d}` + raw.slice(10);
}
async function set(scriptId, i, iso) {
  const c = list(), p = (c.projects || []).find(x => x.scriptId === scriptId), d = p && p.dueDates[i];
  if (!d || d.raw == null) return { ok: false, err: 'constant not found in source' };
  const file = path.join(REPO, d.file.split(':')[0]), src = fs.readFileSync(file, 'utf8'), re = new RegExp(d.match);
  const m = re.exec(src); const next = src.slice(0, m.index) + m[0].replace(m[1], fromISO(iso, m[1])) + src.slice(m.index + m[0].length);
  fs.writeFileSync(file, next);
  const dir = path.join(REPO, p.dir);
  const push = await sh('clasp', ['push', '--force'], dir);
  if (!push.ok) { fs.writeFileSync(file, src); return { ok: false, err: 'clasp push failed (file restored): ' + push.err.slice(0, 200) }; }
  await sh('git', ['add', path.relative(REPO, file)], REPO);
  await sh('git', ['-c', 'user.name=Spigen1', '-c', 'user.email=kjw@spigen.com', 'commit', '-m', `chore(${path.basename(p.dir)}): ${d.label} → ${iso} (via GCX Mesh Schedules)`], REPO);
  await sh('/bin/bash', [path.join(REPO, 'gcx-sync.sh'), 'push'], REPO);
  return { ok: true };
}
module.exports = { list, set };
