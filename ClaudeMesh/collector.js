// Reads live Claude Code state from disk:
//   ~/.claude/sessions/<pid>.json        → which sessions are alive, name, busy/idle
//   ~/.claude/projects/*/<sid>.jsonl     → per-message model + token usage (spend, speed, context)
//   ~/.claude/projects/*/<sid>/subagents → sub-agent transcripts (counted toward parent)
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');

const HOME = os.homedir();
const SESS_DIR = path.join(HOME, '.claude', 'sessions');
const PROJ_DIR = path.join(HOME, '.claude', 'projects');

// $ per million tokens. cacheWrite = 1.25x input (5m TTL) / 2x input (1h TTL).
const PRICES = [
  ['claude-fable',     { in: 10, out: 50, cr: 0.25 }],
  ['claude-mythos',    { in: 10, out: 50, cr: 0.25 }],
  ['claude-opus-5-5',  { in: 4,  out: 20, cr: 0.20 }],
  ['claude-opus',      { in: 5,  out: 25, cr: 0.50 }],
  ['claude-sonnet-5',  { in: 2,  out: 10, cr: 0.20 }],   // sonnet-5 and sonnet-5-5
  ['claude-sonnet',    { in: 3,  out: 15, cr: 0.30 }],
  ['claude-haiku',     { in: 1,  out: 5,  cr: 0.10 }],
];
function priceFor(model) {
  for (const [p, v] of PRICES) if (model.startsWith(p)) return v;
  return { in: 3, out: 15, cr: 0.3 };
}
function family(model) {
  const m = /claude-(fable|mythos|opus|sonnet|haiku)/.exec(model || '');
  return m ? m[1] : 'other';
}
function ctxWindow(model) { return /haiku/.test(model) ? 200_000 : 1_000_000; }

function costOf(model, u) {
  const p = priceFor(model);
  const cc = u.cache_creation || {};
  const c1h = cc.ephemeral_1h_input_tokens || 0;
  const c5m = cc.ephemeral_5m_input_tokens != null ? cc.ephemeral_5m_input_tokens
    : Math.max(0, (u.cache_creation_input_tokens || 0) - c1h);
  let usd = ((u.input_tokens || 0) * p.in + (u.output_tokens || 0) * p.out +
    (u.cache_read_input_tokens || 0) * p.cr + c5m * p.in * 1.25 + c1h * p.in * 2) / 1e6;
  if (u.speed === 'fast') usd *= 2;
  return usd;
}

const dayKey = (t) => { const d = new Date(t); return `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`; };

const SERVER_RE = /(\b(ssh|scp|rsync|sftp|mosh)\b[^\n]*\b(gcx-server|claude-server|100\.115\.156\.121|laptop-n9abvord)\b)|tailscale ssh|run-on-server|run_on_server\.sh|sync_to_server\.sh|skill:run-on-server|\bpeek\.sh\b.*gcx|gcx-server:/i;

class SessionStats {
  constructor(sid) {
    this.sid = sid;
    this.total = { usd: 0, in: 0, out: 0, cr: 0, cw: 0 };
    this.byModel = {};          // model -> {usd,out,in,cr,cw}
    this.today = { usd: 0, out: 0 };
    this.todayByModel = {};
    this.events = [];           // recent {t, out, usd, model, dur} for rate windows (last 15 min)
    this.lastModel = null;
    this.lastCtx = 0;
    this.lastSpeed = 0;         // tok/s of the most recent completed response
    this.lastAssistantAt = 0;
    this.lastWriteAt = 0;
    this.errors = [];           // timestamps of API error messages
    this.lastError = '';
    this.log = [];              // [{t, kind, text}]
    this.turns = 0;
    this.turnTimes = [];        // timestamps of replies in the last 24 h (activity frequency)
  }
  pushLog(t, kind, text) {
    this.log.push({ t, kind, text: String(text).replace(/\s+/g, ' ').slice(0, 220) });
    if (this.log.length > 60) this.log.splice(0, this.log.length - 60);
  }
}

class FileTail {
  constructor(file, sid, isSub) {
    this.file = file; this.sid = sid; this.isSub = isSub;
    this.offset = 0; this.rest = '';
    this.prevTs = 0;            // timestamp of last non-assistant line (response start proxy)
    this.msgs = new Map();      // message.id -> {model, usage, start, end}
    this.size = 0;
  }
}

class Collector {
  constructor() {
    this.tails = new Map();     // file -> FileTail
    this.stats = new Map();     // sid -> SessionStats
    this.sessions = new Map();  // pid -> session json (+ proc info)
    this.sidToFile = new Map();
    this.ps = new Map();        // pid -> {ppid, tty, cpu, rss, etime}
    this.dayStart = new Date().setHours(0, 0, 0, 0);
    this.scanning = false;
    this.ready = false;
  }

  stat(sid) {
    let s = this.stats.get(sid);
    if (!s) { s = new SessionStats(sid); this.stats.set(sid, s); }
    return s;
  }

  // ---- session registry -------------------------------------------------
  readSessions() {
    const out = new Map();
    let files = [];
    try { files = fs.readdirSync(SESS_DIR).filter(f => /^\d+\.json$/.test(f)); } catch { }
    for (const f of files) {
      try {
        const j = JSON.parse(fs.readFileSync(path.join(SESS_DIR, f), 'utf8'));
        if (!j.pid) continue;
        try { process.kill(j.pid, 0); } catch { continue; }   // dead pid → skip
        out.set(j.pid, j);
      } catch { }
    }
    this.sessions = out;
  }

  refreshPs() {
    return new Promise(res => {
      execFile('/bin/ps', ['-axo', 'pid=,ppid=,tty=,%cpu=,rss=,etime='], { maxBuffer: 8e6 }, (err, so) => {
        if (!err) {
          const m = new Map();
          for (const line of so.split('\n')) {
            const p = line.trim().split(/\s+/);
            if (p.length < 6) continue;
            m.set(+p[0], { ppid: +p[1], tty: p[2], cpu: +p[3], rss: +p[4], etime: p[5] });
          }
          this.ps = m;
        }
        res();
      });
    });
  }

  // Is `pid` a descendant of `ancestor`?
  descends(pid, ancestor) {
    for (let i = 0; i < 12 && pid > 1; i++) {
      if (pid === ancestor) return true;
      const p = this.ps.get(pid); if (!p) return false; pid = p.ppid;
    }
    return false;
  }

  // ---- transcript discovery ------------------------------------------------
  discoverFiles() {
    const cutoff = Date.now() - 36 * 3600e3;
    let dirs = [];
    try { dirs = fs.readdirSync(PROJ_DIR); } catch { }
    const liveSids = new Set([...this.sessions.values()].map(s => s.sessionId));
    for (const d of dirs) {
      const pd = path.join(PROJ_DIR, d);
      let ents; try { ents = fs.readdirSync(pd, { withFileTypes: true }); } catch { continue; }
      for (const e of ents) {
        if (e.isFile() && e.name.endsWith('.jsonl')) {
          const sid = e.name.slice(0, -6);
          const f = path.join(pd, e.name);
          let st; try { st = fs.statSync(f); } catch { continue; }
          if (liveSids.has(sid) || st.mtimeMs > cutoff) {
            if (!this.tails.has(f)) this.tails.set(f, new FileTail(f, sid, false));
            this.sidToFile.set(sid, f);
          }
        } else if (e.isDirectory()) {
          const sub = path.join(pd, e.name, 'subagents');
          let subs; try { subs = fs.readdirSync(sub); } catch { continue; }
          for (const s of subs) {
            if (!s.endsWith('.jsonl')) continue;
            const f = path.join(sub, s);
            if (this.tails.has(f)) continue;
            let st; try { st = fs.statSync(f); } catch { continue; }
            if (liveSids.has(e.name) || st.mtimeMs > cutoff) this.tails.set(f, new FileTail(f, e.name, true));
          }
        }
      }
    }
  }

  // ---- incremental read ----------------------------------------------------
  async readTail(t) {
    let st; try { st = fs.statSync(t.file); } catch { return; }
    if (st.size < t.offset) { t.offset = 0; t.rest = ''; }   // truncated/rewritten
    if (st.size === t.offset) return;
    t.size = st.size;
    const initial = t.offset === 0;
    const fd = fs.openSync(t.file, 'r');
    const CH = 4 << 20;
    const buf = Buffer.alloc(CH);
    try {
      while (t.offset < st.size) {
        const n = fs.readSync(fd, buf, 0, Math.min(CH, st.size - t.offset), t.offset);
        if (n <= 0) break;
        t.offset += n;
        const text = t.rest + buf.toString('utf8', 0, n);
        const lines = text.split('\n');
        t.rest = lines.pop();
        const nearEnd = st.size - t.offset < (2 << 20);
        for (const l of lines) this.processLine(t, l, initial && !nearEnd);
        if (initial) await new Promise(r => setImmediate(r));   // stay responsive on big files
      }
    } finally { fs.closeSync(fd); }
    if (!t.isSub) this.stat(t.sid).lastWriteAt = st.mtimeMs;
  }

  processLine(t, line, quiet) {
    if (!line) return;
    const tsm = /"timestamp":"([^"]+)"/.exec(line);
    const ts = tsm ? Date.parse(tsm[1]) : 0;
    const isAsst = line.includes('"type":"assistant"');
    if (!isAsst) {
      if (ts) t.prevTs = ts;
      if (!t.isSub && line.startsWith('{"type":"custom-title"')) { try { this.stat(t.sid).customTitle = JSON.parse(line).customTitle || ''; } catch { } return; }
      if (!t.isSub && line.startsWith('{"type":"ai-title"')) { try { this.stat(t.sid).aiTitle = JSON.parse(line).aiTitle || ''; } catch { } return; }
      if (quiet) return;
      if (line.includes('"type":"user"') && !t.isSub) {
        let j; try { j = JSON.parse(line); } catch { return; }
        const c = j.message && j.message.content;
        const s = this.stat(t.sid);
        if (typeof c === 'string') { if (!j.isMeta && !c.startsWith('<')) s.pushLog(ts, 'you', c); }
        else if (Array.isArray(c)) {
          for (const b of c) {
            if (b.type === 'text' && !j.isMeta && !b.text.startsWith('<')) s.pushLog(ts, 'you', b.text);
            if (b.type === 'tool_result' && b.is_error) s.pushLog(ts, 'err', 'tool error: ' + (typeof b.content === 'string' ? b.content : '').slice(0, 120));
          }
        }
      }
      return;
    }
    let j; try { j = JSON.parse(line); } catch { return; }
    const m = j.message || {};
    const s = this.stat(t.sid);
    if (j.isApiErrorMessage) {
      s.errors.push(ts);
      const txt = (m.content || []).map(b => b.text || '').join(' ');
      s.lastError = txt.slice(0, 200);
      if (!quiet) s.pushLog(ts, 'err', txt);
      return;
    }
    const u = m.usage; if (!u || !m.model || m.model === '<synthetic>') return;
    const id = m.id || j.requestId || j.uuid;
    let rec = t.msgs.get(id);
    if (!rec) {
      rec = { model: m.model, start: t.prevTs || ts, end: ts, cost: 0, out: 0, inp: 0, cr: 0, cw: 0, ev: null };
      t.msgs.set(id, rec);
      if (t.msgs.size > 400) t.msgs.delete(t.msgs.keys().next().value);
      if (!t.isSub) { s.turns++; if (ts > Date.now() - 864e5) s.turnTimes.push(ts); }   // for 24 h frequency
    }
    // Same message is re-logged once per content block with cumulative usage → apply deltas.
    const cost = costOf(m.model, u);
    const out = u.output_tokens || 0, inp = u.input_tokens || 0,
      cr = u.cache_read_input_tokens || 0, cw = u.cache_creation_input_tokens || 0;
    const d = { usd: cost - rec.cost, out: out - rec.out, in: inp - rec.inp, cr: cr - rec.cr, cw: cw - rec.cw };
    Object.assign(rec, { cost, out, inp, cr, cw, end: ts });
    const bm = s.byModel[m.model] || (s.byModel[m.model] = { usd: 0, out: 0, in: 0, cr: 0, cw: 0 });
    for (const k of ['usd', 'out', 'in', 'cr', 'cw']) { s.total[k] += d[k]; bm[k] += d[k]; }
    if (ts >= this.dayStart) {
      s.today.usd += d.usd; s.today.out += d.out;
      const tm = s.todayByModel[m.model] || (s.todayByModel[m.model] = { usd: 0, out: 0 });
      tm.usd += d.usd; tm.out += d.out;
    }
    const dur = Math.max(0.3, (rec.end - rec.start) / 1000);
    if (ts > Date.now() - 15 * 60e3) {
      if (!rec.ev) { rec.ev = { t: ts, out: 0, usd: 0, model: m.model, sub: t.isSub }; s.events.push(rec.ev); }
      rec.ev.t = ts; rec.ev.out = out; rec.ev.usd = cost; rec.ev.dur = dur;
    }
    // does this session reach the 24/7 server? (ssh/scp/rsync to gcx-server / claude-server, Tailscale SSH,
    // the run-on-server skill or its scripts) — sticky per session, with the time of the latest use
    for (const b of m.content || []) {
      if (b.type !== 'tool_use') continue;
      const inp = b.input || {}, txt = b.name === 'Skill' ? 'skill:' + (inp.skill || '') : String(inp.command || inp.cmd || '');
      if (SERVER_RE.test(txt)) { s.serverUses = (s.serverUses || 0) + 1; s.serverAt = Math.max(s.serverAt || 0, ts || Date.now()); }
    }
    if (!t.isSub) {
      s.lastModel = m.model;
      s.lastCtx = inp + cr + cw;
      s.lastAssistantAt = ts;
      if (out > 20) s.lastSpeed = out / dur;
      if (!quiet) for (const b of m.content || []) {
        if (b.type === 'text' && b.text.trim()) s.pushLog(ts, 'claude', b.text);
        else if (b.type === 'tool_use') s.pushLog(ts, 'tool', b.name + ' ' + summarizeInput(b.input));
      }
    }
  }

  // ---- main loop ------------------------------------------------------------
  async tick() {
    if (this.scanning) return;
    this.scanning = true;
    try {
      const now = Date.now();
      const ds = new Date().setHours(0, 0, 0, 0);
      if (ds !== this.dayStart) {           // midnight rollover
        this.dayStart = ds;
        for (const s of this.stats.values()) { s.today = { usd: 0, out: 0 }; s.todayByModel = {}; }
      }
      this.readSessions();
      await this.refreshPs();
      if (!this._lastDisc || now - this._lastDisc > 5000) { this.discoverFiles(); this._lastDisc = now; }
      for (const t of this.tails.values()) await this.readTail(t);
      for (const s of this.stats.values()) {
        s.events = s.events.filter(e => e.t > now - 15 * 60e3);
        s.errors = s.errors.filter(e => e > now - 3600e3);
        if (s.turnTimes.length && s.turnTimes[0] < now - 864e5) s.turnTimes = s.turnTimes.filter(x => x > now - 864e5);
      }
      this.ready = true;
    } finally { this.scanning = false; }
  }

  snapshot(ptyMap) {
    const now = Date.now();
    const sessions = [];
    const models = {};
    const addModel = (model, usd, out) => {
      const f = family(model);
      const m = models[f] || (models[f] = { family: f, todayUsd: 0, todayOut: 0, tps: 0, liveUsdHr: 0, sessions: 0 });
      m.todayUsd += usd; m.todayOut += out;
    };
    for (const s of this.stats.values()) for (const [mod, v] of Object.entries(s.todayByModel)) addModel(mod, v.usd, v.out);

    for (const [pid, j] of this.sessions) {
      const s = this.stat(j.sessionId);
      const p = this.ps.get(pid) || {};
      const win = s.events.filter(e => e.t > now - 60e3);
      const tps = win.reduce((a, e) => a + e.out, 0) / 60;
      const usdHr = s.events.filter(e => e.t > now - 300e3).reduce((a, e) => a + e.usd, 0) * 12;
      const subActive = new Set(s.events.filter(e => e.sub && e.t > now - 90e3).map(e => e.t)).size;
      const spark = [];
      for (let i = 29; i >= 0; i--) {    // 30 bins × 10 s = last 5 minutes, tok/s per bin
        const a = now - (i + 1) * 10e3, b = now - i * 10e3;
        spark.push(s.events.filter(e => e.t > a && e.t <= b).reduce((x, e) => x + e.out, 0) / 10);
      }
      const err5 = s.errors.filter(t => t > now - 5 * 60e3).length;
      const sinceWrite = (now - (s.lastWriteAt || j.updatedAt || now)) / 1000;
      let health = j.status === 'busy' ? 'working' : 'idle';
      if (j.status === 'busy' && sinceWrite > 300) health = 'quiet';
      if (err5) health = 'error';
      if (j.status && !['busy', 'idle'].includes(j.status)) health = 'waiting';
      const model = s.lastModel || '';
      if (model && !models[family(model)]) addModel(model, 0, 0);
      const fam = family(model);
      if (models[fam]) { models[fam].tps += tps; models[fam].liveUsdHr += usdHr; models[fam].sessions++; }
      let owner = null;
      for (const [ptyId, ptyPid] of ptyMap) if (this.descends(pid, ptyPid)) owner = ptyId;
      sessions.push({
        pid, sid: j.sessionId, name: sessionName(j, s), cwd: j.cwd, kind: j.kind,
        status: j.status, health, tty: p.tty && p.tty !== '??' ? '/dev/' + p.tty : null,
        cpu: p.cpu || 0, rssMb: Math.round((p.rss || 0) / 1024), etime: p.etime || '',
        startedAt: j.startedAt, version: j.version, remote: !!j.bridgeSessionId,
        model, family: fam, ctx: s.lastCtx, ctxMax: ctxWindow(model),
        totalUsd: s.total.usd, todayUsd: s.today.usd, tokIn: s.total.in, tokOut: s.total.out,
        tokCacheRead: s.total.cr, tokCacheWrite: s.total.cw,
        byModel: s.byModel, tps, lastSpeed: s.lastSpeed, usdHr, turns: s.turns,
        // activity frequency: replies over 24 h, recent ones weigh more (half-life 3 h)
        freq: s.turnTimes.reduce((a, x) => a + Math.pow(.5, (now - x) / 108e5), 0),
        subagents: subActive, errors5m: err5, errors1h: s.errors.length, lastError: s.lastError, serverUses: s.serverUses || 0, serverAt: s.serverAt || 0,
        sinceWrite, spark, log: s.log.slice(-40), owner,
      });
    }
    sessions.sort((a, b) => a.startedAt - b.startedAt);
    let todayAll = 0; for (const m of Object.values(models)) todayAll += m.todayUsd;
    return { t: now, ready: this.ready, sessions, models: Object.values(models), todayUsd: todayAll };
  }
}

// Display name: an explicit /rename wins, then a name Claude set itself, then the conversation's
// title (custom/AI) — the auto "user-xx" derived name only as a last resort.
function sessionName(j, s) {
  if (j.name && j.nameSource && !['derived', 'auto'].includes(j.nameSource)) return j.name;
  return s.customTitle || (j.nameSource === 'auto' && j.name) || s.aiTitle || j.name || j.sessionId.slice(0, 8);
}

function summarizeInput(inp) {
  if (!inp) return '';
  const v = inp.command || inp.file_path || inp.pattern || inp.description || inp.url || inp.query || inp.skill || '';
  return String(v).replace(HOME, '~').slice(0, 140);
}

module.exports = { Collector, costOf, family };
