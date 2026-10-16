// GCX Mesh — the other machine. The Mac app and the server program each run a phone server; with a
// peer configured, each one also shows the OTHER machine's live sessions (fetched over Tailscale) and
// forwards send / interrupt / rename / details / terminal for them, so the iPhone sees one mesh whichever
// machine it is connected to. When the peer goes offline its last-seen sessions stay listed as offline
// (with "Resume here" once their transcript has been synced to this machine).
//
// Config: <userData>/peer.json  { "url": "http://100.112.43.88:47320", "token": "<the peer's phone token>",
//                                 "kind": "mac" | "server", "label": "mini" (optional display name) }
// Nothing here touches Electron.
const fs = require('fs');
const path = require('path');
const http = require('http');

const POLL_MS = 4000, OFFLINE_KEEP_MS = 7 * 24 * 3600 * 1000;
const PREFIX = 'peer:';                     // terminal ids of the peer's panes, as seen by the phone

function createPeer({ userData, selfKind }) {
  const file = path.join(userData, 'peer.json');
  let cfg = null, last = null, lastOk = 0, online = false, timer = null;

  function loadCfg() {
    try { const c = JSON.parse(fs.readFileSync(file, 'utf8')); cfg = c && c.url && c.token ? c : null; } catch { cfg = null; }
    return cfg;
  }
  const kind = () => (cfg && cfg.kind) || (selfKind === 'server' ? 'mac' : 'server');
  const label = () => (cfg && cfg.label) || (kind() === 'mac' ? 'Mac' : 'server');

  function request(method, p, { body, query, timeout = 6000, stream } = {}) {
    return new Promise((resolve, reject) => {
      if (!cfg) return reject(new Error('no peer configured'));
      const u = new URL(p, cfg.url);
      u.searchParams.set('t', cfg.token);
      for (const [k, v] of Object.entries(query || {})) u.searchParams.set(k, v);
      const data = body === undefined ? null : Buffer.from(JSON.stringify(body));
      const req = http.request(u, { method, timeout, headers: data ? { 'Content-Type': 'application/json', 'Content-Length': data.length } : {} }, res => {
        if (stream) return resolve(res);
        let s = ''; res.setEncoding('utf8'); res.on('data', d => s += d);
        res.on('end', () => { try { resolve({ status: res.statusCode, json: JSON.parse(s || '{}') }); } catch { resolve({ status: res.statusCode, json: {} }); } });
      });
      req.on('timeout', () => req.destroy(new Error('timeout')));
      req.on('error', reject);
      if (data) req.write(data);
      req.end();
    });
  }

  async function poll() {
    if (!loadCfg()) { online = false; return; }
    try {
      const r = await request('GET', '/api/local-snapshot', { timeout: 3500 });
      if (r.status === 200 && r.json && Array.isArray(r.json.sessions)) { last = r.json; lastOk = Date.now(); online = true; }
      else online = false;
    } catch { online = false; }
  }
  function start() { if (!timer) { poll(); timer = setInterval(poll, POLL_MS); } }

  const remoteSids = () => new Set(((last && last.sessions) || []).map(s => s.sid));
  const owns = sid => online && remoteSids().has(sid);
  const ownsTerm = id => typeof id === 'string' && id.startsWith(PREFIX);

  // local snapshot + the peer's sessions (marked remote; offline ones greyed) — never mutates `local`
  function merge(local) {
    if (!local) return local;
    const out = { ...local, peer: cfg ? { kind: kind(), label: label(), online, lastSeen: lastOk || null } : null };
    out.sessions = local.sessions.map(s => ({ ...s, hostKind: selfKind }));
    if (!last || !cfg || (!online && Date.now() - lastOk > OFFLINE_KEEP_MS)) return out;
    const mine = new Set(out.sessions.map(s => s.sid));
    for (const s of last.sessions || []) {
      if (mine.has(s.sid)) continue;                 // same conversation resumed here — the local copy wins
      out.sessions.push({ ...s, hostKind: kind(), hostLabel: label(), onPeer: true, offline: !online,
        owner: online && s.owner ? PREFIX + s.owner : null,
        health: online ? s.health : 'offline', tps: online ? s.tps : 0 });
    }
    if (online) out.ptys = [...(local.ptys || []), ...((last.ptys || []).map(p => ({ ...p, id: PREFIX + p.id, onPeer: true })))];
    return out;
  }

  return {
    start, merge, owns, ownsTerm, online: () => online, label,
    configured: () => !!loadCfg(),
    // forward a JSON API call for a peer session
    forward: (method, p, body, query) => request(method, p, { body, query }).then(r => r.json),
    // pipe the peer's terminal SSE stream for terminal `id` (with the peer: prefix) into `res`
    async pipeTerm(id, res, onClose) {
      const up = await request('GET', '/api/term/' + encodeURIComponent(id.slice(PREFIX.length)), { stream: true, timeout: 0 });
      res.writeHead(up.statusCode, { 'Content-Type': up.headers['content-type'] || 'text/event-stream', 'Cache-Control': 'no-store', 'X-Accel-Buffering': 'no' });
      up.pipe(res);
      onClose(() => up.destroy());
    },
    // raw SSE response of the peer's terminal `id` (with the peer: prefix), for the desktop window's mirror
    openTerm: id => request('GET', '/api/term/' + encodeURIComponent(id.slice(PREFIX.length)), { stream: true, timeout: 0 }),
    upload: (name, b64) => request('POST', '/api/upload', { body: { name, b64 }, timeout: 30000 }).then(r => r.json),
    termResize: (id, cols, rows) => request('POST', '/api/term/' + encodeURIComponent(id.slice(PREFIX.length)), { body: { resize: { cols, rows } } }).then(r => r.json),
    termWrite: (id, data) => request('POST', '/api/term/' + encodeURIComponent(id.slice(PREFIX.length)), { body: { data } }).then(r => r.json),
  };
}

module.exports = { createPeer, PEER_PREFIX: PREFIX };
