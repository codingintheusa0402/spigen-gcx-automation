// Phone access: a tiny HTTP server that serves the iPhone web app (phone/) and a JSON + SSE API
// backed by the same functions the desktop window uses. Off by default; enabled from the 📱 button.
// Every data/control endpoint requires the pairing token (?t=…); static code files don't.
// Reach it on the Wi-Fi IP or the Tailscale IP — it is never exposed through Tailscale Funnel.
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

const PORT = +process.env.MESH_PHONE_PORT || 47320;   // env override: isolated test runs
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.json': 'application/json', '.webmanifest': 'application/manifest+json' };

function createPhoneServer(ctx) {
  // ctx: { root, userData, getSnap, getHistory, sendToSession, interruptSession, resumeOnMac, rename, convert,
  //        commandsFor, ptys, ptyBus, termOpen?, termClose?, bind?, peer? }
  // peer (peer.js): the other machine — its sessions are merged into what the phone sees, and calls for them
  // are forwarded there, so the Mac and the server show one mesh.
  // Mac: ptys are node-pty terminals. Server (server/mesh-server.js): ptys are tmux panes, termOpen/termClose
  // start and stop streaming a pane while a phone is watching it; bind limits the listening address.
  const cfgFile = path.join(ctx.userData, 'phone.json');
  let cfg = { enabled: false, token: '' };
  try { cfg = { ...cfg, ...JSON.parse(fs.readFileSync(cfgFile, 'utf8')) }; } catch { }
  if (!cfg.token) cfg.token = crypto.randomBytes(18).toString('base64url');
  const save = () => { try { fs.writeFileSync(cfgFile, JSON.stringify(cfg)); } catch { } };
  save();

  let server = null;
  const clients = new Set();                // SSE telemetry streams
  let peerPolledAt = 0;                     // the other machine is mirroring us → keep the snapshot fresh
  const okToken = t => typeof t === 'string' && t.length === cfg.token.length && crypto.timingSafeEqual(Buffer.from(t), Buffer.from(cfg.token));

  // static files: phone/ app + the shared renderer engine + xterm
  const STATIC = {
    '/': 'phone/index.html', '/phone.js': 'phone/phone.js', '/phone.css': 'phone/phone.css', '/icon.png': 'renderer/icon.png',
    '/apple-touch-icon.png': 'build/icon-ios.png',
    '/mesh.js': 'renderer/mesh.js', '/mother.js': 'renderer/mother.js', '/bubbles.js': 'renderer/bubbles.js', '/organism.js': 'renderer/organism.js', '/pixel.js': 'renderer/pixel.js', '/md.js': 'renderer/md.js',
    '/xterm.js': 'node_modules/@xterm/xterm/lib/xterm.js', '/xterm.css': 'node_modules/@xterm/xterm/css/xterm.css', '/addon-fit.js': 'node_modules/@xterm/addon-fit/lib/addon-fit.js',
  };

  const peer = ctx.peer || null;
  const view = s => peer && s ? peer.merge(s) : s;              // what the phone sees: local + the other machine
  const remote = sid => !!(peer && sid && peer.owns(sid));
  const termId = p => { try { return decodeURIComponent(p.slice(10)); } catch { return p.slice(10); } };   // tmux pane ids (%0) arrive encoded
  const body = (req, max = 2e6) => new Promise(res => { let b = ''; req.on('data', c => { b += c; if (b.length > max) req.destroy(); }); req.on('end', () => { try { res(JSON.parse(b || '{}')); } catch { res({}); } }); });
  const json = (res, code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(obj)); };
  const sse = res => { res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-store', Connection: 'keep-alive', 'X-Accel-Buffering': 'no' }); res.write('retry: 2000\n\n'); };
  const send = (res, ev, data) => res.write(`event: ${ev}\ndata: ${JSON.stringify(data)}\n\n`);

  async function handle(req, res) {
    const u = new URL(req.url, 'http://x'), p = u.pathname;
    res.setHeader('X-Content-Type-Options', 'nosniff');
    if (req.method === 'GET' && STATIC[p]) {
      const f = path.join(ctx.root, STATIC[p]);
      fs.readFile(f, (err, data) => {
        if (err) { res.writeHead(404); return res.end(); }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream', 'Cache-Control': 'no-cache' }); res.end(data);
      });
      return;
    }
    const tok = u.searchParams.get('t') || req.headers['x-mesh-token'];
    if (p === '/manifest.webmanifest') {      // start_url carries the token so the home-screen app is paired
      if (!okToken(tok)) { res.writeHead(401); return res.end(); }
      res.writeHead(200, { 'Content-Type': MIME['.webmanifest'] });
      return res.end(JSON.stringify({ name: 'GCX Mesh', short_name: 'Mesh', start_url: '/?t=' + cfg.token, display: 'standalone', background_color: '#000000', theme_color: '#000000',
        icons: [{ src: '/apple-touch-icon.png', sizes: '1024x1024', type: 'image/png' }] }));
    }
    if (!p.startsWith('/api/')) { res.writeHead(404); return res.end(); }
    if (!okToken(tok)) return json(res, 401, { error: 'not paired — scan the QR code in GCX Mesh on your Mac (or pair with the server link)' });

    if (p === '/api/events' && req.method === 'GET') {
      sse(res); const c = { res }; clients.add(c);
      const s = view(ctx.getSnap()); if (s) send(res, 'telemetry', s);
      req.on('close', () => clients.delete(c)); return;
    }
    if (p === '/api/local-snapshot') { peerPolledAt = Date.now(); return json(res, 200, ctx.getSnap() || { sessions: [] }); }   // for the peer (never merged)
    if (p === '/api/history') return json(res, 200, await ctx.getHistory());
    if (p === '/api/details') {
      const sid = u.searchParams.get('sid') || '';
      if (remote(sid)) return json(res, 200, await peer.forward('GET', '/api/details', undefined, { sid }).catch(() => ({})));
      return json(res, 200, await ctx.details(sid));
    }
    if (p === '/api/commands') return json(res, 200, await ctx.commandsFor(u.searchParams.get('cwd') || ''));
    if (p.startsWith('/api/term/') && req.method === 'GET') {           // live terminal of an in-app session
      const id = termId(p);
      if (peer && peer.ownsTerm(id)) return peer.pipeTerm(id, res, fn => req.on('close', fn)).catch(e => json(res, 502, { error: e.message }));
      const pt = ctx.ptys.get(id); if (!pt) return json(res, 404, { error: 'no such terminal' });
      sse(res);
      if (ctx.termOpen) await ctx.termOpen(id);                           // server: start streaming the tmux pane
      // listen first and hold output back until init went out, so nothing falls between the snapshot and the stream
      let held = [];
      const on = (pid, d) => { if (pid === id) held ? held.push(d) : send(res, 'data', d); }, off = pid => { if (pid === id) send(res, 'exit', {}); };
      const sz = (pid, cols, rows) => { if (pid === id) send(res, 'size', { cols, rows }); };   // the terminal was resized on the host
      ctx.ptyBus.on('data', on); ctx.ptyBus.on('exit', off); ctx.ptyBus.on('resize', sz);
      const screen = ctx.termSnapshot ? await ctx.termSnapshot(id).catch(() => null) : null;   // Mac app: the window's exact screen
      send(res, 'init', { cols: pt.proc.cols, rows: pt.proc.rows, backlog: screen != null ? screen : pt.buf || '' });
      for (const d of held) send(res, 'data', d); held = null;
      req.on('close', () => { ctx.ptyBus.off('data', on); ctx.ptyBus.off('exit', off); ctx.ptyBus.off('resize', sz); if (ctx.termClose) ctx.termClose(id); }); return;
    }
    if (req.method !== 'POST') return json(res, 405, {});
    const b = await body(req, p === '/api/upload' ? 40e6 : 2e6);
    if (p === '/api/upload') {                                           // a file pasted/dropped into the peer's mirror of our terminal
      const name = path.basename(String(b.name || 'file')).replace(/[^\w.\- ]+/g, '_').slice(-80) || 'file';
      const dir = path.join(os.tmpdir(), 'claude-mesh-paste'); fs.mkdirSync(dir, { recursive: true });
      const f = path.join(dir, `${Date.now().toString(36)}-${name}`);
      fs.writeFileSync(f, Buffer.from(String(b.b64 || ''), 'base64'));
      return json(res, 200, { ok: true, path: f });
    }
    if (['/api/send', '/api/interrupt', '/api/rename'].includes(p) && remote(b.sid))
      return json(res, 200, await peer.forward('POST', p, b).catch(e => ({ ok: false, err: `${peer.label()}: ${e.message}` })));
    if (p.startsWith('/api/term/') && peer && peer.ownsTerm(termId(p)))
      return json(res, 200, await peer.termWrite(termId(p), String(b.data || '')).catch(() => ({ ok: false })));
    if (p === '/api/send') return json(res, 200, await ctx.sendToSession(b.sid, String(b.text || '')));
    if (p === '/api/interrupt') return json(res, 200, await ctx.interruptSession(b.sid));
    if (p === '/api/resume') return json(res, 200, ctx.resumeOnMac(b.sid));
    if (p === '/api/rename') return json(res, 200, await ctx.rename(b.sid, String(b.name || '').slice(0, 120)));
    if (p === '/api/convert') return json(res, 200, await ctx.convert(b.sid, b.task || ''));
    if (p.startsWith('/api/term/')) {                                    // keystrokes from the phone's terminal
      const pt = ctx.ptys.get(termId(p)); if (!pt) return json(res, 404, {});
      pt.proc.write(String(b.data || '')); return json(res, 200, { ok: true });
    }
    json(res, 404, {});
  }

  function start() {
    if (server) return;
    server = http.createServer((req, res) => handle(req, res).catch(e => { try { json(res, 500, { error: String(e.message || e) }); } catch { } }));
    server.on('error', e => { console.error('phone server', e.message); server = null; });
    server.listen(PORT, ctx.bind || '0.0.0.0');
  }
  function stop() { if (!server) return; for (const c of clients) try { c.res.end(); } catch { } clients.clear(); server.close(); server = null; }

  // addresses the phone can use: Wi-Fi/LAN and Tailscale (100.64.0.0/10)
  function urls() {
    const out = [];
    for (const [name, list] of Object.entries(os.networkInterfaces())) for (const a of list || []) {
      if (a.family !== 'IPv4' || a.internal) continue;
      const ts = /^100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\./.test(a.address);
      if (ts || /^(en|bridge)/.test(name)) out.push({ kind: ts ? 'Tailscale (anywhere)' : 'Wi-Fi (same network)', url: `http://${a.address}:${PORT}/?t=${cfg.token}` });
    }
    return out.sort((a, b) => a.kind.localeCompare(b.kind)).reverse();
  }

  if (cfg.enabled && !ctx.passive) start();        // passive: a CLI run (--pair / --status) beside the running service
  return {
    broadcast(snap) { if (!clients.size) return; const v = view(snap); for (const c of clients) try { send(c.res, 'telemetry', v); } catch { } },
    status: () => ({ enabled: cfg.enabled, running: !!server, port: PORT, urls: urls(), clients: clients.size,
                     peerWatching: Date.now() - peerPolledAt < 15000, peer: peer ? { online: peer.online(), label: peer.label() } : null }),
    setEnabled(v) { cfg.enabled = !!v; save(); v ? start() : stop(); },
    regenerate() { cfg.token = crypto.randomBytes(18).toString('base64url'); save(); for (const c of clients) try { c.res.end(); } catch { } clients.clear(); },
  };
}

module.exports = { createPhoneServer };
