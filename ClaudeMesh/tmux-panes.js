// Claude sessions that live in tmux panes (the mini keeps them in tmux gcx:*, so they survive app restarts).
// They get the same live terminal as in-app sessions: the pane is streamed only while someone watches it
// (tmux pipe-pane → file → ptyBus 'data'), opened on a snapshot of its screen (capture-pane), typed into with
// send-keys and resized with resize-window. Ported from server/mesh-server.js (the old Windows server).
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile, execFileSync } = require('child_process');
const { StringDecoder } = require('string_decoder');

module.exports = function createPanes({ ptyBus }) {
  const BIN = ['/opt/homebrew/bin/tmux', '/usr/local/bin/tmux'].find(p => fs.existsSync(p)) || 'tmux';
  const DIR = path.join(os.tmpdir(), 'gcx-mesh-panes');
  // an app started from the Dock has no locale: tmux then turns tabs (and any non-ASCII text) into '_' — always UTF-8
  const env = { ...process.env, LANG: 'en_US.UTF-8', LC_CTYPE: 'en_US.UTF-8' };
  const tmux = args => new Promise(res => execFile(BIN, args, { timeout: 10000, env }, (err, so) => res(err ? null : String(so))));
  const tmuxSync = args => { try { return execFileSync(BIN, args, { timeout: 10000, env }).toString(); } catch { return ''; } };

  let panes = new Map();          // pane id (%N) -> { id, pid, tty, cols, rows, title }
  const streams = new Map();      // pane id -> { n, file, pos, dec, timer }

  async function refresh() {
    const out = await tmux(['list-panes', '-a', '-F', '#{pane_id}|#{pane_pid}|#{pane_tty}|#{pane_width}|#{pane_height}|#{session_name}:#{window_name}']);
    if (out == null) { if (panes.size) gone(new Map()); return; }   // no tmux server
    const m = new Map();
    for (const l of out.split('\n')) {
      const [id, pid, tty, w, h, ...rest] = l.split('|'); if (!id) continue;
      const title = rest.join('|');
      const p = { id, pid: +pid, tty, cols: +w || 120, rows: +h || 32, title };
      const o = panes.get(id);
      if (o && (o.cols !== p.cols || o.rows !== p.rows)) ptyBus.emit('resize', id, p.cols, p.rows);
      m.set(id, p);
    }
    gone(m);
  }
  function gone(m) {
    for (const id of streams.keys()) if (!m.has(id)) { ptyBus.emit('exit', id); stop(id); }
    panes = m;
  }

  // keystrokes: printable runs as literal text, control bytes one by one (Esc, arrows, ^C …)
  function write(id, data) {
    const s = String(data || ''); if (!s || !panes.has(id)) return;
    let run = '';
    const flush = () => { if (run) { tmuxSync(['send-keys', '-t', id, '-l', run]); run = ''; } };
    for (const ch of s) {
      const c = ch.codePointAt(0);
      if (c < 32 || c === 127) { flush(); tmuxSync(['send-keys', '-t', id, '-H', c.toString(16).padStart(2, '0')]); }
      else run += ch;
    }
    flush();
  }

  // what the pane shows now: scrollback, then exactly one screen, then the cursor where tmux has it
  function snapshot(id) {
    const p = panes.get(id); if (!p) return '';
    const hist = tmuxSync(['capture-pane', '-p', '-e', '-t', id, '-S', '-2000', '-E', '-1']).replace(/\n$/, '');
    const screen = tmuxSync(['capture-pane', '-p', '-e', '-t', id]).replace(/\n$/, '').split('\n');
    while (screen.length < p.rows) screen.push('');
    const [cx, cy] = tmuxSync(['display', '-p', '-t', id, '#{cursor_x} #{cursor_y}']).trim().split(' ').map(Number);
    return (hist ? hist.split('\n').join('\r\n') + '\r\n' : '') + screen.slice(-p.rows).join('\x1b[0m\r\n')
      + `\x1b[0m\x1b[${(cy || 0) + 1};${(cx || 0) + 1}H`;
  }

  const paneFile = id => path.join(DIR, id.replace(/[^\w]/g, '_') + '.raw');
  async function open(id) {
    let s = streams.get(id);
    if (s) { s.n++; return; }
    fs.mkdirSync(DIR, { recursive: true });
    const file = paneFile(id); try { fs.unlinkSync(file); } catch { }
    fs.writeFileSync(file, '');
    s = { n: 1, file, pos: 0, dec: new StringDecoder('utf8'), timer: null };
    streams.set(id, s);
    await tmux(['pipe-pane', '-t', id, `cat >> '${file}'`]);
    s.timer = setInterval(() => {
      let st; try { st = fs.statSync(file); } catch { return; }
      if (st.size <= s.pos) return;
      const fd = fs.openSync(file, 'r');
      try {
        const b = Buffer.alloc(Math.min(st.size - s.pos, 1 << 20)); const n = fs.readSync(fd, b, 0, b.length, s.pos); s.pos += n;
        const text = s.dec.write(b.subarray(0, n)); if (text) ptyBus.emit('data', id, text);
      } finally { fs.closeSync(fd); }
      if (s.pos > 64 << 20) { s.pos = 0; fs.truncateSync(file, 0); }     // long watch: don't let the file grow forever
    }, 80);
  }
  function close(id) { const s = streams.get(id); if (s && --s.n <= 0) stop(id); }
  function stop(id) {
    const s = streams.get(id); if (!s) return;
    clearInterval(s.timer); streams.delete(id);
    tmuxSync(['pipe-pane', '-t', id]);                                      // no command = stop piping
    try { fs.unlinkSync(s.file); } catch { }
  }
  function resize(id, c, r) {
    const p = panes.get(id); if (!p) return false;
    if (c !== p.cols || r !== p.rows) { tmuxSync(['resize-window', '-t', id, '-x', String(c), '-y', String(r)]); p.cols = c; p.rows = r; }
    return true;
  }
  // a crash while streaming would leave pipe-pane writing forever → clean up at start
  try { for (const f of fs.readdirSync(DIR)) { tmuxSync(['pipe-pane', '-t', '%' + f.replace(/^_/, '').replace(/\.raw$/, '')]); fs.unlinkSync(path.join(DIR, f)); } } catch { }

  return {
    refresh, write, snapshot, open, close, resize,
    has: id => panes.has(id),
    list: () => [...panes.values()],
    // phone-server.js wants pty-like entries: { proc: { cols, rows, pid, write }, buf }
    proc: id => { const p = panes.get(id); return p && { proc: { cols: p.cols, rows: p.rows, pid: p.pid, write: d => write(id, d) }, get buf() { return snapshot(id); } }; },
    stopAll: () => { for (const id of [...streams.keys()]) stop(id); },
  };
};
