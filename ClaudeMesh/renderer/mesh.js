// Osmos-style mesh. The core is an Attractor; each Claude session is a translucent mote that
// orbits it and ejects matter (tokens) into it. The cursor behaves like the player's mote:
// moving ejects a propulsion trail and pushes the medium, touching small motes absorbs them,
// clicking empty space fires an ejection burst (shockwave); pinch / scroll zooms the view.
const FAMILY_COLOR = { fable: '#ff5c8a', mythos: '#ff5c8a', opus: "#7d7bff", sonnet: "#2ef0a8", haiku: '#ffb347', other: '#9fb4d8' };
const HEALTH_COLOR = { working: '#7dffc4', idle: '#7c86b8', quiet: '#ffcf6b', waiting: '#ffcf6b', error: '#ff6b7d' };
const BG = ['#1a1850', '#0d0b30', '#05041a'], NEB = ['#5a3cff', '#2a8cff', '#b04cff'];
const AMBIENT_COLORS = ['#5ec8ff', '#3ee6b0', '#8f8bff', '#ff6b8b', '#b98bff', '#4fa3ff'];

function rgb(hex) { const n = parseInt(hex.slice(1), 16); return [n >> 16, (n >> 8) & 255, n & 255]; }
function hexA(hex, a) { const [r, g, b] = rgb(hex); return `rgba(${r},${g},${b},${a})`; }
function mix(hex, t) { const c = rgb(hex).map(v => Math.round(v + (255 - v) * t)); return "#" + c.map(v => v.toString(16).padStart(2, "0")).join(""); }
// session bubbles darken as today's spend grows ($0 → bright, ~$40+ → darkest); resets at midnight with todayUsd
function spendDim(s) { return Math.min(.62, Math.log10(1 + Math.max(0, s.todayUsd || 0)) / Math.log10(41) * .62); }
function darken(hex, k) { return '#' + rgb(hex).map(v => Math.round(v * (1 - k)).toString(16).padStart(2, '0')).join(''); }
function sessionColor(s) { return darken(FAMILY_COLOR[s.family] || '#9fb4d8', spendDim(s)); }
function hash(s) { let h = 2166136261; for (const c of s) h = Math.imul(h ^ c.charCodeAt(0), 16777619); return (h >>> 0) / 4294967295; }
function rnd(seed) { let s = seed * 1e9 | 0 || 1; return () => ((s = Math.imul(s ^ (s >>> 15), 2246822507) ^ Math.imul(s ^ (s >>> 13), 3266489909)) >>> 0) / 4294967296; }

// Pre-rendered translucent mote sprite (rim-lit cell with nebula speckle), cached per colour.
const spriteCache = new Map();
function moteSprite(color, speckles = true) {
  const key = color + speckles;
  if (spriteCache.has(key)) return spriteCache.get(key);
  const S = 256, c = document.createElement('canvas'); c.width = c.height = S;
  const x = c.getContext('2d'), R = S / 2 - 6, C = S / 2;
  const body = x.createRadialGradient(C, C, 0, C, C, R);
  body.addColorStop(0, hexA(color, .22)); body.addColorStop(.62, hexA(color, .26));
  body.addColorStop(.84, hexA(color, .6)); body.addColorStop(.96, hexA(mix(color, .6), 1)); body.addColorStop(1, hexA(color, 0));
  x.fillStyle = body; x.beginPath(); x.arc(C, C, R, 0, 7); x.fill();
  if (speckles) {
    const r = rnd(hash(color));
    for (let i = 0; i < 260; i++) {
      const a = r() * 6.283, d = Math.pow(r(), .7) * R * .86;
      x.fillStyle = hexA(r() > .8 ? '#ffffff' : mix(color, r() * .6), .25 + r() * .55);
      x.beginPath(); x.arc(C + Math.cos(a) * d, C + Math.sin(a) * d, .6 + r() * 2.2, 0, 7); x.fill();
    }
  }
  const spec = x.createRadialGradient(C - R * .38, C - R * .42, 0, C - R * .38, C - R * .42, R * .55);
  spec.addColorStop(0, 'rgba(255,255,255,.35)'); spec.addColorStop(1, 'rgba(255,255,255,0)');
  x.fillStyle = spec; x.beginPath(); x.arc(C, C, R, 0, 7); x.fill();
  spriteCache.set(key, c); return c;
}
function glowSprite(color) {
  const key = 'g' + color; if (spriteCache.has(key)) return spriteCache.get(key);
  const S = 128, c = document.createElement('canvas'); c.width = c.height = S;
  const x = c.getContext('2d'), g = x.createRadialGradient(64, 64, 0, 64, 64, 64);
  g.addColorStop(0, hexA(color, .9)); g.addColorStop(.25, hexA(color, .35)); g.addColorStop(1, hexA(color, 0));
  x.fillStyle = g; x.fillRect(0, 0, S, S); spriteCache.set(key, c); return c;
}

class Mesh {
  constructor(canvas, { onSelect, onHub }) {
    this.c = canvas; this.ctx = canvas.getContext('2d');
    this.onSelect = onSelect; this.onHub = onHub; this.onHold = null;   // onHold(sid, clientX, clientY): press-and-hold
    this.nodes = new Map(); this.matter = []; this.fx = [];
    this.selected = null; this.hover = null;
    this.hub = { x: 0, y: 0, tps: 0, usd: 0, count: 0, mass: 0 };
    this.time = 0;
    this.cursor = { x: -999, y: -999, px: -999, py: -999, vx: 0, vy: 0, r: 11, in: false, mass: 0 };
    this.stars = [0, 1, 2].map(layer => Array.from({ length: 90 + layer * 40 }, () => ({ x: Math.random(), y: Math.random(), s: Math.random(), tw: Math.random() * 6.28, layer })));
    this.ambient = [];                    // bubbles blown by the sessions (bubbles.js)
    canvas.addEventListener('mousemove', e => { this.lastInput = performance.now(); this.move(e); });
    canvas.addEventListener('mouseenter', () => { this.cursor.in = true; });
    canvas.addEventListener('mouseleave', () => { this.cursor.in = false; this.hover = null; });
    canvas.addEventListener('mousedown', e => this.down(e));
    this.cam = { s: 1, tx: 0, ty: 0, ts: 1, ttx: 0, tty: 0 };      // view zoom (t* = target, eased toward)
    canvas.addEventListener('wheel', e => { this.lastInput = performance.now(); this.wheelZoom(e); }, { passive: false });
    canvas.addEventListener('dblclick', e => { const h = this.hit(e); if (!h) Object.assign(this.cam, { ts: 1, ttx: 0, tty: 0 }); else if (h !== 'hub') this.unpin(h); });
    addEventListener('mousemove', e => {
      if (this.drag) return this.dragMove(e);
      const p = this.pan; if (!p) return;
      if (!p.moved && Math.hypot(e.clientX - p.x0, e.clientY - p.y0) > 4) { p.moved = true; this.c.style.cursor = 'grabbing'; }
      if (p.moved) { this.cam.ttx = p.tx + e.clientX - p.x0; this.cam.tty = p.ty + e.clientY - p.y0; this.clampCam(); this.cam.tx = this.cam.ttx; this.cam.ty = this.cam.tty; }
    });
    addEventListener('mouseup', () => {
      this.dragEnd();
      const p = this.pan; this.pan = null; if (!p) return;
      this.c.style.cursor = ''; if (!p.moved) this.burst(p.e);
    });
    new ResizeObserver(() => this.resize()).observe(canvas);
    this.resize();
    requestAnimationFrame(t => this.frame(t));
  }

  newAmbient(anywhere) {
    const v = this.cam ? this.view() : { x0: 0, y0: 0, x1: this.w || 1200, y1: this.h || 700 }, w = v.x1 - v.x0, h = v.y1 - v.y0;
    const r = 2 + Math.pow(Math.random(), 2.4) * 16, edge = Math.random() * 4 | 0;
    return {
      x: v.x0 + (anywhere ? Math.random() * w : edge === 0 ? -20 : edge === 1 ? w + 20 : Math.random() * w),
      y: v.y0 + (anywhere ? Math.random() * h : edge === 2 ? -20 : edge === 3 ? h + 20 : Math.random() * h),
      vx: (Math.random() - .5) * 8, vy: (Math.random() - .5) * 8, r, depth: .3 + Math.random() * .7,
      color: AMBIENT_COLORS[Math.random() * AMBIENT_COLORS.length | 0], born: this.time,
    };
  }

  // ambient medium: motes drift, wrap back in from the edges, and part around the cursor
  simAmbient(dt, cspd) {
    const cur = this.cursor, v = this.view(), target = Math.round((this.pixel ? 24 : 40) * Math.min(2.2, 1 / (this.cam.s * this.cam.s)) ** .5);
    while (this.ambient.length < target) this.ambient.push(this.newAmbient(false));
    if (this.ambient.length > target) this.ambient.length = target;
    for (const m of this.ambient) {
      if (cur.in) {
        const dx = m.x - cur.x, dy = m.y - cur.y, d = Math.hypot(dx, dy), R = 70 + m.r;
        if (d < R && d > 1) { const f = (1 - d / R) * (40 + cspd * .4); m.vx += dx / d * f * dt; m.vy += dy / d * f * dt; }
      }
      const damp = Math.pow(.7, dt); m.vx = m.vx * damp + (Math.random() - .5) * 4 * dt; m.vy = m.vy * damp + (Math.random() - .5) * 4 * dt;
      m.x += m.vx * dt; m.y += m.vy * dt;
      if (m.x < v.x0 - 40) m.x = v.x1 + 30; else if (m.x > v.x1 + 40) m.x = v.x0 - 30;
      if (m.y < v.y0 - 40) m.y = v.y1 + 30; else if (m.y > v.y1 + 40) m.y = v.y0 - 30;
    }
  }

  // ---- zoom: pinch or scroll, anchored at the pointer, 0.65× (wide) … 2× (close) ----
  wheelZoom(e) {
    e.preventDefault();
    const r = this.c.getBoundingClientRect(), px = e.clientX - r.left, py = e.clientY - r.top, c = this.cam;
    const k = Math.exp(-e.deltaY * (e.ctrlKey ? .012 : .0018));          // ctrlKey = trackpad pinch
    const s = Math.min(2, Math.max(.65, c.ts * k));
    c.ttx = px - (px - c.ttx) * s / c.ts; c.tty = py - (py - c.tty) * s / c.ts; c.ts = s;
    this.clampCam();
  }

  // the map extends 40% beyond the window on every side; the view can roam inside that
  clampCam() {
    const c = this.cam, W = this.w, H = this.h, M = .4;
    const fit = (t, size) => {
      const lo = -(1 + M) * size * c.ts + size, hi = M * size * c.ts;      // view edge stays within [-M·size, (1+M)·size]
      return lo > hi ? (lo + hi) / 2 : Math.min(hi, Math.max(lo, t));
    };
    c.ttx = fit(c.ttx, W); c.tty = fit(c.tty, H);
  }

  stepCam(dt) {
    const c = this.cam, a = 1 - Math.pow(.0005, dt);
    c.s += (c.ts - c.s) * a; c.tx += (c.ttx - c.tx) * a; c.ty += (c.tty - c.ty) * a;
  }

  view() { const c = this.cam; return { x0: -c.tx / c.s, y0: -c.ty / c.s, x1: (this.w - c.tx) / c.s, y1: (this.h - c.ty) / c.s }; }
  toWorld(e) { const r = this.c.getBoundingClientRect(), c = this.cam; return { x: (e.clientX - r.left - c.tx) / c.s, y: (e.clientY - r.top - c.ty) / c.s }; }

  drawBackdrop(ctx, W, H, parX, parY) {
    const d = devicePixelRatio || 1, key = W + 'x' + H + '@' + d;
    if (!this.bg || this.bg.key !== key) {
      const mk = () => { const c = document.createElement('canvas'); c.width = Math.ceil(W * d); c.height = Math.ceil(H * d); const x = c.getContext('2d'); x.setTransform(d, 0, 0, d, 0, 0); return [c, x]; };
      const [base, bx] = mk();
      const g = bx.createRadialGradient(W * .5, H * .45, 0, W * .5, H * .5, Math.max(W, H) * .8);
      g.addColorStop(0, BG[0]); g.addColorStop(.45, BG[1]); g.addColorStop(1, BG[2]); bx.fillStyle = g; bx.fillRect(0, 0, W, H);
      for (const [nx, ny, col] of [[.22, .28, NEB[0]], [.8, .72, NEB[1]], [.7, .2, NEB[2]]]) {
        const ng = bx.createRadialGradient(W * nx, H * ny, 0, W * nx, H * ny, Math.max(W, H) * .42);
        ng.addColorStop(0, hexA(col, .28)); ng.addColorStop(1, hexA(col, 0)); bx.fillStyle = ng; bx.fillRect(0, 0, W, H);
      }
      // all star layers in one canvas with a margin, so parallax is one blit (no wrap-around tiling)
      const M = 160, sc = document.createElement('canvas'); sc.width = Math.ceil((W + 2 * M) * d); sc.height = Math.ceil((H + 2 * M) * d);
      const sx = sc.getContext('2d'); sx.setTransform(d, 0, 0, d, 0, 0);
      for (const layer of this.stars) for (const s of layer) {
        sx.fillStyle = `rgba(200,210,255,${(.15 + .5 * s.s) * .8})`; const z = .5 + s.s * (s.layer * .5 + .6);
        sx.fillRect(s.x * (W + 2 * M), s.y * (H + 2 * M), z, z);
      }
      const layers = { c: sc, M };
      const tw = this.stars.flat().filter(s => s.s > .55).slice(0, 45);           // the brightest ones twinkle live
      this.bg = { key, base, layers, tw };
    }
    ctx.drawImage(this.bg.base, 0, 0, W, H);
    { const { c, M } = this.bg.layers, ox = Math.max(-M, Math.min(M, -parX * 12 + this.cam.tx * .12)), oy = Math.max(-M, Math.min(M, -parY * 12 + this.cam.ty * .12));
      ctx.drawImage(c, ox - M, oy - M, W + 2 * M, H + 2 * M); }
    for (const s of this.bg.tw) {
      const { M } = this.bg.layers, ox = Math.max(-M, Math.min(M, -parX * 12 + this.cam.tx * .12)), oy = Math.max(-M, Math.min(M, -parY * 12 + this.cam.ty * .12));
      const x = s.x * (W + 2 * M) - M + ox, y = s.y * (H + 2 * M) - M + oy;
      const a = (.15 + .5 * s.s) * .5 * Math.max(0, Math.sin(this.time * (1 + s.s) + s.tw));
      if (a > .02) { ctx.fillStyle = `rgba(220,228,255,${a})`; const z = 1 + s.s * (s.layer * .5 + .6); ctx.fillRect(x - .3, y - .3, z, z); }
    }
  }

  resize() {
    if (this.pixel) return this.resizePixel();
    const r = this.c.getBoundingClientRect(), d = devicePixelRatio || 1;
    this.w = r.width; this.h = r.height;
    this.c.width = Math.max(1, r.width * d); this.c.height = Math.max(1, r.height * d);
    this.ctx.setTransform(d, 0, 0, d, 0, 0);
    this.hub.x = this.w / 2; this.hub.y = this.motherY();
    if (this.cam) { this.clampCam(); Object.assign(this.cam, { tx: this.cam.ttx, ty: this.cam.tty }); }
  }

  update(snap) {
    const seen = new Set();
    for (const s of snap.sessions) {
      seen.add(s.sid);
      let n = this.nodes.get(s.sid);
      if (!n) {
        const seed = hash(s.sid);
        n = { sid: s.sid, seed, x: this.hub.x, y: this.hub.y, vx: 0, vy: 0, born: this.time, acc: 0, s, flash: 0, spin: seed * 6.28, r: 10, swell: 0, phase: 0 };
        const rr = rnd(seed); n.inner = Array.from({ length: 22 }, () => ({ a: rr() * 6.28, d: .15 + rr() * .7, sp: (rr() - .5) * 1.4, sz: .7 + rr() * 1.8 }));
        this.nodes.set(s.sid, n);
      }
      if (n.s.health !== 'error' && s.health === 'error') n.flash = 1;
      n.s = s; n.dying = 0;   // a session that reappears after a missed snapshot is alive again
    }
    for (const [sid, n] of this.nodes) if (!seen.has(sid) && !n.dying) n.dying = this.time;
    this.hub.tps = snap.sessions.reduce((a, s) => a + s.tps, 0);
    this.hub.usd = snap.todayUsd; this.hub.count = snap.sessions.length; this.hub.usage = snap.usage;
    this.hub.working = snap.sessions.filter(s => s.health === 'working').length;
  }

  // home slot on an elliptical orbit around the attractor; slots drift slowly (≈ one lap / 15 min)
  home(n, list, i) {
    const lap = this.time * (6.283 / 900);
    const a = -Math.PI / 2 + (i / Math.max(1, list.length)) * 6.283 + lap;
    const ring = list.length > 7 && i % 2 ? .66 : 1;
    const rx = Math.max(this.galaxyRadius() + 110, this.w / 2 - 130) * ring, ry = Math.max(this.galaxyRadius() + 90, this.h / 2 - 90) * ring;
    return { x: this.hub.x + Math.cos(a) * rx, y: this.hub.y + Math.sin(a) * ry };
  }

  command(sid) {
    const n = this.nodes.get(sid); if (!n) return;
    for (let i = 0; i < 26; i++) {
      const a = Math.atan2(n.y - this.hub.y, n.x - this.hub.x) + (Math.random() - .5) * .5, sp = 220 + Math.random() * 160;
      this.matter.push({ x: this.hub.x, y: this.hub.y, vx: Math.cos(a) * sp, vy: Math.sin(a) * sp, r: 1.5 + Math.random() * 2, color: '#ffffff', target: n, life: 3 });
    }
    n.flash = .9;
  }

  nodeRadius(n) { return 22 + Math.min(30, Math.log10(1 + (n.s.totalUsd || 0)) * 10); }

  // Frame governor: 60 fps while you interact (pointer, drag, zoom, pan), 30 otherwise;
  // nothing at all while the window is hidden/minimised or the mesh is off-screen (Grid view).
  frame(now) {
    if (document.hidden || this.offscreen || document.body.dataset.view === 'grid' || !this.w) {
      this.last = this.lastPx = 0; setTimeout(() => requestAnimationFrame(t => this.frame(t)), 250); return;
    }
    const c = this.cam, busy = now - (this.lastInput || 0) < 2500 || this.drag || this.pan || Math.abs(c.s - c.ts) > .002 || Math.abs(c.tx - c.ttx) > .5;
    const fps = busy ? 60 : 30;
    if (now - (this.lastDraw || 0) < 1000 / fps - 3) { requestAnimationFrame(t => this.frame(t)); return; }
    this.lastDraw = now;
    if (this.pixel) { this.framePixel(now); requestAnimationFrame(t => this.frame(t)); return; }
    const realDt = Math.min(.1, (now - (this.last || now)) / 1000); this.last = now;
    const dt = realDt; this.time += dt; this.stepCam(realDt);
    const ctx = this.ctx, W = this.w, H = this.h, cur = this.cursor;
    cur.vx = cur.vx * .85 + (cur.x - cur.px) / Math.max(realDt, .001) * .15;
    cur.vy = cur.vy * .85 + (cur.y - cur.py) / Math.max(realDt, .001) * .15;
    cur.px = cur.x; cur.py = cur.y;
    const cspd = Math.hypot(cur.vx, cur.vy);
    const parX = cur.in ? (cur.x / W - .5) : 0, parY = cur.in ? (cur.y / H - .5) : 0;

    // --- deep indigo nebula backdrop + parallax starfield (pre-rendered; only a few stars twinkle live)
    this.drawBackdrop(ctx, W, H, parX, parY);

    const d = devicePixelRatio || 1, cam = this.cam;
    ctx.setTransform(d * cam.s, 0, 0, d * cam.s, d * cam.tx, d * cam.ty);   // world space from here on
    // --- sessions orbit the mother; their bubbles drift, shimmer and fall into it
    this.simAmbient(dt, cspd);
    this.stepSessions(dt, realDt, cspd);
    this.drawLinks(ctx);
    this.drawBubbles(ctx, parX, parY);

    ctx.globalCompositeOperation = 'lighter';
    this.matter = this.matter.filter(p => {
      p.life -= dt;
      if (p.target) {                      // command matter flies to its session
        if (!this.nodes.has(p.target.sid)) return false;
        const dx = p.target.x - p.x, dy = p.target.y - p.y, d = Math.hypot(dx, dy);
        if (d < this.nodeRadius(p.target)) { p.target.flash = Math.min(1, p.target.flash + .06); return false; }
        p.vx += dx / d * 900 * dt; p.vy += dy / d * 900 * dt; p.vx *= .96; p.vy *= .96;
      } else if (!p.free) {                // attractor gravity
        const dx = this.hub.x - p.x, dy = this.hub.y - p.y, d = Math.max(20, Math.hypot(dx, dy));
        if (d < this.galaxyRadius() * .2 + this.hub.mass) { this.hub.mass = Math.min(14, this.hub.mass + p.r * .25); return false; }
        const G = 90000 / (d * d) + 30; p.vx += dx / d * G * dt; p.vy += dy / d * G * dt;
        p.vx *= Math.pow(.6, dt); p.vy *= Math.pow(.6, dt);
      } else { p.vx *= Math.pow(.25, dt); p.vy *= Math.pow(.25, dt); }
      p.x += p.vx * dt; p.y += p.vy * dt;
      const a = Math.min(1, p.life / 1.2);
      const gs = glowSprite(p.color), sz = p.r * 6;
      ctx.globalAlpha = a; ctx.drawImage(gs, p.x - sz / 2, p.y - sz / 2, sz, sz); ctx.globalAlpha = 1;
      return p.life > 0;
    });
    ctx.globalCompositeOperation = 'source-over';
    if (this.matter.length > 1500) this.matter.splice(0, this.matter.length - 1500);

    this.drawAttractor(realDt);
    for (const n of this.nodes.values()) this.drawMote(n, realDt);
    this.drawMotherInfo();
    this.drawFx(realDt);
    this.drawCursor(realDt, cspd);

    ctx.setTransform(d, 0, 0, d, 0, 0);
    requestAnimationFrame(t => this.frame(t));
  }

  drawAttractor(dt) {
    const ctx = this.ctx, h = this.hub, t = this.time;
    if (!Number.isFinite(h.mass)) h.mass = 0;
    h.mass *= Math.pow(.5, dt);
    const R = 30 + h.mass + Math.sin(t * 1.6) * 1.5;
    ctx.globalCompositeOperation = 'lighter';
    const halo = ctx.createRadialGradient(h.x, h.y, R * .5, h.x, h.y, R * 5);
    halo.addColorStop(0, 'rgba(140,200,255,.30)'); halo.addColorStop(.4, 'rgba(90,120,255,.08)'); halo.addColorStop(1, 'rgba(60,80,255,0)');
    ctx.fillStyle = halo; ctx.beginPath(); ctx.arc(h.x, h.y, R * 5, 0, 6.29); ctx.fill();
    // corona rays
    for (let i = 0; i < 14; i++) {
      const a = i / 14 * 6.283 + t * .05 * (i % 2 ? 1 : -1), L = R * (1.8 + .6 * Math.sin(t * 1.3 + i * 2.1));
      const rg = ctx.createLinearGradient(h.x, h.y, h.x + Math.cos(a) * L, h.y + Math.sin(a) * L);
      rg.addColorStop(0, 'rgba(190,225,255,.35)'); rg.addColorStop(1, 'rgba(190,225,255,0)');
      ctx.strokeStyle = rg; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(h.x, h.y); ctx.lineTo(h.x + Math.cos(a) * L, h.y + Math.sin(a) * L); ctx.stroke();
    }
    // swirling accretion disk
    for (let i = 0; i < 70; i++) {
      const a = i * 2.39996 + t * (.6 + (i % 5) * .08), d = R * (1.1 + (i * 37 % 100) / 100 * 1.1);
      ctx.fillStyle = `rgba(170,215,255,${.25 + .35 * ((i * 13) % 10) / 10})`;
      ctx.fillRect(h.x + Math.cos(a) * d, h.y + Math.sin(a) * d * .9, 1.6, 1.6);
    }
    ctx.globalCompositeOperation = 'source-over';
    const core = ctx.createRadialGradient(h.x - R * .3, h.y - R * .35, 1, h.x, h.y, R);
    core.addColorStop(0, '#ffffff'); core.addColorStop(.3, '#d6ecff'); core.addColorStop(.75, '#6fb6ff'); core.addColorStop(1, 'rgba(80,140,255,.6)');
    ctx.fillStyle = core; ctx.beginPath(); ctx.arc(h.x, h.y, R, 0, 6.29); ctx.fill();
    ctx.strokeStyle = 'rgba(220,240,255,.8)'; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.arc(h.x, h.y, R, 0, 6.29); ctx.stroke();
    if (this.hover === 'hub') { ctx.strokeStyle = 'rgba(255,255,255,.45)'; ctx.setLineDash([2, 5]); ctx.beginPath(); ctx.arc(h.x, h.y, R + 10, t, t + 6.283); ctx.stroke(); ctx.setLineDash([]); }
    ctx.textAlign = 'center'; ctx.fillStyle = '#0a1240';
    ctx.font = '600 15px system-ui'; ctx.fillText(`${Math.round(h.tps)}`, h.x, h.y + 1);
    ctx.font = '500 9px system-ui'; ctx.fillText('tok/s', h.x, h.y + 12);
    ctx.fillStyle = 'rgba(220,230,255,.9)'; ctx.font = '300 13px system-ui';
    ctx.fillText(`$${(h.usd || 0).toFixed(2)}  today`, h.x, h.y + R + 44);
    ctx.fillStyle = 'rgba(160,170,220,.75)'; ctx.font = '300 11px system-ui';
    ctx.fillText(`${h.working || 0} / ${h.count}  active`.split('').join(String.fromCharCode(8202)), h.x, h.y + R + 60);
  }

  drawMote(n, dt) {
    const ctx = this.ctx, s = n.s, t = this.time;
    const alive = n.dying ? Math.max(0, 1 - (t - n.dying) / 2.5) : Math.min(1, (t - n.born) / 1.2);
    let R = this.nodeRadius(n) * (.5 + .5 * alive) * (1 + n.swell * .12);
    let col = sessionColor(s);
    const energy = { working: 1, idle: .62, quiet: .7, waiting: .8, error: .9 }[s.health] ?? .5;
    if (s.health === 'working') R *= 1 + .035 * Math.sin(t * 5 + n.seed * 9);
    n.phase *= Math.pow(.1, dt); n.flash *= Math.pow(.25, dt);
    n.spin += dt * (s.health === 'working' ? .9 : .15) * (n.seed > .5 ? 1 : -1);

    ctx.globalAlpha = alive * (.55 + .45 * energy);
    // outer glow
    ctx.globalCompositeOperation = 'lighter';
    const gs = glowSprite(s.health === 'error' ? '#ff4d6a' : col), G = R * (3 + n.phase * .6 + n.flash * 2);
    ctx.globalAlpha = alive * (.18 + .25 * energy + n.flash * .4);
    ctx.drawImage(gs, n.x - G, n.y - G, G * 2, G * 2);
    ctx.globalCompositeOperation = 'source-over';
    ctx.globalAlpha = alive * (.55 + .45 * energy);
    // translucent body (rotating sprite)
    ctx.save(); ctx.translate(n.x, n.y); ctx.rotate(n.spin);
    ctx.drawImage(moteSprite(col), -R, -R, R * 2, R * 2);
    ctx.restore();
    // swirling interior particles
    ctx.globalCompositeOperation = 'lighter';
    for (const p of n.inner) {
      p.a += p.sp * dt * (s.health === 'working' ? 2.4 : .5) * (1 + n.swell * 2);
      const d = p.d * R * (1 + n.swell * .08 * Math.sin(t * 6 + p.a));
      ctx.fillStyle = hexA(mix(col, .5), .5 + .4 * energy);
      ctx.beginPath(); ctx.arc(n.x + Math.cos(p.a) * d, n.y + Math.sin(p.a) * d, p.sz * (R / 30), 0, 6.29); ctx.fill();
    }
    ctx.globalCompositeOperation = 'source-over';
    // glowing, shimmering rim (the reference's neon border) + a faint outer halo ring
    const shim = .75 + .25 * Math.sin(t * 3.1 + n.seed * 17) * Math.sin(t * 1.7 + n.seed * 5);
    const rimCol = s.health === 'error' && Math.random() > .3 ? '#ff6b7d' : mix(col, .45);
    ctx.save();
    ctx.globalCompositeOperation = 'lighter';
    ctx.shadowColor = rimCol; ctx.shadowBlur = (14 + 10 * energy) * shim + n.flash * 20;
    ctx.strokeStyle = hexA(rimCol, .75 + .25 * energy); ctx.lineWidth = 2.2;
    ctx.beginPath(); ctx.arc(n.x, n.y, R, 0, 6.29); ctx.stroke();
    ctx.shadowBlur = 0; ctx.strokeStyle = `rgba(255,255,255,${.35 * shim})`; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.arc(n.x, n.y, R - .5, 0, 6.29); ctx.stroke();
    const sa = t * (s.health === 'working' ? 1.6 : .5) * (n.seed > .5 ? 1 : -1) + n.seed * 9;   // travelling glint
    ctx.strokeStyle = `rgba(255,255,255,${.55 * shim})`; ctx.lineWidth = 2.4; ctx.lineCap = 'round';
    ctx.beginPath(); ctx.arc(n.x, n.y, R, sa, sa + .55); ctx.stroke();
    ctx.strokeStyle = hexA(col, .18 + .1 * shim); ctx.lineWidth = 1;
    ctx.beginPath(); ctx.arc(n.x, n.y, R + 13, 0, 6.29); ctx.stroke();
    ctx.restore();
    // attention pulse
    if (s.health === 'waiting' || s.health === 'quiet') {
      const k = (t * .8) % 1; ctx.strokeStyle = `rgba(255,207,107,${(1 - k) * .7})`; ctx.lineWidth = 1.5;
      ctx.beginPath(); ctx.arc(n.x, n.y, R + 4 + k * 26, 0, 6.29); ctx.stroke();
    }
    // context gauge (thin outer arc)
    const ctxPct = s.ctxMax ? Math.min(1, s.ctx / s.ctxMax) : 0;
    ctx.strokeStyle = 'rgba(255,255,255,.07)'; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(n.x, n.y, R + 7, 0, 6.29); ctx.stroke();
    ctx.strokeStyle = ctxPct > .85 ? '#ff6b7d' : ctxPct > .6 ? '#ffcf6b' : hexA(mix(col, .3), .9);
    ctx.beginPath(); ctx.arc(n.x, n.y, R + 7, -Math.PI / 2, -Math.PI / 2 + ctxPct * 6.283); ctx.stroke();
    // sub-agents = small motes in orbit
    const subs = Math.min(8, s.subagents || 0);
    for (let i = 0; i < subs; i++) {
      const a = t * 1.3 + i / subs * 6.283 + n.seed * 5, rr = R + 22, r = 4.5;
      ctx.drawImage(moteSprite(col, false), n.x + Math.cos(a) * rr - r, n.y + Math.sin(a) * rr - r, r * 2, r * 2);
    }
    if (this.selected === n.sid) {
      ctx.strokeStyle = 'rgba(255,255,255,.8)'; ctx.lineWidth = 1; ctx.setLineDash([2, 5]);
      ctx.beginPath(); ctx.arc(n.x, n.y, R + 15, t * .6, t * .6 + 6.283); ctx.stroke(); ctx.setLineDash([]);
    }
    // label pill (dark glass, like the reference's node tags)
    ctx.globalAlpha = alive * (.85 + .15 * n.swell);
    const name = s.name.length > 30 ? s.name.slice(0, 29) + '…' : s.name;
    const l2 = `${(s.model || '—').replace('claude-', '')}  ·  $${s.todayUsd.toFixed(2)}`;
    const l3 = (s.health === 'working' ? `working · ${Math.round(s.tps)} tok/s` : s.health) + (s.owner ? '  ·  in-app' : '');
    ctx.font = '600 12px system-ui, "Apple SD Gothic Neo"'; const w1 = ctx.measureText(name).width;
    ctx.font = '400 10.5px system-ui'; const pw = Math.max(w1, ctx.measureText(l2).width, ctx.measureText(l3).width) + 22;
    const py = n.y + R + 20, ph = 46;
    ctx.fillStyle = 'rgba(12,12,40,.72)'; ctx.strokeStyle = hexA(mix(col, .3), .35 + .25 * n.swell); ctx.lineWidth = 1;
    ctx.beginPath(); ctx.roundRect(n.x - pw / 2, py, pw, ph, 8); ctx.fill(); ctx.stroke();
    ctx.textAlign = 'center';
    ctx.font = '600 12px system-ui, "Apple SD Gothic Neo"'; ctx.fillStyle = '#f2f3ff'; ctx.fillText(name, n.x, py + 15);
    ctx.font = '400 10.5px system-ui'; ctx.fillStyle = hexA(mix(col, .35), .95); ctx.fillText(l2, n.x, py + 29);
    ctx.fillStyle = HEALTH_COLOR[s.health] || '#9fb4d8'; ctx.fillText(l3, n.x, py + 41);
    ctx.globalAlpha = 1;
  }

  drawFx(dt) {
    const ctx = this.ctx;
    this.fx = this.fx.filter(f => {
      f.t += dt;
      if (f.kind === 'absorb') {          // mote drawn into the cursor
        const k = Math.min(1, f.t / .35), tx = f.to ? f.to.x : this.cursor.x, ty = f.to ? f.to.y : this.cursor.y; const x = f.x + (tx - f.x) * k, y = f.y + (ty - f.y) * k;
        ctx.globalAlpha = 1 - k; const S = f.r * BUBBLE_SPRITE_K * (1 - k) + .1; ctx.drawImage(bubbleSprite(f.color), x - S, y - S, S * 2, S * 2); ctx.globalAlpha = 1;
        return k < 1;
      }
      if (f.kind === 'ko') {              // a rival was absorbed
        const k = f.t / 1.6; ctx.strokeStyle = hexA(f.color, (1 - k) * .9); ctx.lineWidth = 2;
        for (let i = 0; i < 2; i++) { ctx.beginPath(); ctx.arc(f.x, f.y, 10 + k * (70 + i * 40), 0, 6.29); ctx.stroke(); }
        ctx.fillStyle = hexA(f.color, 1 - k); ctx.font = '500 12px system-ui'; ctx.textAlign = 'center';
        ctx.fillText(f.by === 'you' ? `${f.name} absorbed!` : `${f.by} ate ${f.name}`, f.x, f.y - 20 - k * 16); return k < 1;
      }
      if (f.kind === 'eaten') {           // the cursor mote was swallowed
        const k = f.t / 1.2; ctx.strokeStyle = `rgba(255,92,108,${(1 - k) * .8})`; ctx.lineWidth = 2;
        ctx.beginPath(); ctx.arc(f.x, f.y, 6 + k * 50, 0, 6.29); ctx.stroke();
        ctx.fillStyle = `rgba(255,170,180,${1 - k})`; ctx.font = '300 11px system-ui'; ctx.textAlign = 'center';
        ctx.fillText('A B S O R B E D', f.x, f.y - 16 - k * 10); return k < 1;
      }
      if (f.kind === 'wave') {            // ejection shockwave
        const k = f.t / .7; ctx.strokeStyle = `rgba(190,220,255,${(1 - k) * .55})`; ctx.lineWidth = 2 * (1 - k) + .5;
        ctx.beginPath(); ctx.arc(f.x, f.y, 8 + k * 160, 0, 6.29); ctx.stroke(); return k < 1;
      }
      return false;
    });
  }

  drawCursor(dt, spd) {
    const c = this.cursor, ctx = this.ctx; if (!c.in) return;
    if (spd > 60 && Math.random() < Math.min(1, spd / 600)) {     // propulsion trail
      const a = Math.atan2(-c.vy, -c.vx) + (Math.random() - .5) * .7, sp = 40 + Math.random() * 60;
      this.matter.push({ x: c.x + Math.cos(a) * c.r, y: c.y + Math.sin(a) * c.r, vx: Math.cos(a) * sp, vy: Math.sin(a) * sp, r: .8 + Math.random(), color: '#9fd8ff', life: 1.2, free: true });
    }
    ctx.globalCompositeOperation = 'lighter';
    ctx.globalAlpha = .35; ctx.drawImage(glowSprite('#7cc4ff'), c.x - c.r * 2.6, c.y - c.r * 2.6, c.r * 5.2, c.r * 5.2); ctx.globalAlpha = 1;
    ctx.globalCompositeOperation = 'source-over';
    ctx.drawImage(moteSprite('#7cc4ff', false), c.x - c.r, c.y - c.r, c.r * 2, c.r * 2);
    ctx.strokeStyle = 'rgba(210,235,255,.85)'; ctx.lineWidth = 1; ctx.beginPath(); ctx.arc(c.x, c.y, c.r, 0, 6.29); ctx.stroke();
  }

  hit(e) {
    const { x, y } = this.toWorld(e);
    if (Math.hypot(x - this.hub.x, y - this.hub.y) < this.motherRadius()) return 'hub';
    for (const n of this.nodes.values()) if (!n.dying && Math.hypot(x - n.x, y - n.y) < this.nodeRadius(n) + 12) return n.sid;
    return null;
  }
  move(e) {
    const p = this.toWorld(e); this.cursor.x = p.x; this.cursor.y = p.y;
    if (this.drag) return;
    if (this.cursor.px < -900) { this.cursor.px = this.cursor.x; this.cursor.py = this.cursor.y; }
    this.cursor.in = true; this.hover = this.hit(e);
    this.c.style.cursor = this.hover && this.hover !== 'hub' ? 'grab' : this.hover === 'hub' ? 'pointer' : '';
  }
  down(e) {
    const h = this.hit(e);
    if (h === 'hub') return this.onHub();
    if (h) { if (e.button === 0) this.dragStart(e, h); return; }   // click = select, drag = move its orbit
    if (e.button === 0) this.pan = { x0: e.clientX, y0: e.clientY, tx: this.cam.ttx, ty: this.cam.tty, moved: false, e };   // drag empty space = pan
  }

  // a click (no drag) on empty space: Osmos-style ejection burst + shockwave, and deselect
  burst(e) {
    const { x, y } = this.toWorld(e);
    this.fx.push({ kind: 'wave', x, y, t: 0 });
    for (let i = 0; i < 26; i++) {
      const a = Math.random() * 6.283, sp = 80 + Math.random() * 200;
      this.matter.push({ x, y, vx: Math.cos(a) * sp, vy: Math.sin(a) * sp, r: .8 + Math.random() * 1.6, color: '#9fd8ff', life: 1.4, free: true });
    }
    const push = (o, k) => { const dx = o.x - x, dy = o.y - y, d = Math.hypot(dx, dy); if (d < 220 && d > 1) { o.vx += dx / d * (1 - d / 220) * k; o.vy += dy / d * (1 - d / 220) * k; } };
    for (const m of this.ambient) push(m, 260);
    this.onSelect(null);
  }
}
window.Mesh = Mesh;
window.FAMILY_COLOR = FAMILY_COLOR;
