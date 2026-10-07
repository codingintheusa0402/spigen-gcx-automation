// Pixel mode: a lightweight retro renderer for the mesh. The scene is drawn straight into a
// low-res backing store (1 art pixel = PX css px) with integer rects only — no gradients,
// blur or compositing — capped at 30 fps; labels go on a thin full-res text layer.
const PX = 3;
const PAL = {
  bg: '#0b0f2a', bg2: '#141a44', star: '#5f6aa8', star2: '#c2c9ff', white: '#fff1e8', ink: '#05061a',
  fable: '#ff77a8', mythos: '#ff77a8', opus: '#8b7bff', sonnet: '#00e436', haiku: '#ffa300', other: '#83a6d4',
  working: '#00e436', idle: '#6b74b0', quiet: '#ffec27', waiting: '#ffec27', error: '#ff004d', core: '#29adff',
};
const AMB = ['#29adff', '#00e436', '#8b7bff', '#ff004d', '#ff77a8'];

function shade(hex, k) { // k<1 darker, k>1 lighter
  const n = parseInt(hex.slice(1), 16), f = v => Math.max(0, Math.min(255, Math.round(k < 1 ? v * k : v + (255 - v) * (k - 1))));
  return '#' + [n >> 16, (n >> 8) & 255, n & 255].map(v => f(v).toString(16).padStart(2, '0')).join('');
}

Object.assign(Mesh.prototype, {
  setPixel(on) {
    this.pixel = !!on;
    this.c.style.imageRendering = on ? 'pixelated' : '';
    this.matter.length = 0; this.fx.length = 0;   // particles from one renderer must not leak into the other
    if (on) this.ambient.length = Math.min(this.ambient.length, 30);   // ecosystem refills to its own target
    this.resize();
    if (!on && this.tc) this.tctx.clearRect(0, 0, this.tc.width, this.tc.height);
  },

  resizePixel() {
    const r = this.c.getBoundingClientRect();
    this.w = r.width; this.h = r.height;
    this.pw = Math.max(1, Math.ceil(r.width / PX)); this.ph = Math.max(1, Math.ceil(r.height / PX));
    this.c.width = this.pw; this.c.height = this.ph;
    this.ctx.setTransform(1, 0, 0, 1, 0, 0); this.ctx.imageSmoothingEnabled = false;
    this.hub.x = this.w / 2; this.hub.y = this.h / 2;
    if (this.cam) this.clampCam();
    if (!this.tc) {
      this.tc = document.createElement('canvas'); this.tc.id = 'meshText';
      Object.assign(this.tc.style, { position: 'absolute', inset: '0', width: '100%', height: '100%', pointerEvents: 'none' });
      this.c.after(this.tc); this.tctx = this.tc.getContext('2d');
    }
    const d = devicePixelRatio || 1;
    this.tc.width = r.width * d; this.tc.height = r.height * d; this.tctx.setTransform(d, 0, 0, d, 0, 0);
    this.pstars = Array.from({ length: Math.round(this.pw * this.ph / 900) }, () => ({ x: Math.random() * this.pw | 0, y: Math.random() * this.ph | 0, b: Math.random() }));
  },

  // filled pixel disc in art-pixel space
  pdisc(cx, cy, r, color) {
    const x = this.ctx; x.fillStyle = color; cx = Math.round(cx); cy = Math.round(cy); r = Math.max(1, Math.round(r));
    for (let y = -r; y <= r; y++) { const w = Math.floor(Math.sqrt(r * r - y * y + r * .8)); x.fillRect(cx - w, cy + y, w * 2 + 1, 1); }
  },
  pring(cx, cy, r, color, from = 0, to = 6.283) {
    const x = this.ctx; x.fillStyle = color; cx = Math.round(cx); cy = Math.round(cy);
    const steps = Math.max(12, Math.round(r * 7 * (to - from) / 6.283));
    let lx = null, ly = null;
    for (let i = 0; i <= steps; i++) {
      const a = from + (to - from) * i / steps, px = Math.round(cx + Math.cos(a) * r), py = Math.round(cy + Math.sin(a) * r);
      if (px !== lx || py !== ly) x.fillRect(px, py, 1, 1); lx = px; ly = py;
    }
  },

  framePixel(now) {
    if (now - (this.lastPx || 0) < 33) return;           // 30 fps cap
    const realDt = Math.min(.1, (now - (this.lastPx || now)) / 1000); this.lastPx = now;
    const dt = realDt; this.time += dt; this.stepCam(realDt);
    if (this.c.width !== Math.ceil(this.w / PX)) this.resizePixel();
    const ctx = this.ctx, cur = this.cursor, t = this.time, k = this.cam.s / PX;
    cur.vx = cur.vx * .7 + (cur.x - cur.px) / Math.max(realDt, .001) * .3; cur.vy = cur.vy * .7 + (cur.y - cur.py) / Math.max(realDt, .001) * .3;
    cur.px = cur.x; cur.py = cur.y;
    const cspd = Math.hypot(cur.vx, cur.vy);

    // backdrop: flat colour, dithered horizon band, twinkling stars
    ctx.fillStyle = PAL.bg; ctx.fillRect(0, 0, this.pw, this.ph);
    ctx.fillStyle = PAL.bg2;
    for (let y = 0; y < this.ph; y += 2) for (let x = (y / 2) % 2; x < this.pw; x += 4) if (Math.abs(y - this.ph / 2) < this.ph * .18) ctx.fillRect(x, y, 1, 1);
    for (const s of this.pstars) { const on = Math.sin(t * (1 + s.b * 2) + s.x) > -.6; ctx.fillStyle = s.b > .85 && on ? PAL.star2 : PAL.star; if (on || s.b < .5) ctx.fillRect(s.x, s.y, 1, 1); }

    ctx.save(); ctx.translate(Math.round(this.cam.tx / PX), Math.round(this.cam.ty / PX));   // world space (whole art pixels keep it crisp)
    // bubbles (model-coloured, glinting) + sessions orbiting the mother; links as dotted pixel arcs
    this.simAmbient(dt, cspd);
    this.stepSessions(dt, realDt, cspd);
    const live = [...this.nodes.values()].filter(n => !n.dying);
    for (const n of live) {                         // load strands as dotted pixel arcs, ageing with load
      const h = this.heaviness(n), N = Math.max(1, Math.round(1 + h * 23)), st = starColor('outer', .08 + h * .92, 1.7).map(Math.round);
      const col = '#' + st.map(v => v.toString(16).padStart(2, '0')).join(''), off = (t * (n.s.health === 'working' ? 4 + 8 * h : 1)) % 1;
      for (let j = 0; j < N; j++) {
        const g = this.strandGeom(n, j, t), L = Math.hypot(g.x1 - g.x0, g.y1 - g.y0), steps = Math.max(8, L * k / 3 | 0);
        ctx.fillStyle = j === 0 || Math.floor(t * 6 + j) % 3 ? col : shade(col, .55);
        for (let i = 0; i < steps; i++) { const [px, py] = this.strandPoint(g, (i + off) / steps); ctx.fillRect(Math.round(px * k), Math.round(py * k), 1, 1); }
      }
    }
    for (const m of this.ambient) {
      const col = PAL[m.fam] || PAL.other, r = Math.max(1, m.r * k * 1.1), x = m.x * k, y = m.y * k, glint = Math.sin(t * 4.2 + m.ph) > .2;
      this.pdisc(x, y, r, shade(col, .35)); this.pring(x, y, r, glint ? shade(col, 1.4) : col);
    }

    if (this.matter.length > 300) this.matter.splice(0, this.matter.length - 300);
    this.matter = this.matter.filter(p => {
      p.life -= dt;
      if (p.target) {
        if (!this.nodes.has(p.target.sid)) return false;
        const dx = p.target.x - p.x, dy = p.target.y - p.y, d = Math.hypot(dx, dy);
        if (d < this.nodeRadius(p.target)) return false;
        p.vx += dx / d * 900 * dt; p.vy += dy / d * 900 * dt; p.vx *= .96; p.vy *= .96;
      } else if (!p.free) {
        const dx = this.hub.x - p.x, dy = this.hub.y - p.y, d = Math.max(20, Math.hypot(dx, dy));
        if (d < this.galaxyRadius() * .2) { this.hub.mass = Math.min(12, this.hub.mass + .4); return false; }
        const G = 90000 / (d * d) + 30; p.vx += dx / d * G * dt; p.vy += dy / d * G * dt; p.vx *= Math.pow(.6, dt); p.vy *= Math.pow(.6, dt);
      } else { p.vx *= Math.pow(.25, dt); p.vy *= Math.pow(.25, dt); }
      p.x += p.vx * dt; p.y += p.vy * dt;
      ctx.fillStyle = p.color; ctx.fillRect(Math.round(p.x * k), Math.round(p.y * k), 1, 1);
      return p.life > 0;
    });

    // mother bubble: the burning texture drawn chunky into the low-res buffer, still heat-hazing
    const h = this.hub; h.mass *= Math.pow(.6, dt);
    const GR = this.motherRadius(), hx = h.x * k, hy = h.y * k;
    for (let i = 0; i < 2; i++) { const kk = (t * .35 + i / 2) % 1; if (Math.floor(t * 8 + i) % 2) this.pring(hx, hy, GR * k * (1.05 + kk * .6), kk < .5 ? '#ff8a3d' : '#a8336b'); }
    this.drawMother(ctx, Math.round(hx), Math.round(hy), Math.round(GR * k), false);
    ctx.fillStyle = PAL.white; const cw = Math.max(1, Math.round(GR * k * (.1 + .03 * Math.sin(t * 6))));
    ctx.fillRect(Math.round(hx - cw), Math.round(hy - cw), cw * 2, cw * 2);
    if (this.hover === 'hub') this.pring(hx, hy, GR * k * 1.15, PAL.white, t * 2, t * 2 + 4.5);

    // session motes: dithered body, outline, inner sparkles, ctx gauge
    for (const n of this.nodes.values()) {
      const s = n.s, col = shade(s.serverUses ? '#d84cf2' : PAL[s.family] || PAL.other, 1 - spendDim(s)), alive = n.dying ? Math.max(0, 1 - (t - n.dying) / 2.5) : Math.min(1, (t - n.born) / 1.2);
      if (alive <= 0) continue;
      const R = this.nodeRadius(n) * k * (.5 + .5 * alive) * (1 + n.swell * .12), x = n.x * k, y = n.y * k;
      const dim = s.health === 'idle' ? .55 : 1;
      this.pdisc(x, y, R, shade(col, .28 * dim));
      ctx.fillStyle = shade(col, .5 * dim);                       // checker dither inner shell
      for (let yy = -R; yy <= R; yy++) for (let xx = -R; xx <= R; xx++) {
        const d2 = xx * xx + yy * yy; if (d2 < (R - 1) * (R - 1) && d2 > (R * .62) ** 2 && ((xx + yy) & 1)) ctx.fillRect(Math.round(x + xx), Math.round(y + yy), 1, 1);
      }
      const spin = t * (s.health === 'working' ? 1.6 : .3) * (n.seed > .5 ? 1 : -1);
      ctx.fillStyle = s.health === 'working' ? PAL.white : shade(col, 1.35);
      for (const p of n.inner.slice(0, 9)) { const a = p.a + spin * p.sp, d = p.d * (R - 2); ctx.fillRect(Math.round(x + Math.cos(a) * d), Math.round(y + Math.sin(a) * d), 1, 1); }
      const err = s.health === 'error' && Math.floor(t * 6) % 2;
      this.pring(x, y, R, err ? PAL.error : shade(col, dim < 1 ? .8 : 1.15));
      ctx.fillStyle = PAL.white; ctx.fillRect(Math.round(x - R * .45), Math.round(y - R * .55), 2, 1); ctx.fillRect(Math.round(x - R * .55), Math.round(y - R * .45), 1, 1);
      const pct = s.ctxMax ? Math.min(1, s.ctx / s.ctxMax) : 0;
      this.pring(x, y, R + 3, pct > .85 ? PAL.error : pct > .6 ? PAL.quiet : shade(col, .9), -Math.PI / 2, -Math.PI / 2 + pct * 6.283);
      if ((s.health === 'waiting' || s.health === 'quiet') && Math.floor(t * 3) % 2) this.pring(x, y, R + 6, PAL.quiet);
      for (let i = 0, subs = Math.min(6, s.subagents || 0); i < subs; i++) {
        const a = t * 1.3 + i / subs * 6.283; ctx.fillStyle = col;
        ctx.fillRect(Math.round(x + Math.cos(a) * (R + 7)) - 1, Math.round(y + Math.sin(a) * (R + 7)) - 1, 2, 2);
      }
      if (this.selected === n.sid) {     // blinking corner brackets
        if (Math.floor(t * 3) % 2) { ctx.fillStyle = PAL.white; const b = Math.round(R + 9), cx = Math.round(x), cy = Math.round(y);
          for (const [sx, sy] of [[-1, -1], [1, -1], [-1, 1], [1, 1]]) { ctx.fillRect(cx + sx * b - (sx > 0 ? 2 : 0), cy + sy * b, 3, 1); ctx.fillRect(cx + sx * b, cy + sy * b - (sy > 0 ? 2 : 0), 1, 3); } }
      }
    }

    // cursor: pixel ring + crosshair
    if (cur.in) {
      const cx = cur.x * k, cy = cur.y * k, r = cur.r * k;
      this.pdisc(cx, cy, r, '#1d5fa0'); this.pring(cx, cy, r, '#7fd4ff');
      ctx.fillStyle = PAL.white; ctx.fillRect(Math.round(cx), Math.round(cy), 1, 1);
      if (cspd > 60) { ctx.fillStyle = '#7fd4ff'; for (let i = 1; i <= 3; i++) ctx.fillRect(Math.round(cx - cur.vx * k * .02 * i * 2), Math.round(cy - cur.vy * k * .02 * i * 2), 1, 1); }
    }
    for (const f of this.fx) if (f.kind === 'wave') { f.t += realDt; const kk = f.t / .7; if (kk < 1) this.pring(f.x * k, f.y * k, (8 + kk * 160) * k, kk < .5 ? PAL.white : '#7fd4ff'); }
    this.fx = this.fx.filter(f => f.kind === 'wave' && f.t < .7);

    ctx.restore();
    this.drawPixelText();
  },

  // labels on the full-res layer, in a chunky mono face
  drawPixelText() {
    const x = this.tctx, d = devicePixelRatio || 1, cam = this.cam;
    x.setTransform(d, 0, 0, d, 0, 0); x.clearRect(0, 0, this.w, this.h);
    x.setTransform(d * cam.s, 0, 0, d * cam.s, d * cam.tx, d * cam.ty);
    x.textAlign = 'center'; x.textBaseline = 'alphabetic';
    const txt = (s, X, Y, col, size = 10, bold = true) => {
      x.font = `${bold ? 'bold ' : ''}${size}px Monaco, Menlo, "Apple SD Gothic Neo", monospace`;
      x.fillStyle = PAL.ink; x.fillText(s, Math.round(X) + 1, Math.round(Y) + 1); x.fillStyle = col; x.fillText(s, Math.round(X), Math.round(Y));
    };
    const h = this.hub, GR = this.motherRadius(); let below = 0;
    { const L = (h.usage && h.usage.limits) || [], RH = 32, H = 22 + Math.max(1, L.length) * RH, W = 170, x0 = h.x - W / 2, y0 = h.y + GR * 1.22 + 10; below = y0 + H;
      x.fillStyle = 'rgba(16,21,58,.88)'; x.fillRect(x0, y0, W, H); x.strokeStyle = PAL.white; x.lineWidth = 2; x.strokeRect(x0, y0, W, H);
      txt('USAGE LEFT', h.x, y0 + 14, '#ffec27', 9.5);
      L.forEach((l, i) => { const left = Math.max(0, 100 - l.used), ry = y0 + 20 + i * RH, col = left < 15 ? PAL.error : left < 40 ? PAL.quiet : PAL.working;
        x.textAlign = 'left'; txt(l.label.toUpperCase(), x0 + 10, ry + 10, PAL.white, 9.5);
        x.textAlign = 'right'; txt(Math.round(left) + '%', x0 + W - 10, ry + 10, col, 10);
        x.fillStyle = '#29466f'; x.fillRect(x0 + 10, ry + 14, W - 20, 3); x.fillStyle = col; x.fillRect(x0 + 10, ry + 14, (W - 20) * left / 100, 3);
        x.textAlign = 'left'; txt('RESETS ' + fmtReset(l.resetsAt).toUpperCase(), x0 + 10, ry + 24, '#7fd4ff', 8.5, false); });
      x.textAlign = 'center'; }
    txt(`$${(h.usd || 0).toFixed(2)} TODAY`, h.x, below + 16, PAL.white, 11);
    txt(`${h.working || 0}/${h.count} ACTIVE`, h.x, below + 30, '#7fd4ff', 10);
    for (const n of this.nodes.values()) {
      if (n.dying) continue;
      const s = n.s, R = this.nodeRadius(n) * (1 + n.swell * .12), ly = n.y + R + 22, col = PAL[s.family] || PAL.other;
      const name = (s.name.length > 22 ? s.name.slice(0, 21) + '…' : s.name).toUpperCase();
      txt(name, n.x, ly, PAL.white, 10.5);
      txt(`${(s.model || '—').replace('claude-', '').toUpperCase()} $${s.todayUsd.toFixed(2)}`, n.x, ly + 13, col, 9.5, false);
      txt((s.health === 'working' ? `▶ ${Math.round(s.tps)} TOK/S` : s.health.toUpperCase()) + (s.owner ? ' · APP' : ''), n.x, ly + 25, PAL[s.health] || PAL.idle, 9.5, false);
    }
    x.setTransform(d, 0, 0, d, 0, 0);
  },
});
