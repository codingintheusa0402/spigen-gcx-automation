// Session bubbles + the bubbles they produce.
//  • Every session slowly orbits the mother bubble. Drag one anywhere: the drop point becomes
//    its orbit (saved per session). Farther orbits turn slower, closer ones faster (Kepler-ish).
//    Double-click a session bubble to send it back to its automatic slot.
//  • Each session is tied to the mother by glowing strands: more strands, and an older star
//    colour, the harder it works (drawLinks).
const ORBIT_LAP = 900;
// load colour 0…1: light blue → yellow → orange → red
const LOAD_RAMP = [[0, [150, 210, 255]], [.4, [255, 236, 110]], [.7, [255, 152, 44]], [1, [236, 52, 40]]];
function loadColor(x) {
  for (let i = 1; i < LOAD_RAMP.length; i++) if (x <= LOAD_RAMP[i][0]) {
    const [a, ca] = LOAD_RAMP[i - 1], [b, cb] = LOAD_RAMP[i], k = (x - a) / (b - a);
    return ca.map((v, j) => Math.round(v + (cb[j] - v) * k));
  }
  return LOAD_RAMP[LOAD_RAMP.length - 1][1];
}                    // seconds per lap for an automatic slot (≈ 15 min)
const PIN_KEY = 'mesh.orbits';

function bubbleSprite(color) {
  const key = 'b' + color; if (spriteCache.has(key)) return spriteCache.get(key);
  const S = 96, c = document.createElement('canvas'); c.width = c.height = S;
  const x = c.getContext('2d'), C = S / 2, R = S * .31;         // body radius; the rest is glow
  const glow = x.createRadialGradient(C, C, R * .7, C, C, S / 2);
  glow.addColorStop(0, hexA(color, 0)); glow.addColorStop(.28, hexA(color, .55)); glow.addColorStop(.42, hexA(color, .22)); glow.addColorStop(1, hexA(color, 0));
  x.fillStyle = glow; x.fillRect(0, 0, S, S);
  const body = x.createRadialGradient(C, C, 0, C, C, R);
  body.addColorStop(0, hexA(color, .08)); body.addColorStop(.7, hexA(color, .16)); body.addColorStop(.9, hexA(mix(color, .55), .85)); body.addColorStop(1, hexA(mix(color, .8), 1));
  x.fillStyle = body; x.beginPath(); x.arc(C, C, R, 0, 6.29); x.fill();
  const spec = x.createRadialGradient(C - R * .4, C - R * .45, 0, C - R * .4, C - R * .45, R * .5);
  spec.addColorStop(0, 'rgba(255,255,255,.55)'); spec.addColorStop(1, 'rgba(255,255,255,0)');
  x.fillStyle = spec; x.beginPath(); x.arc(C, C, R, 0, 6.29); x.fill();
  spriteCache.set(key, c); return c;
}
const BUBBLE_SPRITE_K = 1 / .31;          // sprite size / body radius

Object.assign(Mesh.prototype, {
  // ---------------- orbits ----------------
  loadOrbits() { if (!this.orbits) { try { this.orbits = JSON.parse(localStorage.getItem(PIN_KEY)) || {}; } catch { this.orbits = {}; } } return this.orbits; },
  saveOrbits() { try { localStorage.setItem(PIN_KEY, JSON.stringify(this.orbits)); } catch { } },
  orbitShape() {                          // the automatic orbit ellipse (also the speed reference)
    const rx = Math.max(this.motherRadius() + 130, this.w / 2 - 130), ry = Math.max(this.motherRadius() + 100, this.h / 2 - 120);
    return { rx, ry, k: ry / rx };
  },
  // keep a session's whole footprint (glow ring above, label pill + hint line below) inside the mesh
  // window at any window/dock size — orbits that would leave it slide along the edge instead
  fitInView(n, p) {
    const R = this.nodeRadius(n) * 1.12, half = Math.max(R + 20, (n.labelW || 200) / 2) + 8;
    const top = R + 22, bottom = R + 22 + 46 + 10 + 26;       // ring | label gap + pill + pad + #meshHint
    const x0 = half, x1 = this.w - half, y0 = top, y1 = this.h - bottom;
    const cx = v => x1 > x0 ? Math.min(x1, Math.max(x0, v)) : this.w / 2;
    let x = cx(p.x); const y = y1 > y0 ? Math.min(y1, Math.max(y0, p.y)) : (y0 + y1) / 2;
    // squeezed into the mother's column (short window)? slide sideways so neither covers the other
    const m = this.motherBox(), nTop = y - top, nBot = y + R + 22 + 46;
    if (nBot > m.y0 && nTop < m.y1 && x + half > m.x0 && x - half < m.x1)
      x = cx(x < this.hub.x ? m.x0 - half - 6 : m.x1 + half + 6);
    return { x, y };
  },
  // screen box of the mother bubble + its usage card + "$ today" line (see drawMotherInfo)
  motherBox() {
    const R = this.motherRadius(), L = (this.hub.usage && this.hub.usage.limits) || [];
    const cardH = 22 + Math.max(1, L.length) * 24, hw = Math.max(R * 1.22, 107) + 4;
    return { x0: this.hub.x - hw, x1: this.hub.x + hw, y0: this.hub.y - R * 1.22, y1: this.hub.y + R * 1.22 + 8 + cardH + 24 };
  },
  // mother sits at the centre unless that would push its usage card below the mesh (short window/tall dock)
  motherY() {
    const R = this.motherRadius(), below = R * 1.22 + 8 + (22 + 3 * 24) + 24 + 30;   // card + today line + #meshHint
    return Math.max(R * 1.22 + 6, Math.min(this.h / 2, this.h - below));
  },
  omegaAt(rx) {                           // rad/s; reference speed at the automatic orbit, ∝ r^-1.5
    const ref = this.orbitShape().rx, w0 = 6.283 / ORBIT_LAP;
    return w0 * Math.min(4, Math.max(.2, Math.pow(ref / Math.max(20, rx), 1.5)));
  },
  orbitPos(o) {                           // o = { rx, a0, t0, auto? } — angle advances with wall-clock time
    const sh = this.orbitShape(), rx = o.auto ? sh.rx : o.rx, a = o.a0 + this.omegaAt(rx) * (Date.now() - o.t0) / 1000;
    return { x: this.hub.x + Math.cos(a) * rx, y: this.hub.y + Math.sin(a) * rx * sh.k };
  },
  // a new (or released) session gets its own automatic orbit in the widest free gap — once.
  // Sessions never get re-slotted when others arrive, leave or are dragged.
  autoOrbit(n, live) {
    const sh = this.orbitShape(), orb = this.loadOrbits(), now = Date.now();
    const angs = live.filter(o => o !== n && orb[o.sid]).map(o => { const p = this.orbitPos(orb[o.sid]); return Math.atan2((p.y - this.hub.y) / sh.k, p.x - this.hub.x); });
    let best = -Math.PI / 2, bestGap = -1;
    for (let i = 0; i < 72; i++) {
      const a = -Math.PI / 2 + i / 72 * 6.283;
      const gap = angs.length ? Math.min(...angs.map(b => Math.abs(Math.atan2(Math.sin(a - b), Math.cos(a - b))))) : 9;
      if (gap > bestGap + 1e-6) { bestGap = gap; best = a; }
    }
    orb[n.sid] = { auto: true, rx: sh.rx, a0: best, t0: now }; this.saveOrbits();
  },
  dropOrbit(n) {                          // turn the drop point into an orbit on the same ellipse family
    const sh = this.orbitShape(), dx = n.x - this.hub.x, dy = (n.y - this.hub.y) / sh.k;
    const rx = Math.max(this.motherRadius() + this.nodeRadius(n) + 20, Math.hypot(dx, dy));
    this.loadOrbits()[n.sid] = { rx, a0: Math.atan2(dy, dx), t0: Date.now() }; this.saveOrbits();
  },

  // ---------------- session motion (shared by both renderers) ----------------
  stepSessions(dt, realDt, cspd) {
    const orb = this.loadOrbits();
    const live = [...this.nodes.values()].filter(n => !n.dying);
    for (const n of live) if (!orb[n.sid]) this.autoOrbit(n, live);
    live.forEach(n => {
      n.swell += ((this.hover === n.sid || n.dragging ? 1 : 0) - n.swell) * Math.min(1, realDt * 8);
      if (n.dragging) { n.vx = n.vy = 0; return; }
      const h = this.fitInView(n, this.orbitPos(orb[n.sid]));   // each session follows only its own orbit, kept on-screen
      n.vx += (h.x - n.x) * 2.2 * dt; n.vy += (h.y - n.y) * 2.2 * dt;
      const damp = Math.pow(.12, dt); n.vx *= damp; n.vy *= damp;
      n.x += n.vx * dt; n.y += n.vy * dt;
    });
    for (const [sid, n] of this.nodes) if (n.dying && this.time - n.dying > 2.5) this.nodes.delete(sid);
  },

    // token bubbles were removed (too costly); load now shows as strands in drawLinks
  simAmbient() { this.ambient.length = 0; },
  // shimmering, glow-rimmed bubbles
  drawBubbles(ctx, parX, parY) {
    const t = this.time;
    for (const m of this.ambient) {
      const fade = Math.min(1, (t - m.born) / .6), sh = .72 + .28 * Math.sin(t * 4.2 + m.ph) * Math.sin(t * 2.3 + m.ph * 2);
      const r = m.r * (1 + .05 * Math.sin(t * 6 + m.ph)), S = r * BUBBLE_SPRITE_K;
      ctx.globalAlpha = fade * sh;
      ctx.drawImage(bubbleSprite(m.color), m.x - S, m.y - S, S * 2, S * 2);
    }
    ctx.globalAlpha = 1;
  },

  // smoothed load 0…1 from tok/s (√ so moderate work already shows a visible bundle)
  heaviness(n) {
    const target = n.s.health === 'working' ? Math.min(1, Math.sqrt((n.s.tps || 0) / 300)) : 0;
    n.heavy = (n.heavy || 0) + (target - (n.heavy || 0)) * .04;      // eases over ~1–2 s
    // wiggle phase advances faster the harder the session works (accumulated, so speed changes stay smooth)
    const now = this.time; n.wig = (n.wig || 0) + Math.max(0, now - (n._wt ?? now)) * (.35 + 5.5 * n.heavy); n._wt = now;
    return n.heavy;
  },
  // Strand j has a fixed place in the bundle (golden-ratio spread), so new strands fan in
  // around the existing ones instead of everything shifting as the count grows.
  strandGeom(n, j, t) {
    const hub = this.hub, MR = this.motherRadius(), R = this.nodeRadius(n);
    const dx = hub.x - n.x, dy = hub.y - n.y, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
    const side = n.seed > .5 ? 1 : -1, u = j === 0 ? 0 : ((j * .6180339887 + n.seed) % 1) * 2 - 1;   // −1…1, 0 = centre line
    const rot = (x, y, a) => [x * Math.cos(a) - y * Math.sin(a), x * Math.sin(a) + y * Math.cos(a)];
    const [sx, sy] = rot(ux, uy, u * .26), [ex, ey] = rot(-ux, -uy, -u * .13);       // a narrow bundle around the centre line
    const x0 = n.x + sx * R, y0 = n.y + sy * R, x1 = hub.x + ex * MR * 1.03, y1 = hub.y + ey * MR * 1.03;
    const bend = (side * .2 + u * .04 + .01 * Math.sin(t * (.9 + (j % 7) * .23) + n.seed * 9 + j)) * L;
    // most strands follow a smooth path to the star, each with its own slight random bow (a gentle
    // outward sag that grows with distance from the centre line); only ~1 in 8 squiggles
    const far = Math.abs(u), h = ((j * 7919 + Math.round(n.seed * 1e4)) % 97) / 97, squig = j > 0 && h < .125;
    const amp = squig ? L * (.005 + .012 * Math.pow(far, 1.4)) : 0;
    const sag = j === 0 ? 0 : (Math.sign(u || 1) * .022 * far * far + (h - .5) * .016) * L;
    return { x0, y0, x1, y1, mx: (x0 + x1) / 2 - uy * bend, my: (y0 + y1) / 2 + ux * bend, nx: -uy, ny: ux,
      amp, sag, squig, freq: 1.4 + h * 8, ph: h * 6.283 + (n.wig || 0) * (1 + h * 3) * (h > .06 ? 1 : -1), wobbly: j > 0 };
  },
  // point at s∈[0,1] along a strand: the quadratic curve plus its dangle/wiggle (zero at both ends)
  strandPoint(g, s) {
    const q = 1 - s, x = q * q * g.x0 + 2 * q * s * g.mx + s * s * g.x1, y = q * q * g.y0 + 2 * q * s * g.my + s * s * g.y1;
    if (!g.wobbly) return [x, y];
    const env = Math.sin(Math.PI * s), off = env * (g.sag + g.amp * Math.sin(g.freq * s * 6.283 + g.ph));
    return [x + g.nx * off, y + g.ny * off];
  },

  // Load lines: 1 strand when idle, up to 50 under heavy load, ageing young → old star colour.
  // Strand 0 is the bright main link; the rest are batched into one path per session (one stroke).
  drawLinks(ctx) {
    const t = this.time, cap = 20;
    ctx.globalCompositeOperation = 'lighter'; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    for (const n of this.nodes.values()) {
      if (n.dying) continue;
      const h = this.heaviness(n), N = Math.max(1, Math.round(1 + h * (cap - 1))), work = n.s.health === 'working' ? 1 : 0;
      const rim = (ORB_PAL[n.s.serverUses ? 'server' : n.s.family] || ORB_PAL.other).rim;
      const G = [];
      for (let j = 0; j < N; j++) {
        const g = this.strandGeom(n, j, t), jit = ((j * 9301 + Math.round(n.seed * 1e4)) % 233) / 233 - .5;   // per-strand variety
        const c = loadColor(Math.min(1, Math.max(0, h + jit * .28))), w = (j === 0 ? 1.6 : 1) * (.7 + h * 2.6) * (1 + jit * .5);
        const vis = Math.min(1, N - j);                                // the newest strand fades in
        const steps = g.squig ? Math.ceil(g.freq * 12) : 16, path = new Path2D(); path.moveTo(g.x0, g.y0);
        for (let i = 1; i <= steps; i++) { const [x, y] = this.strandPoint(g, i / steps); path.lineTo(x, y); }
        if (j === 0) {                                                 // main link: from the orb's rim colour into the load colour
          const grad = ctx.createLinearGradient(g.x0, g.y0, g.x1, g.y1); grad.addColorStop(0, hexA(rim, .9)); grad.addColorStop(.3, `rgba(${c},.95)`); grad.addColorStop(1, `rgba(${c},1)`);
          ctx.strokeStyle = grad;
        } else ctx.strokeStyle = `rgb(${c})`;
        ctx.globalAlpha = (.08 + .07 * work) * vis; ctx.lineWidth = w * 3.2; ctx.stroke(path);        // soft glow
        ctx.globalCompositeOperation = 'source-over';
        ctx.globalAlpha = (j === 0 ? .9 : .6 + .25 * h) * vis; ctx.lineWidth = w; ctx.stroke(path);    // true-colour core
        ctx.globalCompositeOperation = 'lighter';
        G.push({ g, c });
      }
      if (work) {                                                      // light pulses run down every 3rd strand
        G.filter((_, i) => i % 3 === 0).slice(0, 8).forEach(({ g, c }, i) => {
          const kk = (t * (.3 + h * .9) + i * .37 + n.seed) % 1, sz = 5 + 4 * h, [px, py] = this.strandPoint(g, kk);
          ctx.globalAlpha = .55 + .4 * h; ctx.drawImage(glowSprite('#' + c.map(v => v.toString(16).padStart(2, '0')).join('')), px - sz, py - sz, sz * 2, sz * 2);
        });
      }
    }
    ctx.globalAlpha = 1; ctx.globalCompositeOperation = 'source-over';
  },

  // ---------------- drag & drop ----------------
  dragStart(e, sid) {
    this.drag = { sid, x0: e.clientX, y0: e.clientY, moved: false, held: false };
    clearTimeout(this.holdTimer);                                  // press-and-hold (no movement) → details dropdown
    this.holdTimer = setTimeout(() => { const d = this.drag; if (d && !d.moved && this.onHold) { d.held = true; this.onHold(d.sid, d.x0, d.y0); } }, 520);
  },
  dragMove(e) {
    const d = this.drag; if (!d) return false;
    if (!d.moved && !d.held && Math.hypot(e.clientX - d.x0, e.clientY - d.y0) > 4) { d.moved = true; clearTimeout(this.holdTimer); }
    const n = this.nodes.get(d.sid);
    if (d.moved && n) { const p = this.toWorld(e); n.x = p.x; n.y = p.y; n.dragging = true; this.c.style.cursor = 'grabbing'; }
    return true;
  },
  dragEnd() {
    const d = this.drag; if (!d) return; this.drag = null; this.c.style.cursor = ''; clearTimeout(this.holdTimer);
    if (d.held) return;                                            // the hold already opened the details
    const n = this.nodes.get(d.sid);
    if (!d.moved) return this.onSelect(d.sid);
    if (n) { n.dragging = false; this.dropOrbit(n); }
  },
  unpin(sid) { const o = this.loadOrbits(); if (o[sid]) { delete o[sid]; this.saveOrbits(); } },
});
