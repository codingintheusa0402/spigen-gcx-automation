// Session bubbles + the bubbles they produce.
//  • Every session slowly orbits the mother bubble. Drag one anywhere: the drop point becomes
//    its orbit (saved per session). Farther orbits turn slower, closer ones faster (Kepler-ish).
//    Double-click a session bubble to send it back to its automatic slot.
//  • Each session is tied to the mother by glowing strands: more strands, and an older star
//    colour, the harder it works (drawLinks).
const ORBIT_LAP = 900;                    // seconds per lap for an automatic slot (≈ 15 min)
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
    const rx = Math.max(this.motherRadius() + 130, this.w / 2 - 130), ry = Math.max(this.motherRadius() + 100, this.h / 2 - 90);
    return { rx, ry, k: ry / rx };
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
      const h = this.orbitPos(orb[n.sid]);             // each session follows only its own orbit
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

  // glowing curved links from each session to the mother; a light pulse travels while it works
  // Load lines: every session is tied to the mother by glowing strands. The harder it works
  // (tok/s, smoothed), the more strands (1 → 7) and the older their star colour — young
  // blue-white when light, cooling through gold to red-orange under heavy load.
  heaviness(n) {
    const target = n.s.health === 'working' ? Math.min(1, Math.sqrt((n.s.tps || 0) / 260)) : 0;
    n.heavy = (n.heavy || 0) + (target - (n.heavy || 0)) * .04;      // eases over ~1–2 s
    return n.heavy;
  },
  strandGeom(n, j, t) {
    const hub = this.hub, MR = this.motherRadius(), R = this.nodeRadius(n);
    const dx = hub.x - n.x, dy = hub.y - n.y, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
    const side = n.seed > .5 ? 1 : -1, fan = j === 0 ? 0 : (j % 2 ? 1 : -1) * Math.ceil(j / 2);   // 0, +1, -1, +2, -2 …
    const sa = fan * .16, sx = ux * Math.cos(sa) - uy * Math.sin(sa), sy = ux * Math.sin(sa) + uy * Math.cos(sa);
    const x0 = n.x + sx * R, y0 = n.y + sy * R, x1 = hub.x - ux * MR * 1.04, y1 = hub.y - uy * MR * 1.04;
    const bend = (side * .22 + fan * .09 + .025 * Math.sin(t * (1.1 + j * .37) + n.seed * 9 + j)) * L;
    return { x0, y0, x1, y1, mx: (x0 + x1) / 2 - uy * bend, my: (y0 + y1) / 2 + ux * bend };
  },
  drawLinks(ctx) {
    const t = this.time;
    ctx.globalCompositeOperation = 'lighter'; ctx.lineCap = 'round';
    for (const n of this.nodes.values()) {
      if (n.dying) continue;
      const h = this.heaviness(n), strands = 1 + h * 6, work = n.s.health === 'working' ? 1 : 0;
      const star = starColor('outer', .08 + h * .92, 1.7).map(Math.round), starRGB = `${star}`;
      const rim = (ORB_PAL[n.s.family] || ORB_PAL.other).rim;
      for (let j = 0; j < Math.ceil(strands); j++) {
        const vis = Math.min(1, strands - j), g = this.strandGeom(n, j, t);   // the newest strand fades in
        const grad = ctx.createLinearGradient(g.x0, g.y0, g.x1, g.y1);
        grad.addColorStop(0, hexA(rim, .85)); grad.addColorStop(.35, `rgba(${starRGB},.9)`); grad.addColorStop(1, `rgba(${starRGB},.95)`);
        ctx.strokeStyle = grad;
        const passes = j === 0 ? [[7, .07 + .06 * work], [3, .16 + .12 * work], [1.2, .55 + .3 * work]] : [[4, .1 + .08 * h], [1, .45 + .3 * h]];
        for (const [w, a] of passes) {
          ctx.globalAlpha = a * vis; ctx.lineWidth = w;
          ctx.beginPath(); ctx.moveTo(g.x0, g.y0); ctx.quadraticCurveTo(g.mx, g.my, g.x1, g.y1); ctx.stroke();
        }
        if (work) {                                                    // a light pulse runs down each strand
          const k = (t * (.3 + h * .9) + j * .37 + n.seed) % 1, q = 1 - k, sz = 6 + 4 * h;
          const px = q * q * g.x0 + 2 * q * k * g.mx + k * k * g.x1, py = q * q * g.y0 + 2 * q * k * g.my + k * k * g.y1;
          ctx.globalAlpha = vis * (.6 + .4 * h); ctx.drawImage(glowSprite('#' + star.map(v => v.toString(16).padStart(2, '0')).join('')), px - sz, py - sz, sz * 2, sz * 2);
        }
      }
    }
    ctx.globalAlpha = 1; ctx.globalCompositeOperation = 'source-over';
  },


  // ---------------- drag & drop ----------------
  dragStart(e, sid) { this.drag = { sid, x0: e.clientX, y0: e.clientY, moved: false }; },
  dragMove(e) {
    const d = this.drag; if (!d) return false;
    if (!d.moved && Math.hypot(e.clientX - d.x0, e.clientY - d.y0) > 4) d.moved = true;
    const n = this.nodes.get(d.sid);
    if (d.moved && n) { const p = this.toWorld(e); n.x = p.x; n.y = p.y; n.dragging = true; this.c.style.cursor = 'grabbing'; }
    return true;
  },
  dragEnd() {
    const d = this.drag; if (!d) return; this.drag = null; this.c.style.cursor = '';
    const n = this.nodes.get(d.sid);
    if (!d.moved) return this.onSelect(d.sid);
    if (n) { n.dragging = false; this.dropOrbit(n); }
  },
  unpin(sid) { const o = this.loadOrbits(); if (o[sid]) { delete o[sid]; this.saveOrbits(); } },
});
