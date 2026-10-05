// Session bubbles + the bubbles they produce.
//  • Every session slowly orbits the mother bubble. Drag one anywhere: the drop point becomes
//    its orbit (saved per session). Farther orbits turn slower, closer ones faster (Kepler-ish).
//    Double-click a session bubble to send it back to its automatic slot.
//  • Sessions blow bubbles in their model's colour — more and bigger the more tokens they burn.
//  • The mother bubble's gravity pulls every bubble in (stronger the closer it gets) and eats it.
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
  orbitPos(o) {                           // o = { rx, a0, t0 } — angle advances with wall-clock time
    const sh = this.orbitShape(), a = o.a0 + this.omegaAt(o.rx) * (Date.now() - o.t0) / 1000;
    return { x: this.hub.x + Math.cos(a) * o.rx, y: this.hub.y + Math.sin(a) * o.rx * sh.k };
  },
  dropOrbit(n) {                          // turn the drop point into an orbit on the same ellipse family
    const sh = this.orbitShape(), dx = n.x - this.hub.x, dy = (n.y - this.hub.y) / sh.k;
    const rx = Math.max(this.motherRadius() + this.nodeRadius(n) + 20, Math.hypot(dx, dy));
    this.loadOrbits()[n.sid] = { rx, a0: Math.atan2(dy, dx), t0: Date.now() }; this.saveOrbits();
  },

  // ---------------- session motion (shared by both renderers) ----------------
  stepSessions(dt, realDt, cspd) {
    const cur = this.cursor, orb = this.loadOrbits();
    const live = [...this.nodes.values()].filter(n => !n.dying);
    const auto = live.filter(n => !orb[n.sid]);
    live.forEach(n => {
      n.swell += ((this.hover === n.sid || n.dragging ? 1 : 0) - n.swell) * Math.min(1, realDt * 8);
      if (n.dragging) { n.vx = n.vy = 0; return; }
      const h = orb[n.sid] ? this.orbitPos(orb[n.sid]) : this.home(n, auto, auto.indexOf(n));
      n.vx += (h.x - n.x) * 2.2 * dt; n.vy += (h.y - n.y) * 2.2 * dt;
      if (cur.in && !orb[n.sid]) {
        const dx = n.x - cur.x, dy = n.y - cur.y, d = Math.hypot(dx, dy), R = this.nodeRadius(n) + 90;
        if (d < R && d > 1 && this.hover !== n.sid) { const f = (1 - d / R) * Math.max(0, cspd - 180) * .9; n.vx += dx / d * f * dt; n.vy += dy / d * f * dt; }
      }
      if (!orb[n.sid]) for (const o of live) if (o !== n) {
        const dx = n.x - o.x, dy = n.y - o.y, d = Math.hypot(dx, dy), min = this.nodeRadius(n) + this.nodeRadius(o) + 50;
        if (d < min && d > .1) { n.vx += dx / d * (min - d) * 3 * dt; n.vy += dy / d * (min - d) * 3 * dt; }
      }
      const damp = Math.pow(.12, dt); n.vx *= damp; n.vy *= damp;
      n.x += n.vx * dt; n.y += n.vy * dt;
    });
    for (const [sid, n] of this.nodes) if (n.dying && this.time - n.dying > 2.5) this.nodes.delete(sid);
  },

  // ---------------- bubbles: emitted by sessions, eaten by the mother ----------------
  emitBubble(n) {
    // leave mostly from the side facing the mother: a cone around that direction (bell-shaped spread), rare strays wider
    const toHub = Math.atan2(this.hub.y - n.y, this.hub.x - n.x), spread = (Math.random() + Math.random() + Math.random() - 1.5) * (Math.random() < .12 ? 2.2 : .75);
    const s = n.s, R = this.nodeRadius(n), a = toHub + spread, sp = 50 + Math.random() * 80;
    const tang = (n.seed > .5 ? 1 : -1) * (12 + Math.random() * 28);
    const q = Math.random(), r = .7 + q * q * q * 4.3;                       // mostly tiny, the odd larger one
    n.flash = Math.min(.35, n.flash + .02);
    return { x: n.x + Math.cos(a) * (R + 2), y: n.y + Math.sin(a) * (R + 2), vx: Math.cos(a) * sp - Math.sin(toHub) * tang, vy: Math.sin(a) * sp + Math.cos(toHub) * tang,
      r, fam: s.family, color: FAMILY_COLOR[s.family] || FAMILY_COLOR.other, born: this.time, ph: Math.random() * 6.28, depth: 1 };
  },
  seedBubble() {
    const v = this.view(), fams = ['opus', 'sonnet', 'fable', 'haiku'], fam = fams[Math.random() * fams.length | 0];
    return { x: v.x0 + Math.random() * (v.x1 - v.x0), y: v.y0 + Math.random() * (v.y1 - v.y0), vx: (Math.random() - .5) * 20, vy: (Math.random() - .5) * 20,
      r: 2 + Math.random() * 4, fam, color: FAMILY_COLOR[fam], born: this.time, ph: Math.random() * 6.28, depth: 1 };
  },

  simAmbient(dt, cspd) {
    if (!this.seeded && this.w) { this.seeded = true; for (let i = 0; i < 18; i++) this.ambient.push(this.seedBubble()); }
    const A = this.ambient, hub = this.hub, cur = this.cursor, v = this.view();
    const MR = this.motherRadius(), cap = this.pixel ? 140 : 520, live = [...this.nodes.values()].filter(n => !n.dying);
    // emission: rate and size follow each session's token speed
    for (const n of live) {
      const s = n.s, rate = s.tps > 0 ? Math.min(40, 2 + s.tps / 4) : s.health === 'working' ? 1.5 : .04;
      n.bacc = (n.bacc || 0) + rate * dt;
      while (n.bacc >= 1) { n.bacc -= 1; if (A.length < cap) A.push(this.emitBubble(n)); n.phase = 1; }
    }
    for (const m of A) {
      // mother gravity: ∝ 1/d², so it tugs gently from afar and hard up close
      const dx = hub.x - m.x, dy = hub.y - m.y, dist = Math.hypot(dx, dy) || 1, d2 = dist * dist + 900;
      const g = 4.2e6 / d2; m.vx += dx / dist * g * dt; m.vy += dy / dist * g * dt;     // a = GM / r² (softened)
      if (dist < MR * .78) {
        m.dead = true; hub.mass = Math.min(14, (hub.mass || 0) + m.r * .18);
        if (!this.pixel) this.fx.push({ kind: 'absorb', x: m.x, y: m.y, r: m.r, color: m.color, t: 0, to: hub });
        continue;
      }
      if (cur.in) {
        const cx = m.x - cur.x, cy = m.y - cur.y, cd = Math.hypot(cx, cy), R = 70 + m.r;
        if (cd < R && cd > 1) { const f = (1 - cd / R) * (40 + cspd * .4); m.vx += cx / cd * f * dt; m.vy += cy / cd * f * dt; }
      }
      for (const n of live) {                // slide off session bubbles instead of passing through
        const R = this.nodeRadius(n) + m.r + 2, sx = m.x - n.x, sy = m.y - n.y, sd = Math.hypot(sx, sy);
        if (sd < R && sd > .1 && this.time - m.born > .3) { m.vx += sx / sd * (R - sd) * 8 * dt; m.vy += sy / sd * (R - sd) * 8 * dt; }
      }
      const damp = Math.pow(.93, dt); m.vx *= damp; m.vy *= damp;       // a little drag → orbits decay into the mother
      m.x += m.vx * dt; m.y += m.vy * dt;
      if (m.x < v.x0 - 500 || m.x > v.x1 + 500 || m.y < v.y0 - 500 || m.y > v.y1 + 500) m.dead = true;
    }
    this.ambient = A.filter(m => !m.dead);
  },

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
  drawLinks(ctx) {
    const hub = this.hub, t = this.time, MR = this.motherRadius();
    ctx.globalCompositeOperation = 'lighter'; ctx.lineCap = 'round';
    for (const n of this.nodes.values()) {
      if (n.dying) continue;
      const col = sessionColor(n.s), R = this.nodeRadius(n);
      const dx = hub.x - n.x, dy = hub.y - n.y, L = Math.hypot(dx, dy) || 1, ux = dx / L, uy = dy / L;
      const x0 = n.x + ux * R, y0 = n.y + uy * R, x1 = hub.x - ux * MR * 1.05, y1 = hub.y - uy * MR * 1.05;
      const bend = (n.seed > .5 ? 1 : -1) * L * .22, mx = (x0 + x1) / 2 - uy * bend, my = (y0 + y1) / 2 + ux * bend;
      const work = n.s.health === 'working' ? 1 : 0, g = ctx.createLinearGradient(x0, y0, x1, y1);
      g.addColorStop(0, hexA(col, .9)); g.addColorStop(1, 'rgba(255,150,90,.8)');
      for (const [w, a] of [[7, .07 + .05 * work], [3, .16 + .1 * work], [1.2, .55 + .3 * work]]) {
        ctx.strokeStyle = g; ctx.globalAlpha = a; ctx.lineWidth = w;
        ctx.beginPath(); ctx.moveTo(x0, y0); ctx.quadraticCurveTo(mx, my, x1, y1); ctx.stroke();
      }
      ctx.globalAlpha = 1;
      const pulses = work ? 1 + Math.min(3, (n.s.tps || 0) / 60 | 0) : 0;
      for (let i = 0; i < pulses; i++) {
        const k = ((t * (.35 + (n.s.tps || 0) / 400) + i / pulses + n.seed) % 1), q = 1 - k;
        const px = q * q * x0 + 2 * q * k * mx + k * k * x1, py = q * q * y0 + 2 * q * k * my + k * k * y1, sz = 9;
        ctx.drawImage(glowSprite(mix(col, .4)), px - sz, py - sz, sz * 2, sz * 2);
      }
    }
    ctx.globalCompositeOperation = 'source-over';
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
