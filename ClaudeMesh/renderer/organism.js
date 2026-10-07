// Session bubbles as living micro-universes (after the user's reference: a magenta-rimmed red
// nebula, a cyan-rimmed violet one, a lime-rimmed fibrous green one with a glowing knot).
// Per session, three layers are rendered once — a swirled noise nebula, a sheet of radial
// fibres, and a soft glowing glass rim with light crescents. Each frame they're composited with
// counter-rotation and breathing, plus 3-D orbiting sparks and a pulsing heart-knot.
const ORB_PAL = {
  // sessions that use the 24/7 server (user's reference orb): thick magenta-violet glass, fiery red-orange
  // nebula with swirling flame fibres, a dense swarm of golden sparks and a small green-gold knot at the core
  server: { rim: '#e055dc', deep: [150, 16, 14], mid: [236, 62, 24], hot: [255, 196, 64], fib: '#ff6a2a', spark: '#ffe04a', knot: '#b6ff6a', sparks: 210, fibres: 320, rimWide: true, sparkBoost: 1.6, bigRatio: .82 },
  fable:  { rim: '#ff4fd8', deep: [120, 8, 18],  mid: [214, 36, 30],  hot: [255, 176, 46], fib: '#ff8a3a', spark: '#ffe9a6', knot: '#ffd36b' },
  mythos: { rim: '#ff4fd8', deep: [120, 8, 18],  mid: [214, 36, 30],  hot: [255, 176, 46], fib: '#ff8a3a', spark: '#ffe9a6', knot: '#ffd36b' },
  opus:   { rim: '#4fd0ff', deep: [62, 22, 128], mid: [140, 52, 205], hot: [255, 96, 176], fib: '#7a6cff', spark: '#ffffff', knot: '#ff7ad1' },
  sonnet: { rim: '#8dff5a', deep: [4, 38, 14],   mid: [22, 120, 40], hot: [150, 230, 80], fib: '#6dff7a', spark: '#e4ffb8', knot: '#fff6a0' },
  haiku:  { rim: '#ffb347', deep: [70, 24, 4],   mid: [168, 74, 12], hot: [255, 210, 110], fib: '#ffc46a', spark: '#fff1c4', knot: '#fff0a0' },
  other:  { rim: '#9fb4d8', deep: [18, 24, 48],  mid: [60, 76, 120], hot: [170, 190, 230], fib: '#a8bce8', spark: '#ffffff', knot: '#dfe8ff' },
};
const orbCache = new Map();

function makeOrb(fam, seed) {
  const key = fam + ':' + Math.round(seed * 1e4); if (orbCache.has(key)) return orbCache.get(key);
  const P = ORB_PAL[fam] || ORB_PAL.other, S = 256, C = S / 2, R = S / 2 - 2;
  const mk = (w = S) => { const c = document.createElement('canvas'); c.width = c.height = w; return c; };
  const n1 = vnoise(seed * .91 + .13), n2 = vnoise(seed * .53 + .71), r = rnd(seed + .5);

  // nebula: swirled fbm, dark lanes, hot heart — like a galaxy seen inside a soap bubble
  const neb = mk(), nx = neb.getContext('2d'), img = nx.createImageData(S, S);
  for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
    const dx = (x - C) / R, dy = (y - C) / R, d = Math.hypot(dx, dy); if (d > 1) continue;
    const sw = 2.4 * (1 - d), ca = Math.cos(sw), sa = Math.sin(sw), qx = dx * ca - dy * sa, qy = dx * sa + dy * ca;   // swirl
    const a = n1(qx * 3.2 + 7, qy * 3.2 + 3, 5), b = n2(qx * 7 - 2, qy * 7 + 9, 4);
    const dens = Math.pow(Math.max(0, a * 1.35 - .3), 1.4), heart = Math.pow(Math.max(0, 1 - d / .55), 2);
    const lane = b < .42 ? (.42 - b) * 2.2 : 0;                          // dark dust lanes
    const k = Math.min(1, dens * .9 + heart * .8);
    let c = [0, 1, 2].map(i => P.deep[i] + (P.mid[i] - P.deep[i]) * Math.min(1, dens * 1.4) + (P.hot[i] - P.mid[i]) * heart * (.6 + .6 * a));
    const lum = (.55 + .7 * k) * (1 - lane * .7) * (.75 + .25 * (1 - d * d));
    const o = (y * S + x) * 4;
    img.data[o] = Math.min(255, c[0] * lum); img.data[o + 1] = Math.min(255, c[1] * lum); img.data[o + 2] = Math.min(255, c[2] * lum);
    img.data[o + 3] = 255 * (d > .96 ? (1 - d) / .04 : 1);
  }
  nx.putImageData(img, 0, 0);

  // fibres: curved radial filaments (dense for the green "creature", sparse veils for the others)
  const fib = mk(), fx = fib.getContext('2d'), nF = P.fibres || (fam === 'sonnet' ? 420 : 170);
  fx.globalCompositeOperation = 'lighter'; fx.lineCap = 'round';
  for (let i = 0; i < nF; i++) {
    const a = r() * 6.283, r0 = R * (.08 + r() * .3), r1 = R * (.55 + r() * .42), bend = (r() - .5) * .7;
    const am = a + bend, rm = (r0 + r1) / 2;
    fx.strokeStyle = hexA(r() > .85 ? mix(P.fib, .5) : P.fib, .06 + r() * (fam === 'sonnet' ? .22 : .16)); fx.lineWidth = .5 + r() * 1.3;
    fx.beginPath(); fx.moveTo(C + Math.cos(a) * r0, C + Math.sin(a) * r0);
    fx.quadraticCurveTo(C + Math.cos(am) * rm, C + Math.sin(am) * rm, C + Math.cos(a + bend * 1.6) * r1, C + Math.sin(a + bend * 1.6) * r1); fx.stroke();
  }
  // a dusting of fixed embedded stars
  for (let i = 0; i < 160; i++) {
    const a = r() * 6.283, d = Math.sqrt(r()) * R * .95, z = r();
    fx.fillStyle = hexA(z > .9 ? '#ffffff' : P.spark, .25 + .6 * z); fx.fillRect(C + Math.cos(a) * d, C + Math.sin(a) * d, z > .92 ? 2 : 1.1, z > .92 ? 2 : 1.1);
  }

  // glass rim: soft glowing band (like the mother bubble's edge) + two light crescents
  const RS = 384, RC = RS / 2, RR = RS / 2 / 1.45, rim = mk(RS), rx = rim.getContext('2d');
  const g = rx.createRadialGradient(RC, RC, RR * .6, RC, RC, RR * 1.45);   // stop k ↔ radius (.6 + .85k)·R; peak at the sphere's edge
  const rc = P.rim, rw = mix(P.rim, .55);
  if (P.rimWide) {                                       // the reference's thick, soft glass band + inner pink haze
    g.addColorStop(0, hexA(rc, 0)); g.addColorStop(.18, hexA(rc, .12)); g.addColorStop(.34, hexA(rc, .42)); g.addColorStop(.43, hexA(rw, .9));
    g.addColorStop(.5, hexA(rc, .95)); g.addColorStop(.56, hexA(rc, .6)); g.addColorStop(.68, hexA(rc, .22)); g.addColorStop(1, hexA(rc, 0));
  } else {
    g.addColorStop(0, hexA(rc, 0)); g.addColorStop(.28, hexA(rc, .14)); g.addColorStop(.42, hexA(rc, .5));
    g.addColorStop(.47, hexA(rw, .95)); g.addColorStop(.52, hexA(rc, .65)); g.addColorStop(.62, hexA(rc, .2)); g.addColorStop(1, hexA(rc, 0));
  }
  rx.fillStyle = g; rx.fillRect(0, 0, RS, RS);
  rx.filter = 'blur(4px)'; rx.lineCap = 'round';
  rx.strokeStyle = 'rgba(255,255,255,.42)'; rx.lineWidth = RR * .09;
  rx.beginPath(); rx.arc(RC, RC, RR * .84, Math.PI * 1.08, Math.PI * 1.42); rx.stroke();
  rx.strokeStyle = hexA(mix(P.rim, .7), .3); rx.lineWidth = RR * .07;
  rx.beginPath(); rx.arc(RC, RC, RR * .84, Math.PI * .12, Math.PI * .38); rx.stroke();
  rx.filter = 'none';

  // per-orb animated sparks on a sphere + the heart-knot's lobes
  const sparks = Array.from({ length: P.sparks || 70 }, () => ({ th: r() * 6.283, ph: Math.acos(2 * r() - 1), rr: .3 + r() * .65, sp: (.2 + r() * .8) * (r() > .5 ? 1 : -1), tw: r() * 6.28, big: r() > (P.bigRatio || .86) }));
  const orb = { neb, fib, rim, sparks, P, lobes: 3 + (r() * 3 | 0), RK: 1.45 };
  orbCache.set(key, orb); return orb;
}

Object.assign(Mesh.prototype, {
  nodeRadius(n) { return 30 + Math.min(28, Math.log10(1 + (n.s.totalUsd || 0)) * 10); },

  drawMote(n, dt) {
    const ctx = this.ctx, s = n.s, t = this.time, fam = s.serverUses ? 'server' : ORB_PAL[s.family] ? s.family : 'other';
    const orb = makeOrb(fam, n.seed), P = orb.P;
    const alive = n.dying ? Math.max(0, 1 - (t - n.dying) / 2.5) : Math.min(1, (t - n.born) / 1.2);
    const work = s.health === 'working', energy = { working: 1, idle: .55, quiet: .7, waiting: .8, error: .9 }[s.health] ?? .5;
    n.phase *= Math.pow(.1, dt); n.flash *= Math.pow(.25, dt);
    n.spin += dt * (work ? .5 + Math.min(1.2, (s.tps || 0) / 150) : .08) * (n.seed > .5 ? 1 : -1);
    const breathe = 1 + .025 * Math.sin(t * (work ? 3.4 : 1.3) + n.seed * 9) + .012 * Math.sin(t * 7.1 + n.seed * 3);
    const R = this.nodeRadius(n) * (.5 + .5 * alive) * (1 + n.swell * .1) * breathe;
    const dim = spendDim(s), shim = .78 + .22 * Math.sin(t * 2.6 + n.seed * 17) * Math.sin(t * 1.3 + n.seed * 5);
    const rimCol = s.health === 'error' ? '#ff4d6a' : P.rim;

    // soft outer aura
    ctx.globalCompositeOperation = 'lighter';
    ctx.globalAlpha = alive * (.16 + .2 * energy + n.flash * .4) * shim * (1 - dim * .5);
    const G = R * (2.3 + n.phase * .4 + n.flash); ctx.drawImage(glowSprite(rimCol), n.x - G, n.y - G, G * 2, G * 2);
    ctx.globalCompositeOperation = 'source-over';

    // interior, clipped to the sphere
    ctx.save(); ctx.globalAlpha = alive;
    ctx.translate(n.x, n.y);
    ctx.save(); ctx.rotate(n.spin * .6); ctx.drawImage(orb.neb, -R, -R, R * 2, R * 2); ctx.restore();
    ctx.globalCompositeOperation = 'lighter';
    ctx.save(); ctx.rotate(-n.spin * .9 + Math.sin(t * .7 + n.seed) * .08);
    const fs = 1 + .03 * Math.sin(t * 2.2 + n.seed * 4); ctx.globalAlpha = alive * (.55 + .35 * energy) * (.85 + .15 * Math.sin(t * 3 + n.seed));
    ctx.drawImage(orb.fib, -R * fs, -R * fs, R * 2 * fs, R * 2 * fs); ctx.restore();
    // sparks orbiting on a sphere: brighter in front, twinkling
    const spd = work ? 1 + Math.min(2, (s.tps || 0) / 80) : .25;
    // batched: dots are bucketed by brightness and filled as one path per bucket (4 draws, not 70)
    const B = [new Path2D(), new Path2D(), new Path2D(), new Path2D()], gs = glowSprite(P.spark);
    for (const p of orb.sparks) {
      const th = p.th + n.spin * p.sp * spd * 1.6, x = Math.sin(p.ph) * Math.cos(th), z = Math.sin(p.ph) * Math.sin(th), y = Math.cos(p.ph);
      const front = .35 + .65 * (z + 1) / 2, tw = .55 + .45 * Math.sin(t * (2 + p.sp * 3) + p.tw);
      const px = x * p.rr * R, py = y * p.rr * R, a = front * tw;
      if (p.big) { const g = R * .16 * front; ctx.globalAlpha = Math.min(1, alive * a * .8 * (.45 + .55 * energy) * (P.sparkBoost || 1)); ctx.drawImage(gs, px - g, py - g, g * 2, g * 2); }
      const z2 = (p.big ? 1.05 : .62) * Math.max(1, R / 40), bk = B[Math.min(3, a * 4 | 0)]; bk.moveTo(px + z2, py); bk.arc(px, py, z2, 0, 6.283);   // round, glowing dots
    }
    ctx.fillStyle = P.spark;
    B.forEach((path, i) => { ctx.globalAlpha = Math.min(1, alive * (i + .6) / 4 * (.45 + .55 * energy) * (P.sparkBoost || 1)); ctx.fill(path); });
    // the heart: a slowly turning rose-knot that beats faster while the session works
    const beat = .8 + .2 * Math.sin(t * (work ? 6 : 2) + n.seed * 7), kr = R * .3 * beat, L = orb.lobes;
    ctx.globalAlpha = alive * (.35 + .5 * energy); ctx.strokeStyle = P.knot; ctx.lineWidth = 1.1;
    for (let j = 0; j < 2; j++) {
      ctx.beginPath();
      for (let i = 0; i <= 120; i++) {
        const u = i / 120 * 6.283, rr = kr * Math.cos(L * u + j * .6) * (1 - j * .25), w = u + t * (.4 + j * .25) * (n.seed > .5 ? 1 : -1);
        const X = Math.cos(w) * rr, Y = Math.sin(w) * rr * (.75 + .25 * Math.sin(t + j));
        i ? ctx.lineTo(X, Y) : ctx.moveTo(X, Y);
      }
      const lw = ctx.lineWidth, ga = ctx.globalAlpha;            // soft glow = a wide faint pass (cheap vs shadowBlur)
      ctx.lineWidth = 4.5; ctx.globalAlpha = ga * .22; ctx.stroke(); ctx.lineWidth = lw; ctx.globalAlpha = ga; ctx.stroke();
    }
    const hg = R * (.22 + .06 * beat); ctx.globalAlpha = alive * (.5 + .4 * energy); ctx.drawImage(glowSprite(P.knot), -hg, -hg, hg * 2, hg * 2);
    ctx.globalCompositeOperation = 'source-over';
    if (dim > .01) { ctx.globalAlpha = alive * dim * .75; ctx.fillStyle = '#05030f'; ctx.beginPath(); ctx.arc(0, 0, R, 0, 6.29); ctx.fill(); }   // spend darkens it
    ctx.restore();

    // glowing glass rim (soft, like the mother bubble's edge) — shimmering
    ctx.globalCompositeOperation = 'lighter';
    const RK = R * orb.RK * (1 + .01 * Math.sin(t * 5 + n.seed));
    ctx.globalAlpha = alive * (.7 + .3 * shim) * (1 - dim * .45) * (s.health === 'error' && Math.random() > .5 ? .5 : 1);
    ctx.drawImage(orb.rim, n.x - RK, n.y - RK, RK * 2, RK * 2);
    if (n.flash > .02) { ctx.globalAlpha = n.flash; ctx.drawImage(orb.rim, n.x - RK, n.y - RK, RK * 2, RK * 2); }
    ctx.globalCompositeOperation = 'source-over'; ctx.globalAlpha = 1;
    const col = sessionColor(s);

    // attention pulse
    if (s.health === 'waiting' || s.health === 'quiet') {
      const k = (t * .8) % 1; ctx.strokeStyle = `rgba(255,207,107,${(1 - k) * .7})`; ctx.lineWidth = 1.5;
      ctx.beginPath(); ctx.arc(n.x, n.y, R + 6 + k * 26, 0, 6.29); ctx.stroke();
    }
    // context gauge (thin outer arc)
    const ctxPct = s.ctxMax ? Math.min(1, s.ctx / s.ctxMax) : 0;
    ctx.globalAlpha = alive;
    ctx.strokeStyle = 'rgba(255,255,255,.06)'; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.arc(n.x, n.y, R + 14, 0, 6.29); ctx.stroke();
    ctx.strokeStyle = ctxPct > .85 ? '#ff6b7d' : ctxPct > .6 ? '#ffcf6b' : hexA(mix(P.rim, .3), .8);
    ctx.beginPath(); ctx.arc(n.x, n.y, R + 14, -Math.PI / 2, -Math.PI / 2 + ctxPct * 6.283); ctx.stroke();
    // (sub-agent satellite orbs removed at the user's request — the count stays in the session panel)
    if (this.selected === n.sid) {
      ctx.strokeStyle = 'rgba(255,255,255,.8)'; ctx.lineWidth = 1; ctx.setLineDash([2, 5]);
      ctx.beginPath(); ctx.arc(n.x, n.y, R + 17, t * .6, t * .6 + 6.283); ctx.stroke(); ctx.setLineDash([]);
    }
    // label pill
    ctx.globalAlpha = alive * (.85 + .15 * n.swell);
    const name = s.name.length > 30 ? s.name.slice(0, 29) + '…' : s.name;
    const l2 = `${(s.model || '—').replace('claude-', '')}  ·  $${s.todayUsd.toFixed(2)}`;
    const l3 = (work ? `working · ${Math.round(s.tps)} tok/s` : s.health) + (s.owner ? '  ·  in-app' : '') + (s.serverUses ? '  ·  ⇄ server' : '');
    ctx.font = '600 12px system-ui, "Apple SD Gothic Neo"'; const w1 = ctx.measureText(name).width;
    ctx.font = '400 10.5px system-ui'; const pw = Math.max(w1, ctx.measureText(l2).width, ctx.measureText(l3).width) + 22;
    const py = n.y + R + 22, ph = 46; n.labelW = pw;
    ctx.fillStyle = 'rgba(12,12,40,.72)'; ctx.strokeStyle = hexA(mix(P.rim, .3), .3 + .25 * n.swell); ctx.lineWidth = 1;
    ctx.beginPath(); ctx.roundRect(n.x - pw / 2, py, pw, ph, 8); ctx.fill(); ctx.stroke();
    ctx.textAlign = 'center';
    ctx.font = '600 12px system-ui, "Apple SD Gothic Neo"'; ctx.fillStyle = '#f2f3ff'; ctx.fillText(name, n.x, py + 15);
    ctx.font = '400 10.5px system-ui'; ctx.fillStyle = hexA(mix(col, .4), .95); ctx.fillText(l2, n.x, py + 29);
    ctx.fillStyle = HEALTH_COLOR[s.health] || '#9fb4d8'; ctx.fillText(l3, n.x, py + 41);
    ctx.globalAlpha = 1;
  },
});
