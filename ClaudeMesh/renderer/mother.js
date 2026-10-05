// The mother bubble at the centre: a burning star whose layers age with usage (see starColor) generated once as a mottled texture, then drawn in thin horizontal strips
// with drifting offsets every frame so it ripples like heat haze. Its gravity swallows every
// bubble the sessions produce (bubbles.js); the plan-usage card sits just below it.
function vnoise(seed) {
  const P = new Uint8Array(512), r = rnd(seed);
  for (let i = 0; i < 256; i++) P[i] = i;
  for (let i = 255; i > 0; i--) { const j = r() * (i + 1) | 0; [P[i], P[j]] = [P[j], P[i]]; }
  for (let i = 0; i < 256; i++) P[i + 256] = P[i];
  const G = Float32Array.from({ length: 256 }, () => r());
  const sm = t => t * t * (3 - 2 * t);
  const n2 = (x, y) => {
    const xi = Math.floor(x) & 255, yi = Math.floor(y) & 255, xf = x - Math.floor(x), yf = y - Math.floor(y);
    const a = G[P[xi + P[yi]]], b = G[P[xi + 1 + P[yi]]], c = G[P[xi + P[yi + 1]]], d = G[P[xi + 1 + P[yi + 1]]];
    const u = sm(xf), v = sm(yf); return a + (b - a) * u + (c - a) * v + (a - b - c + d) * u * v;
  };
  return (x, y, oct = 4) => { let s = 0, amp = .5, f = 1; for (let i = 0; i < oct; i++) { s += amp * n2(x * f, y * f); amp *= .5; f *= 2.03; } return s / (1 - Math.pow(.5, oct)); };
}

// Star ageing. Each layer is a black-body colour that cools from a young hot star (blue-white)
// to an old one (orange-red) as its usage window is consumed:
//   outer rim ← 5-hour session · middle ← this week · core ← this month (see main.js monthlyFrom)
function kelvinRGB(T) {
  const t = T / 100, c = v => Math.max(0, Math.min(255, v));
  const r = t <= 66 ? 255 : 329.698727446 * Math.pow(t - 60, -0.1332047592);
  const g = t <= 66 ? 99.4708025861 * Math.log(t) - 161.1195681661 : 288.1221695283 * Math.pow(t - 60, -0.0755148492);
  const b = t >= 66 ? 255 : t <= 19 ? 0 : 138.5177312231 * Math.log(t - 10) - 305.0447927307;
  return [c(r), c(g), c(b)];
}
const STAR_AGE = { outer: [18000, 1700], mid: [14000, 2700], core: [30000, 3800] };
function starColor(zone, age, sat = 1.6) {
  const [T0, T1] = STAR_AGE[zone], T = Math.exp(Math.log(T0) + (Math.log(T1) - Math.log(T0)) * Math.min(1, Math.max(0, age)));
  const [r, g, b] = kelvinRGB(T), m = (r + g + b) / 3;                      // black-body hues are pale: boost saturation
  return [r, g, b].map(v => Math.max(0, Math.min(255, m + (v - m) * sat)));
}
const sstep = (a, b, x) => { const t = Math.min(1, Math.max(0, (x - a) / (b - a))); return t * t * (3 - 2 * t); };

let MOTHER_NOISE = null;
function makeMotherTextures(ages, S = 360) {
  const C = S / 2;
  if (!MOTHER_NOISE) {                                                       // noise is fixed; only colours change
    const n1 = vnoise(.37), n2 = vnoise(.81), N = S * S;
    MOTHER_NOISE = { m: new Float32Array(N), m2: new Float32Array(N), f: new Float32Array(N) };
    for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
      const i = y * S + x; MOTHER_NOISE.m[i] = n1(x / 26, y / 26); MOTHER_NOISE.m2[i] = n2(x / 11, y / 11); MOTHER_NOISE.f[i] = n2(x / 14 + 9, y / 14 - 3);
    }
  }
  const { m: NM, m2: NM2, f: NF } = MOTHER_NOISE;
  const mk = () => { const c = document.createElement('canvas'); c.width = c.height = S; return c; };
  const body = mk(), flame = mk(), bx = body.getContext('2d'), fx = flame.getContext('2d');
  const bi = bx.createImageData(S, S), fi = fx.createImageData(S, S);
  const cCore = starColor('core', ages.core, 1.25), cMid = starColor('mid', ages.mid), cOut = starColor('outer', ages.outer, 1.8);
  const cMid2 = starColor('mid', Math.min(1, ages.mid + .18), 1.9);        // deeper tone toward the rim for depth
  for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
    const dx = (x - C) / C, dy = (y - C) / C, r = Math.hypot(dx, dy), o = (y * S + x) * 4, i = y * S + x;
    if (r > 1) continue;
    const m = NM[i], m2 = NM2[i], t = Math.min(1, Math.max(0, r + (m - .5) * .09 + (m2 - .5) * .03));
    const wc = 1 - sstep(.08, .26, t), wo = sstep(.62, .8, t), wm = Math.max(0, 1 - wc - wo), mid = sstep(.25, .62, t);
    let col = [0, 1, 2].map(k => wc * (cCore[k] * .55 + 255 * .45) + wm * (cMid[k] * (1 - mid) + cMid2[k] * mid) + wo * cOut[k]);
    const lum = (t < .1 ? 1.1 : .58 + .36 * m2) * (1 + .3 * Math.max(0, 1 - Math.abs(t - .87) / .07));   // mottled body, hot rim band
    const a = r < .9 ? 1 : Math.max(0, 1 - (r - .9) / .1);
    bi.data[o] = Math.min(255, col[0] * lum); bi.data[o + 1] = Math.min(255, col[1] * lum); bi.data[o + 2] = Math.min(255, col[2] * lum); bi.data[o + 3] = 255 * a;
    const rim = Math.max(0, 1 - Math.abs(r - .86) / .14), f = Math.pow(NF[i], 2.2) * rim;
    fi.data[o] = Math.min(255, cOut[0] * 1.1 + 40 * f); fi.data[o + 1] = Math.min(255, cOut[1] * 1.05 + 60 * f); fi.data[o + 2] = Math.min(255, cOut[2] + 40 * f); fi.data[o + 3] = 255 * Math.min(1, f * 1.6);
  }
  bx.putImageData(bi, 0, 0); fx.putImageData(fi, 0, 0);
  return { body, flame, S, outer: cOut };
}

function limitColor(left) { return left < 15 ? '#ff6b7d' : left < 40 ? '#ffcf6b' : '#7dffc4'; }
function fmtReset(t) {
  const m = Math.max(0, (t - Date.now()) / 60000);
  if (!t) return '—';
  if (m < 60) return Math.ceil(m) + 'm';
  if (m < 1440) return Math.floor(m / 60) + 'h ' + Math.floor(m % 60) + 'm';
  return Math.floor(m / 1440) + 'd ' + Math.floor(m % 1440 / 60) + 'h';
}

Object.assign(Mesh.prototype, {
  motherRadius() { return Math.max(56, Math.min(104, Math.min(this.w, this.h) * .13)); },
  galaxyRadius() { return this.motherRadius(); },          // older call sites (orbit sizing, pixel mode)
  starAges() {
    const L = (this.hub.usage && this.hub.usage.limits) || [], pct = k => { const l = L.find(x => x.kind === k); return l ? l.used / 100 : 0; };
    return { outer: pct('session'), mid: pct('weekly_all'), core: (this.hub.usage && this.hub.usage.monthly) || 0 };
  },

  // heat haze: the body texture in thin strips, each nudged sideways by drifting sine waves
  drawMother(ctx, cx, cy, R, smooth = true) {
    const ages = this.starAges(), key = [ages.outer, ages.mid, ages.core].map(v => Math.round(v * 50)).join();
    if (!this.mtex || (this.mtex.key !== key && this.time - (this.mtexAt || 0) > 2)) { this.mtex = makeMotherTextures(ages); this.mtex.key = key; this.mtexAt = this.time; }
    const { body, flame, S } = this.mtex, t = this.time, heat = Math.min(1, (this.hub.mass || 0) / 10);
    R *= 1 + .014 * Math.sin(t * 1.7) + heat * .05;
    ctx.save(); ctx.imageSmoothingEnabled = smooth;
    const N = smooth ? 34 : 24, sh = S / N, dh = 2 * R / N, amp = R * (.022 + heat * .02);
    for (let i = 0; i < N; i++) {
      const off = Math.sin(i * .43 + t * 3.6) * amp + Math.sin(i * .19 - t * 2.1) * amp * .6;
      const x = cx - R + (smooth ? off : Math.round(off)), y = cy - R + i * dh;
      ctx.drawImage(body, 0, i * sh, S, sh + .6, x, y, 2 * R, dh + .6);
    }
    // flickering flame licks on the rim (two counter-rotating copies)
    ctx.globalCompositeOperation = 'lighter';
    ctx.translate(cx, cy);
    ctx.rotate(t * .21); ctx.globalAlpha = .28 + .12 * Math.sin(t * 5.3); ctx.drawImage(flame, -R, -R, 2 * R, 2 * R);
    ctx.rotate(-t * .47); ctx.globalAlpha = .22 + .1 * Math.sin(t * 4.1 + 1); ctx.drawImage(flame, -R * 1.02, -R * 1.02, 2.04 * R, 2.04 * R);
    ctx.restore();
    if (!smooth) return;
    ctx.globalCompositeOperation = 'lighter';
    // white-hot core, pulsing (brighter while bubbles are being swallowed)
    const cr = R * (.16 + .04 * Math.sin(t * 6.2) + heat * .08);
    const core = ctx.createRadialGradient(cx, cy, 0, cx, cy, cr);
    core.addColorStop(0, 'rgba(255,255,255,1)'); core.addColorStop(.35, 'rgba(255,245,255,.65)'); core.addColorStop(1, 'rgba(220,200,255,0)');
    ctx.fillStyle = core; ctx.beginPath(); ctx.arc(cx, cy, cr, 0, 6.29); ctx.fill();
    ctx.globalCompositeOperation = 'source-over';
  },

  // full-res renderer's mother bubble
  drawAttractor(dt) {
    const ctx = this.ctx, h = this.hub, t = this.time;
    if (!Number.isFinite(h.mass)) h.mass = 0;
    h.mass *= Math.pow(.6, dt);
    const R = this.motherRadius();
    // hot outer haze: soft orange corona that breathes, plus a faint shimmering ring of heat
    ctx.globalCompositeOperation = 'lighter';
    const pulse = .5 + .5 * Math.sin(t * 1.3);
    const oc = (this.mtex && this.mtex.outer || [255, 120, 60]).map(Math.round);
    const halo = ctx.createRadialGradient(h.x, h.y, R * .8, h.x, h.y, R * 2.1);
    halo.addColorStop(0, `rgba(${oc},${.32 + .1 * pulse})`); halo.addColorStop(.4, `rgba(${oc},${.1 + .04 * pulse})`); halo.addColorStop(1, `rgba(${oc},0)`);
    ctx.fillStyle = halo; ctx.beginPath(); ctx.arc(h.x, h.y, R * 2.1, 0, 6.29); ctx.fill();
    for (let i = 0; i < 3; i++) {
      const k = ((t * .35 + i / 3) % 1), rr = R * (1.02 + k * .7);
      ctx.strokeStyle = `rgba(${oc},${(1 - k) * .16})`; ctx.lineWidth = 1.5 + (1 - k) * 2;
      ctx.beginPath(); ctx.arc(h.x, h.y, rr, 0, 6.29); ctx.stroke();
    }
    ctx.globalCompositeOperation = 'source-over';
    this.drawMother(ctx, h.x, h.y, R, true);
    if (this.hover === 'hub') { ctx.strokeStyle = 'rgba(255,230,210,.55)'; ctx.setLineDash([2, 6]); ctx.lineWidth = 1; ctx.beginPath(); ctx.arc(h.x, h.y, R * 1.16, t, t + 6.283); ctx.stroke(); ctx.setLineDash([]); }
  },

  // drawn after the sessions so an orbiting bubble never hides it
  drawMotherInfo() {
    const ctx = this.ctx, h = this.hub, R = this.motherRadius();
    const below = this.drawLimits(ctx, h.x, h.y + R * 1.22 + 8);
    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(236,230,255,.92)'; ctx.font = '300 13px system-ui';
    ctx.fillText(`$${(h.usd || 0).toFixed(2)}  today  ·  ${h.working || 0} / ${h.count} active`, h.x, below + 17);
  },

  // plan usage left (what /usage shows): a compact glass card under the mother bubble; returns its bottom y
  drawLimits(ctx, cx, top) {
    const L = (this.hub.usage && this.hub.usage.limits) || [];
    const W = 214, RH = 24, H = 22 + Math.max(1, L.length) * RH, x = cx - W / 2, y = top;
    ctx.save();
    ctx.fillStyle = 'rgba(14,10,46,.72)'; ctx.strokeStyle = 'rgba(200,210,255,.26)'; ctx.lineWidth = 1;
    ctx.beginPath(); ctx.roundRect(x, y, W, H, 12); ctx.fill(); ctx.stroke();
    ctx.textAlign = 'center'; ctx.fillStyle = 'rgba(190,200,255,.75)'; ctx.font = '500 9px system-ui';
    ctx.fillText('U S A G E   L E F T', cx, y + 14);
    if (!L.length) { ctx.fillStyle = 'rgba(220,225,255,.7)'; ctx.font = '300 11px system-ui'; ctx.fillText('loading…', cx, y + 34); ctx.restore(); return y + H; }
    L.forEach((l, i) => {
      const left = Math.max(0, 100 - l.used), ry = y + 20 + i * RH, col = limitColor(left);
      ctx.textAlign = 'left'; ctx.fillStyle = '#eef0ff'; ctx.font = `${l.active ? 600 : 400} 11.5px system-ui`;
      ctx.fillText(l.label, x + 12, ry + 11);
      ctx.textAlign = 'right'; ctx.fillStyle = 'rgba(180,188,235,.75)'; ctx.font = '300 9.5px system-ui';
      ctx.fillText(fmtReset(l.resetsAt), x + W - 50, ry + 11);
      ctx.fillStyle = col; ctx.font = '600 11.5px system-ui'; ctx.fillText(Math.round(left) + '%', x + W - 12, ry + 11);
      ctx.fillStyle = 'rgba(255,255,255,.12)'; ctx.fillRect(x + 12, ry + 16, W - 24, 2.5);
      ctx.fillStyle = col; ctx.fillRect(x + 12, ry + 16, (W - 24) * left / 100, 2.5);
    });
    ctx.restore();
    return y + H;
  },
});
