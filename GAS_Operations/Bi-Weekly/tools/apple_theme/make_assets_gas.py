"""Writes ../../AppleAssets.js — a TEMP Apps Script that, on the Apple deck:
  1) places the Material Symbols icons next to card spec labels / source line / index numbers /
     gauge headers / Appendix headings (skips slides that already have icons),
  2) swaps the slide-3 weekly chart (image titled 'apl_chart', or the old Looker screenshot) with weekly_chart.png,
  3) puts icons on the Apple layouts (p5 gauges, p7 card, p9 appendix).
Push with clasp, run `aaRun` from the editor (it is the first function of AppleAssets.gs), then delete the file and push again.
Icons: Material Symbols Rounded (Apache-2.0); render with icons_render.py if icons/*.png are missing."""
import os, json, base64
HERE = os.path.dirname(os.path.abspath(__file__))
from cfg import CFG
b64 = lambda p: base64.b64encode(open(p, 'rb').read()).decode()
icons = {k[:-4]: b64(os.path.join(HERE, 'icons', k)) for k in os.listdir(os.path.join(HERE, 'icons')) if k.endswith('.png') and k != 'contact.png'}
chart = b64(os.path.join(HERE, 'weekly_chart.png')) if os.path.exists(os.path.join(HERE, 'weekly_chart.png')) else ''
LAYOUT = [('p5','category',146,58,12),('p5','donut',146,232,12),('p7','bag',466,83,10),('p7','globe',466,207.5,9),('p7','report',586,207.5,9),
          ('p7','code',466,267.5,9),('p7','box',586,267.5,9),('p7','star',466,327.5,9),('p9','photos',43,85,10),('p9','sheet',175,85,10),
          ('p9','db',307,85,10),('p9','siren',439,85,10),('p9','doc',571,85,10)]
js = '''// TEMP (bi-weekly-builder): icons + weekly chart + layout icons for the Apple deck. Delete after running.
function aaRun() { aaPlaceIcons(); aaSwapChart(); aaLayoutIcons(); }
const AA_DECK = '%(deck)s';
const AA_ICONS = %(icons)s;
const AA_CHART = "%(chart)s";
const AA_LAYOUT = %(layout)s;
function _aaBlob(b64, n) { return Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', n); }
function _aaIcon(page, key, x, y, s, title) { const im = page.insertImage(_aaBlob(AA_ICONS[key], key + '.png'), x, y, s, s); im.setTitle(title || 'apl_icon'); return im; }
function _aaText(el) { try { return el.asShape().getText().asString().trim(); } catch (e) { return ''; } }
function aaPlaceIcons() {
  const pres = SlidesApp.openById(AA_DECK);
  const cardKeys = ['globe', 'report', 'code', 'box', 'star'], appx = ['photos', 'sheet', 'db', 'siren', 'doc'];
  const idxKeys = { '01': 'insights', '02': 'fold', '03': 'phone', '04': 'iphone', '05': 'siren' };
  pres.getSlides().forEach(function(slide) {
    const els = slide.getPageElements();
    if (els.some(function(e) { return e.getTitle && e.getTitle() === 'apl_icon'; })) return;
    els.forEach(function(el) {
      const id = el.getObjectId(), t = _aaText(el); let m = id.match(/^at_lb(\\d)_/);
      if (m) { _aaIcon(slide, cardKeys[+m[1]], el.getLeft() + 7, el.getTop() + 1.5, 9); el.setLeft(el.getLeft() + 11); return; }
      if (/^at_src_/.test(id)) { _aaIcon(slide, t.indexOf('Amazon') >= 0 ? 'bag' : 'agent', el.getLeft() + 7, el.getTop() + 2, 10); el.setLeft(el.getLeft() + 12); return; }
      m = id.match(/^ap3_a_c(\\d)$/);
      if (m) { _aaIcon(slide, appx[+m[1]], el.getLeft() + 7, el.getTop() + 3, 10); el.setLeft(el.getLeft() + 13); return; }
      if (t === '모델별 TOP3' || t === '인입사유별 TOP3') {
        const L = el.getLeft(), T = el.getTop();
        els.forEach(function(o) { if (o !== el && Math.abs(o.getTop() - T) < 3 && o.getLeft() > L) o.setLeft(o.getLeft() + 15); });
        _aaIcon(slide, t === '모델별 TOP3' ? 'category' : 'donut', L + 7, T + 7, 12); el.setLeft(L + 15);
      }
    });
    // index slide: icon right of each section number (also inside groups)
    const all = []; (function walk(list) { list.forEach(function(e) { if (e.getPageElementType() === SlidesApp.PageElementType.GROUP) walk(e.asGroup().getChildren()); else all.push(e); }); })(els);
    if (all.filter(function(e) { return idxKeys[_aaText(e)]; }).length >= 4)
      all.forEach(function(e) { const k = idxKeys[_aaText(e)]; if (k) _aaIcon(slide, k, e.getLeft() + 48, e.getTop() + 7, 20); });
  });
  pres.saveAndClose();
}
function aaSwapChart() {
  if (!AA_CHART) return;
  const pres = SlidesApp.openById(AA_DECK);
  pres.getSlides().some(function(s) {
    const imgs = s.getPageElements().filter(function(e) { return e.getPageElementType() === SlidesApp.PageElementType.IMAGE; });
    const old = imgs.filter(function(e) { return e.getTitle() === 'apl_chart'; })[0] ||
                imgs.filter(function(e) { return e.getWidth() > 500 && e.getHeight() > 200 && s.getPageElements().some(function(o) { return _aaText(o).indexOf('누적 클레임 인입건') >= 0; }); })[0];
    if (!old) return false;
    const im = s.insertImage(_aaBlob(AA_CHART, 'weekly_chart.png'), old.getLeft(), old.getTop(), old.getWidth(), old.getHeight());
    im.setTitle('apl_chart'); im.sendToBack(); old.remove(); return true;
  });
  pres.saveAndClose();
}
function aaLayoutIcons() {
  const pres = SlidesApp.openById(AA_DECK); const byId = {};
  pres.getLayouts().forEach(function(l) { byId[l.getObjectId()] = l; l.getPageElements().forEach(function(e) { if (e.getTitle && e.getTitle() === 'lay_icon') e.remove(); }); });
  AA_LAYOUT.forEach(function(p) { if (byId[p[0]]) _aaIcon(byId[p[0]], p[1], p[2], p[3], p[4], 'lay_icon'); });
  pres.saveAndClose();
}
''' % {'deck': CFG['apple_deck'], 'icons': json.dumps(icons), 'chart': chart, 'layout': json.dumps(LAYOUT)}
out = os.path.join(HERE, '..', '..', 'AppleAssets.js')
open(out, 'w').write(js); print('wrote', os.path.normpath(out), len(js), 'chars; icons', sorted(icons))
