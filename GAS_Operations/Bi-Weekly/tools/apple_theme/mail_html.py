"""Bi-Weekly mail as live HTML (Apple theme): reads the final deck (report.json apple_deck) and renders the Overview
dashboard, TOP3 gauge tiles, TOP-7 tables and SIREN tables as real HTML (text, numbers, working links) — no screenshots.
Writes mail.html next to this file. Run with the homebrew python3."""
import os, sys, json, re, html as H
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, '..'))
from rate_pipeline import creds
from googleapiclient.discovery import build
from cfg import CFG
import spigen_images as SI

E = 12700
F = "-apple-system,BlinkMacSystemFont,'SF Pro Text','Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',Arial,sans-serif"
INK, GRAY, GRAY2, BLUE, LINK, RED, CANVAS, HAIR = '#1D1D1F', '#6E6E73', '#86868B', '#0071E3', '#0066CC', '#FF3B30', '#F5F5F7', '#E5E5EA'
RING = ['#0071E3', '#64D2FF', '#5E5CE6', '#D2D2D7']
DECK = CFG['apple_deck']; DECK_URL = f'https://docs.google.com/presentation/d/{DECK}/edit?usp=sharing'

p = build('slides', 'v1', credentials=creds()).presentations().get(presentationId=DECK).execute()
def walk(es):
    for e in es:
        if 'elementGroup' in e: yield from walk(e['elementGroup']['children'])
        else: yield e
def runs(e): return [x['textRun'] for x in e.get('shape', {}).get('text', {}).get('textElements', []) if 'textRun' in x]
def txt(e): return ''.join(r['content'] for r in runs(e)).strip()
def xy(e): t = e['transform']; return t.get('translateX', 0) / E, t.get('translateY', 0) / E
def esc(s): return H.escape(s)
def pname(n):   # same rule as Code.js _chartProductName: drop SP_, "_new_biz__ev25mfaq" -> "EV25MFAQ"
    n = n.strip(); n = n[3:] if n.startswith('SP_') else n
    if '_' not in n: return n
    parts = [t for t in n.split('_') if t]; return parts[-1].upper() if parts else n
UNKNOWN = re.compile(r'^(알\s*수\s*없음|해당\s*없음|unknown|n/?a|없음|-)(\s*\(.*\))?$', re.I)
NW = 'white-space:nowrap;'
def title_of(sl): return next((txt(e) for e in walk(sl['pageElements']) if xy(e)[1] < 30 and txt(e)), '')
def slide_link(sid): return f'https://docs.google.com/presentation/d/{DECK}/edit#slide=id.{sid}'

# ---------- building blocks (inline CSS only; Gmail-safe tables) ----------
def tile(inner, pad='22px 24px', mb=14): return f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:18px;margin:0 0 {mb}px"><tr><td style="padding:{pad};font-size: 11px;color:#1D1D1F">{inner}</td></tr></table>'
def h2(t, sub='', href=''):
    a = f'<a href="{href}" style="color:#1d1d1f;text-decoration:none">{esc(t)}</a>' if href else esc(t)
    s = f'<div style="font:600 12px {F};color:{GRAY2};margin-top:3px">{esc(sub)}</div>' if sub else ''
    return f'<div style="font:700 19px {F};color:{INK};margin:4px 0 12px;line-height:1.25">{a}{s}</div>'
def lab(t, extra=''): return f'<div style="font:600 13px {F};color:{INK};margin:2px 0 10px">{esc(t)} <span style="color:{GRAY2};font-weight:400">{esc(extra)}</span></div>'
def num(n): return f'{n:,}'

# ---------- dashboard (slide 3) ----------
def dashboard(sl):
    els = list(walk(sl['pageElements']))
    find = lambda pred: next((e for e in els if pred(txt(e))), None)
    mon = find(lambda t: t.startswith('아마존 리뷰 모니터링'))
    mon_html = ''
    for r in runs(mon):
        c = r['style'].get('foregroundColor', {}).get('opaqueColor', {}).get('rgbColor', {})
        red = c.get('red', 0) > .8 and c.get('green', 1) < .5
        mon_html += f'<span style="color:{RED if red else GRAY}">{esc(r["content"]).replace(chr(10), "<br>")}</span>'
    total = txt(find(lambda t: re.fullmatch(r'[\d,]+', t or '') is not None and len(t) >= 5))
    cum_lab = txt(find(lambda t: '누적 클레임 인입건' in t)); yoy = txt(find(lambda t: 'YoY' in t))
    imgs = [e for e in els if 'image' in e and e['image'].get('imageProperties', {}).get('link')]
    # top-3 reason tiles (rank text '1st/2nd/3rd' + '<reason> <n>' + english) with their Looker link (icon nearest in x)
    groups = [[c for c in walk([g])] for g in sl['pageElements'] if 'elementGroup' in g]
    icons = sorted([i for i in imgs if xy(i)[1] > min(xy(k)[1] for k in imgs) + 1] or imgs, key=lambda i: xy(i)[0])
    tiles = []
    for n, rank in enumerate(('1st', '2nd', '3rd')):
        g = next(g for g in groups if any(txt(c) == rank for c in g)); ts = [txt(c) for c in g if txt(c) and txt(c) != rank]
        main = next(t for t in ts if re.search(r'\d', t)); eng = next((t for t in ts if re.fullmatch(r'[A-Za-z ]+', t)), '')
        tiles.append((rank, main, eng, icons[min(n, len(icons) - 1)]['image']['imageProperties']['link']['url']))
    head_link = min(imgs, key=lambda i: xy(i)[1])['image']['imageProperties']['link']['url']
    cells = ''.join(f'<td width="33%" valign="top" style="padding:0 {6 if i < 2 else 0}px 0 0"><a href="{u}" style="text-decoration:none;display:block;background:{CANVAS};border-radius:14px;padding:12px 14px">'
                    f'<div style="font:700 12px {F};color:{RED if i == 0 else BLUE}">{rk}</div><div style="font:700 15px {F};color:#1d1d1f;margin:4px 0 2px;{NW}">{esc(m)}</div>'
                    f'<div style="font:400 11px {F};color:{GRAY};{NW}">{esc(en)} &rsaquo;</div></a></td>' for i, (rk, m, en, u) in enumerate(tiles))
    # weekly bar chart as HTML (table cells), same data/window as chart.py
    wk = json.load(open(os.path.join(HERE, 'weekly.json')))['weeks']
    key = lambda l: tuple(int(x) for x in re.match(r'(\d+)월 (\d+)주 (\d+)년', l).groups()[::-1][:1] + re.match(r'(\d+)월 (\d+)주', l).groups())
    data = sorted([(key(l), l, n) for l, n, _ in wk]); data = [d for d in data if d[0] >= (2026, 3, 5)]
    mx = max(d[2] for d in data); BH = 120
    fmt = lambda v: f'{v/1000:.1f}k' if v >= 1000 else str(v)
    bars = ''
    for i, (k, l, v) in enumerate(data):
        last = i == len(data) - 1; h = max(3, round(v / mx * BH))
        bars += (f'<td valign="bottom" align="center" style="padding:0 1px"><div style="font:600 8px {F};color:{BLUE if last else GRAY2};margin-bottom:2px">{fmt(v)}</div>'
                 f'<div style="height:{h}px;width:12px;background:{BLUE if last else "#D2D2D7"};border-radius:4px 4px 0 0"></div></td>')
    xl = ''.join(f'<td align="center" style="font:{"700" if k[2] == 1 or i == 0 else "400"} 8px {F};color:{GRAY2};padding-top:4px">{f"{k[1]}월" if (k[2] == 1 or i == 0) else ""}</td>' for i, (k, l, v) in enumerate(data))
    chart = f'<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;border-bottom:1px solid {HAIR}"><tr>{bars}</tr></table><table role="presentation" cellpadding="0" cellspacing="0" style="width:100%"><tr>{xl}</tr></table>'
    inner = (h2('1. Overview', href=slide_link(sl['objectId'])) +
             '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
             f'<td width="38%" valign="top" style="padding-right:14px"><div style="font:600 13px {F};color:{INK};margin-bottom:6px">집중 모니터링 대상</div><div style="font:400 12px {F};line-height:1.6">{mon_html}</div></td>'
             f'<td valign="top"><div style="font:600 13px {F};color:{INK};margin-bottom:8px">2026 최다 인입사유 <a href="{head_link}" style="font-weight:400;color:{LINK};text-decoration:none;font-size:12px">상세 대시보드 &rsaquo;</a></div>'
             f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>{cells}</tr></table></td></tr></table>'
             f'<div style="margin:18px 0 8px"><span style="font:400 12px {F};color:{GRAY}">{esc(cum_lab)}</span><br><span style="font:700 30px {F};color:{INK}">{esc(total)}</span> '
             f'<span style="font:600 13px {F};color:{BLUE}">{esc(yoy)}</span></div>' + chart)
    return tile(inner)

# ---------- TOP3 gauge slides ----------
def gauges(sl):
    els = list(walk(sl['pageElements']))
    t = title_of(sl); m = re.match(r'^(?:1\.\s*)?(Overview \([^)]*\))\s*(.*)$', t)
    out = h2(m.group(1) if m else t, m.group(2) if m else '', slide_link(sl['objectId']))
    for sec in ('모델별 TOP3', '인입사유별 TOP3'):
        s = next(e for e in els if txt(e) == sec); sy = xy(s)[1]
        kind = next((txt(e) for e in els if txt(e) in ('클레임', '배드리뷰') and abs(xy(e)[1] - sy) < 4), '')
        counts = sorted([e for e in els if re.fullmatch(r'[\d,]+건', txt(e)) and 30 < xy(e)[1] - sy < 90], key=lambda e: xy(e)[0])
        cells = ''
        for i, c in enumerate(counts):
            cx, cy = xy(c)
            name = next((txt(e) for e in els if abs(xy(e)[0] - cx) < 3 and 10 < xy(e)[1] - cy < 25), '')
            L = next((txt(e).split('\n') for e in els if abs(xy(e)[0] - cx - 7) < 3 and 25 < xy(e)[1] - cy < 45), [])
            V = next((txt(e).split('\n') for e in els if 60 < xy(e)[0] - cx < 100 and 25 < xy(e)[1] - cy < 45), [])
            L = [pname(l) for l in L]; name = pname(name)
            pairs = list(zip(L, V)); unk = sum(int(v) for l, v in pairs if UNKNOWN.match(l) and v.strip().isdigit())
            pairs = [(l, v) for l, v in pairs if not UNKNOWN.match(l)]
            if unk:
                pairs = [(l, str(int(v) + unk) if l == '그 외' else v) for l, v in pairs]
                if not any(l == '그 외' for l, _ in pairs): pairs.append(('그 외', str(unk)))
            L, V = [l for l, _ in pairs], [v for _, v in pairs]
            vals = [int(v) for v in V if v.strip().isdigit()]; tot = sum(vals) or 1
            bar = ''.join(f'<td style="background:{RING[j % 4]};height:6px;width:{max(1, round(v / tot * 100))}%"></td>' for j, v in enumerate(vals) if v)
            rows = ''.join(f'<tr><td style="font:400 11px {F};color:{INK};padding:2px 0;{NW}max-width:132px;overflow:hidden;text-overflow:ellipsis"><span style="display:inline-block;width:7px;height:7px;border-radius:4px;background:{RING[j % 4]};margin-right:5px"></span>{esc(l)}</td>'
                           f'<td align="right" style="font:400 11px {F};color:{GRAY};padding:2px 0 2px 6px;{NW}">{v}</td></tr>' for j, (l, v) in enumerate(zip(L, V)))
            cells += (f'<td width="33%" valign="top" style="padding:0 {8 if i < 2 else 0}px 0 0"><div style="background:{CANVAS};border-radius:14px;padding:14px 14px 10px">'
                      f'<div style="font:600 10px {F};color:{GRAY2}">{i + 1}</div><div style="font:700 20px {F};color:{INK};line-height:1.2;{NW}">{esc(txt(c))}</div>'
                      f'<div style="font:400 11px {F};color:{GRAY};margin:2px 0 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{esc(name)}</div>'
                      f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-radius:3px;overflow:hidden;margin-bottom:8px"><tr>{bar}</tr></table>'
                      f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0">{rows}</table></div></td>')
        out += lab(sec, kind) + f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:14px"><tr>{cells}</tr></table>'
    return tile(out, '22px 24px 8px')

# ---------- TOP-7 tables ----------
cat = SI.catalog()
def top7(sl):
    tb = next(e for e in walk(sl['pageElements']) if 'table' in e)['table']
    cell = lambda r, c: ''.join(x.get('textRun', {}).get('content', '') for x in tb['tableRows'][r]['tableCells'][c].get('text', {}).get('textElements', [])).strip()
    foot = next((txt(e) for e in walk(sl['pageElements']) if txt(e).startswith('출처')), '')
    HL = {1: '제품 (기종 / 모델 / 색상)', 2: 'SKU', 3: 'ASIN', 4: '판매량(EU)', 5: '클레임', 6: '배드리뷰', 7: '비율'}
    hdr = ''.join(f'<td style="font:600 10px {F};color:{GRAY2};padding:6px 4px;border-bottom:1px solid {HAIR};{NW}" align="{"left" if c == 1 else "center"}">{HL.get(c, esc(cell(0, c)))}</td>' for c in range(tb['columns']) if c != 0)
    rows = ''
    for r in range(1, tb['rows']):
        name = cell(r, 1).split('\n'); sku, asin = cell(r, 2), cell(r, 3); rate = cell(r, 7)
        img = SI.resolve(cat, sku, name[0])
        pic = f'<img src="{img}" width="30" height="30" style="display:block;width:30px;height:30px;object-fit:contain;border-radius:8px;background:{CANVAS}">' if img else f'<div style="width:30px;height:30px;border-radius:8px;background:{CANVAS}"></div>'
        red = float(rate.rstrip('%') or 0) >= 2
        rows += (f'<tr><td style="padding:7px 4px;border-bottom:1px solid {HAIR}"><table role="presentation" cellpadding="0" cellspacing="0"><tr><td style="padding-right:8px">{pic}</td>'
                 f'<td style="font:600 11px {F};color:{INK}">{esc(name[0])}<div style="font:400 10px {F};color:{GRAY2}">{esc(" ".join(name[1:]))}</div></td></tr></table></td>'
                 + ''.join(f'<td align="center" style="font:400 11px {F};color:{INK};padding:7px 4px;border-bottom:1px solid {HAIR};{NW}">{esc(v)}</td>' for v in (sku, asin, cell(r, 4), cell(r, 5), cell(r, 6)))
                 + f'<td align="center" style="font:600 11px {F};color:{RED if red else INK};padding:7px 4px;border-bottom:1px solid {HAIR};{NW}">{esc(rate)}</td></tr>')
    t = title_of(sl); m = re.match(r'^(Overview \([^)]*\))\s*(.*)$', t)
    return tile(h2(m.group(1), m.group(2), slide_link(sl['objectId'])) + f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>{hdr}</tr>{rows}</table>'
                f'<div style="font:400 9px {F};color:{GRAY2};margin-top:8px;line-height:1.5">{esc(foot)}</div>')

# ---------- SIREN tables ----------
def siren(slides):
    rows, count, sheet = '', '', ''
    for sl in slides:
        els = list(walk(sl['pageElements']))
        count = count or next((txt(e) for e in els if re.fullmatch(r'\d{1,4}', txt(e) or '') and xy(e)[1] > 280), '')
        for e in els:
            for r in runs(e):
                if r['content'].strip() and 'SIREN 시트' in txt(e) and r['style'].get('link'): sheet = r['style']['link']['url']
        tb = next(e for e in els if 'table' in e)['table']
        for r in tb['tableRows'][1:]:
            cs = r['tableCells']
            def ct(c): return ''.join(x.get('textRun', {}).get('content', '') for x in cs[c].get('text', {}).get('textElements', [])).strip()
            link = next((x['textRun']['style']['link']['url'] for x in cs[1].get('text', {}).get('textElements', []) if x.get('textRun', {}).get('style', {}).get('link')), '')
            view = f'<a href="{link}" style="color:{LINK};text-decoration:none;font-weight:600">보기</a>' if link else ''
            new = ct(0) >= CFG['period']['start'].replace('-', '/')
            rows += (f'<tr><td style="font:400 11px {F};color:{RED if new else INK};padding:7px 4px;border-bottom:1px solid {HAIR};white-space:nowrap">{esc(ct(0))}</td>'
                     f'<td align="center" style="font:400 11px {F};padding:7px 4px;border-bottom:1px solid {HAIR};{NW}">{view}</td>'
                     f'<td style="font:400 11px {F};color:{INK};padding:7px 4px;border-bottom:1px solid {HAIR}">{esc(ct(2))}</td>'
                     f'<td style="font:400 11px {F};color:{INK};padding:7px 4px;border-bottom:1px solid {HAIR};{NW}">{esc(ct(3)).replace(chr(10), "<br>")}</td>'
                     f'<td style="font:400 11px {F};color:{INK};padding:7px 4px;border-bottom:1px solid {HAIR}">{esc(ct(4))}</td>'
                     f'<td align="center" style="font:400 11px {F};color:{INK};padding:7px 4px;border-bottom:1px solid {HAIR}">{esc(ct(5))}</td></tr>')
    hdr = ''.join(f'<td style="font:600 10px {F};color:{GRAY2};padding:6px 4px;border-bottom:1px solid {HAIR};{NW}" align="{a}">{h}</td>' for h, a in (('날짜', 'left'), ('SIREN', 'center'), ('기종+제품', 'left'), ('SKU', 'left'), ('이슈', 'left'), ('등록', 'center')))
    foot = (f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px"><tr><td><a href="{sheet}" style="display:inline-block;border:1.5px solid {BLUE};color:{BLUE};font:600 12px {F};text-decoration:none;padding:8px 16px;border-radius:980px">26년 SIREN 시트 열기</a></td>'
            f'<td align="right" style="{NW}"><span style="font:400 12px {F};color:{GRAY}">2026 SIREN Registered by GCX</span> <span style="font:700 26px {F};color:{INK};vertical-align:middle">{esc(count)}</span></td></tr></table>')
    return tile(h2(title_of(slides[0]), href=slide_link(slides[0]['objectId'])) + f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>{hdr}</tr>{rows}</table>' + foot)

# ---------- assemble ----------
S = p['slides']
dash = next(s for s in S if title_of(s) in ('1. Overview', 'Overview'))
gauge_sl = [s for s in S[:12] if any(txt(e) == '모델별 TOP3' for e in walk(s['pageElements']))]
top7_sl = [s for s in S if '판매량 TOP 7' in title_of(s)]
siren_sl = [s for s in S if 'GCX SIREN (26년' in title_of(s)]
order = {'iphone18': 0, 'pixel11': 1, 'glxZ8': 2}
top7_sl.sort(key=lambda s: next((v for k, v in (('iPhone', 0), ('Pixel', 1), ('Galaxy', 2)) if k in title_of(s)), 9))
body = dashboard(dash) + ''.join(gauges(s) for s in gauge_sl) + ''.join(top7(s) for s in top7_sl) + siren(siren_sl)
body = re.sub(r'font:(\d+) (\d+)px ' + re.escape(F), lambda m: f'font-weight:{m.group(1)};font-size:{m.group(2)}px', body)
body = re.sub(r'<span style="display:inline-block;width:7px;height:7px;border-radius:4px;background:#[0-9A-F]+;margin-right:5px"></span>', '', body)
body = body.replace('color:#1D1D1F;', '').replace(';color:#1D1D1F"', '"').replace('font-weight:400;', '').replace('font-weight:400', '')
body = body.replace('font-weight:400;font-size:11px;', '').replace('font-size:11px;', '').replace(';font-size:11px"', '"').replace('font-size:11px"', '"')
body = body.replace(' style=""', '')
body = re.sub(r'<table width="100%" cellspacing="0"(?! style)', '<table width="100%" cellspacing="0" style="font-size:11px"', body).replace('padding:7px 4px;border-bottom:1px solid #E5E5EA', 'padding:7px 4px;border-bottom:1px solid #EEE')
body = body.replace(' role="presentation"', '').replace(' cellpadding="0" cellspacing="0"', ' cellspacing="0"')
open(os.path.join(HERE, 'mail_sections.html'), 'w').write(body)
print('sections', len(body), 'chars', len(body.encode()), 'bytes;', len(gauge_sl), 'gauge,', len(top7_sl), 'top7,', len(siren_sl), 'siren')
