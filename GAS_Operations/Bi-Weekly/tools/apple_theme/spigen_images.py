"""Official Spigen product shots (spigen.com public Shopify catalog) on the Apple deck:
  - claim/review cards: rounded #F5F5F7 product tile + official shot left of the product name
  - TOP-7 tables: a thumbnail in each product row
Matching: exact base SKU (first 8 chars of the variant SKU); fallback = same model name + same series title.
Idempotent: removes its own `sp_` elements first. Run after apple_cards / apple_tables (they reset positions)."""
import os, sys, re, json, time, urllib.request
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
from apple_cards import svc, P, E, rgb, txt
CACHE = os.path.join(HERE, 'spigen_catalog.json')

def catalog(refresh=False):
    if os.path.exists(CACHE) and not refresh and time.time() - os.path.getmtime(CACHE) < 7 * 86400:
        return json.load(open(CACHE))
    prods, page = [], 1
    while True:
        req = urllib.request.Request(f'https://www.spigen.com/products.json?limit=250&page={page}', headers={'User-Agent': 'Mozilla/5.0'})
        ps = json.load(urllib.request.urlopen(req, timeout=60)).get('products', [])
        if not ps: break
        prods += ps; page += 1; time.sleep(0.5)
    sku, titles = {}, {}
    for p in prods:
        if not p.get('images'): continue
        titles[p['title']] = p['images'][0]['src']
        for v in p['variants']:
            s = (v.get('sku') or '').strip().upper()[:8]
            if re.match(r'^[A-Z]{3}\d{5}$', s):
                sku.setdefault(s, (v.get('featured_image') or {}).get('src') or p['images'][0]['src'])
    json.dump({'sku': sku, 'titles': titles}, open(CACHE, 'w'))
    return {'sku': sku, 'titles': titles}

def norm(s): return re.sub(r'[^a-z0-9]+', ' ', s.lower().replace('magfit', 'mag fit')).strip()
def series_of(device):
    d = device.lower()
    for k, t in (('fold 8', 'galaxy z fold 8'), ('flip 8', 'galaxy z flip 8'), ('pixel 11', 'pixel 11'), ('iphone 18', 'iphone 18'), ('iphone 17', 'iphone 17')):
        if k in d: return t
    return None
def resolve(cat, sku, product_line=''):
    url = cat['sku'].get((sku or '').upper())
    if not url and 'ㅣ' in product_line:
        model, device = [x.strip() for x in product_line.split('ㅣ', 1)]
        ser = series_of(device)
        if ser:
            mod = norm(model)
            for title, u in cat['titles'].items():
                tn = norm(title)
                if tn.startswith(ser) and tn.split(' series ')[-1].startswith(mod): url = u; break
    return url and url + ('&' if '?' in url else '?') + 'width=300'

def geo(e):
    tr = e['transform']; return (tr.get('translateX', 0)/E, tr.get('translateY', 0)/E,
                                 e['size']['width']['magnitude']*tr.get('scaleX', 1)/E, e['size']['height']['magnitude']*tr.get('scaleY', 1)/E)
def image(sid, oid, url, x, y, w, h):
    return {'createImage': {'objectId': oid, 'url': url, 'elementProperties': {'pageObjectId': sid,
            'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': h*E, 'unit': 'EMU'}},
            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}}

def main():
    from apple_cards import rtile
    cat = catalog('--refresh' in sys.argv)
    p = svc.get(presentationId=P).execute()
    reqs, hit, miss = [], 0, []
    for s in p['slides']:
        sid, els = s['objectId'], s['pageElements']; k = sid[-10:].replace('_', '')
        reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in els if e['objectId'].startswith('sp_')]
        # --- claim/review card ---
        if any(e['objectId'].startswith('at_ph_') for e in els):
            sku = next((txt(e) for e in els if 'shape' in e and re.fullmatch(r'A[A-Z]{2}\d{5}', txt(e) or '') and geo(e)[0] > 560), '')
            prod = next((e for e in els if 'shape' in e and 'ㅣ' in (txt(e) or '') and geo(e)[1] < 95 and geo(e)[0] < 120), None)
            date = next((e for e in els if 'shape' in e and re.fullmatch(r'\d{4}\.\d{2}\.\d{2}', txt(e) or '') and geo(e)[0] < 120), None)
            url = resolve(cat, sku, txt(prod) if prod else '')
            if url and prod:
                r, ids = rtile(sid, f'sp_t_{k}', 50, 79, 38, 38, '#F5F5F7'); reqs += r
                reqs.append(image(sid, f'sp_i_{k}', url, 53, 82, 32, 32))
                for e, y in ((prod, 79), (date, 99)):
                    if e:
                        tr = dict(e['transform']); tr['translateX'] = 92*E; tr['unit'] = 'EMU'
                        reqs.append({'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': tr}})
                hit += 1
            else:
                miss.append(sku or sid)
                for e in (prod, date):   # no image → keep the standard text position
                    if e:
                        tr = dict(e['transform']); tr['translateX'] = 50*E; tr['unit'] = 'EMU'
                        reqs.append({'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': tr}})
        # --- TOP-7 table rows ---
        for e in els:
            if 'table' not in e or not e['objectId'].startswith('rate_'): continue
            tb = e['table']; tid = e['objectId']; tx, ty, _, _ = geo(e)
            hdr = [txt({'shape': {'text': c.get('text', {})}}) for c in tb['tableRows'][0]['tableCells']]
            if 'SKU' not in hdr: continue
            ci_sku, ci_prod = hdr.index('SKU'), 1
            x0 = tx + sum(c['columnWidth']['magnitude'] for c in tb['tableColumns'][:ci_prod]) / E
            # Slides renders rows taller than rowHeight when text wraps, so thumbnails drifted.
            # Pin every row to a height that fits the tallest content (2-line name + 사유 line):
            # rendered height == set height, and thumbnail centers can be computed exactly.
            HH, RH = 24, 40
            heights = [HH] + [RH] * (len(tb['tableRows']) - 1)
            reqs.append({'updateTableRowProperties': {'objectId': tid, 'rowIndices': [0], 'tableRowProperties': {'minRowHeight': {'magnitude': HH, 'unit': 'PT'}}, 'fields': 'minRowHeight'}})
            reqs.append({'updateTableRowProperties': {'objectId': tid, 'rowIndices': list(range(1, len(heights))), 'tableRowProperties': {'minRowHeight': {'magnitude': RH, 'unit': 'PT'}}, 'fields': 'minRowHeight'}})
            reqs.append({'updateTableCellProperties': {'objectId': tid, 'tableRange': {'location': {'rowIndex': 0, 'columnIndex': 0}, 'rowSpan': len(heights), 'columnSpan': tb['columns']},
                         'tableCellProperties': {'contentAlignment': 'MIDDLE'}, 'fields': 'contentAlignment'}})
            # white tile behind the table, resized to the pinned height
            tw = sum(c['columnWidth']['magnitude'] for c in tb['tableColumns']) / E
            old = [x['objectId'] for x in els if x['objectId'].startswith('at_tb_' + tid[-8:])]
            reqs += [{'deleteObject': {'objectId': o}} for o in old]
            r, tids = rtile(sid, 'sp_tb_' + tid[-8:], tx - 12, ty - 8, tw + 24, sum(heights) + 16); reqs += r
            reqs.append({'updatePageElementsZOrder': {'pageElementObjectIds': tids, 'operation': 'SEND_TO_BACK'}})
            y = ty
            for ri, row in enumerate(tb['tableRows']):
                h = heights[ri]
                if ri > 0:
                    sku = txt({'shape': {'text': row['tableCells'][ci_sku].get('text', {})}})
                    first = txt({'shape': {'text': row['tableCells'][ci_prod].get('text', {})}}).split('\n')[0]
                    parts = [x.strip() for x in re.split(r'[|｜│ㅣ]', first)]
                    url = resolve(cat, sku, f'{parts[1]} ㅣ{parts[0]}' if len(parts) >= 2 else '')
                    if url:
                        reqs.append(image(sid, f'sp_r_{tid[-6:]}_{ri}', url, x0 + 5, y + (h - 26) / 2, 26, 26)); hit += 1
                    else:            # keep the thumbnail column even: empty light-gray tile
                        import apple_cards as _ac; _r = _ac.R; _ac.R = 6
                        reqs += rtile(sid, f'sp_e_{tid[-6:]}_{ri}', x0 + 5, y + (h - 26) / 2, 26, 26, '#F5F5F7')[0]; _ac.R = _r
                        miss.append(sku)
                    reqs.append({'updateParagraphStyle': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': ci_prod}, 'textRange': {'type': 'ALL'},
                                 'style': {'indentStart': {'magnitude': 30, 'unit': 'PT'}, 'indentFirstLine': {'magnitude': 30, 'unit': 'PT'}}, 'fields': 'indentStart,indentFirstLine'}})
                y += h
    for i in range(0, len(reqs), 200):
        for _ in range(8):
            try: svc.batchUpdate(presentationId=P, body={'requests': reqs[i:i+200]}).execute(); break
            except Exception as ex:
                if '429' in str(ex): time.sleep(20); continue
                raise
    print('images placed', hit, '| no official image:', sorted(set(miss)))

if __name__ == '__main__':
    main()
