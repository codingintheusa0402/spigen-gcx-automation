"""SIREN issue cards (§7c): one Apple claim/review card per pick in siren_cards.json, inserted right after the last
'5. GCX SIREN (26년 …)' table slide of the FINAL deck (report.json apple_deck). Idempotent per pick (card id = siren_card_<n>).
siren_cards.json: {"picks": [{"deck": <SIREN deck id>, "slide": <slide objectId>, "type": "claim"|"review",
  "sku": "ACS11561", "asin": "B0…", "reason": "카메라링 변형", "product": "Air Skin Aramid MagFit ㅣGlx. Z Fold 8",
  "rating": "4.5점  Global Ratings: 427", "detail": "<optional shortened text>"}]}
Source slide layout (siren-report decks): table cells 구분/국가/구매일/인입일/결함기간/링크 header row, then values, then 상세 내용."""
import sys, os, json, re
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, '..'))
from rate_pipeline import creds
from googleapiclient.discovery import build
from cfg import CFG
S = build('slides', 'v1', credentials=creds()).presentations()
N = os.environ.get('BW_DECK') or CFG['apple_deck']; E = 12700
T = lambda e: ''.join(x.get('textRun', {}).get('content', '') for x in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()
picks = json.load(open(os.path.join(HERE, 'siren_cards.json')))['picks']

def source(deck, sid):
    sl = next(x for x in S.get(presentationId=deck).execute()['slides'] if x['objectId'] == sid)
    cells, imgs, links = [], [], []
    def w(es):
        for e in es:
            if 'elementGroup' in e: w(e['elementGroup']['children'])
            if 'image' in e:
                tr, sz = e['transform'], e['size']; W = sz['width']['magnitude'] * tr.get('scaleX', 1) / E; Hh = sz['height']['magnitude'] * tr.get('scaleY', 1) / E
                if W > 60: imgs.append({'url': e['image']['contentUrl'], 'w': W, 'h': Hh})
            for r in e.get('table', {}).get('tableRows', []):
                for c in r['tableCells']:
                    te = c.get('text', {}).get('textElements', [])
                    links.extend(x['textRun']['style']['link'].get('url') for x in te if x.get('textRun', {}).get('style', {}).get('link'))
                    t = ''.join(x.get('textRun', {}).get('content', '') for x in te).strip()
                    if t: cells.append(t)
    w(sl['pageElements'])
    v = cells[6:12]   # country, purchase, inbound, period, ref
    detail = re.sub(r'^\[[A-Z]{3}\d{5}\]\s*', '', cells[12]).replace('\n', ' ').strip()
    return {'country': v[0], 'date': v[2].replace('-', '.'), 'ref': v[4], 'detail': detail, 'imgs': imgs, 'link': next((l for l in links if l), '')}

p = S.get(presentationId=N).execute()
ids = [x['objectId'] for x in p['slides']]
tpl = next(sl for sl in p['slides'] if any(T(e) == 'SIREN 등록됨' for e in sl['pageElements']) and any(T(e) == '클레임 내용' for e in sl['pageElements'])
           and not sl['objectId'].startswith('siren_card'))
els = {T(e) or e['objectId']: e for e in tpl['pageElements']}
role = lambda pred: next(e['objectId'] for e in tpl['pageElements'] if pred(T(e)))
R = {'title': role(lambda t: t.endswith('Claims / Reviews')), 'detail': role(lambda t: len(t) > 25 and not t.endswith('Reviews') and 'ㅣ' not in t and 'Global' not in t),
     'date': role(lambda t: re.fullmatch(r'\d{4}\.\d{2}\.\d{2}', t) is not None), 'product': role(lambda t: 'ㅣ' in t),
     'rating': role(lambda t: 'Global Ratings' in t), 'sku': role(lambda t: re.fullmatch(r'A[A-Z]{2}\d{5}', t) is not None),
     'asin': role(lambda t: re.fullmatch(r'B0[A-Z0-9]{8}', t) is not None), 'btn': role(lambda t: t.endswith('바로가기')),
     'badge': role(lambda t: t == 'SIREN 등록됨'), 'head': role(lambda t: t in ('클레임 내용', '리뷰 내용')), 'src': role(lambda t: t.startswith('From '))}
geo = lambda e: (e['transform'].get('translateX', 0) / E, e['transform'].get('translateY', 0) / E)
cty = next(e for e in tpl['pageElements'] if re.fullmatch(r'[A-Z]{2}', T(e)))
rsn = min([e for e in tpl['pageElements'] if T(e) and e is not cty and abs(geo(e)[1] - geo(cty)[1]) < 3 and geo(e)[0] > geo(cty)[0]], key=lambda e: geo(e)[0])
R['country'], R['reason'] = cty['objectId'], rsn['objectId']
# button text style from an existing review card ('배드리뷰 바로가기' fits there) and the template claim button
btn_style = lambda txt: next((dict((k, v) for k, v in x['textRun']['style'].items() if k != 'link') for sl in p['slides'] if not sl['objectId'].startswith('siren_card')
                              for e in sl['pageElements'] if T(e) == txt for x in e['shape']['text']['textElements'] if 'textRun' in x), None)
photos = [e for e in tpl['pageElements'] if 'image' in e and e['size']['width']['magnitude'] * e['transform'].get('scaleX', 1) / E > 60]
src_icon = next((e for e in tpl['pageElements'] if e.get('title') == 'apl_icon' and abs(geo(e)[1] - geo(next(x for x in tpl['pageElements'] if x['objectId'] == R['src']))[1]) < 10), None)
amz_icon = None
for sl in p['slides']:
    if any(T(e) == 'From Amazon' for e in sl['pageElements']):
        a = next((e for e in sl['pageElements'] if e.get('title') == 'apl_icon' and src_icon and abs(geo(e)[1] - geo(src_icon)[1]) < 3), None)
        if a: amz_icon = a['image']['contentUrl']; break
old = lambda oid: T(next(e for e in tpl['pageElements'] if e['objectId'] == oid))
reqs, new_ids = [], []
for n, pk in enumerate(picks):
    new = f'siren_card_{n}'
    if new in ids: reqs.append({'deleteObject': {'objectId': new}})
    sdat = source(pk['deck'], pk['slide']); new_ids.append(new)
    m = {e['objectId']: f'{new}_{i}' for i, e in enumerate(tpl['pageElements'])}
    reqs.append({'duplicateObject': {'objectId': tpl['objectId'], 'objectIds': {tpl['objectId']: new, **m}}})
    def st(key, val):
        o = R[key]; L = len(old(o))
        return [{'insertText': {'objectId': m[o], 'insertionIndex': 0, 'text': val}},
                {'deleteText': {'objectId': m[o], 'textRange': {'type': 'FIXED_RANGE', 'startIndex': len(val), 'endIndex': len(val) + L}}}]
    review = pk['type'] == 'review'
    for k, val in (('title', '5. GCX SIREN 신규 등록 이슈 사례'), ('detail', pk.get('detail') or sdat['detail']), ('date', sdat['date']),
                   ('country', sdat['country']), ('reason', pk['reason']), ('asin', pk['asin']), ('sku', pk['sku']),
                   ('rating', pk['rating']), ('product', pk['product'])):
        reqs += st(k, val)
    if review: reqs += st('btn', '배드리뷰 바로가기') + st('head', '리뷰 내용') + st('src', 'From Amazon')
    link = sdat['link'] or (f"https://spigenhelp.zendesk.com/agent/tickets/{sdat['ref']}" if not review else '')
    bs = btn_style('배드리뷰 바로가기' if review else '클레임 바로가기')
    if bs: reqs.append({'updateTextStyle': {'objectId': m[R['btn']], 'textRange': {'type': 'ALL'}, 'style': bs, 'fields': ','.join(bs)}})
    if link: reqs.append({'updateTextStyle': {'objectId': m[R['btn']], 'textRange': {'type': 'ALL'}, 'style': {'link': {'url': link}}, 'fields': 'link'}})
    reqs.append({'updateTextStyle': {'objectId': m[R['btn']], 'textRange': {'type': 'ALL'}, 'style': {'underline': False,
                 'foregroundColor': {'opaqueColor': {'rgbColor': {'red': 1, 'green': 1, 'blue': 1}}}}, 'fields': 'underline,foregroundColor'}})
    reqs.append({'updateTextStyle': {'objectId': m[R['badge']], 'textRange': {'type': 'ALL'}, 'style': {'link': {'url': f"https://docs.google.com/presentation/d/{pk['deck']}/edit"}}, 'fields': 'link'}})
    if review and amz_icon and src_icon: reqs.append({'replaceImage': {'imageObjectId': m[src_icon['objectId']], 'url': amz_icon}})
    reqs += [{'deleteObject': {'objectId': m[ph['objectId']]}} for ph in photos]
    AX, AY, AW, AH, G = 50, 124, 376, 248, 9; imgs = sdat['imgs'][:3]; cw = (AW - G * (len(imgs) - 1)) / max(1, len(imgs))
    for k, im in enumerate(imgs):
        sc = min(cw / im['w'], AH / im['h']); w, h = im['w'] * sc, im['h'] * sc
        reqs.append({'createImage': {'objectId': f'{new}_ph{k}', 'url': im['url'], 'elementProperties': {'pageObjectId': new,
                     'size': {'width': {'magnitude': w * E, 'unit': 'EMU'}, 'height': {'magnitude': h * E, 'unit': 'EMU'}},
                     'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': (AX + k * (cw + G) + (cw - w) / 2) * E, 'translateY': (AY + (AH - h) / 2) * E, 'unit': 'EMU'}}}})
S.batchUpdate(presentationId=N, body={'requests': reqs}).execute()
# position: move cards one by one (re-reading the order each time) to sit right after the last SIREN table slide
for nid in new_ids:
    sl = S.get(presentationId=N, fields='slides(objectId,pageElements(shape(text(textElements(textRun(content))))))').execute()['slides']
    order = [x['objectId'] for x in sl]
    last = max(i for i, x in enumerate(sl) if any('GCX SIREN (26년' in T(e) for e in x.get('pageElements', [])))
    prev = [i for i, o in enumerate(order) if o in new_ids[:new_ids.index(nid)]]
    at = (max(prev) if prev else last) + 1
    S.batchUpdate(presentationId=N, body={'requests': [{'updateSlidesPosition': {'slideObjectIds': [nid], 'insertionIndex': at}}]}).execute()
print('cards', len(new_ids), '— now run apple_type.py fix, align_icons.py, then render')
