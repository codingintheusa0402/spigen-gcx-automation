import sys, os, json, re, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from rate_pipeline import creds
from googleapiclient.discovery import build
s = build('slides', 'v1', credentials=creds()).presentations()
from cfg import CFG
N = CFG['apple_deck']; E = 12700
SP = os.path.dirname(os.path.abspath(__file__)) + '/'
SRC = json.load(open(SP + 'siren_src.json'))
TPL = 'SLIDES_API1905798936_9485'
TESLA = 'https://docs.google.com/presentation/d/1KSykjC0-9eEKYQWY3Vi1gY8LS4VJrEiUAMVlYrgSWJ8/edit'
FOLD = 'https://docs.google.com/presentation/d/18URXgcRQoo0j4F_V3rCWVhmqOZ62CVfMqlOSWNJfDYA/edit'
TESLA_NAME = '전면 카메라 보호 커버 (TO331) ㅣTesla Model Y'
FOLD_NAME = 'Air Skin Aramid MagFit ㅣGlx. Z Fold 8'
CARDS = [  # (source slide, type, sku, asin, reason, product, rating, siren deck)
    ('g3bbe1be880a_0_24', 'claim', 'ACS11561', 'B0GMSKSKX8', '카메라링 변형', FOLD_NAME, '4.5점  Global Ratings: 427', FOLD),
    ('g2e14a09965c_0_19', 'claim', 'ACS11561', 'B0GMSKSKX8', '카메라링 변형', FOLD_NAME, '4.5점  Global Ratings: 427', FOLD),
    ('siren_slide_0035', 'claim', 'ACP12263', 'B0HC1H9BXW', '결로', TESLA_NAME, '2.5점 (DE)', TESLA),
    ('siren_slide_0040', 'claim', 'ACP12263', 'B0HC1H9BXW', '결로', TESLA_NAME, '2.5점 (DE)', TESLA),
    ('siren_slide_0056', 'review', 'ACP11231', 'B0G3YGP6QM', '결로', TESLA_NAME, '2.5점 (DE)', TESLA),
]
p = s.get(presentationId=N).execute()
ids = [x['objectId'] for x in p['slides']]
T = lambda e: ''.join(x.get('textRun', {}).get('content', '') for x in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()
siren_ix = max(i for i, sl in enumerate(p['slides']) if any('GCX SIREN (26년' in T(e) for e in sl['pageElements']))
tpl = next(x for x in p['slides'] if x['objectId'] == TPL)
role = {}  # template element text -> id
for e in tpl['pageElements']:
    role[T(e) or e['objectId']] = e['objectId']
# Amazon source icon from any review card (for the review card)
amz_icon = None
for sl in p['slides']:
    els = sl['pageElements']
    if any(T(e) == 'From Amazon' for e in els):
        ic = [e for e in els if e.get('title') == 'apl_icon' and round(e['transform'].get('translateY', 0) / E) == 84]
        if ic: amz_icon = ic[0]['image']['contentUrl']; break

def settext(oid, new):
    old = next(T(e) for e in tpl['pageElements'] if e['objectId'] == oid)
    return [{'insertText': {'objectId': oid, 'insertionIndex': 0, 'text': new}},
            {'deleteText': {'objectId': oid, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': len(new), 'endIndex': len(new) + len(old)}}}]

reqs = []
for n, (sid, kind, sku, asin, reason, prod, rating, deck) in enumerate(CARDS):
    c = SRC[sid]['cells']; v = dict(zip(c[:6], c[6:12]))
    country = c[6]; date = (c[9] if '인입' in c[3] else c[9]).replace('-', '.')
    ref = c[10]; detail = re.sub(r'^\[[A-Z]{3}\d{5}\]\s*', '', c[12]).replace('\n', ' ').strip()
    new = f'siren_card_{n}'
    m = {o: f'{new}_{o[-5:]}' for o in [e['objectId'] for e in tpl['pageElements']]}
    reqs.append({'duplicateObject': {'objectId': TPL, 'objectIds': {TPL: new, **m}}})
    reqs.append({'updateSlidesPosition': {'slideObjectIds': [new], 'insertionIndex': siren_ix + 1 + n}})
    R = lambda key: m[role[key]]
    def st(key, val):
        oid = role[key]; old = key
        return [{'insertText': {'objectId': m[oid], 'insertionIndex': 0, 'text': val}},
                {'deleteText': {'objectId': m[oid], 'textRange': {'type': 'FIXED_RANGE', 'startIndex': len(val), 'endIndex': len(val) + len(old)}}}]
    tpl_detail = next(T(e) for e in tpl['pageElements'] if e['objectId'] == 'SLIDES_API1905798936_9529')
    reqs += st('4. Galaxy Z8 Series Claims / Reviews', '5. GCX SIREN 신규 등록 이슈 사례')
    reqs += st(tpl_detail, detail)
    reqs += st('2026.10.06', date)
    reqs += st('FR', country)
    reqs += st('유격', reason)
    reqs += st('B0GMSKSKX8', asin)
    reqs += st('ACS11561', sku)
    reqs += st('4.5점  Global Ratings: 427', rating)
    reqs += st('Air Skin Aramid MagFit ㅣGlx. Z Fold 8', prod)
    link = f'https://spigenhelp.zendesk.com/agent/tickets/{ref}' if kind == 'claim' else next(iter(SRC[sid]['links']))
    btn = m[role['클레임 바로가기']]
    if kind == 'review':
        reqs += st('클레임 바로가기', '배드리뷰 바로가기') + st('클레임 내용', '리뷰 내용') + st('From Zendesk', 'From Amazon')
    reqs.append({'updateTextStyle': {'objectId': btn, 'textRange': {'type': 'ALL'}, 'style': {'link': {'url': link}}, 'fields': 'link'}})
    reqs.append({'updateTextStyle': {'objectId': m[role['SIREN 등록됨']], 'textRange': {'type': 'ALL'}, 'style': {'link': {'url': deck}}, 'fields': 'link'}})
    if kind == 'review' and amz_icon:
        reqs.append({'replaceImage': {'imageObjectId': m['SLIDES_API1905798936_9553'], 'url': amz_icon}})
    # photos: fit inside the photo tile area (50,124)-(426,372)
    reqs.append({'deleteObject': {'objectId': m['SLIDES_API1905798936_9538']}})
    imgs = SRC[sid]['imgs']; AX, AY, AW, AH, G = 50, 124, 376, 248, 9
    cw = (AW - G * (len(imgs) - 1)) / len(imgs)
    for k, im in enumerate(imgs):
        sc = min(cw / im['w'], AH / im['h']); w, h = im['w'] * sc, im['h'] * sc
        x = AX + k * (cw + G) + (cw - w) / 2; y = AY + (AH - h) / 2
        reqs.append({'createImage': {'objectId': f'{new}_ph{k}', 'url': im['url'], 'elementProperties': {'pageObjectId': new,
                     'size': {'width': {'magnitude': w * E, 'unit': 'EMU'}, 'height': {'magnitude': h * E, 'unit': 'EMU'}},
                     'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x * E, 'translateY': y * E, 'unit': 'EMU'}}}})
s.batchUpdate(presentationId=N, body={'requests': reqs}).execute()
print('cards', len(CARDS), 'requests', len(reqs), 'after slide', siren_ix + 1)
