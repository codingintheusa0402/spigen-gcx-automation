"""One-off repair of Slide-Maker cards already in the 260910 deck:
- photos: put each back into its template slot box and re-crop CENTER_CROP
- stale template photos (video thumbnail failed) removed
- long 클레임/리뷰 내용 text shrunk to fit its box
"""
import json, math, re, sys
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

PID = '1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc'
EMU = 12700
p = '/Users/kevinkim/.config/gws_shim/token.json'
d = json.load(open(p)); c = Credentials.from_authorized_user_info(d); c.refresh(Request()); d['token'] = c.token; json.dump(d, open(p, 'w'))
svc = build('slides', 'v1', credentials=c).presentations()
pres = svc.get(presentationId=PID).execute()
slides = pres['slides']

def geo(e):
    t = e.get('transform', {}); s = e.get('size', {})
    nw = s.get('width', {}).get('magnitude', 0) / EMU; nh = s.get('height', {}).get('magnitude', 0) / EMU
    sx = t.get('scaleX', 1); sy = t.get('scaleY', 1)
    return dict(L=t.get('translateX', 0) / EMU, T=t.get('translateY', 0) / EMU, W=nw * sx, H=nh * sy, nw=nw, nh=nh, sx=sx, sy=sy)

def text_of(e):
    return ''.join(te.get('textRun', {}).get('content', '') for te in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()

def in_photo_area(g):
    return 95 <= g['L'] <= 480 and 110 <= g['T'] <= 360 and 60 <= g['W'] <= 330 and 60 <= g['H'] <= 330

def card_info(s):
    title = ''; content = None; product = None; photos = []
    for e in s.get('pageElements', []):
        if 'shape' in e:
            t = text_of(e); g = geo(e)
            if 'Claims / Reviews' in t: title = t
            if abs(g['L'] - 495.3) < 1 and abs(g['T'] - 97.5) < 1: content = e
            if abs(g['L'] - 90.4) < 1 and abs(g['T'] - 60.2) < 1 and g['W'] > 200: product = e
        elif 'image' in e:
            g = geo(e)
            if in_photo_area(g): photos.append((e, g))
    return title, content, product, photos

def units(text):
    return sum(1 if re.match(r'[ᄀ-ᇿ㄰-㆏가-힯　-ヿ一-鿿＀-￯]', ch) else 0.55 for ch in text)

def fit_size(text, size, maxW, maxH, minSize, lineK=1.5):
    s = size
    while s > minSize:
        lines = sum(max(1, math.ceil(units(para) / max(1, maxW / s))) for para in text.split('\n'))
        if lines * s * lineK <= maxH: break
        s -= 0.5
    return max(s, minSize)

# family templates = last original card before the generated ones
families = {'Galaxy Z8': None, 'Pixel 11': None}
for i, s in enumerate(slides):
    if s['objectId'].startswith('SLIDES_API'): continue
    title, content, product, photos = card_info(s)
    for fam in families:
        if fam in title and len(photos) >= 1: families[fam] = (i, photos)
reqs = []; log = []
for i, s in enumerate(slides):
    if not s['objectId'].startswith('SLIDES_API'): continue
    title, content, product, photos = card_info(s)
    fam = next((f for f in families if f in title), None)
    if not fam: continue
    tidx, tslots = families[fam]
    slots = sorted([g for _, g in tslots], key=lambda g: g['L'])
    sigs = [(round(g['nw'], 2), round(g['nh'], 2), round(g['sx'], 1)) for g in slots]
    real = []
    for e, g in photos:
        sig = (round(g['nw'], 2), round(g['nh'], 2), round(g['sx'], 1))
        if not e['image'].get('sourceUrl') and sig in sigs:
            reqs.append({'deleteObject': {'objectId': e['objectId']}}); log.append('%d: removed stale template photo' % (i + 1))
        else:
            real.append((e, g))
    real.sort(key=lambda x: x[1]['L'])
    if len(real) == 1 and len(slots) > 1:
        big = max(slots, key=lambda g: g['W'] * g['H'])
        areaL = min(g['L'] for g in slots); areaR = max(g['L'] + g['W'] for g in slots)
        targets = [dict(L=areaL + (areaR - areaL - big['W']) / 2, T=big['T'], W=big['W'], H=big['H'])]
    else:
        targets = slots
    for (e, g), tg in zip(real, targets):
        reqs.append({'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': {
            'scaleX': tg['W'] / g['nw'], 'scaleY': tg['H'] / g['nh'], 'shearX': 0, 'shearY': 0,
            'translateX': tg['L'] * EMU, 'translateY': tg['T'] * EMU, 'unit': 'EMU'}}})
        url = e['image'].get('contentUrl')
        if url:
            reqs.append({'replaceImage': {'imageObjectId': e['objectId'], 'url': url, 'imageReplaceMethod': 'CENTER_CROP'}})
    if content:
        txt = text_of(content)
        cur = None
        for te in content['shape']['text']['textElements']:
            if 'textRun' in te: cur = te['textRun'].get('style', {}).get('fontSize', {}).get('magnitude'); break
        if txt and cur:
            new = fit_size(txt, cur, 170, 60, 6)
            if new < cur:
                reqs.append({'updateTextStyle': {'objectId': content['objectId'], 'style': {'fontSize': {'magnitude': new, 'unit': 'PT'}}, 'fields': 'fontSize'}})
                log.append('%d: content %spt -> %spt' % (i + 1, cur, new))
    if product:
        txt = text_of(product); cur = None
        for te in product['shape']['text']['textElements']:
            if 'textRun' in te: cur = te['textRun'].get('style', {}).get('fontSize', {}).get('magnitude'); break
        if txt and cur:
            new = fit_size(txt, cur, 240, cur * 1.5, 8, 1.4)
            if new < cur:
                reqs.append({'updateTextStyle': {'objectId': product['objectId'], 'style': {'fontSize': {'magnitude': new, 'unit': 'PT'}}, 'fields': 'fontSize'}})
                log.append('%d: product %spt -> %spt' % (i + 1, cur, new))
print('\n'.join(log)); print(len(reqs), 'requests')
if '--apply' in sys.argv and reqs:
    # replaceImage per element must follow its transform; batch is ordered
    for k in range(0, len(reqs), 60):
        svc.batchUpdate(presentationId=PID, body={'requests': reqs[k:k + 60]}).execute()
    print('applied')
