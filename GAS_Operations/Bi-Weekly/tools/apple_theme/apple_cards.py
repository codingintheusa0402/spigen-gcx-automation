import sys, json, time
sys.path.insert(0, '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools')
from rate_pipeline import creds
from googleapiclient.discovery import build
P = '1quCr9Xj-pSsVXKrYuaEOq0LPILZMN2LPUwBkY1f_GFI'
E = 12700
FONT = 'Noto Sans KR'
CANVAS, TILE, INK, GRAY, GRAY2, BLUE, RED = '#F5F5F7', '#FFFFFF', '#1D1D1F', '#6E6E73', '#86868B', '#0066CC', '#FF3B30'
svc = build('slides', 'v1', credentials=creds()).presentations()
ORIG = '12NxCxbW3z0fH1KKEVzX_uqBGH_APlkZpBWdPgET_aCk'
R = 12  # tile corner radius (pt)
def rtile(sid, oid, x, y, w, h, color=TILE):
    """Rounded tile with a controlled corner radius: 2 rects + 4 corner circles."""
    out, parts = [], [(x+R, y, w-2*R, h, 'RECTANGLE'), (x, y+R, w, h-2*R, 'RECTANGLE'),
                      (x, y, 2*R, 2*R, 'ELLIPSE'), (x+w-2*R, y, 2*R, 2*R, 'ELLIPSE'),
                      (x, y+h-2*R, 2*R, 2*R, 'ELLIPSE'), (x+w-2*R, y+h-2*R, 2*R, 2*R, 'ELLIPSE')]
    ids = []
    for i, (px, py, pw, ph, kind) in enumerate(parts):
        pid = f'{oid}_{i}'; ids.append(pid)
        out += [{'createShape': {'objectId': pid, 'shapeType': kind, 'elementProperties': {'pageObjectId': sid,
                  'size': {'width': {'magnitude': pw*E, 'unit': 'EMU'}, 'height': {'magnitude': ph*E, 'unit': 'EMU'}},
                  'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': px*E, 'translateY': py*E, 'unit': 'EMU'}}}},
                {'updateShapeProperties': {'objectId': pid, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(color)}}},
                  'outline': {'propertyState': 'NOT_RENDERED'}}, 'fields': 'shapeBackgroundFill.solidFill.color,outline.propertyState'}}]
    return out, ids

def rgb(h): h = h.lstrip('#'); return {'red': int(h[0:2], 16)/255, 'green': int(h[2:4], 16)/255, 'blue': int(h[4:6], 16)/255}
def txt(e): return ''.join(t.get('textRun', {}).get('content', '') for t in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()
def link_of(e): return ((e.get('image', {}).get('imageProperties', {}) or {}).get('link') or {}).get('url', '')

def place(e, x, y, w, h):
    """Absolute move/resize keeping the element's own size (scale = target/size)."""
    sw, sh = e['size']['width']['magnitude'], e['size']['height']['magnitude']
    return {'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE',
            'transform': {'scaleX': w*E/sw, 'scaleY': h*E/sh, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}
def tile(sid, oid, x, y, w, h, color=TILE):
    return [{'createShape': {'objectId': oid, 'shapeType': 'ROUND_RECTANGLE', 'elementProperties': {'pageObjectId': sid,
              'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': h*E, 'unit': 'EMU'}},
              'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}},
            {'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(color)}}},
              'outline': {'propertyState': 'NOT_RENDERED'}}, 'fields': 'shapeBackgroundFill.solidFill.color,outline.propertyState'}}]
def label(sid, oid, x, y, w, h, text, size, color, bold=False, url=None, align='START'):
    style = {'fontFamily': FONT, 'fontSize': {'magnitude': size, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}, 'bold': bold}
    fields = 'fontFamily,fontSize,foregroundColor,bold'
    if url: style['link'] = {'url': url}; style['underline'] = False; fields += ',link,underline'
    return [{'createShape': {'objectId': oid, 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': sid,
              'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': h*E, 'unit': 'EMU'}},
              'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}},
            {'insertText': {'objectId': oid, 'text': text}},
            {'updateTextStyle': {'objectId': oid, 'style': style, 'fields': fields, 'textRange': {'type': 'ALL'}}},
            {'updateParagraphStyle': {'objectId': oid, 'style': {'alignment': align, 'lineSpacing': 100, 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}}, 'fields': 'alignment,lineSpacing,spaceAbove,spaceBelow', 'textRange': {'type': 'ALL'}}},
            {'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'contentAlignment': 'TOP'}, 'fields': 'contentAlignment'}}]
def restyle(e, size, color, bold=False, align=None, valign='TOP'):
    if not txt(e): return []
    r = [{'updateTextStyle': {'objectId': e['objectId'], 'style': {'fontFamily': FONT, 'fontSize': {'magnitude': size, 'unit': 'PT'},
           'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}, 'bold': bold}, 'fields': 'fontFamily,fontSize,foregroundColor,bold', 'textRange': {'type': 'ALL'}}},
         {'updateShapeProperties': {'objectId': e['objectId'], 'shapeProperties': {'contentAlignment': valign, 'autofit': {'autofitType': 'NONE'},
           'shapeBackgroundFill': {'propertyState': 'NOT_RENDERED'}, 'outline': {'propertyState': 'NOT_RENDERED'}},
           'fields': 'contentAlignment,autofit.autofitType,shapeBackgroundFill.propertyState,outline.propertyState'}},
         {'updateParagraphStyle': {'objectId': e['objectId'], 'style': {'lineSpacing': 105, 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}, **({'alignment': align} if align else {})},
           'fields': 'lineSpacing,spaceAbove,spaceBelow' + (',alignment' if align else ''), 'textRange': {'type': 'ALL'}}}]
    return r

def near(e, x, y, tol=3):
    tr = e.get('transform', {}); return abs(tr.get('translateX', 0)/E - x) < tol and abs(tr.get('translateY', 0)/E - y) < tol

def card_requests(s, orig_slide):
    sid = s['objectId']; k = sid[-10:].replace('_', '')
    cur_ids = {e['objectId'] for e in s['pageElements']}
    els = [e for e in orig_slide['pageElements']]
    find = lambda x, y: next((e for e in els if 'shape' in e and near(e, x, y)), None)
    panel = next((e for e in els if 'image' in e and near(e, 90.4, 60.2) and e['size']['width']['magnitude']*e['transform'].get('scaleX', 1)/E > 590), None)
    if not panel: return None
    badge = next((e for e in els if 'image' in e and link_of(e) and e['transform'].get('translateX', 0)/E > 560 and e['transform'].get('translateY', 0)/E < 90), None)
    patch = [e for e in els if 'image' in e and near(e, 204, 92.6)]
    title, product, date = find(82, 17.4), find(90.4, 60.2), find(391.3, 86.3)
    heading, content = find(495.9, 77.1), find(495.3, 97.5)
    vals = [find(494.8, 170.5), find(494.8, 217.9), find(495.5, 261.7), find(495.5, 308.7), find(495.9, 355.0)]
    siren = next((e for e in els if 'shape' in e and txt(e) == 'SIREN 등록됨'), None)
    photos = [e for e in els if ('image' in e or 'video' in e) and e not in [panel, badge] + patch and 95 <= e['transform'].get('translateX', 0)/E <= 480 and 105 <= e['transform'].get('translateY', 0)/E <= 365]
    is_review = heading and '리뷰' in txt(heading)
    url = link_of(badge) if badge else ''
    reqs = [{'deleteObject': {'objectId': i}} for i in cur_ids if i.startswith('at_')]
    els = [e for e in els if e['objectId'] in cur_ids or e in (panel, badge) or e in patch]
    reqs += [{'updatePageProperties': {'objectId': sid, 'pageProperties': {'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(CANVAS)}}}}, 'fields': 'pageBackgroundFill.solidFill.color'}}]
    reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in [panel, badge] + patch if e and e['objectId'] in cur_ids]
    # tiles: photo tile left, story tile + spec grid right
    tiles = []
    grid = [(452, 198), (572, 198), (452, 258), (572, 258), (452, 318)]
    gw = [112, 112, 112, 112, 232]
    for oid, (tx, ty, tw, th) in [(f'at_ph_{k}', (36, 72, 404, 312)), (f'at_st_{k}', (452, 72, 232, 118))] + \
            [(f'at_g{i}_{k}', (gx, gy, gw[i], 52 if i < 4 else 66)) for i, (gx, gy) in enumerate(grid)]:
        r, ids = rtile(sid, oid, tx, ty, tw, th); reqs += r; tiles += ids
    # headline
    if title: reqs += [place(title, 36, 18, 640, 40)] + restyle(title, 22, INK, bold=True)
    # photo tile header: product + date
    if product: reqs += [place(product, 50, 82, 300, 20)] + restyle(product, 11, INK, bold=True)
    if date: reqs += [place(date, 50, 100, 120, 16)] + restyle(date, 9, GRAY2)
    if siren:
        from apple_v3 import pill
        sl = (siren['shape'].get('shapeProperties', {}).get('link') or {}).get('url') or next(
            (t['textRun']['style']['link'].get('url') for t in siren['shape'].get('text', {}).get('textElements', [])
             if t.get('textRun', {}).get('style', {}).get('link')), None)
        if siren['objectId'] in cur_ids: reqs.append({'deleteObject': {'objectId': siren['objectId']}})
        reqs += pill(sid, f'at_sir_{k}', 350, 81, 78, 18, 'SIREN 등록됨', 'red', 7.5, url=sl)
    # photos: map old photo area (95..480 x 110..360) into tile (50..426 x 124..372)
    photos = [e for e in photos if e['objectId'] in cur_ids]
    if photos:
        boxes = []
        for e in photos:
            tr = e['transform']; w = e['size']['width']['magnitude']*tr.get('scaleX', 1)/E; h = e['size']['height']['magnitude']*tr.get('scaleY', 1)/E
            boxes.append((e, tr['translateX']/E, tr['translateY']/E, w, h))
        bx0 = min(b[1] for b in boxes); by0 = min(b[2] for b in boxes)
        bx1 = max(b[1]+b[3] for b in boxes); by1 = max(b[2]+b[4] for b in boxes)
        AX, AY, AW, AH = 50, 124, 376, 248
        f = min(AW/(bx1-bx0), AH/(by1-by0))
        ox = AX + (AW - (bx1-bx0)*f)/2; oy = AY + (AH - (by1-by0)*f)/2
        for e, x, y, w, h in boxes:
            reqs.append(place(e, ox+(x-bx0)*f, oy+(y-by0)*f, w*f, h*f))
    # story tile
    reqs += label(sid, f'at_src_{k}', 464, 82, 120, 14, 'From Amazon' if is_review else 'From Zendesk', 8, GRAY2)
    if url:
        from apple_v3 import pill
        reqs += pill(sid, f'at_lnk_{k}', 584, 80, 88, 18, '배드리뷰 바로가기' if is_review else '클레임 바로가기', 'filled', 7.5, url=url)
    if heading: reqs += [place(heading, 464, 98, 210, 18)] + restyle(heading, 9, GRAY)
    if content: reqs += [place(content, 464, 116, 210, 70)] + restyle(content, 10.5, INK, bold=True)
    # spec grid
    names = ['국가', '인입사유', 'ASIN', 'SKU', '아마존 리뷰 평점 / 갯수']
    for i, ((gx, gy), v) in enumerate(zip(grid, vals)):
        reqs += label(sid, f'at_lb{i}_{k}', gx+12, gy+8, gw[i]-20, 12, names[i], 7.5, GRAY2)
        if v: reqs += [place(v, gx+12, gy+22, gw[i]-14, 26 if i < 4 else 36)] + restyle(v, 11.5, INK, bold=True)
    # tiles to back (behind texts/photos)
    reqs.append({'updatePageElementsZOrder': {'pageElementObjectIds': tiles, 'operation': 'SEND_TO_BACK'}})
    return reqs

if __name__ == '__main__':
    only = [int(a) for a in sys.argv[1:]]
    p = svc.get(presentationId=P).execute()
    orig = {x['objectId']: x for x in svc.get(presentationId=ORIG).execute()['slides']}
    pend = []
    for i, s in enumerate(p['slides']):
        if only and (i+1) not in only: continue
        if s['objectId'] not in orig: continue
        r = card_requests(s, orig[s['objectId']])
        if not r: continue
        pend += r
        if len(pend) > 350:
            for t in range(8):
                try: svc.batchUpdate(presentationId=P, body={'requests': pend}).execute(); break
                except Exception as ex:
                    if '429' in str(ex): time.sleep(20); continue
                    raise
            pend = []
        print('card', i+1, flush=True)
    if pend: svc.batchUpdate(presentationId=P, body={'requests': pend}).execute()
