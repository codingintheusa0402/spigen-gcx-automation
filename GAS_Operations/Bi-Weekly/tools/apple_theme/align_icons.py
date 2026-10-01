"""Vertically aligns every apl_icon image with the text it labels (run after icons are placed or fonts change).
Each label box is grown to fit one line of its text (center kept), text is middle-aligned, and the icon is
centered on the same line. Pairs icon -> nearest text box starting just right of the icon on the same row."""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
from apple_cards import svc, P, E, txt

def geo(e):
    tr = e['transform']; return (tr.get('translateX', 0)/E, tr.get('translateY', 0)/E,
                                 e['size']['width']['magnitude']*tr.get('scaleX', 1)/E, e['size']['height']['magnitude']*tr.get('scaleY', 1)/E)
def fsize(e):
    return next((t['textRun']['style']['fontSize']['magnitude'] for t in e['shape']['text']['textElements']
                 if t.get('textRun', {}).get('style', {}).get('fontSize')), 10)

def main():
    p = svc.get(presentationId=P).execute(); reqs = []
    for s in p['slides']:
        els = s['pageElements']
        texts = [e for e in els if 'shape' in e and txt(e)]
        for ic in (e for e in els if 'image' in e and e.get('title') == 'apl_icon'):
            ix, iy, iw, ih = geo(ic); icy = iy + ih/2
            cands = [(abs(geo(t)[1] + geo(t)[3]/2 - icy) + abs(geo(t)[0] - ix), t) for t in texts
                     if ix - 12 <= geo(t)[0] <= ix + iw + 8 and geo(t)[1] - 14 <= icy <= geo(t)[1] + geo(t)[3] + 14]
            if not cands: continue
            t = min(cands, key=lambda c: c[0])[1]; tx, ty, tw, th = geo(t)
            nh = max(th, fsize(t)*1.45 + 8); cy = ty + th/2
            reqs += [{'updatePageElementTransform': {'objectId': t['objectId'], 'applyMode': 'ABSOLUTE', 'transform': {
                        'scaleX': t['transform'].get('scaleX', 1), 'scaleY': nh*E/t['size']['height']['magnitude'],
                        'translateX': tx*E, 'translateY': (cy - nh/2)*E, 'unit': 'EMU'}}},
                     {'updateShapeProperties': {'objectId': t['objectId'], 'shapeProperties': {'contentAlignment': 'MIDDLE', 'autofit': {'autofitType': 'NONE'}}, 'fields': 'contentAlignment,autofit.autofitType'}},
                     {'updatePageElementTransform': {'objectId': ic['objectId'], 'applyMode': 'ABSOLUTE', 'transform': dict(ic['transform'], translateY=(cy - ih/2)*E, unit='EMU')}}]
    for i in range(0, len(reqs), 400):
        for _ in range(8):
            try: svc.batchUpdate(presentationId=P, body={'requests': reqs[i:i+400]}).execute(); break
            except Exception as ex:
                if '429' in str(ex): time.sleep(20); continue
                raise
    print('aligned', len(reqs)//3)

if __name__ == '__main__':
    main()
