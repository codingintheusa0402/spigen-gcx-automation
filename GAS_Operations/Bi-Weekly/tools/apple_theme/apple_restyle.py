import sys, json
sys.path.insert(0, '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools')
from rate_pipeline import creds
from googleapiclient.discovery import build

sys.path.insert(0, __import__('os').path.dirname(__file__))
from cfg import CFG
P = CFG['apple_deck']
E = 12700
FONT = 'Noto Sans KR'
INK, SECOND, HAIR, TILE, WHITE = '#1D1D1F', '#6E6E73', '#D2D2D7', '#F5F5F7', '#FFFFFF'
D_SECOND, BLUE, D_BLUE, RED, BLACK = '#A1A1A6', '#0071E3', '#2997FF', '#FF3B30', '#000000'

svc = build('slides', 'v1', credentials=creds()).presentations()
p = svc.get(presentationId=P).execute()

def rgb(h): h = h.lstrip('#'); return {'red': int(h[0:2], 16)/255, 'green': int(h[2:4], 16)/255, 'blue': int(h[4:6], 16)/255}
def hx(c):
    if not c: return None
    oc = c.get('opaqueColor', c)
    if 'rgbColor' in oc:
        r = oc['rgbColor']; return '#%02X%02X%02X' % tuple(round(r.get(k, 0)*255) for k in ('red', 'green', 'blue'))
    return {'LIGHT1': '#FFFFFF', 'DARK1': '#000000', 'DARK2': '#595959', 'ACCENT1': '#FFAB40'}.get(oc.get('themeColor'), oc.get('themeColor'))
def bbox(e):
    if 'elementGroup' in e:
        bs = [bbox(ch) for ch in e['elementGroup']['children']]
        return [min(x[0] for x in bs), min(x[1] for x in bs), max(x[2] for x in bs), max(x[3] for x in bs)]
    tr = e.get('transform', {}); a, b, c, d = tr.get('scaleX', 1), tr.get('shearX', 0), tr.get('shearY', 0), tr.get('scaleY', 1)
    tx, ty = tr.get('translateX', 0), tr.get('translateY', 0)
    w = e.get('size', {}).get('width', {}).get('magnitude', 0); h = e.get('size', {}).get('height', {}).get('magnitude', 0)
    pts = [(a*x+b*y+tx, c*x+d*y+ty) for x, y in ((0, 0), (w, 0), (0, h), (w, h))]
    return [min(q[0] for q in pts)/E, min(q[1] for q in pts)/E, max(q[0] for q in pts)/E, max(q[1] for q in pts)/E]
def flat(els):
    for e in els:
        if 'elementGroup' in e: yield from flat(e['elementGroup']['children'])
        else: yield e
def inside(bb, region):
    cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
    return region[0] <= cx <= region[2] and region[1] <= cy <= region[3]

GRAYS = {'#7F89AC', '#7278B2', '#7680A2', '#8F95B3', '#9AA0C0', '#C9CDD8', '#595959'}
DARKS = {'#0A1739', '#121735', '#303665', '#000000', '#473821'}
REDS = {'#FF5252', '#FF6B6B', '#DD7E6B'}
def map_color(c, on_dark):
    if c is None: return None if on_dark else INK
    if c in ('#FFFFFF',): return None if on_dark else INK
    if c in GRAYS: return D_SECOND if on_dark else SECOND
    if c in DARKS: return TILE if on_dark else INK
    if c in REDS: return RED
    if c == '#FF5A00': return D_BLUE if on_dark else BLUE
    return None

def text_reqs(oid, text, on_dark, cell=None, force=None, title=False):
    reqs = []
    for t in text.get('textElements', []):
        tr = t.get('textRun')
        if not tr or not tr.get('content', '').strip('\n'): continue
        st = tr.get('style', {})
        style, fields = {'fontFamily': FONT}, ['fontFamily']
        col = force or map_color(hx(st.get('foregroundColor')), on_dark)
        if col: style['foregroundColor'] = {'opaqueColor': {'rgbColor': rgb(col)}}; fields.append('foregroundColor')
        if title and st.get('fontSize', {}).get('magnitude', 0) >= 18: style['bold'] = True; fields.append('bold')
        r = {'updateTextStyle': {'objectId': oid, 'style': style, 'fields': ','.join(fields),
             'textRange': {'type': 'FIXED_RANGE', 'startIndex': t.get('startIndex', 0), 'endIndex': t['endIndex']}}}
        if cell: r['updateTextStyle']['cellLocation'] = cell
        reqs.append(r)
    return reqs

def bg(sid, h): return {'updatePageProperties': {'objectId': sid, 'pageProperties': {'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(h)}}}}, 'fields': 'pageBackgroundFill.solidFill.color'}}
def fill(oid, h): return {'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(h)}}}}, 'fields': 'shapeBackgroundFill.solidFill.color'}}

n_slides = len(p['slides'])
total = 0
import time
START = int(sys.argv[1]) if len(sys.argv) > 1 else 0
pending = []
def send(batch):
    for attempt in range(8):
        try:
            svc.batchUpdate(presentationId=P, body={'requests': batch}).execute(); return
        except Exception as ex:
            if '429' in str(ex): time.sleep(20); continue
            raise
for si, s in enumerate(p['slides']):
    if si < START: continue
    sid = s['objectId']; reqs = []
    is_cover, is_end = si == 0, si == n_slides - 1
    if is_cover or is_end: continue      # final design keeps the orange Spigen cover + black logo closing as-is
    reqs.append(bg(sid, BLACK if (is_cover or is_end) else WHITE))
    top = s.get('pageElements', [])
    # 1) sidebar chrome + hard drop shadows
    for e in top:
        bb = bbox(e)
        if 'elementGroup' in e and bb[0] < 1 and bb[2] > 300 and bb[3] > 400: reqs.append({'deleteObject': {'objectId': e['objectId']}}); continue
        if 'image' in e and bb[0] < 1 and bb[2] <= 53: reqs.append({'deleteObject': {'objectId': e['objectId']}}); continue
        f = e.get('shape', {}).get('shapeProperties', {}).get('shapeBackgroundFill', {})
        if 'solidFill' in f and f.get('propertyState') != 'NOT_RENDERED' and hx(f['solidFill']['color']) == '#000000' and not e['shape'].get('text'):
            reqs.append({'deleteObject': {'objectId': e['objectId']}})  # SIREN table shadow
    deleted = {r['deleteObject']['objectId'] for r in reqs if 'deleteObject' in r}
    els = [e for e in flat([x for x in top if x['objectId'] not in deleted])]
    # 2) dark regions: chart gauges, claim/review card panels, dark-filled tiles
    dark = []
    for e in els:
        bb = bbox(e)
        if 'image' in e and (str(e.get('title', '')).startswith('GEN_') or (bb[2]-bb[0] > 590 and bb[3]-bb[1] > 300 and bb[0] >= 85)):
            dark.append(bb)
        f = e.get('shape', {}).get('shapeProperties', {}).get('shapeBackgroundFill', {})
        if 'solidFill' in f and f.get('propertyState') != 'NOT_RENDERED':
            c = hx(f['solidFill']['color'])
            if c in ('#303665', INK): reqs.append(fill(e['objectId'], INK)); dark.append(bb)
            if c == '#4A1426': reqs.append(fill(e['objectId'], RED))
    # 3) text + tables
    for e in els:
        bb = bbox(e)
        if 'shape' in e and e['shape'].get('text'):
            f = e['shape'].get('shapeProperties', {}).get('shapeBackgroundFill', {})
            chip = 'solidFill' in f and hx(f['solidFill']['color']) == '#4A1426'
            on_dark = is_cover or is_end or any(inside(bb, r) for r in dark)
            title = abs(bb[0]-82) < 2 and abs(bb[1]-17) < 2
            force = WHITE if chip else (('#F5F5F7' if title else '#86868B') if is_cover else (INK if title else None))
            if is_cover and 'Bi-weekly' in ''.join(t.get('textRun', {}).get('content', '') for t in e['shape']['text']['textElements']): force, title = '#F5F5F7', True
            reqs += text_reqs(e['objectId'], e['shape']['text'], on_dark, force=force, title=title)
        if 'table' in e:
            t = e['table']
            reqs.append({'updateTableBorderProperties': {'objectId': e['objectId'], 'borderPosition': 'ALL',
                'tableBorderProperties': {'tableBorderFill': {'solidFill': {'color': {'rgbColor': rgb(HAIR)}}}, 'weight': {'magnitude': 0.75, 'unit': 'PT'}, 'dashStyle': 'SOLID'},
                'fields': 'tableBorderFill.solidFill.color,weight,dashStyle'}})
            for ri, row in enumerate(t['tableRows']):
                reqs.append({'updateTableCellProperties': {'objectId': e['objectId'],
                    'tableRange': {'location': {'rowIndex': ri, 'columnIndex': 0}, 'rowSpan': 1, 'columnSpan': t['columns']},
                    'tableCellProperties': {'tableCellBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(TILE if ri == 0 else WHITE)}}}},
                    'fields': 'tableCellBackgroundFill.solidFill.color'}})
                for ci, cell in enumerate(row['tableCells']):
                    if cell.get('text'):
                        reqs += text_reqs(e['objectId'], cell['text'], False, cell={'rowIndex': ri, 'columnIndex': ci},
                                          force=(INK if ri == 0 else None))
    pending += reqs
    if len(pending) > 300: send(pending); pending = []
    total += len(reqs)
    print(si+1, len(reqs), flush=True)
if pending: send(pending)
print('total requests', total)
