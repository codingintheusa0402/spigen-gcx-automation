import sys, re, time
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_v3 import svc, P, E, rgb, txt, INK
GRAY2 = '#86868B'
LATIN, HANGUL = 'Inter', 'Noto Sans KR'
TITLE_RE = re.compile(r'^(\d\.\s|Overview|Appendix)')

def send(reqs):
    for i in range(0, len(reqs), 400):
        for t in range(8):
            try: svc.batchUpdate(presentationId=P, body={'requests': reqs[i:i+400]}).execute(); break
            except Exception as ex:
                if '429' in str(ex): time.sleep(20); continue
                raise

def titles():
    p = svc.get(presentationId=P).execute(); reqs = []
    for si, s in enumerate(p['slides']):
        if si in (0, 1) or si == len(p['slides']) - 1: continue      # cover / index / closing have their own hero type
        for e in s['pageElements']:
            if 'shape' not in e or e['transform'].get('translateY', 0)/E > 30: continue
            t = txt(e)
            if not TITLE_RE.match(t): continue
            oid = e['objectId']; raw = ''.join(x.get('textRun', {}).get('content', '') for x in e['shape']['text']['textElements'])
            # "1.\tOverview" -> "1. Overview"
            if re.match(r'^\d\.[^\s]', raw):
                reqs.append({'insertText': {'objectId': oid, 'insertionIndex': 2, 'text': ' '}}); t = t[:2] + ' ' + t[2:]
            tab = raw.find('\t')
            if 0 <= tab < 4:
                reqs += [{'deleteText': {'objectId': oid, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': tab, 'endIndex': tab + 1}}},
                         {'insertText': {'objectId': oid, 'insertionIndex': tab, 'text': ' '}}]
            sw, sh = e['size']['width']['magnitude'], e['size']['height']['magnitude']
            reqs.append({'updatePageElementTransform': {'objectId': oid, 'applyMode': 'ABSOLUTE', 'transform': {
                'scaleX': 648*E/sw, 'scaleY': 34*E/sh, 'translateX': 36*E, 'translateY': 16*E, 'unit': 'EMU'}}})
            reqs.append({'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'contentAlignment': 'MIDDLE', 'autofit': {'autofitType': 'NONE'}}, 'fields': 'contentAlignment,autofit.autofitType'}})
            reqs.append({'updateParagraphStyle': {'objectId': oid, 'textRange': {'type': 'ALL'}, 'style': {'alignment': 'START', 'lineSpacing': 100, 'indentStart': {'magnitude': 0, 'unit': 'PT'}, 'indentFirstLine': {'magnitude': 0, 'unit': 'PT'}, 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}},
                         'fields': 'alignment,lineSpacing,indentStart,indentFirstLine,spaceAbove,spaceBelow'}})
            reqs.append({'updateTextStyle': {'objectId': oid, 'textRange': {'type': 'ALL'}, 'style': {'fontSize': {'magnitude': 18, 'unit': 'PT'}, 'bold': True,
                         'foregroundColor': {'opaqueColor': {'rgbColor': rgb(INK)}}}, 'fields': 'fontSize,bold,foregroundColor'}})
            q = t.find(' 대상 국가')
            if q > 0:
                reqs.append({'updateTextStyle': {'objectId': oid, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': q + 1, 'endIndex': len(t)},
                             'style': {'foregroundColor': {'opaqueColor': {'rgbColor': rgb(GRAY2)}}}, 'fields': 'foregroundColor'}})
    send(reqs); print('title requests', len(reqs))

def is_hangul(ch): return '가' <= ch <= '힣' or '㄰' <= ch <= '㆏'
def latin_ranges(text_obj):
    """[(start,end)] of runs that contain no Hangul (whitespace/punct sticks to its neighbour)."""
    out = []
    for te in text_obj.get('textElements', []):
        tr = te.get('textRun')
        if not tr: continue
        s0 = te.get('startIndex', 0); c = tr['content']; i = 0; bold = bool(tr.get('style', {}).get('bold'))
        while i < len(c):
            if is_hangul(c[i]) or not c[i].strip(): i += 1; continue
            j = i
            while j < len(c) and not is_hangul(c[j]): j += 1
            seg = c[i:j].rstrip()
            if seg.strip(): out.append((s0 + i, s0 + i + len(seg), bold))
            i = j
    return out

def fonts():
    p = svc.get(presentationId=P).execute(); reqs = []
    def walk(els):
        for e in els:
            if 'elementGroup' in e: walk(e['elementGroup']['children']); continue
            if 'shape' in e and e['shape'].get('text'):
                for a, b, bd in latin_ranges(e['shape']['text']):
                    reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'FIXED_RANGE', 'startIndex': a, 'endIndex': b}, 'style': {'weightedFontFamily': {'fontFamily': LATIN, 'weight': 600 if bd else 400}}, 'fields': 'weightedFontFamily'}})
            if 'table' in e:
                for ri, row in enumerate(e['table']['tableRows']):
                    for ci, cell in enumerate(row['tableCells']):
                        for a, b, bd in latin_ranges(cell.get('text', {})):
                            reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'cellLocation': {'rowIndex': ri, 'columnIndex': ci}, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': a, 'endIndex': b}, 'style': {'weightedFontFamily': {'fontFamily': LATIN, 'weight': 600 if bd else 400}}, 'fields': 'weightedFontFamily'}})
    for s in p['slides'][1:-1]: walk(s['pageElements'])     # cover + closing keep their brand fonts
    send(reqs); print('font requests', len(reqs))

if __name__ == '__main__':
    if 'titles' in sys.argv: titles()
    if 'fonts' in sys.argv: fonts()

def bullets_and_bold():
    """'1.' list bullets -> literal '1. '; restore bold on elements that are bold by design."""
    from apple_cards import near, ORIG
    p = svc.get(presentationId=P).execute()
    orig = {x['objectId']: x for x in svc.get(presentationId=ORIG).execute()['slides']}
    reqs = []
    def bold(oid, cell=None):
        r = {'updateTextStyle': {'objectId': oid, 'textRange': {'type': 'ALL'}, 'style': {'bold': True}, 'fields': 'bold'}}
        if cell: r['updateTextStyle']['cellLocation'] = cell
        reqs.append(r)
    for si, s in enumerate(p['slides']):
        ids = {e['objectId']: e for e in s['pageElements']}
        for e in s['pageElements']:
            if 'shape' not in e or not e['shape'].get('text'): continue
            t = txt(e); oid = e['objectId']
            has_bullet = any(te.get('paragraphMarker', {}).get('bullet') for te in e['shape']['text']['textElements'])
            if has_bullet and e['transform'].get('translateY', 0)/E < 30 and t.startswith('Overview'):
                reqs += [{'deleteParagraphBullets': {'objectId': oid, 'textRange': {'type': 'ALL'}}},
                         {'insertText': {'objectId': oid, 'insertionIndex': 0, 'text': '1. '}}]
            if t.endswith('건') or t in ('모델별 TOP3', '인입사유별 TOP3') or oid.startswith(('ap3_a_c', 'ap3_a_title', 'ap4_end_t')) and not oid.startswith('ap3_a_c0_') \
               :
                if not re.match(r'^ap3_a_c\d_\d', oid): bold(oid)
        for e in s['pageElements']:   # table header rows
            if 'table' in e:
                for ci, c in enumerate(e['table']['tableRows'][0]['tableCells']):
                    if txt({'shape': {'text': c.get('text', {})}}): bold(e['objectId'], {'rowIndex': 0, 'columnIndex': ci})
        o = orig.get(s['objectId'])        # card roles from the original deck (same ids)
        if o:
            for x, y in [(90.4, 60.2), (495.3, 97.5), (494.8, 170.5), (494.8, 217.9), (495.5, 261.7), (495.5, 308.7), (495.9, 355.0)]:
                el = next((q for q in o['pageElements'] if 'shape' in q and near(q, x, y)), None)
                if el and el['objectId'] in ids and txt(ids[el['objectId']]): bold(el['objectId'])
    send(reqs); print('bullet/bold requests', len(reqs))

if __name__ == '__main__' and 'fix' in sys.argv:
    bullets_and_bold(); titles(); fonts()
