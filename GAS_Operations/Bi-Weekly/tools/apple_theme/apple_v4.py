import sys
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_v3 import svc, P, E, rgb, txt, pill, text, INK, GRAY, FONT, CANVAS
from cfg import CFG
from apple_cards import rtile
import re as _re

LABEL = '2026 SIREN Registered by GCX'
_isnum = lambda t: _re.fullmatch(r'\d{1,4}', t or '') is not None

def _bbox(e):
    tr = e['transform']; sz = e.get('size', {})
    return (tr.get('translateX', 0)/E, tr.get('translateY', 0)/E,
            sz.get('width', {}).get('magnitude', 0)*tr.get('scaleX', 1)/E, sz.get('height', {}).get('magnitude', 0)*tr.get('scaleY', 1)/E)

def stat_tiles():
    """Classic yellow SIREN count badge -> white rounded stat tile 'at_st_*' (204x41, right-aligned to the table tile, 8pt below it)."""
    p = svc.get(presentationId=P).execute(); reqs = []
    for s in p['slides']:
        sid = s['objectId']; els = s['pageElements']
        lab = next((e for e in els if txt(e) == LABEL), None)
        if not lab or any(e['objectId'].startswith(('at_st_', 'apl_badge_')) for e in els): continue
        lx, ly, lw, lh = _bbox(lab); cx, cy = lx + lw/2, ly + lh/2
        for e in els:   # classic badge box / shadow: text-less shapes under the label
            if not ('shape' in e or 'image' in e) or txt(e) or e['objectId'].startswith(('at_', 'ap')): continue
            x, y, w, h = _bbox(e)
            if x <= cx <= x + w and y <= cy <= y + h: reqs.append({'deleteObject': {'objectId': e['objectId']}})
        tb = [_bbox(e) for e in els if e['objectId'].startswith('at_tb_')]
        right = max(x + w for x, y, w, h in tb) if tb else lx + lw + 40
        bottom = max(y + h for x, y, w, h in tb) if tb else ly - 8
        r, ids = rtile(sid, 'at_st_' + sid[-6:].replace('_', ''), right - 204, bottom + 8, 204, 41)
        reqs += r
        num = [e['objectId'] for e in els if _isnum(txt(e)) and abs(_bbox(e)[1] - ly) < 30]
        reqs.append({'updatePageElementsZOrder': {'pageElementObjectIds': [lab['objectId']] + num, 'operation': 'BRING_TO_FRONT'}})
    if reqs: svc.batchUpdate(presentationId=P, body={'requests': reqs}).execute()
    print('stat tiles', len(reqs))

def siren_and_end():
    p = svc.get(presentationId=P).execute()
    reqs = []
    for s in p['slides']:
        sid = s['objectId']
        badge = [_bbox(e) for e in s['pageElements'] if e['objectId'].startswith(('apl_badge_', 'at_st_'))]
        if badge:
            bx, by = min(b[0] for b in badge), min(b[1] for b in badge)
            bw = max(b[0] + b[2] for b in badge) - bx; bh = max(b[1] + b[3] for b in badge) - by
            reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in s['pageElements'] if e['objectId'].startswith(('ap3_sb', 'ap4_'))]
            # stat tile: label + number share one vertical center
            for e in s['pageElements']:
                t = txt(e)
                if t == LABEL or (_isnum(t) and bx - 20 <= _bbox(e)[0] + _bbox(e)[2]/2 <= bx + bw + 20 and by - 30 <= _bbox(e)[1] + _bbox(e)[3]/2 <= by + bh + 30):
                    sw, sh = e['size']['width']['magnitude'], e['size']['height']['magnitude']
                    x, w = (bx + 10, bw - 58) if t == LABEL else (bx + bw - 58, 46)
                    reqs.append({'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': {
                        'scaleX': w*E/sw, 'scaleY': bh*E/sh, 'translateX': x*E, 'translateY': by*E, 'unit': 'EMU'}}})
                    reqs.append({'updateShapeProperties': {'objectId': e['objectId'], 'shapeProperties': {'contentAlignment': 'MIDDLE', 'autofit': {'autofitType': 'NONE'}}, 'fields': 'contentAlignment,autofit.autofitType'}})
                    reqs.append({'updateParagraphStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'alignment': 'START' if t == LABEL else 'END', 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}}, 'fields': 'alignment,spaceAbove,spaceBelow'}})
                    reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'fontSize': {'magnitude': 9 if t == LABEL else 24, 'unit': 'PT'}}, 'fields': 'fontSize'}})
            reqs += pill(sid, 'ap4_sb' + sid[-6:].replace('_', ''), bx - 150, by + bh/2 - 13, 136, 26, '26년 SIREN 시트 열기', 'outline', 9,
                         url='https://docs.google.com/spreadsheets/d/15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI/edit#gid=1840076165', under=CANVAS)
            # table: column alignment (header follows its column), wider last column
            for e in s['pageElements']:
                if 'table' not in e: continue
                tb = e['table']; tid = e['objectId']
                hdr = [txt({'shape': {'text': c.get('text', {})}}) for c in tb['tableRows'][0]['tableCells']]
                if 'SIREN 보기' not in hdr: continue
                align = {'날짜': 'START', 'SIREN 보기': 'CENTER', '기종+제품': 'START', 'SKU': 'START', '이슈': 'START', 'SIREN 등록 여부': 'CENTER'}
                for ci, h in enumerate(hdr):
                    for ri in range(tb['rows']):
                        cell = tb['tableRows'][ri]['tableCells'][ci]
                        if txt({'shape': {'text': cell.get('text', {})}}):
                            reqs.append({'updateParagraphStyle': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': ci}, 'textRange': {'type': 'ALL'}, 'style': {'alignment': align.get(h, 'START')}, 'fields': 'alignment'}})
                widths = [c['columnWidth']['magnitude'] for c in tb['tableColumns']]
                if widths[5] < 70*E:
                    reqs.append({'updateTableColumnProperties': {'objectId': tid, 'columnIndices': [2], 'tableColumnProperties': {'columnWidth': {'magnitude': widths[2] - 22*E, 'unit': 'EMU'}}, 'fields': 'columnWidth'}})
                    reqs.append({'updateTableColumnProperties': {'objectId': tid, 'columnIndices': [5], 'tableColumnProperties': {'columnWidth': {'magnitude': widths[5] + 22*E, 'unit': 'EMU'}}, 'fields': 'columnWidth'}})
    # closing slide: kept as the classic black Spigen-logo slide (approved 261002 design)
    svc.batchUpdate(presentationId=P, body={'requests': reqs}).execute()
    print('siren/end requests', len(reqs))

if __name__ == '__main__':
    stat_tiles()
    siren_and_end()
