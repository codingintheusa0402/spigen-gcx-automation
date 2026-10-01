import sys
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_v3 import svc, P, E, rgb, txt, pill, text, INK, GRAY, FONT, CANVAS
from cfg import CFG

def siren_and_end():
    p = svc.get(presentationId=P).execute()
    reqs = []
    for s in p['slides']:
        sid = s['objectId']
        badge = [e for e in s['pageElements'] if e['objectId'].startswith('apl_badge_')]
        if badge:
            b = badge[0]; tr = b['transform']
            bx, by = tr['translateX']/E, tr['translateY']/E
            bw = b['size']['width']['magnitude']*tr.get('scaleX', 1)/E; bh = b['size']['height']['magnitude']*tr.get('scaleY', 1)/E
            reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in s['pageElements'] if e['objectId'].startswith(('ap3_sb', 'ap4_'))]
            # stat tile: label + number share one vertical center
            for e in s['pageElements']:
                t = txt(e)
                if t in ('2026 SIREN Registered by GCX', '40'):
                    sw, sh = e['size']['width']['magnitude'], e['size']['height']['magnitude']
                    x, w = (bx + 14, bw - 70) if t != '40' else (bx + bw - 58, 46)
                    reqs.append({'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': {
                        'scaleX': w*E/sw, 'scaleY': bh*E/sh, 'translateX': x*E, 'translateY': by*E, 'unit': 'EMU'}}})
                    reqs.append({'updateShapeProperties': {'objectId': e['objectId'], 'shapeProperties': {'contentAlignment': 'MIDDLE', 'autofit': {'autofitType': 'NONE'}}, 'fields': 'contentAlignment,autofit.autofitType'}})
                    reqs.append({'updateParagraphStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'alignment': 'START' if t != '40' else 'END', 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}}, 'fields': 'alignment,spaceAbove,spaceBelow'}})
                    reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'fontSize': {'magnitude': 9 if t != '40' else 24, 'unit': 'PT'}}, 'fields': 'fontSize'}})
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
    siren_and_end()
