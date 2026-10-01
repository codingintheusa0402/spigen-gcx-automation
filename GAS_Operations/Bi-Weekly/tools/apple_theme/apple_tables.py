import sys, json, time
sys.path.insert(0, '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools')
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_cards import svc, P, E, rgb, rtile, FONT, CANVAS, TILE, INK, GRAY, GRAY2, BLUE, txt
HAIR = '#E5E5EA'
p = svc.get(presentationId=P).execute()
reqs = []
for s in p['slides']:
    sid = s['objectId']
    tables = [e for e in s['pageElements'] if 'table' in e]
    if not tables: continue
    reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in s['pageElements'] if e['objectId'].startswith('at_tb')]
    reqs.append({'updatePageProperties': {'objectId': sid, 'pageProperties': {'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(CANVAS)}}}}, 'fields': 'pageBackgroundFill.solidFill.color'}})
    for t in tables:
        tid, tb = t['objectId'], t['table']
        tr = t['transform']; x, y = tr['translateX']/E, tr['translateY']/E
        w = sum(c['columnWidth']['magnitude'] for c in tb['tableColumns'])/E
        h = sum(r['rowHeight']['magnitude'] for r in tb['tableRows'])/E
        # borders: none, then hairlines between rows
        reqs.append({'updateTableBorderProperties': {'objectId': tid, 'borderPosition': 'ALL', 'tableBorderProperties': {
            'tableBorderFill': {'solidFill': {'color': {'rgbColor': rgb(TILE)}}}, 'weight': {'magnitude': 0.75, 'unit': 'PT'}}, 'fields': 'tableBorderFill.solidFill.color,weight'}})
        reqs.append({'updateTableBorderProperties': {'objectId': tid, 'borderPosition': 'INNER_HORIZONTAL', 'tableBorderProperties': {
            'tableBorderFill': {'solidFill': {'color': {'rgbColor': rgb(HAIR)}}}, 'weight': {'magnitude': 0.75, 'unit': 'PT'}}, 'fields': 'tableBorderFill.solidFill.color,weight'}})
        reqs.append({'updateTableCellProperties': {'objectId': tid, 'tableRange': {'location': {'rowIndex': 0, 'columnIndex': 0}, 'rowSpan': tb['rows'], 'columnSpan': tb['columns']},
            'tableCellProperties': {'tableCellBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(TILE)}}}, 'contentAlignment': 'MIDDLE'},
            'fields': 'tableCellBackgroundFill.solidFill.color,contentAlignment'}})
        for ci, cell in enumerate(tb['tableRows'][0]['tableCells']):
            if txt({'shape': {'text': cell.get('text', {})}}):
                reqs.append({'updateTextStyle': {'objectId': tid, 'cellLocation': {'rowIndex': 0, 'columnIndex': ci}, 'textRange': {'type': 'ALL'},
                    'style': {'fontFamily': FONT, 'bold': True, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(GRAY)}}}, 'fields': 'fontFamily,bold,foregroundColor'}})
        # SIREN: icon images -> blue "보기" links in column 1
        hdr = [txt({'shape': {'text': c.get('text', {})}}) for c in tb['tableRows'][0]['tableCells']]
        if 'SIREN 보기' in hdr:
            col = hdr.index('SIREN 보기')
            icons = sorted([e for e in s['pageElements'] if 'image' in e and abs(e['transform']['translateX']/E-198) < 2],
                           key=lambda e: e['transform']['translateY'])
            for ri, ic in enumerate(icons, start=1):
                if ri >= tb['rows']: break
                url = (ic['image'].get('imageProperties', {}).get('link') or {}).get('url')
                reqs.append({'deleteObject': {'objectId': ic['objectId']}})
                if not url: continue
                if not txt({'shape': {'text': tb['tableRows'][ri]['tableCells'][col].get('text', {})}}):
                    reqs.append({'insertText': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': col}, 'text': '보기', 'insertionIndex': 0}})
                reqs.append({'updateTextStyle': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': col}, 'textRange': {'type': 'ALL'},
                    'style': {'fontFamily': FONT, 'link': {'url': url}, 'underline': False, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(BLUE)}}}, 'fields': 'fontFamily,link,underline,foregroundColor'}})
                reqs.append({'updateParagraphStyle': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': col}, 'textRange': {'type': 'ALL'},
                    'style': {'alignment': 'CENTER'}, 'fields': 'alignment'}})
        # white rounded tile behind the table
        r, ids = rtile(sid, f'at_tb_{tid[-8:]}', x-12, y-8, w+24, h+30)
        reqs += r
        reqs.append({'updatePageElementsZOrder': {'pageElementObjectIds': ids, 'operation': 'SEND_TO_BACK'}})
    # SIREN count badge -> Apple stat tile
    for e in s['pageElements']:
        if e['objectId'].startswith('apl_badge_'):
            reqs.append({'updateShapeProperties': {'objectId': e['objectId'], 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(TILE)}}}}, 'fields': 'shapeBackgroundFill.solidFill.color'}})
        if txt(e) == '2026 SIREN Registered by GCX':
            reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'foregroundColor': {'opaqueColor': {'rgbColor': rgb(GRAY)}}}, 'fields': 'foregroundColor'}})
        if txt(e) == '40':
            reqs.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'ALL'}, 'style': {'bold': True, 'fontSize': {'magnitude': 22, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(INK)}}}, 'fields': 'bold,fontSize,foregroundColor'}})
    print('table slide', sid)
for i in range(0, len(reqs), 300):
    for t in range(8):
        try: svc.batchUpdate(presentationId=P, body={'requests': reqs[i:i+300]}).execute(); break
        except Exception as ex:
            if '429' in str(ex): time.sleep(20); continue
            raise
print('requests', len(reqs))
