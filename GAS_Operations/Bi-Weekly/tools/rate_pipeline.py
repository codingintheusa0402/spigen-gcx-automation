#!/usr/bin/env python3
"""Bi-Weekly builder — 클레임 + 배드리뷰 / 판매량 TOP N slides.

Subcommands (run from anywhere; state lives in tools/state/):

  prep      --series glxZ8,pixel11 --start YYYY-MM-DD --end YYYY-MM-DD
            Reads each series' 신제품 라인업 catalogue, counts Zendesk
            Product-Issue claims and 1-3점 bad reviews (EU, date-windowed),
            collects the top-2 인입사유 per SKU, and prints the Caspi LM SQL
            for the sales query (the agent runs it through the Caspi MCP tool
            and saves the JSON result to tools/state/caspi_sales.json).

  aggregate --caspi tools/state/caspi_sales.json
            Joins sales onto the prep data → tools/state/rate_tables.json.

  slides    --deck <presentationId>
            Inserts one "Overview (<series>) 클레임 + 배드리뷰 / 판매량 TOP N"
            table slide per series right after the last Overview slide of the
            deck. Idempotent: earlier rate_* slides are replaced.

  sheet     [--title ...]
            Optional: a Google Sheet with one tab per series, whole catalogue
            ranked by 판매량(EU).

Auth: ~/.config/gws_shim/token.json (kjw@spigen.com; Sheets/Slides/Drive).
"""
import argparse, collections, json, os, re, sys, time
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, 'state'); os.makedirs(STATE, exist_ok=True)
CFG = json.load(open(os.path.join(HERE, 'series.json')))
EMU = 12700

def creds():
    p = os.path.expanduser('~/.config/gws_shim/token.json')
    d = json.load(open(p)); c = Credentials.from_authorized_user_info(d); c.refresh(Request())
    d['token'] = c.token; json.dump(d, open(p, 'w'))
    return c

def col(h, name):
    if name in h: return h.index(name)
    for i, x in enumerate(h):
        if x.strip().startswith(name): return i
    raise KeyError(name)

def in_window(s, start, end):
    m = re.search(r'(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})', str(s or ''))
    if not m: return False
    iso = '%s-%02d-%02d' % (m.group(1), int(m.group(2)), int(m.group(3)))
    return (not start or iso >= start) and (not end or iso <= end)

# ----------------------------------------------------------------- prep
def cmd_prep(a):
    ss = build('sheets', 'v4', credentials=creds()).spreadsheets()
    series = [s.strip() for s in a.series.split(',') if s.strip()]
    for s in series:
        if s not in CFG['series']: sys.exit('unknown series %s (see series.json)' % s)
    eu_claim, eu_rev = set(CFG['euClaimCountries']), set(CFG['euReviewCountries'])
    out = {'window': {'start': a.start, 'end': a.end}, 'series': {}}
    # Zendesk claims (Product Issue, EU) once
    zd = CFG['zendesk']
    v = ss.values().get(spreadsheetId=zd['sheet'], range="'%s'!A1:AZ40000" % zd['tab']).execute().get('values', [])
    h = v[0]; zi = {k: col(h, k) for k in ('SKU', 'Category', 'Country', '인입사유', 'Ticket created - Date')}
    zrows = []
    for r in v[1:]:
        g = lambda k: (r[zi[k]] if zi[k] < len(r) else '').strip()
        if g('Category') != zd['category'] or g('Country') not in eu_claim: continue
        if a.start or a.end:
            if not in_window(g('Ticket created - Date'), a.start, a.end): continue
        zrows.append((g('SKU'), g('인입사유')))
    for key in series:
        sc = CFG['series'][key]
        cat = ss.values().get(spreadsheetId=sc['sheet'], range="'%s'!A1:H500" % sc['catalogTab']).execute().get('values', [])
        ch = cat[0]; ci = {k: col(ch, k) for k in ('ASIN', 'SKU', '기종명', '모델명', '색상명')}
        items = {}
        for r in cat[1:]:
            g = lambda k: (r[ci[k]] if ci[k] < len(r) else '').strip()
            if g('SKU') and g('SKU') not in items:
                items[g('SKU')] = {'sku': g('SKU'), 'asin': g('ASIN'), 'device': g('기종명'), 'model': g('모델명'), 'color': g('색상명'),
                                   'claims': 0, 'bad': 0, 'sold': 0, 'refund': 0, 'reasons': collections.Counter()}
        for sku, reason in zrows:
            if sku in items:
                items[sku]['claims'] += 1
                if reason: items[sku]['reasons'][reason] += 1
        rv = ss.values().get(spreadsheetId=sc['sheet'], range="'%s'!A1:Z8000" % sc['badReviewTab']).execute().get('values', [])
        rh = rv[0]; si, co, ti, di = col(rh, 'SKU'), col(rh, '국가(tag)'), col(rh, '인입사유(tag)'), col(rh, 'Created 날짜')
        for r in rv[1:]:
            g = lambda i: (r[i] if i < len(r) else '').strip()
            if g(si) not in items or g(co) not in eu_rev: continue
            if (a.start or a.end) and not in_window(g(di), a.start, a.end): continue
            items[g(si)]['bad'] += 1
            if g(ti) and g(ti) != '긍정 리뷰': items[g(si)]['reasons'][g(ti)] += 1
        for it in items.values():
            it['reasons'] = [[k, n] for k, n in it['reasons'].most_common(2)]
        out['series'][key] = {'label': sc['label'], 'items': list(items.values())}
        print('%s: %d SKUs, claims %d, bad reviews %d' % (sc['label'], len(items), sum(i['claims'] for i in items.values()), sum(i['bad'] for i in items.values())))
    json.dump(out, open(os.path.join(STATE, 'prep.json'), 'w'), ensure_ascii=False, indent=1)
    skus = sorted({it['sku'] for s in out['series'].values() for it in s['items']})
    yr = (a.start or a.end or time.strftime('%Y-%m-%d'))[:4]
    sql = ("SELECT SUBSTR(SELLER_SKU,1,8) AS SKU, ASIN, TRANSACTION_TYPE, MARKETPLACE, SUM(TRY_TO_NUMBER(QTY)) AS QTY, COUNT(*) AS N, "
           "MIN(DATA_MONTH) AS FIRST_MONTH, MAX(DATA_MONTH) AS LAST_MONTH\nFROM %s\nWHERE DATA_YEAR = '%s' AND SUBSTR(SELLER_SKU,1,8) IN (%s)\nGROUP BY 1,2,3,4 ORDER BY 1,3,4"
           % (CFG['caspiTable'], yr, ','.join("'%s'" % s for s in skus)))
    open(os.path.join(STATE, 'caspi_sales.sql'), 'w').write(sql)
    print('\nCaspi LM SQL written to tools/state/caspi_sales.sql — run it with the Caspi MCP run_query tool (limit 5000) and save the JSON result to tools/state/caspi_sales.json')

# ------------------------------------------------------------ aggregate
def cmd_aggregate(a):
    prep = json.load(open(os.path.join(STATE, 'prep.json')))
    raw = open(a.caspi).read(); data = json.loads(raw[raw.index('{'):])
    sale, refund = collections.Counter(), collections.Counter()
    for r in data['rows']:
        q = float(r.get('QTY') or 0)
        if r['TRANSACTION_TYPE'] == 'SALE': sale[r['SKU']] += q
        elif r['TRANSACTION_TYPE'] == 'REFUND': refund[r['SKU']] += q
    tables = {}
    for key, s in prep['series'].items():
        rows = []
        for it in s['items']:
            it = dict(it); it['sold'] = int(sale.get(it['sku'], 0)); it['refund'] = int(refund.get(it['sku'], 0))
            it['rate'] = (it['claims'] + it['bad']) / it['sold'] * 100 if it['sold'] else None
            rows.append(it)
        tables[key] = rows
        top = sorted([t for t in rows if t['sold'] >= CFG['minSold']], key=lambda t: -t['rate'])[:CFG['topN']]
        print('\n== %s (EU sold %s)' % (s['label'], format(sum(t['sold'] for t in rows), ',')))
        for i, t in enumerate(top, 1):
            print(' %d %s %s | sold %d | claims %d bad %d | %.2f%% | %s' % (i, t['sku'], t['model'], t['sold'], t['claims'], t['bad'], t['rate'], ', '.join(x for x, _ in t['reasons'])))
    json.dump({'window': prep['window'], 'tables': tables, 'labels': {k: s['label'] for k, s in prep['series'].items()}},
              open(os.path.join(STATE, 'rate_tables.json'), 'w'), ensure_ascii=False, indent=1)

# --------------------------------------------------------------- slides
def rgb(h): h = h.lstrip('#'); return {'red': int(h[:2], 16) / 255, 'green': int(h[2:4], 16) / 255, 'blue': int(h[4:], 16) / 255}
def txt(e): return ''.join(te.get('textRun', {}).get('content', '') for te in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()

def cmd_slides(a):
    svc = build('slides', 'v1', credentials=creds()).presentations()
    R = json.load(open(os.path.join(STATE, 'rate_tables.json')))
    PID = a.deck; N = CFG['topN']
    win = R['window']; wtxt = ('%s~%s' % (win['start'] or '2026.01.01', win['end'] or '')) if (win['start'] or win['end']) else '2026 YTD'
    pres = svc.get(presentationId=PID).execute(); slides = pres['slides']
    old = [x['objectId'] for x in slides if x['objectId'].startswith('rate_')]
    if old:
        svc.batchUpdate(presentationId=PID, body={'requests': [{'deleteObject': {'objectId': o}} for o in old]}).execute()
        pres = svc.get(presentationId=PID).execute(); slides = pres['slides']
    def find_last(pred):
        idx = None
        for i, s in enumerate(slides):
            for e in s.get('pageElements', []):
                if 'shape' in e and pred(txt(e)): idx = i
        return idx
    COLS = ['순위', '제품 (기종 / 모델 / 색상)', 'SKU', 'ASIN', '판매량 (EU)', '클레임', '배드리뷰', '비율 (%)']
    W = [32, 250, 56, 72, 56, 40, 48, 46]
    for key, rows_all in R['tables'].items():
        label = R['labels'][key]; fm = CFG['series'][key]['familyMatch']
        src_idx = find_last(lambda t: fm in t and 'Claims / Reviews' in t)
        if src_idx is None: print('no card slide for', label, '— skipped'); continue
        dest = find_last(lambda t: t.startswith('Overview')) + 1
        src = slides[src_idx]; new_id = 'rate_%s_%d' % (key, int(time.time()))
        svc.batchUpdate(presentationId=PID, body={'requests': [
            {'duplicateObject': {'objectId': src['objectId'], 'objectIds': {src['objectId']: new_id}}},
            {'updateSlidesPosition': {'slideObjectIds': [new_id], 'insertionIndex': dest}}]}).execute()
        new = svc.get(presentationId=PID).execute()['slides'][dest]
        reqs = []; title_id = None
        for e in new['pageElements']:
            t = e.get('transform', {}); L = t.get('translateX', 0) / EMU; T = t.get('translateY', 0) / EMU
            is_title = 'shape' in e and abs(L - 82) < 1 and abs(T - 17.4) < 1
            keep = ('elementGroup' in e) or ('image' in e and L < 5 and 150 < T < 200) or is_title
            if is_title: title_id = e['objectId']
            if not keep: reqs.append({'deleteObject': {'objectId': e['objectId']}})
        reqs += [{'deleteText': {'objectId': title_id, 'textRange': {'type': 'ALL'}}},
                 {'insertText': {'objectId': title_id, 'text': 'Overview (%s) 클레임 + 배드리뷰 / 판매량 TOP %d' % (label, N)}},
                 {'updateTextStyle': {'objectId': title_id, 'style': {'fontSize': {'magnitude': 15, 'unit': 'PT'}}, 'fields': 'fontSize'}}]
        rows = sorted([t for t in rows_all if t['sold'] >= CFG['minSold']], key=lambda t: -t['rate'])[:N]
        tid = new_id + '_tbl'
        reqs.append({'createTable': {'objectId': tid, 'elementProperties': {'pageObjectId': new_id,
            'size': {'width': {'magnitude': 600 * EMU, 'unit': 'EMU'}, 'height': {'magnitude': 280 * EMU, 'unit': 'EMU'}},
            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 82 * EMU, 'translateY': 60 * EMU, 'unit': 'EMU'}}, 'rows': N + 1, 'columns': 8}})
        for ci, w in enumerate(W):
            reqs.append({'updateTableColumnProperties': {'objectId': tid, 'columnIndices': [ci], 'tableColumnProperties': {'columnWidth': {'magnitude': w * EMU, 'unit': 'EMU'}}, 'fields': 'columnWidth'}})
        for ri in range(N + 1):
            reqs.append({'updateTableRowProperties': {'objectId': tid, 'rowIndices': [ri], 'tableRowProperties': {'minRowHeight': {'magnitude': (30 if ri else 22) * EMU, 'unit': 'EMU'}}, 'fields': 'minRowHeight'}})
        def cell(ri, ci, text, bold=False, color='#ffffff', size=8, align='CENTER', fill=None):
            loc = {'rowIndex': ri, 'columnIndex': ci}
            if text: reqs.append({'insertText': {'objectId': tid, 'cellLocation': loc, 'text': text}})
            reqs.append({'updateTextStyle': {'objectId': tid, 'cellLocation': loc, 'style': {'fontFamily': 'Arial', 'fontSize': {'magnitude': size, 'unit': 'PT'}, 'bold': bold, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}}, 'fields': 'fontFamily,fontSize,bold,foregroundColor'}})
            reqs.append({'updateParagraphStyle': {'objectId': tid, 'cellLocation': loc, 'style': {'alignment': align}, 'fields': 'alignment'}})
            reqs.append({'updateTableCellProperties': {'objectId': tid, 'tableRange': {'location': loc, 'rowSpan': 1, 'columnSpan': 1}, 'tableCellProperties': {
                'tableCellBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(fill or ('#1c2352' if ri % 2 else '#161c45'))}, 'alpha': 1}}, 'contentAlignment': 'MIDDLE'}, 'fields': 'tableCellBackgroundFill,contentAlignment'}})
        for ci, h in enumerate(COLS): cell(0, ci, h, True, '#c9cdd8', 8, fill='#0b1030')
        for ri, t in enumerate(rows, 1):
            cell(ri, 0, str(ri), True)
            name = '%s ㅣ %s ㅣ %s' % (t['device'], t['model'], t['color'])
            sub = '주요 인입사유: ' + (', '.join(x for x, _ in t['reasons']) if t['reasons'] else '-')
            cell(ri, 1, name + '\n' + sub, False, '#ffffff', 7.5, 'START')
            reqs.append({'updateTextStyle': {'objectId': tid, 'cellLocation': {'rowIndex': ri, 'columnIndex': 1}, 'textRange': {'type': 'FROM_START_INDEX', 'startIndex': len(name) + 1},
                'style': {'fontSize': {'magnitude': 6.5, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb('#9aa0c0')}}, 'bold': False}, 'fields': 'fontSize,foregroundColor,bold'}})
            cell(ri, 2, t['sku']); cell(ri, 3, t['asin']); cell(ri, 4, '{:,}'.format(t['sold'])); cell(ri, 5, str(t['claims'])); cell(ri, 6, str(t['bad']))
            cell(ri, 7, '%.2f%%' % t['rate'], True, '#ff6b6b' if t['rate'] >= 2 else '#ffffff')
        reqs.append({'updateTableBorderProperties': {'objectId': tid, 'tableRange': {'location': {'rowIndex': 0, 'columnIndex': 0}, 'rowSpan': N + 1, 'columnSpan': 8}, 'borderPosition': 'ALL',
            'tableBorderProperties': {'tableBorderFill': {'solidFill': {'color': {'rgbColor': rgb('#2a3168')}, 'alpha': 1}}, 'weight': {'magnitude': 0.75 * EMU, 'unit': 'EMU'}}, 'fields': 'tableBorderFill,weight'}})
        note = ('출처: Caspi LM (Amazon Seller VAT 거래 데이터) ㅣ 기준: 판매량 = Amazon EU+UK 판매 수량(%s 기준 연도, SALE 기준) ㅣ 클레임 = Zendesk 4. Product Issue 티켓(EU 국가, %s) ㅣ '
                '배드리뷰 = Amazon 1~3점 리뷰(DE/FR/IT/ES/UK, %s) ㅣ 비율 = (클레임+배드리뷰) ÷ 판매량 ㅣ 판매 %d개 미만 SKU 제외 ㅣ 주요 인입사유 = 클레임+배드리뷰 합산 상위 2개(긍정 리뷰 제외) ㅣ '
                '미국/일본/인도 판매 데이터는 현재 Caspi LM 권한 범위 밖') % (wtxt[:4], wtxt, wtxt, CFG['minSold'])
        nid = new_id + '_note'
        reqs.append({'createShape': {'objectId': nid, 'shapeType': 'TEXT_BOX', 'elementProperties': {'pageObjectId': new_id,
            'size': {'width': {'magnitude': 600 * EMU, 'unit': 'EMU'}, 'height': {'magnitude': 26 * EMU, 'unit': 'EMU'}},
            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 82 * EMU, 'translateY': 372 * EMU, 'unit': 'EMU'}}}})
        reqs.append({'insertText': {'objectId': nid, 'text': note}})
        reqs.append({'updateTextStyle': {'objectId': nid, 'style': {'fontFamily': 'Arial', 'fontSize': {'magnitude': 6, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb('#8f95b3')}}}, 'fields': 'fontFamily,fontSize,foregroundColor'}})
        for k in range(0, len(reqs), 80): svc.batchUpdate(presentationId=PID, body={'requests': reqs[k:k + 80]}).execute()
        print('%s → slide %d (%s)' % (label, dest + 1, new_id))
        pres = svc.get(presentationId=PID).execute(); slides = pres['slides']

# ---------------------------------------------------------------- sheet
def cmd_sheet(a):
    ss = build('sheets', 'v4', credentials=creds()).spreadsheets()
    R = json.load(open(os.path.join(STATE, 'rate_tables.json')))
    tabs = [(R['labels'][k], k) for k in R['tables']]
    book = ss.create(body={'properties': {'title': a.title}, 'sheets': [{'properties': {'title': t, 'gridProperties': {'frozenRowCount': 3}}} for t, _ in tabs]}).execute()
    sid = book['spreadsheetId']; gids = {s['properties']['title']: s['properties']['sheetId'] for s in book['sheets']}
    hdr = ['순위', '기종', '모델', '색상', 'SKU', 'ASIN', '판매량(EU)', '환불(EU)', '클레임', '배드리뷰', '클레임+배드리뷰', '비율(%)', '주요 인입사유']
    fmt = []
    for title, key in tabs:
        rows = sorted(R['tables'][key], key=lambda t: -t['sold'])
        values = [[title + ' — 판매량(EU) 순위'], ['출처: Caspi LM · %s (Amazon EU+UK, SALE) ㅣ 클레임 = Zendesk Product Issue(EU) ㅣ 배드리뷰 = 1~3점(DE/FR/IT/ES/UK)' % CFG['caspiTable']], hdr]
        for i, t in enumerate(rows, 1):
            values.append([i, t['device'], t['model'], t['color'], t['sku'], t['asin'], t['sold'], t['refund'], t['claims'], t['bad'], t['claims'] + t['bad'],
                           (t['rate'] / 100) if t['rate'] is not None else '', ', '.join(x for x, _ in t['reasons'])])
        ss.values().update(spreadsheetId=sid, range="'%s'!A1" % title, valueInputOption='RAW', body={'values': values}).execute()
        g = gids[title]; n = len(rows)
        fmt += [{'repeatCell': {'range': {'sheetId': g, 'startRowIndex': 2, 'endRowIndex': 3}, 'cell': {'userEnteredFormat': {'backgroundColor': {'red': .07, 'green': .09, 'blue': .21}, 'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}}}, 'fields': 'userEnteredFormat(backgroundColor,textFormat)'}},
                {'repeatCell': {'range': {'sheetId': g, 'startRowIndex': 3, 'endRowIndex': 3 + n, 'startColumnIndex': 11, 'endColumnIndex': 12}, 'cell': {'userEnteredFormat': {'numberFormat': {'type': 'PERCENT', 'pattern': '0.00%'}}}, 'fields': 'userEnteredFormat.numberFormat'}},
                {'setBasicFilter': {'filter': {'range': {'sheetId': g, 'startRowIndex': 2, 'endRowIndex': 3 + n, 'startColumnIndex': 0, 'endColumnIndex': 13}}}}]
    ss.batchUpdate(spreadsheetId=sid, body={'requests': fmt}).execute()
    print('https://docs.google.com/spreadsheets/d/%s/edit' % sid)

if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('prep'); p.add_argument('--series', required=True); p.add_argument('--start', default=''); p.add_argument('--end', default='')
    p = sub.add_parser('aggregate'); p.add_argument('--caspi', default=os.path.join(STATE, 'caspi_sales.json'))
    p = sub.add_parser('slides'); p.add_argument('--deck', required=True)
    p = sub.add_parser('sheet'); p.add_argument('--title', default='GCX Bi-weekly — 판매량(EU) 대비 클레임·배드리뷰')
    a = ap.parse_args()
    {'prep': cmd_prep, 'aggregate': cmd_aggregate, 'slides': cmd_slides, 'sheet': cmd_sheet}[a.cmd](a)
