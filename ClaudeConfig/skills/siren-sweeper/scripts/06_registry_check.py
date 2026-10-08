"""Step 6 — 기등록 check against the GCX/영업 SIREN registries (siren-finder only checks CQ Emergency Net).
Reads every registered row from:
  - GCX KPI '26년 SIREN' (+ previous year's '상반기/하반기 KPI 실적 검증' tabs) — link = 'PPT 링크' smart chip
  - [🚨 SIREN 🚨] '영업' tab — link = the LAST 'LINK' column (품질, cells show "Link"/"CQ SIREN"); user rule 2026-10-02
Outputs:
  regrows.json   all registered rows (src, date, product, skus, issue, imp, result, url)
  regmatch.json  {case no (siren deck numbering 1..N): [matched rows]} — same SKUs + same failure (Claude-judged)
  regreport.json registered issues whose same-SKU same-defect VOC kept coming after the latest registration
Usage: python3 06_registry_check.py"""
import json, re, os, datetime, concurrent.futures as cf, glob, common  # noqa
import find_siren as F
import build_siren_slides as B
from googleapiclient.discovery import build
sh = build('sheets', 'v4', credentials=B._creds()).spreadsheets()
SKU = re.compile(r'\b[A-Z]{3}\d{5}\b')

def grid(sid, rng):
    r = sh.get(spreadsheetId=sid, ranges=[rng], includeGridData=True,
               fields='sheets.data.rowData.values(formattedValue,hyperlink,chipRuns,textFormatRuns)').execute()
    rows = []
    for rd in r['sheets'][0]['data'][0].get('rowData', []):
        out = []
        for v in rd.get('values', []):
            urls = [c['chip']['richLinkProperties']['uri'] for c in v.get('chipRuns', []) if 'chip' in c]
            urls += [t['format']['link']['uri'] for t in v.get('textFormatRuns', []) if t.get('format', {}).get('link')]
            if v.get('hyperlink'): urls.append(v['hyperlink'])
            out.append((v.get('formattedValue', '') or '', urls))
        rows.append(out)
    return rows

def norm_date(s):
    m = re.match(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', s.strip())
    return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}' if m else ''

rows = []
y = datetime.date.today().year
GCX_TABS = [(common.GCX_KPI_26[0], common.GCX_KPI_26[1], 19, dict(date=1, title=2, product=3, sku=4, issue=5, reg=7, fb=8, imp=9, result=10), f"{y % 100}년 SIREN"),
            (common.GCX_KPI_25, '상반기 KPI 실적 검증', 54, dict(date=1, title=2, product=3, sku=4, issue=5, reg=9, fb=10, imp=11, result=12), f"'{(y - 1) % 100} 상반기"),
            (common.GCX_KPI_25, '하반기 KPI 실적 검증', 60, dict(date=1, title=2, product=3, sku=4, issue=5, reg=7, fb=8, imp=9, result=10), f"'{(y - 1) % 100} 하반기")]
for sid, tab, start, cm, label in GCX_TABS:
    for k, r in enumerate(grid(sid, f"'{tab}'!A{start}:N"), start):
        r = r + [('', [])] * 15
        if not re.match(r'^\d+$', r[0][0].strip()) or not norm_date(r[1][0]): continue
        if not r[cm['reg']][0].strip().upper().startswith('O'): continue
        rows.append({'src': f"{label} #{r[0][0].strip()}", 'date': norm_date(r[1][0]), 'title': r[cm['title']][0][:150],
                     'product': r[cm['product']][0][:100], 'skus': sorted(set(SKU.findall(r[cm['sku']][0]))), 'sku_text': r[cm['sku']][0],
                     'issue': r[cm['issue']][0], 'imp': r[cm['imp']][0], 'result': r[cm['result']][0][:300],
                     'url': (r[cm['title']][1] or [''])[0]})
sid, tab, start = common.SALES_SIREN
for r in grid(sid, f"'{tab}'!A{start}:AG"):
    r = r + [('', [])] * 34
    if not norm_date(r[1][0]): continue
    url = (r[32][1] or r[17][1] or [''])[0]          # last LINK (품질) column first, then the first LINK column
    rows.append({'src': f"영업 #{r[0][0].strip()}", 'date': norm_date(r[1][0]), 'title': r[16][0][:150], 'product': r[6][0][:100],
                 'skus': sorted(set(SKU.findall(r[7][0]))), 'sku_text': r[7][0], 'issue': r[11][0] + ' / ' + r[16][0][:120],
                 'imp': r[30][0], 'result': r[31][0][:300], 'url': url})
json.dump(rows, open('regrows.json', 'w'), ensure_ascii=False, indent=1)
print('registered rows:', len(rows))

PROMPT = """SIREN 등록 행(row)들과 비교 대상(case)이 있다. 각 case 에 대해, 같은 SKU/라인업이면서 '같은 불량 현상'인 등록 행의 idx만 고른다.
(예: '내장 자석 이탈'≈'자석탈락', '글라스 파손(내구성)'≈'글라스깨짐', '케이스 파손'≈'프레임파손', '스탠드가 튀어오름'≈'킥스탠드이슈';
'하단부 립 패임'≠'힌지파손', '형합'≠'코팅벗겨짐', '힌지 파손'≠'힌지 돌기 파손' — 부위(돌기·버튼·카메라 등)가 따로 명시되면 다른 불량). SKU가 '외 N개'로 일부만 적힌 행은 제품명이 같으면 possible=true 로 표시.
판단 근거는 case 의 samples(실제 고객 VOC 요약)를 우선한다 — defect 이름이 '제품품질불만'처럼 포괄적이면 samples 내용으로 불량 현상을 판단.
같은 불량 현상이 확실할 때만 match, 관련은 있지만 다른 부위/현상이거나 애매하면 possible, 아니면 제외.
JSON 배열만: [{"id": <case id>, "match": [<idx>...], "possible": [<idx>...]}]

%s"""
def ask(items):
    out = {}
    chunks = [items[i:i + 10] for i in range(0, len(items), 10)]
    with cf.ThreadPoolExecutor(6) as ex:
        for res in ex.map(lambda ch: F.claude_json(PROMPT % json.dumps(ch, ensure_ascii=False)), chunks):
            for it in res: out[it['id']] = it
    return out

# --- cases built in this run (case_NN.json)
items = []
for f in sorted(glob.glob('case_[0-9][0-9].json')):
    n = int(f[5:7]); d = json.load(open(f))
    skus = {x['sku'] for x in d['product']['skus']}
    cand = [k for k, r in enumerate(rows) if set(r['skus']) & skus or ('외' in r['sku_text'] and d['product']['device'].split()[0].lower() in r['product'].lower())]
    if cand:
        fin = json.load(open(f'case_{n:02d}.final.json'))['claims'] if os.path.exists(f'case_{n:02d}.final.json') else d['claims']
        samples = [x['detail_text'][:90] for x in fin if len(x.get('detail_text', '')) > 25][:5]
        items.append({'id': n, 'product': d['product']['product_name'], 'defect': d['product']['defect_type'], 'skus': sorted(skus), 'samples': samples,
                      'rows': [{'idx': k, 'product': rows[k]['product'], 'sku': rows[k]['sku_text'][:60], 'issue': rows[k]['issue'][:160]} for k in cand]})
res = ask(items)
regmatch = {}
for n, it in res.items():
    m = [dict(rows[k], possible=False) for k in it.get('match', [])] + [dict(rows[k], possible=True) for k in it.get('possible', [])]
    if m: regmatch[str(n)] = sorted(m, key=lambda r: r['date'])
json.dump(regmatch, open('regmatch.json', 'w'), ensure_ascii=False, indent=1)
for n in sorted(regmatch, key=int): print('case', n, [(r['src'], r['date'], r['possible']) for r in regmatch[n]])

# --- registered issues whose VOC continues (all siren-finder groups, current year)
cands = json.load(open(common.latest_candidates()))
items = []
for k, r in enumerate(rows):
    g = [c for c in cands if set(c['skus']) & set(r['skus'])]
    if g: items.append({'id': k, 'product': r['product'], 'defect': r['issue'][:120], 'skus': r['skus'],
                        'rows': [{'idx': j, 'product': c['name'], 'sku': ','.join(c['skus'][:4]), 'issue': c['defect']} for j, c in enumerate(g)]})
    r['_groups'] = g
res = ask(items)       # here 'rows' = candidate groups; idx = index into that row's group list
groups = {}
for k, it in res.items():
    r = rows[k]
    for j in it.get('match', []):
        c = r['_groups'][j]; groups.setdefault(c['key'], {'c': c, 'regs': []})['regs'].append(r)
report = []
for key, v in groups.items():
    c, regs = v['c'], v['regs']; latest = max(x['date'] for x in regs)
    cl = [t for t in c['claims'] if t['created'] > latest]
    rv = [t for t in c['reviews'] if re.match(r'\d{4}-', t['created']) and t['created'] > latest]
    last30 = (datetime.date.today() - datetime.timedelta(days=30)).isoformat()
    report.append({'group': c['name'], 'defect': c['defect'].split('_', 1)[-1], 'regs': [{x: y for x, y in r.items() if x != '_groups'} for r in regs],
                   'latest': latest, 'after': len(cl) + len(rv), 'claims': len(cl), 'reviews': len(rv),
                   'last30': sum(1 for t in cl + rv if t['created'] >= last30), 'last': max([t['created'] for t in cl + rv], default='-'),
                   'links': [f"https://spigenhelp.zendesk.com/agent/tickets/{t['id']}" for t in sorted(cl, key=lambda t: t['created'], reverse=True)[:3]]
                            + [t['url'] for t in sorted(rv, key=lambda t: t['created'], reverse=True)[:max(0, 3 - len(cl))]]})
report.sort(key=lambda x: -x['after'])
json.dump(report, open('regreport.json', 'w'), ensure_ascii=False, indent=1)
print('registered issues with VOC after registration:', sum(1 for x in report if x['after'] >= 10 or x['last30'] >= 5))
