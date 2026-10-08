"""Step 9 — data for the summary (Apple Bi-weekly style) deck -> cases.json.
Default selection = every case except Screen Protector ((SP) division; glass is reported by another team).
Numbering = sequential #1..N in SLIDE order (sections: Case → SDA·New Biz → Power Accessories → others).
Needs: case_NN.final.json + deck_NN.txt (step 5), tickets/ko (2–3), media.json (4), regmatch.json (6), eu_sales.json (9b).
Usage: python3 09_summary_data.py [--cases 1,2,5,...] [--include-sp]"""
import sys, json, re, os, argparse, common  # noqa
from collections import Counter
from spigen_images import catalog, resolve
import build_siren_slides as B
from googleapiclient.discovery import build
ap = argparse.ArgumentParser(); ap.add_argument('--cases'); ap.add_argument('--include-sp', action='store_true'); a = ap.parse_args()
cands = [c for c in json.load(open(common.latest_candidates())) if not c['registered']]
regmatch = json.load(open('regmatch.json')) if os.path.exists('regmatch.json') else {}
eu = json.load(open('eu_sales.json'))
cat = catalog()
N = len([f for f in os.listdir('.') if re.match(r'case_\d\d\.json$', f)])
chosen = [int(x) for x in a.cases.split(',')] if a.cases else \
         [n for n in range(1, N + 1) if a.include_sp or '(SP)' not in cands[n - 1]['defect']]
SECTION = {'Case': 0, 'SDA': 1, 'New Biz': 1, 'PAcc.': 2}
# 생산지: SKU_Master (먼데이보드) first, product master Data as fallback
sh = build('sheets', 'v4', credentials=B._creds()).spreadsheets()
def table(sid, key, man, org):
    v = sh.values().get(spreadsheetId=sid, range="'Data'").execute()['values']; h = v[0]
    out = {}
    for r in v[1:]:
        r = r + [''] * 80
        if r[h.index(man)].strip() or r[h.index(org)].strip(): out.setdefault(r[h.index(key)].strip()[:8], (r[h.index(man)].strip(), r[h.index(org)].strip()))
    return out
m1 = table(common.SKU_MASTER, 'name', '생산업체', '원산지'); m2 = table(common.PRODUCT_MASTER, 'SKU', '생산업체', '원산지정보')
out = []
for n in chosen:
    c = cands[n - 1]; fin = json.load(open(f'case_{n:02d}.final.json'))['claims']
    div = re.match(r'\(([^)]*)\)', c['defect']).group(1)
    months = Counter(x['created'][:7] for x in c['claims'] + c['reviews'] if re.match(r'\d{4}-\d{2}', x['created']))
    units = sum(eu.get(s, 0) for s in c['skus'])
    eu_voc = sum(1 for x in c['claims'] + c['reviews'] if x['country'].upper() in common.EU_COUNTRIES)
    mf, og = Counter(), Counter()
    for s in c['skus']:
        x = m1.get(s) or ('', ''); y = m2.get(s) or ('', '')
        if x[0] or y[0]: mf[(x[0] or y[0]).replace(' (구Wireless)', '')] += 1
        if x[1] or y[1]: og[x[1] or y[1]] += 1
    reg = [r for r in regmatch.get(str(n), []) if not r['possible']]
    newest = sorted(fin, key=lambda x: x['inflow_date'], reverse=True)[:2]
    out.append(dict(n=n, sec=SECTION.get(div, 3), div=div, name=c['name'], defect=c['defect'].split('_', 1)[-1], total=c['total'],
        claims=len(c['claims']), reviews=len(c['reviews']), first=c['first'], last=c['last'], skus=c['skus'],
        months=dict(sorted(months.items())), countries=Counter(x['country'] for x in c['claims'] + c['reviews']).most_common(),
        img=next((resolve(cat, s) for s in c['skus'] if resolve(cat, s)), None),
        photos=[u for cl in fin for u in cl.get('image_urls', [])][:3],          # replaced by 10_contact_sheets --apply
        deck=open(f'deck_{n:02d}.txt').read().strip(),
        reg=(reg[0]['src'] + (' (CQ)' if 'CQ' in reg[0]['imp'] else '')) if reg else None,
        reg_url=reg[0]['url'] if reg else None,
        reg_date=min(r['date'] for r in reg)[2:].replace('-', '.') if reg else None,
        eu_units=units, eu_voc=eu_voc, eu_rate=(eu_voc / units * 100 if units else None),
        maker=[k for k, _ in mf.most_common()], origin=[k for k, _ in og.most_common()],
        voc=[{'text': x['detail_text'], 'country': x['country'], 'date': x['inflow_date'], 'kind': '클레임' if x['label'] == '클레임' else '배드리뷰',
              'link': x.get('link_url') or f"https://spigenhelp.zendesk.com/agent/tickets/{x['link_label']}"} for x in newest]))
out.sort(key=lambda c: (c['sec'], chosen.index(c['n'])))
for i, c in enumerate(out, 1): c['no'] = i
json.dump(out, open('cases.json', 'w'), ensure_ascii=False, indent=1)
for c in out: print(c['no'], f"(deck {c['n']})", c['div'], c['name'], c['defect'], c['total'], c['eu_units'], f"{c['eu_rate']:.2f}%" if c['eu_rate'] else '-', c['reg'] or '')
