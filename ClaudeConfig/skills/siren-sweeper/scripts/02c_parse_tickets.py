"""Step 2c — parse the persisted get_page_text result into tickets.json.
Usage: python3 02c_parse_tickets.py <path to tool-results/*.json printed in the get_page_text output>"""
import sys, json, common  # noqa
from collections import Counter
raw = json.load(open(sys.argv[1]))
t = raw[0]['text'] if isinstance(raw, list) else raw
assert '@@END' in t, 'dump incomplete — re-run get_page_text with a larger max_chars'
recs = {}
for ln in t.split('\n'):
    if ln.startswith('@@R '):
        r = json.loads(ln[4:]); recs[r['id']] = r
json.dump(recs, open('tickets.json', 'w'), ensure_ascii=False)
ids = json.load(open('ids.json'))
print(len(recs), 'tickets | missing:', len(set(ids) - set(recs)))
print(Counter(a[2] for r in recs.values() for a in r['a']))
