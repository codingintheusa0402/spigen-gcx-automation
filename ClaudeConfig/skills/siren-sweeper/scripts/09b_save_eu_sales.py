"""Step 9b — find the latest run_query result with BASE_SKU/UNITS/UK_UNITS in the newest session transcript and
save {base_sku: units} to eu_sales.json. Usage: python3 09b_save_eu_sales.py [transcript.jsonl]"""
import sys, json, glob, os, common  # noqa
f = sys.argv[1] if len(sys.argv) > 1 else max(glob.glob(os.path.expanduser('~/.claude/projects/-Users-kevinkim/*.jsonl')), key=os.path.getmtime)
res = None
for ln in open(f):
    if 'UK_UNITS' in ln and 'BASE_SKU' in ln and 'tool_result' in ln:
        for c in json.loads(ln).get('message', {}).get('content', []):
            if isinstance(c, dict) and c.get('type') == 'tool_result':
                t = c['content'] if isinstance(c['content'], str) else c['content'][0].get('text', '')
                if '"rows"' in t: res = json.loads(t)
assert res and not res.get('truncated'), 'no (complete) EU sales result found'
eu = {r['BASE_SKU']: int(float(r['UNITS'])) for r in res['rows']}
json.dump(eu, open('eu_sales.json', 'w'), indent=1); print(len(eu), 'SKUs,', sum(eu.values()), 'EU units')
