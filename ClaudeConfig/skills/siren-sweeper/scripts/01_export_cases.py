"""Step 1 — export the top-N open (not CQ-registered) siren-finder candidates as siren-report case skeletons.
Usage: python3 01_export_cases.py [--top 20]   -> case_01.json … case_NN.json, ids.json (claim ticket ids)"""
import argparse, json, common  # noqa
import find_siren as F
from lineup import ProductMaster
ap = argparse.ArgumentParser(); ap.add_argument('--top', type=int, default=20); a = ap.parse_args()
c = json.load(open(common.latest_candidates()))
open_c = [x for x in c if not x['registered']][:a.top]
pm = ProductMaster(F.read_tab(F.sheets(), *F.PRODUCT_MASTER))
ids = set()
for i, x in enumerate(open_c, 1):
    x['lineup'] = tuple(x['lineup'])
    d = F.export_case(x, pm)
    json.dump(d, open(f'case_{i:02d}.json', 'w'), ensure_ascii=False, indent=1)
    ids |= {cl['link_label'] for cl in d['claims'] if cl['label'] == '클레임'}
    print(i, x['name'], x['defect'], len(x['claims']), len(x['reviews']))
json.dump(sorted(ids), open('ids.json', 'w'))
print('work dir:', common.WORK, '| claim tickets:', len(ids))
