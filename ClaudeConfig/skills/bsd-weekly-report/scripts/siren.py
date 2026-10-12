#!/usr/bin/env python3
"""26년 SIREN registry rows relevant to the weekly block.
  siren.py FROM TO "<ACS/AGL SKU,...>"   (SKUs = the items currently listed in the block's [🚨SIREN] section)
Prints: rows registered in FROM..TO (new items), and current status of the listed SKUs (CQ 피드백 날짜 + 개선 결과)."""
import sys, json
from _g import *
lo, hi = md(sys.argv[1]), md(sys.argv[2]); skus = [s.strip() for s in (sys.argv[3] if len(sys.argv) > 3 else '').split(',') if s.strip()]
v = vals(SIREN_SID, f"'{SIREN_TAB}'!A18:P400"); h = v[0]
rows = [dict(zip(h, x)) for x in v[1:] if len(x) > 2 and md(x[1] or '')]
new = [r for r in rows if lo <= md(r['날짜']) <= hi]
cur = [r for r in rows if any(s in r.get('SKU', '') for s in skus)]
pick = lambda r: {k: r.get(k, '') for k in ['날짜', '`', 'SKU', '이슈', '작성자', 'SIREN 등록', '피드백 날짜', 'SIREN 개선', '개선 결과']}
print(json.dumps({'new_in_window': [pick(r) for r in new], 'listed_status': [pick(r) for r in cur]}, ensure_ascii=False, indent=1))
