#!/usr/bin/env python3
"""Monitoring numbers for the weekly block.
  stats.py FROM TO      e.g. stats.py 10/3 10/9
- 누적 totals (case/SP, 해외/국내) = '결과' tab summed up to TO  (matches what was reported before, verified 2026-10-12)
- ※TOP 3 인입사유 (누적)        = '1-3점' rows Created<=TO by 인입사유(tag)
- ①②③ (금주)                   = '1-3점' rows FROM<=Created<=TO by (기종명용 모델명, SKU, 인입사유(tag)), ties -> cumulative count
긍정 리뷰 / 기타사항 / blank tags are excluded from TOP3 and ①②③."""
import sys, json, collections
from _g import *
lo, hi = md(sys.argv[1]), md(sys.argv[2])
SKIP = ('긍정 리뷰', '기타사항', '')
out = {}
for label, sid, kr in SERIES:
    res = vals(sid, "'결과'!A1:ZZ30"); h = [x.replace(' ', '') for x in res[0]]
    j = h.index(f'{hi[0]}/{hi[1]}')
    def tot(cat, krflag=None):
        return sum(int(x[i]) for x in res[1:] if x and x[0] == cat and (krflag is None or (x[1] == 'KR') == krflag)
                   for i in range(2, j + 1) if i < len(x) and x[i].lstrip('-').isdigit())
    r = vals(sid, "'1-3점'!A1:AZ8000"); ix = {n: i for i, n in enumerate(r[0])}
    C = lambda x, n: x[ix[n]].strip() if ix[n] < len(x) else ''
    rows = [x for x in r[1:] if x and md(C(x, 'Created 날짜')) and C(x, '인입사유(tag)') not in SKIP]
    key = lambda x: (f"{C(x,'기종명')}용 {C(x,'모델명')}", C(x, 'SKU'), C(x, '인입사유(tag)'))
    o = {}
    for cat, nm in [('휴대폰케이스', 'case'), ('휴대폰보호필름', 'sp')]:
        rc = [x for x in rows if C(x, '대분류') == cat]
        cum = [x for x in rc if md(C(x, 'Created 날짜')) <= hi]
        wk = [x for x in cum if md(C(x, 'Created 날짜')) >= lo]
        cc = collections.Counter(key(x) for x in cum); wc = collections.Counter(key(x) for x in wk)
        o[nm] = {'total': tot(cat), **({'해외': tot(cat, False), '국내': tot(cat, True)} if kr else {}),
                 'top3_reasons_cum': collections.Counter(key(x)[2] for x in cum).most_common(5),
                 'week_top': [{'item': f"{k[0]} ({k[1]}) {k[2]}", 'week': wc[k], 'cum': cc[k]}
                              for k in sorted(wc, key=lambda k: (-wc[k], -cc[k]))[:5]]}
    out[label] = o
print(json.dumps(out, ensure_ascii=False, indent=0))
