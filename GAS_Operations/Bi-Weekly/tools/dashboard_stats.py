"""Dashboard numbers for the Overview slide: cumulative claims (26년 전체문의 rows ≤ cutoff), YoY vs 25년 same cutoff,
top-3 defect 인입사유 (non-inquiry). Usage: python3 dashboard_stats.py <cutoff YYYY-MM-DD> [--apply <deckId> --old '22,767|(~09.30)|-1.58%(YoY)|황변 2,843|충전불량 2,085|파손 943']"""
import sys, re
from collections import Counter
from datetime import datetime
from rate_pipeline import creds
from googleapiclient.discovery import build
ZD26, ZD25 = '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc', '1t5CJLsVw1hAVPspt_gBVCiM6vWSrEfvqQ8aV7xvoGD8'
SKIP = {'invoice발행관련', '상품문의', '기타사항', '호환문의'}          # inquiries, not defects
cut = sys.argv[1]; sh = build('sheets', 'v4', credentials=creds()).spreadsheets()
v = sh.values().get(spreadsheetId=ZD26, range="'26년 전체문의'!A1:BZ").execute()['values']; H = v[0]
di, ri = H.index('Ticket created - Date'), H.index('인입사유')
rows = [r for r in v[1:] if len(r) > di and r[di][:10] and r[di][:10] <= cut]
top = [(k, n) for k, n in Counter((r[ri] if len(r) > ri else '').strip() for r in rows).most_common(12) if k and k not in SKIP][:3]
d25 = [x[0] if x else '' for x in sh.values().get(spreadsheetId=ZD25, range="'25년 전체문의'!A:A").execute()['values'][1:]]
h25 = sh.values().get(spreadsheetId=ZD25, range="'25년 전체문의'!1:1").execute()['values'][0]
c25 = sh.values().get(spreadsheetId=ZD25, range=f"'25년 전체문의'!{chr(65 + h25.index('Ticket created - Date'))}2:{chr(65 + h25.index('Ticket created - Date'))}").execute()['values']
cut25 = cut.replace('2026', '2025')
n25 = sum(1 for x in c25 if x and x[0][:10] <= cut25)
yoy = (len(rows) - n25) / n25 * 100
new = [f'{len(rows):,}', f'(~{cut[5:7]}.{cut[8:10]})', f'{yoy:.2f}%(YoY)'] + [f'{k} {n:,}' for k, n in top]
print('cumulative', new[0], 'cutoff', new[1], 'YoY', new[2], 'top3', new[3:])
if '--apply' in sys.argv:
    deck = sys.argv[sys.argv.index('--apply') + 1]; old = sys.argv[sys.argv.index('--old') + 1].split('|')
    s = build('slides', 'v1', credentials=creds()).presentations()
    dash = next(sl['objectId'] for sl in s.get(presentationId=deck).execute()['slides'][:4]
                if any('누적 클레임' in str(e) for e in sl.get('pageElements', [])))
    r = s.batchUpdate(presentationId=deck, body={'requests': [{'replaceAllText': {'containsText': {'text': a, 'matchCase': True}, 'replaceText': b, 'pageObjectIds': [dash]}} for a, b in zip(old, new)]}).execute()
    print('replaced', [x.get('replaceAllText', {}).get('occurrencesChanged', 0) for x in r['replies']])
