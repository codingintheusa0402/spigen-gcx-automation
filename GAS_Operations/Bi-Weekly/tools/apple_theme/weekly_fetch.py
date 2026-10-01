"""Weekly claim counts (26년 전체문의 '월/주차별') -> weekly.json for chart.py. Run with the homebrew python3."""
import os, sys, json, collections
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, os.path.join(HERE, '..'))
from rate_pipeline import creds
from googleapiclient.discovery import build
from cfg import CFG
ss = build('sheets', 'v4', credentials=creds()).spreadsheets()
v = ss.values().get(spreadsheetId=CFG['zendesk_sheet'], range="'26년 전체문의'!A1:AZ60000").execute()['values']
h = v[0]; ci = h.index('월/주차별')
c = collections.Counter(r[ci].strip() for r in v[1:] if len(r) > ci and r[ci].strip())
json.dump({'weeks': [(k, n, '') for k, n in c.items()], 'total_rows': len(v) - 1}, open(os.path.join(HERE, 'weekly.json'), 'w'), ensure_ascii=False)
print('weeks', len(c), 'rows', len(v) - 1)
