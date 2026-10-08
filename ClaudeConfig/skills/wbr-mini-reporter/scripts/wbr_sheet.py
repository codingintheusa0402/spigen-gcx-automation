#!/usr/bin/env python3
"""Read / backup / write one person's column block of a WBR tab via Sheets API (gws_shim token).
  wbr_sheet.py read   SPREADSHEET_ID "TAB" A1RANGE            -> prints JSON {cell: value} incl. row labels (cols A-C)
  wbr_sheet.py write  SPREADSHEET_ID "TAB" A1RANGE edits.json -> backs up the range, then writes {cell: value}
Writes use RAW so text like "- 0건" is never parsed as a formula; only values change, formatting stays."""
import json, os, re, sys, time
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
P=os.path.expanduser('~/.config/gws_shim/token.json'); d=json.load(open(P))
c=Credentials(token=d.get('token'),refresh_token=d['refresh_token'],client_id=d['client_id'],
  client_secret=d['client_secret'],token_uri=d.get('token_uri','https://oauth2.googleapis.com/token'),scopes=d.get('scopes'))
c.refresh(Request()); d['token']=c.token; json.dump(d,open(P,'w'))
S=build('sheets','v4',credentials=c).spreadsheets()
cmd,sid,tab,rng=sys.argv[1:5]
m=re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)',rng); col,r1,_,r2=m.group(1),int(m.group(2)),m.group(3),int(m.group(4))
def get(r): return S.values().get(spreadsheetId=sid,range=f"'{tab}'!{r}").execute().get('values',[])
cur={f"{col}{r1+i}":(row[0] if row else '') for i,row in enumerate(get(rng)+[[]]*(r2-r1+1))}
cur=dict(list(cur.items())[:r2-r1+1])
if cmd=='read':
    labels=get(f"A{r1}:C{r2}")
    out={k:{'label':' / '.join(x.strip() for x in (labels[i] if i<len(labels) else []) if x.strip()),'value':v} for i,(k,v) in enumerate(cur.items())}
    print(json.dumps(out,ensure_ascii=False,indent=1))
elif cmd=='write':
    edits=json.load(open(sys.argv[5]))
    bk=os.path.expanduser(f"~/.config/wbr_mini_reporter/backup_{sid[:8]}_{re.sub(r'[^0-9A-Za-z]','',tab)}_{rng.replace(':','-')}_{time.strftime('%Y%m%d_%H%M%S')}.json")
    os.makedirs(os.path.dirname(bk),exist_ok=True); json.dump(cur,open(bk,'w'),ensure_ascii=False,indent=1)
    new=[[edits.get(k,v)] for k,v in cur.items()]
    S.values().update(spreadsheetId=sid,range=f"'{tab}'!{rng}",valueInputOption='RAW',body={'values':new}).execute()
    print('backup:',bk); print('changed cells:',[k for k,v in cur.items() if edits.get(k,v)!=v])
