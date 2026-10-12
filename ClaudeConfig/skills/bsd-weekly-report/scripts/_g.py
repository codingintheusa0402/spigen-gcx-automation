"""Shared Sheets client (gws_shim token, kjw@spigen.com) + helpers."""
import json, os, re
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
P = os.path.expanduser('~/.config/gws_shim/token.json'); _d = json.load(open(P))
_c = Credentials(token=_d.get('token'), refresh_token=_d['refresh_token'], client_id=_d['client_id'],
                 client_secret=_d['client_secret'], token_uri=_d.get('token_uri', 'https://oauth2.googleapis.com/token'), scopes=_d.get('scopes'))
_c.refresh(Request()); _d['token'] = _c.token; json.dump(_d, open(P, 'w'))
S = build('sheets', 'v4', credentials=_c).spreadsheets()

WEEKLY_SID = '1J4l0m4bJCthuGtKTBk51a0UuZcPen0w_x1zyPpFb39U'   # 사업지원실 '26 주간업무
SIREN_SID, SIREN_TAB = '15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI', '26년 SIREN'
TEAM = '글로벌CX전략팀'
# Monitoring books shown in the block, in block order. Add/remove when a series starts/ends.
SERIES = [
    ('Galaxy Z Fold 8 / Flip 8 / Fold 8 Ultra Series', '19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4', True),
    ('Google Pixel 11 Series', '12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI', False),
    ('iPhone 18 Series', '1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU', True),
]  # (label, sheet id, has 국내(KR) split)

def vals(sid, rng): return S.values().get(spreadsheetId=sid, range=rng).execute().get('values', [])
def tabs(sid): return [s['properties'] for s in S.get(spreadsheetId=sid, fields='sheets.properties(title,sheetId,index)').execute()['sheets']]
def md(s):
    """'2026-10-09' / '2026/10/09' / '10/9' -> (10, 9)"""
    p = re.split(r'[-/.]', s.strip().split(' ')[0])
    try: return (int(p[1]), int(p[2])) if len(p[0]) == 4 else (int(p[0]), int(p[1]))
    except Exception: return None
def a1(rng):
    m = re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rng); return m.group(1), int(m.group(2)), int(m.group(4))
