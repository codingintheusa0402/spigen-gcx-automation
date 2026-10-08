"""Render every slide (LARGE thumbnails) to render_<deck>/NN.png for the visual check (always look at them:
unintended wraps, overlaps, missing images). Usage: python3 render.py [deck_id] [slide numbers...]"""
import sys, os, time, urllib.request, common  # noqa
from rate_pipeline import creds
from googleapiclient.discovery import build
s = build('slides', 'v1', credentials=creds()).presentations()
pid = sys.argv[1] if len(sys.argv) > 1 and len(sys.argv[1]) > 20 else open('summary_deck.txt').read().strip()
only = {int(x) for x in sys.argv[1:] if x.isdigit()}
out = f'render_{pid[:8]}'; os.makedirs(out, exist_ok=True)
for k in range(4):
    try: p = s.get(presentationId=pid, fields='slides.objectId').execute(); break
    except Exception: time.sleep(10)
for i, sl in enumerate(p['slides'], 1):
    if only and i not in only: continue
    th = s.pages().getThumbnail(presentationId=pid, pageObjectId=sl['objectId'], thumbnailProperties_thumbnailSize='LARGE').execute()
    urllib.request.urlretrieve(th['contentUrl'], f'{out}/{i:02d}.png')
print(os.path.abspath(out))
