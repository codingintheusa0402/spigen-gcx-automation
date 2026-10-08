"""Step 4 — first 3 customer images/videos per ticket + review photos: download, verify (sips), convert HEIC/
oversized/unknown to JPEG (-Z 2000) and re-host on Drive (anyone-reader); videos -> Drive (spigen.com reader)
for Slides createVideo. Result cache: media.json (key -> img url | vid drive id). Resumable; failed keys retried
if you delete them from media.json."""
import urllib.parse, sys, json, glob, os, re, subprocess, urllib.request, concurrent.futures as cf, threading
import common  # noqa  (chdir to work dir)
import build_siren_slides as B
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
recs = json.load(open('tickets.json'))
OUT = 'media.json'
os.makedirs('media', exist_ok=True)
res = json.load(open(OUT)) if os.path.exists(OUT) else {}   # key -> {"kind":"img","url":..} | {"kind":"vid","drive":..} | {"kind":"bad"}
lock = threading.Lock()
creds = B._creds()
_local = threading.local()
def drive():
    if not hasattr(_local, 'd'):
        _local.d = build('drive', 'v3', credentials=creds)
    return _local.d
ZD = 'https://spigenhelp.zendesk.com/attachments/token/%s/?name=%s'

def folder(name, anyone):
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false and 'me' in owners"
    r = drive().files().list(q=q, fields='files(id)').execute()['files']
    if r: return r[0]['id']
    fid = drive().files().create(body={'name': name, 'mimeType': 'application/vnd.google-apps.folder'}, fields='id').execute()['id']
    return fid
import datetime
TAG = datetime.date.today().strftime('%y%m%d')
IMG_FOLDER = folder(f'SIREN_고객사진_변환_{TAG}', True)
VID_FOLDER = folder(f'SIREN_고객영상_{TAG}', False)

def upload(path, mime, parent, perm):
    f = drive().files().create(body={'name': os.path.basename(path), 'parents': [parent]},
                               media_body=MediaFileUpload(path, mimetype=mime, resumable=True), fields='id').execute()
    drive().permissions().create(fileId=f['id'], body=perm, fields='id').execute()
    return f['id']

def dims(path):
    o = subprocess.run(['sips', '-g', 'pixelWidth', '-g', 'pixelHeight', '-g', 'format', path], capture_output=True, text=True).stdout
    w = re.search(r'pixelWidth: (\d+)', o); h = re.search(r'pixelHeight: (\d+)', o); fm = re.search(r'format: (\S+)', o)
    return (int(w.group(1)) if w else 0, int(h.group(1)) if h else 0, fm.group(1) if fm else '')

def handle(key, url, ctype, name):
    local = os.path.join('media', re.sub(r'[^A-Za-z0-9._-]', '_', key)[:80])
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=120) as r, open(local, 'wb') as f:
        f.write(r.read())
    size = os.path.getsize(local)
    if ctype.startswith('video') or re.search(r'\.(mov|mp4)$', name, re.I):
        if size > 300e6: return {'kind': 'bad', 'why': 'video too big'}
        fid = upload(local, ctype if ctype.startswith('video') else 'video/mp4', VID_FOLDER,
                     {'type': 'domain', 'domain': 'spigen.com', 'role': 'reader'})
        return {'kind': 'vid', 'drive': fid}
    w, h, fm = dims(local)
    if not w: return {'kind': 'bad', 'why': 'not an image'}
    if fm in ('jpeg', 'png', 'gif') and w * h <= 25e6 and size < 45e6:
        return {'kind': 'img', 'url': url, 'w': w, 'h': h}
    out = local + '.jpg'
    subprocess.run(['sips', '-s', 'format', 'jpeg', '-Z', '2000', local, '--out', out], capture_output=True)
    if not os.path.exists(out): return {'kind': 'bad', 'why': 'convert failed ' + fm}
    w, h, _ = dims(out)
    fid = upload(out, 'image/jpeg', IMG_FOLDER, {'type': 'anyone', 'role': 'reader'})
    return {'kind': 'img', 'url': f'https://drive.google.com/uc?export=view&id={fid}', 'w': w, 'h': h, 'converted': True}

jobs = []
for tid, r in recs.items():
    for a in r['a'][:3]:
        tok, name, ctype, size = a
        jobs.append((f'zd:{tok}', ZD % (tok, urllib.parse.quote(name)), ctype or '', name))
for f in sorted(glob.glob('case_*.json')):
    for c in json.load(open(f))['claims']:
        for u in c.get('image_urls') or []:
            jobs.append((f'rv:{u}', u, 'image/jpeg', u))
jobs = [j for j in jobs if j[0] not in res]
print('jobs', len(jobs), flush=True)
def run(j):
    try: return j[0], handle(*j)
    except Exception as e: return j[0], {'kind': 'bad', 'why': str(e)[:200]}
with cf.ThreadPoolExecutor(12) as ex:
    for n, (k, v) in enumerate(ex.map(run, jobs), 1):
        with lock:
            res[k] = v
            if n % 50 == 0 or n == len(jobs):
                json.dump(res, open(OUT, 'w')); print(n, flush=True)
json.dump(res, open(OUT, 'w'))
from collections import Counter
print(Counter(v['kind'] for v in res.values()), Counter(v.get('why','')[:40] for v in res.values() if v['kind']=='bad'))
