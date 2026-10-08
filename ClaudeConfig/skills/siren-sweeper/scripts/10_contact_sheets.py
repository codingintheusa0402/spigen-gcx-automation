"""Step 10 — pick representative photos (user rule: photos that SHOW the defect; never invoices, order/Amazon
screenshots, packaging labels). Makes cs<no>.jpg contact sheets (up to 30 numbered thumbnails per case, newest VOC
first) for the agent to LOOK AT, then the agent writes picks.json {"<no>": [i, j, k]} and runs --apply.
Run with /usr/bin/python3 (Pillow). Usage: /usr/bin/python3 10_contact_sheets.py   |   ... --apply"""
import json, re, os, sys, common  # noqa
cases = json.load(open('cases.json'))
if '--apply' in sys.argv:
    pool, picks = json.load(open('pool.json')), json.load(open('picks.json'))
    for c in cases:
        if str(c['no']) in picks: c['photos'] = [pool[str(c['no'])][i]['url'] for i in picks[str(c['no'])]]
    json.dump(cases, open('cases.json', 'w'), ensure_ascii=False, indent=1); print('applied', len(picks)); sys.exit()
from PIL import Image, ImageDraw, ImageFont, ImageOps
media = json.load(open('media.json'))
url2key = {v['url']: k for k, v in media.items() if v.get('kind') == 'img'}
font = ImageFont.truetype('/System/Library/Fonts/AppleSDGothicNeo.ttc', 26)
pool = {}
for c in cases:
    fin = sorted(json.load(open(f"case_{c['n']:02d}.final.json"))['claims'], key=lambda x: x['inflow_date'], reverse=True)
    items = []
    for x in fin:
        for u in x.get('image_urls', []):
            k = url2key.get(u)
            if not k: continue
            local = os.path.join('media', re.sub(r'[^A-Za-z0-9._-]', '_', k)[:80]) + ('.jpg' if media[k].get('converted') else '')
            if os.path.exists(local): items.append({'url': u, 'local': local, 'ticket': x['link_label']})
        if len(items) >= 30: break
    pool[c['no']] = items = items[:30]
    W, cols = 240, 6; rows = max(1, (len(items) + cols - 1) // cols)
    sheet = Image.new('RGB', (cols * W, rows * W), 'white'); d = ImageDraw.Draw(sheet)
    for i, it in enumerate(items):
        x0, y0 = (i % cols) * W, (i // cols) * W
        try:
            im = ImageOps.exif_transpose(Image.open(it['local'])).convert('RGB'); im.thumbnail((W - 8, W - 8)); sheet.paste(im, (x0 + 4, y0 + 4))
        except Exception: d.text((x0 + 10, y0 + 100), 'ERR', fill='red', font=font)
        d.rectangle([x0 + 4, y0 + 4, x0 + 44, y0 + 36], fill='black'); d.text((x0 + 8, y0 + 4), str(i), fill='yellow', font=font)
    sheet.save(f"cs{c['no']:02d}.jpg", quality=80); print(c['no'], c['name'], c['defect'], len(items))
json.dump(pool, open('pool.json', 'w'), ensure_ascii=False, indent=1)
