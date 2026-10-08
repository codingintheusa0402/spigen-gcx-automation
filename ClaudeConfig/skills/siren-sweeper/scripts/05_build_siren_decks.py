"""Step 5 — build one SIREN deck per case with the siren-report skill (build_siren_slides.build_deck, unchanged).
Usage: python3 05_build_siren_decks.py [N ...]   (run two processes on halves to go faster). Writes deck_NN.txt."""
"""Builds the 20 SIREN decks with the siren-report skill's build_deck() (unchanged).
Only DeckBuilder.flush is wrapped: smaller chunks + if a chunk fails, its non-media requests are
applied first and each image/video request is retried alone, skipping any media Slides rejects."""
import sys, json, glob, os, re, time
import common  # noqa  (chdir to work dir)
import build_siren_slides as B
from googleapiclient.errors import HttpError

SKIPPED = []
def _exec(b, reqs):
    for k in range(5):
        try:
            return b.slides.presentations().batchUpdate(presentationId=b.pid, body={"requests": reqs}).execute()
        except HttpError as e:
            if e.resp.status in (429, 500, 502, 503) and k < 4 and not any('createImage' in r or 'createVideo' in r for r in reqs):
                time.sleep(5 * (k + 1)); continue
            if e.resp.status == 429 and k < 4:
                time.sleep(20); continue
            raise
def flush(self):
    reqs, self.reqs = self.reqs, []
    for i in range(0, len(reqs), 120):
        chunk = reqs[i:i + 120]
        try:
            _exec(self, chunk)
        except HttpError:
            media = [r for r in chunk if 'createImage' in r or 'createVideo' in r]
            _exec(self, [r for r in chunk if r not in media])
            for r in media:
                try:
                    _exec(self, [r])
                except HttpError as e:
                    SKIPPED.append((self.pid, json.dumps(r)[:300], str(e)[:200]))
B.DeckBuilder.flush = flush

tickets = json.load(open('tickets.json'))
ko = json.load(open('ko.json'))
media = json.load(open('media.json'))
ZD = 'https://spigenhelp.zendesk.com/attachments/token/%s/?name=%s'

def prepare(n):
    d = json.load(open(f'case_{n:02d}.json'))
    d['date'] = '2026. 10. 02'
    out = []
    for c in d['claims']:
        c = dict(c)
        if c['label'] == '클레임':
            tid = c['link_label']
            m = re.match(r'\[([^/\]]*)/([^\]]*)\]', c['detail_text'])
            c['_asin'] = (m.group(2).strip() if m else '') or d['product']['asin']
            c['detail_text'] = ko.get(tid) or tickets[tid]['b'][:300]
            imgs, vids = [], []
            for tok, name, ctype, size in tickets[tid]['a'][:3]:
                v = media.get(f'zd:{tok}', {})
                if v.get('kind') == 'img': imgs.append(v['url'])
                elif v.get('kind') == 'vid': vids.append(v['drive'])
            c['image_urls'], c['video_drive_ids'] = imgs, vids
        else:
            m = re.match(r'\[([^/\]]*)/', c['detail_text'])
            c['_asin'] = (m.group(1).strip() if m else '') or d['product']['asin']
            c['detail_text'] = re.sub(r'^\[[^\]]*\]\s*', '', c['detail_text']) or c.get('original_text', '')[:200]
            c['image_urls'] = [u for u in c.get('image_urls') or [] if media.get(f'rv:{u}', {}).get('kind') == 'img']
        out.append(c)
    # newest first within each type, same as the latest hand-built SIREN deck
    cl = sorted([c for c in out if c['label'] == '클레임'], key=lambda c: c['inflow_date'], reverse=True)
    rv = sorted([c for c in out if c['label'] != '클레임'], key=lambda c: c['inflow_date'], reverse=True)
    d['claims'] = cl + rv
    return d

if __name__ == '__main__':
    todo = [int(x) for x in sys.argv[1:]] or range(1, 21)
    for n in todo:
        if os.path.exists(f'deck_{n:02d}.txt'): continue
        d = prepare(n)
        json.dump(d, open(f'case_{n:02d}.final.json', 'w'), ensure_ascii=False, indent=1)
        t = time.time()
        url = B.build_deck(d)
        open(f'deck_{n:02d}.txt', 'w').write(url)
        print(n, url, f'{time.time()-t:.0f}s', 'skipped media so far:', len(SKIPPED), flush=True)
    json.dump(SKIPPED, open(f'skipped_{"_".join(map(str,todo))}.json', 'w'), ensure_ascii=False)
