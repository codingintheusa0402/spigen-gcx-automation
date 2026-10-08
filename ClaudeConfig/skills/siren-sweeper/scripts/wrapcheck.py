"""Usage: python3 wrapcheck.py [deck_id] — flags text boxes whose text would wrap unintentionally: estimates each paragraph's rendered width
(Hangul ≈ 1.0em, Latin/digits ≈ 0.6em, Inter) against box width minus the default 7.2pt L/R insets,
and reports boxes where the wrapped line count exceeds what the box height can hold or exceeds the
number of explicit lines (i.e. a line the author meant as one line breaks)."""
import sys, json, math, common  # noqa
from rate_pipeline import creds
from googleapiclient.discovery import build
s = build('slides', 'v1', credentials=creds()).presentations()
pid = sys.argv[1] if len(sys.argv) > 1 else open('summary_deck.txt').read().strip()
p = s.get(presentationId=pid).execute()
E = 12700
def w_of(t, size):
    return sum(size * (1.0 if ord(ch) > 0x2E80 else 0.30 if ch in ' .,:;·|()/-~' else 0.66 if ch.isupper() or ch.isdigit() else 0.55) for ch in t)
bad = []
for si, sl in enumerate(p['slides'][1:-1], 2):
    for e in sl.get('pageElements', []):
        tx = e.get('shape', {}).get('text')
        if not tx: continue
        size_ = None; content = ''
        for te in tx['textElements']:
            r = te.get('textRun')
            if r:
                content += r['content']
                fs = r.get('style', {}).get('fontSize', {}).get('magnitude')
                if fs: size_ = max(size_ or 0, fs)
        if not content.strip() or not size_: continue
        tr = e['transform']; W = e['size']['width']['magnitude'] * tr.get('scaleX', 1) / E
        H = e['size']['height']['magnitude'] * tr.get('scaleY', 1) / E
        paras = content.rstrip('\n').split('\n')
        avail = W - 14.4
        wrapped = sum(max(1, math.ceil(w_of(pp, size_) / max(avail, 1))) for pp in paras)
        cap = max(len(paras), int(H // (size_ * 1.45)))      # lines the box was sized for
        if wrapped > cap:
            bad.append((si, round(W), round(H), size_, len(paras), wrapped, content.strip()[:60].replace('\n', ' / ')))
for b in bad: print(b)
print('flagged', len(bad))
