"""Shared per-period config + section-anchor lookup for the apple_theme scripts."""
import json, os, re
HERE = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(HERE, 'report.json')))

def section_anchors(pres):
    """[(chip label, slideObjectId)] — first slide of each section, found by title text (no hardcoded ids)."""
    def title(s):
        for e in s.get('pageElements', []):
            t = ''.join(x.get('textRun', {}).get('content', '') for x in e.get('shape', {}).get('text', {}).get('textElements', [])).strip()
            if t and e.get('transform', {}).get('translateY', 0) / 12700 < 30: return t
        return ''
    titles = [(s['objectId'], title(s)) for s in pres['slides']]
    first = lambda pred: next((sid for sid, t in titles if pred(t)), None)
    out = [('Overview', first(lambda t: re.match(r'^(1\.\s*)?Overview$', t) is not None) or first(lambda t: t.startswith(('1. Overview', 'Overview'))))]
    for s in CFG['series']:
        out.append((s['chip'], first(lambda t, p=s['title_prefix']: t.startswith(p))))
    out.append(('SIREN', first(lambda t: 'GCX SIREN' in t and t[:1].isdigit())))
    out.append(('Appendix', 'apl_appendix'))
    return [(k, v) for k, v in out if v]
