import sys
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_cards import svc, P, E, rgb, txt, FONT, INK, GRAY, GRAY2
BLUE, CHIP, HAIR, CANVAS, LINKGRAY = '#0071E3', '#E8E8ED', '#D2D2D7', '#F5F5F7', '#424245'

def box(sid, oid, kind, x, y, w, h):
    return {'createShape': {'objectId': oid, 'shapeType': kind, 'elementProperties': {'pageObjectId': sid,
            'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': h*E, 'unit': 'EMU'}},
            'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}}
def link_style(url=None, slide=None):
    if url: return {'url': url}
    if slide: return {'pageObjectId': slide}
def text(sid, oid, x, y, w, h, s, size, color, bold=False, align='START', url=None, slide=None, underline=False):
    st = {'fontFamily': FONT, 'fontSize': {'magnitude': size, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}, 'bold': bold, 'underline': underline}
    f = 'fontFamily,fontSize,foregroundColor,bold,underline'
    if url or slide: st['link'] = link_style(url, slide); f += ',link'
    return [box(sid, oid, 'TEXT_BOX', x, y, w, h), {'insertText': {'objectId': oid, 'text': s}},
            {'updateTextStyle': {'objectId': oid, 'style': st, 'fields': f, 'textRange': {'type': 'ALL'}}},
            {'updateParagraphStyle': {'objectId': oid, 'style': {'alignment': align, 'lineSpacing': 100}, 'fields': 'alignment,lineSpacing', 'textRange': {'type': 'ALL'}}},
            {'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'contentAlignment': 'MIDDLE'}, 'fields': 'contentAlignment'}}]
def _capsule(sid, oid, x, y, w, h, color):
    """True capsule (semicircular ends): rect + 2 circles of diameter h."""
    r = h / 2; out = []
    for i, (px, py, pw, ph, kind) in enumerate([(x + r, y, max(w - h, 0.1), h, 'RECTANGLE'), (x, y, h, h, 'ELLIPSE'), (x + w - h, y, h, h, 'ELLIPSE')]):
        pid = f'{oid}_c{i}'
        out += [box(sid, pid, kind, px, py, pw, ph),
                {'updateShapeProperties': {'objectId': pid, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(color)}}},
                  'outline': {'propertyState': 'NOT_RENDERED'}}, 'fields': 'shapeBackgroundFill.solidFill.color,outline.propertyState'}}]
    return out
def pill(sid, oid, x, y, w, h, label, style='filled', size=9, url=None, slide=None, under='#FFFFFF'):
    """Apple capsule button. filled/red/chip/dark = solid capsule; outline = 1pt blue ring over `under`."""
    fill = {'filled': BLUE, 'outline': BLUE, 'chip': CHIP, 'dark': INK, 'red': '#FF3B30'}[style]
    color = {'filled': '#FFFFFF', 'outline': BLUE, 'chip': INK, 'dark': '#FFFFFF', 'red': '#FFFFFF'}[style]
    out = _capsule(sid, oid + 'a', x, y, w, h, fill)
    if style == 'outline': out += _capsule(sid, oid + 'b', x + 1, y + 1, w - 2, h - 2, under)
    st = {'fontFamily': FONT, 'fontSize': {'magnitude': size, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}, 'underline': False}
    f = 'fontFamily,fontSize,foregroundColor,underline'
    if url or slide: st['link'] = link_style(url, slide); f += ',link'
    out += [box(sid, oid, 'TEXT_BOX', x, y, w, h), {'insertText': {'objectId': oid, 'text': label}},
            {'updateTextStyle': {'objectId': oid, 'style': st, 'fields': f, 'textRange': {'type': 'ALL'}}},
            {'updateParagraphStyle': {'objectId': oid, 'style': {'alignment': 'CENTER', 'lineSpacing': 100}, 'fields': 'alignment,lineSpacing', 'textRange': {'type': 'ALL'}}},
            {'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'contentAlignment': 'MIDDLE'}, 'fields': 'contentAlignment'}}]
    return out
def hair(sid, oid, x, y, w):
    return [{'createLine': {'objectId': oid, 'lineCategory': 'STRAIGHT', 'elementProperties': {'pageObjectId': sid,
             'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': 0, 'unit': 'EMU'}},
             'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}},
            {'updateLineProperties': {'objectId': oid, 'lineProperties': {'lineFill': {'solidFill': {'color': {'rgbColor': rgb(HAIR)}}}, 'weight': {'magnitude': 0.75, 'unit': 'PT'}}, 'fields': 'lineFill.solidFill.color,weight'}}]
def bg(sid, c): return {'updatePageProperties': {'objectId': sid, 'pageProperties': {'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(c)}}}}, 'fields': 'pageBackgroundFill.solidFill.color'}}
def restyle(oid, size, color, bold=False, align='CENTER'):
    return [{'updateTextStyle': {'objectId': oid, 'style': {'fontFamily': FONT, 'fontSize': {'magnitude': size, 'unit': 'PT'}, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}, 'bold': bold}, 'fields': 'fontFamily,fontSize,foregroundColor,bold', 'textRange': {'type': 'ALL'}}},
            {'updateParagraphStyle': {'objectId': oid, 'style': {'alignment': align}, 'fields': 'alignment', 'textRange': {'type': 'ALL'}}}]
def move(oid, x, y, sw, size):  # absolute move with target size via scale
    return {'updatePageElementTransform': {'objectId': oid, 'applyMode': 'ABSOLUTE', 'transform': {'scaleX': sw[0]*E/size[0], 'scaleY': sw[1]*E/size[1], 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}

from cfg import CFG, section_anchors
D = 'https://docs.google.com/presentation/d/%s/edit'
S = 'https://docs.google.com/spreadsheets/d/%s/edit'
def appendix_columns():
    S2 = lambda sid, gid=None: S % sid + (f'#gid={gid}' if gid else '')
    return [
      ('클레임·배드리뷰 원본', [(f"{x['chip']} 고객사진 모음", D % x['source_deck']) for x in CFG['series']]),
      ('리뷰 모니터링 시트', [(f"{x['chip']} Series", S2(x['sheet'], 957652957)) for x in CFG['series']]),
      ('클레임 데이터', [('Zendesk Raw Data_2026년', S2(CFG['zendesk_sheet'])), ('최다 인입사유 대시보드', CFG['looker_url'])]),
      ('SIREN', [('26년 SIREN 등록 현황', S2(CFG['siren_sheet'], CFG['siren_gid']))]),
      ('보고서', [(f"원본 보고서 ({CFG['report_code']})", D % CFG['source_deck']), ('Caspi 데이터 포털', 'https://caspilm.spigen.com')]),
    ]

def build_all():
    p = svc.get(presentationId=P).execute()
    ids = [s['objectId'] for s in p['slides']]
    SECTIONS = section_anchors(p); APPENDIX = appendix_columns()
    overview = dict(SECTIONS).get('Overview')
    reqs = []
    # ---- cover: Apple hero only when report.json apple_cover=true (approved 261002 deck keeps the classic orange cover) ----
    if CFG.get('apple_cover', False):
        cov = p['slides'][0]; cid = cov['objectId']
        reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in cov['pageElements'] if e['objectId'].startswith('ap3_')]
        reqs.append(bg(cid, '#FFFFFF'))
        sizes = {e['objectId']: (e['size']['width']['magnitude'], e['size']['height']['magnitude']) for e in cov['pageElements'] if 'size' in e}
        import re as _re
        find = lambda pred: next(e['objectId'] for e in cov['pageElements'] if 'shape' in e and pred(txt(e)))
        eyebrow = find(lambda t: '글로벌CX전략팀' in t)
        title = find(lambda t: t.startswith('GCX Bi-weekly'))
        date = find(lambda t: _re.fullmatch(r'\d{4}\.\d{2}\.\d{2}', t or '') is not None)
        reqs += [move(eyebrow, 110, 100, (500, 22), sizes[eyebrow]), move(title, 60, 120, (600, 60), sizes[title]), move(date, 160, 194, (400, 28), sizes[date])]
        reqs += [{'updateShapeProperties': {'objectId': o, 'shapeProperties': {'contentAlignment': 'MIDDLE'}, 'fields': 'contentAlignment'}} for o in (eyebrow, title, date)]
        reqs += restyle(eyebrow, 12, INK, bold=True) + restyle(title, 40, INK, bold=True) + restyle(date, 20, INK)
        # sub-nav chips
        widths = [max(52, 14 + len(lab)*6.2) for lab, _ in SECTIONS]; gap = 8; x = (720 - (sum(widths) + gap*(len(widths)-1))) / 2
        for i, ((lab, target), w) in enumerate(zip(SECTIONS, widths)):
            reqs += pill(cid, f'ap3_chip{i}', x, 28, w, 18, lab, 'chip', 8, slide=target); x += w + gap
        # hero buttons
        reqs += pill(cid, 'ap3_btn1', 238, 248, 116, 30, 'Overview 보기', 'filled', 11, slide=overview)
        reqs += pill(cid, 'ap3_btn2', 366, 248, 116, 30, 'Appendix 보기', 'outline', 11, slide='apl_appendix', under='#FFFFFF')
    # ---- appendix slide (before the closing slide) ----
    if 'apl_appendix' not in ids:
        reqs.append({'createSlide': {'objectId': 'apl_appendix', 'insertionIndex': len(ids) - 1,
                                     'slideLayoutReference': {'layoutId': 'p12'}}})
    else:
        app = next(s for s in p['slides'] if s['objectId'] == 'apl_appendix')
        reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in app['pageElements'] if e['objectId'].startswith('ap3_')]
    A = 'apl_appendix'
    reqs.append(bg(A, CANVAS))
    reqs += text(A, 'ap3_a_title', 36, 24, 400, 34, 'Appendix', 22, INK, bold=True)
    reqs += hair(A, 'ap3_a_h1', 36, 70, 648)
    colx = [36, 168, 300, 432, 564]
    for ci, (head, items) in enumerate(APPENDIX):
        reqs += text(A, f'ap3_a_c{ci}', colx[ci], 82, 126, 16, head, 9, INK, bold=True)
        for ii, (lab, url) in enumerate(items):
            reqs += text(A, f'ap3_a_c{ci}_{ii}', colx[ci], 104 + ii*19, 126, 16, lab, 8.5, LINKGRAY, url=url)
    more = 'ap3_a_more'
    msg = '젠데스크, 아마존 배드리뷰 데이터 접근 권한이 필요하면 Caspi 접근 신청 페이지에서 요청하세요.'
    reqs += text(A, more, 36, 300, 648, 16, msg, 8.5, GRAY)
    s0 = msg.index('Caspi 접근 신청'); s1 = s0 + len('Caspi 접근 신청')
    reqs.append({'updateTextStyle': {'objectId': more, 'style': {'link': {'url': 'https://caspilm.spigen.com/access-request'}, 'underline': True, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb('#0066CC')}}},
                 'fields': 'link,underline,foregroundColor', 'textRange': {'type': 'FIXED_RANGE', 'startIndex': s0, 'endIndex': s1}}})
    reqs += hair(A, 'ap3_a_h2', 36, 326, 648)
    reqs += text(A, 'ap3_a_copy', 36, 334, 220, 16, 'Copyright © 2026 Spigen Inc. 글로벌CX전략팀', 8, GRAY)
    x = 262
    NAV = SECTIONS[:-2] + [('SPIGEN', None)]          # approved footer nav: … | SPIGEN
    for i, (lab, target) in enumerate(NAV):
        w = 8 + len(lab)*4.6
        reqs += text(A, f'ap3_a_nav{i}', x, 334, w + 8, 16, lab, 8, LINKGRAY, slide=target); x += w + 6
        if i < len(NAV) - 1: reqs += text(A, f'ap3_a_bar{i}', x, 334, 12, 16, '|', 8, HAIR); x += 14
    # closing slide stays the classic black Spigen-logo slide
    svc.batchUpdate(presentationId=P, body={'requests': reqs}).execute()
    print('requests', len(reqs))

if __name__ == '__main__':
    build_all()
