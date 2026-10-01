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

SECTIONS = [('Overview', 'g3f2ca1aded2_0_19'), ('Galaxy Z8', 'SLIDES_API939767698_0'), ('Pixel 11', 'SLIDES_API939767698_775'),
            ('iPhone 18', 'SLIDES_API939767698_999'), ('SIREN', 'g3f2ca1aded2_0_340'), ('Appendix', 'apl_appendix')]
D = 'https://docs.google.com/presentation/d/%s/edit'
S = 'https://docs.google.com/spreadsheets/d/%s/edit'
APPENDIX = [
  ('클레임·배드리뷰 원본', [('Galaxy Z8 고객사진 모음', D % '1VC5WAoiufinAPz9bPn1OrBnAef9JkDZEZxlAGF6DDho'),
                       ('Pixel 11 고객사진 모음', D % '1JJKzzBnm9no89mocr6Xqzwqgz8YWoiU5S44Em7gJYSc'),
                       ('iPhone 18 고객사진 모음', D % '1uuHcoTZxxLYlxdaHb0KFUBbI2hEPMvkOV0cL8ELU9dU')]),
  ('리뷰 모니터링 시트', [('Galaxy Z8 Series', S % '19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4' + '#gid=957652957'),
                     ('Pixel 11 Series', S % '12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI' + '#gid=957652957'),
                     ('iPhone 18 Series', S % '1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU' + '#gid=957652957')]),
  ('클레임 데이터', [('Zendesk Raw Data_2026년', S % '1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc'),
                ('최다 인입사유 대시보드', 'https://lookerstudio.google.com/u/0/reporting/654b75ad-c824-4ab2-ac9e-9d9e0f30aa35/page/O36tF')]),
  ('SIREN', [('26년 SIREN 등록 현황', S % '15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI' + '#gid=1840076165')]),
  ('보고서', [('원본 보고서 (261002)', D % '12NxCxbW3z0fH1KKEVzX_uqBGH_APlkZpBWdPgET_aCk'),
           ('Caspi 데이터 포털', 'https://caspilm.spigen.com')]),
]

def build_all():
    p = svc.get(presentationId=P).execute()
    ids = [s['objectId'] for s in p['slides']]
    reqs = []
    # ---- cover: bright hero ----
    cov = p['slides'][0]; cid = cov['objectId']
    reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in cov['pageElements'] if e['objectId'].startswith('ap3_')]
    reqs.append(bg(cid, '#FFFFFF'))
    sizes = {e['objectId']: (e['size']['width']['magnitude'], e['size']['height']['magnitude']) for e in cov['pageElements'] if 'size' in e}
    eyebrow, title, date = 'g3aa52d0de84_0_1', 'g3aa9ffb5090_2_95', 'g3aa52d0de84_0_2'
    reqs += [move(eyebrow, 110, 100, (500, 22), sizes[eyebrow]), move(title, 60, 120, (600, 60), sizes[title]), move(date, 160, 194, (400, 28), sizes[date])]
    reqs += [{'updateShapeProperties': {'objectId': o, 'shapeProperties': {'contentAlignment': 'MIDDLE'}, 'fields': 'contentAlignment'}} for o in (eyebrow, title, date)]
    reqs += restyle(eyebrow, 12, INK, bold=True) + restyle(title, 40, INK, bold=True) + restyle(date, 20, INK)
    # sub-nav chips
    widths = [64, 70, 62, 66, 52, 66]; gap = 8; x = (720 - (sum(widths) + gap*(len(widths)-1))) / 2
    for i, ((lab, target), w) in enumerate(zip(SECTIONS, widths)):
        reqs += pill(cid, f'ap3_chip{i}', x, 28, w, 18, lab, 'chip', 8, slide=target); x += w + gap
    # hero buttons
    reqs += pill(cid, 'ap3_btn1', 238, 248, 116, 30, 'Overview 보기', 'filled', 11, slide='g3f2ca1aded2_0_19')
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
    msg = '데이터 접근 권한이 필요하면 Caspi 접근 신청 페이지에서 요청하세요. 문의: 글로벌CX전략팀 김지우.'
    reqs += text(A, more, 36, 300, 648, 16, msg, 8.5, GRAY)
    s0 = msg.index('Caspi 접근 신청'); s1 = s0 + len('Caspi 접근 신청')
    reqs.append({'updateTextStyle': {'objectId': more, 'style': {'link': {'url': 'https://caspilm.spigen.com/access-request'}, 'underline': True, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb('#0066CC')}}},
                 'fields': 'link,underline,foregroundColor', 'textRange': {'type': 'FIXED_RANGE', 'startIndex': s0, 'endIndex': s1}}})
    reqs += hair(A, 'ap3_a_h2', 36, 326, 648)
    reqs += text(A, 'ap3_a_copy', 36, 334, 220, 16, 'Copyright © 2026 Spigen Inc. 글로벌CX전략팀', 8, GRAY)
    x = 262
    for i, (lab, target) in enumerate(SECTIONS[:5]):
        w = 8 + len(lab)*4.6
        reqs += text(A, f'ap3_a_nav{i}', x, 334, w + 8, 16, lab, 8, LINKGRAY, slide=target); x += w + 6
        if i < 4: reqs += text(A, f'ap3_a_bar{i}', x, 334, 12, 16, '|', 8, HAIR); x += 14
    reqs += text(A, 'ap3_a_loc', 600, 334, 84, 16, 'Korea', 8, LINKGRAY, align='END')
    # ---- closing slide: bright ----
    reqs.append(bg(ids[-1], '#FFFFFF'))
    svc.batchUpdate(presentationId=P, body={'requests': reqs}).execute()
    print('requests', len(reqs))

if __name__ == '__main__':
    build_all()
