import sys, time
sys.path.insert(0, __import__('os').path.dirname(__file__))
from apple_cards import svc, P, E, rgb, rtile
from apple_v3 import pill, text, hair, bg, INK, GRAY, GRAY2, CANVAS, LINKGRAY, BLUE

MASTER = 'simple-light-2'
def send(reqs):
    for i in range(0, len(reqs), 300):
        for _ in range(8):
            try: svc.batchUpdate(presentationId=P, body={'requests': reqs[i:i+300]}).execute(); break
            except Exception as ex:
                if '429' in str(ex): time.sleep(20); continue
                raise

def place_ph(e, x, y, w, h):
    sw, sh = e['size']['width']['magnitude'], e['size']['height']['magnitude']
    return {'updatePageElementTransform': {'objectId': e['objectId'], 'applyMode': 'ABSOLUTE', 'transform': {
        'scaleX': w*E/sw, 'scaleY': h*E/sh, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}
OPTIONAL = []   # text-style requests that only apply when the placeholder carries text
def style_ph(e, size, color, bold=False, align='START', valign='MIDDLE', font='Noto Sans KR'):
    oid = e['objectId']; out = [{'updateShapeProperties': {'objectId': oid, 'shapeProperties': {'contentAlignment': valign}, 'fields': 'contentAlignment'}}]
    if True:
        OPTIONAL.append([{'updateTextStyle': {'objectId': oid, 'textRange': {'type': 'ALL'}, 'style': {'fontFamily': font, 'fontSize': {'magnitude': size, 'unit': 'PT'},
                 'bold': bold, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(color)}}}, 'fields': 'fontFamily,fontSize,bold,foregroundColor'}},
                {'updateParagraphStyle': {'objectId': oid, 'textRange': {'type': 'ALL'}, 'style': {'alignment': align, 'lineSpacing': 100,
                 'spaceAbove': {'magnitude': 0, 'unit': 'PT'}, 'spaceBelow': {'magnitude': 0, 'unit': 'PT'}}, 'fields': 'alignment,lineSpacing,spaceAbove,spaceBelow'}}])
    return out
def tiles(lid, key, boxes, color='#FFFFFF'):
    out, ids = [], []
    for i, (x, y, w, h) in enumerate(boxes):
        r, t = rtile(lid, f'lay_{key}_t{i}', x, y, w, h, color); out += r; ids += t
    return out + [{'updatePageElementsZOrder': {'pageElementObjectIds': ids, 'operation': 'SEND_TO_BACK'}}]

def build():
    p = svc.get(presentationId=P).execute()
    L = {l['objectId']: l for l in p['layouts']}
    reqs = []
    # ---- theme (master): Apple palette + default type ----
    scheme = {'DARK1': '#1D1D1F', 'LIGHT1': '#FFFFFF', 'DARK2': '#6E6E73', 'LIGHT2': '#F5F5F7', 'ACCENT1': '#0071E3', 'ACCENT2': '#FF3B30',
              'ACCENT3': '#64D2FF', 'ACCENT4': '#5E5CE6', 'ACCENT5': '#34C759', 'ACCENT6': '#FF9F0A', 'HYPERLINK': '#0066CC', 'FOLLOWED_HYPERLINK': '#0066CC',
              'TEXT1': '#1D1D1F', 'BACKGROUND1': '#FFFFFF', 'TEXT2': '#6E6E73', 'BACKGROUND2': '#F5F5F7'}
    reqs.append({'updatePageProperties': {'objectId': MASTER, 'pageProperties': {'colorScheme': {'colors': [{'type': k, 'color': rgb(v)} for k, v in scheme.items()]},
                 'pageBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(CANVAS)}}}}, 'fields': 'colorScheme,pageBackgroundFill.solidFill.color'}})
    m = next(x for x in p['masters'] if x['objectId'] == MASTER)
    for e in m['pageElements']:
        ph = e.get('shape', {}).get('placeholder', {}).get('type')
        if ph == 'TITLE': reqs += style_ph(e, 18, INK, bold=True)
        if ph == 'BODY': reqs += style_ph(e, 11, INK, valign='TOP')
        if ph == 'SLIDE_NUMBER': reqs += style_ph(e, 8, GRAY2, align='END')
    # ---- clean previous runs ----
    for lid in ('p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8', 'p9', 'p10', 'p11'):
        reqs += [{'deleteObject': {'objectId': e['objectId']}} for e in L[lid].get('pageElements', []) if e['objectId'].startswith('lay_')]
    def ph(lid, kind, n=0):
        hits = [e for e in L[lid].get('pageElements', []) if e.get('shape', {}).get('placeholder', {}).get('type') == kind]
        return hits[n] if len(hits) > n else None
    def drop(lid, kinds):
        return [{'deleteObject': {'objectId': e['objectId']}} for e in L[lid].get('pageElements', [])
                if e.get('shape', {}).get('placeholder', {}).get('type') in kinds or (e.get('shape', {}).get('placeholder') is None and not e['objectId'].startswith('lay_') and 'shape' in e)]
    def std_title(lid):
        t = ph(lid, 'TITLE'); return [place_ph(t, 36, 16, 648, 34)] + style_ph(t, 18, INK, bold=True) if t else []

    # p2 — Cover
    reqs.append(bg('p2', '#FFFFFF'))
    t, s = ph('p2', 'CENTERED_TITLE'), ph('p2', 'SUBTITLE')
    reqs += [place_ph(t, 60, 120, 600, 60)] + style_ph(t, 40, INK, bold=True, align='CENTER')
    reqs += [place_ph(s, 160, 194, 400, 28)] + style_ph(s, 20, INK, align='CENTER')
    reqs += text('p2', 'lay_cov_eyebrow', 110, 100, 500, 22, '경영지원부문ㅣ사업지원실ㅣ글로벌CX전략팀', 12, INK, bold=True, align='CENTER')
    widths, x = [64, 70, 62, 66, 52, 66], None
    x = (720 - (sum(widths) + 8*5)) / 2
    for i, (lab, w) in enumerate(zip(['Overview', 'Galaxy Z8', 'Pixel 11', 'iPhone 18', 'SIREN', 'Appendix'], widths)):
        reqs += pill('p2', f'lay_cov_chip{i}', x, 28, w, 18, lab, 'chip', 8); x += w + 8
    reqs += pill('p2', 'lay_cov_btn1', 238, 248, 116, 30, 'Overview 보기', 'filled', 11)
    reqs += pill('p2', 'lay_cov_btn2', 366, 248, 116, 30, 'Appendix 보기', 'outline', 11, under='#FFFFFF')
    # p3 — Index (3-column section grid)
    reqs.append(bg('p3', '#FFFFFF')); reqs += drop('p3', ['TITLE'])
    for i, xx in enumerate((240, 432, 624)):
        reqs.append({'createLine': {'objectId': f'lay_idx_l{i}', 'lineCategory': 'STRAIGHT', 'elementProperties': {'pageObjectId': 'p3',
                     'size': {'width': {'magnitude': 0, 'unit': 'EMU'}, 'height': {'magnitude': 286*E, 'unit': 'EMU'}},
                     'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': xx*E, 'translateY': 80*E, 'unit': 'EMU'}}}})
        reqs.append({'updateLineProperties': {'objectId': f'lay_idx_l{i}', 'lineProperties': {'lineFill': {'solidFill': {'color': {'rgbColor': rgb('#E5E5EA')}}}, 'weight': {'magnitude': 0.75, 'unit': 'PT'}}, 'fields': 'lineFill.solidFill.color,weight'}})
    for i, (xx, yy, num) in enumerate([(68, 69, '01'), (254, 69, '02'), (444, 69, '03'), (68, 205, '04'), (254, 205, '05')]):
        reqs += text('p3', f'lay_idx_n{i}', xx, yy, 68, 40, num, 26, BLUE)
    # p4 — Overview dashboard
    reqs.append(bg('p4', '#FFFFFF')); reqs += std_title('p4')
    b = ph('p4', 'BODY'); reqs += [place_ph(b, 124, 57, 230, 90)] + style_ph(b, 9, GRAY, valign='TOP')
    reqs += text('p4', 'lay_ov_h1', 336, 55, 330, 18, '2026 최다 인입사유', 10, INK)
    reqs += tiles('p4', 'ov', [(382, 86, 95, 51), (480, 86, 95, 51), (578, 86, 95, 51)], color='#F5F5F7')
    # p5 — TOP3 gauges
    reqs.append(bg('p5', CANVAS)); reqs += std_title('p5') + drop('p5', ['BODY'])
    reqs += text('p5', 'lay_g_h1', 154, 51, 106, 26, '모델별 TOP3', 11.5, INK, bold=True) + text('p5', 'lay_g_l1', 244, 52, 70, 26, '클레임', 10.5, GRAY2)
    reqs += text('p5', 'lay_g_h2', 154, 225, 106, 26, '인입사유별 TOP3', 11.5, INK, bold=True) + text('p5', 'lay_g_l2', 244, 226, 70, 26, '클레임', 10.5, GRAY2)
    reqs += tiles('p5', 'g', [(x, y, 160, 123) for y in (82, 256) for x in (142, 326, 509)])
    for r, y in enumerate((180, 354)):
        for c, x in enumerate((120, 304, 487)):
            reqs += text('p5', f'lay_g_r{r}{c}', x, y, 26, 33, str(c + 1), 11, GRAY2)
    # p6 — Table (TOP 7)
    reqs.append(bg('p6', CANVAS)); reqs += std_title('p6')
    reqs += tiles('p6', 'tb', [(70, 52, 620, 258)])
    reqs += text('p6', 'lay_tb_note', 82, 372, 600, 26, '출처: ', 6.5, GRAY2)
    # p7 — Claim / Review card
    reqs.append(bg('p7', CANVAS)); reqs += std_title('p7')
    reqs += tiles('p7', 'cd', [(36, 72, 404, 312), (452, 72, 232, 118), (452, 198, 112, 52), (572, 198, 112, 52), (452, 258, 112, 52), (572, 258, 112, 52), (452, 318, 232, 66)])
    b = ph('p7', 'BODY'); reqs += [place_ph(b, 464, 116, 210, 70)] + style_ph(b, 10.5, INK, bold=True, valign='TOP')
    reqs += text('p7', 'lay_cd_src', 476, 79, 120, 18, 'From Amazon', 8, GRAY2) + text('p7', 'lay_cd_head', 464, 98, 210, 18, '리뷰 내용', 9, GRAY)
    reqs += pill('p7', 'lay_cd_btn', 584, 80, 88, 18, '배드리뷰 바로가기', 'filled', 7.5)
    for i, ((gx, gy, gw), lab) in enumerate(zip([(452, 198, 112), (572, 198, 112), (452, 258, 112), (572, 258, 112), (452, 318, 232)],
                                                ['국가', '인입사유', 'ASIN', 'SKU', '아마존 리뷰 평점 / 갯수'])):
        reqs += text('p7', f'lay_cd_lb{i}', gx + 23, gy + 5, gw - 30, 18, lab, 7.5, GRAY2)
    # p8 — SIREN
    reqs.append(bg('p8', CANVAS)); reqs += std_title('p8')
    reqs += tiles('p8', 'sr', [(121, 76, 522, 228), (443, 312, 200, 40)])
    reqs += text('p8', 'lay_sr_lbl', 453, 312, 140, 40, '2026 SIREN Registered by GCX', 8.5, GRAY)
    reqs += pill('p8', 'lay_sr_btn', 293, 319, 136, 26, '26년 SIREN 시트 열기', 'outline', 9, under=CANVAS)
    # p9 — Appendix (apple.com footer)
    reqs.append(bg('p9', CANVAS)); reqs += drop('p9', ['SUBTITLE', 'BODY'])
    reqs += std_title('p9') + hair('p9', 'lay_ap_h1', 36, 70, 648) + hair('p9', 'lay_ap_h2', 36, 326, 648)
    for i, (xx, head) in enumerate(zip([36, 168, 300, 432, 564], ['클레임·배드리뷰 원본', '리뷰 모니터링 시트', '클레임 데이터', 'SIREN', '보고서'])):
        reqs += text('p9', f'lay_ap_c{i}', xx + 13, 80, 120, 20, head, 9, INK, bold=True)
    reqs += text('p9', 'lay_ap_copy', 36, 334, 260, 16, 'Copyright © 2026 Spigen Inc. 글로벌CX전략팀', 8, GRAY)
    # p10 — Closing
    reqs.append(bg('p10', '#FFFFFF'))
    b = ph('p10', 'BODY'); reqs += [place_ph(b, 110, 150, 500, 44)] + style_ph(b, 28, INK, bold=True, align='CENTER', valign='MIDDLE')
    reqs += text('p10', 'lay_end_sub', 110, 194, 500, 20, '글로벌CX전략팀', 11, GRAY, align='CENTER')
    reqs += pill('p10', 'lay_end_btn', 306, 228, 108, 26, '처음으로', 'outline', 10, under='#FFFFFF')
    # p11 — Big-number stat
    reqs.append(bg('p11', CANVAS)); reqs += std_title('p11')
    reqs += tiles('p11', 'bn', [(36, 72, 648, 300)])
    b = ph('p11', 'BODY'); reqs += [place_ph(b, 60, 130, 600, 160)] + style_ph(b, 72, INK, bold=True, align='CENTER', valign='MIDDLE')
    send(reqs); print('layout requests', len(reqs))
    ok = 0
    for r in OPTIONAL:
        try: svc.batchUpdate(presentationId=P, body={'requests': r}).execute(); ok += 1
        except Exception as ex:
            if '429' in str(ex): time.sleep(20)
    print('placeholder styles applied', ok, '/', len(OPTIONAL))

if __name__ == '__main__':
    build()
