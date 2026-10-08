"""Step 12 — summary deck 'GCX SIREN <상/하반기> 등록현황 및 VOC 점검' in the Apple Bi-weekly design system
(bi-weekly-builder skill §7 design: #F5F5F7 canvas, white 12pt-radius tiles, capsules, Inter + Noto Sans KR, Material
icons, classic orange cover + black closing). Starts from a Drive copy of the approved Apple Bi-weekly golden deck
(keeps cover/closing/theme), deletes its other slides, then adds: index → 1. Overview dashboard → 1. Overview table →
one card slide per case (sections Case / SDA·New Biz / Power Accessories / 기타) → Appendix.
Encodes every user/manager rule from 2026-10-02 (see SKILL.md 'Approved design rules').
Usage:
  python3 12_build_summary.py                 # new deck (Drive copy of the golden deck)
  python3 12_build_summary.py --deck <id>     # rebuild in place (keeps the URL) — ONLY if the user has not hand-edited it
"""
import sys, json, os, time, datetime, argparse, common  # noqa
from rate_pipeline import creds
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
import apple_cards as AC
from apple_cards import rtile, rgb, E
from apple_v3 import pill, hair, bg, box, text
from apple_type import latin_ranges, is_hangul

ap_ = argparse.ArgumentParser(); ap_.add_argument('--deck'); A = ap_.parse_args()
TODAY = datetime.date.today(); Y = TODAY.year; HALF = common.half_label(TODAY)
TITLE = f'GCX SIREN {HALF} 등록현황 및 VOC 점검'
INK, GRAY, GRAY2, BLUE, LINK, RED, CANVAS = '#1D1D1F', '#6E6E73', '#86868B', '#0071E3', '#0066CC', '#FF3B30', '#F5F5F7'
ICONS = '/Users/kevinkim/Desktop/GCX/GAS_Operations/Bi-Weekly/tools/apple_theme/icons'
cr = creds(); slides = build('slides', 'v1', credentials=cr).presentations(); drive = build('drive', 'v3', credentials=cr)
cases = json.load(open('cases.json')); by_n = {c['no']: c for c in cases}; NC = len(cases)
SEC_DEF = [('Case', ('Case',), 'iphone'), ('SDA · New Biz', ('SDA', 'New Biz'), 'bag'), ('Power Accessories', ('PAcc.',), 'box')]
SECTIONS = [(name, [c['no'] for c in cases if c['div'] in divs], ic) for name, divs, ic in SEC_DEF]
rest = [c['no'] for c in cases if not any(c['div'] in d for _, d, _ in SEC_DEF)]
if rest: SECTIONS.append(('기타', rest, 'category'))
SECTIONS = [s for s in SECTIONS if s[1]]

# ---------- assets (icons + charts) -> public Drive files
state = json.load(open('assets.json')) if os.path.exists('assets.json') else {}
def asset(path):
    key = f'{path}:{os.path.getmtime(path)}'
    if key in state: return state[key]
    if 'folder' not in state:
        state['folder'] = drive.files().create(body={'name': f'SIREN_Sweep_assets_{TODAY:%y%m%d}', 'mimeType': 'application/vnd.google-apps.folder'}, fields='id').execute()['id']
    f = drive.files().create(body={'name': os.path.basename(path), 'parents': [state['folder']]}, media_body=MediaFileUpload(path, mimetype='image/png'), fields='id').execute()
    drive.permissions().create(fileId=f['id'], body={'type': 'anyone', 'role': 'reader'}).execute()
    state[key] = f'https://drive.google.com/uc?export=view&id={f["id"]}'; json.dump(state, open('assets.json', 'w'), indent=1)
    return state[key]
ICON = {k: asset(f'{ICONS}/{k}.png') for k in ['insights', 'iphone', 'bag', 'box', 'doc', 'globe', 'category', 'donut', 'report', 'sheet']}
CH = {c['no']: asset(os.path.abspath(f"m{c['n']:02d}.png")) for c in cases}
OVER = asset(os.path.abspath('overview.png'))

# ---------- deck
if A.deck:
    pid = A.deck
else:
    pid = drive.files().copy(fileId=common.GOLDEN_APPLE_DECK, body={'name': TITLE}, fields='id').execute()['id']
drive.files().update(fileId=pid, body={'name': TITLE}).execute()
p = slides.get(presentationId=pid).execute()
cover = p['slides'][0]['objectId']
reqs = [{'deleteObject': {'objectId': s['objectId']}} for s in p['slides'][1:-1]]
cover_txt = ''.join(x.get('textRun', {}).get('content', '') for e in p['slides'][0]['pageElements'] for x in e.get('shape', {}).get('text', {}).get('textElements', []))
for old in ('GCX Bi-weekly Report',):
    if old in cover_txt:
        reqs.append({'replaceAllText': {'containsText': {'text': old, 'matchCase': True}, 'replaceText': TITLE, 'pageObjectIds': [cover]}})
K = [0]; order = []
def oid(p_): K[0] += 1; return f'sw_{p_}_{K[0]:04d}'
def new_slide(color):
    sid = oid('slide'); reqs.append({'createSlide': {'objectId': sid, 'insertionIndex': len(order) + 1, 'slideLayoutReference': {'predefinedLayout': 'BLANK'}}})
    reqs.append(bg(sid, color)); order.append(sid); return sid
def img(sid, url, x, y, w, h):
    o = oid('img'); reqs.append({'createImage': {'objectId': o, 'url': url, 'elementProperties': {'pageObjectId': sid,
        'size': {'width': {'magnitude': w*E, 'unit': 'EMU'}, 'height': {'magnitude': h*E, 'unit': 'EMU'}},
        'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': x*E, 'translateY': y*E, 'unit': 'EMU'}}}}); return o
def T(sid, x, y, w, h, s, size, color=INK, bold=False, align='START', url=None, slide=None, valign='TOP'):
    o = oid('t'); r = text(sid, o, x, y, w, h, s, size, color, bold, align, url, slide)
    r[-1]['updateShapeProperties']['shapeProperties']['contentAlignment'] = valign
    reqs.extend(r); return o
def tile(sid, x, y, w, h, color='#FFFFFF'):
    AC.R = min(12, w / 4, h / 4)          # 12pt radius; small thumbnail tiles scale it down (rtile breaks below 24pt)
    r, _ = rtile(sid, oid('tile'), x, y, w, h, color); reqs.extend(r); AC.R = 12
def title(sid, s, qual=None):
    o = T(sid, 36, 16, 648, 28, s + (' ' + qual if qual else ''), 18, INK, True, valign='MIDDLE')
    if qual:
        reqs.append({'updateTextStyle': {'objectId': o, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': len(s) + 1, 'endIndex': len(s) + 1 + len(qual)},
                     'style': {'fontSize': {'magnitude': 9, 'unit': 'PT'}, 'bold': False, 'foregroundColor': {'opaqueColor': {'rgbColor': rgb(GRAY2)}}}, 'fields': 'fontSize,bold,foregroundColor'}})
def short_name(c): return c['name'].replace(' 시리즈용 ', ' · ')
def full_name(c): return c['name'].replace(' 시리즈용 ', ' ')
def fmt_rate(c): return f"{c['eu_rate']:.2f}%" if c['eu_rate'] is not None else '-'
def hot(c): return c['eu_rate'] is not None and c['eu_rate'] >= 2
TOTAL = sum(c['total'] for c in cases); CL = sum(c['claims'] for c in cases); RV = sum(c['reviews'] for c in cases)
REGS = [c for c in cases if c['reg']]
SCOPE = 'Screen Protector 제외'           # manager rule: never 'non-SP' — spell it out + footnote why

# ---------- index (white)
idx = new_slide('#FFFFFF')
blocks = [('01', 'Overview', [f'{SCOPE} SIREN {NC}건 대시보드', '이슈별 VOC · 판매량(Amazon EU) · VOC율 표'], 'insights')]
for i, (sec, ns, ic) in enumerate(SECTIONS, 2):
    blocks.append((f'{i:02d}', f'{sec}\nSIREN {len(ns)}건', [f"{by_n[n]['name'].split(' 시리즈용')[0]} {by_n[n]['defect']}" for n in ns], ic))
blocks.append((f'{len(blocks) + 1:02d}', 'Appendix', ['SIREN 슬라이드 · 데이터 원본 링크'], 'doc'))
cols = [(70, 70), (285, 70), (500, 70), (70, 225), (285, 225), (500, 225)]
for (num, head, items, ic), (x, y) in zip(blocks, cols):
    T(idx, x, y, 50, 26, num, 20, BLUE, valign='MIDDLE'); img(idx, ICON[ic], x + 52, y + 5, 15, 15)
    T(idx, x, y + 32, 190, 40, head, 13, INK)
    T(idx, x, y + 76, 190, 70, '\n'.join(f'{k}) {t}' for k, t in enumerate(items, 1)), 6.5, INK)
for x in (265, 480, 695):
    reqs.extend(hair(idx, oid('ln'), x, 70, 0.1)); reqs[-2]['createLine']['elementProperties']['size'] = {'width': {'magnitude': 1, 'unit': 'EMU'}, 'height': {'magnitude': 320*E, 'unit': 'EMU'}}

# ---------- 1. Overview dashboard (white)
ov = new_slide('#FFFFFF')
title(ov, '1. Overview', f'({SCOPE} · {Y}.01.01~{TODAY:%Y.%m.%d})')
T(ov, 72, 58, 250, 14, '작성 취지', 8.5, INK, True)
T(ov, 72, 71, 310, 32, f'SIREN 등록 누락 방지를 위한 정기 점검으로,\n동일 라인업 × 동일 불량 클레임+배드리뷰 5건 이상 이슈 {NC}건과 기등록 이슈의\n'
                       '등록 이후 VOC 지속 여부를 함께 확인 (Screen Protector는 별도 팀 보고로 제외)', 6.5, GRAY)
if REGS: T(ov, 72, 106, 300, 14, f"기등록 SIREN {len(REGS)}건 ({'·'.join('#' + str(c['no']) for c in REGS)}) — 등록 후에도 VOC 지속", 7, RED, True)
T(ov, 72, 128, 300, 14, f'추적 이슈 {NC}건의 누적 VOC (클레임 + 배드리뷰)', 8.5, INK, True)
T(ov, 72, 142, 200, 34, f'{TOTAL:,}', 26, INK, valign='MIDDLE')
T(ov, 72, 176, 260, 12, f'클레임 {CL:,} · 배드리뷰 {RV:,}', 7.5, BLUE)
T(ov, 380, 58, 300, 14, f'판매 대비 VOC율 TOP3 (Amazon EU 판매량 기준, {Y})', 8.5, INK, True)
for k, c in enumerate(sorted([c for c in cases if c['eu_rate']], key=lambda c: -c['eu_rate'])[:3]):
    x = 380 + k * 103; tile(ov, x, 76, 97, 58, CANVAS)
    o = T(ov, x + 7, 80, 40, 16, f"{k + 1}{['st', 'nd', 'rd'][k]}", 11, BLUE)
    reqs.append({'updateTextStyle': {'objectId': o, 'textRange': {'type': 'FIXED_RANGE', 'startIndex': 1, 'endIndex': 3}, 'style': {'baselineOffset': 'SUPERSCRIPT'}, 'fields': 'baselineOffset'}})
    T(ov, x + 7, 94, 88, 16, fmt_rate(c), 11.5, RED if hot(c) else INK, True)
    T(ov, x + 3, 111, 94, 20, f"{full_name(c)} {c['defect']}", 6, GRAY)          # full product name, no #N (user rule)
T(ov, 72, 196, 120, 12, '이슈별 VOC', 8.5, INK, True)
for k, (col, lab) in enumerate([(BLUE, '클레임'), ('#64D2FF', '배드리뷰')]):       # legend squares (not text in the heading)
    lx = 560 + k * 58; sq = oid('lg'); reqs.append(box(ov, sq, 'RECTANGLE', lx, 199, 7, 7))
    reqs.append({'updateShapeProperties': {'objectId': sq, 'shapeProperties': {'shapeBackgroundFill': {'solidFill': {'color': {'rgbColor': rgb(col)}}}, 'outline': {'propertyState': 'NOT_RENDERED'}}, 'fields': 'shapeBackgroundFill.solidFill.color,outline.propertyState'}})
    T(ov, lx + 3, 195, 54, 14, lab, 7, GRAY, valign='MIDDLE')
CH_H = 190; CH_W = CH_H * 620 / 215                                                  # overview.png is 620x215pt
chart = img(ov, OVER, 360 - CH_W / 2, 213, CH_W, CH_H)
reqs.append({'updatePageElementsZOrder': {'pageElementObjectIds': [chart], 'operation': 'SEND_TO_BACK'}})

# ---------- 1. Overview table (canvas)
tb = new_slide(CANVAS)
title(tb, '1. Overview', f'{SCOPE} SIREN VOC / 판매량 (Amazon EU)')
tile(tb, 36, 52, 648, 318)
COLS = [('No.', 36, 22), ('제품', 58, 182), ('불량 유형', 240, 54), ('생산지', 294, 66), ('합계', 360, 26), ('클레임', 386, 28), ('리뷰', 414, 24),
        ('판매량(EU)', 438, 40), ('VOC율(EU)', 478, 40), ('기간', 518, 58), ('기등록', 578, 50), ('슬라이드', 630, 46)]
LEFT = ('제품', '불량 유형', '생산지', '기간'); y0, RH = 60, 25.5
# text boxes are widened by 7pt each side (Slides' fixed ~7.2pt insets) so text uses the whole column (no wraps)
for h, x, w in COLS: T(tb, x - 7, y0, w + 14, 18, h, 6.5, GRAY2, align='START' if h in LEFT else 'CENTER', valign='MIDDLE')
reqs.extend(hair(tb, oid('ln'), 40, y0 + 20, 640))
for r, c in enumerate(cases):
    y = y0 + 22 + r * RH
    vals = {'No.': f"#{c['no']}", '불량 유형': c['defect'], '생산지': f"{', '.join(c['maker']) or '-'} · {'/'.join(c['origin']) or '-'}",
            '합계': str(c['total']), '클레임': str(c['claims']), '리뷰': str(c['reviews']), '판매량(EU)': f"{c['eu_units']:,}", 'VOC율(EU)': fmt_rate(c),
            '기간': f"{c['first'][2:].replace('-', '.')}~{c['last'][5:].replace('-', '.')}"}
    for h, x, w in COLS:
        if h == '제품':
            tile(tb, x, y + 3, 19, 19, CANVAS)
            if c['img']: img(tb, c['img'], x + 1.5, y + 4.5, 16, 16)
            T(tb, x + 17, y, w - 10, RH, short_name(c), 6.5, INK, True, valign='MIDDLE'); continue
        if h == '기등록':                                  # red capsule -> earlier SIREN slide/sheet link, date under it
            if c['reg']:
                reqs.extend(pill(tb, oid('pill'), x + 4, y + 2.5, w - 8, 12, '기등록', 'red', 6, url=c['reg_url']))
                T(tb, x - 7, y + 14.5, w + 14, 10, c['reg_date'], 5.5, GRAY2, align='CENTER', valign='MIDDLE')
            else: T(tb, x - 7, y, w + 14, RH, '-', 6.5, INK, align='CENTER', valign='MIDDLE')
            continue
        if h == '슬라이드':
            T(tb, x - 7, y, w + 14, RH, '보기', 6.5, LINK, align='CENTER', url=c['deck'], valign='MIDDLE'); continue
        red = h == 'VOC율(EU)' and hot(c)
        T(tb, x - 7, y, w + 14, RH, vals[h], 6.5, RED if red else INK, h == '합계' or red, align='START' if h in LEFT else 'CENTER', valign='MIDDLE')
    if r < NC - 1: reqs.extend(hair(tb, oid('ln'), 40, y + RH, 640))
T(tb, 36, 372, 648, 30, '※ Screen Protector(글라스) 제품은 별도 팀에 보고되어 본 자료에서 제외 · 생산지 = SKU_Master 생산업체·원산지\n'
  f'판매량(EU) = Caspi Amazon 주문, Amazon EU(DE·FR·IT·ES·NL·SE·BE·IE·PL·UK)만, {Y}, 취소 제외 — JP·US·IN 미포함 · VOC율(EU) = EU 국가 클레임+배드리뷰 ÷ EU 판매량, 2% 이상 빨강 · 합계/클레임/리뷰는 전 국가', 5.5, GRAY2)

# ---------- case slides
CASE_SLIDE = {}
for si, (sec, ns, ic) in enumerate(SECTIONS, 2):
    for n in ns:
        c = by_n[n]; sid = new_slide(CANVAS); CASE_SLIDE[n] = sid
        title(sid, f'{si}. {sec} SIREN', f'#{n}')
        tile(sid, 36, 52, 410, 330); tile(sid, 50, 64, 38, 38, CANVAS)
        if c['img']: img(sid, c['img'], 53, 67, 32, 32)
        T(sid, 96, 64, 330, 18, short_name(c), 11.5, INK, True, valign='MIDDLE')
        more = f" 외 {len(c['skus']) - 2}종" if len(c['skus']) > 2 else ''
        T(sid, 96, 84, 340, 14, f"불량 유형 · {c['defect']}   |   SKU {', '.join(c['skus'][:2])}{more}", 7, GRAY, valign='MIDDLE')
        for k, u in enumerate(c['photos'][:3]):
            tile(sid, 50 + k * 130, 110, 122, 150, CANVAS); img(sid, u, 52 + k * 130, 112, 118, 146)
        if not c['photos']: T(sid, 50, 110, 380, 150, '첨부 사진 없음', 8, GRAY2, align='CENTER', valign='MIDDLE')
        T(sid, 50, 268, 200, 12, '최근 VOC', 7.5, GRAY)
        yy = 282
        for v in c['voc'][:2]:                       # each recent VOC has a link button to its original (ticket / review)
            T(sid, 50, yy, 300, 12, f"{v['kind']} · {v['date']} · {v['country']}", 6.5, GRAY2, valign='MIDDLE')
            reqs.extend(pill(sid, oid('pill'), 372, yy - 1, 60, 14, '원문 보기', 'outline', 6.5, url=v['link']))
            T(sid, 50, yy + 14, 382, 32, v['text'][:150], 8, INK); yy += 48
        tile(sid, 456, 52, 228, 72)
        T(sid, 468, 60, 100, 12, 'VOC 합계', 7.5, GRAY)
        T(sid, 468, 74, 120, 28, f"{c['total']}건", 20, INK, True, valign='MIDDLE')
        T(sid, 468, 104, 200, 12, f"클레임 {c['claims']} · 배드리뷰 {c['reviews']}", 7.5, GRAY)
        reqs.extend(pill(sid, oid('pill'), 584, 60, 90, 18, 'SIREN 슬라이드 열기', 'filled', 7.5, url=c['deck']))
        if c['reg']: reqs.extend(pill(sid, oid('pill'), 584, 84, 90, 15, f"기등록 {c['reg']}", 'red', 6, url=c['reg_url']))
        small = [('bag', '판매량 (Amazon EU)', f"{c['eu_units']:,}개", INK), ('donut', 'VOC율 (EU)', fmt_rate(c), RED if hot(c) else INK),
                 ('category', '인입 기간', f"{c['first'][2:].replace('-', '.')}~{c['last'][5:].replace('-', '.')}", INK),
                 ('globe', '국가 TOP', ' · '.join(f'{k} {v}' for k, v in c['countries'][:3]), INK)]
        for k, (ic_, lab, val, colr) in enumerate(small):
            x = 456 + (k % 2) * 118; y = 132 + (k // 2) * 60
            tile(sid, x, y, 110, 52); img(sid, ICON[ic_], x + 10, y + 9, 9, 9)
            T(sid, x + 15, y + 7, 98, 12, lab, 6.5, GRAY, valign='MIDDLE')
            T(sid, x + 3, y + 24, 107, 22, val, 11.5 if len(val) < 12 else 9, colr, True, valign='MIDDLE')
        tile(sid, 456, 252, 228, 130); img(sid, ICON['report'], 466, 261, 9, 9)
        T(sid, 478, 259, 180, 12, '월별 인입 추이 (클레임+배드리뷰)', 6.5, GRAY, valign='MIDDLE')
        img(sid, CH[n], 462, 276, 216, 85)
        T(sid, 466, 362, 210, 14, '■ 최근 월 = 파랑 · 점선 = 추세', 6, GRAY2)

# ---------- Appendix (white)
ap = new_slide('#FFFFFF'); title(ap, 'Appendix')
T(ap, 36, 50, 648, 14, 'SIREN 슬라이드 (siren-report 스킬 생성, 클레임·배드리뷰 1건당 1슬라이드)', 8.5, INK, True)
for k, c in enumerate(cases):
    T(ap, 36 + (k % 2) * 324, 70 + (k // 2) * 22, 316, 18, f"#{c['no']} {short_name(c)} — {c['defect']}", 7.5, LINK, url=c['deck'], valign='MIDDLE')
ya = 70 + ((NC + 1) // 2) * 22 + 6
reqs.extend(hair(ap, oid('ln'), 36, ya, 648)); T(ap, 36, ya + 8, 648, 14, '데이터 원본', 8.5, INK, True)
src = [('Zendesk 전체문의', 'https://docs.google.com/spreadsheets/d/1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc/edit'),
       ('아마존 배드리뷰 (SC 시트)', 'https://docs.google.com/spreadsheets/d/1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw/edit'),
       ('GCX SIREN 등록 현황', f'https://docs.google.com/spreadsheets/d/{common.GCX_KPI_26[0]}/edit'),
       ('SIREN 등록 현황 (영업)', f'https://docs.google.com/spreadsheets/d/{common.SALES_SIREN[0]}/edit?gid=1510756422#gid=1510756422'),
       ('CQ Emergency Net (기등록 확인)', 'https://docs.google.com/spreadsheets/d/137K4hpNfHoyb6PEb64gO7Wxi5b3CPlnMETQt-bbPKKE/edit'),
       ('SKU_Master (생산지)', f'https://docs.google.com/spreadsheets/d/{common.SKU_MASTER}/edit')]
for k, (lab, u) in enumerate(src):
    img(ap, ICON['sheet'], 36 + (k % 3) * 216, ya + 32 + (k // 3) * 24, 10, 10)
    T(ap, 50 + (k % 3) * 216, ya + 28 + (k // 3) * 24, 200, 18, lab, 7.5, LINK, url=u, valign='MIDDLE')
reqs.extend(hair(ap, oid('ln'), 36, ya + 84, 648))
T(ap, 36, ya + 92, 648, 30, f'젠데스크, 아마존 배드리뷰 데이터 접근 권한이 필요하면 Caspi 접근 신청 페이지에서 요청하세요.\nCopyright © {Y} Spigen Inc. 글로벌CX전략팀. All rights reserved.', 6.5, GRAY)
x = 36
for k, w in enumerate(['Overview'] + [s[0] for s in SECTIONS] + ['SPIGEN']):
    url = 'https://www.spigen.com/' if w == 'SPIGEN' else None
    sl = None if url else (ov if w == 'Overview' else CASE_SLIDE[SECTIONS[k - 1][1][0]])
    ww = sum(7 * (1.0 if ord(ch) > 0x2E80 else 0.62) for ch in w) + 20
    T(ap, x, ya + 132, ww, 14, w, 7, INK, url=url, slide=sl, valign='MIDDLE'); x += ww + 6

def send(rs):
    """Chunks of 300; retries 429 / transient image-fetch failures; a chunk that keeps failing on an image is split
    (non-image requests first, then each image alone; a broken image is skipped and reported)."""
    for i in range(0, len(rs), 300):
        chunk = rs[i:i + 300]
        for k in range(4):
            try:
                slides.batchUpdate(presentationId=pid, body={'requests': chunk}).execute(); break
            except HttpError as e:
                if e.resp.status in (429, 500, 502, 503) or 'retrieving the image' in str(e):
                    time.sleep(20 if e.resp.status == 429 else 6); continue
                raise
        else:
            imgs = [r for r in chunk if 'createImage' in r]; ids = {r['createImage']['objectId'] for r in imgs}
            slides.batchUpdate(presentationId=pid, body={'requests': [r for r in chunk if r not in imgs and
                               r.get('updatePageElementsZOrder', {}).get('pageElementObjectIds', [''])[0] not in ids]}).execute()
            for r in imgs:
                try: slides.batchUpdate(presentationId=pid, body={'requests': [r]}).execute()
                except HttpError as e: print('SKIPPED image', r['createImage']['url'], str(e)[:120])
send(reqs)

# ---------- one font pass (cover/closing keep brand fonts): Latin -> Inter, Hangul -> Noto Sans KR, weight from the run's bold
p = slides.get(presentationId=pid).execute(); fr = []
for s in p['slides'][1:-1]:
    for e in s.get('pageElements', []):
        tx = e.get('shape', {}).get('text')
        if not tx: continue
        for te in tx['textElements']:
            r = te.get('textRun')
            if not r: continue
            a0 = te.get('startIndex', 0); c = r['content']; b = bool(r.get('style', {}).get('bold')); i = 0
            while i < len(c):
                if not c[i].strip(): i += 1; continue
                hg = is_hangul(c[i]); j = i
                while j < len(c) and (is_hangul(c[j]) == hg or not c[j].strip()) and c[j] != '\n': j += 1
                seg_end = j
                while seg_end > i and not c[seg_end - 1].strip(): seg_end -= 1
                fam, wt = ('Noto Sans KR', 700 if b else 400) if hg else ('Inter', 600 if b else 400)
                fr.append({'updateTextStyle': {'objectId': e['objectId'], 'textRange': {'type': 'FIXED_RANGE', 'startIndex': a0 + i, 'endIndex': a0 + seg_end},
                           'style': {'weightedFontFamily': {'fontFamily': fam, 'weight': wt}}, 'fields': 'weightedFontFamily'}})
                i = j
for i in range(0, len(fr), 400): slides.batchUpdate(presentationId=pid, body={'requests': fr[i:i + 400]}).execute()
open('summary_deck.txt', 'w').write(pid)
print(f'https://docs.google.com/presentation/d/{pid}/edit', len(p['slides']), 'slides')
