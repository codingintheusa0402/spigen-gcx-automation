"""Wrap mail_sections.html (from mail_html.py) into the full bi-weekly mail -> mail.html. Edit MAIL below per period."""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__))
from cfg import CFG
MAIL = json.load(open(os.path.join(HERE, 'mail.json')))
F = "-apple-system,BlinkMacSystemFont,'SF Pro Text','Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',Arial,sans-serif"
INK, GRAY, BLUE, RED, CANVAS = '#1D1D1F', '#6E6E73', '#0071E3', '#FF3B30', '#F5F5F7'
LINK = f"https://docs.google.com/presentation/d/{CFG['apple_deck']}/edit?usp=sharing"
def tile(inner, pad='22px 26px'): return f'<table width="100%" cellspacing="0" style="background:#fff;border-radius:18px;margin:0 0 14px"><tr><td style="padding:{pad}">{inner}</td></tr></table>'
def label(t): return f'<div style="font-weight:600;font-size:12px;color:{GRAY};margin:0 0 8px">{t}</div>'
def items(rows): return ''.join(f'<div style="font-size:14px;padding:4px 0">• {a} <b style="color:{RED}">{b}</b></div>' for a, b in rows)
sec = open(os.path.join(HERE, 'mail_sections.html')).read()
html = (f'<div style="background:{CANVAS};padding:24px 0;font-family:{F};color:{INK}"><table align="center" width="680" cellspacing="0" style="width:680px;max-width:100%"><tr><td style="padding:0 14px">'
  + tile(f'<div style="font-weight:600;font-size:12px;color:{GRAY}">GCX Bi-weekly Report · 글로벌CX전략팀</div>'
         f'<div style="font-weight:700;font-size:30px;margin:8px 0 6px;line-height:1.15"><a href="{LINK}" style="color:{INK};text-decoration:none">{MAIL["title"]}</a></div>'
         f'<div style="font-size:14px;color:{GRAY};margin:0 0 18px">{MAIL["subtitle"]}</div>'
         f'<a href="{LINK}" style="display:inline-block;background:{BLUE};color:#fff;font-weight:600;font-size:14px;text-decoration:none;padding:11px 22px;border-radius:980px">리포트 슬라이드 열기 &rsaquo;</a>', '28px 28px')
  + tile(f'<div style="font-size:14px;line-height:1.7">안녕하세요.<br>글로벌CX전략팀 김지우입니다.<br><br>{MAIL["intro"]}</div>')
  + tile(label('이번 모니터링 기간 중 상당량 인입된 VOC') + items(MAIL['voc']))
  + tile(label('금주 신규 SIREN 등록') + items(MAIL['siren']))
  + sec
  + tile(f'<div style="font-size:14px;line-height:1.7">자세한 내용은 상단 <a href="{LINK}" style="color:#0066CC;text-decoration:none;font-weight:600">리포트 슬라이드</a>에서 확인을 부탁 드립니다. 감사합니다!</div>')
  + '</td></tr></table></div>'
  + '<div style="font-family:Arial,sans-serif;font-size:13px;color:#222;margin-top:12px">-- <br><b>김지우</b> 글로벌CX전략팀 | 담당 | <b>M</b> +82.10.2872.2066 | <b>E</b> <a href="mailto:kjw@spigen.com">kjw@spigen.com</a><br>'
    '<img src="https://ci3.googleusercontent.com/mail-sig/AIorK4w-54OvB6nZ_GMuFrUspHQCAearahPbswha8CuA--8GlqZa_pDQBfoPfmf8VV4vi_6D3v88AkK-1htz"></div>')
open(os.path.join(HERE, 'mail.html'), 'w').write(html)
open(os.path.join(HERE, 'mail_preview.html'), 'w').write('<!doctype html><meta charset="utf-8"><body style="margin:0">' + html)
print('mail.html', len(html.encode()), 'bytes')
