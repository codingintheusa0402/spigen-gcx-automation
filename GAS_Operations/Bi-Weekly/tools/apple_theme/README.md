# Apple-style (apple.com / macOS 27 page) restyle of a Bi-Weekly deck copy

Run order on a Drive copy of the live deck (set P in apple_cards.py to the copy's id):
1. apple_restyle.py  — fonts/colors/backgrounds, sidebar chrome removal
2. apple_cards.py    — claim/review cards -> photo tile + spec tiles + pill buttons (roles read from the ORIGINAL deck by element id)
3. apple_tables.py   — TOP-7 / SIREN tables on white tiles, hairlines, blue 보기 links
4. apple_v3.py       — cover hero + section chips, Appendix (apple.com footer style), closing slide
5. GAS runOverviewApple261002 (Code.js) — TOP3 gauges in CHART_THEMES.apple

All scripts are idempotent (they delete their own at_/ap3_ shapes before redrawing).

## Round 3 (2026-10-01)
6. apple_v4.py        — capsule buttons (rect + 2 circles; outline = layered ring), SIREN table alignment + stat tile, closing slide
7. icons_render.py    — renders Material Symbols Rounded (Apache-2.0) icons to PNG; chart.py renders the weekly bar chart
   (SF Pro / Apple SD Gothic Neo, system /usr/bin/python3 with matplotlib). Both were inserted via a one-off GAS
   (base64 blobs, no public files) — re-create it from these PNGs if needed.
8. apple_type.py fix  — slide titles: same box (x36,y16), 18pt bold ink, gray "대상 국가" qualifier; '1.' list bullets -> text;
   restores bold, then Latin runs -> Inter (600 for bold), Hangul stays Noto Sans KR.
9. align_icons.py   — centers each icon on its label's text line (run after any font/label change)
10. apple_layouts.py  — Apple theme on master `simple-light-2` (color scheme + default type) and 10 layouts built on the
    unused stock layouts: p2 Cover, p3 Index, p4 Overview dashboard, p5 TOP3 gauges, p6 Table, p7 Claim/Review card,
    p8 SIREN, p9 Appendix, p10 Closing, p11 Big number. Layout icons were inserted by a one-off GAS (title 'lay_icon').
    Layout names can't be set through the API — rename them in View → Theme builder.
