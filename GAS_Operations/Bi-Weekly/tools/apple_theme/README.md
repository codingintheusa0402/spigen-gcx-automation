# Apple-style (apple.com / macOS 27 page) restyle of a Bi-Weekly deck copy

Run order on a Drive copy of the live deck (set P in apple_cards.py to the copy's id):
1. apple_restyle.py  — fonts/colors/backgrounds, sidebar chrome removal
2. apple_cards.py    — claim/review cards -> photo tile + spec tiles + pill buttons (roles read from the ORIGINAL deck by element id)
3. apple_tables.py   — TOP-7 / SIREN tables on white tiles, hairlines, blue 보기 links
4. apple_v3.py       — cover hero + section chips, Appendix (apple.com footer style), closing slide
5. GAS runOverviewApple261002 (Code.js) — TOP3 gauges in CHART_THEMES.apple

All scripts are idempotent (they delete their own at_/ap3_ shapes before redrawing).
