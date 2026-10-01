# Apple-style Bi-Weekly deck (apple.com / macOS 27 page design)

Per-period settings: `report.json` (deck ids, date, series). Full runbook: the `bi-weekly-builder` skill (§7).
Run order: apple_restyle → apple_cards → apple_tables → apple_v3 → apple_v4 → GAS bwRunOverviewApple →
weekly_fetch + chart (/usr/bin/python3) → make_assets_gas (temp GAS `aaRun`) → apple_layouts → apple_type fix → align_icons.
All scripts are idempotent (they delete their own at_/ap3_/ap4_/lay_ elements before redrawing).
Icons: Material Symbols Rounded (Apache-2.0) via icons_render.py; the font is downloaded on demand, not committed.
11. spigen_images.py — official spigen.com product shots: card header tile + TOP-7 row thumbnails (run after apple_cards/apple_tables)
