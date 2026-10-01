---
name: bi-weekly-builder
description: >-
  Builds the complete GCX Bi-weekly Report for a period, end to end, the way it
  was built and approved on 2026-10-01 (261002): the classic deck (claim/review
  card slides from the 클레임 및 배드리뷰 고객사진 모음 decks, Overview TOP3
  gauge slides, "클레임 + 배드리뷰 / 판매량 TOP 7" tables from Caspi, SIREN
  status slides, index/cover updates) AND — as the default deliverable — its
  apple.com-style copy ("<code> GCX Bi-weekly Report-apple": bright #F5F5F7
  canvas, white rounded tiles, capsule buttons, Material-Symbols icons, Apple
  bar chart, uniform titles, Inter/Noto Sans KR, Appendix footer page, Apple
  theme + slide-type layouts, official spigen.com product images). The Apple
  version is the one SENT, under the standard name "<yymmdd> GCX Bi-weekly
  Report". Trigger when the user says "run bi-weekly
  builder", "build the bi-weekly slides / report", "바이위클리 만들어줘",
  "make this period's bi-weekly", or any close paraphrase. ALWAYS ask for the
  date range, the monitored series, and the deck / Apps Script URL first.
metadata:
  category: automation
  locale: ko-KR
  phase: v2.1.0-live
---

# bi-weekly-builder

Project folder: `~/Desktop/GCX/GAS_Operations/Bi-Weekly/` (deck-bound Apps Script via clasp
+ `tools/` Python). **Golden reference = the final deck sent for 261002 (approved 2026-10-01):**
https://docs.google.com/presentation/d/1quCr9Xj-pSsVXKrYuaEOq0LPILZMN2LPUwBkY1f_GFI/edit
(titled `261002 GCX Bi-weekly Report`; built from the classic working deck `12NxCxbW…`).
"If the skill makes one like this, it did it correctly" — render it next to your output when in doubt.

### Final deck composition (81 slides for 3 series — reproduce this structure)
| slides | content |
|---|---|
| 1 | Cover — classic orange Spigen cover (title, team line, date) |
| 2 | Index — 01 Overview · 02/03/04 one block per series · 05 GCX SIREN, icons beside numbers |
| 3 | 1. Overview dashboard — 집중 모니터링 대상 (newest series red), 최다 인입사유 top-3 tiles, 누적 클레임 + YoY, Apple weekly bar chart |
| 4 | 1. Overview (2026년 전제품) — TOP3 gauges, claims 2026-01-01~latest |
| 5–7 | 1. Overview (<series>) 배드리뷰 TOP3 — one per series, cumulative |
| 8–10 | 1. Overview (<series>) 클레임 TOP3 — one per series, cumulative |
| 11–13 | Overview (<series>) 클레임 + 배드리뷰 / 판매량 TOP 7 — product thumbnails |
| 14… | 2./3./4. <Series> Claims / Reviews — this period's cards only, product tile + photos + spec tiles |
| next 1–2 | 5. GCX SIREN (26년 하반기 등록 현황) — 9 rows per slide, stat tile + sheet button |
| next | Appendix — apple.com footer |
| last | Closing — black Spigen-logo slide |
Memory with the history/gotchas: `biweekly_slide_maker_and_lg_designs`.

Python: Slides/Sheets/Drive via `tools/rate_pipeline.creds()` (gws_shim token,
kjw@spigen.com). Charts/icons render with **`/usr/bin/python3`** (has matplotlib + Pillow;
the homebrew 3.14 can't install them). Slides API write quota is 60 req/min — batch
(≤400 requests per call) and retry on 429 (all scripts here already do).

---------------------------------------------------------------------------------------------
## 0. Ask first (one AskUserQuestion turn, never assume)

1. **Date range** for the cards (작성 날짜), e.g. `2026-09-11 ~ today`.
2. **Series to monitor** (e.g. Galaxy Z8 / Pixel 11 / iPhone 18). Each needs its
   고객사진 모음 deck id + product monitoring sheet id (tabs `신제품 라인업`, `1-3점`, `1-5점`, `DE`).
   Find unknown ones with Drive search `title contains '고객사진'`.
3. **This period's deck URL + its bound Apps Script URL** (the user copies the last
   classic deck and renames it `<yymmdd> GCX Bi-weekly Report`; the bound script id changes
   with every copy → set `.clasp.json` `scriptId`).
Default: also produce the Apple copy (§7). Only skip it if the user says so.

---------------------------------------------------------------------------------------------
## 1. Prepare the classic deck

- `.clasp.json` → this period's scriptId. New series → add to `CLAIM_SLIDE_SOURCES`,
  `BAD_REVIEW_SOURCES` (Code.js), the SlideMaker.html checkbox, and `tools/series.json`.
- Edit the **`BW_RUN` block at the top of Code.js**: `start`, `end`, `sources`,
  `appleDeck` (fill after §7 copy), and `overview` (one job per Overview TOP3 slide id —
  ids stay stable across copies; new series → duplicate the Pixel 11 TOP3 slides with
  Slides API `duplicateObject`, replace "Pixel 11 Series" in the title, add jobs).
- `node -e "new Function(require('fs').readFileSync('Code.js','utf8'))"` then `clasp push --force`.
- Running from the editor: the Run dropdown can't be clicked reliably (ghost-click) —
  **put the wrapper you need as the first function of Code.js**, push, reload the editor,
  open Code.gs, Run. A new copy needs OAuth: the user clicks Review permissions → Allow.

## 2. Claim / review cards (`bwRunCards`)

- A family with **no card in the deck is skipped** → seed one first: duplicate the last card of
  another family, set its title to `N. <Series> Claims / Reviews`; delete seeds after the run.
- Run `bwRunCards` (re-runs only add missing cards; 6-min cap → run again).
- Source decks may contain a dated "템플릿 예시" sample slide — the maker skips it (`_readSourceSlide`).
- **Delete the copied previous-period cards** (card date < start) — the deck holds only this period.
- Verify 2–3 cards via `presentations.pages.getThumbnail` (not the browser).

## 3. Overview TOP3 gauge slides (`bwRunOverview`)

Data definitions (user-approved; reproduce the old numbers):
- **Claims** = Zendesk `26년 전체문의`, Category `4. Product Issue`, Device contains the series
  (e.g. `Galaxy Z Fold 8`/`Galaxy Z Flip 8`, `Google Pixel 11`, `iPhone 18`).
- **Bad reviews** = the series sheet `1-3점`, **excluding the `긍정 리뷰` tag**.
- **Series slides are cumulative (all dates); "Overview (2026년 전제품)" = 2026-01-01 ~ latest.**
- Display: product names drop the `SP_` prefix; every card uses the same type (count 13.5/12,
  title 8.5, legend/value 7.5 @115%), long labels lose their "(…)" instead of shrinking.
`rebuildChartGridOnSlide` redraws in place (keeps title/sidebar); Zendesk sheet read once per run.

## 4. TOP-7 slides (`tools/rate_pipeline.py`)

```
cd tools
python3 rate_pipeline.py prep --series glxZ8,pixel11,iphone18      # writes state/caspi_sales.sql
#  run state/caspi_sales.sql with mcp__claude_ai_CaspiLM__run_query (limit 10000),
#  copy the saved tool-result file to state/caspi_sales.json
python3 rate_pipeline.py aggregate      # REFUSES a truncated Caspi result (truncated:true)
python3 rate_pipeline.py slides --deck <classicDeckId>
```
- Sales = `S3.AMAZON_SELLER.VAT_TRANSACTION_DATA`, SALE, EU+UK, 2026, grouped **by SKU only**
  (per-ASIN/marketplace rows exceed Caspi's ~100KB cap and get silently truncated).
- Rate = (claims + bad reviews) ÷ sold, min 20 sold, ≥2% red. Footnote says EU+UK / VAT data / data month.
- If the user asks whether 판매량 is real: spot-check 3–5 SKUs with a direct Caspi COUNT/SUM.

## 5. SIREN slides

- Rows = `26년 SIREN` (sheet `15Jh6ZFD…`, header row 18) with SIREN 등록 = O since the half-year start.
  The table holds 9 rows → overflow goes on a duplicated continuation slide (delete extra rows,
  refill cells insert-then-delete, relink the per-row icons/links, shrink the shadow rect).
- Count badge = number of rows with SIREN 등록 = O. Deck links: find the review deck in Drive
  by the sheet's C title (try keyword searches when the exact title misses).

## 6. Cover / index / dashboard

- Cover date = report date. Index slide: add a section block per series (04 …), SIREN last;
  section titles renumbered (`N. <Series> Claims / Reviews`, `N. GCX SIREN …`).
- Overview dashboard (slide 3): `2026 전제품 누적 클레임 인입건 수(~MM.DD)` = row count of
  `26년 전체문의` (created ≤ cutoff), YoY via `getYoYStats` logic, top-3 `인입사유` counts
  (all categories), 집중 모니터링 대상 list incl. each series' review-monitoring window —
  **the newest series' line goes first, in red (#FF3B30)**, e.g. `iPhone 18 Series: 26.9.21~26.12.21`,
  older series in gray below it.

---------------------------------------------------------------------------------------------
## 7. Apple version (default deliverable) — `tools/apple_theme/`

1. Drive-copy the finished classic deck → `<code> GCX Bi-weekly Report-apple` (same folder) and
   work on the copy. When it's done and approved, the Apple deck becomes the sent report: rename it
   `<code> GCX Bi-weekly Report` (ask the user first — the classic working deck has the same name;
   suggest renaming that one `<code> GCX Bi-weekly Report (classic)`).
2. Fill `tools/apple_theme/report.json` (report_code, report_date, period, source_deck,
   apple_deck, series[chip/title_prefix/source_deck/sheet], zendesk/siren/looker) and set
   `BW_RUN.appleDeck` in Code.js.
3. Run in this order (homebrew `python3` unless noted; every script is idempotent):
   | step | command | what it does |
   |---|---|---|
   | a | `python3 apple_restyle.py` | backgrounds, removes navy sidebar chrome, Noto Sans KR, role-based colors, light tables |
   | b | `python3 apple_cards.py` | each card → photo tile (photos fit-scaled) + content tile + 5 spec tiles, labels, blue capsule button, red SIREN capsule (roles read from the CLASSIC deck by element id) |
   | c | `python3 apple_tables.py` | TOP-7/SIREN tables on white tiles, hairline rows, blue "보기" links instead of yellow icons |
   | d | `python3 apple_v3.py` | Appendix (apple.com footer: 5 icon-headed link columns, line "젠데스크, 아마존 배드리뷰 데이터 접근 권한이 필요하면 **Caspi 접근 신청** 페이지에서 요청하세요.", copyright + nav row `Overview · <series…> · SPIGEN` (SPIGEN → https://www.spigen.com/ as one link), no country). The Apple cover hero only runs if `report.json` `apple_cover: true` — default false |
   | e | `python3 apple_v4.py` | SIREN header/column alignment, stat tile (rounded tile aligned to the table's right edge, 8pt gap), outline capsule |
   | f | GAS `bwRunOverviewApple` | TOP3 gauges in `CHART_THEMES.apple` (white tiles, blue/light-blue/indigo/gray rings) |
   | g | `python3 weekly_fetch.py` then `/usr/bin/python3 chart.py` | Apple weekly bar chart (gray capsule-top bars, latest week blue, faint trend line, SF Pro + Apple SD Gothic Neo) |
   | h | `/usr/bin/python3 icons_render.py` (only if `icons/*.png` missing) then `python3 make_assets_gas.py` | writes temp `AppleAssets.js` (icons on slides, chart swap, layout icons) → `clasp push`, run `aaRun` (first function of AppleAssets.gs), then **delete AppleAssets.js and push again** |
   | i | `python3 apple_layouts.py` | Apple theme (master `simple-light-2` color scheme + default type) and 10 layouts on the unused stock layouts p2–p11 (Cover, Index, Overview dashboard, TOP3 gauges, Table, Claim/Review card, SIREN, Appendix, Closing, Big number). Re-run step h afterwards for layout icons. Layout names can't be set via API → tell the user: View → Theme builder → right-click → Rename |
   | j | `python3 apple_type.py fix` | '1.' list bullets → text, restores bold, every title same box (x36 y16, 18pt bold #1D1D1F, gray "대상 국가…" qualifier), Latin runs → Inter (600 for bold), Hangul → Noto Sans KR |
   | k | `python3 spigen_images.py` | official spigen.com product shots (public Shopify `products.json`, matched by base SKU, fallback model+series title; cached 7 days): 38pt #F5F5F7 rounded tile + shot left of each card's product name, 24pt thumbnail in each TOP-7 product cell. Reports SKUs with no official image (not yet listed on spigen.com) |
   | l | `python3 align_icons.py` | centers every icon on its label's text line (run after any font change) |
   After step j, set gauge counts to 12pt if 4-digit counts touch the ring.
4. Render and look at: cover, index, slide 3, one gauge, one TOP-7, 2 cards (1 and 2 photos),
   SIREN, Appendix, closing. Fix anything off before reporting.

### Design system (keep identical every period)
- Canvas `#F5F5F7` (cover/index/dashboard/closing white), tiles white, **corner radius 12pt**
  built as 2 rects + 4 circles (Slides' ROUND_RECTANGLE radius can't be controlled — never use it).
- Ink `#1D1D1F`, secondary `#6E6E73`, tertiary `#86868B`, hairline `#E5E5EA`/`#D2D2D7`,
  blue `#0071E3` (links `#0066CC`), red `#FF3B30`, chart ring `#0071E3 #64D2FF #5E5CE6 #D2D2D7`.
- Buttons/chips = **true capsules** (rect + 2 circles of diameter h; outline = blue capsule with an
  inset capsule in the background color). Never FLOW_CHART_TERMINATOR (elliptical ends).
- Type: Inter (Latin) + Noto Sans KR (Hangul) — closest available to SF Pro / Apple SD Gothic Neo.
  Titles 18pt bold, card values 11.5 bold, labels 7.5 gray.
- Icons: Material Symbols Rounded (Apache-2.0) in Apple blue (gray for source line, red for SIREN).
  Do not use SF Symbols or Apple PNGs (license).
- Bright backgrounds only (user: "the bg of the slides should be bright not dark") — **except the
  cover and closing: they stay the classic Spigen brand slides** (orange cover with logo, team line,
  date in Gothic A1 / Nanum Gothic; black closing with the white Spigen logo). No script restyles,
  re-fonts or re-colors slide 1 or the last slide.

---------------------------------------------------------------------------------------------
## 8. Gotchas (all hit on 261002)
- Caspi `run_query` silently truncates at ~100KB (`truncated:true`) — aggregate in SQL.
- New Slides shapes store size as 3000000 EMU × scale: move with RELATIVE transforms or
  ABSOLUTE with recomputed scale, never `scaleX:1`.
- Text boxes shorter than their text overflow downward → icons look high; fix with align_icons.
- Setting `fontFamily` drops `bold` on runs → always set weightedFontFamily with weight.
- Thumbnails can be stale right after a GAS run — re-render.
- The user edits the Apple deck by hand (e.g. Appendix footer text, pasted cover/closing):
  scripts must not overwrite user text; if a slide was replaced, ask before rebuilding it.
- CaspiSalesBackfill (sheet Z col 판매량) is a separate job — cumulative EU+UK orders since launch.

## 9. Report back
Links to both decks; slide ranges per series; maker notes (missing ASIN/rating, extra photos,
SIREN badges); TOP-7 highlights (≥2%); SIREN count; Caspi data-month caveat; anything skipped.
