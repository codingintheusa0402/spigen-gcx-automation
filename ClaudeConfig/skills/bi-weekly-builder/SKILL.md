---
name: bi-weekly-builder
description: >-
  Builds the complete GCX Bi-weekly Report for a period, end to end, the way 261016 was built and approved
  (2026-10-16): classic working deck (claim/review cards from the 고객사진 모음 decks, Overview TOP3 gauges,
  TOP-7 from Caspi, SIREN tables, dashboard numbers) → Apple-style working copy → transplanted into a copy of the
  previous SENT deck (golden base, hand-made design kept) → SIREN 신규 이슈 사례 cards → share + rename (deck and
  its Apps Script) → FINAL step: Apple-themed live-HTML Gmail draft (never sent). Trigger on "run bi-weekly builder",
  "build the bi-weekly slides / report", "바이위클리 만들어줘", "make this period's bi-weekly", "bi-weekly report
  for this time", or any close paraphrase. ALWAYS confirm the date range and series first.
metadata:
  category: automation
  locale: ko-KR
  phase: v3.0.0-live
---

# bi-weekly-builder (v3 runbook)

**Project:** `~/Desktop/GCX/GAS_Operations/Bi-Weekly/` — `Code.js` (GAS, clasp) + `tools/` (Python) + `tools/apple_theme/`.
**Golden base (latest SENT deck) = 261016** `1sJ49LsEqMFrFyykAsgpWIREO9SCnWa5GW-ILEBLbG8Q` (bound GAS `1KwZrx…`);
before: 261002 `1quCr9Xj…`. Update this line every period with the new sent deck.

**Runner script:** GAS project `15U2Db0iWmLAA7Z8l2FVZ5rHe8617jrufJ5Hdc9r4oHneWC_Bpl4hF7Hs` (bound to the old 261002
classic deck, already authorized incl. Gmail). It drives ANY deck via ids in `BW_RUN` (`deck`, `appleDeck`) — no
need for the new copy's own script. Run a function = make it the **first function** of the file (or a temp
`Aa*.js` file, which sorts first), `clasp push --force`, reload the editor, open that file, click Run
(ghost-click: the Run-dropdown is unreliable). Temp files: delete → **touch Code.js** → push (a pure delete
says "already up to date" and leaves the remote file). `.claspignore` excludes `tools/**`.
Python: homebrew `python3` with `tools/rate_pipeline.creds()` (gws_shim: Drive+Slides+Sheets, no Gmail);
charts/icons with `/usr/bin/python3` (matplotlib/Pillow). Render checks: `presentations.pages.getThumbnail`.
Slides writes ≤60/min — batch, retry 429.

## Standing user rules (all ALWAYS)
1. Series order = **newest launch first** (iPhone 18 → Pixel 11 → Galaxy Z8 …) in index, sections, Overview.
2. Apple deliverable = **copy of the previous SENT deck**, its cover/index/dashboard/Appendix/closing kept; never
   regenerate the look from scratch. Bright backgrounds; cover (orange) + closing (black logo) untouched.
3. **SIREN 신규 이슈 사례**: for each SIREN registered in the period, 2–3 cards from its SIREN deck's claim/review slides
   that show defect photos.
4. Product names: drop `SP_`; underscore codes → last token in caps (`_new_biz__ev25mfaq` → `EV25MFAQ`). Lives in
   `Code.js _chartProductName` and `mail_html.pname`.
5. "알 수 없음 (No product info)" / 해당없음 / unknown are **never ranked** — merged into 그 외 (`_isUnknownLabel`/`_rankKnown`).
6. Claims everywhere = Zendesk `4. Product Issue` only (EU+UK for TOP-7). Bad reviews exclude the `긍정 리뷰` tag.
7. Every deck created → Drive `anyone`/`writer` (link-only). Every deck's **bound Apps Script renamed = deck name**.
8. Final step = Gmail **draft** (never send), **live HTML** (no screenshots), report link at the TOP, subject
   `Bi-Weekly 제품 이슈 Report, <yy.mm.dd> ~ <yy.mm.dd>` (no "n of 41"), no line-wrapping labels, < 100 KB.
9. Chat notices only where the user says; never test-post to live rooms (private room `AAQAc9NQmJQ` for tests).

## 0. Inputs (one AskUserQuestion)
- Date range: default = day after the previous report's end → today (report code = today `yymmdd`).
- Series: default = same as last period. Deck: **copy it yourself** (don't wait for a URL) — Drive-copy the previous
  classic deck → `<code> GCX Bi-weekly Report (classic)` and use `BW_RUN.deck` + the 15U2 runner.

## 1. Classic working deck
1. Drive-copy previous classic deck (same folder `0ADTRUOsN0M72Uk9PVA`), share anyone-edit.
2. `Code.js` BW_RUN: `start`, `end`, `deck` (new id), `sources`. `node -e "new Function(…)"`, `clasp push --force`.
3. **TOP-7** (`tools/`): `python3 rate_pipeline.py prep --series glxZ8,pixel11,iphone18` → run `state/caspi_sales.sql`
   with Caspi `run_query` (limit 10000) → copy the saved tool-result file to `state/caspi_sales.json` (check
   `"truncated": false`) → `python3 rate_pipeline.py aggregate` → `python3 rate_pipeline.py slides --deck <classic>`.
4. **Cards**: run `bwRunCards` (twice if `remaining>0`; ~5 min each). Then delete cards dated < start (Python: card
   slides contain '클레임 내용'/'리뷰 내용'; date = first `YYYY.MM.DD`). The maker strips inherited SIREN badges and
   re-adds only real matches (26년 SIREN SKU + issue) — still eyeball one badged card.
5. **SIREN tables**: rows = `26년 SIREN` (sheet `15Jh6ZFD…`, header row 18) with 날짜 ≥ half-year start, 등록=O, 9 rows/slide.
   New rows: `insertTableRows` + per-cell insertText with the reference row's text/paragraph style + cell fill,
   duplicate a row icon and relink it (SIREN deck = C-column chip `richLinkProperties.uri`), grow the shadow rect,
   `replaceAllText` the count (= all 2026 rows with 등록=O).
6. **Dashboard + cover**: `python3 dashboard_stats.py <cutoff = last data day> --apply <deck> --old '<old values|…>'`
   (cumulative, (~MM.DD), YoY, top-3 defect 인입사유); cover date via replaceAllText.
7. Gauges on classic are optional (the Apple working copy gets its own in step 2.3).

## 2. Apple working copy (`tools/apple_theme/`, all idempotent)
1. Drive-copy the classic → `<code> …-apple-work`; set `report.json` (report_code/date/period/source_deck =
   classic, apple_deck = this copy) and `BW_RUN.appleDeck`.
2. Reorder to newest-first + index swap if the classic isn't already (insert-then-delete texts, one
   `updateSlidesPosition` per slide, card titles renumbered via temp-token replaceAllText).
3. Run: `apple_restyle.py` → `apple_cards.py` → `apple_tables.py` → `apple_v3.py` (Appendix after last SIREN slide) →
   `apple_v4.py` (SIREN stat tile; any count) → GAS `bwRunOverviewApple` (uses `BW_RUN.overview` ids when
   `appleOverview` is empty — **clear `appleOverview` for a working copy**) → `weekly_fetch.py` + `/usr/bin/python3 chart.py`
   → `make_assets_gas.py` + run `aaRun` (icons, chart swap — finds the chart inside groups — layout icons, index
   icons follow report.json series order) → `apple_type.py fix` → `spigen_images.py` → `align_icons.py`.

## 3. Final deck = golden base + this period's slides
1. Drive-copy the golden base (latest SENT deck) → `<code> GCX Bi-weekly Report`; share anyone-edit.
2. `transplant_gas.txt` → `AaTransplant.js` (set DST/SRC), push, run `aaTransplant`: keeps cover/index/dashboard/
   Appendix/closing of DST, inserts SRC's 전제품 → claim TOP3 → TOP-7 → cards → SIREN tables, then bad-review backups
   after the Appendix, swaps the chart image. Delete temp file, touch Code.js, push.
3. Cover date + dashboard numbers: `dashboard_stats.py … --apply <final>`; Appendix `원본 보고서 (<code>)` text + link →
   classic deck. Fix titles to golden style (claim TOP3 titles without "1. ").
4. **SIREN issue cards**: write `siren_cards.json` (picks = 2–3 photo slides per new SIREN; open the SIREN deck from the
   26년 SIREN C-column chip; prefer 인입일 inside the period; shorten long details to ≤4 lines) →
   `python3 siren_cards.py` (template + roles auto-detected, inserted after the last SIREN table, `BW_DECK=<id>` to test
   on a copy) → `apple_type.py fix` → `align_icons.py`. Add `2026 GCX 신규 SIREN 이슈 사례` to index 05.
5. If gauges need a rebuild on the final deck: put its gauge slide ids in `BW_RUN.appleOverview`, bump the done-key
   version in `bwRunOverviewApple`, run, then `apple_type.py fix` + `aaRun` (icons) + `align_icons.py`.
6. Rename decks (`… (classic)` for the working classic), trash the `-apple-work` copy, then rename each deck's bound
   Apps Script (open deck → Extensions → Apps Script → click title → Rename; a fresh API copy creates a new empty
   project — name it too). Update the "Golden base" line at the top of this file + memory.
7. Render and compare with the golden base: cover, index, dashboard, 전제품, one series TOP3, a TOP-7, 2 cards, SIREN
   tables, SIREN issue cards, Appendix, closing.

## 4. FINAL — Gmail draft (live HTML)
```
cd tools/apple_theme
# mail.json: subject, title "Bi-Weekly Report <yy. mm. dd>", subtitle, intro, voc [[product, issue]] (top Zendesk
#   4. Product Issue product×인입사유 within the period), siren [[date · product, issue]], replace_draft ("" or old draft id)
python3 weekly_fetch.py && python3 mail_html.py && python3 mail_build.py     # prints bytes — must be < ~100 KB
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --window-size=760,5400 \
   --screenshot=<scratch>/mail.png "file://$PWD/mail_preview.html"          # LOOK at it (wraps, alignment)
python3 mail_gas.py   # → ../../AaMailDraft.js ; push, run aaMailDraft, delete file, touch Code.js, push
```
mail_html reads the FINAL deck: dashboard (monitoring list w/ red line, top-3 tiles → Looker links, cumulative + YoY,
weekly bars as table cells), every TOP3 slide (stacked bar + rows, unknown merged, names normalized, long labels
truncated with …), TOP-7 tables (spigen.com thumbs, ≥2% red), SIREN table (보기 links, period dates red, sheet button,
count). Section titles link to their slides; header title + blue capsule link to the deck (박기세님 feedback).
To `wbd@spigen.com`, CC = 43 (list in mail_gas.py — refresh from the last sent mail `in:sent subject:Bi-Weekly`).
Compaction keeps size down: font stack once on the wrapper, color/weight inherited, `nowrap` set per table.
Then open the draft in Gmail (Drafts → expand) and confirm it renders; never click Send.

## 5. Report back
Final deck link, classic link, Gmail draft subject, TOP-7 ≥2% highlights, new SIRENs, maker notes (missing images /
ratings), anything skipped. Commit + push (`~/Desktop/GCX` repo; auto-sync also commits) and update memory.

## Design system (unchanged)
Canvas `#F5F5F7`, white tiles r=12pt (2 rects + 4 circles; never ROUND_RECTANGLE), ink `#1D1D1F`, gray `#6E6E73`/`#86868B`,
hairline `#E5E5EA`, blue `#0071E3` (links `#0066CC`), red `#FF3B30`, ring `#0071E3 #64D2FF #5E5CE6 #D2D2D7`; capsules =
rect + 2 circles; Inter (600 bold) + Noto Sans KR; titles x36 y16 18pt bold; Material Symbols Rounded icons.

## Gotchas
- Caspi `run_query` truncates at ~100 KB → group in SQL; aggregate refuses truncated input.
- Slides element transforms inside groups are group-relative → match by group membership, not x/y.
- `updateSlidesPosition` with several ids needs them in current deck order → move one by one, re-reading order.
- Duplicated card templates carry their SIREN badge → the maker removes it (fixed 261016).
- New shapes: 3000000 EMU × scale; set `weightedFontFamily` (fontFamily alone drops bold); icons drift after font
  changes → `align_icons.py`; thumbnails can be stale right after a GAS run.
- Gmail: clips > 102 KB; standards-mode tables don't inherit font-size from the wrapper — set it on each table;
  links turn blue unless the anchor has its own lowercase color.
- Deck copies via the Drive API don't carry a bound script (opening Extensions → Apps Script creates an empty one).

## Linked
`siren-sweeper` reuses `tools/apple_theme` helpers (`apple_cards.rtile`, `apple_v3.pill/hair/text`, `apple_type.*`,
`spigen_images.catalog/resolve`) — keep signatures. `siren-report` makes the SIREN decks the issue cards read from.
