---
name: siren-finder
description: Find SIREN-able product issues — same product line-up × same defect with Zendesk claims + Amazon bad reviews ≥ 5 — skip ones already in the CQ Emergency Net registry, report the rest (with each SKU's Amazon 판매량 from Caspi) to the user's private Google Chat room, then build SIREN decks with the siren-report skill for the cases the user picks. Trigger when the user asks to "find SIRENable cases", "SIREN 후보 찾아줘", "SIREN 등록할 만한 이슈 찾아서 보고해줘", "run siren finder", or any close paraphrase.
---

# siren-finder

Finds SIREN candidates across ALL products and hands chosen ones to the
`siren-report` skill (`~/.claude/skills/siren-report/`) for deck building.
Established 2026-09-28 (after the manual Classic LS 이염/변색 and TA MagFit 킥스탠드 cases).

## Rules (user-set 2026-09-28 — don't change without being asked)

- **Case** = one product **line-up** × one 1차 Defect type. Line-up = device generation +
  product model with all sizes/colors merged (iPhone 18 Pro + 18 Pro Max Tough Armor Pro
  = one case). Derived via SKU → product master `Data` 기종명/모델명 (first device in
  `기종명` decides the generation, so 17 Pro Max claims merge with 18 Pro Max SKUs).
  Power-accessory SKUs whose model is generic are keyed per SKU. See `lineup.py`.
- **Threshold**: claims + bad reviews **≥ 5** (`--threshold`).
- **Window**: all of the current year by default (`--since`, default `YYYY-01-01`).
- **Excluded defects**: 황변, anything `(Delivery Issue)_…`, 배송·중고품 related.
- **Registry check first**: a case is dropped if the same line-up SKUs already have a row
  for a similar issue in `2026_CQ_Spigen Issue Report Emergency Net Sheet_R01`
  (`137K4hpNfHoyb6PEb64gO7Wxi5b3CPlnMETQt-bbPKKE`) — compare `SKU` + `불량 유형` +
  `사유 ( 상세 )`. Tabs checked: Glass, Case (the ones the user named) + 전기전자, 생활용품.
  Similarity is judged by the Claude CLI and cached.
- **Report** to the user's private room first (`spaces/AAQAc9NQmJQ`, default webhook in
  the script) — never to a team room before the user says so.
- **Every case in the report shows 판매량** per SKU (and total + claim/review rate) from
  Caspi `S3.AMAZON_SELLER.FLAT_FILE_ALL_ORDERS_DATA_BY_ORDER_DATE_GENERAL` (all Amazon
  channels, cancelled + `amzn.gr.` warehouse resales excluded, since `--since`), via
  registered query `pq_abe464407bdbd0de98` + the account API key in
  `~/.config/siren_finder/secrets.json`. Raw skus look like `ACS09826PAN` / `ACS09826SGP` /
  `amzn.gr.ACS09826PAN-…` → base SKU is `REGEXP_SUBSTR(sku,'[A-Z]{3}[0-9]{5}')`.
  **Page with limit 1000** — a 2000-row page came back silently cut (nextOffset null).

## Sources

| What | Sheet | Notes |
|---|---|---|
| Claims (whole RAW) | `1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc` / `26년 전체문의` | Category `4. Product Issue`; defect = `1차 Defect Reason or Inquiries`; SKU = `★문의SKU` |
| Bad reviews (whole Caspi/SC data) | `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw` / `SC` | ★1–3 only. No defect column → each review classified once by `claude -p --model sonnet` into the same 1차 Defect taxonomy (or `none`), cached in `~/.config/siren_finder/review_cls.json` |
| Registry | `137K4hp…` Glass / Case / 전기전자 / 생활용품 | header row is the one containing `SKU` (row 3) |
| Product master | `1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo` / `Data` | SKU/ASIN → 기종명, 모델명, 생산업체 |

Sheets are read with the gws_shim token (`~/.config/gws_shim/token.json`).

## Workflow

1. **Analyse** (no send): `python3 ~/.claude/skills/siren-finder/find_siren.py`
   → prints the chat preview; writes `~/.config/siren_finder/runs/candidates_<yymmdd>.json`
   (all ≥threshold cases with `registered` flag + matching registry rows).
   First run classifies every bad review (~5.5k, ~20–30 min with 6 workers); later runs
   only classify new reviews. A failed classifier batch is logged and retried next run.
2. **Sanity-check** a few top cases in the preview (line-up grouping, defect label,
   registry decision). Spot-check any case whose registry row text is borderline.
3. **Send** to the private room: `find_siren.py --send`. It first creates a new spreadsheet
   `SIREN 후보 리포트_<yymmdd>` (tabs 미등록 / 기등록: every case with counts, dates, SKUs,
   판매량 total + per SKU, claim+review rate, registry match, all ticket/review links), then
   posts: header (totals, registry-excluded count, sheet link) + the top `--top` (30) cases
   (name, defect, total, date span, 판매량 + rate, first 3 ticket/review links), split into
   ≤3.8k-char messages. With all-year data there are ~500 open cases, so the sheet is the
   full list — Chat only carries the top of it.
4. **Wait for the user** to pick which cases to take to SIREN.
5. For each picked case N (numbering = the report's):
   `find_siren.py --case N --export /path/case_N.json` → siren-report data skeleton
   (claims + reviews with country/dates/links, overview by-country counts, SKU/ASIN list;
   reviews already carry `original_text`, Korean `detail_text` and photo).
   Then, like the Classic LS / TA MagFit decks:
   - open each claim ticket in Zendesk (claude-in-chrome, `/api/v2/tickets/<id>/comments.json`
     with `credentials:'include'`), read the customer's messages, write a Korean
     `detail_text`, and take up to 3 customer images/videos (HEIC → convert, videos → Drive;
     see siren-report SKILL.md "API gotchas").
   - send the Chat message in the established format (review-request line, `[클레임N]` /
     `[배드리뷰N]` labels hyperlinked, product line with ASIN, image grid) to the room the
     user names, then build the deck:
     `python3 ~/.claude/skills/siren-report/build_siren_slides.py --data case_N.json`.

## Output-size gotchas

- Chrome `javascript_tool` output is capped (~1k chars) and blocks query strings — store
  results in `window.__x` and read them in slices with `?`/`&`/`=` replaced.
- Chat text messages: keep each under ~4k chars (the script splits automatically).

## Linked skills
- **`siren-sweeper`** (GCX SIREN 반기별 Sweeper) runs on top of this skill's candidates: builds SIREN decks for the top N via `siren-report`,
  checks the GCX/영업 SIREN registries this skill does not check (it only checks CQ Emergency Net), and makes the
  'GCX SIREN <상/하반기> 등록현황 및 VOC 점검' summary deck in the `bi-weekly-builder` Apple design.
- `siren-report` builds the decks for picked cases.
