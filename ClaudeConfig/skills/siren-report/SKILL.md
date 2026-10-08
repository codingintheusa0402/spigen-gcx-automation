---
name: siren-report
description: Build the SIREN 제품 클레임 조치 사항 검토 Google Slides registration deck (title → 제품 클레임 개요 → one slide per claim/bad review) for a defect case that crossed the SIREN threshold. Trigger on "SIREN 덱 만들어줘", "build the SIREN report/deck", or after siren-finder picks cases. Never generates the CQ investigation section.
---

# siren-report

Builds a **SIREN 제품 클레임 조치 사항 검토** Google Slides deck — the formal report
deck GCX/CX registers whenever a defect pattern crosses the SIREN threshold (see
`ticket-reporter/SKILL.md`'s "SIREN check" section for how a case gets flagged).

This skill produces the **GCX/CX-side registration deck only**: title → 제품 클레임
개요 → one slide per matching claim/bad-review. The later CQ(품질관리팀) investigation
section seen in mature decks (Field Issue / 제품 검증 / CQ 검토 의견 slides, embedded
verification videos) is added manually by the CQ team after physical sample testing —
**out of scope, never auto-generate it.**

## Template source (reverse-engineered 2026-09-21)

Read directly via the Slides API from two real decks linked (as Sheets smart chips,
`richLinkProperties.uri`, in the `제목` column) from the **"26년 SIREN"** tracking tab
— sheet `15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI`, tab "'26 GCX KPI (NEW
ver.)_26.01.07", header row 18, cases from row 19:
- `19KWYQpoh-8W6WJGGK30UKl6a5njutrTtNF8yc8QcIGg` — "iPhone 17 Pro/Pro Max용 Classic LS
  MagFit 유격 이슈" (7 slides — the larger reference, includes the later CQ section).
- `1oziw7rbEL1pPLlfbiu_eg6uFOmJgPlSKcbLLwKuPZqU` — "갤럭시 S26 Ultra용 Glas.tR EZ Fit
  Pro(AGL11068)_지문인식 이슈" (22 slides).

If the template ever needs re-verification or extension, read a real deck from that
same tracking tab with the Slides API (`presentations().get()`) rather than guessing —
don't rebuild from memory of this doc alone.

### Exact layout (do not deviate — font size/alignment/position all matter)

Page: 720pt × 405pt (10in × 5.625in widescreen).

**Slide 1 — title (updated 2026-09-21 to match Spigen's own cover-slide template)**
- Background: solid orange, sampled directly from the `spigen-slides` skill's
  `LIGHT_TEMPLATE_ID` (`1BBG9PR6ZBsEABbJLhbUUfRMkgGYQtNMOWAmLQgPhr70`) cover slide —
  `rgb(1.0, 0.3529412, 0.0)` (`COVER_ORANGE`), NOT the `#FF6B1A` brand-guide value (close
  but not identical — verify against the live template if this ever needs re-checking).
- Title textbox at (27.6, 34.3), size 594.8×99.4, Noto Sans KR, black, two separate runs
  in one box: line 1 `제품 클레임 조치 사항 검토(SIREN)` at 36pt bold, line 2 (the
  specific issue title) at 20pt bold — deliberately smaller than line 1, not the same
  size (confirmed against a real example image the user provided).
- Spigen logo mark image at (634.5, 35.4), size 59.1×55.4 (includes 3pt crop padding) —
  see "Logo asset" below for where this comes from.
- Author/dept line at (27.6, 324.4), size 394.6×40.3, Noto Sans KR 12pt regular, black:
  `경영지원부문ㅣ사업지원실ㅣ글로벌CX전략팀 <담당자> 담당`.
- Date at (607.7, 323.3), size 76.4×42.5, Noto Sans KR 12pt regular, black, START-aligned.
  **Must use the compact `YYYY.MM.DD` format (no spaces)** — this box is narrow enough
  that the report's usual `YYYY. MM. DD` (with spaces) wraps to two lines. The generator
  strips spaces from `data["date"]` for this box only.

**Logo asset**: the template's own logo image lives at an internal
`lh7-rt.googleusercontent.com/slidesz/...` URL that is session-gated — Slides API's
`createImage` fetches server-side with no auth, so that URL 400s. Fixed by: fetching the
template cover slide's `presentations().pages().getThumbnail()` (a real public URL),
cropping just the logo region out of that PNG (pure-Python PNG decode/crop/encode, no
Pillow needed — see the crop script pattern used 2026-09-21 if this needs redoing), then
re-uploading the crop to Drive with `anyone/reader` permission for a stable public URL.
That URL is hardcoded as `SPIGEN_LOGO_URL` in `build_siren_slides.py`. If it ever
disappears (file deleted, permission revoked), regenerate it the same way rather than
reusing the original session-gated URL.

**Slide 2 — 제품 클레임 개요**
- Header textbox at (27.6, 10), Gothic A1 23pt bold: ` 제품 클레임 개요` (leading space
  matches the real decks).
- Table 1 (2×5) at (35, 52), ~650×90: `기종 | 제품명 | 불량 유형 | SKU/ASIN | 제조 업체명`.
  **SKU/ASIN cell (user rule 2026-09-28):** header is `SKU/ASIN` (not `SKU`); value lists
  every affected SKU with its ASIN, one per line: `ACS10478 / B0FH5S3JMK`. Pass
  `product.skus: [{"sku","asin"}, ...]` (legacy `sku`/`asin` fields still work).
  **제조 업체명 (user rule 2026-09-28):** taken from the product master sheet
  `1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo`, tab `Data`, column `생산업체`, matched
  on the `SKU` column (`lookup_manufacturer()` runs automatically when
  `product.manufacturer` is empty or `-`). Do NOT use 신제품 라인업 for this — it often
  shows `입고처리 미진행`.
  If `product.asin` is set, the 제품명 body cell is hyperlinked to
  `https://www.amazon.de/dp/<asin>` (user request 2026-09-21 — always `.de`, not a
  country-specific TLD keyed off the claim's country).
  **Multi-SKU line-ups (fixed 2026-10-02):** one pair per line at 11pt in the default 130pt
  column wraps every pair, and with ≥2 SKUs table 1 grew under table 2 (values hidden). Now a
  single SKU keeps the original layout above; with ≥2 SKUs the SKU/ASIN column is widened
  (column widths 95/140/95/220/100), the cell uses 9pt one-per-line (2–6 pairs) or 8pt two-per-line
  (7–12), >12 shows the first 10 + `외 N종`, and tables 1/2 shrink to content height
  so tables 2/3 move down and table 3 still ends on the page. Table 3 gets a 120pt label column
  when there are >5 countries (else `국가별 인입 채널` wraps into 4 lines and runs off the page).
- Table 2 (2×4) at (35, 147) (single-SKU layout; multi-SKU: directly below table 1): `총 인입건 수 | 전체 배드 리뷰 수 | 전체 클레임 수
  (Zendesk) | Global 리뷰 평점`. The `Global 리뷰 평점` **header** cell (not the value
  cell below it) is filled with `ACCENT_ORANGE` (`#FF6B1A`, the Spigen brand orange from
  `spigen-slides.md`) instead of the usual black header fill — user request 2026-09-21 to
  visually distinguish the one Global/cross-market metric from the other three
  Zendesk-sourced counts.
- **Global 리뷰 평점 value (user rule 2026-10-08)** = the star rating on the Amazon page the 제품명 cell hyperlinks to
  (`amazon.de/dp/<product.asin>`, the `#acrPopover` average next to the title), written as `4.5(DE)`. If amazon.de shows no
  rating (out of stock / not sold there / blocked), try amazon.com → amazon.co.uk → amazon.co.jp and label with that country
  (`(US)`/`(UK)`/`(JP)`); none → `N/A`. Auto-filled by `amazon_rating.global_rating()` when `overview.global_rating` is empty or
  `-`. Check one by hand: `python3 amazon_rating.py <ASIN>`. To backfill an existing deck, edit only the value cell (row 1,
  col 3) of the table whose header has `Global 리뷰 평점` on slide 2 — don't rebuild. `product.asin` is forced to one of
  `product.skus[].asin` (first one) if it isn't among them — claim-level ASINs have pointed at other products.
- Table 3 (4×(1+N countries)) at (35, 252): row0 `국가별 인입 채널 | <countries>`, row1
  `배드 리뷰 수 | ...`, row2 `고객 클레임 수(Zendesk) | ...`, row3 `Total | ...`.
  Country columns = only countries that actually have data (fallback to the full
  `COUNTRY_ORDER` list if none do). Order: UK, DE, IT, FR, ES, JP, IN, US, KR.

**Slide 3+ — one per claim/bad-review**
- Header at (27.6, 8.9), Gothic A1 23pt bold: **separate counters per type (user rule
  2026-09-28)** — `클레임 #01`…`클레임 #15`, then `배드 리뷰 #01`, `배드 리뷰 #02`… The
  generator renumbers automatically from the label prefix (`배드…` = 배드 리뷰, else 클레임),
  so input label numbers don't matter.
  **The `#NN` part is colored (user rule 2026-09-28): blue (`#0000FF`) on 클레임 slides,
  red (`#FF0000`) on 배드 리뷰 slides**; the label word itself stays black.
- Table (4 rows × 6 cols) at (27.6, 56.3), ~665×150:
  - row0 header: `구분 | 국가 | 구매일 | 인입일 | 주문후 결함 발생 기간 | 링크`
  - row1 data: (구분 cell blank) `country | purchase_date | inflow_date |
    days_elapsed | link_label`
  - **링크 cell is ALWAYS hyperlinked (user rule 2026-09-28, default for every deck).**
    클레임 slides: link_label = Zendesk ticket ID, linked to
    `https://spigenhelp.zendesk.com/agent/tickets/<id>`. 배드 리뷰 slides: link_label =
    Amazon Review ID (e.g. `R2NDQ4MTPI3GO8`), linked to the real review page
    `https://<country amazon domain>/gp/customer-reviews/<id>/ref_=bcr_shw_rev_dtl`.
    Pass `link_url` explicitly when known; otherwise `_claim_link_url()` derives it from
    link_label (digits → Zendesk, `R…` → Amazon on the claim's country). Never leave a
    plain label like `클레임` — always use the real ticket/review ID.
  - row2: `상세 내용` merged across the remaining 5 cols, full complaint detail text.
  - row3: `고객 첨부 사진` merged across all 6 cols (label only; `고객 첨부 사진/영상` when
    videos are attached).
    **배드 리뷰 with no customer photo/video (user rule 2026-09-28):** row3 becomes
    `배드 리뷰 원문` | the review's original title + body text (merged across 5 cols) —
    pass it as `original_text`. Claims keep the plain label row even without media.
- Photo(s)/video(s) placed as separate floating elements below the table (not inside a
  cell). Images via `createImage` (`image_urls`); **videos via `createVideo` source DRIVE
  (`video_drive_ids`)** — download the Zendesk video attachment (token URLs are public),
  upload it to a Drive folder (`SIREN_<product>_고객영상_<yymmdd>`) shared domain
  `spigen.com` reader, and pass the file IDs. Cap 3 media per slide; height is clamped
  so media never runs past the 405pt page bottom — size/position varies with aspect ratio; the generator
  lays multiple images out left-to-right with a 12pt gap, capped at ~200pt width each.

**Cell styling** (confirmed from raw API JSON): header cells = black background +
white bold text, Arial 12pt. Body cells = white background + near-black (DARK1) bold
text, Noto Sans 11pt, vertically centered (`contentAlignment: MIDDLE`).

## Files

- `build_siren_slides.py` — the generator. `DeckBuilder` class wraps `presentations
  .create()` + one atomic `batchUpdate()` (chunked at 400 requests). Helper methods:
  `new_slide()`, `add_textbox()`, `add_table()` (handles per-cell styling + merges via
  `mergeTableCells`), `add_image()`.
- `siren_case.example.json` — worked example built from real ticket #1000162836 data
  (iPhone 18 Pro / SP_Glas.tR EZ Fit Slim / 레인보우현상, 1 claim, no bad reviews).

## Usage

```
python3 build_siren_slides.py --data siren_case.json [--dry-run]
```

`--data` JSON schema (see `siren_case.example.json`):

```jsonc
{
  "issue_title": "...",           // slide 1 line 2
  "author": "...",                // 담당자 name for the author line
  "date": "YYYY. MM. DD",         // optional — defaults to today (KST-agnostic, just today's date)
  "product": {
    "device": "...", "product_name": "...", "defect_type": "...",
    "skus": [{"sku": "...", "asin": "..."}], // every affected SKU+ASIN -> SKU/ASIN cell
    "manufacturer": "",                      // leave empty -> auto-lookup from product master 'Data' 생산업체
    "asin": "..."                            // optional — hyperlinks 제품명 to amazon.de/dp/<asin>
  },
  "overview": {
    "total_count": 0, "bad_review_count": 0, "zendesk_count": 0,
    "global_rating": "",                     // leave empty -> auto: Amazon rating of the linked ASIN, e.g. "4.5(DE)" / "N/A"
    "by_country_bad_review": {"DE": 2, ...}, // omit countries with 0/no data
    "by_country_zendesk": {"UK": 1, ...}
  },
  "claims": [
    {
      "label": "클레임 #01",                 // or "배드 리뷰 #02" — number sequentially across both types
      "country": "...", "purchase_date": "YYYY-MM-DD", "inflow_date": "YYYY-MM-DD",
      "days_elapsed": "N일", "link_label": "1000162836",  // Zendesk ticket ID or Amazon Review ID — always hyperlinked
      "link_url": "...",                     // optional — auto-derived from link_label if omitted
      "detail_text": "...",
      "image_urls": ["..."]                  // Zendesk attachment token URLs work directly
    }
  ]
}
```

Auth: reuses the `gws_shim` OAuth token (`~/.config/gws_shim/token.json`, scopes
`drive` + `presentations`) — same token already used elsewhere for Sheets access.

## API gotchas (hard-won, don't re-discover)

- `cellLocation` (in `insertText`/`updateTextStyle`) and `tableRange.location` (in
  `updateTableCellProperties`/`mergeTableCells`) take **only** `{rowIndex,
  columnIndex}` — do NOT nest a `tableObjectId` key inside them. The outer request's
  `objectId` already identifies the table; nesting causes `Unknown name
  "tableObjectId"`.
- `objectId` must be **≥5 characters** or `createShape`/`createTable`/etc. fail with
  a 400. The generator uses `f"siren_{prefix}_{self._counter:04d}"`.
- `batchUpdate` is **atomic** — one bad request fails the whole batch, including a
  `deleteObject` on the default slide that was bundled with it. If a build fails
  partway, the freshly-`create()`d presentation is left behind and essentially empty
  — trash it via `drive.files().update(fileId=..., body={"trashed": True})` before
  retrying, don't leave orphaned files.

- **HEIC images make `batchUpdate` fail with a bare 500 "Internal error encountered"**
  (Slides can't import HEIC; Zendesk serves iPhone photos as `image/heic`). Check each
  attachment's content-type first; convert HEIC with `sips -s format jpeg -Z 2000`, upload
  to the case's Drive folder with `anyone/reader`, and use
  `https://drive.google.com/uc?export=view&id=<id>`. Same applies to Chat image grids.
- When cleaning up failed test decks, trash only by the exact presentation ID you
  created — never by a Drive name search (it matches other people's shared SIREN decks).

## Open items / not yet built

- Not wired into the SIREN alert flow yet (`ticket-reporter/siren.py`'s
  `post_alert()`) — currently a standalone script invoked manually with a hand-built
  JSON. Wiring it up means: on a first-for-cluster SIREN trigger, gather ALL matching
  Zendesk claims + bad reviews (not just the one ticket that fired the alert),
  compute the by-country breakdowns, look up manufacturer/SKU from 신제품 라인업, and
  call `build_deck()` — then post the resulting Slides link back to the alert thread.
- No chart-generation for the "동일 생산업체 및 동일 제품 클레임 인입 현황" trend-chart
  slide seen in some real decks (slide 3 of the smaller 7-slide example) — would need
  a charting library/image-generation step, not implemented.

## Linked skills
- Called by **`siren-sweeper`** (GCX SIREN 반기별 Sweeper, step 5) to build one deck per candidate; also by `siren-finder` step 5.
- Input comes from `siren-finder` (`find_siren.py --case N --export`).
