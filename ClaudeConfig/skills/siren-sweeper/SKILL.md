---
name: siren-sweeper
description: >-
  GCX SIREN 반기별 Sweeper — the half-yearly sweep that makes sure no SIREN-able issue is missed. From the latest siren-finder run it builds a SIREN registration deck for each top candidate (siren-report skill), posts review cards, checks every candidate against ALL SIREN registries (GCX KPI 26년 SIREN + previous year's 상/하반기 KPI 실적 검증 + [🚨 SIREN 🚨] 영업 sheet), reports registered issues whose VOC keeps coming, and produces the summary deck "GCX SIREN <상/하반기> 등록현황 및 VOC 점검" in the apple.com-style Bi-weekly design (bi-weekly-builder §7). Everything goes to the user's PRIVATE chat room only. Trigger on "run SIREN sweeper", "SIREN 반기 sweep", "SIREN 등록현황 및 VOC 점검 만들어줘", "하반기/상반기 SIREN 점검", "make SIREN reports for the top N and the summary deck", or any close paraphrase.
metadata:
  category: automation
  locale: ko-KR
  built: "2026-10-02 — approved deck 1_ncdmLrNjG5WjuCltG_mZcNZke7DM72pCkC5ttlbsa8 (GCX SIREN 하반기 등록현황 및 VOC 점검)"
---

# siren-sweeper — GCX SIREN 반기별 Sweeper

Built and approved 2026-10-02 (manager 박기세 실장님 feedback applied). Reference output:
https://docs.google.com/presentation/d/1_ncdmLrNjG5WjuCltG_mZcNZke7DM72pCkC5ttlbsa8/edit — "if it looks like this, it's right".

## Linked skills (this skill orchestrates them — don't re-implement them)
| skill | used for |
|---|---|
| `siren-finder` | step 0: candidates (`find_siren.py --top N --send`) → `~/.config/siren_finder/runs/candidates_<yymmdd>.json`; `export_case()` skeletons; `claude_json()` |
| `siren-report` | step 5: one SIREN registration deck per case (`build_siren_slides.build_deck`, unchanged; multi-SKU overview layout fixed 2026-10-02) |
| `bi-weekly-builder` | design system + helpers for the summary deck (§7 Apple version: `tools/apple_theme/apple_cards.rtile`, `apple_v3.pill/hair/text`, `apple_type`, `spigen_images`, icons, chart style) and the golden deck it copies |
| `frontend-design` (anthropic) | the aesthetic direction behind the apple.com-style deck (bright canvas, white tiles, capsules) — consult when changing the look |

## Rules (user-set 2026-10-02 — keep)
- **Private room only** (`common.PRIVATE_WEBHOOK`, spaces/AAQAc9NQmJQ). Never post to team rooms unless told.
- Build **all N** SIREN decks the user asks for (default top 20 open candidates), glass included; the **summary deck excludes Screen Protector** cases (SP division — reported by another team).
- Never write "non-SP" / "Non-Glass" — write **"Screen Protector 제외"** and footnote *"Screen Protector(글라스) 제품은 별도 팀에 보고되어 본 자료에서 제외"*.
- No "후보" wording on the summary deck. Use "슬라이드/slide", never "덱/deck" on slides (cards may say 덱).
- **판매량 = Amazon EU only** (DE·FR·IT·ES·NL·SE·BE·IE·PL·UK; no JP/US/IN), labeled everywhere; **VOC율(EU) = EU-country VOC ÷ EU 판매량** (2% ⇒ red). 합계/클레임/리뷰 stay all-country.
- Summary numbering **#1..N in slide order** (sections Case → SDA·New Biz → Power Accessories → 기타), the same on every slide; index shows no #N.
- 기등록 = the case matches a registry row (same SKU line-up + same failure). Show it as a **red capsule button linking to the earlier SIREN material** + the **registration date** (earliest across registries) on the table; on case slides a red capsule `기등록 <src>`. Link rules: GCX rows → 'PPT 링크' chip; **영업 rows → the last 'LINK' column (품질, cells "Link"/"CQ SIREN")**.
- Representative photos: **look at the attachments** (contact sheets) and pick ones that show the defect — never invoices, Amazon/order screenshots, packaging labels.
- Each 최근 VOC has a "원문 보기" button (Zendesk ticket / Amazon review).
- Chart: one rounded shape per bar (same radius), legend squares (클레임 blue / 배드리뷰 light-blue), x labels = full product name + defect row (no #N).
- TOP3 tiles show the **full product name** (e.g. "Apple Watch Series 9 Rugged Armor Pro 파손").
- Table has 생산지 (SKU_Master 생산업체·원산지). The 상태(status) column was tried and **removed by the user** — don't add it back unless asked (statuses are summarised in SKILL notes below if needed).
- Title: **"GCX SIREN <상/하반기> 등록현황 및 VOC 점검"** (file name too). Slide 3 starts with a **작성 취지** block; the share message must state 취지/목적 + 범위.
- **After any build: render every slide and fix unintended text wraps** (memory `feedback_slides_text_wrap_check`) — text boxes lose ~7.2pt per side to insets.

## Workflow (scripts in `scripts/`, each `import common` → work dir `~/.config/siren_sweeper/<yymmdd>/`, override with `SWEEP_DIR`)
0. **Ask** (one AskUserQuestion): how many top candidates (default 20), which siren-finder run (default latest), half label (auto from date).
   If no fresh run: `python3 ~/.claude/skills/siren-finder/find_siren.py --top 20 --send` (private room).
1. `python3 01_export_cases.py --top 20` → case_NN.json, ids.json
2. **Zendesk (browser, claude-in-chrome)** — navigate a tab to `https://spigenhelp.zendesk.com/api/v2/tickets/<any>.json`,
   `python3 02a_zendesk_js.py > zendesk_collect.js` and run its content with `javascript_tool`; poll `JSON.stringify(window.__p)`
   until `fin:true` (~1–2 min / 1000 tickets); run `02b_dump_js.txt`; then `get_page_text(max_chars=800000)` — the harness saves it
   to a `tool-results/*.json` file; `python3 02c_parse_tickets.py <that file>` → tickets.json.
   (Don't pass Google tokens into the page — blocked; don't POST to localhost — PNA prompt hangs the tab.)
3. `python3 03_summarize.py` (re-run until all summarised) → ko.json
4. `python3 04_media.py` → media.json (HEIC→JPEG, videos→Drive). Delete `bad` keys and re-run to retry.
5. `python3 05_build_siren_decks.py 1 2 … 10 & python3 05_build_siren_decks.py 11 … 20` (background) → deck_NN.txt.
   Spot-check slide counts (= claims+reviews+2) and a few thumbnails.
6. `python3 06_registry_check.py` → regrows.json, regmatch.json, regreport.json. **Review regmatch.json by hand**: only `possible:false`
   rows become 기등록; Claude judges with the cases' VOC samples (rule: same failure mode; '힌지 파손'≠'힌지 돌기 파손').
7. `python3 07_send_cards.py 1 … N` — review cards (기등록 notes from regmatch) → private room.
8. `python3 08_send_reg_report.py` — "기등록 SIREN — VOC 지속" report → private room.
9. EU sales: `python3 09a_eu_sales_sql.py` → run the SQL with `mcp__claude_ai_CaspiLM__run_query` → `python3 09b_save_eu_sales.py`;
   then `python3 09_summary_data.py` (default = all non-SP cases; `--cases 1,2,5…` to override) → cases.json
10. `/usr/bin/python3 10_contact_sheets.py` → **Read every cs<no>.jpg**, choose 3 defect photos each, write picks.json
    `{"1":[i,j,k],…}`, `/usr/bin/python3 10_contact_sheets.py --apply`
11. `/usr/bin/python3 11_charts.py` → overview.png, m<deck>.png
12. `python3 12_build_summary.py` (new deck from the golden copy) → summary_deck.txt; `python3 wrapcheck.py`; `python3 render.py` and
    **look at all slides**.
13. `python3 13_send_summary_card.py` → private room. Give the user the share text (취지/범위 + link).

## Editing an existing deck (lessons from 2026-10-02)
- Before rebuilding, **diff the live deck against the expected build** — the user hand-edits (e.g. TOP3 captions, "1. Overview"
  title). If edited: make **targeted in-place edits** (replaceAllText keeps styles; replace one image with `replaceImage`; rebuild
  only the affected slide), never a full rebuild.
- `deleteText`+`insertText` on a button **drops its link** — re-apply `link` after changing button text, then verify every 기등록
  button has a link.
- replaceAllText inherits the first run's font → Korean can end up in Inter; re-run a Hangul→Noto Sans KR pass.
- A font pass must preserve bold: set weight 600/700 only where the run is bold (read the run's `bold`); don't infer element IDs from
  a dry run after data changed (IDs shift).
- Chart image swap: keep it below the "이슈별 VOC" heading/legend (y≥213) and `SEND_TO_BACK`.
- Transient `There was a problem retrieving the image` / 502 → retry (the send() helper does).

## Status notes (2026-10-02 run, for reference if the status column is ever requested again)
개선 불가 (구조 제약) / 조치 완료 (양산 적용) / 노이슈 (외력 추정) / 개선 예정 (금형 개선) / 개선 진행중 / 검토중 (CQ 이관) / 미등록 —
derived from GCX '개선 결과' + 영업 '진행 단계/품질' columns.
