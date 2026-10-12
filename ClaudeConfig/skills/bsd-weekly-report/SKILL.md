---
name: bsd-weekly-report
description: 사업지원실 '26 주간업무 reporter — fills 김지우's 글로벌CX전략팀 block (column D) of the weekly tab (e.g. `261012`) in the "사업지원실 '26 주간업무" spreadsheet, the way the 261012 version was built and approved on 2026-10-12: [🚨SIREN] items with CQ 피드백 status, then [💡모니터링] blocks for each running series (Galaxy Z Fold 8/Flip 8, Pixel 11, iPhone 18 …) with 누적 counts, ※TOP 3 인입사유 (누적) and ①②③ 금주 SKU×인입사유, plus any extra ■ sections the user pastes. Trigger on "사업지원실 주간업무 써줘", "주간업무 내 파트 채워줘", "make the 261019 weekly", "run bsd weekly report", "draft my 사업지원실 weekly block", or any close paraphrase. ALWAYS ask the three inputs first. Not the GCX WBR sheet — that is wbr-mini-reporter.
---

# 사업지원실 '26 주간업무 — 글로벌CX전략팀 block

Spreadsheet: `1J4l0m4bJCthuGtKTBk51a0UuZcPen0w_x1zyPpFb39U` ("사업지원실 '26 주간업무"). One tab per week named
`yymmdd` (Monday, e.g. `261012`); newest tab sits right after `주간업무 포맷_매주복사`. Team rows: col B = team
name (`글로벌CX전략팀`), the block runs down to the next team's name in col B. **My part is column D**; col E of the same
rows belongs to a teammate (Google store / Shopify / 인도 아마존 …) — never touch it. Other teams edit the tab
live while you work, so row numbers shift: re-read right before writing, never reuse old row numbers.

Scripts: `~/.claude/skills/bsd-weekly-report/scripts/` (gws_shim token, kjw@spigen.com).

## 0. Ask first (one message, all three) — do not start before all are answered
1. **Sheet (tab) name** of the new week (e.g. `261019`). If it doesn't exist yet, say so and ask whether to create
   it — normally the team creates it from `주간업무 포맷_매주복사`; don't create it on your own.
2. **Range of dates to report** (e.g. `10/10 ~ 10/16`). Default suggestion: day after last week's `*M/D 기준`
   date through the Friday before the new tab's Monday. The end date is the `*M/D 기준` date; the start date bounds
   the ①②③ 금주 lines.
3. **Cell range** in the sheet for my block (e.g. `D167:D228`). If empty, it is seeded from last week.

Also ask whether there are extra ■ items to add this week (e.g. budget/기안 news) — the user may paste drafts.

## 1. Read last week + the target
```
cd ~/.claude/skills/bsd-weekly-report/scripts
python3 block.py prev <PREVTAB>          # my block in last week's tab (auto-detected D-range)
python3 block.py read <TAB> <RANGE>      # target; [chip] marks cells holding smart chips (deck/sheet/photo links)
```
PREVTAB = the tab right after TAB in the tab list (one week older).

If the target range is empty → seed it (inserts blank rows above the next team only if needed, copies last
week's block with copyPaste so smart chips + formatting survive; refuses if the range already has text):
```
python3 block.py seed <TAB> <RANGE-start-cell-range> <PREVTAB> <PREVRANGE>
```
Then `block.py read` again to get the real row numbers. If the range already has text (user/team pre-filled),
edit in place.

## 2. Gather
```
python3 siren.py <FROM> <TO> "<SKUs listed in the current [🚨SIREN] section, comma-separated>"
python3 stats.py <FROM> <TO>
```
- **siren.py** → `new_in_window` (registry `26년 SIREN` rows dated FROM..TO — any 작성자, the section lists the
  team's registrations) and `listed_status` (피드백 날짜 / 개선 결과 for listed items; ignore old rows of the same
  SKU with an earlier date).
- **stats.py** (verified 2026-10-12 to reproduce the previously reported numbers exactly):
  - 누적 totals = `결과` tab summed to TO (case/SP; 해외/국내 for Z8 & iPhone 18; Pixel has no KR).
  - `top3_reasons_cum` = `1-3점` rows Created ≤ TO by `인입사유(tag)`.
  - `week_top` = `1-3점` rows FROM ≤ Created ≤ TO by (기종명용 모델명, SKU, 인입사유(tag)); ties ranked by cumulative.
  - 긍정 리뷰 / 기타사항 / blank tags excluded everywhere.
  - Series list lives in `_g.py` `SERIES` — add a new launch / drop an ended one there (and its block section).

## 3. Draft — keep the exact format
```
[🚨SIREN] 
1. <기종>용 <제품> (<SKU>) <이슈> (SIREN 등록, M/D)
<deck title — smart chip to the deck>
    - 이슈: <one line>
   - CQ 피드백 대기중            ← or "   - CQ 피드백 (M/D): <summary of 개선 결과>"

[💡모니터링 - <series>]
<monitoring sheet title — smart chip>
   1) Spigen Case 배드리뷰 인입 건수 (*M/D 기준 누적 총 N건 ㅣ해외: N건 ㅣ국내: N건)
※TOP 3 CASE 인입사유 (누적): <r1> N건 ㅣ <r2> N건 ㅣ <r3> N건
① <기종>용 <제품> (<SKU>) <인입사유> N건      ← 금주 (FROM~TO) counts, NOT cumulative (user rule 2026-10-12)
② …
③ …
   2) Spigen Clear Protection 배드리뷰 인입 건수 (*M/D 기준 누적 총 …)
※TOP 3 SP 인입사유 (누적): …
① ② ③ …
```
Rules:
- **SIREN lifecycle:** new registrations in the window are appended (number continues). An item gets
  "CQ 피드백 (M/D): …" the week the registry shows a 피드백 날짜; the following week it is dropped (as 260928 → 261006).
  New items need the deck link as a chip: find the deck in Drive by its title (`… SIREN 검토_글로벌CX전략팀_yy.mm.dd`)
  and write the chip (see `block.py` chip handling / updateCells with `chipRuns`); if not found, write plain text and tell the user.
- `재질` is written `재질 (*그립감, 미끄러움 등)`. Keep existing parenthetical notes (e.g. `(*9/4 클프개발팀: 제품 이상없음)`)
  when the same SKU×사유 is still listed; drop them with the line otherwise.
- The `*디자인 … / L 사진: <photo chip> *사진 업데이트 완료 >> > N건 (~M/D)` lines under Z8 case: N = 디자인 누적 count.
- Notes like `L ★… 클프개발팀 확인 …` stay; refresh any count inside them.
- Extra ■ sections the user gives (e.g. "■ 구글 제미나이 배드리뷰 키워드 추출AI 토큰 …"): business tone, `  1.` / `      - ` / `L ` levels,
  facts only from what the user pasted (기안 content, amounts, dates); ask about anything ambiguous (e.g. 200 vs 250 USD).
- Never invent work or numbers; if a number can't be computed, leave last week's and flag it.

## 4. Write (auto-backup, chip-safe)
Put `{cell: text}` for changed cells in a JSON file (scratchpad), then:
```
python3 block.py write <TAB> edits.json
```
Backs up old text to `~/.config/bsd_weekly_report/backup_*.json`, writes plain cells RAW, and rewrites chip cells
keeping the chip (its display text must stay verbatim in the new text). Afterwards run `block.py read` and check
every chip from last week is still `[chip]` — users' manual edits sometimes drop them; restore from PREVTAB.

## 5. Report
Give the clickable link `https://docs.google.com/spreadsheets/d/1J4l0m4bJCthuGtKTBk51a0UuZcPen0w_x1zyPpFb39U/edit#gid=<gid>&range=<RANGE>`,
a short table of the new totals / ①②③, and what to double-check (tie-breaks on 1건 weeks, dropped notes,
SIREN status changes, anything inferred in extra ■ sections).
