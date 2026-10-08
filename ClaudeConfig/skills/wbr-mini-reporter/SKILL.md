---
name: wbr-mini-reporter
description: WBR mini reporter — rewrites 김지우's own column block of the weekly GCX meeting sheet (WBR, "GCX_weekly meeting 25.03.26~") from what was actually done in the last 7 days (Claude Code sessions, git commits, memory, SIREN registry). The new week's tab starts as a copy of last week's, so the block holds already-reported content; this overwrites it with this week's version in the same format. Trigger on "run WBR mini reporter", "WBR 내 파트 업데이트해줘", "주간보고 내 칸 채워줘", "update my weekly report portion", or any close paraphrase. ALWAYS ask the three inputs first.
---

# WBR mini reporter

## 0. Ask first (one message, all three)
1. **Cell range** of the user's portion (e.g. `F14:F67`) — a single column block.
2. **Sheet (tab) name** of the new week (e.g. `10.08 (목)`).
3. **Spreadsheet URL** (default seen so far: `1jNqTdtR40Prp-qUINHtZg9B8k-8CNT2ylc36UQwBsDY`, "GCX_weekly meeting 25.03.26~").

Do not start until all three are answered. Confirm the header cell of the range (row 1 of the block) reads **김지우** — columns rotate between weeks (10.01 = E, 10.08 = F). If it doesn't, stop and ask.

## 1. Read current block + previous week
```
python3 ~/.claude/skills/wbr-mini-reporter/scripts/wbr_sheet.py read <SID> "<TAB>" <RANGE>
```
Prints every cell with its row label (cols A–C: KPI/SIREN/MCF/Review/프로젝트/금주/차주/스케쥴/연차). Also read the previous tab (the one right after it in the tab list) for the same person's column to see what was already reported — never repeat it as "this week" news.

Window = previous report date+1 … new report date (normally the last 7 days).

## 2. Gather what actually happened (in parallel)
- **Sessions:** `python3 ~/.claude/skills/wbr-mini-reporter/scripts/collect_activity.py <SINCE> <UNTIL> --exclude <this session id>` — the user's own prompts per session. Ignore `claude -p` classifier prompts (they start "You classify…/You label…/For each SIREN candidate…") and scratchpad sweeps — those are automated runs, only count them as evidence a job ran.
- **Git:** `cd ~/Desktop/GCX && git log --since=<SINCE> --pretty='%ad %s' --date=short` — gives shipped versions/fixes with dates.
- **Memory:** files in `~/.claude/projects/-Users-kevinkim/memory/` modified in the window (`find … -newermt <SINCE>`) — gives outcomes and root causes.
- **SIREN count:** read the `26년 SIREN` tab (`15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI`, gid 1840076165); count rows with 담당자 = 김지우 dated in the window (금주) and in the current half (누적). A request for review sent to Chat (no registry row) is "검토 요청", not 등록.

## 3. Draft — same format as the existing cells
- Keep the rows' roles; only change cells whose content is time-bound:
  - SIREN 금주/누적 (numbers + titles as in previous weeks), MCF numbers only if they changed.
  - **금주 (2. 업무 진척)**: the main cell. Sections in brackets — `[Zendesk T2 티켓 처리]`, `[SIREN]`, `[Bi - Weekly]`, `[리뷰 수집]`, `[배드리뷰 Chat Broadcasting]`, `[GCX Reply]`, `[MCF Tracking]`, `[자동화 인프라]` … only the ones with real work this week. `- ` bullets, `    - ` sub-bullets or `L ` lines, completion date `(M/D)` at the end of each item. Korean, business tone, outcome first (what changed for the team), no code/file names.
  - **차주 계획**: 2–4 `- ` lines; carry over ongoing items (e.g. GStore 고객 문의 대응 운영).
  - **스케쥴 금주/차주**: last week's "차주" meetings become this week's "금주"; clear stale ones. There's no calendar access — mark guesses and tell the user to check.
  - **연차/교육**: drop items whose date has passed; keep upcoming ones; add public holidays in the window.
- Leave static cells (Review 관리 sheet list, 프로젝트 PJ명, monday link) untouched unless the user says otherwise.
- Never invent work. If something is ambiguous (e.g. which SIREN deck was updated), state it generically.

## 4. Write (auto-backup)
Put `{cell: value}` for changed cells in a JSON file, then:
```
python3 ~/.claude/skills/wbr-mini-reporter/scripts/wbr_sheet.py write <SID> "<TAB>" <RANGE> edits.json
```
It backs up the whole block to `~/.config/wbr_mini_reporter/backup_*.json` first, writes with `RAW` (values only, formatting kept), and prints the changed cells. Write only inside the given range.

## 5. Report
Tell the user: changed cells, a short summary of the 금주 cell, the backup path, and what to double-check (meetings/schedule guesses, SIREN counts). Link the tab: `https://docs.google.com/spreadsheets/d/<SID>/edit#gid=<gid>`.
