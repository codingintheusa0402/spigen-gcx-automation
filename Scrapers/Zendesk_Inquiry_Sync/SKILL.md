---
name: zendesk-inquiry-sync
description: >-
  Append new Solved Zendesk tickets (fetched from Caspi S3.ZENDESK.*) to the
  "26년 전체문의" tab of the "Zendesk Raw Data_2026년" spreadsheet — A:AD only
  (Ticket ID … On-hold time), deduped by Ticket ID, appended below the existing
  rows, with the formula columns AE~ extended automatically — and schedule it to
  run N times per week at a chosen time. Trigger when the user says "sync
  zendesk to 26년 전체문의", "전체문의 시트 업데이트해줘", "append solved tickets
  to the 2026 zendesk sheet", "set up / change the zendesk inquiry sync
  schedule", "run zendesk inquiry sync", or any close paraphrase.
metadata:
  category: automation
  locale: ko-KR
  phase: v1.0.0-live
---

# zendesk-inquiry-sync

Keeps `26년 전체문의` (gid `483971768`) in spreadsheet
`1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc` up to date with Solved Zendesk
tickets. Code: `sync.py` next to this file (plain Python 3; Caspi data-api over
HTTPS + Google Sheets API v4). Runs locally — the scheduled job is a macOS
launchd agent, so no Claude session is needed once it's set up.

## What gets appended (the rules)

| Rule | Detail |
|---|---|
| Source | Caspi registered query (`query.sql`) over `S3.ZENDESK.TICKETS` + `TICKET_FIELDS` (option value → display name) + `TICKET_METRIC_SETS`. Caspi lags Zendesk by ~1 day. |
| Status | **Solved**. Zendesk auto-flips Solved → **Closed** after a few days (Closed = solved & locked), so the default includes both — otherwise any ticket that closes between two runs would be lost forever. `run --solved-only` = strict `solved`. |
| Year | 2026 KST = Ticket ID ≥ `1000132837` (Caspi's `created_at` is a UTC date, so the boundary is by ID). |
| Channel | `SQ_website`, `Amazon Buyer Message`, `spigen.support@spigen.com` (the only 3 in the sheet). |
| Agent replies ≥ 1 | The sheet never held 0-reply tickets (verified 0 / 22,368). Drops merged duplicates, untriaged Amazon buyer messages, auto-closed noise. |
| Dedupe | Col A (header reads "Last updated" but holds the **Ticket ID**). A ticket already present anywhere in col A is never written again. |
| Columns | Writes **A:AD only** (30 cols, `Last updated` … `On-hold time (min)`). Header A1:AD1 must match exactly or the run aborts without writing. |
| Formula cols AE~ | Not written from data; formulas + format are copied down from the last existing row (relative refs shift), so 충성고객 / SKU / 브랜드명 … fill in automatically. |
| Values | Written RAW: IDs/counts as numbers, dates as real dates (`yyyy-mm-dd` format copied), `Tier 2 Escalated` as TRUE/FALSE, `Agent replies brackets` stays text (`3-5`, `>5`). |

Known caveat: Caspi dates are UTC dates, so `Ticket created/updated - Date` can
be one day earlier than the KST date the old Explore exports showed for tickets
created 00:00–08:59 KST.

## Commands

```bash
S=~/Desktop/GCX/Scrapers/Zendesk_Inquiry_Sync/sync.py   # or wherever the repo is cloned
python3 $S status                     # creds / schedule / launchd / last run
python3 $S run --dry-run              # show what would be appended, write nothing
python3 $S run                        # append now
python3 $S schedule --days mon,thu --time 09:00   # (re)install launchd job, KST
python3 $S schedule --days weekdays --time 08:30  # also: daily
python3 $S unschedule
```

A full run takes ~2–3 min (≈28k rows paged from Caspi; 429s are retried).
Logs: `~/Library/Logs/zendesk-inquiry-sync/sync.log`.

## How to handle requests

**"Run it now"** → `run --dry-run` first, report the count + a few sample
Ticket IDs, then `run`. Report the appended row range and ID range.

**Schedule set/change** → ALWAYS ask the user (AskUserQuestion) for
1. how many times per week / which weekdays (e.g. 월·목, 평일 매일, 월수금), and
2. what time of day (KST, HH:MM),

then run `schedule --days <mon,tue,…|weekdays|daily> --time HH:MM`, and confirm
with `status`. The launchd job ticks every 30 min and `run --scheduled`
self-gates: it runs once on each chosen weekday at/after the chosen time, so a
Mac that was asleep at that time catches up on wake the same day.

## Setup for a new user (once per machine)

Secrets live only in `~/.config/zendesk_inquiry_sync/` (chmod 600) — never in the repo.

1. `bash <repo>/Scrapers/Zendesk_Inquiry_Sync/install.sh` — symlinks this skill
   into `~/.claude/skills/` and installs the Python deps.
2. **Caspi (Zendesk data).** Needs a Spigen Claude account with the Caspi
   connector and Zendesk-table access. Using the Caspi `data_api` tool:
   - `action: register`, `name: "26년 전체문의 sync"`, `params: ["min_ticket_id"]`,
     `sql:` the full contents of `query.sql` → gives a `queryId` (`pq_…`).
     Each user registers their own copy — queries run under the registrant's
     own permissions.
   - `action: key_issue` (`confirm: true` only after the user explicitly says
     to issue a key) → `ak_…`. If the user already has a Caspi API key, reuse it
     (`action: key_list`).
   - `python3 sync.py setup --caspi-key ak_… --query-id pq_…`
3. **Google Sheets.** If `~/.config/gws_shim/token.json` exists it's used
   automatically. Otherwise get the team's OAuth *desktop client* JSON from the
   repo owner (never commit it) and run
   `python3 sync.py setup --client-secret /path/client.json` (browser consent
   with your @spigen.com account; you need edit access to the spreadsheet).
4. `python3 sync.py run --dry-run` to verify, then ask for the schedule (above).

## If something breaks

- `Header A1:AD1 changed — refusing to write` → someone renamed/inserted a
  column; update `HEADER` in `sync.py` (and the SQL/`build_row` if a field moved).
- Only a few hundred rows fetched → the Caspi response is size-capped (~100KB)
  and flags `truncated` with `nextOffset: null`; `caspi_fetch` pages by rows
  received — keep that logic.
- New Zendesk custom field needed → add it to `query.sql`, re-register (new
  `queryId` → `setup --query-id`), and map it in `build_row`.
