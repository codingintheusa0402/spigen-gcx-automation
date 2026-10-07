# BadReview → Google Chat app-card report

Daily bad-review (★1~3) report for the review-monitoring sheets (**iPhone 18 Series**,
**Galaxy Z8 Series**, **Pixel 11 Series**), sent to the GCX cross-team Google Chat rooms
as Google Chat **cardsV2** app cards. Since 2026-09-21 the default format is **one
swipeable carousel message per room** (iPhone 18 → Z8 → Pixel 11), sent automatically
every weekday at 10:30 KST by `auto_broadcast.py`.

This folder is the version-controlled twin of the Claude Code skills
`pixel11-badreview-chat-report`, `glxz8-badreview-chat-report`,
`iphone18-badreview-chat-report` and `badreview-chat-broadcast` (`~/.claude/skills/…`).

## Screenshots

![Daily 배드리뷰 carousel card (iPhone 18 → Galaxy Z8 → Pixel 11), shown in the private test room](docs/carousel.jpg)
*Daily 배드리뷰 carousel card (iPhone 18 → Galaxy Z8 → Pixel 11), shown in the private test room*

## Files

| File | Purpose |
|------|---------|
| `auto_broadcast.py` | **Production path.** Unattended weekday broadcast of the 3-product carousel to all 13 rooms (no test-send). Imports card/room/carousel code from the skills — see [`AUTO_BROADCAST.md`](AUTO_BROADCAST.md). |
| `badreview_chat_report.py` | Standalone manual sender for the older **2-product, separate-card** format (Pixel 11 + Z8 only, 12 rooms). Self-contained (own card builder + room list). Kept as a fallback; iPhone 18 / carousel / significance highlighting are **not** in it. |
| `chat_app/` | Interactive Google Chat app (Kevin identity → Pixel 11) with date-range + 국가/기종 filters. See `chat_app/SETUP.md`. |
| `chat_app_jane/` | Same app, second identity (Jane → Galaxy Z8). See `chat_app_jane/README.md`. |
| `logs/`, `state/` | runtime log + `held_/ran_<date>.flag` markers (not committed). |

## Where it runs (2026-10-07)

The scheduled jobs moved to the 24/7 **GCX server** (`gcx-server`, WSL2) crontab —
`ServerBootstrap/crontab.txt`:

```
30 10 * * 1-5  python3 auto_broadcast.py                 # main run
30 11 * * 1-5  python3 auto_broadcast.py --retry-if-held # Z8 KR-gate deadline (main + 1 h)
*/5 * * * *    python3 auto_broadcast.py --catchup        # missed-trigger catch-up
```

The Mac launchd copies (`com.spigen.gcx.badreview-broadcast{,-retry,-catchup}.plist`)
are parked in `~/Library/LaunchAgents/disabled-moved-to-server/`. **Never enable both
sides** — the rooms would get double broadcasts.

## What a card shows

| Part | Content |
|------|---------|
| title | `✔️ M/D(요일) <product> 배드리뷰 (1~3점) (총 N건)` (carousel: outer header `✔️ M/D(요일) 배드리뷰 (1~3점)`, page title `<product> 배드리뷰 (총 N건)`) |
| subtitle / image | `고객 리뷰 ★1~3점 · <date range>` + product thumbnail (carousel pages: `image` widget, squared through `wsrv.nl`) |
| **Top 5 인입사유(누적)** | per `대분류` (`휴대폰보호필름` red `#EA4335`, `휴대폰케이스` blue `#4285F4`): `{tot}건` then 5 rows `n위 · 이유 · c건 · p%`. Counted over the whole `1-3점` tab. |
| **오늘 M/D(요일) 최다 인입사유** | top `인입사유(tag)` among today's rows + fixed 5-line breakdown (6th+ → `…외 N건`). The skill builders also highlight a significant day vs. the trailing-7-day average (`recentAvg`). |
| button | **배드리뷰** → that product's `1-3점` sheet |

`N` = rows whose `Update 날짜` (fallback `Exported Date`) is today (KST).
Rows tagged **`긍정 리뷰`** are dropped before any counting (`EXCLUDED_TAGS`; permanent
user rule 2026-09-08). Inside a carousel, `decoratedText` and `columns` don't render, so
`carousel.py` flattens them into `textParagraph` lines.

## Data source

`1-3점` tab of each spreadsheet, read via **Sheets API v4** with the gws_shim OAuth token
`~/.config/gws_shim/token.json` (GCP project `gcxbot`, `kjw@spigen.com`; refreshed and
written back on every run). Columns: `Update 날짜`, `인입사유(tag)`, `대분류`, `국가(tag)`.

| Product | Spreadsheet ID |
|---------|----------------|
| iPhone 18 Series | `1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU` |
| Galaxy Z8 Series | `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4` |
| Pixel 11 Series | `12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI` |

`1-3점` tab gid = `970309432` on Pixel 11 / Z8.

## auto_broadcast.py — behaviour summary

Full detail in [`AUTO_BROADCAST.md`](AUTO_BROADCAST.md).

- Skips weekends and Korean public holidays (Nager.Date API, hardcoded 2026 fallback).
- Waits for the AI tagger: if any of today's rows has a blank `인입사유(tag)`, re-checks
  every 10 min, up to 3 times, then sends anyway (2026-09-18).
- **Z8 KR gate** (2026-09-18): 0 `국가(tag)=KR` rows in today's Z8 data → the whole
  carousel is held, an alert + preview goes to the private test room only, and
  `state/held_<date>.flag` is written. The **11:30 (main + 1 h) `--retry-if-held`** run then sends
  unconditionally (2026-10-01).
- **`--catchup`** (2026-10-02): every 5 min, weekdays 10:30–18:00, runs the normal flow
  if `state/ran_<date>.flag` is missing (machine was asleep at 10:30).
- **Run lock** (2026-10-07): every real run holds `state/running_<date>.lock` while it works (stale after 90 min). Any other real run that starts meanwhile — `--catchup`, `--retry-if-held` or a manual one — logs `SKIP … another run is in progress` and exits. Added after the 10:30 run (15 min waiting on tags) and the 10:33 catch-up both sent the carousel to all 13 rooms on 2026-10-07.
- Room list (13 rooms, incl. `GCX전략 Spigen x TCK` added 2026-09-29) and the test room
  come from `~/.claude/skills/badreview-chat-broadcast/broadcast.py`.

```bash
python3 auto_broadcast.py --test-only --dry-run   # build only, lists target = test room
python3 auto_broadcast.py --test-only             # send carousel to the private test room ONLY
python3 auto_broadcast.py --force --ignore-kr-gate  # LIVE manual resend after a KR-gate hold
```

**Testing rule (permanent, 2026-09-21): always `--test-only`.** Never a bare run or
`--force` with a fake `--date` — that leaked a wrongly-dated card to all live rooms on
2026-09-18.

## badreview_chat_report.py — manual 2-product sender

```bash
python3 badreview_chat_report.py --dry-run --print-data   # build + print cardsV2 JSON, send nothing
python3 badreview_chat_report.py --test                    # both cards → TEST room only
python3 badreview_chat_report.py --broadcast --yes [--only "ADS1,JP Sales"]
python3 badreview_chat_report.py --test --product glxz8 --date 2026-09-02
```

`--broadcast` refuses to run without `--yes`. Per-product webhook routing: each room has a
default webhook and (except 리더들방) a separate `glxz8` webhook used only for the Z8 card
(`room_url()`). Webhook messages cannot be edited or deleted once sent.

Requirements: Python 3.9+, `google-auth`, `google-auth-oauthlib`,
`google-api-python-client`, a valid gws_shim token.

## chat_app/ + chat_app_jane/ — interactive Chat app

Apps Script **Google Chat app (Workspace add-on mode)**, live since 2026-09-11, running
*alongside* the webhook broadcast (webhooks can't receive picker events). One product per
identity (`Config.gs` → `APP_PRODUCT`): Kevin = `pixel11`, Jane = `glxz8`. Control card
has 시작일/종료일 pickers, 국가(`국가(tag)`) and 기종(`기종명`) dropdowns and **[조회]**;
the report card is scoped to the chosen range + filters (unlike the cumulative Top 5 of
the webhook card). Text `9/1~9/11` or `9/11` in a message also works.

Entry points: `onMessage`, `onAddToSpace`, `onRemoveFromSpace`, button handler
`refreshReport` (`onCardClick` = legacy shim). Card builders: `buildCards_(q)` →
`controlCard_` / `reportCard_`. Scope: `spreadsheets` only.

Deploy: edit `chat_app/Code.gs`, copy it to `chat_app_jane/`, then in each folder
`clasp push --force && clasp deploy -i <deploymentId>` (IDs in `chat_app/SETUP.md`).

## Safety flow (manual runs)

1. Re-read the sheets every run (never reuse stale numbers).
2. Test → private test room (`AAQAc9NQmJQ`) only.
3. Show the result + room list, get an explicit human "yes".
4. Broadcast.
5. A product with 0 rows today still sends ("오늘 업로드된 배드리뷰 없음").
