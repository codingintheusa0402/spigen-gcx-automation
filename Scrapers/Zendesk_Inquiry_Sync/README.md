# Zendesk_Inquiry_Sync

`sync.py` appends Solved Zendesk tickets, including auto-Closed ones, from Caspi (`S3.ZENDESK.*`) to the `26년 전체문의` tab of `Zendesk Raw Data_2026년` (`1sjcCj_P4DRD8rywkmYJhbsrzwFfgiJQuF9nIKwCiKlc`). It writes only A:AD, dedupes by Ticket ID (col A), adds rows below the existing ones and copies the formula columns AE~ down. It runs on a weekly schedule each user picks, through launchd on macOS or Task Scheduler on Windows, and can post a Google Chat notice after each scheduled run.

This is the backend of the Claude Code skill **`zendesk-inquiry-sync`** ([`SKILL.md`](SKILL.md): rules, per-user setup, commands). `README_Zendesk_Inquiry_Sync.md` is the older short version of this file.

## Screenshots

![`python3 sync.py status`](docs/status.jpg)
*`python3 sync.py status`*

## What gets appended

- Tickets created in 2026 KST. The year boundary is the first 2026 ticket ID, `MIN_TICKET_ID = 1000132837`, because Caspi's `created_at` is a UTC date.
- Status `solved` or `closed`. `run --solved-only` keeps only `solved`.
- Channels `SQ_website`, `Amazon Buyer Message` and `spigen.support@spigen.com`.
- At least one agent reply. This also drops merged duplicates and auto-closed noise; on 2026-09-30, none of the sheet's 22,368 rows had 0 replies.
- Not already in col A of the sheet.

Before writing, the script checks that header A1:AD1 matches its built-in `HEADER` exactly. If it doesn't, it refuses to write. Col A is labelled "Last updated" but holds the Ticket ID (a Zendesk Explore export quirk).

## Files

| File | Purpose |
|---|---|
| `sync.py` | CLI: `run` / `setup` / `schedule` / `unschedule` / `status` |
| `query.sql` | Caspi registered-query SQL for the ticket rows (param `min_ticket_id`) |
| `intake_query.sql` | Optional Caspi registered query (param `created_date`): ticket counts per status for one day, used for the notice's intake line |
| `SKILL.md` | Claude Code skill `zendesk-inquiry-sync` |
| `install.sh` | Symlinks `SKILL.md` into `~/.claude/skills/zendesk-inquiry-sync/` and pip-installs `requirements.txt` |
| `requirements.txt` | `google-api-python-client`, `google-auth`, `google-auth-oauthlib` |

## Usage

```bash
python3 sync.py status                       # config, schedule, launchd/Task state, last run (read-only)
python3 sync.py run --dry-run                # fetch + show the first rows it would append, write nothing
python3 sync.py run                          # append now
python3 sync.py run --until-yesterday        # only tickets created up to yesterday (KST)

python3 sync.py setup --caspi-key <key> --query-id <pq_…> \
    [--client-secret <oauth_client.json>] [--chat-webhook <url>] [--intake-query-id <pq_…>]
python3 sync.py schedule --days mon,thu --time 09:00 [--until-yesterday]   # also: weekdays | daily
python3 sync.py unschedule
```

`run --scheduled` is what the scheduler calls. The job wakes every 30 minutes (launchd `StartInterval` 1800) or every 5 minutes (Task Scheduler). It runs only on a scheduled day at or after the scheduled time, at most once per day. If the computer was off at the scheduled time, it catches up within minutes of waking. A file lock prevents overlapping runs.

## Chat notice

When `--chat-webhook` is set up, each scheduled run posts a plain-text notice: "✅ Zendesk Raw Data 업데이트 완료". It lists the run time, the intake for the previous day(s) (completed vs in progress), how many rows were added to the sheet, and a link to the tab. The intake line needs `--intake-query-id`. After unscheduled days it covers all of them; for example, Monday on a weekdays schedule shows the Fri–Sun total. A failure notice is posted at most once per day.

## Config & state (per user, never committed)

| Path | Content |
|---|---|
| `~/.config/zendesk_inquiry_sync/credentials.json` | Caspi API key, queryIds, Chat webhook |
| `~/.config/zendesk_inquiry_sync/google_token.json` | Google OAuth token. Falls back to `~/.config/gws_shim/token.json`, or `$GOOGLE_TOKEN_PATH` |
| `~/.config/zendesk_inquiry_sync/schedule.json` / `state.json` / `run.lock` | Schedule, last run, lock |
| `~/Library/LaunchAgents/com.spigen.gcx.zendesk-inquiry-sync.plist` | launchd job (macOS) |
| Task Scheduler "Spigen GCX Zendesk Inquiry Sync" | Windows job |
| `~/Library/Logs/zendesk-inquiry-sync/` (Windows: `%LOCALAPPDATA%\zendesk-inquiry-sync\Logs`) | Logs |

Set the env `ZIS_SPREADSHEET_ID` to point the script at a test copy of the spreadsheet.

## Quick start (teammates)

```bash
git clone git@github.com:spigenHQ/HQ_GCX.git ~/HQ_GCX
bash ~/HQ_GCX/Scrapers/Zendesk_Inquiry_Sync/install.sh
```

Then in Claude Code say **"set up zendesk inquiry sync"**. Claude registers your own Caspi query and key, checks them with a dry run, and asks which days per week and what time (KST) the sync should run.

## Recent changes (2026-09-30 – 10-01)

- First version: the Caspi → `26년 전체문의` Solved-ticket sync skill, which works from any clone location and asks for a schedule on first use.
- Windows support (Task Scheduler checks every 5 min, runs on battery and catches up) and the `--until-yesterday` mode.
- Google Chat notice after scheduled runs. It now uses bullets, shows yesterday's intake, gives a clearer appended count, and reports intake for unscheduled days in the next run's notice.
