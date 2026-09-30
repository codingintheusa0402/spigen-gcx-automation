# Zendesk_Inquiry_Sync

Appends Solved (incl. auto-Closed) Zendesk tickets from Caspi to
`Zendesk Raw Data_2026년` → `26년 전체문의` (A:AD, dedupe by Ticket ID, formula
columns AE~ extended), on a user-chosen weekly schedule via launchd.

| File | Purpose |
|---|---|
| `sync.py` | `run` / `setup` / `schedule` / `unschedule` / `status` |
| `query.sql` | Caspi registered-query SQL (param `min_ticket_id`) |
| `SKILL.md` | Claude Code skill `zendesk-inquiry-sync` (rules, setup, commands) |
| `install.sh` | Symlink skill into `~/.claude/skills/` + pip deps |

Quick start (teammates):

```bash
git clone git@github.com:spigenHQ/HQ_GCX.git ~/HQ_GCX
bash ~/HQ_GCX/Scrapers/Zendesk_Inquiry_Sync/install.sh
```

Then in Claude Code say **"set up zendesk inquiry sync"** — Claude registers your
own Caspi query/key, verifies with a dry run, and asks which days per week and what
time (KST) you want it to run.
Full rules and per-user setup (own Caspi query + key, Google token) are in `SKILL.md`.
Per-user secrets/state: `~/.config/zendesk_inquiry_sync/` (never committed).
