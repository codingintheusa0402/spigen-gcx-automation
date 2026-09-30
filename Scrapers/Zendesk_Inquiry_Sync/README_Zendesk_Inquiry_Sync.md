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

Quick start: `bash install.sh`, then in Claude Code: "set up zendesk inquiry sync".
Full rules and per-user setup (own Caspi query + key, Google token) are in `SKILL.md`.
Per-user secrets/state: `~/.config/zendesk_inquiry_sync/` (never committed).
