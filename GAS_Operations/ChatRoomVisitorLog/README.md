# ChatRoomVisitorLog

Google Apps Script project that keeps a running log of **who visits the GCX Google Chat rooms** — who is in each room, who joined, who left, and who posted — in a Google Sheet.

**Script ID:** _(create with `clasp create` — see Setup)_
**Log spreadsheet:** auto-created on first run as `GCX Chatroom Visitor Log` (ID stored in Script Properties → `LOG_SPREADSHEET_ID`)

---

## What "visited" means here

Google Chat has no read receipts / "who opened the room" API for other people. The
two signals the Chat API *does* expose are membership and messages, so the log records:

| Event | Meaning | Timestamp |
|-------|---------|-----------|
| `BASELINE` | First run for a space — every current member written once so the roster has a starting point | run time |
| `JOINED` | A person became a member of the space | membership `createTime` from Chat (when they actually joined) |
| `LEFT` | A person is no longer a member | run time (detected on the next poll) |
| `POSTED` | A person sent a message in the space | message `createTime` |

Bots / Chat apps are skipped everywhere.

---

## Files

| File | Purpose |
|------|---------|
| `Code.js` | `logChatRoomVisitors()` (trigger target), `listMySpaces()`, `setupChatRoomVisitorLogTriggers()`, `test_dryRunChatRoomVisitors()`, `clearMessageCheckpoints()` |
| `appsscript.json` | GAS manifest — enables the **Chat** and **People** advanced services and declares the read-only Chat scopes |

---

## Sheet layout

| Tab | Columns | Notes |
|-----|---------|-------|
| `Visitor_Log` | `Timestamp · Space · Space ID · Event · User · Email · User ID · Detail` | Append-only. `Detail` = role for BASELINE/JOINED, `first seen …` for LEFT, `<message id> · <80-char snippet>` for POSTED |
| `Members` | `Space · Space ID · User · Email · User ID · Role · Joined (Chat) · First Seen · Last Seen` | Current roster, rewritten every run. `First Seen` is preserved across runs |
| `Users` | `User ID · Name · Email · Resolved At` | Cache of `users/<id>` → name/email so the People API isn't hit every run |

"Unique visitors per day" is a pivot on `Visitor_Log`: rows = `Timestamp` (group by day), values = `COUNTUNIQUE` of `User ID`, filter `Event` = `POSTED` (or `JOINED`).

---

## How it runs

Time-based trigger (default **every 10 min**). Each run, per space in `SPACE_IDS`:

1. `Chat.Spaces.Members.list` → current human members, diffed against the `Members` tab → `BASELINE` / `JOINED` / `LEFT` rows.
2. `Chat.Spaces.Messages.list` with `createTime > <checkpoint>` (oldest first, max 500 per run) → `POSTED` rows. The checkpoint lives in Script Properties as `MSG_CKPT_<space id>`; the first run looks back `INITIAL_LOOKBACK_HOURS` (24h).
3. Appends to `Visitor_Log`, rewrites `Members`, flushes the `Users` cache, advances the checkpoint.

A `LockService` lock prevents overlapping runs. Name/email resolution order: `Users` cache → `displayName` sent by Chat → People API (`people/<id>`, domain profile; best-effort, needs the `directory.readonly` scope) → raw `users/<id>`.

---

## Script Properties

| Key | Required | Meaning |
|-----|----------|---------|
| `SPACE_IDS` | ✅ | Comma-separated space IDs, e.g. `spaces/AAAAabc123,spaces/AAAAdef456`. Run `listMySpaces()` to print the ones you belong to |
| `LOG_SPREADSHEET_ID` | – | Target spreadsheet. Left empty → created on the first real run and stored here |
| `LOG_MESSAGES` | – | `false` to log membership only (no `POSTED` rows). Default `true` |
| `LOG_MESSAGE_TEXT` | – | `false` to store only the message ID in `Detail`, no text snippet. Default `true` |
| `INITIAL_LOOKBACK_HOURS` | – | Message look-back on the very first run per space. Default `24` |
| `POLL_MINUTES` | – | Trigger interval used by `setupChatRoomVisitorLogTriggers()`: 1 / 5 / 10 / 15 / 30. Default `10` |

---

## Setup

The script calls the Chat API **as the authorizing user** (no bot needed), but Google still requires a
configured Chat app on the attached GCP project before any Chat API call succeeds. The `gcxbot`
project (#64325928759) already has one (the `T2 Report` app), so reuse it.

1. Create the project and push:
   ```bash
   cd ~/Desktop/GCX/GAS_Operations/ChatRoomVisitorLog
   clasp create --type standalone --title "ChatRoomVisitorLog"
   clasp push --force
   ```
2. Apps Script → **Project Settings → Google Cloud Platform (GCP) Project → Change project** → enter `64325928759` (`gcxbot`). Make sure **Google Chat API** and **People API** are enabled in that project (APIs & Services → Enable APIs).
3. In the editor run **`listMySpaces`** once. Approve the OAuth prompt (Chat read-only, directory read-only, Sheets). The log shows `spaces/… | SPACE | <room name>` for every room the account is in.
4. **Project Settings → Script Properties** → add `SPACE_IDS` with the rooms to track.
5. Run **`test_dryRunChatRoomVisitors`** — prints what would be logged, writes nothing.
6. Run **`logChatRoomVisitors`** once by hand → creates the spreadsheet (URL in the execution log) and writes the `BASELINE` roster.
7. Run **`setupChatRoomVisitorLogTriggers`** → recurring trigger every `POLL_MINUTES`.

The authorizing account must be a **member of every space** in `SPACE_IDS` — the Chat API only returns rooms the caller belongs to.

---

## Limits / notes

- `LEFT` is detected on the next poll, so its timestamp is up to `POLL_MINUTES` late. `JOINED` uses Chat's own join time and is exact.
- Someone who opens the room but never posts only ever shows up as `BASELINE`/`JOINED` — there is no per-user read signal in the API.
- If a space gets more than 500 new messages between polls the remainder is picked up on the following run (checkpoint carries over).
- The Chat API read quota is per user; at 10-min polling across a handful of rooms this is nowhere near it.
- Email lookup silently falls back to name-only if the People/Directory call is refused (e.g. directory sharing off for the domain); the row still gets written.
- `clearMessageCheckpoints()` makes the next run re-read the last `INITIAL_LOOKBACK_HOURS` of messages — expect duplicate `POSTED` rows for that window.
