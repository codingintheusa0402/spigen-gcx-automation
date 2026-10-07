# TicketReporterCard

Interactive Google Chat app ("Ticket Reporter") that turns the static TCK 전달 전 보고
reports posted by the `ticket-reporter` Claude skill into an in-chat internal-note workflow —
pick a canned phrase, write a freeform note, and post it straight to Zendesk as an internal
note, with a per-sender confirmer code appended automatically.

**Script ID:** `1KK5FzhKq39NrAjETOvgQKja5YQPqS-aJMki1rGuy14CwQnDWwYPz5EOf`
**GCP project:** `gcx-zendesk-decision-maker`
**Deployment ID:** `AKfycbwVTKgXn9WBU4vk-MLuj1wPzv9y_Ru0E_1VSyMwvb4T0PGSiKixJZfOaPM18r73WLVQ` (pinned in Chat API → Connection settings → Apps Script)

---

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | Event handlers (`onMessage`, `onCardClick`), card building, Zendesk API calls, Sheet-based thread↔ticket mapping |
| `Config.gs` | Zendesk subdomain, Script Property keys, `CONFIRMERS` dropdown list, `CONFIRMER_BY_EMAIL` map, the 43-item `REFERENCE_PHRASES` list |
| `appsscript.json` | GAS manifest (Chat app config) |

---

## Five ways to trigger it

### 1. `"티켓 <번호>"` or a pasted report → interactive card
`onMessage` renders a card: the report text, a single reference dropdown (43 canned phrases
that each append a line into one freeform note box), a confirmer dropdown, and a submit
button. `submitNote` posts the note as a Zendesk internal note — **never public**, hard rule.

### 2. Thread-reply direct-post
`send.py` (in the sibling `ticket-reporter` skill) writes `{ticketId, threadId}` into the
`TicketQueue` Google Sheet right after posting each static report webhook. If a human later
replies in that thread mentioning the app with **no ticket number** in the text, `onMessage`
resolves the ticket via `lookupTicketByThread_` (pure Sheet lookup — no `chat.bot` API read
needed, see limitation below) and posts the reply text straight to Zendesk via
`postInternalNoteAndReopen_`, which files the internal note **and** reopens the ticket to
`open` in the same PUT call.

The confirmer code is resolved from the sender's email via `CONFIRMER_BY_EMAIL` in
`Config.gs` (not the `CONFIRMERS` dropdown used by path 1):

| Email | Code | Name |
|-------|------|------|
| `kjw@spigen.com` | `KJW` | 김지우 Kevin |
| `yangsr@spigen.com` | `L` | 양숙랑 Grace |
| `arrha@spigen.com` | `NAR` | 나아름 Jane |

Unmapped senders fall back to `CONFIRMERS[0]` (`KJW`). The confirmation card title is plain
`내부 노트 전송 완료` — no ticket number, emoji, or "스레드 답장" suffix.

### 3. `/revision <feedback>` → feedback channel
A pure feedback channel that never touches Zendesk. `onMessage` checks this prefix **before**
the ticket-number regex, since pasted revision feedback often quotes a full report (which
contains its own ticket numbers) and would otherwise be misread as a ticket trigger.
`handleRevisionFeedback_` logs `{ts, submitterEmail, feedbackText, appliedAt}` to a lazily-created
`Feedback` tab in the same `TicketQueue` sheet and replies with a "피드백 접수 완료" card.

The `ticket-reporter` Claude session applies these: `check_feedback.py list` at the start of
every monitor tick returns unapplied rows (empty `appliedAt`), the session updates
`SKILL.md`'s writing rules accordingly, then `check_feedback.py mark-applied <row>` stamps it
done.

### 4. `/manual` → full SKILL.md text
Checked before the ticket-number regex for the same reason as `/revision`. `handleManualRequest_`
reads the ticket-reporter `SKILL.md` text back out of the `Manual` tab of the same `TicketQueue`
sheet (mirrored there by the monitor session whenever `SKILL.md` changes — the file itself only
lives on the machine running the monitor) and replies with a short preview (1,500 chars) plus a
link button straight to that tab. A full-text card was tried first but silently failed once
`SKILL.md` passed ~24KB (Chat's payload limit, with no error surfaced on the Apps Script side).

### 5. `/btw <question>` → general Q&A, not a ticket action
Anything after `/btw` is a free-form question, not a ticket trigger/report — checked before
the ticket-number regex for the same reason as `/revision`/`/manual`, since a question often
contains its own digit sequences (order numbers, dates, example ticket numbers) that must not
get misread as "open ticket card for this id".

By design (사용자 지시 2026-10-07), this app never calls an LLM API directly — no Anthropic key
lives in Script Properties. `handleBtwQuestion_` just appends `{question, threadId, email}` to
a lazily-created `BtwQueue` tab in the same `TicketQueue` sheet and immediately acks "⏳ 질문
접수됨". The actual answer comes from the **ticket-reporter monitor session itself** — the
already-running Claude Code `/loop` on the server — which checks `BtwQueue` every tick
(`check_btw_queue.py list`), answers using its own reasoning, and replies in-thread with
`send.py --thread "<threadId>"` (no `--ticket-id` — this never touches Zendesk), then stamps
`answeredAt` via `check_btw_queue.py mark-answered <row>`. So the answer lands on the
monitor's normal cadence rather than instantly — worst case ~5 minutes during business hours.

---

## Known hard limitation

`chat.bot` is **not** a user-consentable OAuth scope in Apps Script's authorization flow
(`invalid_scope`). This means the script can never call the Chat REST API to post a message
*as the app itself* — it can only reply within `chatCreate_`/`chatUpdate_` envelopes returned
from an event handler that a human message already triggered. The `TicketQueue` sheet
thread-mapping (path 2 above) exists specifically to work around this: it lets a plain human
thread-reply resolve back to a ticket without the app ever needing to read the thread's
history via the API.

---

## Image grid (in `send.py`, not this app)

Report attachment images render as a `grid` widget (2 columns for exactly 2 images, 3 for 3+)
instead of one `image` widget per row, to keep long reports shorter. `grid.items[]` does not
accept a per-item `onClick` on the webhook endpoint (confirmed: 400 "Unknown name onClick") —
click-through to full size is a "원본1 · 원본2 · …" link row placed below the grid instead.

---

## Deployment

```bash
cd ~/Desktop/GCX/GAS_Operations/TicketReporterCard
clasp push --force
clasp deploy -i AKfycbwVTKgXn9WBU4vk-MLuj1wPzv9y_Ru0E_1VSyMwvb4T0PGSiKixJZfOaPM18r73WLVQ -d "<description>"
```

`clasp push` alone updates script content/HEAD but does **not** move the live pinned
deployment — the `clasp deploy -i` step is required for changes to reach the installed Chat
app.

Live rooms: real chatroom `spaces/AAQAdqYt1ro`; test room `Private` (`spaces/AAQAc9NQmJQ`).
