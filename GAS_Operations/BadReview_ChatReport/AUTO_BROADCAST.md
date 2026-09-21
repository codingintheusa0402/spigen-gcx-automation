# Auto-broadcast (unattended weekday schedule)

Sends the Pixel 11 + Galaxy Z8 + iPhone 18 (added 2026-09-21) 배드리뷰(1~3점) cards to
all 12 GCX rooms **every weekday at 10:30 AM KST**, skipping Korean public holidays
automatically — **no test-send, no confirmation prompt.** Added 2026-09-15 per
explicit user request. This is a *separate*
path from the interactive `badreview-chat-broadcast` skill, which still requires a
test-send + explicit "yes" on every manual run — that hard rule is untouched.

**Combined carousel format (2026-09-21, default):** all 3 products go out as **ONE
message per room** — a swipeable Cards v2 carousel, pages iPhone 18 → Galaxy Z8 →
Pixel 11 — via each room's default `token` webhook, built by `carousel.py` in
`badreview-chat-broadcast/` (shared with the interactive skill). Replaces the old
3-separate-messages-per-room format. Because it's now one message, the Z8 KR-gate
(below) holds back the **entire carousel**, not just Z8, when it trips — there's no
way to omit a single page from an already-sent message.

## Files

| File | Purpose |
|------|---------|
| `auto_broadcast.py` | the unattended script `launchd` runs |
| `~/Library/LaunchAgents/com.spigen.gcx.badreview-broadcast.plist` | the schedule (Mon–Fri, 10:30 local time = KST) |
| `logs/auto_broadcast.log` | one line per run: skipped (weekend/holiday) or per-room OK/ERR |
| `logs/launchd.out.log` / `launchd.err.log` | raw stdout/stderr from launchd itself |

## How it works

1. `launchd` fires the script at 10:30 AM every weekday (`StartCalendarInterval`, one
   entry per Weekday 1–5). Requires the Mac to be **on and awake** at that time — if
   it's asleep/off, that day's run is simply skipped (launchd does not queue/catch up
   missed fires for `StartCalendarInterval`, unlike `cron`'s behavior on some systems).
2. The script checks `today.weekday() >= 5` (weekend safety net, launchd shouldn't fire
   then anyway) and calls the free **Nager.Date API**
   (`https://date.nager.at/api/v3/PublicHolidays/{year}/KR`) for that year's Korean
   public holidays. If today is in that set → log `SKIP` and exit, nothing sent.
   - Nager's KR list does **not** include 근로자의날/Labour Day (May 1) — matches
     "national holiday" (관공서 공휴일) intent, not "day off for private companies."
     If that's wrong for GCX's actual calendar, adjust `kr_holidays()`.
   - If the API call fails (network hiccup), falls back to a **hardcoded 2026 list**
     baked into the script (`KR_HOLIDAYS_FALLBACK_2026`) — re-derive this every
     January from the same API for the new year, or the fallback silently stops
     covering real holidays.
3. Otherwise: refreshes the `gws_shim` Sheets API token, then for each sheet checks
   whether every row with today's `Update 날짜` already has its `인입사유(tag)` filled
   in (the AI tagging agents can lag behind newly-added rows — this is exactly what
   triggered a manual Z8 resend on 2026-09-18). If not, **waits 10 minutes and
   rechecks, up to 3 retries** (30 min max) before giving up and sending anyway —
   logging how many rows were still untagged rather than staying silent about it.
   Then reads both sheets' `1-3점`
   tab directly (Sheets API v4 — **not** the browser/`gviz` method the interactive
   skill uses, since there's no Chrome session in an unattended launchd run), computes
   the same `{todayCount, todayTags, recentAvg, film, case}` shape as the interactive
   flow (including the `recentAvg` trailing-7-day baseline for significance
   highlighting), then **imports** (doesn't duplicate) `report.py` from all three
   `~/.claude/skills/{pixel11,glxz8,iphone18}-badreview-chat-report/` for card-building,
   `broadcast.py` from `~/.claude/skills/badreview-chat-broadcast/` for the room list
   + POST helper, and `carousel.py` (same directory) to assemble the 3 cards into one
   combined swipeable message. Any card-layout, room-list, or carousel-layout change
   made to those files takes effect here automatically — nothing to keep in sync
   manually.
4. **Z8-only KR gate** (2026-09-18, updated 2026-09-21 for the combined carousel):
   checks whether any of today's Z8 rows have `국가(tag)` == `KR` — Z8's single
   largest country segment, unlike Pixel 11 which has none. KR reviews occasionally
   upload after 11 AM (past this 10:30 run), so 0 KR rows is treated as "maybe still
   incomplete," not a real zero day. If so, the script **does not** broadcast the
   carousel to the 12 rooms at all (Pixel 11 and iPhone 18 no longer send separately
   either, since it's one message now) — it posts an alert plus a full carousel
   preview to the **private test room only**. Resend once you've confirmed it's a
   real zero day, or once KR reviews land, with:
   ```bash
   python3 auto_broadcast.py --force --ignore-kr-gate
   ```
   ⚠️ `--force` only bypasses the weekday/holiday skip — it does **not** hold back
   Pixel 11, and does **not** need `--ignore-kr-gate` to still send PX. (Learned the
   hard way 2026-09-18: testing the KR-gate alert with `--force --date 2026-09-21`
   sent a real, wrongly-dated Pixel 11 card to all 12 rooms.)

## Testing — NEVER a bare live run

**Rule (2026-09-21, permanent): when testing anything in this script, always pass
`--test-only` — never a bare run, and never `--force` alone with a fake `--date`.**
`--test-only` sends the combined carousel to the private test room only (space `AAQAc9NQmJQ`) and
never touches `broadcast.ROOMS`, no matter what `--date` is given — it implies
`--force` (a test send shouldn't also get skipped by the weekday/holiday check) and
disables the KR-gate hold (nothing to hold back when it's already private). This is
what should have been used on 2026-09-18 instead of `--force --date 2026-09-21`
without `--test-only`, which leaked a live card to all 12 rooms.

```bash
python3 auto_broadcast.py --test-only --dry-run --date 2026-09-27   # inspect first
python3 auto_broadcast.py --test-only --date 2026-09-27             # then actually send, still private-only
```
5. Posts Z8 (unless held by the KR gate) then PX to all 12 rooms (same order/pacing as
   the interactive `--all`), logs each result.

## Manual controls

```bash
# See what launchd currently has loaded
launchctl print gui/$(id -u)/com.spigen.gcx.badreview-broadcast

# Test the logic without sending anything
python3 auto_broadcast.py --dry-run                    # as if run today
python3 auto_broadcast.py --dry-run --date 2026-09-25   # simulate a holiday (should SKIP)

# Force a real send right now, bypassing the weekday/holiday check (careful — this is live)
python3 auto_broadcast.py --force

# Unload / reload after editing the plist
launchctl bootout gui/$(id -u)/com.spigen.gcx.badreview-broadcast
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.spigen.gcx.badreview-broadcast.plist

# Disable temporarily without deleting anything
launchctl bootout gui/$(id -u)/com.spigen.gcx.badreview-broadcast
```

There is no delete/uninstall step needed to try this out — `bootout` stops it, and it
won't restart until the plist is `bootstrap`ped again (or the Mac reboots and something
re-loads LaunchAgents automatically, which does NOT happen for a bootout'd agent — it
stays off until manually bootstrapped again).
