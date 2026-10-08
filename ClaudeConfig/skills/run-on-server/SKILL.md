---
name: run-on-server
description: Run any skill, task or Claude session on the 24/7 GCX Windows server (claude-server / WSL2 Ubuntu `gcx-server`) instead of this Mac. Trigger whenever the user says to do something "on the server", "server-side", "서버에서 돌려줘", "서버에서 실행해", "run it on the server PC", "use the server for this", or names a skill/task together with server (e.g. "run sc-scraper on the server", "ticket-reporter 모니터 서버에서 돌려"). Also for checking / attaching to / stopping server sessions ("what's running on the server", "서버 세션 보여줘").
---

# Run on Server

The user has a 24/7 server: Windows laptop `claude-server` (Tailscale) running WSL2 Ubuntu `gcx-server`
(user `kevinkim`, same Claude Max account, same skills/memory/GCX repo as this Mac). Every Claude session on it
lives in the shared tmux session **`gcx`**, which is shown live on the laptop's screen ("GCX Live" window) and
attachable from the Mac: `ssh -t kevinkim@gcx-server tmux attach -t gcx`.

When the user asks to run something server-side, **do not run it on the Mac**. Instead start the SAME task as its own
Claude session on the server, so it is identical to what would run here (same skill, same rules, same memory).

## Steps

1. **Sync first** (Mac → server: skills, commands, CLAUDE.md, settings, memory, GCX uncommitted files; aborts if the
   server repo has its own uncommitted/unpushed work so nothing is overwritten):
   ```bash
   bash ~/.claude/skills/run-on-server/sync_to_server.sh
   ```
2. **Launch** the task as a new server session. Write the full instruction you would otherwise act on here — the
   user's exact request plus every answer already collected in this chat (dates, options, URLs, room choice) — so the
   server session doesn't need to ask again:
   ```bash
   bash ~/.claude/skills/run-on-server/run_on_server.sh <short-name> "<full prompt>"
   ```
   - `<short-name>`: kebab-case, e.g. `sc-scraper-261007`, `ticket-monitor`, `siren-finder`. Becomes the tmux
     window name AND the Remote Control session name (visible at claude.ai/code as `gcx-<short-name>`).
   - The prompt should start with the skill trigger if it's a skill (e.g. "Run the sc-scraper skill with: …").
   - For skills that normally ask questions one by one (sc-scraper, bi-weekly-builder, siren-*), ask those questions
     HERE first, then pass all answers in the prompt.
3. **Watch / relay**: `bash ~/.claude/skills/run-on-server/peek.sh <short-name>` prints the last screen of that
   session. To answer a question the server session asks: `bash ~/.claude/skills/run-on-server/send.sh <short-name> "<text>"`.
   For long tasks arm a Monitor that runs peek.sh every 30–60 s and emits only on new questions / errors / completion.
4. **Report back** to the user: what is running, window name, how to watch it (GCX Live window on the laptop,
   `ssh -t kevinkim@gcx-server tmux attach -t gcx` on the Mac, or claude.ai/code → `gcx-<short-name>`).
5. **After it finishes**, if the server session committed+pushed to git, run `cd ~/Desktop/GCX && git pull --ff-only`
   on the Mac so both machines stay identical.

## Other commands
- List server sessions: `ssh kevinkim@gcx-server 'tmux list-windows -t gcx -F "#I #W"'`
- Stop one: `ssh kevinkim@gcx-server 'tmux kill-window -t gcx:<short-name>'`
- Windows-side admin commands (power, reboot, scheduled tasks): SSH `user@claude-server` with
  `-i ~/.ssh/id_ed25519_gcx_server` (default shell PowerShell).

## Rules that still apply on the server
- All hard rules in memory apply unchanged (test-room-first for Chat broadcasts, never test-send to live rooms,
  no test messages to real ABM customers, etc.) — the server has the same memory.
- Browser tasks: the server's Chrome profile is `~/.chrome-scraper-profile` (signed in as kjw@spigen.com, shared with
  the SC scraper; desktop shortcut "GCX Chrome (kjw@spigen.com)"). Logins that expire need the user at the laptop
  or via Remote Desktop — tell them which site and tab.
- Hammerspoon / AppleScript / `osascript` steps don't exist on the server; replace "open a Terminal tailing X" with
  a tmux window (`tmux new-window -t gcx -n <name> "tail -F X"`).
- If the server is unreachable (`ssh kevinkim@gcx-server true` fails), say so and ask before falling back to the Mac.
