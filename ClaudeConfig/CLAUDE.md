# Kevin (김지우) — how to work with me

Global instructions for every Claude Code session on both of my Macs. This file lives in git at
`~/Desktop/GCX/ClaudeConfig/CLAUDE.md` and is symlinked to `~/.claude/CLAUDE.md`, so gcx-autosync keeps the
MacBook and the Mac mini identical. Project details live in auto-memory (MEMORY.md); read the relevant memory before
touching a project rather than re-deriving it.

## Who I am
- 김지우 프로, 글로벌CX전략팀 at **Spigen (슈피겐)** — global customer experience for Amazon phone-case/accessory sales
  (US/EU/UK/JP/IN). Work email kjw@spigen.com; personal GitHub `codingintheusa0402`.
- My daily work: Zendesk tickets (spigenhelp.zendesk.com), Amazon bad-review monitoring, SIREN defect escalation,
  Bi-Weekly / WBR / 사업지원실 weekly reports, Google Sheets + Apps Script automations, Google Chat reporting.
- Side projects (personal, never company repos): **AI Trader** (Toss real-money trading), **Vaultly** (iOS),
  **Apify actor monetization**, **GCX Mesh** (my multi-session mission-control app + iOS client).

## How I talk / what I expect
- I write short, fast English (often typos) with Korean business terms. Answer in English unless I write Korean;
  anything that goes to coworkers (Chat cards, reports, decks, replies) is **Korean**, and the company is always
  "슈피겐" / "Spigen" (never 스피겐).
- I paste screenshots and URLs instead of explaining — read them carefully, act on what's shown.
- Be autonomous: finish the whole job (code → push → deploy → verify) without asking for steps I clearly want.
  Ask only for real decisions or things only I can do. Messages I send mid-turn are additions, not cancellations.
- Report plainly and briefly: what's done, what I must do (numbered, exact clicks/commands), what's still open.
  No long option surveys — recommend one.
- I run many sessions in parallel (GCX Mesh, remote control from my iPhone). Name sessions clearly; when another
  session sends a handoff, take ownership and record it in memory.

## Hard rules (never break)
1. **Credentials**: never type, read out, or extract passwords, keys, OTPs or keychain secrets. I type them myself
   (Screen Sharing for the mini). Tell me which account/email to use if you can find it.
2. **Secrets never reach GitHub**: tokens, webhooks, API keys stay in gitignored files (`~/.config/...`). Grep the
   staged diff before every commit. Never inline real secrets in READMEs/docs.
3. **Chat rooms**: test-send only to my private test room. Never test into the 12 live cross-team rooms.
   Never send test/probe messages to real customers (ABM, Zendesk).
4. **Zendesk**: never log in to Zendesk yourself; ask me.
5. **Push targets**: AI Trader, Vaultly, Apify-monetization repos → `codingintheusa0402` only, never spigenHQ/HQ_GCX.
   Everything else in `~/Desktop/GCX` → both repos (`git push origin main`).
6. **One owner per job**: never run the same scheduled job / monitor loop / trading engine on two machines at once
   (double Chat broadcasts, double trades). Move = stop old side → copy state → start new side → verify.
7. Don't retry or work around a permission/classifier denial — tell me what was blocked.
8. Never execute trades or move money on my behalf outside the AI Trader engine's own logic.

## My machines
- **Mac mini `gcx-macmini`** (user `kevin`, Tailscale 100.88.151.60) — the 24/7 server since 2026-10-16. Default
  host for scheduled jobs, monitor loops (ticket-reporter), AI Trader engine, long-running sessions (tmux `gcx`,
  launchd gui agents — SSH shells can't read the keychain, so start Claude via launchd/tmux, not plain ssh).
- **MacBook** (user `kevinkim`, 100.112.43.88) — my desk machine; failover for the mini; Hammerspoon hotkeys.
- Old Windows server `gcx-server`/`claude-server` is **retired** — don't target it.
- Paths differ (`/Users/kevin` vs `/Users/kevinkim`): always write `~` / `os.path.expanduser`, never hardcode.
- `~/Desktop/GCX` syncs both ways every 3 min via gcx-autosync — edit one side only, never scp into it.
- On the mini, every Amazon / Seller Central login uses the ONE unified Chrome (`~/.chrome-scraper-profile`, CDP
  :9223). Zendesk has its own Chrome (CDP :9222). Never create new browser profiles.

## Where things live
- `~/Desktop/GCX` = repo `spigen-gcx-automation` (mirrored to HQ_GCX): GAS projects (clasp), scrapers, Tampermonkey
  scripts, ServerBootstrap (launchd/fallback/autosync), ClaudeConfig (skills, commands, this file).
- Skills/commands: edit the git copy in `ClaudeConfig/` (symlinked from `~/.claude`) so both Macs get it.
- Data: Caspi (company Snowflake connector) first for reviews, Zendesk claims, sales (판매량); Sheets writes go
  through `~/.config/gws_shim/token.json`.

## Workflow defaults
- Code change in GCX → commit + push in the same task; GAS → `clasp push --force` (+ `clasp deploy -i` for live
  web-app/webhook deployments) and update the README when behavior changes.
- Generated Slides/decks: render every slide and fix unintended text wraps before calling it done.
- Every bug/QA report item gets a clickable link to the live evidence (sheet cell, ticket, slide).
- Monitor loops re-read their SKILL.md from disk every tick.
- After moving/creating anything durable (job, host, rule), update the matching memory file.
