# ServerBootstrap

Setup kit and job catalog for the **24/7 GCX server**: a Windows 11 laptop (Tailscale name `claude-server`) running WSL2 Ubuntu as its own Tailscale node, `gcx-server`. Unattended GCX jobs run there under cron instead of under the Mac's launchd. Claude Code sessions run there in a shared tmux session, `gcx`, that shows both on the laptop screen and from the Mac (`ssh -t kevinkim@gcx-server tmux attach -t gcx`).

The same folder holds the two catalogs that **GCX Mesh**'s ⏱ Schedules menu reads ([`../ClaudeMesh`](../ClaudeMesh/README.md)): `jobs.json` for Mac/server jobs and `gas_due_dates.json` for Apps Script projects.

## Screenshots

![`crontab -l` on gcx-server (2026-10-07)](docs/crontab.jpg)
*`crontab -l` on gcx-server (2026-10-07)*

## Files

| File | Run where / when | What it does |
|---|---|---|
| `0_enable_ssh.ps1` | Admin PowerShell on the laptop, after Tailscale login | Installs and starts OpenSSH Server, authorizes the Mac's public key, and sets PowerShell as the SSH shell, so Claude on the Mac can do the rest remotely. On the corp laptop `Add-WindowsCapability` hung, so OpenSSH was installed from the winget MSI instead. |
| `1_windows_setup.ps1` | Admin PowerShell, once | Turns off sleep/hibernate on AC and the lid action, installs Tailscale (winget), writes `.wslconfig` (`vmIdleTimeout=-1`, 8 GB), registers the hidden **WSL-KeepAlive** logon task, and runs `wsl --install -d Ubuntu`. |
| `2_wsl_setup.sh` | Inside Ubuntu as the normal user, once | Sets TZ to Asia/Seoul. Installs apt basics (python3, git, tmux, cron, rsync, jq…), Node 22, Claude Code, Google Chrome + Noto CJK (headed via WSLg), and the Python libs the jobs import (`google-api-python-client`, `google-auth*`, `playwright`, `requests`) plus Playwright Chromium. Adds Mac-path shims: `/Users/kevinkim → /home/kevinkim` and `/opt/homebrew/bin/python3 → /usr/bin/python3`. Installs Tailscale with SSH (`sudo tailscale up --ssh --hostname gcx-server`, done by hand). |
| `3_push_from_mac.sh [host]` | On the Mac, after the server is up; safe to re-run | rsyncs the GCX repo, `~/.claude/skills`, the memory folder, `~/CLAUDE.md`, Claude settings/statusline and the per-job `~/.config/*` token/state folders to the server (default host `kevinkim@gcx-server`). It also installs `crontab.txt`, creates the job log folders and starts tmux. `SYNC_ONLY=1` copies files only. Progress is mirrored to the laptop's on-screen log `C:\GCX-Setup\setup.log`. |
| `crontab.txt` | Installed on gcx-server | The server's jobs, converted 1:1 from the Mac launchd agents (see below). |
| `tmux_start.sh` | Copied to `~/tmux_start.sh` on the server | Creates the shared tmux session `gcx` (windows: `main`, `jobs-log` tailing every job log, `setup-log`) and adds an `@reboot` cron line so it comes back after a restart. |
| `jobs.json` | Read by GCX Mesh | Catalog of scheduled jobs: id, what it does, skills, dir, command, log, `state` files that must move with the job, default schedule, and whether it is `movable` between Mac and server. Also lists Claude session loops (the ticket monitor in `gcx:ticket-monitor`) and the cloud GAS triggers. |
| `gas_due_dates.json` | Read and edited by GCX Mesh | 18 Apps Script projects: script ID, what it does, trigger schedule, setup function, and editable date constants (file:line + regex), e.g. MasterTrigger `MASTER_END_DATE`, the Apify `endDate`s, Bi-Weekly `BW_RUN`. Saving a date in GCX Mesh rewrites the literal, runs `clasp push`, and commits. |
| `test/sc_test_report.py` | Server, manual | Waits for an SC scraper run to finish, parses `/tmp/sc_scraper.log`, and posts a `[SERVER TEST]` result card to the **test room only**. Takes the webhook from `$TEST_WEBHOOK`, never hard-coded. |

## Jobs on the server (`crontab.txt`, TZ Asia/Seoul)

| When | Job |
|---|---|
| Mon–Fri 10:30 | `GAS_Operations/BadReview_ChatReport/auto_broadcast.py`: bad-review carousel to the GCX rooms |
| Mon–Fri 11:00 | same, `--retry-if-held` (if the 10:30 run was held by the KR-tagging gate) |
| every 5 min | same, `--catchup` (sends if today's run was missed) |
| every 30 min | `GAS_Operations/CaspiSalesBackfill/backfill.py` (self-gates to once per day per sheet) |
| every 30 min | `GAS_Operations/DiscolorationReport/report.py` (sends once per Mon/Fri due day) |

Logs go to each job's `logs/launchd.{out,err}.log`. The **SKU/ASIN filler** (`Scrapers/SKU_ASIN_Filler`, every 5 min) stays on the Mac because it needs the Seller Central inventory login with OTP in `~/.chrome-sc-inventory-profile`. Its cron line is present but commented out. The live **SC Review Scraper** is also Mac-only for now (EU OTP).

Moving a job between machines is done from GCX Mesh → ⏱ Schedules. It switches the source off, copies the job's `state` files, then switches the target on. Server cron lines it manages are tagged `# mesh:<id>`.

## Setup order

1. On the laptop, log in to Tailscale and run `0_enable_ssh.ps1` (admin).
2. Run `1_windows_setup.ps1` (admin) → reboot → open Ubuntu and create user `kevinkim`.
3. In Ubuntu: `bash 2_wsl_setup.sh`, then `sudo tailscale up --ssh --hostname gcx-server`, then run `claude` once and log in.
4. On the Mac: `bash ServerBootstrap/3_push_from_mac.sh` (later syncs: `SYNC_ONLY=1 bash ServerBootstrap/3_push_from_mac.sh`).
5. Attach: `ssh -t kevinkim@gcx-server tmux attach -t gcx`.

To launch any skill or task as a server Claude session from the Mac, use the `run-on-server` skill ("서버에서 돌려줘").

## Gotchas

- Start WSL from the desktop session, not over SSH. Otherwise WSLg has no display and Chrome fails with "Missing X server". Fix: `wsl --shutdown`, then start the WSL-KeepAlive / GCX-Live scheduled tasks again.
- Windows auto-login is not set up. After a reboot, WSL, cron and tmux stay down until someone logs in at the laptop.
- Chrome cookies can't be copied from the Mac (they're encrypted with the Keychain). Sign the server's Chrome profiles in on the server itself.
- Mac `~/Desktop` reads are slow (iCloud). For large updates, keep the server's GCX as a git clone with the same remotes rather than re-rsyncing everything.

## Recent changes (2026-10-07)

- Kit added. Crontab installed on gcx-server for the badreview ×3, Caspi backfill and discoloration jobs; the matching Mac launchd agents were booted out. The SKU/ASIN filler stays on the Mac.
- `jobs.json` and `gas_due_dates.json` were added for the GCX Mesh Schedules menu, which shows Mac + server jobs and lets you edit the Apps Script due dates.
