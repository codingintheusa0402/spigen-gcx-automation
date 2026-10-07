#!/usr/bin/env bash
# Run on the MAC after the server is up:  bash 3_push_from_mac.sh [host]
# Copies GCX code, skills, memory, CLAUDE.md, tokens -> server over Tailscale SSH,
# installs crontab, starts tmux sessions. Safe to re-run (also used for syncing).
set -euo pipefail
H="${1:-kevinkim@gcx-server}"
R="rsync -az --stats --exclude=__pycache__ --exclude=*.bak-* --exclude=node_modules --exclude=.DS_Store --exclude=/.git --exclude=.venv"

LOG(){ ssh "$H" "printf '%s\\r\\n' \"[\$(date +%H:%M:%S)] [mac push] $*\" >> /mnt/c/GCX-Setup/setup.log" 2>/dev/null || true; }
ssh "$H" 'mkdir -p ~/Desktop ~/.claude/projects ~/.config'
LOG "copying GCX code, skills, memory, CLAUDE.md, settings, tokens from Mac..."
$R ~/Desktop/GCX/                         "$H":Desktop/GCX/
$R ~/.claude/skills/                      "$H":.claude/skills/
$R ~/.claude/projects/-Users-kevinkim/memory/ "$H":.claude/projects/-home-kevinkim/memory/
$R ~/CLAUDE.md                            "$H":CLAUDE.md
for f in settings.json statusline.sh; do [ -e ~/.claude/$f ] && $R ~/.claude/$f "$H":.claude/$f; done
for c in gws_shim caspi_sales_backfill discoloration_report siren_finder siren_sweeper zendesk_inquiry_sync monday_key.txt; do
  [ -e ~/.config/$c ] && $R ~/.config/$c "$H":.config/
done
# same memory also reachable when Claude is started from /Users/kevinkim (symlinked home)
ssh "$H" 'cd ~/.claude/projects && [ -e -Users-kevinkim ] || ln -s -home-kevinkim -Users-kevinkim'

if [ "${SYNC_ONLY:-0}" != 1 ]; then
  scp ~/Desktop/GCX/ServerBootstrap/crontab.txt "$H":/tmp/gcx.cron
  ssh "$H" 'for d in Desktop/GCX/GAS_Operations/{BadReview_ChatReport,CaspiSalesBackfill,DiscolorationReport}/logs; do mkdir -p ~/$d; done; crontab /tmp/gcx.cron && crontab -l | grep -c python3'
  scp ~/Desktop/GCX/ServerBootstrap/tmux_start.sh "$H":tmux_start.sh
  ssh "$H" 'bash ~/tmux_start.sh'
fi
LOG "push complete"
echo "Pushed to $H. Attach with:  ssh -t $H tmux attach -t gcx"
