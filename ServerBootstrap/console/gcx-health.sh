#!/usr/bin/env bash
# Linux-side health for the GCX console. One line per card: KEY|level(ok/warn/bad)|text
export PATH=$HOME/.local/bin:/usr/local/bin:/usr/bin:/bin
up=$(awk '{printf "%dh %dm", $1/3600, ($1%3600)/60}' /proc/uptime)
echo "wsl|ok|Ubuntu up $up"
# Claude login + sessions
if [ -s ~/.claude/.credentials.json ]; then cl=ok; ct="logged in"; else cl=bad; ct="NOT logged in"; fi
n=$(pgrep -fc "[c]laude --remote-control" 2>/dev/null || echo 0)
echo "claude|$cl|$ct · $n live session(s)"
if tmux list-windows -t gcx -F '#W' 2>/dev/null | grep -q '^ticket-monitor$' && pgrep -f "[g]cx-ticket-monitor" >/dev/null; then
  echo "monitor|ok|ticket monitor running"; else echo "monitor|bad|ticket monitor NOT running"; fi
# cron + last job run
if systemctl is-active -q cron; then
  last=$(sudo -n journalctl -u cron --since "-2h" --no-pager 2>/dev/null | grep -E "CMD \(cd \\\$G" | tail -1 | awk '{print $3}')
  echo "cron|ok|$(crontab -l 2>/dev/null | grep -cE '^[0-9*]') jobs · last ${last:-—}"
else echo "cron|bad|cron service stopped"; fi
# git autosync
l=$(tail -1 ~/.gcx-autosync.log 2>/dev/null); t=$(echo "$l" | grep -oE '[0-9]{2}:[0-9]{2}:[0-9]{2}')
age=$(( ( $(date +%s) - $(date -r ~/.gcx-autosync.heartbeat +%s 2>/dev/null || echo 0) ) / 60 ))
if [ -f ~/.gcx-autosync.conflict ]; then echo "git|bad|CONFLICT — needs manual merge"
elif [ "$age" -gt 15 ]; then echo "git|warn|sync job not run for ${age} min"
else echo "git|ok|in sync · checked ${age} min ago · last change $t"; fi
# google sheets token
if [ -s ~/.config/gws_shim/token.json ]; then echo "sheets|ok|Sheets token present"; else echo "sheets|bad|Sheets token missing"; fi
