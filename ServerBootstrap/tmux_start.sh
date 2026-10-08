#!/usr/bin/env bash
# Shared tmux session 'gcx' — same live view on the server screen AND from the Mac:
#   Mac:    ssh -t kevinkim@gcx-server tmux attach -t gcx
#   Server: "GCX Live" window opens at logon (wsl tmux attach)
export PATH=$HOME/.local/bin:$PATH
( crontab -l 2>/dev/null | grep -v tmux_start; echo "@reboot sleep 30 && bash ~/tmux_start.sh" ) | crontab -
tmux has-session -t gcx 2>/dev/null && exit 0
cd ~/Desktop/GCX
tmux new-session -d -s gcx -n main -x 200 -y 50 "cd ~/Desktop/GCX && export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0; claude --remote-control gcx-main; exec bash"
tmux set -g -t gcx mouse on
tmux set -g -t gcx status-right ' #H | %Y-%m-%d %H:%M '
tmux set -g -t gcx history-limit 50000
tmux new-window -t gcx -n jobs-log 'touch ~/gcx-jobs.log; tail -F ~/gcx-jobs.log ~/Desktop/GCX/GAS_Operations/*/logs/launchd.*.log ~/Desktop/GCX/Scrapers/SKU_ASIN_Filler/launchd.*.log'
tmux new-window -t gcx -n setup-log 'tail -F /mnt/c/GCX-Setup/setup.log'
tmux select-window -t gcx:main
