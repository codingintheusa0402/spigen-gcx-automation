#!/usr/bin/env bash
# GCX server desktop launchers (called from Windows shortcuts). Usage: gcx-launch.sh live|new|resume
export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0
bash ~/tmux_start.sh >/dev/null 2>&1
case "$1" in
  new)    n="claude-$(date +%H%M)"
          tmux new-window -t gcx -n "$n" "cd ~/Desktop/GCX && claude --remote-control gcx-$n; exec bash" ;;
  resume) tmux new-window -t gcx -n "resume-$(date +%H%M)" "cd ~ && claude --resume --remote-control; exec bash" ;;
esac
exec tmux attach -t gcx
