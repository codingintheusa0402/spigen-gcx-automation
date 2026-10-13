#!/usr/bin/env bash
# GCX server desktop launchers (called from Windows shortcuts). Usage: gcx-launch.sh live|new|resume
export PATH=$HOME/.local/bin:$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0
bash ~/tmux_start.sh >/dev/null 2>&1
case "$1" in
  new)    n="claude-$(date +%H%M)"
          tmux new-window -t gcx -n "$n" "cd ~/Desktop/GCX && claude --remote-control gcx-$n; exec bash" ;;
  resume) tmux new-window -t gcx -n "resume-picker" "python3 $HOME/Desktop/GCX/ServerBootstrap/gcx-resume-picker.py" ;;   # picker: a LIVE session switches to its tab, never a 2nd copy
esac
exec tmux attach -t gcx
