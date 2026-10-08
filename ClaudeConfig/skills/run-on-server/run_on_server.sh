#!/usr/bin/env bash
# usage: run_on_server.sh <short-name> "<prompt>"  -> new Claude session in tmux window gcx:<short-name> on the server
set -euo pipefail
N="$1"; P="$2"; H=kevinkim@gcx-server
printf '%s' "$P" | ssh $H "mkdir -p ~/.gcx-tasks && cat > ~/.gcx-tasks/$N.prompt"
ssh $H "bash ~/tmux_start.sh >/dev/null 2>&1; tmux kill-window -t gcx:$N 2>/dev/null; \
  tmux new-window -t gcx -n $N 'cd ~ && export PATH=\$HOME/.local/bin:\$PATH DISPLAY=:0 WAYLAND_DISPLAY=wayland-0; claude --remote-control gcx-$N \"\$(cat ~/.gcx-tasks/$N.prompt)\"; echo; echo \"[session ended]\"; exec bash'; \
  echo \"started gcx:$N\""
