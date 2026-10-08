#!/usr/bin/env bash
# usage: send.sh <short-name> "<text>"  -> types text + Enter into that server session
printf '%s' "$2" | ssh kevinkim@gcx-server "cat > /tmp/_send.txt && tmux load-buffer -b s /tmp/_send.txt && tmux paste-buffer -b s -t gcx:$1 && sleep 0.3 && tmux send-keys -t gcx:$1 Enter"
