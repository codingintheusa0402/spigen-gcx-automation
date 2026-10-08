#!/usr/bin/env bash
# usage: peek.sh <short-name> [lines]  -> last screen of that server session
ssh kevinkim@gcx-server "tmux capture-pane -p -J -t gcx:$1 -S -${2:-60} | grep -v '^\s*\$' | tail -${2:-40}"
