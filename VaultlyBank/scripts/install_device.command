#!/bin/bash
# Double-clickable / `open`-able wrapper: runs the device install inside the user's GUI session
# (unlocked keychain, so automatic signing can create certs/profiles). Log: ~/Vaultly-install.log
LOG="$HOME/Vaultly-install.log"
cd "$(dirname "$0")/.."
{
  echo "=== $(date) start"
  [ $# -eq 0 ] && set -- 00008140-001804AE3C91401C   # default: Kevin's iPhone 16 Pro (open -a Terminal passes no args)
  ./scripts/install_device.sh "$@"
  echo "=== $(date) exit=$?"
} > "$LOG" 2>&1
