#!/bin/bash
# Double-clickable / `open`-able wrapper: runs the device install inside the user's GUI session
# (unlocked keychain, so automatic signing can create certs/profiles). Log: ~/Vaultly-install.log
LOG="$HOME/Vaultly-install.log"
cd "$(dirname "$0")/.."
{
  echo "=== $(date) start"
  ./scripts/install_device.sh "$@"
  echo "=== $(date) exit=$?"
} > "$LOG" 2>&1
