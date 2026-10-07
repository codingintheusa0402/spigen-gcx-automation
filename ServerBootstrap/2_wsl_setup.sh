#!/usr/bin/env bash
# Run ONCE inside Ubuntu (WSL) as your normal user:
#   curl/paste this file, then:  bash 2_wsl_setup.sh
set -euo pipefail
U="$(whoami)"

sudo timedatectl set-timezone Asia/Seoul 2>/dev/null || sudo ln -sf /usr/share/zoneinfo/Asia/Seoul /etc/localtime

# systemd so cron/tailscaled start by themselves

export DEBIAN_FRONTEND=noninteractive
sudo apt-get update
sudo apt-get install -y python3 python3-pip python3-venv git tmux cron rsync curl jq unzip build-essential

# Node 22 + Claude Code
curl -fsSL https://deb.nodesource.com/setup_22.x | sudo -E bash -
sudo apt-get install -y nodejs
curl -fsSL https://claude.ai/install.sh | bash

# Google Chrome (headed, runs via WSLg) for Zendesk / Seller Central / Playwright jobs
curl -fsSL -o /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt-get install -y /tmp/chrome.deb fonts-noto-cjk

# Python libs the jobs import
pip3 install --user --break-system-packages google-api-python-client google-auth google-auth-oauthlib playwright requests
python3 -m playwright install-deps chromium || true
python3 -m playwright install chromium

# Mac-path compatibility: every hard-coded /Users/kevinkim/... and /opt/homebrew/bin/python3 keeps working
sudo mkdir -p /Users /opt/homebrew/bin
[ -e /Users/kevinkim ] || sudo ln -s "/home/$U" /Users/kevinkim
sudo ln -sf /usr/bin/python3 /opt/homebrew/bin/python3

# Tailscale with built-in SSH -> Mac can `ssh kevinkim@<this-host>` with no key setup
curl -fsSL https://tailscale.com/install.sh | sh
sudo systemctl enable --now cron tailscaled || true
# login done separately: sudo tailscale up --ssh --hostname gcx-server

echo
echo "DONE. Now:  1) run 'claude' once and log in with the Max account"
echo "           2) tell Claude on the Mac: 'server is up' — it pushes everything over Tailscale."
