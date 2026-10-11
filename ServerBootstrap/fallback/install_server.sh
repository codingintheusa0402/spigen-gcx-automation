#!/bin/bash
# Server half of the Mac fallback (run on gcx-server; the Mac runs it automatically over ssh when it's missing).
# Idempotent.  1) mock-trader waits at start while the Mac owns it   2) gcx-owner-watch service stops/starts the
# server's AI Trader + ticket monitor according to the owner   3) tmux_start.sh's ticket-monitor goes through the gate
set -e
F=$HOME/Desktop/GCX/ServerBootstrap/fallback
mkdir -p ~/.config/systemd/user/mock-trader.service.d
cat > ~/.config/systemd/user/mock-trader.service.d/gcx-owner-gate.conf <<EOT
[Service]
ExecStartPre=/usr/bin/python3 $F/gcx_owner.py gate aitrader
TimeoutStartSec=infinity
EOT
cat > ~/.config/systemd/user/gcx-owner-watch.service <<EOT
[Unit]
Description=GCX: hand AI Trader / ticket monitor to the Mac fallback and back
[Service]
ExecStart=/usr/bin/python3 -u $F/gcx_owner.py watch
Restart=always
RestartSec=20
[Install]
WantedBy=default.target
EOT
systemctl --user daemon-reload
systemctl --user enable --now gcx-owner-watch >/dev/null 2>&1
# ticket monitor launched by tmux_start.sh after a reboot: wait for the gate first
for f in ~/tmux_start.sh; do
  grep -q "gcx_owner.py gate ticket-monitor" "$f" || \
    sed -i 's|claude --remote-control gcx-ticket-monitor|python3 '"$F"'/gcx_owner.py gate ticket-monitor; claude --remote-control gcx-ticket-monitor|' "$f"
done
python3 -c "import sys; sys.path.insert(0,'$F'); import gcx_ledger as L; L.put('gates:installed','done',{'v':1})"
echo "gates installed: $(systemctl --user is-active gcx-owner-watch) · tmux_start gated: $(grep -c 'gate ticket-monitor' ~/tmux_start.sh)"
