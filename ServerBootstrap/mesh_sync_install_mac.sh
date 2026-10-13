#!/usr/bin/env bash
# Run once ON THE MAC: installs launchd job com.spigen.gcx.mesh-sync (every 3 min) for mesh_sync_mac.sh,
# then runs it once. Undo: launchctl bootout gui/$(id -u)/com.spigen.gcx.mesh-sync && rm the plist.
set -euo pipefail
S="$HOME/Desktop/GCX/ServerBootstrap/mesh_sync_mac.sh"
P="$HOME/Library/LaunchAgents/com.spigen.gcx.mesh-sync.plist"
cat > "$P" <<PL
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>com.spigen.gcx.mesh-sync</string>
  <key>ProgramArguments</key><array><string>/bin/bash</string><string>$S</string></array>
  <key>StartInterval</key><integer>180</integer>
  <key>RunAtLoad</key><true/>
  <key>EnvironmentVariables</key><dict><key>PATH</key><string>/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin</string></dict>
  <key>StandardErrorPath</key><string>$HOME/.gcx-mesh-sync.err.log</string>
</dict></plist>
PL
launchctl bootout gui/$(id -u)/com.spigen.gcx.mesh-sync 2>/dev/null || true
launchctl bootstrap gui/$(id -u) "$P"
echo "installed; log: ~/.gcx-mesh-sync.log"
