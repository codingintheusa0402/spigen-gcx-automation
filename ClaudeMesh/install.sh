#!/bin/bash
# Build, sign with the stable local identity (so macOS keeps file-access permissions across
# rebuilds), and install to /Applications.
# Safe to run from a Claude session that lives inside Claude Mesh: it builds into dist-install/
# (never the dist/ copy you may be running) and only quits/relaunches the app when no sessions
# are running inside it — otherwise it installs and asks you to restart Claude Mesh yourself.
set -e
cd "$(dirname "$0")"
ID="Claude Mesh Local Signing"
APP="/Applications/Claude Mesh.app"
CSC_IDENTITY_AUTO_DISCOVERY=false npx electron-builder --mac --dir -c.directories.output=dist-install

# Claude Mesh instances that have sessions running inside them (their children are `claude`)
busy=0
for pid in $(pgrep -f 'Claude Mesh.app/Contents/MacOS/Claude Mesh$'); do
  if ps -axo ppid=,comm= | awk -v p="$pid" '$1==p && $2 ~ /claude/' | grep -q .; then busy=1; fi
done
running=$(pgrep -f 'Applications/Claude Mesh.app/Contents/MacOS/Claude Mesh$' || true)
if [ "$busy" = 0 ] && [ -n "$running" ]; then
  osascript -e 'quit app "Claude Mesh"' 2>/dev/null || true
  for i in $(seq 1 20); do pgrep -f 'Applications/Claude Mesh.app/Contents/MacOS/Claude Mesh$' >/dev/null || break; sleep 0.5; done
fi

rm -rf "$APP"
cp -R "dist-install/mac-arm64/Claude Mesh.app" "$APP"
xattr -cr "$APP"
if security find-identity -p codesigning | grep -q "$ID"; then
  codesign --force --deep -s "$ID" "$APP"
else
  echo "warning: '$ID' not in keychain — ad-hoc signing (permissions won't persist)"; codesign --force --deep -s - "$APP"
fi

if [ "$busy" = 1 ]; then
  echo "Installed. Sessions are running inside Claude Mesh, so it was left open — restart it when convenient to load the update."
else
  open "$APP"
fi
