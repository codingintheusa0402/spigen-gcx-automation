#!/bin/bash
# Build, sign with the stable local identity (so macOS keeps file-access permissions across
# rebuilds), and install to /Applications.
set -e
cd "$(dirname "$0")"
ID="Claude Mesh Local Signing"
APP="/Applications/Claude Mesh.app"
CSC_IDENTITY_AUTO_DISCOVERY=false npx electron-builder --mac --dir
osascript -e 'quit app "Claude Mesh"' 2>/dev/null || true
for i in $(seq 1 20); do pgrep -f 'Claude Mesh.app/Contents/MacOS/Claude Mesh$' >/dev/null || break; sleep 0.5; done
rm -rf "$APP"
cp -R "dist/mac-arm64/Claude Mesh.app" "$APP"
xattr -cr "$APP"
if security find-identity -p codesigning | grep -q "$ID"; then
  codesign --force --deep -s "$ID" "$APP"
else
  echo "warning: '$ID' not in keychain — ad-hoc signing (permissions won't persist)"; codesign --force --deep -s - "$APP"
fi
open "$APP"
