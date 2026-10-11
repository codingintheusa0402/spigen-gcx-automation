#!/bin/bash
# Build + install GCX Mesh on the iPhone (free Apple ID signing expires every 7 days; keepalive.sh reruns this automatically).
# Needs: Apple ID signed in at Xcode → Settings → Accounts; iPhone connected (cable or same Wi-Fi), unlocked.
set -e
cd "$(dirname "$0")"
DEV=$(xcrun devicectl list devices 2>/dev/null | awk '/physical/ && /(connected|available)/ {for(i=1;i<=NF;i++) if ($i ~ /^[0-9A-F]{8}-[0-9A-F]{16}$/) print $i}' | head -1)
[ -z "$DEV" ] && { echo "Connect the iPhone (cable or same Wi-Fi) and unlock it."; exit 1; }
xcodegen generate >/dev/null
xcodebuild -project ClaudeMesh.xcodeproj -scheme ClaudeMesh -configuration Release -derivedDataPath build \
  -destination "id=$DEV" DEVELOPMENT_TEAM=6Z44CDGL2Q -allowProvisioningUpdates build | grep -E "error|BUILD (SUCCEEDED|FAILED)"
xcrun devicectl device install app --device "$DEV" "build/Build/Products/Release-iphoneos/GCX Mesh.app" | tail -1
echo "✅ Installed GCX Mesh on iPhone $DEV"
