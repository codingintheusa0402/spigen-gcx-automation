#!/bin/bash
# Build, install and launch Vaultly on the USB-connected iPhone (automatic signing, personal team).
set -euo pipefail
cd "$(dirname "$0")/.."
DEVICE="${1:-$(xcrun devicectl list devices 2>/dev/null | awk '/physical/ {print $(NF-5)}' | head -1)}"
[ -n "$DEVICE" ] || { echo "no physical iPhone connected"; exit 1; }
DD="${TMPDIR:-/tmp}/vaultly-dd-device"
xcodegen generate >/dev/null
xcodebuild build -project Vaultly.xcodeproj -scheme Vaultly -destination "platform=iOS,id=$DEVICE" \
  -derivedDataPath "$DD" -allowProvisioningUpdates -allowProvisioningDeviceRegistration -quiet
xcrun devicectl device install app --device "$DEVICE" "$DD/Build/Products/Debug-iphoneos/Vaultly.app" >/dev/null
xcrun devicectl device process launch --device "$DEVICE" com.kevinkim.vaultly >/dev/null
echo "installed + launched on $DEVICE"
