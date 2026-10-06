#!/bin/bash
# Build Claude Mesh for iPhone and install it on the connected iPhone.
# Needs: Xcode installed (and opened once), your Apple ID added in Xcode → Settings → Accounts,
# the iPhone connected by USB (or paired over Wi-Fi) with Developer Mode on.
set -e
cd "$(dirname "$0")"
if ! xcodebuild -version >/dev/null 2>&1; then
  echo "✗ Xcode isn't active. Install it from the App Store, open it once, then run:"
  echo "    sudo xcode-select -s /Applications/Xcode.app/Contents/Developer"; exit 1
fi
command -v xcodegen >/dev/null && xcodegen generate >/dev/null

# your team id = OU of the "Apple Development" certificate Xcode creates for your Apple ID
TEAM=$(security find-certificate -a -c "Apple Development" -p 2>/dev/null | openssl x509 -noout -subject 2>/dev/null | sed -n 's/.*OU *= *\([A-Z0-9]\{10\}\).*/\1/p' | head -1)
[ -n "$TEAM" ] || TEAM=$(defaults read com.apple.dt.Xcode IDEProvisioningTeams 2>/dev/null | sed -n 's/.*teamID = \([A-Z0-9]*\);.*/\1/p' | head -1)
if [ -z "$TEAM" ]; then
  echo "✗ No Apple ID team found. Open Xcode → Settings → Accounts → + → Apple ID, sign in, then run this again."; exit 1
fi
echo "• team $TEAM"

DEV=$(xcrun devicectl list devices 2>/dev/null | awk '/iPhone/ && /(available|connected)/ {for(i=1;i<=NF;i++) if ($i ~ /^[0-9A-F]{8}-[0-9A-F-]+$/) {print $i; exit}}')
if [ -z "$DEV" ]; then
  echo "✗ No iPhone found. Plug it in with a cable, unlock it, tap 'Trust', and turn on"
  echo "  Settings → Privacy & Security → Developer Mode (the phone restarts). Then run this again."; exit 1
fi
echo "• iPhone $DEV"

xcodebuild -project ClaudeMesh.xcodeproj -scheme ClaudeMesh -configuration Release \
  -destination "id=$DEV" -derivedDataPath build DEVELOPMENT_TEAM="$TEAM" -allowProvisioningUpdates -quiet build
APP=$(ls -d build/Build/Products/Release-iphoneos/*.app | head -1)
xcrun devicectl device install app --device "$DEV" "$APP"
echo "✓ Installed. First launch: Settings → General → VPN & Device Management → trust your Apple ID, then open Claude Mesh."
