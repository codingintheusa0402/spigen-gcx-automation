#!/bin/bash
# Keeps the free-Apple-ID build of GCX Mesh from expiring (7-day signing):
# renews (fresh profile + re-sign + reinstall) when ≤ 3 days remain. Runs from launchd twice a day.
# Xcode keeps the Apple ID signed in, so no re-login is needed; the iPhone just has to be reachable
# (cable, or same Wi-Fi once paired). `touch ~/Library/Logs/GCXMesh/ios_force_renew` forces a renewal.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"
cd "$(dirname "$0")"
LOG="$HOME/Library/Logs/GCXMesh/ios_keepalive.log"; FORCE="$HOME/Library/Logs/GCXMesh/ios_force_renew"
APP="build/Build/Products/Release-iphoneos/GCX Mesh.app"
note() { osascript -e "display notification \"$1\" with title \"GCX Mesh (iPhone)\""; }
exp=$(security cms -D -i "$APP/embedded.mobileprovision" 2>/dev/null | plutil -extract ExpirationDate raw -o - - 2>/dev/null)
left=$(( ( $(date -j -u -f "%Y-%m-%dT%H:%M:%SZ" "$exp" +%s 2>/dev/null || echo 0) - $(date +%s) ) / 3600 ))
echo "$(date '+%F %T') signature expires $exp (${left}h left)" >> "$LOG"
[ "$left" -gt 72 ] && [ ! -f "$FORCE" ] && exit 0
rm -f "$FORCE"
# Xcode reuses a still-valid profile, so drop the cached GCX Mesh profile to force a fresh 7-day one
for f in "$HOME/Library/Developer/Xcode/UserData/Provisioning Profiles/"*.mobileprovision; do
  [ -f "$f" ] || continue
  security cms -D -i "$f" 2>/dev/null | plutil -extract Entitlements.application-identifier raw -o - - 2>/dev/null \
    | grep -q "com.kevinkim.claudemesh.ios" && rm -f "$f" && echo "$(date '+%F %T') dropped cached profile $(basename "$f")" >> "$LOG"
done
if ./install.sh >> "$LOG" 2>&1; then
  new=$(security cms -D -i "$APP/embedded.mobileprovision" | plutil -extract ExpirationDate raw -o - -)
  echo "$(date '+%F %T') renewed → expires $new" >> "$LOG"; note "Renewed on your iPhone — good until ${new%%T*}"
else
  echo "$(date '+%F %T') renew FAILED (iPhone not reachable?)" >> "$LOG"
  note "Can't reach your iPhone to renew (expires in ${left}h). Connect it by cable or same Wi-Fi and unlock it."
fi
