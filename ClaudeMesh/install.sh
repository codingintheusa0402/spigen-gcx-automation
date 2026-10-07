#!/bin/bash
# Build GCX Mesh, sign it with the stable local identity (so macOS keeps file-access permissions
# across rebuilds) and install it to /Applications.
# - Updates the installed app IN PLACE (rsync into the same bundle) — deleting and re-copying it
#   made the Dock drop its pinned icon on every update.
# - Safe to run from a Claude session living inside GCX Mesh: builds into dist-install/ (never the
#   dist/ copy you may be running) and only quits/relaunches the app when no sessions run inside it.
set -e
cd "$(dirname "$0")"
ID="Claude Mesh Local Signing"            # keychain identity (name kept from before the rename)
APP="/Applications/GCX Mesh.app"
OLD="/Applications/Claude Mesh.app"
CSC_IDENTITY_AUTO_DISCOVERY=false npx electron-builder --mac --dir -c.directories.output=dist-install

# running copies (either name), and whether any has Claude sessions inside it
# (ps, not pgrep: pgrep hides its own ancestors — i.e. the app this script may be running inside)
pids() { ps -axo pid=,command= | awk '/(GCX|Claude) Mesh\.app\/Contents\/MacOS\/(GCX|Claude) Mesh$/ {print $1}'; }
busy=0
for pid in $(pids); do
  if ps -axo ppid=,comm= | awk -v p="$pid" '$1==p && $2 ~ /(^|\/)claude$/' | grep -q .; then busy=1; fi
done
if [ "$busy" = 0 ]; then
  for name in "GCX Mesh" "Claude Mesh"; do osascript -e "quit app \"$name\"" 2>/dev/null || true; done
  for i in $(seq 1 20); do [ -z "$(pids)" ] && break; sleep 0.5; done
fi

mkdir -p "$APP"
rsync -a --delete "dist-install/mac-arm64/GCX Mesh.app/" "$APP/"
xattr -cr "$APP"
if security find-identity -p codesigning | grep -q "$ID"; then
  codesign --force --deep -s "$ID" "$APP"
else
  echo "warning: '$ID' not in keychain — ad-hoc signing (permissions won't persist)"; codesign --force --deep -s - "$APP"
fi
touch "$APP"                               # nudge Finder/Dock to refresh the icon

# the pre-rename bundle: remove once nothing runs from it
if [ -d "$OLD" ] && ! ps -axo command= | grep -q "^$OLD/Contents/MacOS/"; then rm -rf "$OLD"; echo "Removed old $OLD"; fi

if [ "$busy" = 1 ]; then
  echo "Installed. Sessions are running inside the app, so it was left open — restart it (GCX Mesh in /Applications) when convenient."
else
  open "$APP"
fi
