#!/bin/bash
# Install / update the GCX Mesh server program on gcx-server (WSL Ubuntu, run as kevinkim):
#   bash ~/Desktop/GCX/ClaudeMesh/server/install.sh
# Dependencies (xterm for the phone page, qrcode for --pair) live in ~/.gcx-mesh/deps and are linked in as
# node_modules (gitignored) so the GCX checkout stays clean for git-autosync.
set -e
ROOT=$(cd "$(dirname "$0")/.." && pwd)
D=$HOME/.gcx-mesh
mkdir -p "$D/deps" "$HOME/.config/systemd/user"
cd "$D/deps"
[ -f package.json ] || echo '{"name":"gcx-mesh-deps","private":true}' > package.json
npm install --omit=dev --ignore-scripts --no-audit --no-fund --loglevel=error '@xterm/xterm@^6.0.0' '@xterm/addon-fit@^0.11.0' 'qrcode@^1.5.4'
[ -e "$ROOT/node_modules" ] && [ ! -L "$ROOT/node_modules" ] && rm -rf "$ROOT/node_modules"
ln -sfn "$D/deps/node_modules" "$ROOT/node_modules"
sed "s|__ROOT__|$ROOT|" "$ROOT/server/gcx-mesh.service" > "$HOME/.config/systemd/user/gcx-mesh.service"
systemctl --user daemon-reload
systemctl --user enable gcx-mesh >/dev/null 2>&1 || true
systemctl --user restart gcx-mesh
sleep 2
systemctl --user is-active gcx-mesh
node "$ROOT/server/mesh-server.js" --pair
