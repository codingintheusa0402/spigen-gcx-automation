#!/bin/bash
# Copy GCX Mesh source (no node_modules/dist/backups) into the GCX repo's ClaudeMesh/ folder.
# Then commit + push from ~/Desktop/GCX as usual (stage ClaudeMesh/ only).
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/Desktop/GCX/ClaudeMesh"
mkdir -p "$DEST"
rsync -a --delete --exclude node_modules --exclude dist --exclude dist-install --exclude _game_backup --exclude ios/build --exclude xcuserdata --exclude .DS_Store "$SRC/" "$DEST/"
echo "synced → $DEST"
