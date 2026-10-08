#!/usr/bin/env bash
# Mac -> server: Claude session transcripts, so Mac sessions can be resumed on the server
# (project dirs -Users-kevinkim* are renamed -home-kevinkim*; memory/ dirs are handled by sync_to_server.sh)
H=kevinkim@gcx-server
cd ~/.claude/projects || exit 1
for d in -Users-kevinkim*; do
  [ -d "$d" ] || continue
  t="-home-kevinkim${d#-Users-kevinkim}"
  rsync -az --exclude=memory --exclude=.DS_Store -- "$d/" "$H:.claude/projects/$t/" 2>/dev/null || echo "failed: $d"
done
echo "transcripts synced ($(ls -d -- -Users-kevinkim* | wc -l | tr -d ' ') project dirs)"
