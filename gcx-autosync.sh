#!/usr/bin/env bash
# gcx-autosync: keep the GCX repo identical on the Mac and the GCX server.
# Runs every 3 min on both (Mac: launchd com.spigen.gcx.git-autosync, server: cron).
#   1. commit whatever changed here (git add -A — .gitignore keeps state/logs/data out)
#   2. pull --rebase the other machine's commits from origin
#   3. push (origin pushes to codingintheusa0402 + spigenHQ)
# Never touches a repo mid-merge/rebase. On a conflict it aborts, leaves the local commit
# unpushed and alerts the private Chat room once.
REPO="$HOME/Desktop/GCX"; LOG="$HOME/.gcx-autosync.log"; LOCK="/tmp/gcx-autosync.lock"
HOST=$( [ "$(uname)" = Darwin ] && echo mac || echo server )
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:$PATH"
log(){ echo "[$(date '+%F %T')] $*" >> "$LOG"; }
alert(){ local u
  u=$( [ -f "$HOME/.config/gcx_autosync_webhook.txt" ] && cat "$HOME/.config/gcx_autosync_webhook.txt" )
  [ -n "$u" ] && curl -s -m 15 -X POST -H 'Content-Type: application/json' "$u" -d "{\"text\":\"⚠️ GCX git auto-sync ($HOST): $1\"}" >/dev/null; }

# single instance (mkdir lock works on macOS + Linux; stale after 15 min)
if ! mkdir "$LOCK" 2>/dev/null; then
  [ -n "$(find "$LOCK" -maxdepth 0 -mmin +15 2>/dev/null)" ] && rm -rf "$LOCK" && mkdir "$LOCK" || exit 0
fi
trap 'rm -rf "$LOCK"' EXIT
cd "$REPO" || exit 0
[ -d .git/rebase-merge ] || [ -d .git/rebase-apply ] || [ -f .git/MERGE_HEAD ] && { log "skip: merge/rebase in progress"; exit 0; }
[ "$(git symbolic-ref --short HEAD 2>/dev/null)" = main ] || { log "skip: not on main"; exit 0; }

git add -A 2>>"$LOG"
if ! git diff --cached --quiet; then
  n=$(git diff --cached --name-only | wc -l | tr -d ' ')
  files=$(git diff --cached --name-only | head -4 | xargs -n1 basename | paste -sd, -)
  git commit -q -m "auto-sync ($HOST): $n file(s) — $files" && log "committed $n file(s): $files"
fi

git fetch -q origin main 2>>"$LOG" || { log "fetch failed"; exit 0; }
if [ -n "$(git rev-list HEAD..origin/main)" ]; then
  if git pull -q --rebase origin main >>"$LOG" 2>&1; then log "pulled $(git rev-list --count ORIG_HEAD..HEAD 2>/dev/null) commit(s)"; rm -f "$HOME/.gcx-autosync.conflict"
  else
    git rebase --abort 2>/dev/null
    log "CONFLICT pulling origin/main — local commits kept, not pushed"
    [ -f "$HOME/.gcx-autosync.conflict" ] || alert "conflict with the other machine's changes — needs a manual merge in ~/Desktop/GCX"
    touch "$HOME/.gcx-autosync.conflict"; exit 0
  fi
fi
if [ -n "$(git rev-list origin/main..HEAD)" ]; then
  git push -q origin main >>"$LOG" 2>&1 && log "pushed" || log "push failed (will retry)"
fi
