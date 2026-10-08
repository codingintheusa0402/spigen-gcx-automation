#!/usr/bin/env bash
# Mac -> server sync of everything Claude needs. Safe: never overwrites server-side git work.
set -uo pipefail
H=kevinkim@gcx-server
export COPYFILE_DISABLE=1
ssh -o ConnectTimeout=10 $H true || { echo "SERVER UNREACHABLE"; exit 2; }

# 1) Claude config, skills, commands, memory
cd ~ && tar --exclude='*.bak-*' --exclude='__pycache__' --exclude='.DS_Store' -czf - \
  .claude/settings.json .claude/statusline.sh CLAUDE.md 2>/dev/null \
  | ssh $H 'tar xzf - -C ~ 2>/dev/null'
cd ~/.claude/projects && tar -czf - -- -Users-kevinkim/memory -Users-kevinkim-Desktop-GCX/memory 2>/dev/null \
  | ssh $H 'cd ~/.claude/projects && mkdir -p _in && tar xzf - -C _in 2>/dev/null; rm -rf -- -home-kevinkim/memory -home-kevinkim-Desktop-GCX/memory && mkdir -p -- -home-kevinkim -home-kevinkim-Desktop-GCX && mv -- _in/-Users-kevinkim/memory -home-kevinkim/ && mv -- _in/-Users-kevinkim-Desktop-GCX/memory -home-kevinkim-Desktop-GCX/ && rm -rf _in'
cd ~ && tar --exclude='.DS_Store' -czf - .claude/plugins 2>/dev/null | ssh $H 'tar xzf - -C ~ 2>/dev/null'
bash ~/.claude/skills/run-on-server/sync_transcripts.sh >/dev/null 2>&1 &   # session history, in the background
echo "claude config/skills/plugins/memory: synced (transcripts syncing in background)"

# 2) GCX repo: same commit, then same uncommitted files
cd ~/Desktop/GCX
MAC_HEAD=$(git rev-parse HEAD)
R=$(ssh $H "cd ~/Desktop/GCX && git fetch -q origin 2>&1; \
  if [ -n \"\$(git log origin/main..HEAD --oneline)\" ]; then echo SERVER_HAS_UNPUSHED; exit; fi; \
  git cat-file -e $MAC_HEAD 2>/dev/null || { echo MAC_HEAD_NOT_PUSHED; exit; }; \
  if [ \"\$(git rev-parse HEAD)\" != $MAC_HEAD ]; then git stash push -q -u -m \"pre-sync \$(date +%F_%T)\" 2>/dev/null; git checkout -q -B main $MAC_HEAD && git branch -q -u origin/main; fi; echo OK")
# (server-only uncommitted work, if any, is kept in `git stash list` on the server — never dropped)
case "$R" in
  *SERVER_HAS_UNPUSHED*) echo "GCX: server has unpushed commits — push them from the server (or pull here) first. Repo NOT synced."; exit 3;;
  *MAC_HEAD_NOT_PUSHED*) echo "GCX: Mac HEAD $MAC_HEAD is not on GitHub — push from the Mac first. Repo NOT synced."; exit 3;;
esac
# server-side edits the Mac doesn't have (e.g. a fix made by a server session) must never be overwritten
SRV_ONLY=$(comm -13 <(git status --porcelain | sort) <(ssh $H 'cd ~/Desktop/GCX && git status --porcelain | sort'))
if [ -n "$SRV_ONLY" ]; then
  echo "GCX: server has changes the Mac doesn't — NOT copying files. Bring them to the Mac first (scp + commit):"; echo "$SRV_ONLY"; exit 4
fi
git status --porcelain -uall | grep -vE '\.venv/|__pycache__|\.DS_Store' | sed -E 's/^(..) //; s/^"//; s/"$//' | grep -v ' -> ' > /tmp/_gcx_delta.txt
DELETED=$(git status --porcelain | awk '$1=="D"{print $2}' | tr '\n' ' ')
[ -s /tmp/_gcx_delta.txt ] && tar -czf - -T /tmp/_gcx_delta.txt 2>/dev/null | ssh $H 'cd ~/Desktop/GCX && tar xzf - 2>/dev/null'
[ -n "$DELETED" ] && ssh $H "cd ~/Desktop/GCX && rm -f $DELETED"
if diff -q <(git status -s | sort) <(ssh $H 'cd ~/Desktop/GCX && git status -s | sort') >/dev/null; then
  echo "GCX: identical to Mac ($(git log -1 --format=%h), $(git status -s | wc -l | tr -d ' ') uncommitted entries)"
else
  echo "GCX: synced but git status differs:"; diff <(git status -s | sort) <(ssh $H 'cd ~/Desktop/GCX && git status -s | sort') | head -20
fi
