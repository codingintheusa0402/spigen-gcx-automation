#!/usr/bin/env bash
# GCX Mesh — keep the MacBook and the Mac mini one mesh. Runs on BOTH Macs every 3 min (launchd
# com.spigen.gcx.mesh-sync), so syncing goes on while either one is off. Each run works out which Mac it is
# on and syncs with the other one:
#   1. peer links: each machine's GCX Mesh gets the other's address + phone key (peer.json), so the desktop
#      window and the iPhone show the sessions of both machines whichever one you look at
#   2. session histories (~/.claude/projects/**/*.jsonl) MERGED both ways line by line (mesh_jsonl_sync.py),
#      so a session from a machine that went offline can be continued on the other ("Resume here"), and a
#      session continued on both machines keeps both parts. Side files (tool-results/ …) copied, newer wins.
#   --histories-only: skip step 1 (used by run-on-server/sync_transcripts.sh)
# Project folders are renamed on the way: MacBook -Users-kevinkim… <-> mini -Users-kevin…; the old Windows
# laptop's -home-kevinkim… folders keep their name on both. memory/ folders are left to the memory sync.
# Never deletes anything.
set -uo pipefail
if [ "$(id -un)" = kevin ]; then                  # on the Mac mini
  H=macbook;     SELF=MacMini; PEER=MacBook; SELF_KIND=server; PEER_KIND=mac
  SELF_IP=100.88.151.60; PEER_IP=100.112.43.88; LP=-Users-kevin; RP=-Users-kevinkim
else                                              # on the MacBook
  H=gcx-macmini; SELF=MacBook; PEER=MacMini; SELF_KIND=mac; PEER_KIND=server
  SELF_IP=100.112.43.88; PEER_IP=100.88.151.60; LP=-Users-kevinkim; RP=-Users-kevin
fi
OLD=-home-kevinkim                                # the old Windows laptop's folders: same name on both sides
UD="Library/Application Support/Claude Mesh"
LOG="$HOME/.gcx-mesh-sync.log"
log(){ echo "[$(date '+%F %T')] $*" >> "$LOG"; }
SSH="ssh -o ConnectTimeout=8 -o BatchMode=yes"
$SSH $H true 2>/dev/null || { log "$PEER unreachable"; exit 0; }

# one run at a time ACROSS both Macs: two overlapping merges would both append the same missing lines.
# The lock lives on the mini (both Macs take the same one); stale after 30 min.
LOCK=.gcx-mesh-sync.lock
if [ "$SELF" = MacMini ]; then lock(){ mkdir "$HOME/$LOCK" 2>/dev/null; }; stale(){ find "$HOME/$LOCK" -maxdepth 0 -mmin +30 | grep -q .; }; unlock(){ rmdir "$HOME/$LOCK" 2>/dev/null; }
else lock(){ $SSH $H "mkdir $LOCK" 2>/dev/null; }; stale(){ $SSH $H "find $LOCK -maxdepth 0 -mmin +30" 2>/dev/null | grep -q .; }; unlock(){ $SSH $H "rmdir $LOCK" 2>/dev/null; }
fi
if ! lock; then
  if stale; then unlock; lock || exit 0; log "stale lock removed"
  else log "a run is still going (here or on $PEER), skipped"; exit 0; fi
fi
trap unlock EXIT

# ---- 1. peer links (only rewritten when something changed) ----
if [ "${1:-}" != "--histories-only" ]; then
tok(){ /usr/bin/python3 -c "import json,os;print(json.load(open(os.path.expanduser('~/$UD/phone.json')))['token'])"; }
MY_TOK=$(tok 2>/dev/null)
PEER_TOK=$($SSH $H "/usr/bin/python3 -c \"import json,os;print(json.load(open(os.path.expanduser('~/$UD/phone.json')))['token'])\"" 2>/dev/null)
if [ -n "$MY_TOK" ] && [ -n "$PEER_TOK" ]; then
  want_there="{\"url\":\"http://$SELF_IP:47320\",\"token\":\"$MY_TOK\",\"kind\":\"$SELF_KIND\",\"label\":\"$SELF\",\"selfLabel\":\"$PEER\"}"
  want_here="{\"url\":\"http://$PEER_IP:47320\",\"token\":\"$PEER_TOK\",\"kind\":\"$PEER_KIND\",\"label\":\"$PEER\",\"selfLabel\":\"$SELF\"}"
  [ "$($SSH $H "cat '$UD/peer.json'" 2>/dev/null)" = "$want_there" ] || { printf '%s' "$want_there" | $SSH $H "umask 077; cat > '$UD/peer.json'"; log "$PEER peer.json updated"; }
  [ "$(cat "$HOME/$UD/peer.json" 2>/dev/null)" = "$want_here" ] || { (umask 077; printf '%s' "$want_here" > "$HOME/$UD/peer.json"); log "$SELF peer.json updated"; }
else
  log "peer link skipped (token here=${MY_TOK:+yes}, $PEER token=${PEER_TOK:+yes})"
fi
fi

# ---- 2. session histories, both ways ----
# the merge script runs on both ends; ship the current copy so the other side never runs a stale one
$SSH $H "mkdir -p .gcx-mesh" && rsync -c "$(dirname "$0")/mesh_jsonl_sync.py" "$H:.gcx-mesh/mesh_jsonl_sync.py" 2>>"$LOG" || log "copy of merge script to $PEER failed"
out=$(MESH_SYNC_HOST=$H python3 "$(dirname "$0")/mesh_jsonl_sync.py" sync 2>&1) && log "histories: $out" || log "history merge failed: $out"
cd "$HOME/.claude/projects" || exit 0
R="rsync -au --exclude=memory --exclude=.DS_Store --exclude=*.jsonl"     # side files only; .jsonl never whole-file
for d in $LP $LP-* $OLD $OLD-*; do
  [ -d "$d" ] && [ ! -L "$d" ] || continue
  case "$d" in $OLD*) t="$d";; *) t="$RP${d#$LP}";; esac
  $R -- "$d/" "$H:.claude/projects/$t/" 2>>"$LOG" || log "side push failed: $d"
done
# (find, not a shell glob: the remote login shell is zsh, which aborts the whole command on a pattern with no match)
for t in $($SSH $H "cd ~/.claude/projects && find . -maxdepth 1 -type d \\( -name '$RP' -o -name '$RP-*' -o -name '$OLD' -o -name '$OLD-*' \\) | sed 's|^\\./||'"); do
  case "$t" in $OLD*) m="$t";; *) m="$LP${t#$RP}";; esac
  mkdir -p -- "$m"
  $R -- "$H:.claude/projects/$t/" "$m/" 2>>"$LOG" || log "side pull failed: $t"
done
log "ok"
