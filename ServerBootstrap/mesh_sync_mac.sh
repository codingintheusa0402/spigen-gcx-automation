#!/usr/bin/env bash
# GCX Mesh — keep the Mac and gcx-server one mesh. Runs ON THE MAC every 3 min (launchd
# com.spigen.gcx.mesh-sync, installed by mesh_sync_install_mac.sh). The Mac drives it because the
# server can't open connections to the Mac.
#   1. peer links: each machine's GCX Mesh gets the other's address + phone key (peer.json), so the
#      iPhone sees the sessions of both machines whichever one it is connected to
#   2. session histories (~/.claude/projects/**/*.jsonl) MERGED both ways line by line (mesh_jsonl_sync.py),
#      so a session from a machine that went offline can be continued on the other ("Resume here"), and a
#      session continued on both machines keeps both parts. Side files (tool-results/ …) copied, newer wins.
#   --histories-only: skip step 1 (used by run-on-server/sync_transcripts.sh)
# Project folders are renamed on the way: Mac -Users-kevinkim… <-> server -home-kevinkim…
# memory/ folders are left to the existing memory sync. Never deletes anything.
set -uo pipefail
H=kevinkim@gcx-server
MAC_UD="$HOME/Library/Application Support/Claude Mesh"
LOG="$HOME/.gcx-mesh-sync.log"
log(){ echo "[$(date '+%F %T')] $*" >> "$LOG"; }
ssh -o ConnectTimeout=8 -o BatchMode=yes $H true 2>/dev/null || { log "server unreachable"; exit 0; }

# ---- 1. peer links (only rewritten when something changed) ----
if [ "${1:-}" != "--histories-only" ]; then
MAC_TS=$(/usr/local/bin/tailscale ip -4 2>/dev/null || /Applications/Tailscale.app/Contents/MacOS/Tailscale ip -4 2>/dev/null | head -1)
MAC_TOK=$(python3 -c "import json,sys;print(json.load(open(sys.argv[1]))['token'])" "$MAC_UD/phone.json" 2>/dev/null)
SRV_TOK=$(ssh $H "python3 -c \"import json;print(json.load(open('/home/kevinkim/.gcx-mesh/phone.json'))['token'])\"" 2>/dev/null)
if [ -n "$MAC_TS" ] && [ -n "$MAC_TOK" ] && [ -n "$SRV_TOK" ]; then
  want_srv="{\"url\":\"http://$MAC_TS:47320\",\"token\":\"$MAC_TOK\",\"kind\":\"mac\"}"
  want_mac="{\"url\":\"http://100.121.202.126:47320\",\"token\":\"$SRV_TOK\",\"kind\":\"server\"}"
  [ "$(ssh $H cat .gcx-mesh/peer.json 2>/dev/null)" = "$want_srv" ] || { printf '%s' "$want_srv" | ssh $H 'umask 077; cat > .gcx-mesh/peer.json'; log "server peer.json updated"; }
  [ "$(cat "$MAC_UD/peer.json" 2>/dev/null)" = "$want_mac" ] || { (umask 077; printf '%s' "$want_mac" > "$MAC_UD/peer.json"); log "mac peer.json updated"; }
else
  log "peer link skipped (mac ts=$MAC_TS, mac token=${MAC_TOK:+yes}, server token=${SRV_TOK:+yes})"
fi
fi

# ---- 2. session histories, both ways ----
out=$(python3 "$(dirname "$0")/mesh_jsonl_sync.py" sync 2>&1) && log "histories: $out" || log "history merge failed: $out"
cd "$HOME/.claude/projects" || exit 0
R="rsync -au --exclude=memory --exclude=.DS_Store --exclude=*.jsonl"     # side files only; .jsonl never whole-file
for d in -Users-kevinkim*; do
  [ -d "$d" ] && [ ! -L "$d" ] || continue
  $R -- "$d/" "$H:.claude/projects/-home-kevinkim${d#-Users-kevinkim}/" 2>>"$LOG" || log "side push failed: $d"
done
for t in $(ssh $H 'cd ~/.claude/projects && for d in -home-kevinkim*; do [ -d "$d" ] && [ ! -L "$d" ] && echo "$d"; done'); do
  m="-Users-kevinkim${t#-home-kevinkim}"; mkdir -p -- "$m"
  $R -- "$H:.claude/projects/$t/" "$m/" 2>>"$LOG" || log "side pull failed: $t"
done
log "ok"
