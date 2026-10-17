#!/bin/bash
# Installs a git pre-push hook: any push containing Vaultly changes runs scripts/qa.sh (both builds) first and is
# refused if it fails. Also refuses pushing Vaultly to any remote that isn't the personal codingintheusa0402 account.
set -e
ROOT=$(git rev-parse --show-toplevel); HOOKS=$(git rev-parse --git-path hooks); mkdir -p "$HOOKS"; HOOK="$HOOKS/pre-push"   # git-path: works in worktrees too
REL=$(python3 -c "import os,sys;print(os.path.relpath(sys.argv[1],sys.argv[2]))" "$(cd "$(dirname "$0")/.." && pwd)" "$ROOT")
PAT="^${REL}/"; [ "$REL" = "." ] && PAT="."      # Vaultly in its own repo: every change counts
cat > "$HOOK" <<H
#!/bin/bash
remote_url="\$2"
changed=0
pushed=""
while read -r lref lsha rref rsha; do
  [ "\$lsha" = "0000000000000000000000000000000000000000" ] && continue
  pushed="\$lsha"
  if [ "\$rsha" = "0000000000000000000000000000000000000000" ]; then
    base=\$(git merge-base "\$lsha" origin/main 2>/dev/null || true)   # new branch: compare with main
  else
    base="\$rsha"
  fi
  if [ -n "\$base" ]; then files=\$(git diff --name-only "\$base" "\$lsha" 2>/dev/null)
  else files=\$(git ls-tree -r --name-only "\$lsha"); fi             # no common base: be conservative
  echo "\$files" | grep -q "${PAT}" && changed=1
done
[ \$changed -eq 0 ] && exit 0
case "\$remote_url" in *codingintheusa0402*) ;; *) echo "⛔ Vaultly may only be pushed to codingintheusa0402 (not \$remote_url)"; exit 1;; esac
# A fresh PASS report for exactly the commit being pushed counts (qa.sh writes the hash into REPORT.md) — no 30-minute re-run.
latest="$ROOT/${REL#.}${REL:+/}build/qa/latest/REPORT.md"
if [ -f "\$latest" ] && grep -q '^\*\*PASS\*\*' "\$latest" && [ -n "\$pushed" ]; then
  short=\$(git rev-parse --short "\$pushed")
  if head -1 "\$latest" | grep -q "· \$short\$"; then echo "✅ QA report for \$short is a PASS — skipping re-run"; exit 0; fi
fi
echo "Vaultly changed — running the QA gate (scripts/qa.sh)…"
bash "$ROOT/${REL#.}${REL:+/}scripts/qa.sh" || { echo "⛔ QA failed — push refused. See $REL/build/qa/latest/REPORT.md"; exit 1; }
H
chmod +x "$HOOK"; echo "✅ pre-push QA hook installed in $HOOK"
