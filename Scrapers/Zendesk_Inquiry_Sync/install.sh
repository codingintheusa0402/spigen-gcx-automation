#!/bin/bash
# Installs the zendesk-inquiry-sync Claude Code skill for the current user:
# symlinks this folder's SKILL.md into ~/.claude/skills and installs Python deps.
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$HOME/.claude/skills/zendesk-inquiry-sync"
mkdir -p "$SKILL_DIR"
ln -sf "$DIR/SKILL.md" "$SKILL_DIR/SKILL.md"
if python3 -c "import googleapiclient, google.oauth2, google_auth_oauthlib" 2>/dev/null; then
  echo "Python deps already installed"
else
  python3 -m pip install --user -q -r "$DIR/requirements.txt" 2>/dev/null \
    || python3 -m pip install --user --break-system-packages -q -r "$DIR/requirements.txt" \
    || echo "WARNING: pip install failed — install manually: pip install -r $DIR/requirements.txt"
fi
echo "Installed skill → $SKILL_DIR (SKILL.md -> $DIR/SKILL.md)"
echo "Next: in Claude Code say \"set up zendesk inquiry sync\" (credentials + schedule)."
