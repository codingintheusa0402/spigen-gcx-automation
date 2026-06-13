# CLAUDE.md

## Session start — always do this first

At the start of every local (Mac) session, before doing anything else:

1. Check for changes pushed from remote-control (cloud) sessions:
   ```bash
   git fetch origin
   git log HEAD..origin/main --oneline
   ```
2. If there are new commits, pull them:
   ```bash
   git pull origin main
   ```
3. Also check if any feature branches were created remotely:
   ```bash
   git branch -r --no-merged main
   ```

This ensures work done via `/remote-control` (cloud Claude Code sessions) is synced locally before starting new work.
