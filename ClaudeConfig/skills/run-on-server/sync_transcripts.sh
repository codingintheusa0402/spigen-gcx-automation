#!/usr/bin/env bash
# Mac <-> server: Claude session histories, MERGED line by line (never whole-file overwrite: a session can
# continue on both machines, and newer-wins would drop one side). See ServerBootstrap/mesh_jsonl_sync.py.
# Side files (tool-results/, custom-title.json …) are copied both ways, newer wins, nothing deleted.
exec bash ~/Desktop/GCX/ServerBootstrap/mesh_sync_mac.sh --histories-only
