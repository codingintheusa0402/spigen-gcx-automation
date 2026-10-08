# GCX git auto-sync

`gcx-autosync.sh` (repo root) runs every 3 minutes on both machines — Mac: launchd `com.spigen.gcx.git-autosync`,
server: cron (`# mesh:git-autosync`). It commits local changes (`git add -A`; `.gitignore` keeps
state/logs/root data files out), pulls the other machine's commits with `--rebase`, and pushes to origin
(codingintheusa0402 + spigenHQ). On a conflict it aborts, keeps the local commit unpushed and alerts the
private Chat room once. Log: `~/.gcx-autosync.log`.
