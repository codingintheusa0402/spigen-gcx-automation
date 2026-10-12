#!/usr/bin/env bash
# Live log for a RUNNING NOW row of the GCX console (opened in a Windows Terminal when the row is clicked).
#   gcx-logview.sh <key>    keys come from gcx-jobs.py (4th field):
#     svc:<unit>:user|system   systemd service (AI Trader → its engine log files)
#     tmux:<window>            read-only live view of that gcx tmux window (Claude sessions, Zendesk Chrome)
#     pid:<pid>                a scheduled job → the file its output goes to (cron >> redirect)
#     health                   watchdog / heartbeat / self-heal / git sync / Zendesk check logs
# Ctrl+C stops following; the window then waits for Enter.
k="$1"; H=$HOME
banner() { printf '\e[1;36m── %s ──\e[0m  (Ctrl+C to stop)\n' "$1"; }
case "$k" in
  svc:mock-trader:*)
    banner "AI Trader engine · ~/mock-trader/state/logs/engine.log + engine.err"
    tail -n 120 -F $H/mock-trader/state/logs/engine.log $H/mock-trader/state/logs/engine.err ;;
  svc:*:system)
    u=${k#svc:}; u=${u%:system}; banner "service $u (journal)"; sudo -n journalctl -u "$u" -n 120 -f ;;
  svc:*:user)
    u=${k#svc:}; u=${u%:user}; banner "service $u (journal)"; journalctl --user -u "$u" -n 120 -f ;;
  tmux:*)
    w=${k#tmux:}; banner "tmux gcx:$w — read-only live view (attach to type: tmux attach -t gcx)"; sleep 1
    trap 'tput cnorm; exit' INT
    tput civis
    while tmux list-windows -t gcx -F '#W' 2>/dev/null | grep -qx "$w"; do
      out=$(tmux capture-pane -p -e -J -t "gcx:$w" -S -400 | tail -n $(( $(tput lines) - 1 )))
      printf '\e[H\e[2J%s' "$out"; sleep 1
    done; tput cnorm; echo; echo "window $w closed" ;;
  pid:*)
    p=${k#pid:}
    if [ ! -d /proc/$p ]; then echo "That job has already finished."; else
      cmd=$(tr '\0' ' ' < /proc/$p/cmdline | cut -c1-160)
      f=$(readlink /proc/$p/fd/1 2>/dev/null)
      # cron runs jobs as  sh -c "cd … && python3 job.py >> log"  — the redirect may sit on the child
      if [ ! -f "$f" ]; then for c in $(pgrep -P $p) ; do g=$(readlink /proc/$c/fd/1 2>/dev/null); [ -f "$g" ] && f=$g && break; done; fi
      banner "$cmd"
      if [ -f "$f" ]; then echo "log: $f"; tail -n 150 -F --pid=$p "$f"; echo; echo "— job finished —"
      else echo "This job doesn't write to a log file (output: ${f:-none}). Showing its progress until it ends…"
        while [ -d /proc/$p ]; do printf '\r  still running · %s elapsed   ' "$(ps -o etime= -p $p | tr -d ' ')"; sleep 2; done; echo; echo "— job finished —"; fi
    fi ;;
  health)
    banner "health checks: watchdog · heartbeat · self-heal · git sync · Zendesk login check"
    tail -n 15 -F $(ls $H/.gcx-*.log /mnt/c/GCX-Setup/selfheal/selfheal.log 2>/dev/null) ;;
  *) echo "unknown log key: $k" ;;
esac
echo; read -r -p "Press Enter to close…" _
