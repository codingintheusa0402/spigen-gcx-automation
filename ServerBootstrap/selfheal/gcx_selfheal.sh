#!/usr/bin/env bash
# GCX server self-heal (Ubuntu side). Every 2 min from cron (# mesh:selfheal) AND from the Windows task GCX-SelfHeal
# (which also starts Ubuntu itself if it stopped). Restarts whatever is down and tells the GCX Server room once:
#   system services cron / tailscaled / gcx-relogin · user services mock-trader (unless the Mac owns it) /
#   gcx-owner-watch · tmux session gcx: Claude gcx-main, Zendesk Chrome (CDP 9222), ticket monitor (unless Mac owns it)
# A thing restarted 3× within 30 min is left alone (crash loop) and reported instead.
exec 9>/tmp/gcx-selfheal.lock; flock -n 9 || exit 0
export PATH=$HOME/.local/bin:/usr/local/bin:/usr/bin:/bin DISPLAY=:0 WAYLAND_DISPLAY=wayland-0
export XDG_RUNTIME_DIR=/run/user/$(id -u) DBUS_SESSION_BUS_ADDRESS=unix:path=/run/user/$(id -u)/bus
G=$HOME/Desktop/GCX; F=$G/ServerBootstrap/fallback; ST=$HOME/.gcx-selfheal; mkdir -p $ST
fixed=(); loops=()
# right after a (re)boot tmux_start.sh / systemd are still bringing things up — don't race them
[ "$(cut -d. -f1 /proc/uptime)" -lt 240 ] && exit 0
log() { echo "$(date '+%m-%d %H:%M:%S') $*" >> $HOME/.gcx-selfheal.log; }

# allowed <name>: false if <name> was restarted 3× in the last 30 min
allowed() { local f=$ST/$1 now=$(date +%s); touch $f
  awk -v n=$now '$1>n-1800' $f > $f.t; mv $f.t $f
  if [ $(wc -l < $f) -ge 3 ]; then [ -f $f.loop ] || { touch $f.loop; loops+=("$1"); }; return 1; fi
  rm -f $f.loop; echo $now >> $f; return 0; }
heal() { allowed "$1" || return; log "restart $1"; eval "$2" >/dev/null 2>&1 && fixed+=("$1") || fixed+=("$1 (restart FAILED)"); }

# keep our own cron line (in case crontab was rewritten)
crontab -l 2>/dev/null | grep -q "mesh:selfheal" || \
  ( crontab -l 2>/dev/null; echo "*/2 * * * * bash $G/ServerBootstrap/selfheal/gcx_selfheal.sh  # mesh:selfheal" ) | crontab -
crontab -l 2>/dev/null | grep -q "tmux_start" || \
  ( crontab -l 2>/dev/null; echo "@reboot sleep 30 && bash ~/tmux_start.sh" ) | crontab -

for s in cron tailscaled gcx-relogin; do
  systemctl list-unit-files "$s.service" >/dev/null 2>&1 || continue
  systemctl is-active -q $s || heal "$s" "sudo systemctl restart $s"
done
for s in gcx-owner-watch; do
  [ -f ~/.config/systemd/user/$s.service ] && ! systemctl --user is-active -q $s && heal "$s" "systemctl --user restart $s"
done
if [ -f ~/.config/systemd/user/mock-trader.service ] && [ ! -f ~/.gcx-owner/aitrader.stopped ] && \
   ! systemctl --user is-active -q mock-trader; then
  st=$(systemctl --user show -p ActiveState --value mock-trader)
  [ "$st" = activating ] || heal "AI Trader engine" "systemctl --user start --no-block mock-trader"
fi

# tmux session + its windows
if ! tmux has-session -t gcx 2>/dev/null; then
  heal "tmux session gcx (all windows)" "bash ~/tmux_start.sh"
else
  win() { tmux list-windows -t gcx -F '#W' | grep -qx "$1"; }
  run() { if win "$1"; then tmux respawn-window -k -t "gcx:$1" "$2"; else tmux new-window -d -t gcx -n "$1" "$2"; fi; }
  MAIN="cd $G && claude --remote-control gcx-main; exec bash"
  CHROME="/opt/google/chrome/chrome --user-data-dir=$HOME/.chrome-zendesk-profile --remote-debugging-port=9222 --no-first-run --no-default-browser-check https://spigenhelp.zendesk.com/agent/filters/360103290632; exec bash"
  TM="cd ~ && python3 $F/gcx_owner.py gate ticket-monitor; claude --remote-control gcx-ticket-monitor \"\$(cat ~/.gcx-tasks/ticket-monitor.prompt)\"; echo; echo [session ended]; exec bash"
  pgrep -f -- "--remote-control gcx-main" >/dev/null || heal "Claude session gcx-main" 'run main "$MAIN"'
  curl -s -m 5 http://127.0.0.1:9222/json/version >/dev/null || heal "Zendesk Chrome (CDP 9222)" 'run zendesk-chrome "$CHROME"'
  if [ -f ~/.gcx-tasks/ticket-monitor.prompt ] && [ ! -f ~/.gcx-owner/ticket-monitor.stopped ] && \
     ! pgrep -f -- "--remote-control gcx-ticket-monitor" >/dev/null && ! pgrep -f "gcx_owner.py gate ticket-monitor" >/dev/null; then
    sleep 5; curl -s -m 5 http://127.0.0.1:9222/json/version >/dev/null && heal "Zendesk ticket monitor" 'run ticket-monitor "$TM"'
  fi
fi

if [ ${#fixed[@]} -gt 0 ] || [ ${#loops[@]} -gt 0 ]; then
  body=""; for x in "${fixed[@]}"; do body+="• restarted: $x<br>"; done
  for x in "${loops[@]}"; do body+="• ⚠️ $x keeps stopping (3× in 30 min) — left alone, needs a look<br>"; done
  lvl=ok; [ ${#loops[@]} -gt 0 ] && lvl=bad
  python3 $G/ServerBootstrap/watchdog/gcx_alert.py "Self-heal: restarted what had stopped" "$body" --level $lvl >/dev/null 2>&1 || log "alert failed"
fi
