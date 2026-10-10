#!/usr/bin/env python3
"""GCX server heartbeat — cron every 2 min on gcx-server.

1. Writes now + Windows boot time to the private "GCX Server Heartbeat" sheet
   (~/.config/gcx_heartbeat_sheet.txt). The Apps Script "GCX Server Monitor" (GAS_Operations/
   GCXServerMonitor) reads it from Google's cloud and posts 🔴 DOWN when it goes quiet.
2. On the first run after a Windows restart, posts 🟢 "back up" with the downtime and current health.
"""
import json, os, subprocess, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gcx_alert import alert  # noqa: E402

STATE = os.path.expanduser("~/.gcx-heartbeat.json")
BOOT_CACHE = "/tmp/gcx-windows-boot-ms"            # /tmp is cleared whenever WSL restarts
PS = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
HEALTH = os.path.expanduser("~/Desktop/GCX/ServerBootstrap/console/gcx-health.sh")


def windows_boot_ms():
    try:
        return int(open(BOOT_CACHE).read())
    except Exception:
        pass
    out = subprocess.run([PS, "-NoProfile", "-Command",
                          "[int64](((Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime()) - [datetime]'1970-01-01').TotalMilliseconds"],
                         capture_output=True, text=True, timeout=60).stdout.strip()
    ms = int(out)
    open(BOOT_CACHE, "w").write(str(ms))
    return ms


def sheets():
    sys.path.insert(0, os.path.expanduser("~/Desktop/GCX/Scrapers/SC_Master_Propagate"))
    import propagate as P
    return P.get_service()


def dur(ms):
    m = round(ms / 60000)
    return f"{m} min" if m < 60 else f"{m // 60} h {m % 60} min"


def main():
    now = int(time.time() * 1000)
    boot = windows_boot_ms()
    try:
        st = json.load(open(STATE))
    except Exception:
        st = {}

    # 🟢 back up — first heartbeat after a new Windows boot (wait until Ubuntu has been up ≥ 2 min so services are settled)
    if st.get("boot") and st["boot"] != boot and float(open("/proc/uptime").read().split()[0]) >= 120:
        down_ms = boot - st.get("last", boot)
        health = subprocess.run(["bash", HEALTH], capture_output=True, text=True, timeout=60).stdout.splitlines()
        bad = [l.split("|", 2) for l in health if l.split("|")[1:2] and l.split("|")[1] != "ok"]
        lines = "<br>".join(f"• {k}: {t}" for k, _, t in bad) or "All checks OK — jobs, Claude sessions and ticket monitor running."
        try:
            alert("GCX Server back up",
                  f"Windows restarted at <b>{time.strftime('%m-%d %H:%M', time.localtime(boot / 1000))}</b> · "
                  f"offline about <b>{dur(max(0, now - st.get('last', boot)))}</b> (last heartbeat before restart "
                  f"{time.strftime('%m-%d %H:%M', time.localtime(st.get('last', boot) / 1000))}).<br>{lines}", level="ok")
            st["boot"] = boot
        except Exception as e:
            print("back-up alert failed:", e)       # retried next run (boot not updated)
    elif not st.get("boot"):
        st["boot"] = boot

    # heartbeat → sheet (read by the cloud monitor)
    sid = open(os.path.expanduser("~/.config/gcx_heartbeat_sheet.txt")).read().strip()
    sheets().spreadsheets().values().update(spreadsheetId=sid, range="heartbeat!A2:E2", valueInputOption="RAW",
        body={"values": [[now, boot, "gcx-server", "", time.strftime("%Y-%m-%d %H:%M:%S")]]}).execute()
    st["last"] = now
    json.dump(st, open(STATE, "w"))


if __name__ == "__main__":
    main()
