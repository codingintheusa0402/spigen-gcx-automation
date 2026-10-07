# Run ONCE in an *Administrator* PowerShell on the Windows server laptop.
# Installs WSL2 Ubuntu + Tailscale, disables sleep, keeps WSL alive 24/7.
$ErrorActionPreference = "Stop"

# --- never sleep / hibernate on AC, lid close = do nothing ---
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
powercfg /change monitor-timeout-ac 10
powercfg /setacvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0
powercfg /setactive SCHEME_CURRENT

# --- Tailscale (Windows side, so you can also RDP in) ---
winget install --id Tailscale.Tailscale -e --accept-source-agreements --accept-package-agreements

# --- WSL2 never idles out ---
@"
[wsl2]
vmIdleTimeout=-1
memory=8GB
"@ | Set-Content -Encoding ascii "$env:USERPROFILE\.wslconfig"

# --- keep Ubuntu running at every logon (hidden window) ---
$action  = New-ScheduledTaskAction -Execute "wsl.exe" -Argument "-d Ubuntu --exec /bin/sleep infinity"
$trigger = New-ScheduledTaskTrigger -AtLogOn
$set     = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0 -Hidden
Register-ScheduledTask -TaskName "WSL-KeepAlive" -Action $action -Trigger $trigger -Settings $set -Force

# --- WSL2 + Ubuntu (asks you to reboot, then to create a Linux username) ---
wsl --install -d Ubuntu

Write-Host "`nDONE. Reboot, open 'Ubuntu' from Start, create user 'kevinkim', then run step 2."
