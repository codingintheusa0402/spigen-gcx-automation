# GCX server self-heal (Windows side) — task GCX-SelfHeal: at sign-in + every 2 min, in the signed-in session, elevated.
# Restarts what stopped: services sshd / Tailscale / tvnserver, the GCX dashboard (shell.ps1 loop), Ubuntu (WSL-KeepAlive),
# then runs the Ubuntu side (gcx_selfheal.sh: cron, tailscaled, tmux gcx, Claude sessions, Zendesk Chrome, AI Trader).
# Copy lives in C:\GCX-Setup\selfheal\; source ServerBootstrap/selfheal/ in the GCX repo.
$ErrorActionPreference = "SilentlyContinue"
$log = "C:\GCX-Setup\selfheal\selfheal.log"
function L($m) { "$(Get-Date -f 'MM-dd HH:mm:ss') $m" | Add-Content $log }
if (((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalMinutes -lt 3) { exit }   # boot still settling
$fixed = @()
foreach ($s in "sshd", "Tailscale", "tvnserver") {
  $x = Get-Service $s
  if ($x -and $x.Status -ne "Running") { Start-Service $s; Start-Sleep 3; $ok = (Get-Service $s).Status -eq "Running"
    $fixed += "$s service" + $(if ($ok) { "" } else { " (restart FAILED)" }); L "service $s -> $((Get-Service $s).Status)" }
}
# GCX dashboard: normally the shell.ps1 loop relaunches GCXConsole. If no dashboard is running (twice, 5 s apart),
# restart the loop itself — it can hang waiting on a leftover child of a killed dashboard (seen 10-12)
$dash = { Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*GCXConsole.ps1*" } }
if (-not (& $dash)) { Start-Sleep 5
  if (-not (& $dash)) {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*console\shell.ps1*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\GCX-Setup\console\shell.ps1" -WindowStyle Hidden
    $fixed += "GCX SERVER dashboard"; L "dashboard restarted"
  }
}
# Ubuntu: started from this (signed-in) session so WSLg / Chrome have a display
$wsl = (wsl.exe -l --running 2>$null | Out-String) -replace "`0", ""
if ($wsl -notmatch "Ubuntu") { Start-ScheduledTask -TaskName "WSL-KeepAlive"; Start-Sleep 15; $fixed += "Ubuntu (WSL)"; L "WSL started" }
elseif ((Get-ScheduledTask -TaskName "WSL-KeepAlive").State -ne "Running") { Start-ScheduledTask -TaskName "WSL-KeepAlive" }
# Ubuntu side (it alerts for what it restarts itself)
wsl.exe -d Ubuntu -u kevinkim -e bash -lc "bash ~/Desktop/GCX/ServerBootstrap/selfheal/gcx_selfheal.sh" 2>$null | Out-Null
if ($fixed.Count) {
  try {
    $hook = (Get-Content "C:\GCX-Setup\alerts\webhook.txt" -Raw).Trim()
    $body = @{ cardsV2 = @(@{ cardId = "selfheal"; card = @{ header = @{ title = "$([char]::ConvertFromUtf32(0x1F7E2)) Self-heal: restarted what had stopped"; subtitle = "GCX Server · Windows · $(Get-Date -f 'MM-dd HH:mm')" };
      sections = @(@{ widgets = @(@{ textParagraph = @{ text = (($fixed | ForEach-Object { "• restarted: $_" }) -join "<br>") } }) }) } }) } | ConvertTo-Json -Depth 12
    Invoke-RestMethod -Uri $hook -Method Post -ContentType "application/json; charset=UTF-8" -Body ([Text.Encoding]::UTF8.GetBytes($body)) | Out-Null
  } catch { L "alert failed: $_" }
}
