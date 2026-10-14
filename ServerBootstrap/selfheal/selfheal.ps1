# GCX server self-heal (Windows side) — task GCX-SelfHeal: at sign-in + every 2 min, in the signed-in session, elevated.
# Restarts what stopped: services sshd / Tailscale / tvnserver, the GCX dashboard (shell.ps1 loop), Ubuntu (WSL-KeepAlive),
# then runs the Ubuntu side (gcx_selfheal.sh: cron, tailscaled, tmux gcx, Claude sessions, Zendesk Chrome, AI Trader).
# Copy lives in C:\GCX-Setup\selfheal\; source ServerBootstrap/selfheal/ in the GCX repo.
$ErrorActionPreference = "SilentlyContinue"
$log = "C:\GCX-Setup\selfheal\selfheal.log"
function L($m) { "$(Get-Date -f 'MM-dd HH:mm:ss') $m" | Add-Content $log }
if ([Environment]::TickCount64 -lt 180000) { exit }   # boot still settling (TickCount, not WMI: a CIM query can hang for minutes on a sick host)
# One-shot requests (2026-10-14): any *.ps1 dropped into C:\GCX-Setup\selfheal\requests\ runs once here — elevated, in the
# signed-in session — then its output sits next to it in requests\done\<name>.ps1.log. Written for the case where Windows
# SSH/VNC are down but Ubuntu still runs: Ubuntu writes the request through /mnt/c (write as .tmp, then rename to .ps1).
# The file is moved to done\ BEFORE it runs, so a request never repeats even if it crashes or restarts the machine.
$rq = "C:\GCX-Setup\selfheal\requests"; New-Item -ItemType Directory -Force "$rq\done" | Out-Null
foreach ($r in (Get-ChildItem "$rq\*.ps1" | Sort-Object Name)) {
  $dst = "$rq\done\$($r.Name)"; Move-Item $r.FullName $dst -Force; L "request $($r.Name): running"
  try { & $dst *>&1 | Out-String -Width 220 | Set-Content "$dst.log" -Encoding utf8; L "request $($r.Name): done" }
  catch { "FAILED: $_" | Add-Content "$dst.log"; L "request $($r.Name): FAILED $_" }
}
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
# Ubuntu reported a frozen WSLg display (no window can open) → restart Ubuntu from here, at most once per 30 min
$flag = "C:\GCX-Setup\selfheal\wslg-broken.flag"; $last = "C:\GCX-Setup\selfheal\wslg-restart.last"
if ((Test-Path $flag) -and (-not (Test-Path $last) -or ((Get-Date) - (Get-Item $last).LastWriteTime).TotalMinutes -gt 30)) {
  L "WSLg display frozen -> restarting Ubuntu"; Set-Content $last (Get-Date); Remove-Item $flag -Force
  wsl.exe --shutdown; Start-Sleep 8; Start-ScheduledTask -TaskName "WSL-KeepAlive"; Start-Sleep 20; Start-ScheduledTask -TaskName "GCX-Live"
  $fixed += "Ubuntu restarted (its display had frozen, so Chrome / Zendesk could not open)"
}
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
