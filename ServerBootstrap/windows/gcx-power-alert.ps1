# GCX Server power alerts (Windows side) — installed to C:\GCX-Setup\alerts\ by install-power-alerts.ps1
#   -Mode boot      : at Windows startup (SYSTEM, before anyone signs in) → 🔵 booting (+ was the last shutdown unexpected?)
#   -Mode shutdown  : on System event 1074 (restart / shutdown starting) → 🟡 rebooting or shutting down, with reason
# Webhook: C:\GCX-Setup\alerts\webhook.txt (not in git; readable by SYSTEM/Administrators only)
param([ValidateSet("boot", "shutdown")][string]$Mode)
$hook = (Get-Content "C:\GCX-Setup\alerts\webhook.txt" -Raw).Trim()
$now = Get-Date -Format "MM-dd HH:mm"
function Post($icon, $title, $text) {
  $card = @{ cardsV2 = @(@{ cardId = "gcx-pwr-$([DateTimeOffset]::Now.ToUnixTimeSeconds())"; card = @{
    header = @{ title = "$icon $title"; subtitle = "GCX Server · claude-server · $now" }
    sections = @(@{ widgets = @(@{ textParagraph = @{ text = $text } }) }) } }) }
  $body = [Text.Encoding]::UTF8.GetBytes(($card | ConvertTo-Json -Depth 10))
  for ($i = 0; $i -lt 12; $i++) {                       # at boot the network may need a minute
    try { Invoke-RestMethod -Uri $hook -Method Post -ContentType "application/json; charset=UTF-8" -Body $body -TimeoutSec 8 | Out-Null; return }
    catch { Start-Sleep 10 }
  }
}
if ($Mode -eq "shutdown") {
  $e = Get-WinEvent -FilterHashtable @{ LogName = "System"; Id = 1074 } -MaxEvents 1
  $p = $e.Properties
  $kind = if ("$($p[4].Value)" -match "restart") { "rebooting" } else { "shutting down" }
  Post "🟡" "GCX Server $kind" ("Started by <b>$($p[0].Value)</b> ($($p[6].Value)).<br>Reason: $($p[2].Value)" +
       "<br>You'll get 🔵 booting and 🟢 back-up messages when it's running again.")
}
if ($Mode -eq "boot") {
  Start-Sleep 20
  $boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
  $unexp = Get-WinEvent -FilterHashtable @{ LogName = "System"; Id = 41, 6008; StartTime = $boot.AddMinutes(-1) } -EA SilentlyContinue | Select -First 1
  $why = if ($unexp) { "⚠️ The previous shutdown was <b>unexpected</b> (power loss, crash or forced power-off)." } else { "Previous shutdown was a normal restart/shutdown." }
  Post "🔵" "GCX Server booting" ("Windows started at <b>$($boot.ToString('MM-dd HH:mm'))</b>. $why<br>Signing in and starting Ubuntu, jobs and Claude sessions — 🟢 back-up message follows when everything is running.")
  # stuck before sign-in? (wrong auto sign-in password, update screen, …) → tell the user how to finish it remotely
  for ($m = 0; $m -lt 12; $m++) {
    Start-Sleep 60
    if (Get-CimInstance Win32_Process -Filter "Name='explorer.exe' or Name='powershell.exe'" |
        Where-Object { $_.CommandLine -like "*shell.ps1*" -or $_.Name -eq "explorer.exe" } |
        Where-Object { (Invoke-CimMethod -InputObject $_ -MethodName GetOwner).User -eq "user" }) { exit }
  }
  $fail = Get-WinEvent -FilterHashtable @{ LogName = "Security"; Id = 4625; StartTime = $boot } -EA SilentlyContinue |
          Where-Object { $_.Properties[5].Value -eq "user" } | Select -First 1
  $hint = if ($fail) { "Auto sign-in was REJECTED (wrong stored password) — after signing in, ask Claude to reopen the auto sign-in window." } else { "Auto sign-in didn't happen." }
  Post "🔴" "GCX Server booted but NOT signed in" ("Windows has been at the sign-in screen for 12 min. $hint<br>" +
       "Sign in remotely: iPhone (Tailscale on) → VNC app → <b>claude-server</b> → type the Windows password on the sign-in screen. " +
       "Everything (dashboard, Ubuntu, jobs, Claude sessions) starts by itself right after.")
}
