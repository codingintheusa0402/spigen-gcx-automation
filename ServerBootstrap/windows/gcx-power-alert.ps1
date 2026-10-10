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
}
