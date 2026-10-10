# Installs the boot + shutdown alert tasks (run as admin). Webhook file must already be at C:\GCX-Setup\alerts\webhook.txt
$dir = "C:\GCX-Setup\alerts"
icacls $dir /inheritance:r /grant "SYSTEM:(OI)(CI)F" /grant "Administrators:(OI)(CI)F" | Out-Null
$ps = "powershell.exe"; $f = "$dir\gcx-power-alert.ps1"
$sys = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -StartWhenAvailable
# boot
$a = New-ScheduledTaskAction -Execute $ps -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$f`" -Mode boot"
Register-ScheduledTask -TaskName "GCX-BootAlert" -Action $a -Trigger (New-ScheduledTaskTrigger -AtStartup) -Principal $sys -Settings $set -Force | Out-Null
# shutdown / restart initiated (System log, User32 event 1074)
$cls = Get-CimClass -Namespace Root/Microsoft/Windows/TaskScheduler -ClassName MSFT_TaskEventTrigger
$t = New-CimInstance -CimClass $cls -ClientOnly
$t.Subscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">*[System[Provider[@Name=''User32''] and EventID=1074]]</Select></Query></QueryList>'
$t.Enabled = $true
$a2 = New-ScheduledTaskAction -Execute $ps -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$f`" -Mode shutdown"
Register-ScheduledTask -TaskName "GCX-ShutdownAlert" -Action $a2 -Trigger $t -Principal $sys -Settings $set -Force | Out-Null
"installed: GCX-BootAlert (at startup), GCX-ShutdownAlert (event 1074)"
