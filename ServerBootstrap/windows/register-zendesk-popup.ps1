# One-off: register the GCX-ZendeskPopup task (weekdays 08:30, interactive session of `user`).
# Run over SSH: powershell -NoProfile -ExecutionPolicy Bypass -File register-zendesk-popup.ps1
$a = New-ScheduledTaskAction -Execute 'conhost.exe' -Argument '--headless powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\GCX-Setup\alerts\gcx-zendesk-popup.ps1'
$t = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 08:30
$s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
$p = New-ScheduledTaskPrincipal -UserId 'user' -LogonType Interactive -RunLevel Highest
Unregister-ScheduledTask -TaskName GCX-ZendeskPopup -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName GCX-ZendeskPopup -Action $a -Trigger $t -Settings $s -Principal $p -Description 'Weekdays 08:30: if the ticket monitor''s Zendesk session expired, raise its sign-in window on top of the GCX Server dashboard.' | Out-Null
Get-ScheduledTask GCX-ZendeskPopup | Select-Object TaskName, State | Format-Table -AutoSize
(Get-ScheduledTaskInfo GCX-ZendeskPopup).NextRunTime
