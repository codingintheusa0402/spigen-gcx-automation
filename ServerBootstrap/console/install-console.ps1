# Make the GCX console the desktop for user "user" (takes effect at next sign-in). Undo: uninstall-console.ps1
$k = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty $k -Name Shell -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\GCX-Setup\console\shell.ps1"
# GCX-Live: start the tmux sessions at logon without popping a window over the console
$a = New-ScheduledTaskAction -Execute "wsl.exe" -Argument "-d Ubuntu -u kevinkim -- bash -lc 'sleep 5; bash ~/tmux_start.sh'"
Set-ScheduledTask -TaskName "GCX-Live" -Action $a | Out-Null
"GCX console set as the desktop shell for 'user' (next sign-in)."
