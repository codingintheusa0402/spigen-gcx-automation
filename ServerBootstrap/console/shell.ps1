# Shell host for the GCX console: keeps it open (relaunches if closed or crashed).
while ($true) {
  Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\GCX-Setup\console\GCXConsole.ps1" -WindowStyle Hidden -Wait
  Start-Sleep 2
}
