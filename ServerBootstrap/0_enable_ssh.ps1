# Run in an *Administrator* PowerShell on the Windows laptop (after Tailscale is logged in).
# Lets Claude on the Mac SSH in over Tailscale and do the whole setup remotely.
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Set-Service sshd -StartupType Automatic
Start-Service sshd
$k = 'ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAICQoGDozRXJvcqY0vCkf5+tQa+3vGs2eIIupxUDSvs8s mac->gcx-server'
Add-Content -Path "$env:ProgramData\ssh\administrators_authorized_keys" -Value $k
icacls "$env:ProgramData\ssh\administrators_authorized_keys" /inheritance:r /grant "Administrators:F" /grant "SYSTEM:F"
New-ItemProperty -Path "HKLM:\SOFTWARE\OpenSSH" -Name DefaultShell -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -PropertyType String -Force
Write-Host "SSH ready. Windows user: $env:USERNAME  |  Tailscale name: $(hostname)"
