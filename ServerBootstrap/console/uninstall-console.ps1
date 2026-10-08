# Restore the normal Windows desktop (Explorer) for user "user" at next sign-in.
Remove-ItemProperty "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name Shell -EA SilentlyContinue
"Normal Windows desktop restored (next sign-in). Run explorer.exe to get it right now."
