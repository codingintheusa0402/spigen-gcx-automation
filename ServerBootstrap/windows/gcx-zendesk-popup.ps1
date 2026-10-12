# GCX-ZendeskPopup — weekdays 08:30 (scheduled task, interactive session of `user`, deployed to C:\GCX-Setup\alerts\).
# Asks the ticket monitor's Chrome (WSL Ubuntu, CDP 9222) whether Zendesk is still signed in. If the session expired,
# zendesk_session.py --popup puts the sign-in page up, and this script raises that window on top of the GCX Server
# dashboard so the user can log in on the server screen (VNC or in the office). Signed in → nothing happens.
#   -Force   raise the window even when signed in (test)
param([switch]$Force)
$log = "C:\GCX-Setup\alerts\zendesk-popup.log"
function Log($m) { "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $m" | Out-File $log -Append -Encoding utf8 }

$cmd = "export PATH=`$HOME/.local/bin:/usr/bin:/bin; cd ~/Desktop/GCX/ServerBootstrap/watchdog && timeout 120 python3 zendesk_session.py --popup"
$out = & wsl.exe -d Ubuntu -u kevinkim -e bash -lc $cmd 2>&1
$code = $LASTEXITCODE
Log "check exit=$code $($out -join ' | ')"
if ($code -ne 1 -and -not $Force) { exit 0 }       # 0 = signed in, 2 = check failed (self-heal handles a dead Chrome)

Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class ZW {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
  public static List<string> Titles = new List<string>(); public static List<IntPtr> Handles = new List<IntPtr>();
  static bool Cb(IntPtr h, IntPtr l) { if (!IsWindowVisible(h)) return true; var sb = new StringBuilder(512); GetWindowText(h, sb, 512); if (sb.Length > 0) { Titles.Add(sb.ToString()); Handles.Add(h); } return true; }
  public static void Scan() { Titles.Clear(); Handles.Clear(); EnumWindows(Cb, IntPtr.Zero); }
}
"@
# the monitor's Chrome is a WSLg window: "… - Google Chrome (Ubuntu)". Prefer the sign-in page, else any Zendesk page.
$target = $null
foreach ($try in 1..6) {
  [ZW]::Scan()
  for ($i = 0; $i -lt [ZW]::Titles.Count; $i++) { if ([ZW]::Titles[$i] -match 'Sign in to Spigen.*Google Chrome \(Ubuntu\)') { $target = [ZW]::Handles[$i] } }
  if (-not $target) { for ($i = 0; $i -lt [ZW]::Titles.Count; $i++) { if ([ZW]::Titles[$i] -match 'Zendesk.*Google Chrome \(Ubuntu\)') { $target = [ZW]::Handles[$i] } } }
  if ($target) { break }
  Start-Sleep 2
}
if (-not $target) { Log "no Chrome (Ubuntu) window found"; exit 0 }
[ZW]::ShowWindow($target, 9) | Out-Null                                       # restore if minimized
[ZW]::SetWindowPos($target, [IntPtr](-1), 0, 0, 0, 0, 0x0043) | Out-Null      # topmost …
[ZW]::SetWindowPos($target, [IntPtr](-2), 0, 0, 0, 0, 0x0043) | Out-Null      # … then normal z-order, now above the dashboard
[ZW]::SetForegroundWindow($target) | Out-Null
Log "raised the Zendesk sign-in window"
