# GCX Server console — full-screen launcher that replaces the Windows desktop (Explorer shell) for user "user".
# Big touch-friendly tiles for driving the server over VNC from a phone, plus a live status strip.
# Installed by install-console.ps1 (HKCU Winlogon\Shell). Escape hatch: the "Windows desktop" tile, or
# Ctrl+Shift+Esc → Task Manager → Run "explorer.exe". Undo: uninstall-console.ps1.
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$WT   = "$env:LOCALAPPDATA\Microsoft\WindowsApps\wt.exe"
$WSL  = "C:\Windows\System32\wsl.exe"
$UB   = @("-d", "Ubuntu", "-u", "kevinkim", "--")
$AMZ  = "https://www.amazon.com/gp/css/homepage.html https://www.amazon.de/gp/css/homepage.html https://www.amazon.co.uk/gp/css/homepage.html https://www.amazon.it/gp/css/homepage.html https://www.amazon.fr/gp/css/homepage.html https://www.amazon.es/gp/css/homepage.html https://www.amazon.in/gp/css/homepage.html"

function Term($title, $cmd) { Start-Process $WT -ArgumentList "--title `"$title`" wsl.exe -d Ubuntu -u kevinkim -- bash -lc `"$cmd`"" }
function Ub($argline)       { Start-Process $WSL -ArgumentList ("-d Ubuntu -u kevinkim -- " + $argline) -WindowStyle Hidden }

$tiles = @(
  @{ t = "New Claude Session";      s = "fresh session · phone-reachable";   i = [char]0xE710; a = { Term "New Claude session" "~/gcx-launch.sh new" } },
  @{ t = "Resume Claude Session";   s = "any past Mac or server session";    i = [char]0xE81C; a = { Term "Resume a Claude session" "~/gcx-launch.sh resume" } },
  @{ t = "GCX Live";                s = "all running sessions + job logs";   i = [char]0xE7F4; a = { Term "GCX Live" "~/gcx-launch.sh live" } },
  @{ t = "Chrome (GCX profile)";    s = "kjw@spigen.com · scraper profile";  i = [char]0xE774; a = { Ub "google-chrome --user-data-dir=/home/kevinkim/.chrome-scraper-profile" } },
  @{ t = "Amazon logins";           s = "photo-check / scraper sessions";    i = [char]0xE8D4; a = { Ub "google-chrome --user-data-dir=/home/kevinkim/.chrome-phaseg-profile $AMZ" } },
  @{ t = "Tailscale";               s = "network · admin console";           i = [char]0xE968; a = { Start-Process "https://login.tailscale.com/admin/machines"; $ts = "C:\Program Files\Tailscale\tailscale-ipn.exe"; if (Test-Path $ts) { Start-Process $ts } } },
  @{ t = "Ubuntu Terminal";         s = "shell on gcx-server";               i = [char]0xE756; a = { Start-Process $WT -ArgumentList "wsl.exe -d Ubuntu -u kevinkim --cd ~/Desktop/GCX" } },
  @{ t = "GCX Repo Files";          s = "~/Desktop/GCX (git, auto-synced)";  i = [char]0xE8B7; a = { Start-Process "explorer.exe" -ArgumentList "\\wsl.localhost\Ubuntu\home\kevinkim\Desktop\GCX" } },
  @{ t = "Windows Desktop";         s = "normal Windows (until sign-out)";   i = [char]0xE7F8; a = { Start-Process "explorer.exe" } },
  @{ t = "Restart Server";          s = "asks first";                         i = [char]0xE777; a = {
        if ([System.Windows.MessageBox]::Show("Restart the server now? Jobs and sessions stop for ~2 minutes.", "GCX Server", "YesNo", "Warning") -eq "Yes") { shutdown /r /t 5 } } }
)

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="GCX Server" WindowStyle="None" WindowState="Maximized" ResizeMode="NoResize" Background="#0B0D17" Topmost="False">
  <Grid Margin="40,30,40,24">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <DockPanel Grid.Row="0" Margin="4,0,4,22">
      <TextBlock x:Name="Clock" DockPanel.Dock="Right" Foreground="#8C93B8" FontSize="22" FontFamily="Segoe UI Light" VerticalAlignment="Bottom"/>
      <StackPanel>
        <TextBlock Text="GCX SERVER" Foreground="#E9EAFF" FontSize="34" FontFamily="Segoe UI Semibold"/>
        <TextBlock x:Name="Sub" Foreground="#6F77A3" FontSize="15" FontFamily="Segoe UI" Margin="2,2,0,0"/>
      </StackPanel>
    </DockPanel>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"><WrapPanel x:Name="Tiles"/></ScrollViewer>
    <Border Grid.Row="2" Background="#121528" CornerRadius="12" Padding="18,12" Margin="4,16,4,0">
      <TextBlock x:Name="Status" Foreground="#9AA2CC" FontSize="14" FontFamily="Cascadia Mono, Consolas" TextWrapping="Wrap"/>
    </Border>
  </Grid>
</Window>
"@
$win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$wrap = $win.FindName("Tiles"); $status = $win.FindName("Status"); $clock = $win.FindName("Clock"); $sub = $win.FindName("Sub")

$bc = New-Object Windows.Media.BrushConverter
foreach ($x in $tiles) {
  $b = New-Object Windows.Controls.Button
  $b.Width = 300; $b.Height = 150; $b.Margin = "10"; $b.Cursor = "Hand"; $b.BorderThickness = "1"
  $b.Background = $bc.ConvertFromString("#161A33"); $b.BorderBrush = $bc.ConvertFromString("#2A3060")
  $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = "20,0,0,0"; $sp.HorizontalAlignment = "Left"; $sp.VerticalAlignment = "Center"
  $ic = New-Object Windows.Controls.TextBlock; $ic.Text = $x.i; $ic.FontFamily = "Segoe Fluent Icons, Segoe MDL2 Assets"; $ic.FontSize = 30
  $ic.Foreground = $bc.ConvertFromString("#8FA2FF"); $ic.Margin = "0,0,0,12"
  $t1 = New-Object Windows.Controls.TextBlock; $t1.Text = $x.t; $t1.FontSize = 20; $t1.FontFamily = "Segoe UI Semibold"; $t1.Foreground = $bc.ConvertFromString("#EEF0FF")
  $t2 = New-Object Windows.Controls.TextBlock; $t2.Text = $x.s; $t2.FontSize = 13; $t2.Foreground = $bc.ConvertFromString("#7A82AE"); $t2.Margin = "0,4,0,0"
  [void]$sp.Children.Add($ic); [void]$sp.Children.Add($t1); [void]$sp.Children.Add($t2); $b.Content = $sp
  $act = $x.a; $b.Add_Click({ try { & $act } catch { [System.Windows.MessageBox]::Show($_.Exception.Message, "GCX Server") } }.GetNewClosure())
  [void]$wrap.Children.Add($b)
}

function Refresh {
  $clock.Text = (Get-Date).ToString("yyyy-MM-dd  HH:mm")
  try {
    $ts = (& "C:\Program Files\Tailscale\tailscale.exe" ip -4 2>$null | Select-Object -First 1)
    $sub.Text = "claude-server  ·  Tailscale $ts  ·  up $([int]((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalHours) h"
  } catch { }
  $job = Start-Job { wsl.exe -d Ubuntu -u kevinkim -- bash -lc 'echo "SESSIONS  $(tmux list-windows -t gcx -F "#W" 2>/dev/null | tr "\n" " ")"; echo "JOBS      $(crontab -l 2>/dev/null | grep -cE "^[0-9*]") scheduled  ·  last: $(sudo -n journalctl -u cron -n 200 --no-pager 2>/dev/null | grep -oE "[0-9:]{8} .*CMD \(cd \$G/[^ ]+" | tail -1 | sed -E "s#.*\\\$G/##")"; echo "GIT SYNC  $(tail -1 ~/.gcx-autosync.log 2>/dev/null)"' 2>$null }
  if (Wait-Job $job -Timeout 20) { $status.Text = ((Receive-Job $job) -join "`n") } ; Remove-Job $job -Force
}
$timer = New-Object Windows.Threading.DispatcherTimer; $timer.Interval = [TimeSpan]::FromSeconds(30); $timer.Add_Tick({ Refresh }); $timer.Start()
$win.Add_ContentRendered({ Refresh })
$win.Add_KeyDown({ if ($_.Key -eq "F5") { Refresh } })
[void]$win.ShowDialog()
