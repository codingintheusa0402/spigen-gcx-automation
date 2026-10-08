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
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="GCX SERVER" Foreground="#E9EAFF" FontSize="34" FontFamily="Segoe UI Semibold"/>
          <Grid Width="34" Height="34" Margin="18,6,6,0" VerticalAlignment="Center">
            <Ellipse x:Name="LiveRing" Width="14" Height="14" Fill="#3DDC97" Opacity="0.55" RenderTransformOrigin="0.5,0.5">
              <Ellipse.RenderTransform><ScaleTransform x:Name="LiveScale" ScaleX="1" ScaleY="1"/></Ellipse.RenderTransform>
            </Ellipse>
            <Ellipse x:Name="LiveDot" Width="14" Height="14" Fill="#3DDC97"/>
          </Grid>
          <TextBlock x:Name="LiveText" Text="LIVE" Foreground="#3DDC97" FontSize="14" FontFamily="Segoe UI Semibold" VerticalAlignment="Center" Margin="0,8,0,0"/>
        </StackPanel>
        <TextBlock x:Name="Sub" Foreground="#6F77A3" FontSize="15" FontFamily="Segoe UI" Margin="2,2,0,0"/>
      </StackPanel>
    </DockPanel>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"><WrapPanel x:Name="Tiles"/></ScrollViewer>
    <StackPanel Grid.Row="2" Margin="4,14,4,0">
      <DockPanel Margin="6,0,6,8">
        <TextBlock x:Name="Checked" DockPanel.Dock="Right" Foreground="#5D648C" FontSize="12" FontFamily="Segoe UI"/>
        <TextBlock Text="HEALTH &amp; STATUS" Foreground="#6F77A3" FontSize="13" FontFamily="Segoe UI Semibold"/>
      </DockPanel>
      <WrapPanel x:Name="Health"/>
      <Grid Margin="6,14,6,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <Border Grid.Column="0" Background="#121528" CornerRadius="10" Padding="16,12">
          <StackPanel><TextBlock x:Name="RunTitle" Text="RUNNING NOW" Foreground="#6F77A3" FontSize="13" FontFamily="Segoe UI Semibold" Margin="0,0,0,6"/>
            <ScrollViewer MaxHeight="172" VerticalScrollBarVisibility="Auto" PanningMode="VerticalOnly"><StackPanel x:Name="Running" Margin="0,0,8,0"/></ScrollViewer></StackPanel></Border>
        <Border Grid.Column="2" Background="#121528" CornerRadius="10" Padding="16,12">
          <StackPanel><TextBlock x:Name="NextTitle" Text="UP NEXT" Foreground="#6F77A3" FontSize="13" FontFamily="Segoe UI Semibold" Margin="0,0,0,6"/>
            <ScrollViewer MaxHeight="172" VerticalScrollBarVisibility="Auto" PanningMode="VerticalOnly"><StackPanel x:Name="Upnext" Margin="0,0,8,0"/></ScrollViewer></StackPanel></Border>
      </Grid>
    </StackPanel>
  </Grid>
</Window>
"@
$win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$wrap = $win.FindName("Tiles"); $health = $win.FindName("Health"); $running = $win.FindName("Running"); $upnext = $win.FindName("Upnext"); $runTitle = $win.FindName("RunTitle"); $nextTitle = $win.FindName("NextTitle"); $checked = $win.FindName("Checked"); $clock = $win.FindName("Clock"); $sub = $win.FindName("Sub")
$liveRing = $win.FindName("LiveRing"); $liveDot = $win.FindName("LiveDot"); $liveText = $win.FindName("LiveText"); $liveScale = $win.FindName("LiveScale")
function Anim($from, $to, $sec, $reverse) {
  $a = New-Object Windows.Media.Animation.DoubleAnimation($from, $to, (New-Object Windows.Duration([TimeSpan]::FromSeconds($sec))))
  $a.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever; $a.AutoReverse = $reverse
  $a.EasingFunction = New-Object Windows.Media.Animation.SineEase; return $a
}
# LIVE beacon: a ring radiating from the dot every 2 s (ring grows ×2.6 while fading out)
try {
  if (-not $liveScale) { $liveScale = $liveRing.RenderTransform }
  $liveScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleXProperty, (Anim 1.0 2.6 2.0 $false))
  $liveScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleYProperty, (Anim 1.0 2.6 2.0 $false))
  $liveRing.BeginAnimation([Windows.UIElement]::OpacityProperty, (Anim 0.55 0.0 2.0 $false))
} catch { try { Add-Content C:\GCX-Setup\console\console.log ("[" + (Get-Date).ToString("MM-dd HH:mm:ss") + "] beacon animation: " + $_.Exception.Message) } catch { } }
function SetBeacon($level) {
  if (-not $liveDot) { return }
  $c = @{ ok = "#3DDC97"; warn = "#FFC857"; bad = "#FF5C7A" }[$level]; $t = @{ ok = "LIVE"; warn = "LIVE · CHECK"; bad = "LIVE · ISSUE" }[$level]
  $b = $bc.ConvertFromString($c); $liveRing.Fill = $b; $liveDot.Fill = $b; $liveText.Foreground = $b; $liveText.Text = $t
}

$bc = New-Object Windows.Media.BrushConverter
foreach ($x in $tiles) {
  $b = New-Object Windows.Controls.Button
  $b.Width = 300; $b.Height = 150; $b.Margin = "10"; $b.Cursor = "Hand"; $b.BorderThickness = "1"
  $b.HorizontalContentAlignment = "Stretch"; $b.VerticalContentAlignment = "Center"; $b.Padding = "24,0,16,0"   # every tile: content flush left at the same x
  $b.Background = $bc.ConvertFromString("#161A33"); $b.BorderBrush = $bc.ConvertFromString("#2A3060")
  $sp = New-Object Windows.Controls.StackPanel; $sp.Margin = "0"; $sp.HorizontalAlignment = "Stretch"; $sp.VerticalAlignment = "Center"
  $ic = New-Object Windows.Controls.TextBlock; $ic.Text = $x.i; $ic.FontFamily = "Segoe Fluent Icons, Segoe MDL2 Assets"; $ic.FontSize = 30; $ic.HorizontalAlignment = "Left"; $ic.TextAlignment = "Left"
  $ic.Foreground = $bc.ConvertFromString("#8FA2FF"); $ic.Margin = "0,0,0,12"
  $t1 = New-Object Windows.Controls.TextBlock; $t1.Text = $x.t; $t1.FontSize = 20; $t1.FontFamily = "Segoe UI Semibold"; $t1.Foreground = $bc.ConvertFromString("#EEF0FF"); $t1.HorizontalAlignment = "Left"; $t1.TextAlignment = "Left"
  $t2 = New-Object Windows.Controls.TextBlock; $t2.Text = $x.s; $t2.FontSize = 13; $t2.Foreground = $bc.ConvertFromString("#7A82AE"); $t2.Margin = "0,4,0,0"; $t2.HorizontalAlignment = "Left"; $t2.TextAlignment = "Left"
  [void]$sp.Children.Add($ic); [void]$sp.Children.Add($t1); [void]$sp.Children.Add($t2); $b.Content = $sp
  $act = $x.a; $b.Add_Click({ try { & $act } catch { [System.Windows.MessageBox]::Show($_.Exception.Message, "GCX Server") } }.GetNewClosure())
  [void]$wrap.Children.Add($b)
}

$LABELS = [ordered]@{ power="Power"; net="Internet"; tailscale="Tailscale"; peers="Mac / iPhone"; remote="VNC · SSH";
  wsl="Ubuntu (WSL)"; claude="Claude"; monitor="Ticket monitor"; cron="Scheduled jobs"; git="Git sync"; sheets="Google Sheets"; disk="Disk · Memory" }
$COLORS = @{ ok = "#3DDC97"; warn = "#FFC857"; bad = "#FF5C7A"; na = "#5D648C" }

function Card($key, $level, $text) {
  $b = New-Object Windows.Controls.Border
  $b.Width = 300; $b.Margin = "6"; $b.Padding = "14,10"; $b.CornerRadius = "10"
  $b.Background = $bc.ConvertFromString("#121528")
  $b.BorderBrush = $bc.ConvertFromString($(if ($level -eq "bad") { "#5A2233" } else { "#1F2448" })); $b.BorderThickness = "1"
  $g = New-Object Windows.Controls.DockPanel
  $dot = New-Object Windows.Shapes.Ellipse; $dot.Width = 12; $dot.Height = 12; $dot.Margin = "0,4,12,0"; $dot.VerticalAlignment = "Top"
  $dot.Fill = $bc.ConvertFromString($COLORS[$level]); [Windows.Controls.DockPanel]::SetDock($dot, "Left")
  $sp = New-Object Windows.Controls.StackPanel
  $t1 = New-Object Windows.Controls.TextBlock; $t1.Text = $LABELS[$key]; $t1.FontSize = 14; $t1.FontFamily = "Segoe UI Semibold"; $t1.Foreground = $bc.ConvertFromString("#E4E6FF")
  $t2 = New-Object Windows.Controls.TextBlock; $t2.Text = $text; $t2.FontSize = 12.5; $t2.TextWrapping = "Wrap"; $t2.Foreground = $bc.ConvertFromString("#8C93B8"); $t2.Margin = "0,2,0,0"
  [void]$sp.Children.Add($t1); [void]$sp.Children.Add($t2); [void]$g.Children.Add($dot); [void]$g.Children.Add($sp); $b.Child = $g
  return $b
}

$HealthJob = {
  $r = @{}
  # power
  $bat = Get-CimInstance Win32_Battery -EA SilentlyContinue
  if ($bat) { $pct = $bat.EstimatedChargeRemaining; $ac = $bat.BatteryStatus -ne 1
    $r.power = if ($ac) { "ok|plugged in · battery $pct%" } elseif ($pct -le 25) { "bad|ON BATTERY $pct% — plug in!" } else { "warn|on battery $pct% — plug in" } }
  else { $r.power = "ok|AC power" }
  # internet
  try { $q = [Net.WebRequest]::Create("https://www.google.com/generate_204"); $q.Timeout = 4000; $sw = [Diagnostics.Stopwatch]::StartNew()
        $q.GetResponse().Close(); $r.net = "ok|online · $($sw.ElapsedMilliseconds) ms" } catch { $r.net = "bad|no internet" }
  # tailscale
  $ts = "C:\Program Files\Tailscale\tailscale.exe"
  try { $j = & $ts status --json 2>$null | ConvertFrom-Json
        $ip = ($j.Self.TailscaleIPs | ? { $_ -like "100.*" } | Select -First 1)
        $r.tailscale = if ($j.BackendState -eq "Running") { "ok|connected · $ip" } else { "bad|$($j.BackendState)" }
        $peers = $j.Peer.PSObject.Properties.Value
        $mac = $peers | ? { $_.OS -eq "macOS" } | Select -First 1; $ph = $peers | ? { $_.OS -eq "iOS" } | Select -First 1
        $f = { param($p, $n) if (-not $p) { "$n —" } elseif ($p.Online) { "$n online" } else { "$n offline" } }
        $r.peers = "ok|" + (& $f $mac "Mac") + " · " + (& $f $ph "iPhone")
  } catch { $r.tailscale = "bad|tailscale not responding"; $r.peers = "na|unknown" }
  # remote access services
  $v = (Get-Service tvnserver -EA SilentlyContinue).Status; $s = (Get-Service sshd -EA SilentlyContinue).Status
  $r.remote = if ($v -eq "Running" -and $s -eq "Running") { "ok|VNC + SSH running" } else { "bad|VNC $v · SSH $s" }
  # disk + memory
  $c = Get-PSDrive C; $free = [math]::Round($c.Free / 1GB); $os = Get-CimInstance Win32_OperatingSystem
  $mem = [math]::Round(100 - 100 * $os.FreePhysicalMemory / $os.TotalVisibleMemorySize)
  $r.disk = "$(if ($free -lt 10) { 'bad' } elseif ($free -lt 25) { 'warn' } else { 'ok' })|C: $free GB free · RAM $mem% used"
  # linux side
  $lx = wsl.exe -d Ubuntu -u kevinkim -- bash -lc "~/Desktop/GCX/ServerBootstrap/console/gcx-health.sh" 2>$null
  if ($lx) { foreach ($l in $lx) { $p = "$l".Split("|", 3); if ($p.Count -eq 3) { $r[$p[0]] = "$($p[1])|$($p[2])" } } }
  else { $r.wsl = "bad|Ubuntu not responding" }
  $r
}

function JobRow($name, $detail, $color) {
  $d = New-Object Windows.Controls.DockPanel; $d.Margin = "0,4,0,4"
  $dot = New-Object Windows.Shapes.Ellipse; $dot.Width = 9; $dot.Height = 9; $dot.Margin = "0,0,10,0"; $dot.Fill = $bc.ConvertFromString($color)
  [Windows.Controls.DockPanel]::SetDock($dot, "Left")
  if ($color -eq "#3DDC97") { try { $dot.BeginAnimation([Windows.UIElement]::OpacityProperty, (Anim 1.0 0.25 0.9 $true)) } catch { } }   # running → breathing dot
  $t2 = New-Object Windows.Controls.TextBlock; $t2.Text = $detail; $t2.FontSize = 13; $t2.Foreground = $bc.ConvertFromString("#8C93B8"); [Windows.Controls.DockPanel]::SetDock($t2, "Right")
  $t1 = New-Object Windows.Controls.TextBlock; $t1.Text = $name; $t1.FontSize = 14; $t1.Foreground = $bc.ConvertFromString("#E4E6FF")
  [void]$d.Children.Add($dot); [void]$d.Children.Add($t2); [void]$d.Children.Add($t1); return $d
}

function Refresh {
  $clock.Text = (Get-Date).ToString("yyyy-MM-dd  HH:mm")
  try { $sub.Text = "claude-server  ·  up $([int]((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).TotalHours) h  ·  tap a tile to start" } catch { }
  $job = Start-Job $HealthJob
  if (Wait-Job $job -Timeout 25) {
    $r = Receive-Job $job; $health.Children.Clear()
    foreach ($k in $LABELS.Keys) { $v = $r[$k]; if (-not $v) { $v = "na|—" }; $lv, $tx = "$v".Split("|", 2); [void]$health.Children.Add((Card $k $lv $tx)) }
    $lvls = @($LABELS.Keys | % { "$($r[$_])".Split("|")[0] })
    SetBeacon $(if ($lvls -contains "bad") { "bad" } elseif ($lvls -contains "warn") { "warn" } else { "ok" })
    $checked.Text = "checked " + (Get-Date).ToString("HH:mm:ss") + "  ·  every 30 s  ·  F5 to refresh"
  } else { Log "health refresh timed out" }
  Remove-Job $job -Force
}

function Log($m) { try { Add-Content C:\GCX-Setup\console\console.log ("[" + (Get-Date).ToString("MM-dd HH:mm:ss") + "] " + $m) } catch { } }

# RUNNING NOW / UP NEXT: own 1-minute update, read live from the server (processes + crontab) each time
function RefreshSchedule {
  try {
    $job = Start-Job { [Console]::OutputEncoding = [Text.Encoding]::UTF8; wsl.exe -d Ubuntu -u kevinkim -- python3 /home/kevinkim/Desktop/GCX/ServerBootstrap/console/gcx-jobs.py 2>&1 }
    if (-not (Wait-Job $job -Timeout 20)) { Log "schedule refresh timed out"; Remove-Job $job -Force; return }
    $lines = @(Receive-Job $job); Remove-Job $job -Force
    $ok = @($lines | ? { "$_" -match "^(run|next)\|" })
    if (-not $ok.Count -and $lines.Count) { Log ("schedule refresh output: " + (($lines | Select -First 3) -join " / ")) }
    $running.Children.Clear(); $upnext.Children.Clear()
    foreach ($l in $ok) { $p = "$l".Split("|", 3)
      if ($p[0] -eq "run")  { [void]$running.Children.Add((JobRow $p[1] $p[2] "#3DDC97")) }
      if ($p[0] -eq "next") { [void]$upnext.Children.Add((JobRow $p[1] $p[2] "#8FA2FF")) } }
    if ($running.Children.Count -eq 0) { [void]$running.Children.Add((JobRow "Nothing running" "" "#5D648C")) }
    if ($upnext.Children.Count -eq 0) { [void]$upnext.Children.Add((JobRow "No scheduled jobs" "" "#5D648C")) }
    $u = (Get-Date).ToString("HH:mm:ss")
    $runTitle.Text = "RUNNING NOW  ·  live, updated $u"; $nextTitle.Text = "UP NEXT  ·  live, updated $u"
  } catch { Log ("schedule refresh error: " + $_.Exception.Message) }
}
$timer = New-Object Windows.Threading.DispatcherTimer; $timer.Interval = [TimeSpan]::FromSeconds(30); $timer.Add_Tick({ Refresh }); $timer.Start()
$timer2 = New-Object Windows.Threading.DispatcherTimer; $timer2.Interval = [TimeSpan]::FromSeconds(60); $timer2.Add_Tick({ RefreshSchedule }); $timer2.Start()
$win.Add_ContentRendered({ Refresh; RefreshSchedule })
$win.Add_KeyDown({ if ($_.Key -eq "F5") { Refresh; RefreshSchedule } })
[void]$win.ShowDialog()
