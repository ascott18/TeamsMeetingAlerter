# The alert window. Launched as a separate STA process by Watch.ps1, one per meeting.

param(
    [Parameter(Mandatory)][string]$PendingFile,
    [switch]$KeepPendingFile
)

. "$PSScriptRoot\Lib.ps1"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$cfg = Get-AlerterConfig
$meeting = Get-Content -LiteralPath $PendingFile -Raw -Encoding UTF8 | ConvertFrom-Json
# DateTimeOffset, not DateTime.Parse: the pending file carries a trailing Z, which
# Parse already converts to local, so relabelling it Utc shifts it a second time.
$startUtc = [datetimeoffset]::Parse($meeting.StartUtc, [Globalization.CultureInfo]::InvariantCulture).UtcDateTime
$startLocal = $startUtc.ToLocalTime()

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Meeting starting" SizeToContent="WidthAndHeight"
        WindowStyle="None" ResizeMode="NoResize" AllowsTransparency="True"
        Background="Transparent" ShowInTaskbar="True" Opacity="0">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Foreground" Value="#FFE8E8EE"/>
      <Setter Property="Background" Value="#FF32323C"/>
      <Setter Property="BorderBrush" Value="#FF4A4A56"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Height" Value="32"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" CornerRadius="5" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Chrome" Property="Opacity" Value="0.82"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Border CornerRadius="10" Background="#FF1C1C22" BorderThickness="1" Padding="18" Width="430">
    <Border.RenderTransform>
      <TranslateTransform x:Name="CardSlide" Y="0"/>
    </Border.RenderTransform>
    <Border.BorderBrush>
      <!-- Diagonal axis: a horizontal sweep lights both vertical edges at once, which
           reads wrong on a rectangle. The highlight band sits off the card at either
           end of the travel, so it passes occasionally rather than pulsing constantly. -->
      <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
        <LinearGradientBrush.RelativeTransform>
          <TranslateTransform x:Name="Shimmer" X="-1" Y="-1"/>
        </LinearGradientBrush.RelativeTransform>
        <GradientStop Color="#FF3C3C48" Offset="0.0"/>
        <GradientStop Color="#FF3C3C48" Offset="0.35"/>
        <GradientStop Color="#FF9CC0FF" Offset="0.5"/>
        <GradientStop Color="#FF3C3C48" Offset="0.65"/>
        <GradientStop Color="#FF3C3C48" Offset="1.0"/>
      </LinearGradientBrush>
    </Border.BorderBrush>
    <Border.Effect>
      <DropShadowEffect BlurRadius="18" ShadowDepth="3" Opacity="0.55" Color="Black"/>
    </Border.Effect>
    <Border.Triggers>
      <EventTrigger RoutedEvent="Border.Loaded">
        <BeginStoryboard>
          <Storyboard>
            <DoubleAnimation Storyboard.TargetName="Shimmer" Storyboard.TargetProperty="X"
                             From="-1" To="1" Duration="0:0:2.6" RepeatBehavior="Forever"/>
            <DoubleAnimation Storyboard.TargetName="Shimmer" Storyboard.TargetProperty="Y"
                             From="-1" To="1" Duration="0:0:2.6" RepeatBehavior="Forever"/>
          </Storyboard>
        </BeginStoryboard>
      </EventTrigger>
    </Border.Triggers>
    <StackPanel>
      <TextBlock x:Name="Kicker" Text="MEETING STARTING" Foreground="#FF8AB4F8" FontSize="11" FontWeight="Bold"/>
      <TextBlock x:Name="Subject" Foreground="#FFFFFFFF" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,7,0,0"/>
      <TextBlock x:Name="Meta" Foreground="#FFA8A8B4" FontSize="12" TextWrapping="Wrap" Margin="0,7,0,0"/>
      <TextBlock x:Name="Countdown" Foreground="#FFE8E8EE" FontSize="14" FontWeight="SemiBold" Margin="0,12,0,0"/>
      <StackPanel Orientation="Horizontal" Margin="0,18,0,0" HorizontalAlignment="Right">
        <Button x:Name="SnoozeBtn" Content="Snooze" Width="88" Margin="0,0,8,0"/>
        <Button x:Name="DismissBtn" Content="Dismiss" Width="88" Margin="0,0,8,0"/>
        <Button x:Name="JoinBtn" Content="Join" Width="104" Background="#FF2B5FD9" BorderBrush="#FF3F73EC" Foreground="White" FontWeight="SemiBold"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

function Get-WindowHandle {
    (New-Object System.Windows.Interop.WindowInteropHelper $window).Handle
}

function Format-Span {
    param([TimeSpan]$Span)
    if ($Span.TotalHours -ge 1) { return '{0}h {1}m' -f [int]$Span.TotalHours, $Span.Minutes }
    if ($Span.TotalMinutes -ge 1) { return '{0}m {1:00}s' -f [int]$Span.TotalMinutes, $Span.Seconds }
    '{0}s' -f [int][Math]::Ceiling($Span.TotalSeconds)
}

$subjectText = $window.FindName('Subject')
$metaText = $window.FindName('Meta')
$countdownText = $window.FindName('Countdown')
$kickerText = $window.FindName('Kicker')
$joinBtn = $window.FindName('JoinBtn')
$snoozeBtn = $window.FindName('SnoozeBtn')
$dismissBtn = $window.FindName('DismissBtn')
$cardSlide = $window.FindName('CardSlide')

$subjectText.Text = $meeting.Subject
$metaParts = @($startLocal.ToString('h:mm tt'))
if ($meeting.Organizer) { $metaParts += $meeting.Organizer }
$metaText.Text = $metaParts -join '  -  '
if (-not $meeting.JoinUrl) { $joinBtn.IsEnabled = $false; $joinBtn.Content = 'No link' }

$quiet = Test-QuietMode -Config $cfg
Write-AlerterLog ("alert window for '{0}' - quiet={1} ({2})" -f $meeting.Subject, $quiet.Quiet, $quiet.Reason)

# While held back, appear without stealing the foreground; the promotion timer
# below raises the window once that state clears.
$startQuiet = [bool]$quiet.Quiet
$window.Topmost = (-not $startQuiet) -and [bool]$cfg.Topmost
$window.ShowActivated = -not $startQuiet
if ($startQuiet) { $kickerText.Text = 'MEETING STARTING - WAITING FOR YOU' }

$script:sound = $null
if ($cfg.Sound -and (Test-Path $cfg.SoundFile) -and -not $startQuiet) {
    try {
        $script:sound = New-Object System.Media.SoundPlayer $cfg.SoundFile
        $script:sound.PlayLooping()
    } catch { }
}

function Stop-AlertSound {
    if ($script:sound) { try { $script:sound.Stop() } catch { } ; $script:sound = $null }
}

function Close-Alert {
    Stop-AlertSound
    if (-not $KeepPendingFile) { Remove-Item -LiteralPath $PendingFile -Force -ErrorAction SilentlyContinue }
    $window.Close()
}

$joinBtn.Add_Click({
    if ($meeting.JoinUrl) {
        try {
            $opened = Open-MeetingJoinUrl -Url $meeting.JoinUrl -PreferAppProtocol ([bool]$cfg.PreferAppProtocol)
            Write-AlerterLog ("joined '{0}' via {1}" -f $meeting.Subject, $opened)
        } catch {
            Write-AlerterLog ("join failed for '{0}': {1}" -f $meeting.Subject, $_.Exception.Message) 'WARN'
        }
    }
    Close-Alert
})

$dismissBtn.Add_Click({
    Write-AlerterLog ("dismissed '{0}'" -f $meeting.Subject)
    Close-Alert
})

$snoozeBtn.Add_Click({
    Stop-AlertSound
    $script:snoozeUntil = (Get-Date).AddSeconds([int]$cfg.SnoozeSeconds)
    $window.Hide()
    Write-AlerterLog ("snoozed '{0}' for {1}s" -f $meeting.Subject, $cfg.SnoozeSeconds)
})

$window.Add_MouseLeftButtonDown({ try { $window.DragMove() } catch { } })

# The window starts at Opacity 0. Its final position needs ActualWidth/ActualHeight,
# which SizeToContent only settles once the content has rendered - by which point a
# visible window would already have been painted at the default 0,0 and would visibly
# jump to the corner. So it is placed while still transparent, then revealed.
$window.Add_ContentRendered({
    $wa = [System.Windows.SystemParameters]::WorkArea
    # Stagger stacked alerts so a second meeting does not hide the first.
    $others = @(Get-ChildItem -LiteralPath (Get-AlerterPath 'state\pending') -Filter '*.json' -ErrorAction SilentlyContinue).Count
    $offset = [Math]::Max(0, $others - 1) * 24
    $window.Left = $wa.Right - $window.ActualWidth - 16
    $window.Top = $wa.Bottom - $window.ActualHeight - 16 - $offset

    try {
        $ease = New-Object System.Windows.Media.Animation.CubicEase
        $ease.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut

        $fade = New-Object System.Windows.Media.Animation.DoubleAnimation
        $fade.From = 0.0
        $fade.To = 1.0
        $fade.Duration = [System.Windows.Duration][TimeSpan]::FromMilliseconds(220)
        $window.BeginAnimation([System.Windows.Window]::OpacityProperty, $fade)

        $slide = New-Object System.Windows.Media.Animation.DoubleAnimation
        $slide.From = 16.0
        $slide.To = 0.0
        $slide.Duration = [System.Windows.Duration][TimeSpan]::FromMilliseconds(260)
        $slide.EasingFunction = $ease
        $cardSlide.BeginAnimation([System.Windows.Media.TranslateTransform]::YProperty, $slide)
    } catch {
        # Never leave the window stuck invisible if the animation cannot start.
        $window.Opacity = 1
    }

    if ($startQuiet) {
        try { Start-TaskbarFlash -WindowHandle (Get-WindowHandle) } catch { }
    }
})

$expiry = (Get-Date).AddMinutes([int]$cfg.AutoDismissMinutes)
$script:snoozeUntil = $null
$script:promoted = -not $startQuiet

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({
    $now = Get-Date
    if ($now -ge $expiry) {
        Write-AlerterLog ("auto-dismissed '{0}' after {1} min" -f $meeting.Subject, $cfg.AutoDismissMinutes)
        Close-Alert
        return
    }

    if ($script:snoozeUntil) {
        if ($now -ge $script:snoozeUntil) {
            $script:snoozeUntil = $null
            $window.Show()
        } else {
            $left = [int][Math]::Ceiling(($script:snoozeUntil - $now).TotalSeconds)
            $countdownText.Text = "Snoozed - back in ${left}s"
            return
        }
    }

    $delta = $startLocal - $now
    if ($delta.TotalSeconds -ge 0) {
        $countdownText.Text = 'Starts in {0}' -f (Format-Span $delta)
    } else {
        $countdownText.Text = 'Started {0} ago' -f (Format-Span $delta.Negate())
    }

    if (-not $script:promoted -and $cfg.PromoteWhenQuietEnds) {
        $q = Test-QuietMode -Config $cfg
        if (-not $q.Quiet) {
            $script:promoted = $true
            $kickerText.Text = 'MEETING STARTING'
            $window.Topmost = [bool]$cfg.Topmost
            try { Start-TaskbarFlash -WindowHandle (Get-WindowHandle) } catch { }
            Write-AlerterLog ("raised alert for '{0}' - quiet state cleared" -f $meeting.Subject)
        }
    }
})

$window.Add_Closed({
    Stop-AlertSound
    $timer.Stop()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
})

$timer.Start()
$window.Show()
if (-not $startQuiet) { try { [void]$window.Activate() } catch { } }
[System.Windows.Threading.Dispatcher]::Run()
