# Watcher: polls the calendar and fires one alert window per meeting.
#
#   .\Watch.ps1 -Login       sign in once (interactive)
#   .\Watch.ps1 -Status      show quiet-mode state and the next meetings
#   .\Watch.ps1 -TestAlert   show a sample alert window right now
#   .\Watch.ps1 -Once        evaluate a single pass and exit
#   .\Watch.ps1              run the watch loop (what the scheduled task does)

[CmdletBinding()]
param(
    [switch]$Login,
    [switch]$Status,
    [switch]$TestAlert,
    [switch]$Once,
    [switch]$NoHide
)

. "$PSScriptRoot\Lib.ps1"

$ErrorActionPreference = 'Stop'
$cfg = Get-AlerterConfig

function Get-FiredStatePath { Get-AlerterPath 'state\fired.json' }
function Get-PendingDir { Get-AlerterPath 'state\pending' }

function Read-FiredState {
    $path = Get-FiredStatePath
    $map = @{}
    if (Test-Path $path) {
        try {
            $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) { $map[$p.Name] = $p.Value }
        } catch { }
    }
    $map
}

function Save-FiredState {
    param([hashtable]$Map)
    $dir = Get-AlerterPath 'state'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    # Keys are only needed long enough to stop a repeat alert for the same instance.
    $cutoff = [datetime]::UtcNow.AddHours(-6)
    $keep = @{}
    foreach ($k in $Map.Keys) {
        try {
            if ([datetime]::Parse($Map[$k], [Globalization.CultureInfo]::InvariantCulture).ToUniversalTime() -ge $cutoff) { $keep[$k] = $Map[$k] }
        } catch { }
    }
    ($keep | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath (Get-FiredStatePath) -Encoding UTF8
    $keep
}

function Start-MeetingAlert {
    param([Parameter(Mandatory)]$Meeting)
    $dir = Get-PendingDir
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $pending = Join-Path $dir ('{0}.json' -f [guid]::NewGuid().ToString('N'))
    [pscustomobject]@{
        Subject   = $Meeting.Subject
        Organizer = $Meeting.Organizer
        StartUtc  = $Meeting.StartUtc.ToString('o')
        JoinUrl   = $Meeting.JoinUrl
        WebLink   = $Meeting.WebLink
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $pending -Encoding UTF8

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}" -PendingFile "{1}"' -f (Get-AlerterPath 'Alert.ps1'), $pending
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    [void][System.Diagnostics.Process]::Start($psi)
    Write-AlerterLog ("alerted: '{0}' at {1}" -f $Meeting.Subject, $Meeting.StartUtc.ToLocalTime().ToString('HH:mm:ss'))
}

# ---------- modes ----------

if ($Login) {
    Invoke-DeviceLogin -Config $cfg | Out-Null
    Write-Host ''
    Write-Host '  Checking calendar access...' -ForegroundColor DarkGray
    $next = @(Get-UpcomingMeetings -Config $cfg)
    Write-Host ("  OK - {0} meeting(s) in the next {1} minutes." -f $next.Count, $cfg.LookaheadMinutes) -ForegroundColor Green
    return
}

if ($Status) {
    $q = Test-QuietMode -Config $cfg
    Write-Host ''
    Write-Host '  Quiet mode : ' -NoNewline
    if ($q.Quiet) { Write-Host "YES - $($q.Reason)" -ForegroundColor Yellow }
    else { Write-Host "no - $($q.Reason)" -ForegroundColor Green }
    Write-Host "  Foreground : $(Get-ForegroundProcessName)"
    Write-Host "  Lead time  : $($cfg.LeadSeconds)s before start"

    $task = Get-ScheduledTask -TaskName 'TeamsMeetingAlerter' -ErrorAction SilentlyContinue
    if ($task) { Write-Host "  Task       : $($task.State)" } else { Write-Host '  Task       : not installed' -ForegroundColor Yellow }

    $running = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine -like '*Watch.ps1*' -and $_.CommandLine -notlike '*-Status*' }
    Write-Host "  Watcher    : $(if ($running) { "running (pid $($running.ProcessId -join ', '))" } else { 'not running' })"

    Write-Host ''
    if (-not (Read-RefreshToken)) {
        Write-Host '  Not signed in. Run:  .\Watch.ps1 -Login' -ForegroundColor Yellow
        return
    }
    try {
        $next = @(Get-UpcomingMeetings -Config $cfg)
        if (-not $next.Count) {
            Write-Host "  No meetings in the next $($cfg.LookaheadMinutes) minutes."
        } else {
            Write-Host "  Next $($next.Count) meeting(s):"
            foreach ($m in $next) {
                $mins = [int]($m.StartUtc - [datetime]::UtcNow).TotalMinutes
                Write-Host ('    {0}  (in {1,4} min)  {2}' -f $m.StartUtc.ToLocalTime().ToString('ddd HH:mm'), $mins, $m.Subject)
            }
        }
    } catch {
        Write-Host "  Calendar read failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ''
    return
}

if ($TestAlert) {
    Start-MeetingAlert -Meeting ([pscustomobject]@{
        Subject   = 'Test alert - this is what a real one looks like'
        Organizer = 'TeamsMeetingAlerter'
        StartUtc  = [datetime]::UtcNow.AddSeconds([int]$cfg.LeadSeconds)
        JoinUrl   = 'https://teams.microsoft.com/l/meetup-join/test'
        WebLink   = $null
    })
    Write-Host '  Test alert launched.' -ForegroundColor Green
    return
}

# ---------- watch loop ----------

if (-not $NoHide) { Hide-ConsoleWindow }
Remove-OldLogs -RetentionDays ([int]$cfg.LogRetentionDays)
Write-AlerterLog ('watcher started (lead {0}s, refresh {1}s, lookahead {2} min)' -f $cfg.LeadSeconds, $cfg.CalendarRefreshSeconds, $cfg.LookaheadMinutes)

$script:fired = Read-FiredState
$script:meetings = @()
$script:tracked = @{}
$script:firstFetch = $true
$script:lastFetchUtc = [datetime]::MinValue
$script:failures = 0
$script:health = 'ok'

function Invoke-WatchPass {
    $nowUtc = [datetime]::UtcNow

    $due = ($nowUtc - $script:lastFetchUtc).TotalSeconds -ge [int]$cfg.CalendarRefreshSeconds
    if ($due) {
        try {
            $script:meetings = @(Get-UpcomingMeetings -Config $cfg)
            $script:lastFetchUtc = $nowUtc
            if ($script:failures) { Write-AlerterLog 'calendar reachable again' }
            $script:failures = 0
            $script:health = 'ok'

            # Log the set only when it changes: a steady state stays quiet, but
            # silently seeing nothing becomes visible instead of looking idle.
            $current = @{}
            foreach ($m in $script:meetings) { $current[$m.Key] = $m }

            if ($script:firstFetch) {
                if ($script:meetings.Count) {
                    $listed = ($script:meetings | ForEach-Object { "{0} '{1}'" -f $_.StartUtc.ToLocalTime().ToString('HH:mm'), $_.Subject }) -join ', '
                    Write-AlerterLog ('tracking {0} meeting(s): {1}' -f $script:meetings.Count, $listed)
                } else {
                    Write-AlerterLog ('tracking 0 meetings in the next {0} min' -f $cfg.LookaheadMinutes)
                }
                $script:firstFetch = $false
            } else {
                foreach ($k in $current.Keys) {
                    if (-not $script:tracked.ContainsKey($k)) {
                        Write-AlerterLog ("found: '{0}' at {1}" -f $current[$k].Subject, $current[$k].StartUtc.ToLocalTime().ToString('HH:mm'))
                    }
                }
                foreach ($k in $script:tracked.Keys) {
                    if ($current.ContainsKey($k)) { continue }
                    # Meetings that already started just age out of the rolling
                    # window; only a future one vanishing means cancelled or moved.
                    if ($script:tracked[$k].StartUtc -gt $nowUtc) {
                        Write-AlerterLog ("gone: '{0}' was at {1}" -f $script:tracked[$k].Subject, $script:tracked[$k].StartUtc.ToLocalTime().ToString('HH:mm'))
                    }
                }
            }
            $script:tracked = $current
        } catch {
            $script:failures++
            $msg = $_.Exception.Message
            if ($msg -eq 'NotSignedIn' -or $msg -match 'invalid_grant|AADSTS') {
                Write-AlerterLog "sign-in needed - run: Watch.ps1 -Login  ($msg)" 'ERROR'
                $script:health = 'error'
                $script:lastFetchUtc = $nowUtc.AddSeconds(-[int]$cfg.CalendarRefreshSeconds).AddMinutes(10)
            } else {
                # Back off on transient failures, but never past a few minutes.
                $backoff = [Math]::Min(300, 15 * $script:failures)
                Write-AlerterLog "calendar read failed (attempt $($script:failures), retry in ${backoff}s): $msg" 'WARN'
                $script:health = 'warn'
                $script:lastFetchUtc = $nowUtc.AddSeconds(-[int]$cfg.CalendarRefreshSeconds).AddSeconds($backoff)
            }
        }
    }

    foreach ($m in $script:meetings) {
        if ($script:fired.ContainsKey($m.Key)) { continue }
        $fireAt = $m.StartUtc.AddSeconds(-[int]$cfg.LeadSeconds)
        $tooLate = $m.StartUtc.AddSeconds([int]$cfg.PostStartGraceSeconds)
        if ($nowUtc -ge $fireAt -and $nowUtc -le $tooLate -and $nowUtc -lt $m.EndUtc) {
            try { Start-MeetingAlert -Meeting $m } catch { Write-AlerterLog "failed to launch alert: $($_.Exception.Message)" 'ERROR' }
            $script:fired[$m.Key] = $m.StartUtc.ToString('o')
            $script:fired = Save-FiredState -Map $script:fired
        }
    }
}

# NotifyIcon.Text throws above 64 characters, so this must stay short.
function Get-TrayStatusText {
    if ($script:health -eq 'error') { return 'Meeting alerter: sign-in needed' }
    if ($script:health -eq 'warn') { return 'Meeting alerter: calendar unreachable' }
    if ($script:firstFetch) { return 'Meeting alerter: starting' }
    $next = @($script:meetings | Where-Object { $_.StartUtc -gt [datetime]::UtcNow })
    if (-not $next.Count) { return 'Meeting alerter: nothing upcoming' }
    $text = 'Next: {0} {1}' -f $next[0].StartUtc.ToLocalTime().ToString('HH:mm'), $next[0].Subject
    if ($text.Length -gt 63) { $text = $text.Substring(0, 60) + '...' }
    $text
}

if ($Once) {
    Invoke-WatchPass
    return
}

$tray = $null
if ($cfg.ShowTrayIcon) {
    try {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $icons = @{
            ok    = New-TrayStateIcon -State 'ok'
            warn  = New-TrayStateIcon -State 'warn'
            error = New-TrayStateIcon -State 'error'
        }

        $tray = New-Object System.Windows.Forms.NotifyIcon
        $tray.Icon = $icons['ok']
        $tray.Text = 'Meeting alerter: starting'
        $tray.Visible = $true

        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $miStatus = $menu.Items.Add('Starting...')
        $miStatus.Enabled = $false
        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

        $miRefresh = $menu.Items.Add('Refresh now')
        $miRefresh.Add_Click({ $script:lastFetchUtc = [datetime]::MinValue })

        $miTest = $menu.Items.Add('Test alert')
        $miTest.Add_Click({
            Start-MeetingAlert -Meeting ([pscustomobject]@{
                Subject   = 'Test alert - this is what a real one looks like'
                Organizer = 'TeamsMeetingAlerter'
                StartUtc  = [datetime]::UtcNow.AddSeconds([int]$cfg.LeadSeconds)
                JoinUrl   = 'https://teams.microsoft.com/l/meetup-join/test'
                WebLink   = $null
            })
        })

        $miLogs = $menu.Items.Add('Open logs folder')
        $miLogs.Add_Click({ Start-Process explorer.exe (Get-AlerterPath 'logs') })

        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $miExit = $menu.Items.Add('Exit')
        $miExit.Add_Click({
            Write-AlerterLog 'exited from tray menu'
            [System.Windows.Forms.Application]::Exit()
        })

        $tray.ContextMenuStrip = $menu
        $tray.Add_MouseDoubleClick({
            $tray.ShowBalloonTip(4000, 'TeamsMeetingAlerter', (Get-TrayStatusText), [System.Windows.Forms.ToolTipIcon]::Info)
        })
    } catch {
        Write-AlerterLog "tray icon unavailable, continuing without it: $($_.Exception.Message)" 'WARN'
        $tray = $null
    }
}

if ($tray) {
    # A tray icon needs a pumping message loop, so the pass runs on a WinForms
    # timer rather than a sleep loop.
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [Math]::Max(1000, 1000 * [int]$cfg.EvaluateIntervalSeconds)
    $timer.Add_Tick({
        try { Invoke-WatchPass } catch { Write-AlerterLog "watch pass failed: $($_.Exception.Message)" 'ERROR' }
        $status = Get-TrayStatusText
        $tray.Icon = $icons[$script:health]
        $tray.Text = $status
        $miStatus.Text = $status
    })
    $timer.Start()
    try {
        [System.Windows.Forms.Application]::Run()
    } finally {
        $timer.Stop()
        $tray.Visible = $false
        $tray.Dispose()
    }
} else {
    while ($true) {
        try { Invoke-WatchPass } catch { Write-AlerterLog "watch pass failed: $($_.Exception.Message)" 'ERROR' }
        Start-Sleep -Seconds ([int]$cfg.EvaluateIntervalSeconds)
    }
}
