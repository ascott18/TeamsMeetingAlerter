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
    $next = Get-UpcomingMeetings -Config $cfg
    Write-Host ("  OK - {0} Teams meeting(s) in the next {1} minutes." -f $next.Count, $cfg.LookaheadMinutes) -ForegroundColor Green
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
        $next = Get-UpcomingMeetings -Config $cfg
        if (-not $next.Count) {
            Write-Host "  No Teams meetings in the next $($cfg.LookaheadMinutes) minutes."
        } else {
            Write-Host "  Next $($next.Count) Teams meeting(s):"
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

$fired = Read-FiredState
$meetings = @()
$lastFetchUtc = [datetime]::MinValue
$failures = 0

while ($true) {
    $nowUtc = [datetime]::UtcNow

    $due = ($nowUtc - $lastFetchUtc).TotalSeconds -ge [int]$cfg.CalendarRefreshSeconds
    if ($due) {
        try {
            $meetings = Get-UpcomingMeetings -Config $cfg
            $lastFetchUtc = $nowUtc
            if ($failures) { Write-AlerterLog 'calendar reachable again' }
            $failures = 0
        } catch {
            $failures++
            $msg = $_.Exception.Message
            if ($msg -eq 'NotSignedIn' -or $msg -match 'invalid_grant|AADSTS') {
                Write-AlerterLog "sign-in needed - run: Watch.ps1 -Login  ($msg)" 'ERROR'
                $lastFetchUtc = $nowUtc.AddSeconds(-[int]$cfg.CalendarRefreshSeconds).AddMinutes(10)
            } else {
                # Back off on transient failures, but never past a few minutes.
                $backoff = [Math]::Min(300, 15 * $failures)
                Write-AlerterLog "calendar read failed (attempt $failures, retry in ${backoff}s): $msg" 'WARN'
                $lastFetchUtc = $nowUtc.AddSeconds(-[int]$cfg.CalendarRefreshSeconds).AddSeconds($backoff)
            }
        }
    }

    foreach ($m in $meetings) {
        if ($fired.ContainsKey($m.Key)) { continue }
        $fireAt = $m.StartUtc.AddSeconds(-[int]$cfg.LeadSeconds)
        $tooLate = $m.StartUtc.AddSeconds([int]$cfg.PostStartGraceSeconds)
        if ($nowUtc -ge $fireAt -and $nowUtc -le $tooLate -and $nowUtc -lt $m.EndUtc) {
            try { Start-MeetingAlert -Meeting $m } catch { Write-AlerterLog "failed to launch alert: $($_.Exception.Message)" 'ERROR' }
            $fired[$m.Key] = $m.StartUtc.ToString('o')
            $fired = Save-FiredState -Map $fired
        }
    }

    if ($Once) { break }
    Start-Sleep -Seconds ([int]$cfg.EvaluateIntervalSeconds)
}
