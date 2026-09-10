# Shared helpers: config, logging, Graph auth, calendar fetch, quiet-mode detection.

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$script:AlerterRoot = $PSScriptRoot
$script:GraphScope = 'https://graph.microsoft.com/Calendars.Read offline_access'
$script:AccessToken = $null
$script:AccessTokenExpiresUtc = [datetime]::MinValue

$script:ConfigDefaults = [ordered]@{
    ClientId                = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    Tenant                  = 'organizations'
    LeadSeconds             = 30
    PostStartGraceSeconds   = 120
    EvaluateIntervalSeconds = 5
    CalendarRefreshSeconds  = 120
    LookaheadMinutes        = 120
    OnlineMeetingsOnly      = $false
    SkipDeclined            = $true
    SkipFreeShowAs          = $true
    SkipAllDay              = $true
    Sound                   = $false
    SoundFile               = 'C:\Windows\Media\Alarm02.wav'
    Topmost                 = $true
    AutoDismissMinutes      = 15
    SnoozeSeconds           = 60
    RespectFullscreenApps   = $true
    RespectFocusAssist      = $true
    PromoteWhenQuietEnds    = $true
    QuietProcesses          = @()
    SubjectExcludePatterns  = @()
    PreferAppProtocol       = $true
    LogRetentionDays        = 14
}

function Get-AlerterPath {
    param([Parameter(Mandatory)][string]$Leaf)
    Join-Path $script:AlerterRoot $Leaf
}

function Write-AlerterLog {
    param([Parameter(Mandatory)][string]$Message, [string]$Level = 'INFO')
    $dir = Get-AlerterPath 'logs'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $file = Join-Path $dir ('alerter-{0}.log' -f (Get-Date).ToString('yyyy-MM-dd'))
    try { Add-Content -LiteralPath $file -Value $line -Encoding UTF8 } catch { }
    if ($Host.Name -eq 'ConsoleHost') { Write-Host $line }
}

function Get-AlerterConfig {
    $cfg = [ordered]@{}
    foreach ($k in $script:ConfigDefaults.Keys) { $cfg[$k] = $script:ConfigDefaults[$k] }
    $path = Get-AlerterPath 'config.json'
    if (Test-Path $path) {
        try {
            $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $json.PSObject.Properties) { $cfg[$p.Name] = $p.Value }
        } catch {
            Write-AlerterLog "config.json unreadable, using defaults: $($_.Exception.Message)" 'WARN'
        }
    }
    [pscustomobject]$cfg
}

function Remove-OldLogs {
    param([int]$RetentionDays = 14)
    $dir = Get-AlerterPath 'logs'
    if (-not (Test-Path $dir)) { return }
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $dir -Filter 'alerter-*.log' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Get-Prop {
    param($Object, [Parameter(Mandatory)][string]$Path)
    $cur = $Object
    foreach ($seg in $Path.Split('.')) {
        if ($null -eq $cur) { return $null }
        $prop = $cur.PSObject.Properties[$seg]
        if (-not $prop) { return $null }
        $cur = $prop.Value
    }
    $cur
}

# ---------- token storage (DPAPI, scoped to this Windows account) ----------

function Get-TokenFilePath { Get-AlerterPath 'state\token.dat' }

function Save-RefreshToken {
    param([Parameter(Mandatory)][string]$Token)
    $dir = Get-AlerterPath 'state'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $enc = ConvertTo-SecureString $Token -AsPlainText -Force | ConvertFrom-SecureString
    Set-Content -LiteralPath (Get-TokenFilePath) -Value $enc -Encoding ASCII -NoNewline
}

function Read-RefreshToken {
    $path = Get-TokenFilePath
    if (-not (Test-Path $path)) { return $null }
    try {
        # Trim matters: a trailing newline makes ConvertTo-SecureString reject the blob.
        $raw = Get-Content -LiteralPath $path -Raw
        if (-not $raw) { return $null }
        $sec = $raw.Trim() | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
        try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch {
        Write-AlerterLog "stored token could not be decrypted: $($_.Exception.Message)" 'WARN'
        $null
    }
}

# ---------- auth ----------

function Invoke-DeviceLogin {
    param($Config)
    $tenant = $Config.Tenant
    $body = @{ client_id = $Config.ClientId; scope = $script:GraphScope }
    $dc = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode" -Body $body

    Write-Host ''
    Write-Host '  Sign in so the alerter can read your calendar.' -ForegroundColor Cyan
    Write-Host ''
    Write-Host "    1. Open  $($dc.verification_uri)"
    Write-Host "    2. Enter code  $($dc.user_code)"
    Write-Host ''
    try { Set-Clipboard -Value $dc.user_code; Write-Host '  (code copied to clipboard)' -ForegroundColor DarkGray } catch { }
    try { Start-Process $dc.verification_uri | Out-Null } catch { }

    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    $interval = [Math]::Max(5, [int]$dc.interval)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        $poll = @{
            client_id   = $Config.ClientId
            grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
            device_code = $dc.device_code
        }
        try {
            $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body $poll
            Save-RefreshToken $tok.refresh_token
            $script:AccessToken = $tok.access_token
            $script:AccessTokenExpiresUtc = [datetime]::UtcNow.AddSeconds([int]$tok.expires_in)
            Write-Host ''
            Write-Host '  Signed in. Refresh token saved, encrypted to this Windows account.' -ForegroundColor Green
            return $true
        } catch {
            $detail = ''
            if ($_.ErrorDetails) { $detail = $_.ErrorDetails.Message }
            if ($detail -match 'authorization_pending') { continue }
            if ($detail -match 'slow_down') { $interval += 5; continue }
            if ($detail -match 'authorization_declined') { throw 'Sign-in was declined.' }
            if ($detail -match 'expired_token') { throw 'The device code expired. Run -Login again.' }
            throw "Sign-in failed: $detail"
        }
    }
    throw 'Sign-in timed out.'
}

function Get-AccessToken {
    param($Config)
    if ($script:AccessToken -and $script:AccessTokenExpiresUtc -gt [datetime]::UtcNow.AddMinutes(2)) {
        return $script:AccessToken
    }
    $refresh = Read-RefreshToken
    if (-not $refresh) { throw 'NotSignedIn' }
    $body = @{
        client_id     = $Config.ClientId
        grant_type    = 'refresh_token'
        refresh_token = $refresh
        scope         = $script:GraphScope
    }
    $tok = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($Config.Tenant)/oauth2/v2.0/token" -Body $body
    if ($tok.refresh_token) { Save-RefreshToken $tok.refresh_token }
    $script:AccessToken = $tok.access_token
    $script:AccessTokenExpiresUtc = [datetime]::UtcNow.AddSeconds([int]$tok.expires_in)
    $script:AccessToken
}

# ---------- calendar ----------

function ConvertFrom-GraphDateTime {
    param([string]$Value)
    if (-not $Value) { return $null }
    # Graph omits any offset while Prefer: outlook.timezone="UTC" is honored, but an
    # explicit Z or offset must not be converted and then relabelled as UTC.
    $dt = [datetime]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    switch ($dt.Kind) {
        'Utc'   { $dt }
        'Local' { $dt.ToUniversalTime() }
        default { [datetime]::SpecifyKind($dt, 'Utc') }
    }
}

function Get-MeetingJoinUrl {
    param($GraphEvent)
    $url = Get-Prop $GraphEvent 'onlineMeeting.joinUrl'
    if (-not $url) { $url = Get-Prop $GraphEvent 'onlineMeetingUrl' }
    if (-not $url) {
        $loc = Get-Prop $GraphEvent 'location.displayName'
        if ($loc -and $loc -match '(https://teams\.microsoft\.com/l/meetup-join/\S+)') { $url = $Matches[1] }
    }
    $url
}

function Get-UpcomingMeetings {
    param($Config)
    $token = Get-AccessToken -Config $Config
    $startUtc = [datetime]::UtcNow.AddMinutes(-10)
    $endUtc = [datetime]::UtcNow.AddMinutes([int]$Config.LookaheadMinutes)
    $select = 'id,subject,start,end,isAllDay,isCancelled,showAs,responseStatus,organizer,isOnlineMeeting,onlineMeetingUrl,onlineMeeting,location,webLink'
    $uri = 'https://graph.microsoft.com/v1.0/me/calendarView' +
           ('?startDateTime={0}&endDateTime={1}' -f $startUtc.ToString('yyyy-MM-ddTHH:mm:ssZ'), $endUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')) +
           '&$select=' + $select + '&$top=100'
    $headers = @{ Authorization = "Bearer $token"; Prefer = 'outlook.timezone="UTC"' }

    $raw = @()
    while ($uri) {
        $page = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
        if ($page.value) { $raw += $page.value }
        $uri = Get-Prop $page '@odata.nextLink'
    }

    $out = @()
    foreach ($e in $raw) {
        if ($Config.SkipAllDay -and (Get-Prop $e 'isAllDay')) { continue }
        if (Get-Prop $e 'isCancelled') { continue }
        if ($Config.SkipFreeShowAs -and (Get-Prop $e 'showAs') -eq 'free') { continue }
        if ($Config.SkipDeclined -and (Get-Prop $e 'responseStatus.response') -eq 'declined') { continue }

        $subject = Get-Prop $e 'subject'
        if (-not $subject) { $subject = '(no subject)' }
        $skip = $false
        foreach ($pat in @($Config.SubjectExcludePatterns)) {
            if ($pat -and $subject -match $pat) { $skip = $true; break }
        }
        if ($skip) { continue }

        $join = Get-MeetingJoinUrl -GraphEvent $e
        if ($Config.OnlineMeetingsOnly -and -not $join) { continue }

        $startAt = ConvertFrom-GraphDateTime (Get-Prop $e 'start.dateTime')
        $endAt = ConvertFrom-GraphDateTime (Get-Prop $e 'end.dateTime')
        if (-not $startAt) { continue }
        if (-not $endAt) { $endAt = $startAt.AddMinutes(30) }

        $out += [pscustomobject]@{
            Key       = '{0}|{1}' -f (Get-Prop $e 'id'), $startAt.ToString('o')
            Subject   = $subject
            Organizer = Get-Prop $e 'organizer.emailAddress.name'
            StartUtc  = $startAt
            EndUtc    = $endAt
            JoinUrl   = $join
            WebLink   = Get-Prop $e 'webLink'
        }
    }
    @($out | Sort-Object StartUtc)
}

# ---------- native interop / quiet-mode detection ----------

function Initialize-NativeTypes {
    if ('TMA.Native' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace TMA {
  public static class Native {
    [DllImport("shell32.dll")] public static extern int SHQueryUserNotificationState(out int state);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern IntPtr MonitorFromWindow(IntPtr hWnd, int flags);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO info);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool FlashWindowEx(ref FLASHWINFO info);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int cmd);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct MONITORINFO { public int cbSize; public RECT rcMonitor; public RECT rcWork; public uint dwFlags; }
    [StructLayout(LayoutKind.Sequential)] public struct FLASHWINFO { public uint cbSize; public IntPtr hwnd; public uint dwFlags; public uint uCount; public uint dwTimeout; }
  }
}
'@
}

function Hide-ConsoleWindow {
    Initialize-NativeTypes
    $h = [TMA.Native]::GetConsoleWindow()
    if ($h -ne [IntPtr]::Zero) { [void][TMA.Native]::ShowWindow($h, 0) }
}

function Get-ForegroundProcessName {
    Initialize-NativeTypes
    $hwnd = [TMA.Native]::GetForegroundWindow()
    if ($hwnd -eq [IntPtr]::Zero) { return $null }
    $procId = 0
    [void][TMA.Native]::GetWindowThreadProcessId($hwnd, [ref]$procId)
    if (-not $procId) { return $null }
    try { (Get-Process -Id $procId -ErrorAction Stop).ProcessName } catch { $null }
}

# A borderless-fullscreen window looks ordinary to SHQueryUserNotificationState, so a
# foreground window that exactly covers its monitor also counts as do-not-interrupt.
function Test-ForegroundIsFullScreen {
    Initialize-NativeTypes
    $hwnd = [TMA.Native]::GetForegroundWindow()
    if ($hwnd -eq [IntPtr]::Zero) { return $false }
    $wr = New-Object 'TMA.Native+RECT'
    if (-not [TMA.Native]::GetWindowRect($hwnd, [ref]$wr)) { return $false }
    $mon = [TMA.Native]::MonitorFromWindow($hwnd, 2)
    if ($mon -eq [IntPtr]::Zero) { return $false }
    $mi = New-Object 'TMA.Native+MONITORINFO'
    $mi.cbSize = [Runtime.InteropServices.Marshal]::SizeOf($mi)
    if (-not [TMA.Native]::GetMonitorInfo($mon, [ref]$mi)) { return $false }

    $covers = ($wr.Left -le $mi.rcMonitor.Left) -and ($wr.Top -le $mi.rcMonitor.Top) -and
              ($wr.Right -ge $mi.rcMonitor.Right) -and ($wr.Bottom -ge $mi.rcMonitor.Bottom)
    if (-not $covers) { return $false }

    # The desktop and shell surfaces always cover the monitor.
    $name = Get-ForegroundProcessName
    if (-not $name) { return $false }
    $shell = @('explorer', 'ShellExperienceHost', 'SearchHost', 'StartMenuExperienceHost', 'LockApp', 'TextInputHost', 'ApplicationFrameHost')
    if ($shell -contains $name) { return $false }
    $true
}

function Test-QuietMode {
    param($Config)
    Initialize-NativeTypes
    $state = 0
    $hr = [TMA.Native]::SHQueryUserNotificationState([ref]$state)

    if ($hr -eq 0) {
        if ($Config.RespectFullscreenApps -and $state -eq 3) { return [pscustomobject]@{ Quiet = $true; Reason = 'fullscreen Direct3D app' } }
        if ($Config.RespectFullscreenApps -and $state -eq 4) { return [pscustomobject]@{ Quiet = $true; Reason = 'presentation mode' } }
        if ($Config.RespectFullscreenApps -and $state -eq 2) { return [pscustomobject]@{ Quiet = $true; Reason = 'fullscreen app / busy' } }
        if ($Config.RespectFocusAssist -and $state -eq 6) { return [pscustomobject]@{ Quiet = $true; Reason = 'focus assist / do not disturb' } }
    }

    $fg = Get-ForegroundProcessName
    if ($fg) {
        foreach ($p in @($Config.QuietProcesses)) {
            if ($p -and $fg -like $p) { return [pscustomobject]@{ Quiet = $true; Reason = "foreground process '$fg' matches QuietProcesses" } }
        }
    }

    if ($Config.RespectFullscreenApps -and (Test-ForegroundIsFullScreen)) {
        return [pscustomobject]@{ Quiet = $true; Reason = "borderless-fullscreen app '$fg'" }
    }

    [pscustomobject]@{ Quiet = $false; Reason = 'notifications accepted' }
}

function Start-TaskbarFlash {
    param([Parameter(Mandatory)][IntPtr]$WindowHandle)
    Initialize-NativeTypes
    $fw = New-Object 'TMA.Native+FLASHWINFO'
    $fw.cbSize = [Runtime.InteropServices.Marshal]::SizeOf($fw)
    $fw.hwnd = $WindowHandle
    $fw.dwFlags = 15   # FLASHW_ALL | FLASHW_TIMERNOFG - keep flashing until brought forward
    $fw.uCount = 0
    $fw.dwTimeout = 0
    [void][TMA.Native]::FlashWindowEx([ref]$fw)
}

function Open-MeetingJoinUrl {
    param([Parameter(Mandatory)][string]$Url, [bool]$PreferAppProtocol = $true)
    if ($PreferAppProtocol -and $Url -like 'https://teams.microsoft.com/*') {
        $app = $Url -replace '^https://teams\.microsoft\.com/', 'msteams:/'
        try { Start-Process $app | Out-Null; return $app } catch { }
    }
    Start-Process $Url | Out-Null
    $Url
}
