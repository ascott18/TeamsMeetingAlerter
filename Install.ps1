# Registers the watcher as a per-user scheduled task. No elevation required.

[CmdletBinding()]
param([switch]$SkipLogin)

. "$PSScriptRoot\Lib.ps1"

$ErrorActionPreference = 'Stop'
$taskName = 'TeamsMeetingAlerter'
$cfg = Get-AlerterConfig

if (-not $SkipLogin -and -not (Read-RefreshToken)) {
    Invoke-DeviceLogin -Config $cfg | Out-Null
}

$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'

# Launched through conhost --headless, which creates no window at all. Hiding the
# console is not enough on Windows 11: conhost hands the session off to Windows
# Terminal, and the resulting tab belongs to Windows Terminal, so GetConsoleWindow()
# and -WindowStyle Hidden both act on the wrong window and the tab stays on screen.
# -STA is required: the tray icon's NotifyIcon and menu need a single-threaded apartment.
$argline = '--headless "{0}" -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "{1}"' -f $powershell, (Get-AlerterPath 'Watch.ps1')
$action = New-ScheduledTaskAction -Execute $conhost -Argument $argline -WorkingDirectory $PSScriptRoot

$userId = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn -User $userId

# Second trigger is a self-heal: if the watcher ever dies, this restarts it,
# and MultipleInstances=IgnoreNew makes it a no-op while one is already running.
$healTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(2) -RepetitionInterval (New-TimeSpan -Minutes 15)

$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($logonTrigger, $healTrigger) `
    -Principal $principal -Settings $settings -Force `
    -Description 'Alerts before Teams meetings start, because Teams often does not.' | Out-Null

# Re-running the installer has to replace a running watcher: the task's
# IgnoreNew policy would otherwise discard the start and leave old code running.
Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and $_.CommandLine -like '*Watch.ps1*' -and $_.CommandLine -notlike '*-Login*' } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Start-ScheduledTask -TaskName $taskName

Write-Host ''
Write-Host "  Installed '$taskName' - starts at logon and is running now." -ForegroundColor Green
Write-Host "  Alert fires $($cfg.LeadSeconds)s before each Teams meeting."
Write-Host ''
Write-Host '  Check it:   .\Watch.ps1 -Status'
Write-Host '  See it:     .\Watch.ps1 -TestAlert'
Write-Host '  Remove it:  .\Uninstall.ps1'
Write-Host ''
