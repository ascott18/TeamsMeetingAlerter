# Removes the scheduled task and stops the watcher. Leaves config, logs and the saved token.

[CmdletBinding()]
param([switch]$ForgetSignIn)

. "$PSScriptRoot\Lib.ps1"

$taskName = 'TeamsMeetingAlerter'

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "  Removed scheduled task '$taskName'." -ForegroundColor Green
} else {
    Write-Host "  Scheduled task '$taskName' was not installed."
}

Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -and $_.CommandLine -like '*Watch.ps1*' } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Host "  Stopped watcher process $($_.ProcessId)."
    }

if ($ForgetSignIn) {
    Remove-Item -LiteralPath (Get-TokenFilePath) -Force -ErrorAction SilentlyContinue
    Write-Host '  Deleted the saved sign-in token.' -ForegroundColor Green
}
