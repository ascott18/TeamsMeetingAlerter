# TeamsMeetingAlerter

Alerts you 30 seconds before a Teams meeting starts, because Teams often doesn't.

An always-on-top window with a **Join** button appears in the bottom-right corner. No sound.
When you're gaming it stays quiet and out of the way, then raises itself once you're out.

## Setup

One time, from this folder:

```powershell
.\Install.ps1
```

It signs you in (device code in a browser — no app registration needed, it uses
Microsoft's own public Graph PowerShell client) and registers a per-user scheduled
task that starts at logon. No admin rights required.

## Commands

| Command | What it does |
| --- | --- |
| `.\Watch.ps1 -Status` | Quiet-mode state, task/watcher state, next meetings |
| `.\Watch.ps1 -TestAlert` | Show a sample alert now |
| `.\Watch.ps1 -Login` | Re-authenticate |
| `.\Watch.ps1 -NoHide` | Run the loop in the foreground with visible logging |
| `.\Uninstall.ps1` | Remove the task and stop the watcher |
| `.\Uninstall.ps1 -ForgetSignIn` | Also delete the saved token |

## Gaming / do-not-disturb

Detection uses `SHQueryUserNotificationState` plus a geometry check, so it catches
both exclusive-fullscreen and borderless-windowed games — the latter looks like an
ordinary window to Windows and is missed by the API alone.

While quiet, the alert appears without stealing focus and flashes in the taskbar.
`PromoteWhenQuietEnds` then raises it as soon as you alt-tab out or exit the game.

Add process names to `QuietProcesses` in `config.json` (wildcards allowed, no `.exe`)
to force quiet mode for specific apps.

## config.json

| Key | Default | Meaning |
| --- | --- | --- |
| `LeadSeconds` | 30 | How far before start to alert |
| `Sound` / `SoundFile` | false | Looping alarm until you act |
| `Topmost` | true | Keep the window above others |
| `RespectGameMode` | true | Stand down for fullscreen apps and games |
| `RespectFocusAssist` | true | Stand down for Do Not Disturb |
| `PromoteWhenQuietEnds` | true | Raise the window once quiet mode clears |
| `QuietProcesses` | `[]` | Extra process names that force quiet mode |
| `OnlineMeetingsOnly` | true | Ignore meetings with no Teams join link |
| `SkipDeclined` / `SkipFreeShowAs` / `SkipAllDay` | true | Which events to ignore |
| `SubjectExcludePatterns` | `[]` | Regexes; matching subjects never alert |
| `PostStartGraceSeconds` | 120 | Don't alert for meetings already this far along |
| `SnoozeSeconds` | 60 | Snooze button duration |
| `AutoDismissMinutes` | 15 | Close an ignored alert after this long |
| `CalendarRefreshSeconds` | 120 | Graph poll interval |
| `LookaheadMinutes` | 120 | How far ahead to fetch |

Changes take effect on the next alert; no restart needed.

## What it can and can't see

It reads scheduled start times from your calendar via Microsoft Graph
(`Calendars.Read`, delegated). So it fires reliably at the *scheduled* start.

It does **not** know when a meeting is actually running. It won't catch an
unscheduled "Meet now", and if an organizer starts 10 minutes late you'll have been
alerted at the scheduled time. Graph exposes no live "is this meeting active" signal
to a delegated app, so that gap isn't closeable from here.

## Layout

    Lib.ps1        auth, Graph, config, logging, quiet-mode detection
    Watch.ps1      polling loop and CLI modes
    Alert.ps1      the WPF window (one process per meeting, launched STA)
    Install.ps1    sign-in + scheduled task registration
    state\         DPAPI-encrypted refresh token, fired-alert dedupe, pending alerts
    logs\          daily log, pruned after LogRetentionDays

The refresh token is encrypted with DPAPI scoped to this Windows account, so
`state\token.dat` is useless if copied to another machine or user.
