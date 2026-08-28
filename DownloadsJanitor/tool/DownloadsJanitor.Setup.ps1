<#
.SYNOPSIS
    Tidy My Downloads - the thing a person actually runs.

.DESCRIPTION
    Three steps, all optional, nothing surprising:
      1. Show a preview. Always. Nothing is moved yet.
      2. Ask permission to tidy. Default is no.
      3. Ask permission to repeat weekly. Default is no.
#>
[CmdletBinding()]
param([switch]$NoPause)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Here = $PSScriptRoot
if (-not $Here) { $Here = Split-Path -Parent $MyInvocation.MyCommand.Path }

# Files that arrive by email or chat are flagged by Windows. Clearing that on
# our own folder needs no admin rights and stops the security prompts.
try {
    Get-ChildItem -LiteralPath (Split-Path -Parent $Here) -Recurse -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue
} catch { }

$Core     = Join-Path $Here 'DownloadsJanitor.Core.ps1'
$TaskName = 'Tidy My Downloads'

function Say { param([string]$T = '', [string]$C = '')
    if ($C) { Write-Host $T -ForegroundColor $C } else { Write-Host $T } }

function Ask {
    param([string]$Question)
    Say ''
    Write-Host "   $Question " -NoNewline -ForegroundColor White
    Write-Host '[y/N] ' -NoNewline -ForegroundColor DarkGray
    $a = ''
    try { $a = Read-Host } catch { return $false }
    if ($null -eq $a) { return $false }
    return (@('y','yes','ja','oui','si','sim','yep','ok') -contains $a.Trim().ToLowerInvariant())
}

function Pause-Here {
    if ($NoPause) { return }
    Say ''
    Write-Host '   Press Enter to close this window...' -ForegroundColor DarkGray -NoNewline
    try { Read-Host | Out-Null } catch { }
}

function Get-CurrentUserSid {
    # A SID always works. DOMAIN\user does not, on Microsoft-account or
    # Entra-joined laptops where the name is mangled or truncated.
    try { return ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value } catch { }
    try {
        foreach ($line in (& whoami.exe /user /fo list 2>$null)) {
            if ($line -match 'S-1-[\d-]+') { return $Matches[0] }
        }
    } catch { }
    return $null
}

function Register-WeeklyTask {
    param([string]$ScriptPath)

    $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arg = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Apply -Quiet' -f $ScriptPath
    $sid = Get-CurrentUserSid

    try {
        $action = New-ScheduledTaskAction -Execute $exe -Argument $arg -WorkingDirectory (Split-Path -Parent $ScriptPath)
        # No trailing Z: a local wall-clock time does not drift when the clocks change.
        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At ([datetime]'12:00:00')
        try { $trigger.RandomDelay = 'PT15M' } catch { }

        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
                        -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew `
                        -ExecutionTimeLimit (New-TimeSpan -Hours 2)

        if ($sid) {
            $principal = New-ScheduledTaskPrincipal -UserId $sid -LogonType Interactive -RunLevel Limited
        } else {
            $principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive -RunLevel Limited
        }

        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
            -Settings $settings -Principal $principal -Force `
            -Description 'Sorts settled files in Downloads into folders by type and month, and lists old clutter. Never deletes anything.' | Out-Null
        return $true
    }
    catch {
        # Some locked-down machines cannot load the ScheduledTasks module.
        try {
            $cmd = '"{0}" {1}' -f $exe, $arg
            & schtasks.exe /Create /TN $TaskName /TR $cmd /SC WEEKLY /D SUN /ST 12:00 /F | Out-Null
            return ($LASTEXITCODE -eq 0)
        } catch { return $false }
    }
}

# ===========================================================================

Say ''
Say '   ============================================='
Say '     TIDY MY DOWNLOADS'
Say '   ============================================='
Say ''
Say '   This sorts your Downloads folder into tidy folders'
Say '   like Photos, Documents and Spreadsheets.'
Say ''
Say '   It never deletes anything. Not ever.' 'Green'
Say '   Anything from the last 30 days stays exactly where it is.'
Say ''

if (-not (Test-Path -LiteralPath $Core)) {
    Say '   Some files are missing.' 'Red'
    Say '   Please extract the whole "Tidy My Downloads" folder from the'
    Say '   zip file, then run this again.'
    Pause-Here
    exit 2
}

# ------------------------------------------------------ step 1: always preview
Say '   ---------------------------------------------'
Say '   STEP 1 of 3   Have a look first'
Say '   ---------------------------------------------'
Say ''
Say '   Nothing will be moved yet. This is just a preview.' 'Cyan'

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Core
if ($LASTEXITCODE -eq 2) { Pause-Here; exit 2 }

# ------------------------------------------------------- step 2: ask to tidy
Say '   ---------------------------------------------'
Say '   STEP 2 of 3   Shall I tidy it up?'
Say '   ---------------------------------------------'
Say ''
Say '   If you say yes, the files listed above move into folders'
Say '   inside your Downloads folder. Nothing leaves Downloads,'
Say '   and nothing is deleted.'
Say ''
Say '   You can undo all of it at any time by double-clicking'
Say '   "Put Everything Back".'

if (-not (Ask 'Tidy up my Downloads folder now?')) {
    Say ''
    Say '   No problem - nothing was changed.' 'Cyan'
    Say '   Run this again whenever you like.'
    Pause-Here
    exit 0
}

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Core -Apply
$tidyExit = $LASTEXITCODE

if ($tidyExit -eq 2) {
    Say '   Nothing was changed.' 'Yellow'
    Pause-Here
    exit 2
}

# --------------------------------------------------- step 3: ask to schedule
Say '   ---------------------------------------------'
Say '   STEP 3 of 3   Keep it tidy automatically?'
Say '   ---------------------------------------------'
Say ''
Say '   I can do this once a week, quietly in the background,'
Say '   every Sunday around midday.'
Say ''
Say '   It will still never delete anything, and it will still'
Say '   leave the last 30 days of downloads alone.'

if (Ask 'Tidy up automatically once a week?') {
    if (Register-WeeklyTask -ScriptPath $Core) {
        Say ''
        Say '   Done. It will run every Sunday around midday.' 'Green'
        Say '   To stop it later, double-click "Remove Downloads Janitor".'
    } else {
        Say ''
        Say '   I could not set up the weekly schedule on this computer.' 'Yellow'
        Say '   (Some work laptops block this. Nothing else is affected.)'
        Say '   Your Downloads folder is still tidied - just run this'
        Say '   again whenever you want to tidy up.'
    }
} else {
    Say ''
    Say '   Fine - no automatic tidying.' 'Cyan'
    Say '   Run this whenever you want a clear-out.'
}

Say ''
Say '   ============================================='
Say '     All done.' 'Green'
Say '   ============================================='
Say ''
Say '   Open "Where are my files.html" in your Downloads folder'
Say '   to see where everything went, with a search box.'
Pause-Here
exit 0
