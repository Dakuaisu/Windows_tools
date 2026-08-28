<#
.SYNOPSIS
    Removes the weekly schedule. Touches no files.

.DESCRIPTION
    Deliberately does NOT un-sort anything. Moving hundreds of files back is a
    far bigger risk than leaving them tidy, and tidy is what was asked for.
    To reverse the sorting, use "Put Everything Back" instead.
#>
[CmdletBinding()]
param([switch]$NoPause)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Say { param([string]$T = '', [string]$C = '')
    if ($C) { Write-Host $T -ForegroundColor $C } else { Write-Host $T } }

$removed = $false
foreach ($name in @('Tidy My Downloads', 'Downloads Janitor')) {
    try {
        if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction Stop
            Say "   Removed the weekly schedule: $name" 'Green'
            $removed = $true
        }
    } catch {
        try {
            & schtasks.exe /Delete /TN $name /F 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) { Say "   Removed the weekly schedule: $name" 'Green'; $removed = $true }
        } catch { }
    }
}

Say ''
if (-not $removed) {
    Say '   There was no weekly schedule to remove.' 'Cyan'
}
Say ''
Say '   Your files have NOT been touched.' 'Green'
Say '   Everything stays exactly where it is, tidy and findable.'
Say ''
Say '   If you would rather have your Downloads folder back the way'
Say '   it was, close this and double-click "Put Everything Back".'
Say ''

if (-not $NoPause) {
    Write-Host '   Press Enter to close this window...' -ForegroundColor DarkGray -NoNewline
    try { Read-Host | Out-Null } catch { }
}
exit 0
