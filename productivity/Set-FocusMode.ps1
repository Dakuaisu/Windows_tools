<#
.SYNOPSIS
    Turns distractions off - closes the apps you named, blocks the sites you
    named, silences toast notifications - and puts every one of them back
    afterwards.

.DESCRIPTION
    Everything this script changes is recorded in a state file before it is
    changed, and -Off restores from that record rather than from assumptions.
    You can close the terminal, reboot, or come back three days later and
    -Off will still know exactly what to undo.

    What it touches:

      APPS      Closes the processes named in the config, politely
                (CloseMainWindow, the same thing clicking the X does) so
                nothing loses unsaved work. Apps that refuse to close are
                reported, not killed, unless you pass -Force.

      SITES     Adds 0.0.0.0 entries to the hosts file between BEGIN/END
                FOCUSMODE markers, so removal is exact and the rest of your
                hosts file is never touched - it is read and written back in
                whatever encoding and line ending it already used, so any
                non-ASCII content in it survives unchanged. The file is backed
                up first. Needs an elevated shell; without one this step is
                skipped with a warning and the rest still runs.

      TOASTS    Sets the per-user ToastEnabled flag to 0, which stops
                notification banners. The previous value is saved and put
                back. This is not Windows' own Focus Assist - that has no
                supported scripting interface - but it has the same effect on
                the thing that actually interrupts you.

    First run writes a config file and tells you where it is. Edit that,
    not the script.

.PARAMETER On
    Enter focus mode.

.PARAMETER Off
    Leave focus mode and restore everything recorded in the state file.

.PARAMETER Status
    Report what is currently in effect. Changes nothing.

.PARAMETER Minutes
    With -On, stay in focus mode for this long, then restore automatically.
    The countdown runs in the foreground; Ctrl+C restores immediately. If the
    window is killed outright, run -Off later and it still restores correctly.

.PARAMETER Force
    Terminate apps that ignore a polite close request. Off by default,
    because that is how unsaved work gets lost.

.PARAMETER ConfigPath
    Config file location.
    Defaults to $env:LOCALAPPDATA\FocusMode\config.json.

.PARAMETER DryRun
    Show what would happen, touch nothing.

.EXAMPLE
    .\Set-FocusMode.ps1 -On -DryRun

.EXAMPLE
    .\Set-FocusMode.ps1 -On -Minutes 50

.EXAMPLE
    .\Set-FocusMode.ps1 -Off

.EXAMPLE
    .\Set-FocusMode.ps1 -Status
#>
[CmdletBinding(DefaultParameterSetName = 'Status')]
param(
    [Parameter(ParameterSetName = 'On', Mandatory = $true)]
    [switch]$On,

    [Parameter(ParameterSetName = 'Off', Mandatory = $true)]
    [switch]$Off,

    [Parameter(ParameterSetName = 'Status')]
    [switch]$Status,

    [Parameter(ParameterSetName = 'On')]
    [ValidateRange(1, 1440)]
    [int]$Minutes = 0,

    [Parameter(ParameterSetName = 'On')]
    [switch]$Force,

    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$ConfigPath = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'FocusMode\config.json'),
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
# A -ConfigPath with no directory part - 'config.json' - makes Split-Path
# return an empty string, and Test-Path on an empty string throws.
if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path (Get-Location).Path $ConfigPath
}

$WorkDir   = Split-Path -Parent $ConfigPath
$StateFile = Join-Path $WorkDir 'focus-state.json'
$BackupDir = Join-Path $WorkDir 'hosts-backups'
$HostsFile = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
$BeginMark = '# BEGIN FOCUSMODE - managed by Set-FocusMode.ps1, do not edit inside this block'
$EndMark   = '# END FOCUSMODE'
$ToastKey  = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\PushNotifications'

foreach ($d in @($WorkDir, $BackupDir)) {
    if (-not (Test-Path -LiteralPath $d)) {
        try { New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null }
        catch { throw "Cannot create '$d' - pick a -ConfigPath you can write to.`n$($_.Exception.Message)" }
    }
}

function Write-Action {
    param([string]$Message, [string]$Colour = '')
    $prefix = '  '
    if ($DryRun) { $prefix = '  [dry-run] ' }
    if ($Colour) { Write-Host ($prefix + $Message) -ForegroundColor $Colour }
    else         { Write-Host ($prefix + $Message) }
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Test-Elevated {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

# -------------------------------------------------------------------- config
$DefaultConfig = [pscustomobject]@{
    apps = @(
        'slack', 'discord', 'steam', 'Spotify', 'WhatsApp', 'Telegram',
        'Teams', 'ms-teams', 'EpicGamesLauncher'
    )
    blockedSites = @(
        'reddit.com', 'x.com', 'twitter.com', 'youtube.com',
        'instagram.com', 'facebook.com', 'tiktok.com', 'news.ycombinator.com'
    )
    blockWww          = $true
    muteNotifications = $true
    restoreApps       = $false
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    try {
        $DefaultConfig | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ConfigPath -Encoding utf8 -ErrorAction Stop
        Write-Host ''
        Write-Host 'First run - wrote a default config. Edit it to match what actually distracts you:' -ForegroundColor Cyan
        Write-Host "  $ConfigPath"
        Write-Host ''
    }
    catch {
        throw "Could not write the config file '$ConfigPath'.`n$($_.Exception.Message)"
    }
}

$Config = $null
try {
    $Config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
}
catch {
    throw "Config file is not valid JSON: $ConfigPath`n$($_.Exception.Message)"
}

$CfgApps    = @(Get-Prop $Config 'apps' @())
$CfgSites   = @(Get-Prop $Config 'blockedSites' @())
$CfgWww     = [bool](Get-Prop $Config 'blockWww' $true)
$CfgMute    = [bool](Get-Prop $Config 'muteNotifications' $true)
$CfgRestore = [bool](Get-Prop $Config 'restoreApps' $false)

# --------------------------------------------------------------------- state
function Read-State {
    if (-not (Test-Path -LiteralPath $StateFile)) { return $null }
    try { return Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json }
    catch { return $null }
}

function Write-State {
    param($State)
    if ($DryRun) { return }
    $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $StateFile -Encoding utf8
}

function Clear-State {
    if ($DryRun) { return }
    if (Test-Path -LiteralPath $StateFile) { Remove-Item -LiteralPath $StateFile -Force }
}

# ------------------------------------------------------------------ hosts io
function Get-HostsFile {
    <#
        Reads the hosts file as text while remembering exactly how it was
        encoded, so the untouched part can be written back byte for byte.

        This used to read and rewrite with -Encoding ascii, which silently
        replaces every non-ASCII character with a literal '?' - permanently, in
        a system file people do hand-edit, and for lines this script never had
        any business touching. Decoding and re-encoding through the same code
        page round-trips those bytes unchanged instead.

        The line ending is remembered for the same reason: a hosts file saved
        with LF endings should not come back as CRLF.
    #>
    if (-not (Test-Path -LiteralPath $HostsFile)) { return $null }

    $bytes = [System.IO.File]::ReadAllBytes($HostsFile)

    $encoding = $null
    $hasBom   = $false
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encoding = New-Object System.Text.UTF8Encoding($true); $hasBom = $true
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encoding = New-Object System.Text.UnicodeEncoding($false, $true); $hasBom = $true
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encoding = New-Object System.Text.UnicodeEncoding($true, $true); $hasBom = $true
    }
    else {
        # No BOM: Windows itself reads the hosts file through the ANSI code
        # page, and a round trip through that same page is byte-preserving.
        $encoding = [System.Text.Encoding]::Default
    }

    $text = $encoding.GetString($bytes)
    if ($hasBom) { $text = $text.TrimStart([char]0xFEFF) }

    $newline = "`r`n"
    if ($text -notmatch "`r`n" -and $text -match "`n") { $newline = "`n" }

    [pscustomobject]@{
        Lines    = @($text -split "`r?`n")
        Encoding = $encoding
        Newline  = $newline
    }
}

function Set-HostsFile {
    param([string[]]$Lines, $Encoding, [string]$Newline)
    if ($DryRun) { return }
    $text = ($Lines -join $Newline)
    # A hosts file that does not end in a newline confuses some resolvers and
    # every text editor; ours always does.
    if ($text -and -not $text.EndsWith($Newline)) { $text += $Newline }
    [System.IO.File]::WriteAllText($HostsFile, $text, $Encoding)
}

function Get-HostsBlockLines {
    $file = Get-HostsFile
    if ($null -eq $file) { return @() }
    $inside = $false
    $block  = New-Object System.Collections.Generic.List[string]
    foreach ($l in $file.Lines) {
        if ($l -eq $BeginMark) { $inside = $true;  continue }
        if ($l -eq $EndMark)   { $inside = $false; continue }
        if ($inside)           { $block.Add($l) }
    }
    $block.ToArray()
}

function Test-HostsBlockPresent {
    # Presence of the marker, not of content: a block someone emptied by hand
    # still leaves the markers behind, and -Off has to clean those up too.
    $file = Get-HostsFile
    if ($null -eq $file) { return $false }
    return ($file.Lines -contains $BeginMark)
}

function Remove-HostsBlock {
    <#
        Strips everything between the markers and leaves the rest byte-for-byte
        as it was. Never rewrites lines it did not add.
    #>
    $file = Get-HostsFile
    if ($null -eq $file) { return $false }
    if ($file.Lines -notcontains $BeginMark) { return $false }

    $kept   = New-Object System.Collections.Generic.List[string]
    $inside = $false
    foreach ($l in $file.Lines) {
        if ($l -eq $BeginMark) { $inside = $true;  continue }
        if ($l -eq $EndMark)   { $inside = $false; continue }
        if (-not $inside)      { $kept.Add($l) }
    }

    # Trim the trailing blank line the block leaves behind.
    while ($kept.Count -gt 0 -and [string]::IsNullOrWhiteSpace($kept[$kept.Count - 1])) {
        $kept.RemoveAt($kept.Count - 1)
    }

    Set-HostsFile -Lines $kept.ToArray() -Encoding $file.Encoding -Newline $file.Newline
    $true
}

function Add-HostsBlock {
    param([string[]]$Domains)

    if ($Domains.Count -eq 0) { return $false }
    if (-not (Test-Path -LiteralPath $HostsFile)) {
        Write-Action 'no hosts file on this machine - skipping site blocking' 'Yellow'
        return $false
    }

    # Back up before the first edit, and keep the last 20.
    if (-not $DryRun) {
        $backup = Join-Path $BackupDir ('hosts_{0}' -f (Get-Date -Format 'yyyy-MM-dd_HHmmss'))
        Copy-Item -LiteralPath $HostsFile -Destination $backup -Force
        Get-ChildItem -LiteralPath $BackupDir -Filter 'hosts_*' -File |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip 20 |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Remove-HostsBlock | Out-Null

    $file = Get-HostsFile
    if ($null -eq $file) { return $false }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($l in $file.Lines) { $lines.Add($l) }
    # Drop trailing blanks so exactly one separator line goes in, whether or not
    # the file already ended with a newline.
    while ($lines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) {
        $lines.RemoveAt($lines.Count - 1)
    }

    $lines.Add('')
    $lines.Add($BeginMark)
    foreach ($d in $Domains) {
        $lines.Add(('0.0.0.0 {0}' -f $d))
        $lines.Add((':: {0}' -f $d))
    }
    $lines.Add($EndMark)

    Set-HostsFile -Lines $lines.ToArray() -Encoding $file.Encoding -Newline $file.Newline
    $true
}

function Get-BlockDomains {
    $domains = New-Object System.Collections.Generic.List[string]
    foreach ($s in $CfgSites) {
        $clean = ([string]$s).Trim().ToLowerInvariant() -replace '^https?://', '' -replace '/.*$', ''
        if (-not $clean) { continue }
        if (-not $domains.Contains($clean)) { $domains.Add($clean) }
        if ($CfgWww -and -not $clean.StartsWith('www.')) {
            $w = 'www.' + $clean
            if (-not $domains.Contains($w)) { $domains.Add($w) }
        }
    }
    $domains.ToArray()
}

function Invoke-FlushDns {
    if ($DryRun) { return }
    try { & ipconfig.exe /flushdns | Out-Null } catch { }
}

# ---------------------------------------------------------------- toast io
function Get-ToastEnabled {
    if (-not (Test-Path -LiteralPath $ToastKey)) { return $null }
    try {
        $v = Get-ItemProperty -LiteralPath $ToastKey -Name 'ToastEnabled' -ErrorAction Stop
        return [int](Get-Prop $v 'ToastEnabled' 1)
    }
    catch { return $null }
}

function Set-ToastEnabled {
    param([int]$Value)
    if ($DryRun) { return }
    if (-not (Test-Path -LiteralPath $ToastKey)) {
        New-Item -Path $ToastKey -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $ToastKey -Name 'ToastEnabled' -Value $Value -PropertyType DWord -Force | Out-Null
}

# ------------------------------------------------------------------- status
function Show-Status {
    $state    = Read-State
    $block    = @(Get-HostsBlockLines)
    $toast    = Get-ToastEnabled
    $running  = @()
    foreach ($a in $CfgApps) {
        $p = @(Get-Process -Name $a -ErrorAction SilentlyContinue)
        if ($p.Count -gt 0) { $running += ('{0} ({1})' -f $a, $p.Count) }
    }

    Write-Host ''
    Write-Host 'Focus mode status'
    Write-Host '-----------------'

    if ($state) {
        $since = [string](Get-Prop $state 'startedAt' 'unknown')
        Write-Host "  Focus mode  : ON since $since" -ForegroundColor Green
        $ends = Get-Prop $state 'endsAt'
        if ($ends) { Write-Host "  Scheduled end: $ends" }
    }
    else {
        Write-Host '  Focus mode  : off'
    }

    # Distinct hostnames, not lines: every domain contributes both an IPv4 and
    # an IPv6 line, so counting lines reported double what was blocked.
    $blockedHosts = @(
        $block |
            Where-Object { $_ -match '^\S+\s+(\S+)' } |
            ForEach-Object { ($_ -split '\s+')[1] } |
            Sort-Object -Unique
    )
    $blockText = 'none'
    if ($blockedHosts.Count -gt 0) { $blockText = "$($blockedHosts.Count) hostname(s) blocked" }
    Write-Host "  Hosts block : $blockText"

    $toastText = 'unset (notifications on)'
    if ($null -ne $toast) {
        if ($toast -eq 0) { $toastText = 'suppressed' } else { $toastText = 'on' }
    }
    Write-Host "  Toasts      : $toastText"

    Write-Host "  Elevated    : $(if (Test-Elevated) { 'yes' } else { 'no - site blocking unavailable' })"
    Write-Host ''
    Write-Host "  Config      : $ConfigPath"
    Write-Host "  Apps watched: $($CfgApps.Count)   Sites listed: $($CfgSites.Count)"

    if ($running.Count -gt 0) {
        Write-Host ''
        Write-Host '  Distracting apps running right now:'
        foreach ($r in $running) { Write-Host "    $r" }
    }
    Write-Host ''
}

# ---------------------------------------------------------------------- OFF
function Disable-FocusMode {
    Write-Host ''
    Write-Host '=== Leaving focus mode ==='

    $state = Read-State
    if (-not $state) {
        Write-Host '  No active focus session recorded.' -ForegroundColor Yellow
        Write-Host '  Cleaning up anything left behind anyway...'
    }

    # --- hosts ---
    # Marker presence, not line count: a block someone emptied by hand still
    # leaves its markers in the file, and those have to go too.
    if (Test-HostsBlockPresent) {
        if (Test-Elevated) {
            if (Remove-HostsBlock) {
                Write-Action 'removed the hosts block' 'Green'
                Invoke-FlushDns
            }
        }
        else {
            Write-Action 'hosts block is still in place - re-run this from an elevated shell to remove it' 'Yellow'
        }
    }
    else {
        Write-Action 'no hosts block to remove'
    }

    # --- toasts ---
    if ($state) {
        $prev = Get-Prop $state 'toastPrevious'
        if ($null -ne $prev) {
            Set-ToastEnabled ([int]$prev)
            Write-Action ("restored notifications (ToastEnabled = {0})" -f $prev) 'Green'
        }
        elseif ((Get-Prop $state 'toastChanged' $false)) {
            # There was no value before us; put it back to the default.
            Set-ToastEnabled 1
            Write-Action 'restored notifications (ToastEnabled = 1)' 'Green'
        }
    }
    else {
        # No state file means we have no idea whether we suppressed toasts or
        # whether the user did. Turning them on would be a guess, and guessing
        # wrong silently un-mutes someone who wanted them off. Report instead.
        $current = Get-ToastEnabled
        if ($current -eq 0) {
            Write-Action 'notifications are suppressed, but no focus session recorded this' 'Yellow'
            Write-Action 'leaving that setting alone - turn it back on yourself if it was us' 'DarkGray'
        }
    }

    # --- apps ---
    if ($state -and $CfgRestore) {
        $closed = @(Get-Prop $state 'appsClosed' @())
        foreach ($c in $closed) {
            $name = [string](Get-Prop $c 'name' '')
            $exe  = [string](Get-Prop $c 'path' '')
            $label = $name
            if (-not $label) { $label = $exe }

            # Do not start a second copy of something reopened by hand since.
            if ($name) {
                $running = @(Get-Process -Name $name -ErrorAction SilentlyContinue)
                if ($running.Count -gt 0) {
                    Write-Action ("{0} is already running again - left alone" -f $label)
                    continue
                }
            }
            if (-not $exe) {
                Write-Action ("no executable path was recorded for {0} - start it yourself" -f $label) 'Yellow'
                continue
            }
            if (-not (Test-Path -LiteralPath $exe)) {
                Write-Action ("{0} is no longer at {1} - start it yourself" -f $label, $exe) 'Yellow'
                continue
            }
            if ($DryRun) { Write-Action ("would relaunch {0}" -f $label); continue }

            try {
                # Started in its own folder: plenty of apps look for resources
                # relative to the executable and misbehave when launched from
                # whatever directory this shell happens to be sitting in.
                Start-Process -FilePath $exe -WorkingDirectory (Split-Path -Parent $exe) -ErrorAction Stop
                Write-Action ("relaunched {0}" -f $label)
            }
            catch {
                # A Store app is launched by package identity through the
                # shell; its executable under WindowsApps generally refuses to
                # start directly, so say that rather than print an access error.
                $hint = ''
                if ($exe -like '*\WindowsApps\*') { $hint = ' (Store app - reopen it from the Start menu)' }
                Write-Action ("could not relaunch {0}{1}: {2}" -f $label, $hint, $_.Exception.Message) 'Yellow'
            }
        }
    }

    Clear-State
    Write-Host ''
    Write-Host '  Focus mode off.' -ForegroundColor Green
    Write-Host ''
}

# ----------------------------------------------------------------------- ON
function Enable-FocusMode {
    Write-Host ''
    Write-Host '=== Entering focus mode ==='
    if ($DryRun) { Write-Host '  (dry run - nothing will actually change)' -ForegroundColor Cyan }

    $existing = Read-State
    if ($existing) {
        Write-Host ''
        Write-Host "  Focus mode is already on (since $([string](Get-Prop $existing 'startedAt' '?')))." -ForegroundColor Yellow
        Write-Host '  Run with -Off first if you want to restart it.' -ForegroundColor Yellow
        Write-Host ''
        return $false
    }

    $state = [pscustomobject]@{
        startedAt     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        endsAt        = $null
        hostsModified = $false
        toastChanged  = $false
        toastPrevious = $null
        appsClosed    = @()
    }
    if ($Minutes -gt 0) {
        $state.endsAt = (Get-Date).AddMinutes($Minutes).ToString('yyyy-MM-dd HH:mm:ss')
    }

    # The .DESCRIPTION promises everything is recorded BEFORE it is changed, so
    # -Off can always restore. The state file used to be written once, at the
    # very end: a failure between closing apps and getting there left them
    # closed with nothing on disk saying so. It is now written up front and
    # again after every change.
    Write-State $state

    # --- apps -----------------------------------------------------------------
    Write-Host ''
    Write-Host '  Apps'
    $closedList = New-Object System.Collections.Generic.List[psobject]
    $stubborn   = New-Object System.Collections.Generic.List[string]

    foreach ($appName in $CfgApps) {
        $procs = @(Get-Process -Name $appName -ErrorAction SilentlyContinue)
        if ($procs.Count -eq 0) { continue }

        # Which of several processes is the app? The one with a window, and
        # failing that the oldest - for a multi-process app like Discord that
        # is the parent, not one of its helper processes. $procs[0] was
        # whichever the process table happened to list first.
        $best     = $null
        $bestTime = $null
        foreach ($p in $procs) {
            $handle = 0
            try { $handle = [int64]$p.MainWindowHandle } catch { }
            if ($handle -ne 0) { $best = $p; break }

            $started = $null
            try { $started = $p.StartTime } catch { }
            if ($started -and ($null -eq $bestTime -or $started -lt $bestTime)) {
                $bestTime = $started
                $best     = $p
            }
        }
        if (-not $best) { $best = $procs[0] }

        $exePath = ''
        try { $exePath = [string]$best.Path } catch { }

        if ($DryRun) {
            Write-Action ("would close {0} ({1} process(es))" -f $appName, $procs.Count)
            continue
        }

        # Ask nicely first - this is the same signal as clicking the X, so
        # anything with unsaved work gets the chance to prompt.
        foreach ($p in $procs) {
            try { $null = $p.CloseMainWindow() } catch { }
        }
        Start-Sleep -Milliseconds 1500

        $still = @(Get-Process -Name $appName -ErrorAction SilentlyContinue)
        if ($still.Count -gt 0 -and $Force) {
            foreach ($p in $still) {
                try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch { }
            }
            Start-Sleep -Milliseconds 500
            $still = @(Get-Process -Name $appName -ErrorAction SilentlyContinue)
        }

        if ($still.Count -eq 0) {
            Write-Action ("closed {0}" -f $appName) 'Green'
            $closedList.Add([pscustomobject]@{ name = $appName; path = $exePath })
            # Recorded as it happens, not at the end of the run: an app that is
            # already closed has to be in the state file before anything else
            # gets the chance to fail.
            $state.appsClosed = @($closedList)
            Write-State $state
        }
        else {
            $stubborn.Add($appName)
            Write-Action ("{0} would not close (it may have unsaved work)" -f $appName) 'Yellow'
        }
    }

    if ($closedList.Count -eq 0 -and $stubborn.Count -eq 0 -and -not $DryRun) {
        Write-Action 'none of the configured apps were running'
    }
    if ($stubborn.Count -gt 0 -and -not $Force) {
        Write-Action 'add -Force to terminate apps that ignore a close request' 'DarkGray'
    }
    $state.appsClosed = @($closedList)

    # --- sites ----------------------------------------------------------------
    Write-Host ''
    Write-Host '  Sites'
    $domains = @(Get-BlockDomains)
    if ($domains.Count -eq 0) {
        Write-Action 'no sites configured'
    }
    elseif (-not (Test-Elevated)) {
        Write-Action 'not elevated - skipping site blocking (the hosts file needs admin)' 'Yellow'
        Write-Action 'run from an elevated PowerShell to block sites too' 'DarkGray'
    }
    else {
        # Recorded before the edit: if the write half-succeeds, -Off still
        # knows to go looking for a block to remove.
        $state.hostsModified = $true
        Write-State $state
        if (Add-HostsBlock -Domains $domains) {
            Invoke-FlushDns
            Write-Action ("blocked {0} hostname(s), DNS cache flushed" -f $domains.Count) 'Green'
        }
        else {
            $state.hostsModified = $false
            Write-State $state
        }
    }

    # --- toasts ---------------------------------------------------------------
    Write-Host ''
    Write-Host '  Notifications'
    if ($CfgMute) {
        $prev = Get-ToastEnabled
        $state.toastPrevious = $prev
        if ($prev -eq 0) {
            # Already off, by the user's own choice or a policy. Recorded as the
            # previous value so -Off puts back the 0 rather than turning
            # notifications on for someone who wanted them off.
            Write-State $state
            Write-Action 'notifications were already suppressed'
        }
        else {
            # The flag is recorded before it is changed, not after.
            $state.toastChanged = $true
            Write-State $state
            Set-ToastEnabled 0
            Write-Action 'toast notifications suppressed' 'Green'
        }
    }
    else {
        Write-Action 'muteNotifications is false in the config - left alone'
    }

    Write-State $state

    Write-Host ''
    if ($DryRun) {
        Write-Host '  DRY RUN - nothing was changed. Re-run without -DryRun to focus for real.' -ForegroundColor Cyan
    }
    else {
        Write-Host '  Focus mode on.' -ForegroundColor Green
        if ($state.endsAt) { Write-Host "  Ends at $($state.endsAt)." }
    }
    Write-Host ''
    $true
}

# ---------------------------------------------------------------------- main
switch ($PSCmdlet.ParameterSetName) {

    'Off' {
        Disable-FocusMode
        exit 0
    }

    'On' {
        # Last value only: anything inside Enable-FocusMode that leaked to the
        # pipeline would otherwise make $started an array, and 'not an array of
        # two things' is false whatever those things are.
        $started = @(Enable-FocusMode) | Select-Object -Last 1
        if (-not $started) { exit 1 }

        if ($Minutes -gt 0 -and -not $DryRun) {
            $endTime = (Get-Date).AddMinutes($Minutes)
            Write-Host "  Counting down. Ctrl+C restores immediately." -ForegroundColor Cyan
            Write-Host ''
            try {
                while ((Get-Date) -lt $endTime) {
                    $left = $endTime - (Get-Date)
                    $pct  = 100 - (100.0 * $left.TotalMinutes / $Minutes)
                    Write-Progress -Activity 'Focus mode' `
                                   -Status ('{0:D2}:{1:D2} remaining' -f [int]$left.TotalMinutes, $left.Seconds) `
                                   -PercentComplete ([math]::Max(0, [math]::Min(100, $pct)))
                    Start-Sleep -Seconds 1
                }
                Write-Progress -Activity 'Focus mode' -Completed
                Write-Host '  Time is up.' -ForegroundColor Cyan
            }
            finally {
                # Runs on Ctrl+C too. If the window is killed outright this is
                # skipped, which is exactly why the state file exists.
                Write-Progress -Activity 'Focus mode' -Completed
                Disable-FocusMode
            }
        }
        exit 0
    }

    default {
        Show-Status
        exit 0
    }
}
