<#
.SYNOPSIS
    Tidy My Downloads - the engine.

.DESCRIPTION
    Sorts settled files in the Downloads folder into <Category>\<yyyy-MM>\
    subfolders, and REPORTS (never deletes) old clutter.

    Safety rules this file must always uphold:
      * Nothing is ever deleted. There is no purge switch and never will be.
      * Move-Item appears in exactly ONE function, and that function has no
        -Force parameter, so an existing file can never be overwritten.
      * Every move is written to the journal BEFORE it happens, so an
        interrupted run can still be undone completely.
      * If the journal cannot be written, the run stops before moving anything.
      * File CONTENTS are never read. No hashing, no sniffing. This is what
        stops OneDrive from downloading gigabytes of cloud-only files.
      * Preview is the default. Moving requires -Apply.

.PARAMETER Path
    Folder to tidy. Defaults to the real Downloads folder, resolved from the
    Windows known-folder registry value (so it follows a OneDrive redirect).

.PARAMETER GraceDays
    Files younger than this are never touched. Default 30. This is what makes
    "I downloaded it Monday and can't find it Friday" impossible: the Downloads
    folder itself is always the last 30 days.

.PARAMETER StaleDays
    Report clutter untouched for this many days. Default 90. Report only.

.PARAMETER Apply
    Actually move files. Without it, this is a preview and nothing changes.

.EXAMPLE
    .\DownloadsJanitor.Core.ps1
    Preview only - shows what would happen, changes nothing.

.EXAMPLE
    .\DownloadsJanitor.Core.ps1 -Apply
#>
[CmdletBinding()]
param(
    [string]$Path,
    [int]   $GraceDays = 30,
    [int]   $StaleDays = 90,
    [switch]$Apply,
    [switch]$Quiet,
    [string]$LogDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LogFile     = $null
$script:JournalFile = $null
$script:LockFile    = $null
$script:HaveLock    = $false
$script:ReportName  = 'Where are my files.html'
$script:NoteName    = '_what-is-this.txt'

# ===========================================================================
# Output helpers
# ===========================================================================

function Write-Line {
    param([string]$Text = '', [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

# Write-through logging: every line hits the disk immediately. If the process
# is killed mid-run, the log still shows exactly how far it got.
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    if (-not $script:LogFile) { return }
    $stamp = Get-Date
    $line = '{0:D4}-{1:D2}-{2:D2} {3:D2}:{4:D2}:{5:D2} [{6}] {7}' -f `
        $stamp.Year, $stamp.Month, $stamp.Day, $stamp.Hour, $stamp.Minute, $stamp.Second, $Level, $Message
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8 -ErrorAction Stop } catch { }
}

function Write-Both {
    param([string]$Message, [string]$Level = 'INFO', [string]$Colour = '')
    Write-Log $Message $Level
    Write-Line $Message $Colour
}

# Measure-Object -Property over an EMPTY pipeline returns $null, and .Sum on
# $null throws under StrictMode. That crash used to kill the script after it
# had already moved files. A plain loop cannot do that.
function Get-SafeSum {
    param($Items, [string]$Property)
    $total = [long]0
    foreach ($i in $Items) {
        if ($null -ne $i) { $total += [long]$i.$Property }
    }
    return $total
}

function Format-Size {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { '{0:N1} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N1} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { '{0:N0} KB' -f ($Bytes / 1KB) }
    else                    { "$Bytes bytes" }
}

# .ToString('yyyy-MM') uses the CURRENT CULTURE'S CALENDAR. On a Thai machine
# that yields 2569-08, on an Arabic one 1448-03. Building the string from the
# integer parts is the only culture-proof way to name a folder.
function Get-MonthBucket {
    param([datetime]$When)
    return ('{0:D4}-{1:D2}' -f $When.Year, $When.Month)
}

function Get-Stamp {
    param([datetime]$When)
    return ('{0:D4}-{1:D2}-{2:D2}_{3:D2}{4:D2}{5:D2}' -f `
        $When.Year, $When.Month, $When.Day, $When.Hour, $When.Minute, $When.Second)
}

function Get-DateOnly {
    param([datetime]$When)
    return ('{0:D4}-{1:D2}-{2:D2}' -f $When.Year, $When.Month, $When.Day)
}

# ===========================================================================
# Locating and vetting the target folder
# ===========================================================================

# %USERPROFILE%\Downloads is a guess. The known-folder registry value is the
# truth, and it follows a OneDrive "back up your folders" redirect.
function Get-DownloadsPath {
    $key  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
    $guid = '{374DE290-123F-4565-9164-39C4925E467B}'
    foreach ($k in @($key, 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')) {
        try {
            $v = Get-ItemPropertyValue -Path $k -Name $guid -ErrorAction Stop
            if ($v) {
                $v = [Environment]::ExpandEnvironmentVariables($v)
                if (Test-Path -LiteralPath $v -PathType Container) { return $v }
            }
        } catch { }
    }
    return (Join-Path $env:USERPROFILE 'Downloads')
}

# Refuse to "tidy" anywhere that isn't a downloads-like folder. Turning this
# loose on a profile root or Documents would be a catastrophe.
function Test-UnsafeTarget {
    param([string]$Target)

    $t = $Target.TrimEnd('\')

    # Drive root (C:\) or UNC share root (\\server\share)
    if ($t -match '^[A-Za-z]:$')            { return 'that is the root of a drive' }
    if ($t -match '^\\\\[^\\]+\\[^\\]+$')   { return 'that is the root of a network share' }

    $forbidden = @()
    foreach ($n in @('UserProfile','Desktop','MyDocuments','MyPictures','MyMusic','MyVideos',
                     'Favorites','ApplicationData','LocalApplicationData','CommonApplicationData',
                     'Windows','ProgramFiles','ProgramFilesX86','System')) {
        try {
            $p = [Environment]::GetFolderPath($n)
            if ($p) { $forbidden += $p.TrimEnd('\') }
        } catch { }
    }
    foreach ($e in @($env:PUBLIC, $env:OneDrive, $env:OneDriveCommercial, $env:ProgramData)) {
        if ($e) { $forbidden += $e.TrimEnd('\') }
    }

    foreach ($f in $forbidden) {
        if ($f -and $t -ieq $f) { return "that is a protected Windows folder ($f)" }
    }
    return $null
}

function Resolve-LogDirectory {
    param([string]$Preferred)
    $candidates = @()
    if ($Preferred) { $candidates += $Preferred }
    if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA 'DownloadsJanitor') }
    if ($env:TEMP)         { $candidates += (Join-Path $env:TEMP 'DownloadsJanitor') }

    foreach ($c in $candidates) {
        try {
            if (-not (Test-Path -LiteralPath $c)) {
                New-Item -ItemType Directory -Path $c -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path $c ('.w' + $PID)
            Set-Content -LiteralPath $probe -Value 'x' -ErrorAction Stop
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            return $c
        } catch { }
    }
    return $null
}

# ===========================================================================
# Classification
# ===========================================================================

function Get-Category {
    param([string]$Name, [string]$Extension, $Config)

    # ToLowerInvariant, never ToLower: on a Turkish machine ToLower turns the
    # I in .AI into a dotless i, and the extension stops matching the map.
    $ext = $Extension.ToLowerInvariant()

    if ($Config.Map.ContainsKey($ext)) { return $Config.Map[$ext] }

    # US tax software emits .tax, .tax2024, .tax2029 - a family, not a list.
    # Anchored so that .taxonomy does not match.
    if ($ext -match '^\.tax\d*$') { return 'Accounting' }

    # A sidecar rides with its parent: contract.pdf.p7s is a signature for a
    # document, not a mystery file.
    $lower = $Name.ToLowerInvariant()
    if ($lower -match '^(.+)(\.[a-z0-9]{1,10})\.[a-z0-9]{1,10}$') {
        $parentExt = $Matches[2]
        if ($Config.Map.ContainsKey($parentExt)) { return $Config.Map[$parentExt] }
    }

    return 'Everything Else'
}

function Test-NeverMove {
    param($File, $Config)

    $name  = $File.Name
    $lower = $name.ToLowerInvariant()

    if ($File.Attributes -band [System.IO.FileAttributes]::System) { return 'system file' }
    if ($Config.SkipNames -contains $lower)      { return 'not ours to move' }
    if ($lower.StartsWith('~$'))                 { return 'Office lock file' }
    if ($Config.NeverMoveExt -contains $File.Extension.ToLowerInvariant()) { return 'shortcut or live data file' }

    # Partial downloads are matched on how the NAME ENDS, because clients use
    # both "file.part" and "big.iso.aria2" shapes.
    foreach ($s in $Config.SkipSuffixes) {
        if ($lower.EndsWith($s)) { return 'still downloading' }
    }
    return $null
}

# ===========================================================================
# Moving - the only place in this tool that relocates a file
# ===========================================================================

function Get-NonClashingPath {
    param([string]$Directory, [string]$FileName, [datetime]$FileDate)

    $candidate = Join-Path $Directory $FileName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $dot  = $FileName.LastIndexOf('.')
    if ($dot -gt 0) {
        $base = $FileName.Substring(0, $dot)
        $ext  = $FileName.Substring($dot)
    } else {
        $base = $FileName
        $ext  = ''
    }

    # A date reads better than a number: "statement (2026-08-14).pdf" tells you
    # which statement it is.
    $candidate = Join-Path $Directory ('{0} ({1}){2}' -f $base, (Get-DateOnly $FileDate), $ext)
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $Directory ('{0} ({1}){2}' -f $base, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    return $null
}

# NOTE: no -Force parameter exists here, and Move-Item is called without it,
# so this function is structurally incapable of overwriting a file.
function Move-OneFile {
    param([string]$Source, [string]$Destination)

    $destDir = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null
    }
    try {
        Move-Item -LiteralPath $Source -Destination $Destination -ErrorAction Stop
    }
    catch [System.InvalidOperationException] {
        # Names like "CON.txt" or "report.pdf." are legal on disk but the Win32
        # layer refuses them unless the path is prefixed. Local paths only -
        # a UNC path needs a different prefix form, so let it throw instead.
        if ($Source -match '^[A-Za-z]:\\' -and $Destination -match '^[A-Za-z]:\\') {
            [System.IO.File]::Move(('\\?\' + $Source), ('\\?\' + $Destination))
        } else { throw }
    }
}

# ===========================================================================
# Reporting
# ===========================================================================

function Get-NewestTimestamp {
    param($Item)
    $t = $Item.CreationTime
    if ($Item.LastWriteTime  -gt $t) { $t = $Item.LastWriteTime }
    if ($Item.LastAccessTime -gt $t) { $t = $Item.LastAccessTime }
    return $t
}

function ConvertTo-HtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}


# ===========================================================================
# HTML report - written into the Downloads folder itself, because that is the
# one place the person will definitely look.
# ===========================================================================

function Write-HtmlReport {
    param(
        [string]$Target, [array]$Moves, [array]$Stale, [array]$LeftAlone,
        [datetime]$When, [int]$GraceDays, [int]$StaleDays, [array]$ExemptCategories
    )

    $h = @()
    $h += '<!doctype html><html><head><meta charset="utf-8">'
    $h += '<title>Where are my files</title>'
    $h += '<style>'
    $h += 'body{font-family:Segoe UI,system-ui,sans-serif;max-width:900px;margin:40px auto;padding:0 20px;color:#1a1a1a;background:#fff;line-height:1.6}'
    $h += 'h1{font-size:26px;margin-bottom:4px} h2{font-size:18px;margin-top:36px;border-bottom:1px solid #e5e5e5;padding-bottom:6px}'
    $h += '.sub{color:#666;margin-top:0} .box{background:#f6f8fa;border:1px solid #e5e5e5;border-radius:8px;padding:14px 18px;margin:18px 0}'
    $h += 'table{border-collapse:collapse;width:100%;font-size:14px} td,th{text-align:left;padding:6px 10px;border-bottom:1px solid #eee;vertical-align:top}'
    $h += 'th{background:#fafafa;font-weight:600} td.r{text-align:right;white-space:nowrap;color:#666}'
    $h += '#f{width:100%;padding:10px 12px;font-size:15px;border:1px solid #ccc;border-radius:6px;margin:10px 0}'
    $h += 'a{color:#0b5fff;text-decoration:none} a:hover{text-decoration:underline}'
    $h += '.none{color:#666;font-style:italic} .ok{color:#137333;font-weight:600}'
    $h += '@media(prefers-color-scheme:dark){body{background:#161616;color:#e8e8e8}.box{background:#1f1f1f;border-color:#333}'
    $h += 'th{background:#1f1f1f}td,th{border-color:#2a2a2a}#f{background:#1f1f1f;color:#e8e8e8;border-color:#444}a{color:#6fa8ff}}'
    $h += '</style></head><body>'

    $h += ('<h1>Where are my files</h1>')
    $h += ('<p class="sub">Tidied on {0}</p>' -f (ConvertTo-HtmlText $When.ToString('D')))

    $movedBytes = Get-SafeSum $Moves 'Size'
    $h += '<div class="box">'
    $h += ('<p><span class="ok">Nothing was deleted.</span> {0} file(s) were moved into folders inside <b>{1}</b>, a total of {2}.</p>' -f `
            $Moves.Count, (ConvertTo-HtmlText $Target), (Format-Size $movedBytes))
    $h += ('<p>Only files older than {0} days are ever moved, so anything you downloaded recently is still sitting in Downloads where you left it.</p>' -f $GraceDays)
    $h += '<p>Changed your mind? Double-click <b>Put Everything Back</b> in the Tidy My Downloads folder and every file returns to exactly where it was.</p>'
    $h += '</div>'

    if ($Moves.Count -gt 0) {
        $h += '<h2>What moved where</h2>'
        $h += '<p>Type below to find a file - for example <b>invoice</b>, or <b>.pdf</b>.</p>'
        $h += '<input id="f" placeholder="Search your moved files..." oninput="flt()">'
        $h += '<table id="t"><tr><th>File</th><th>Now lives in</th><th class="r">Size</th></tr>'
        foreach ($m in $Moves) {
            $link = 'file:///' + ($m.Destination -replace '\\', '/' -replace ' ', '%20')
            $h += ('<tr><td><a href="{0}">{1}</a></td><td>{2}</td><td class="r">{3}</td></tr>' -f `
                    (ConvertTo-HtmlText $link), (ConvertTo-HtmlText $m.Name),
                    (ConvertTo-HtmlText $m.Folder), (Format-Size $m.Size))
        }
        $h += '</table>'
    }

    if ($LeftAlone.Count -gt 0) {
        $h += '<h2>Left exactly where they were</h2>'
        $h += '<table><tr><th>File</th><th>Why</th></tr>'
        foreach ($l in $LeftAlone) {
            $h += ('<tr><td>{0}</td><td>{1}</td></tr>' -f (ConvertTo-HtmlText $l.Name), (ConvertTo-HtmlText $l.Reason))
        }
        $h += '</table>'
    }

    $h += ('<h2>Old clutter you might not need</h2>')
    if ($Stale.Count -gt 0) {
        $staleBytes = Get-SafeSum $Stale 'Size'
        $h += ('<div class="box"><p>These have not changed in over {0} days and take up <b>{1}</b>.' -f $StaleDays, (Format-Size $staleBytes))
        $h += ' <b>Nothing here has been deleted or moved for you</b> - this is only a list. Delete anything you recognise as junk, and ignore the rest.</p>'
        $h += ('<p style="color:#666;font-size:14px">Windows does not record when you last <i>opened</i> a file, only when it last <i>changed</i>. So this list means "unchanged", not "unused". Your {0} are never listed here.</p></div>' -f `
                (ConvertTo-HtmlText (($ExemptCategories | Sort-Object) -join ', ')))
        $h += '<table><tr><th>Item</th><th>Where</th><th class="r">Size</th><th class="r">Age</th></tr>'
        foreach ($s in $Stale) {
            $h += ('<tr><td>{0}</td><td>{1}</td><td class="r">{2}</td><td class="r">{3} days</td></tr>' -f `
                    (ConvertTo-HtmlText $s.Name), (ConvertTo-HtmlText $s.Where),
                    (Format-Size $s.Size), $s.AgeDays)
        }
        $h += '</table>'
    } else {
        $h += '<p class="none">Nothing old enough to bother you about. Your Downloads folder is in good shape.</p>'
    }

    $h += '<h2>Can''t find something?</h2>'
    $h += '<p>Open your Downloads folder and type the file name into the search box at the top right of the window. Windows searches inside all the new folders too.</p>'

    $h += '<script>function flt(){var q=document.getElementById("f").value.toLowerCase();'
    $h += 'var rows=document.getElementById("t").rows;for(var i=1;i<rows.length;i++){'
    $h += 'rows[i].style.display=rows[i].innerText.toLowerCase().indexOf(q)>-1?"":"none";}}</script>'
    $h += '</body></html>'

    $out = Join-Path $Target $script:ReportName
    Set-Content -LiteralPath $out -Value ($h -join "`r`n") -Encoding utf8 -ErrorAction Stop
    return $out
}

$script:CategoryNotes = @{
    'Camera Raw'        = 'Photos straight out of a camera (RAW files), plus the small settings files that record your edits. Keep each pair together.'
    'Brushes & Presets' = 'Brushes, swatches, filters, LUTs and other add-ons for creative apps.'
    'Accounting'        = 'Files that only your accounting software can open - bank exports, ledgers, backups.'
    'Certificates'      = 'Security certificates and keys. Treat these as private and do not share them.'
    'Everything Else'   = 'Files whose type this tool did not recognise. Nothing is wrong with them.'
}

# ===========================================================================
# MAIN
# ===========================================================================

$RunStart = Get-Date
$exitCode = 0

try {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }

    $catFile = Join-Path $here 'DownloadsJanitor.Categories.ps1'
    if (-not (Test-Path -LiteralPath $catFile)) {
        Write-Line ''
        Write-Line "  Some of this tool's files are missing." 'Red'
        Write-Line "  Expected to find: $catFile"
        Write-Line '  Please extract the whole "Tidy My Downloads" folder again and retry.'
        Write-Line ''
        exit 2
    }
    $Config = & $catFile

    # Every folder this tool has ever created, so its own output is never
    # mistaken for the user's junk.
    $managed = @(@($Config.Map.Values) + @($Config.LegacyFolders) | Sort-Object -Unique)

    # ---------------------------------------------------------- target folder
    if (-not $PSBoundParameters.ContainsKey('Path') -or -not $Path) {
        $Path = Get-DownloadsPath
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        Write-Line ''
        Write-Line '  Could not find your Downloads folder.' 'Red'
        Write-Line "  Looked in: $Path"
        Write-Line '  If your Downloads folder lives somewhere unusual, nothing has been changed.'
        Write-Line ''
        exit 2
    }
    $Path = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')

    $unsafe = Test-UnsafeTarget $Path
    if ($unsafe) {
        Write-Line ''
        Write-Line '  Stopping without changing anything.' 'Yellow'
        Write-Line "  $Path"
        Write-Line "  $unsafe, and this tool only ever tidies a Downloads folder."
        Write-Line ''
        exit 2
    }

    # ------------------------------------------------------------ log folder
    $resolvedLogDir = Resolve-LogDirectory $LogDir
    if ($resolvedLogDir) {
        $script:LogFile = Join-Path $resolvedLogDir 'janitor.log'
    }

    Write-Log "=== start (apply=$Apply) target=$Path grace=${GraceDays}d stale=${StaleDays}d ==="

    # --------------------------------------------------------------- preflight
    $onedrive = $false
    if ($env:OneDrive -and $Path.ToLowerInvariant().StartsWith($env:OneDrive.ToLowerInvariant())) { $onedrive = $true }

    Write-Line ''
    Write-Line '  ---------------------------------------------------------'
    Write-Line '   TIDY MY DOWNLOADS'
    Write-Line '  ---------------------------------------------------------'
    Write-Line ''
    Write-Line "   Folder    : $Path"
    if ($onedrive) {
    Write-Line '   OneDrive  : yes - files stay in the cloud, nothing is downloaded'
    }
    Write-Line "   Keeping   : anything from the last $GraceDays days stays put"
    if ($Apply) {
    Write-Line '   Mode      : MOVING FILES' 'Yellow'
    } else {
    Write-Line '   Mode      : preview only, nothing will be changed' 'Cyan'
    }
    if (-not $resolvedLogDir) {
    Write-Line '   Note      : could not open a log folder, continuing without a log'
    }
    Write-Line ''

    # ------------------------------------------------------------------ lock
    if ($Apply -and $resolvedLogDir) {
        $script:LockFile = Join-Path $resolvedLogDir 'running.lock'
        if (Test-Path -LiteralPath $script:LockFile) {
            $stale = $true
            try {
                $info = Get-Content -LiteralPath $script:LockFile -ErrorAction Stop
                $age  = ($RunStart - (Get-Item -LiteralPath $script:LockFile).LastWriteTime).TotalHours
                if ($age -lt 3) {
                    $otherPid = 0
                    if ([int]::TryParse(($info | Select-Object -First 1), [ref]$otherPid)) {
                        if (Get-Process -Id $otherPid -ErrorAction SilentlyContinue) { $stale = $false }
                    }
                }
            } catch { }
            if (-not $stale) {
                Write-Line '   Tidy My Downloads is already running. Nothing to do.' 'Yellow'
                Write-Line ''
                Write-Log 'another instance holds the lock; exiting'
                exit 0
            }
            Remove-Item -LiteralPath $script:LockFile -Force -ErrorAction SilentlyContinue
        }
        try {
            Set-Content -LiteralPath $script:LockFile -Value "$PID" -ErrorAction Stop
            $script:HaveLock = $true
        } catch { }
    }

    # --------------------------------------------------------------- journal
    # Invariant: a move that cannot be journalled does not happen.
    if ($Apply) {
        if (-not $resolvedLogDir) {
            Write-Line '  Cannot write the undo record, so nothing will be moved.' 'Red'
            Write-Line '  (Without it, "Put Everything Back" could not work.)'
            Write-Line ''
            exit 2
        }
        $script:JournalFile = Join-Path $resolvedLogDir ('moves_' + (Get-Stamp $RunStart) + '.tsv')
        try {
            Set-Content -LiteralPath $script:JournalFile -Value "# Tidy My Downloads move journal`tsource`tdestination`tbytes" -Encoding utf8 -ErrorAction Stop
        } catch {
            Write-Line '  Cannot write the undo record, so nothing will be moved.' 'Red'
            Write-Line ''
            exit 2
        }
    }

    # ==================================================== PHASE 1 : sort files
    $graceCutoff = $RunStart.AddDays(-$GraceDays)
    $moves     = @()
    $leftAlone = @()
    $skipped   = 0

    $loose = @(Get-ChildItem -LiteralPath $Path -File -Force -ErrorAction SilentlyContinue)

    foreach ($file in $loose) {
        $name = $file.Name

        $why = Test-NeverMove $file $Config
        if ($why) { $skipped++; Write-Log "skip ($why): $name"; continue }

        # Arrival time: a file extracted from a zip can carry a 2019 timestamp
        # while having landed seconds ago. Taking the later of the two stops it
        # being filed under 2019 and stops the grace period being bypassed.
        $arrived = $file.CreationTime
        if ($file.LastWriteTime -gt $arrived) { $arrived = $file.LastWriteTime }

        if ($arrived -gt $graceCutoff) { $skipped++; Write-Log "skip (recent): $name"; continue }

        $category = Get-Category $name $file.Extension $Config
        $bucket   = Get-MonthBucket $arrived
        $destDir  = Join-Path (Join-Path $Path $category) $bucket
        $destPath = Get-NonClashingPath -Directory $destDir -FileName $name -FileDate $arrived

        if (-not $destPath) {
            $leftAlone += @{ Name = $name; Reason = 'too many files with this name already' }
            Write-Log "left alone (no free name): $name" 'WARN'
            continue
        }

        # Long paths still fail on most machines. Skipping beats truncating.
        if ($destPath.Length -ge 250) {
            $leftAlone += @{ Name = $name; Reason = 'the new folder path would be too long for Windows' }
            Write-Log "left alone (path too long): $name" 'WARN'
            continue
        }

        if (-not (Test-Path -LiteralPath $file.FullName)) {
            $skipped++; Write-Log "skip (vanished before move): $name"; continue
        }

        if ($Apply) {
            try {
                # Journal FIRST. An interrupted run stays fully undoable.
                Add-Content -LiteralPath $script:JournalFile -Encoding utf8 -ErrorAction Stop `
                    -Value ("{0}`t{1}`t{2}" -f $file.FullName, $destPath, $file.Length)
            } catch {
                $leftAlone += @{ Name = $name; Reason = 'could not record the undo step, so it was not moved' }
                Write-Log "left alone (journal write failed): $name" 'ERROR'
                continue
            }
            try {
                Move-OneFile -Source $file.FullName -Destination $destPath
            } catch {
                $leftAlone += @{ Name = $name; Reason = 'it was open in another program' }
                Write-Log ("left alone (move failed): {0} :: {1}" -f $name, $_.Exception.Message) 'WARN'
                continue
            }
        }

        $moves += @{
            Name        = $name
            Destination = $destPath
            Folder      = (Join-Path $category $bucket)
            Size        = $file.Length
        }
        Write-Log ('move: {0} -> {1}' -f $name, $destPath)
        $leaf = Split-Path -Leaf $destPath
        if ($leaf -ne $name) { Write-Log ("renamed to avoid a clash: '{0}' -> '{1}'" -f $name, $leaf) 'WARN' }
    }

    # ================================================== PHASE 2 : stale report
    $staleCutoff = $RunStart.AddDays(-$StaleDays)
    $stale = @()

    $unmanagedDirs = @(Get-ChildItem -LiteralPath $Path -Directory -Force -ErrorAction SilentlyContinue |
                       Where-Object { $managed -notcontains $_.Name })
    $unmanagedRoots = @()
    foreach ($d in $unmanagedDirs) { $unmanagedRoots += ($d.FullName + '\') }

    foreach ($f in @(Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue)) {
        $lower = $f.Name.ToLowerInvariant()
        if ($Config.SkipNames -contains $lower) { continue }
        if ($lower.StartsWith('~$')) { continue }

        # Files inside an unmanaged folder are summarised by that folder's own
        # row instead, so one busy torrent folder cannot flood the report.
        $inUnmanaged = $false
        foreach ($u in $unmanagedRoots) {
            if ($f.FullName.StartsWith($u, [System.StringComparison]::OrdinalIgnoreCase)) { $inUnmanaged = $true; break }
        }
        if ($inUnmanaged) { continue }

        # Paperwork is exempt: an old contract is doing its job.
        $rel = $f.FullName.Substring($Path.Length).TrimStart('\')
        $topFolder = ''
        if ($rel.Contains('\')) { $topFolder = $rel.Substring(0, $rel.IndexOf('\')) }
        if ($Config.RecordsCategories -contains $topFolder) { continue }

        $ts = Get-NewestTimestamp $f
        if ($ts -lt $staleCutoff) {
            $where = Split-Path -Parent $f.FullName
            if ($where.Length -gt $Path.Length) { $where = $where.Substring($Path.Length).TrimStart('\') } else { $where = '(loose in Downloads)' }
            $stale += @{
                Name    = $f.Name
                Where   = $where
                Size    = $f.Length
                AgeDays = [int]($RunStart - $ts).TotalDays
            }
        }
    }

    # One roll-up row per unmanaged folder, emitted whenever it contains
    # anything old - even if the folder itself is in active use.
    foreach ($dir in $unmanagedDirs) {
        $children  = @(Get-ChildItem -LiteralPath $dir.FullName -Recurse -File -Force -ErrorAction SilentlyContinue)
        $lastTouch = Get-NewestTimestamp $dir
        $staleN    = 0
        $staleSize = [long]0
        foreach ($c in $children) {
            $ct = Get-NewestTimestamp $c
            # The age reported is time since the MOST RECENT activity anywhere in
            # the folder, not since its oldest file. Reporting the oldest would
            # make a folder look more abandoned than it is, and this list invites
            # people to delete things.
            if ($ct -gt $lastTouch) { $lastTouch = $ct }
            if ($ct -lt $staleCutoff) {
                $staleN++
                $staleSize += $c.Length
            }
        }
        if ($staleN -gt 0) {
            $label = '{0} ({1} of {2} files unchanged)' -f $dir.Name, $staleN, $children.Count
            $stale += @{
                Name    = $label
                Where   = '(folder in Downloads)'
                Size    = $staleSize
                AgeDays = [int]($RunStart - $lastTouch).TotalDays
            }
        }
    }

    $stale = @($stale | Sort-Object -Property @{ Expression = { $_.Size } } -Descending)

    # ------------------------------------------------------ notes + HTML report
    $reportPath = $null
    if ($Apply) {
        foreach ($m in $moves) {
            $cat = $m.Folder.Split('\')[0]
            if ($script:CategoryNotes.ContainsKey($cat)) {
                $noteFile = Join-Path (Join-Path $Path $cat) $script:NoteName
                if (-not (Test-Path -LiteralPath $noteFile)) {
                    try {
                        Set-Content -LiteralPath $noteFile -Encoding utf8 -ErrorAction Stop -Value @(
                            $cat, ('=' * $cat.Length), '', $script:CategoryNotes[$cat], '',
                            'Put here automatically by Tidy My Downloads. Nothing was deleted.',
                            'To undo, run "Put Everything Back" from the Tidy My Downloads folder.')
                    } catch { }
                }
            }
        }
        try {
            $reportPath = Write-HtmlReport -Target $Path -Moves $moves -Stale $stale -LeftAlone $leftAlone `
                            -When $RunStart -GraceDays $GraceDays -StaleDays $StaleDays `
                            -ExemptCategories $Config.RecordsCategories
        } catch {
            Write-Log ("could not write the HTML report: {0}" -f $_.Exception.Message) 'WARN'
        }
    }

    # ------------------------------------------------------------- summary
    $movedBytes = Get-SafeSum $moves 'Size'
    $staleBytes = Get-SafeSum $stale 'Size'

    Write-Line ''
    if ($Apply) {
        Write-Line ('   Moved      : {0} file(s), {1}' -f $moves.Count, (Format-Size $movedBytes)) 'Green'
    } else {
        Write-Line ('   Would move : {0} file(s), {1}' -f $moves.Count, (Format-Size $movedBytes)) 'Cyan'
    }

    $byCat = @{}
    foreach ($m in $moves) {
        $c = $m.Folder.Split('\')[0]
        if (-not $byCat.ContainsKey($c)) { $byCat[$c] = 0 }
        $byCat[$c]++
    }
    foreach ($c in ($byCat.Keys | Sort-Object)) {
        Write-Line ('                {0,-20} {1}' -f $c, $byCat[$c])
    }

    Write-Line ('   Kept back  : {0} file(s) too recent or still in use' -f $skipped)
    if ($leftAlone.Count -gt 0) {
        Write-Line ('   Left alone : {0} file(s) - listed in the report' -f $leftAlone.Count) 'Yellow'
    }
    Write-Line ('   Old stuff  : {0} item(s), {1} - REPORTED ONLY, nothing deleted' -f $stale.Count, (Format-Size $staleBytes))
    Write-Line ''
    if ($reportPath) {
        Write-Line '   Full details, and a search box to find any file:'
        Write-Line "   $reportPath" 'Cyan'
        Write-Line ''
    }

    Write-Log ("=== done: moved=$($moves.Count) skipped=$skipped leftAlone=$($leftAlone.Count) stale=$($stale.Count) ===")
}
catch {
    # A friend must never see a red stack trace. Log the detail, show a sentence.
    Write-Log ("UNEXPECTED: {0}" -f $_.Exception.ToString()) 'ERROR'
    Write-Line ''
    Write-Line '  Something went wrong, so this stopped early.' 'Red'
    Write-Line '  Your files have not been harmed - this tool never deletes anything,'
    Write-Line '  and anything already moved can be undone with "Put Everything Back".'
    if ($script:LogFile) {
        Write-Line ''
        Write-Line "  Technical details were saved to:"
        Write-Line "  $($script:LogFile)"
    }
    Write-Line ''
    $exitCode = 1
}
finally {
    if ($script:HaveLock -and $script:LockFile) {
        Remove-Item -LiteralPath $script:LockFile -Force -ErrorAction SilentlyContinue
    }
}

exit $exitCode
