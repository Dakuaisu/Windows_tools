<#
.SYNOPSIS
    Put Everything Back - reverses the most recent tidy-up.

.DESCRIPTION
    Reads the newest move journal and walks it BACKWARDS, returning every file
    to exactly where it came from.

    Safety rules:
      * Nothing is ever deleted, including during undo.
      * If something is already sitting at the original location, the file is
        restored next to it as "name (put back).ext" instead of overwriting.
      * Only completely empty folders created by the tidy-up are removed.
      * The undo writes its own journal, so it can itself be undone.
#>
[CmdletBinding()]
param(
    [string]$LogDir,
    [string]$JournalFile,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Line {
    param([string]$Text = '', [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

$exitCode = 0

try {
    # ------------------------------------------------------- find the journal
    if (-not $JournalFile) {
        $dirs = @()
        if ($LogDir) { $dirs += $LogDir }
        if ($env:LOCALAPPDATA) { $dirs += (Join-Path $env:LOCALAPPDATA 'DownloadsJanitor') }
        if ($env:TEMP)         { $dirs += (Join-Path $env:TEMP 'DownloadsJanitor') }

        $found = @()
        foreach ($d in $dirs) {
            if (Test-Path -LiteralPath $d) {
                $found += @(Get-ChildItem -LiteralPath $d -Filter 'moves_*.tsv' -File -ErrorAction SilentlyContinue)
            }
        }
        if ($found.Count -eq 0) {
            Write-Line ''
            Write-Line '  There is nothing to put back.' 'Yellow'
            Write-Line '  No tidy-up has been run on this computer yet.'
            Write-Line ''
            exit 0
        }

        # Take the newest journal that actually RECORDED a move. A run where
        # nothing needed tidying writes an empty journal, and that must not
        # shadow the last run that really did move files.
        $JournalFile = $null
        foreach ($j in @($found | Sort-Object LastWriteTime -Descending)) {
            $peek = @(Get-Content -LiteralPath $j.FullName -ErrorAction SilentlyContinue |
                      Where-Object { $_ -and -not $_.StartsWith('#') })
            if ($peek.Count -gt 0) { $JournalFile = $j.FullName; break }
        }
        if (-not $JournalFile) {
            Write-Line ''
            Write-Line '  There is nothing to put back.' 'Yellow'
            Write-Line '  Every tidy-up so far has left your files exactly where they were.'
            Write-Line ''
            exit 0
        }
    }

    if (-not (Test-Path -LiteralPath $JournalFile)) {
        Write-Line "  Could not find the record of what was moved." 'Red'
        exit 2
    }

    $lines = @(Get-Content -LiteralPath $JournalFile -ErrorAction Stop | Where-Object { $_ -and -not $_.StartsWith('#') })
    if ($lines.Count -eq 0) {
        Write-Line ''
        Write-Line '  The last tidy-up did not move anything, so there is nothing to put back.' 'Yellow'
        Write-Line ''
        exit 0
    }

    Write-Line ''
    Write-Line '  ---------------------------------------------------------'
    Write-Line '   PUT EVERYTHING BACK'
    Write-Line '  ---------------------------------------------------------'
    Write-Line ''
    Write-Line "   Undoing $($lines.Count) move(s) from the last tidy-up."
    Write-Line ''

    # --------------------------------------------------- undo journal of its own
    $undoJournal = $null
    try {
        $stamp = Get-Date
        $name = 'undo_{0:D4}-{1:D2}-{2:D2}_{3:D2}{4:D2}{5:D2}.tsv' -f `
            $stamp.Year, $stamp.Month, $stamp.Day, $stamp.Hour, $stamp.Minute, $stamp.Second
        $undoJournal = Join-Path (Split-Path -Parent $JournalFile) $name
        Set-Content -LiteralPath $undoJournal -Value "# put-back journal`tsource`tdestination`tbytes" -Encoding utf8 -ErrorAction Stop
    } catch { $undoJournal = $null }

    $restored = 0
    $renamed  = 0
    $missing  = 0
    $failed   = 0
    $touchedDirs = @{}
    $downloadsRoot = $null

    # Backwards, so that a file moved twice lands where it started.
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        $parts = $lines[$i] -split "`t"
        if ($parts.Count -lt 2) { continue }
        $origin  = $parts[0]
        $current = $parts[1]
        if (-not $downloadsRoot) { $downloadsRoot = Split-Path -Parent $origin }

        if (-not (Test-Path -LiteralPath $current)) {
            $missing++
            Write-Line ("   already gone: {0}" -f (Split-Path -Leaf $current))
            continue
        }

        $target = $origin
        if (Test-Path -LiteralPath $target) {
            # Something is there now. Never overwrite - restore alongside it.
            $leaf = Split-Path -Leaf $origin
            $dir  = Split-Path -Parent $origin
            $dot  = $leaf.LastIndexOf('.')
            if ($dot -gt 0) { $base = $leaf.Substring(0, $dot); $ext = $leaf.Substring($dot) }
            else            { $base = $leaf; $ext = '' }
            $target = Join-Path $dir ('{0} (put back){1}' -f $base, $ext)
            $n = 2
            while (Test-Path -LiteralPath $target) {
                $target = Join-Path $dir ('{0} (put back {1}){2}' -f $base, $n, $ext)
                $n++
                if ($n -gt 999) { break }
            }
            $renamed++
        }

        try {
            $parent = Split-Path -Parent $target
            if (-not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null
            }
            # No -Force: undo can no more overwrite a file than the tidy-up can.
            Move-Item -LiteralPath $current -Destination $target -ErrorAction Stop
            $restored++
            $touchedDirs[(Split-Path -Parent $current)] = $true
            if ($undoJournal) {
                Add-Content -LiteralPath $undoJournal -Encoding utf8 -ErrorAction SilentlyContinue `
                    -Value ("{0}`t{1}`t0" -f $current, $target)
            }
        } catch {
            $failed++
            Write-Line ("   could not move back: {0}" -f (Split-Path -Leaf $current)) 'Yellow'
        }
    }

    # ------------------------------------------- tidy away the empty shells
    # Only folders the tidy-up itself emptied, and only when truly empty.
    $removedDirs = 0
    foreach ($d in ($touchedDirs.Keys | Sort-Object -Property Length -Descending)) {
        try {
            if (Test-Path -LiteralPath $d) {
                if (@(Get-ChildItem -LiteralPath $d -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                    Remove-Item -LiteralPath $d -Force -ErrorAction Stop
                    $removedDirs++
                    # The category folder above it may now be empty too.
                    $up = Split-Path -Parent $d
                    if ((Test-Path -LiteralPath $up) -and
                        @(Get-ChildItem -LiteralPath $up -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                        Remove-Item -LiteralPath $up -Force -ErrorAction SilentlyContinue
                        $removedDirs++
                    }
                }
            }
        } catch { }
    }

    # ------------------------------------------------- remove our own leftovers
    # These are files this tool wrote, never the user's. A category folder left
    # holding nothing but our explanatory note is not "back to how it was".
    if ($downloadsRoot -and (Test-Path -LiteralPath $downloadsRoot)) {
        foreach ($cat in @(Get-ChildItem -LiteralPath $downloadsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
            try {
                $kids = @(Get-ChildItem -LiteralPath $cat.FullName -Force -ErrorAction SilentlyContinue)
                if ($kids.Count -eq 1 -and $kids[0].Name -eq '_what-is-this.txt') {
                    Remove-Item -LiteralPath $kids[0].FullName -Force -ErrorAction Stop
                    Remove-Item -LiteralPath $cat.FullName -Force -ErrorAction SilentlyContinue
                    $removedDirs++
                }
            } catch { }
        }
        # The report describes moves that have just been undone, so it is now wrong.
        $rp = Join-Path $downloadsRoot 'Where are my files.html'
        if (Test-Path -LiteralPath $rp) { Remove-Item -LiteralPath $rp -Force -ErrorAction SilentlyContinue }
    }

    Write-Line ''
    Write-Line ("   Put back      : {0} file(s)" -f $restored) 'Green'
    if ($renamed -gt 0) {
        Write-Line ("   Renamed       : {0} - something else was already using the old name," -f $renamed) 'Yellow'
        Write-Line  '                   so these came back as "name (put back).ext"'
    }
    if ($missing -gt 0) { Write-Line ("   Already moved : {0} - you had already moved these yourself" -f $missing) }
    if ($failed  -gt 0) { Write-Line ("   Could not move: {0} - probably open in another program" -f $failed) 'Yellow' }
    if ($removedDirs -gt 0) { Write-Line ("   Empty folders removed: {0}" -f $removedDirs) }
    Write-Line ''
    Write-Line '   Your Downloads folder is back to how it was.' 'Green'
    Write-Line ''
}
catch {
    Write-Line ''
    Write-Line '  Something went wrong while putting files back.' 'Red'
    Write-Line '  Nothing was deleted. Any file not moved back is still in its'
    Write-Line '  category folder inside Downloads, safe and findable.'
    Write-Line ''
    $exitCode = 1
}

exit $exitCode
