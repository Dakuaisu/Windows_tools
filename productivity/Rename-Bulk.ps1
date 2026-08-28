<#
.SYNOPSIS
    Batch rename by regex or by template, with a mandatory preview, collision
    detection, and an undo log so any run can be reversed.

.DESCRIPTION
    Preview is the default. The script shows you exactly what it would do and
    changes nothing until you add -Apply. That inversion is deliberate: a bad
    bulk rename is tedious to unpick by hand, and the cost of typing one extra
    switch is nothing next to the cost of getting it wrong.

    Safety properties:
      * Nothing is renamed until every new name has been validated - invalid
        characters, reserved device names (CON, LPT1, ...), empty names, and
        names that would collide with each other or with a file already on
        disk are all caught before the first rename happens. It is all-or-
        nothing, not best-effort.
      * Swaps and rotations (a->b, b->a) work. So do case-only renames on NTFS,
        which a naive Rename-Item refuses. Both are handled by renaming through
        a temporary name where needed.
      * Every applied run writes an undo log. -UndoLast reverses the most
        recent run; -UndoFrom reverses a specific one.

    Two ways to build the new name:

      REGEX     -Pattern / -Replacement, .NET regex against the file name.
                Capture groups are available as $1, $2 or ${name}.

      TEMPLATE  -Template, with tokens substituted:
                  {name}    original name without extension
                  {ext}     extension, including the dot
                  {n}       sequence number, see -StartAt / -Pad
                  {parent}  name of the containing folder
                  {date}    last write date, yyyy-MM-dd
                  {time}    last write time, HHmmss
                Tokens are case-insensitive.

.PARAMETER Path
    Folder to work in. Defaults to the current directory.

.PARAMETER Filter
    Wildcard filter for which items to consider. Default * (everything).

.PARAMETER Pattern
    Regex matched against the file name (including extension).

.PARAMETER Replacement
    Regex replacement string. Use '' to delete what the pattern matched.

.PARAMETER Template
    Name template using the tokens listed above.

.PARAMETER Recurse
    Descend into subfolders.

.PARAMETER IncludeDirectories
    Rename folders as well as files.

.PARAMETER CaseSensitive
    Make -Pattern case-sensitive. Regex is case-insensitive by default.

.PARAMETER SortBy
    Ordering used to assign {n}. Name (default), Date, or Size.

.PARAMETER StartAt
    First sequence number for {n}. Default 1.

.PARAMETER Pad
    Zero-pad {n} to this width. Default 3, so 001, 002, ...

.PARAMETER Apply
    Actually perform the renames. Without it the script only previews.

.PARAMETER UndoLast
    Reverse the most recent applied run and exit.

.PARAMETER UndoFrom
    Reverse the run recorded in the given undo CSV and exit.

.PARAMETER LogDir
    Where undo logs are kept. Defaults to $env:LOCALAPPDATA\RenameBulk.

.EXAMPLE
    # Preview: strip a prefix from every file here
    .\Rename-Bulk.ps1 -Pattern '^IMG_' -Replacement ''

.EXAMPLE
    # Same thing, for real
    .\Rename-Bulk.ps1 -Pattern '^IMG_' -Replacement '' -Apply

.EXAMPLE
    # Number every photo by date taken order: Holiday-001.jpg, Holiday-002.jpg
    .\Rename-Bulk.ps1 -Filter *.jpg -Template 'Holiday-{n}{ext}' -SortBy Date -Apply

.EXAMPLE
    # Move the date to the front: "report 2026-01-04.pdf" -> "2026-01-04 report.pdf"
    .\Rename-Bulk.ps1 -Pattern '^(.+?) (\d{4}-\d{2}-\d{2})\.pdf$' -Replacement '$2 $1.pdf'

.EXAMPLE
    .\Rename-Bulk.ps1 -UndoLast
#>
[CmdletBinding(DefaultParameterSetName = 'Regex')]
param(
    [Parameter(ParameterSetName = 'Regex', Position = 0)]
    [Parameter(ParameterSetName = 'Template', Position = 0)]
    [string]$Path = (Get-Location).Path,

    [Parameter(ParameterSetName = 'Regex')]
    [Parameter(ParameterSetName = 'Template')]
    [string]$Filter = '*',

    [Parameter(ParameterSetName = 'Regex', Mandatory = $true)]
    [string]$Pattern,

    [Parameter(ParameterSetName = 'Regex')]
    [string]$Replacement = '',

    [Parameter(ParameterSetName = 'Template', Mandatory = $true)]
    [string]$Template,

    [Parameter(ParameterSetName = 'Regex')]
    [Parameter(ParameterSetName = 'Template')]
    [switch]$Recurse,

    [Parameter(ParameterSetName = 'Regex')]
    [Parameter(ParameterSetName = 'Template')]
    [switch]$IncludeDirectories,

    [Parameter(ParameterSetName = 'Regex')]
    [switch]$CaseSensitive,

    [Parameter(ParameterSetName = 'Regex')]
    [Parameter(ParameterSetName = 'Template')]
    [ValidateSet('Name','Date','Size')]
    [string]$SortBy = 'Name',

    [Parameter(ParameterSetName = 'Template')]
    [int]$StartAt = 1,

    [Parameter(ParameterSetName = 'Template')]
    [ValidateRange(1, 10)]
    [int]$Pad = 3,

    [Parameter(ParameterSetName = 'Regex')]
    [Parameter(ParameterSetName = 'Template')]
    [switch]$Apply,

    [Parameter(ParameterSetName = 'UndoLast', Mandatory = $true)]
    [switch]$UndoLast,

    [Parameter(ParameterSetName = 'UndoFrom', Mandatory = $true)]
    [string]$UndoFrom,

    [string]$LogDir = (Join-Path $env:LOCALAPPDATA 'RenameBulk')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
function Initialize-LogDir {
    <#
        An applied run is only reversible if its undo log actually gets
        written, so this folder is proved writable before anything is renamed
        rather than after. A redirected or locked-down LOCALAPPDATA falls back
        to TEMP with a warning; if even that fails the run is refused, because
        renaming without an undo log is exactly the outcome this tool exists
        to prevent.
    #>
    param([string]$Preferred)

    $fallback = ''
    if ($env:TEMP) { $fallback = Join-Path $env:TEMP 'RenameBulk' }

    foreach ($candidate in @($Preferred, $fallback)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $candidate)) {
                New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path $candidate ('.write-probe-{0}.tmp' -f $PID)
            [System.IO.File]::WriteAllText($probe, 'x')
            [System.IO.File]::Delete($probe)
            if ($candidate -ne $Preferred) {
                Write-Host "Cannot write to '$Preferred' - keeping undo logs in '$candidate' instead." -ForegroundColor Yellow
            }
            return $candidate
        }
        catch { }
    }
    throw "No writable folder for undo logs (tried '$Preferred' and '$fallback'). Refusing to rename without one."
}

$LogDir = Initialize-LogDir -Preferred $LogDir

$RunStart      = Get-Date
$InvalidChars  = [System.IO.Path]::GetInvalidFileNameChars()
$ReservedNames = @(
    'CON','PRN','AUX','NUL',
    'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
    'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
)

function Test-ValidFileName {
    <#
        Returns an empty string when the name is usable, otherwise the reason
        it is not. Windows rejects more names than people expect.
    #>
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name))              { return 'name would be empty' }
    if ($Name.IndexOfAny($InvalidChars) -ge 0)            { return 'contains an illegal character' }
    if ($Name.EndsWith('.') -or $Name.EndsWith(' '))      { return 'ends with a dot or space' }
    if ($Name.Length -gt 255)                             { return 'longer than 255 characters' }

    # Windows resolves 'CON', 'CON.txt' and 'CON   .txt' to the same device, so
    # the stem is trimmed before the comparison.
    $stem = $Name
    $dot  = $Name.IndexOf('.')
    if ($dot -gt 0) { $stem = $Name.Substring(0, $dot) }
    $stem = $stem.Trim()
    if ($ReservedNames -contains $stem.ToUpperInvariant()) { return "'$stem' is a reserved device name" }

    ''
}

function Expand-Token {
    <#
        Case-insensitive literal token substitution. Deliberately not -replace:
        the replacement side of a regex treats '$' as a capture reference, and
        file names are full of characters that would mean something there.
    #>
    param([string]$Text, [string]$Token, [string]$Value)

    $sb  = New-Object System.Text.StringBuilder
    $pos = 0
    while ($true) {
        $idx = $Text.IndexOf($Token, $pos, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -lt 0) { break }
        [void]$sb.Append($Text.Substring($pos, $idx - $pos))
        [void]$sb.Append($Value)
        $pos = $idx + $Token.Length
    }
    [void]$sb.Append($Text.Substring($pos))
    $sb.ToString()
}

function Get-TempName {
    param([string]$Directory, [string]$Seed)
    for ($i = 0; $i -lt 10000; $i++) {
        $candidate = '.renamebulk-{0}-{1}.tmp' -f $Seed, $i
        if (-not (Test-Path -LiteralPath (Join-Path $Directory $candidate))) { return $candidate }
    }
    throw "Could not find a free temporary name in $Directory"
}

function Get-PathDepth {
    # Separator count, not string length: 'C:\aaaaaaaaaaaa' is shallower than
    # 'C:\a\b\c' even though it is the longer string.
    param([string]$PathText)
    if ([string]::IsNullOrEmpty($PathText)) { return 0 }
    @($PathText.TrimEnd('\').Split('\')).Count
}

function Invoke-RenamePlan {
    <#
        Applies a validated plan of Directory/From/To rows.

        Rows are bucketed by their containing directory and the buckets are
        walked deepest directory first. Renaming a subfolder is recorded
        against its *parent* directory, so depth ordering puts a folder's
        contents strictly before the folder itself: no rename can ever
        invalidate a path another row still has to use. -ShallowestFirst walks
        the same sequence backwards, which is what an undo needs - a renamed
        parent is put back to the name its children were recorded under before
        those children are touched.

        Within one bucket, any row whose target name collides with another
        row's source name, or which differs only by case, is renamed via a
        temporary name first. Without that, swaps fail and NTFS silently
        refuses case-only renames.

        -BestEffort is for undo, where the world may have moved on since the
        log was written: rows whose source has vanished, or whose target name
        has since been taken by something this run is not moving, are skipped
        and reported rather than failing the whole reversal. Nothing is ever
        overwritten in either mode.
    #>
    param([psobject[]]$Plan, [switch]$ShallowestFirst, [switch]$BestEffort)

    $done    = New-Object System.Collections.Generic.List[psobject]
    $skipped = New-Object System.Collections.Generic.List[psobject]

    $buckets = @{}
    foreach ($p in $Plan) {
        $key = ([string]$p.Directory).ToLowerInvariant()
        if (-not $buckets.ContainsKey($key)) {
            $buckets[$key] = New-Object System.Collections.Generic.List[psobject]
        }
        $buckets[$key].Add($p)
    }

    $keys = @($buckets.Keys | Sort-Object @{ Expression = { Get-PathDepth $_ } }, @{ Expression = { $_ } })
    if (-not $ShallowestFirst) { [array]::Reverse($keys) }

    foreach ($key in $keys) {
        $group = @($buckets[$key])

        if ($BestEffort) {
            # Checked here, not up front: after an ancestor bucket has been
            # processed the paths in this bucket are valid again, so an early
            # sweep would wrongly declare every nested row missing.
            $usable = New-Object System.Collections.Generic.List[psobject]
            foreach ($p in $group) {
                if (Test-Path -LiteralPath (Join-Path $p.Directory $p.From)) { $usable.Add($p) }
                else { $skipped.Add([pscustomobject]@{ Item = $p; Reason = 'no longer where the log says it is' }) }
            }
            $group = @($usable)
        }
        if ($group.Count -eq 0) { continue }

        $sources = @{}
        foreach ($p in $group) { $sources[(Join-Path $p.Directory $p.From).ToLowerInvariant()] = $true }

        if ($BestEffort) {
            $free = New-Object System.Collections.Generic.List[psobject]
            foreach ($p in $group) {
                $targetPath = Join-Path $p.Directory $p.To
                $targetLc   = $targetPath.ToLowerInvariant()
                $sourceLc   = (Join-Path $p.Directory $p.From).ToLowerInvariant()
                if ($targetLc -ne $sourceLc -and (Test-Path -LiteralPath $targetPath) -and -not $sources.ContainsKey($targetLc)) {
                    $skipped.Add([pscustomobject]@{ Item = $p; Reason = "'$($p.To)' already exists again" })
                }
                else { $free.Add($p) }
            }
            $group = @($free)
            if ($group.Count -eq 0) { continue }

            $sources = @{}
            foreach ($p in $group) { $sources[(Join-Path $p.Directory $p.From).ToLowerInvariant()] = $true }
        }

        $viaTemp = New-Object System.Collections.Generic.List[psobject]
        $direct  = New-Object System.Collections.Generic.List[psobject]

        for ($i = 0; $i -lt $group.Count; $i++) {
            $p        = $group[$i]
            $targetLc = (Join-Path $p.Directory $p.To).ToLowerInvariant()
            $sourceLc = (Join-Path $p.Directory $p.From).ToLowerInvariant()

            $caseOnly = ($p.From -ieq $p.To) -and ($p.From -cne $p.To)
            $cycle    = ($targetLc -ne $sourceLc) -and $sources.ContainsKey($targetLc)

            if ($caseOnly -or $cycle) { $viaTemp.Add($p) } else { $direct.Add($p) }
        }

        # Phase 1: park the risky ones under temporary names.
        $parked = New-Object System.Collections.Generic.List[psobject]
        for ($i = 0; $i -lt $viaTemp.Count; $i++) {
            $p    = $viaTemp[$i]
            $temp = Get-TempName -Directory $p.Directory -Seed $i
            Rename-Item -LiteralPath (Join-Path $p.Directory $p.From) -NewName $temp -Force -ErrorAction Stop
            $parked.Add([pscustomobject]@{ Directory = $p.Directory; Temp = $temp; To = $p.To; From = $p.From })
        }

        # Phase 2: everything lands on its final name.
        foreach ($p in $direct) {
            Rename-Item -LiteralPath (Join-Path $p.Directory $p.From) -NewName $p.To -Force -ErrorAction Stop
            $done.Add([pscustomobject]@{ Directory = $p.Directory; From = $p.From; To = $p.To })
        }
        foreach ($p in $parked) {
            Rename-Item -LiteralPath (Join-Path $p.Directory $p.Temp) -NewName $p.To -Force -ErrorAction Stop
            $done.Add([pscustomobject]@{ Directory = $p.Directory; From = $p.From; To = $p.To })
        }
    }

    [pscustomobject]@{ Done = $done; Skipped = $skipped }
}

# ---------------------------------------------------------------- undo modes
if ($PSCmdlet.ParameterSetName -eq 'UndoLast' -or $PSCmdlet.ParameterSetName -eq 'UndoFrom') {

    $logFile = $UndoFrom
    if ($PSCmdlet.ParameterSetName -eq 'UndoLast') {
        $latest = Get-ChildItem -LiteralPath $LogDir -Filter 'undo_*.csv' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '*.reversed.csv' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latest) {
            Write-Host 'No undo logs found - nothing to reverse.' -ForegroundColor Yellow
            exit 0
        }
        $logFile = $latest.FullName
    }

    if (-not (Test-Path -LiteralPath $logFile)) { throw "Undo log not found: $logFile" }

    $rows = @(Import-Csv -LiteralPath $logFile)
    if ($rows.Count -eq 0) {
        Write-Host "That undo log is empty - nothing to reverse: $logFile" -ForegroundColor Yellow
        exit 0
    }
    # Strict mode turns a missing CSV column into an error rather than $null,
    # so check the shape before touching a single property.
    foreach ($col in @('Directory','From','To')) {
        if (-not $rows[0].PSObject.Properties[$col]) {
            throw "Not a Rename-Bulk undo log - no '$col' column: $logFile"
        }
    }

    Write-Host ''
    Write-Host "Reversing $($rows.Count) rename(s) from:"
    Write-Host "  $logFile"
    Write-Host ''

    # The reverse plan swaps From and To; -ShallowestFirst then replays the run
    # backwards, so a renamed parent folder is restored to the name its
    # children were recorded under before those children are touched.
    $plan = @(
        $rows | ForEach-Object {
            [pscustomobject]@{
                Directory = [string]$_.Directory
                From      = [string]$_.To
                To        = [string]$_.From
            }
        }
    )

    $result   = Invoke-RenamePlan -Plan $plan -ShallowestFirst -BestEffort
    $reverted = @($result.Done)
    $skipped  = @($result.Skipped)

    foreach ($r in $reverted) {
        Write-Host ('  {0}  ->  {1}' -f $r.From, $r.To)
    }

    if ($skipped.Count -gt 0) {
        Write-Host ''
        Write-Host "  $($skipped.Count) item(s) were skipped:" -ForegroundColor Yellow
        foreach ($s in $skipped) {
            Write-Host ("    {0} : {1}" -f $s.Item.From, $s.Reason) -ForegroundColor Yellow
        }
    }

    Write-Host ''
    Write-Host "Reversed $($reverted.Count) rename(s)." -ForegroundColor Green

    # A reversed log is spent; park it so -UndoLast does not re-run it. The
    # '.reversed.csv' name still matches the undo_*.csv filter above, which is
    # why that filter excludes it explicitly.
    if ($reverted.Count -gt 0) {
        Rename-Item -LiteralPath $logFile -NewName ((Split-Path -Leaf $logFile) -replace '\.csv$', '.reversed.csv') -Force
    }
    # A partial reversal still exits 0: exit 1 means "validation refused the
    # run", and quietly overloading it here would break every caller that
    # already relies on that meaning. Skips are reported on the console.
    exit 0
}

# ------------------------------------------------------------- gather items
if (-not (Test-Path -LiteralPath $Path)) { throw "Path not found: $Path" }
$Path = (Resolve-Path -LiteralPath $Path).ProviderPath

$gatherArgs = @{
    LiteralPath = $Path
    Filter      = $Filter
    Force       = $true
    ErrorAction = 'SilentlyContinue'
}
if ($Recurse) { $gatherArgs['Recurse'] = $true }
if (-not $IncludeDirectories) { $gatherArgs['File'] = $true }

$items = @(Get-ChildItem @gatherArgs | Where-Object { $_.Name -ne 'desktop.ini' })

# This ordering only decides which item gets which {n}. It deliberately says
# nothing about the order the renames happen in - Invoke-RenamePlan derives
# that from directory depth, so a folder is always renamed after its contents.
switch ($SortBy) {
    'Date' { $items = @($items | Sort-Object LastWriteTime) }
    'Size' { $items = @($items | Sort-Object @{ Expression = { if ($_.PSIsContainer) { 0 } else { $_.Length } } }) }
    default { $items = @($items | Sort-Object Name) }
}

if ($items.Count -eq 0) {
    Write-Host "No items matched '$Filter' in $Path" -ForegroundColor Yellow
    exit 0
}

# --------------------------------------------------------------- build plan
$regexOptions = [System.Text.RegularExpressions.RegexOptions]::None
if (-not $CaseSensitive) {
    $regexOptions = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
}

$plan     = New-Object System.Collections.Generic.List[psobject]
$errors   = New-Object System.Collections.Generic.List[string]
$unchanged = 0
$seq      = $StartAt

foreach ($item in $items) {
    $oldName = $item.Name
    $newName = $oldName

    if ($PSCmdlet.ParameterSetName -eq 'Regex') {
        try {
            if (-not [regex]::IsMatch($oldName, $Pattern, $regexOptions)) { $unchanged++; continue }
            $newName = [regex]::Replace($oldName, $Pattern, $Replacement, $regexOptions)
        }
        catch {
            throw "Bad regex: $($_.Exception.Message)"
        }
    }
    else {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($oldName)
        $ext      = [System.IO.Path]::GetExtension($oldName)
        if ($item.PSIsContainer) { $baseName = $oldName; $ext = '' }

        $parentName = Split-Path -Leaf (Split-Path -Parent $item.FullName)
        $number     = $seq.ToString().PadLeft($Pad, '0')

        $newName = $Template
        $newName = Expand-Token $newName '{name}'   $baseName
        $newName = Expand-Token $newName '{ext}'    $ext
        $newName = Expand-Token $newName '{parent}' $parentName
        $newName = Expand-Token $newName '{n}'      $number
        $newName = Expand-Token $newName '{index}'  $number
        $newName = Expand-Token $newName '{date}'   $item.LastWriteTime.ToString('yyyy-MM-dd')
        $newName = Expand-Token $newName '{time}'   $item.LastWriteTime.ToString('HHmmss')
        $seq++
    }

    if ($newName -ceq $oldName) { $unchanged++; continue }

    $reason = Test-ValidFileName $newName
    if ($reason) {
        $errors.Add("'$oldName' -> '$newName' : $reason")
        continue
    }

    $parent = Split-Path -Parent $item.FullName
    $plan.Add([pscustomobject]@{
        Directory = $parent
        From      = $oldName
        To        = $newName
        IsFolder  = [bool]$item.PSIsContainer
    })
}

# ------------------------------------------------------------ validate plan
# Two new names that are equal (case-insensitively, as NTFS sees them) would
# silently clobber each other, so that is a hard stop rather than a warning.
$byTarget = @{}
foreach ($p in $plan) {
    $key = (Join-Path $p.Directory $p.To).ToLowerInvariant()
    if ($byTarget.ContainsKey($key)) {
        $errors.Add("collision: '$($byTarget[$key])' and '$($p.From)' would both become '$($p.To)'")
    }
    else {
        $byTarget[$key] = $p.From
    }
}

# Existing files are only a conflict when nothing in this run is vacating them.
$sourceKeys = @{}
foreach ($p in $plan) { $sourceKeys[(Join-Path $p.Directory $p.From).ToLowerInvariant()] = $true }

foreach ($p in $plan) {
    $targetPath = Join-Path $p.Directory $p.To
    $targetKey  = $targetPath.ToLowerInvariant()
    $sourceKey  = (Join-Path $p.Directory $p.From).ToLowerInvariant()
    if ($targetKey -eq $sourceKey) { continue }          # case-only rename

    if ((Test-Path -LiteralPath $targetPath) -and -not $sourceKeys.ContainsKey($targetKey)) {
        $errors.Add("'$($p.From)' -> '$($p.To)' : a file with that name already exists")
    }
}

# ------------------------------------------------------------------ preview
Write-Host ''
Write-Host "Rename plan"
Write-Host "-----------"
Write-Host "  Folder    : $Path"
Write-Host "  Filter    : $Filter$(if ($Recurse) { ' (recursive)' } else { '' })"
if ($PSCmdlet.ParameterSetName -eq 'Regex') {
    Write-Host "  Pattern   : $Pattern"
    Write-Host "  Replace   : '$Replacement'"
}
else {
    Write-Host "  Template  : $Template"
}
Write-Host "  Matched   : $($items.Count) item(s), $($plan.Count) would change, $unchanged unchanged"
Write-Host ''

if ($plan.Count -gt 0) {
    $width = 0
    foreach ($p in $plan) { if ($p.From.Length -gt $width) { $width = $p.From.Length } }
    if ($width -gt 60) { $width = 60 }

    foreach ($p in $plan) {
        $from = $p.From
        if ($from.Length -gt $width) { $from = $from.Substring(0, $width - 3) + '...' }
        $marker = ' '
        if ($p.IsFolder) { $marker = 'D' }
        Write-Host ('  {0} {1}  ->  {2}' -f $marker, $from.PadRight($width), $p.To)
    }
}

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Host "PROBLEMS ($($errors.Count)) - nothing will be renamed until these are fixed:" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host "  $e" -ForegroundColor Red }
    Write-Host ''
    exit 1
}

if ($plan.Count -eq 0) {
    Write-Host 'Nothing to rename.' -ForegroundColor Yellow
    exit 0
}

if (-not $Apply) {
    Write-Host ''
    Write-Host 'PREVIEW ONLY - nothing was renamed. Add -Apply to run it for real.' -ForegroundColor Cyan
    exit 0
}

# -------------------------------------------------------------------- apply
Write-Host ''
Write-Host "Renaming $($plan.Count) item(s)..."

$undoLog = Join-Path $LogDir ('undo_{0}.csv' -f $RunStart.ToString('yyyy-MM-dd_HHmmss'))
$done    = $null

try {
    $done = @((Invoke-RenamePlan -Plan @($plan)).Done)
}
catch {
    Write-Host ''
    Write-Host "FAILED partway through: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Some items may be left under temporary .renamebulk-*.tmp names in the target folder.' -ForegroundColor Red
    Write-Host 'Check the folder before re-running.' -ForegroundColor Red
    exit 1
}

@($done | ForEach-Object {
    [pscustomobject]@{
        Timestamp = $RunStart.ToString('yyyy-MM-dd HH:mm:ss')
        Directory = $_.Directory
        From      = $_.From
        To        = $_.To
    }
}) | Export-Csv -LiteralPath $undoLog -NoTypeInformation -Encoding utf8

# Keep the 50 most recent undo logs.
Get-ChildItem -LiteralPath $LogDir -Filter 'undo_*.csv' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 50 |
    Remove-Item -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host "Renamed $($done.Count) item(s)." -ForegroundColor Green
Write-Host "Undo log : $undoLog"
Write-Host 'Reverse it with:  .\Rename-Bulk.ps1 -UndoLast'
exit 0
