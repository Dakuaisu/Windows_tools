<#
.SYNOPSIS
    Which applications are crashing and hanging on this machine, how often, and
    whether it is getting worse - as an HTML report plus a rolling daily CSV.

.DESCRIPTION
    Read-only. Counts Application Error (1000) and Application Hang (1002)
    events over the last N days, groups them by executable, works out a friendly
    name for each one, and names the module most of an app's crashes point at.

    The count comes from 1000 and 1002 only. Windows Error Reporting's 1001
    events look like a tempting counter and are not one: on this machine 99% of
    them are BlueScreen and LiveKernelEvent records re-logged on every upload
    retry, and most application crashes never emit a 1001 at all. They are used
    here to enrich a crash that already exists - bucket ID, report folder - and
    never to increment anything.

    Alongside the per-app league table, Windows' own reliability metric
    (Win32_ReliabilityStabilityMetrics, the number behind Reliability Monitor)
    is sampled per day, so "one app is unhappy" can be told apart from "this
    whole machine is degrading".

    The Application log is circular. If it does not reach back as far as -Days
    asks, the window shrinks to the oldest record present and the report says
    so rather than quietly reporting 90 days of data it never had.

    Nothing needs administrator rights. The only thing a standard user cannot
    read is the contents of other accounts' WER report folders; those are
    counted and summarised in one line, never warned about individually.

.PARAMETER Days
    How many days back to look. Default 30. The effective window shrinks to the
    Application log's oldest record when the log does not reach back that far.

.PARAMETER Top
    How many applications the league table lists. Default 15.

.PARAMETER MaxEvents
    Upper bound on any single event-log query, so a pathological log cannot
    balloon the run. Default 5000.

.PARAMETER HistoryLimit
    Rows kept in each rolling CSV. Default 5000; 0 keeps everything. When the
    column set changes the old file is archived rather than deleted.

.PARAMETER OutputDir
    Where the report and CSVs land. Defaults to $env:LOCALAPPDATA\GetCrashTrend.
    If that cannot be written to, the tool falls back to %TEMP%\GetCrashTrend
    and says so.

.PARAMETER Now
    Treat this instant as "now". Only useful for testing the window arithmetic
    against known data; leave it alone in normal use.

.PARAMETER SkipWerStore
    Skip the Windows Error Reporting folder sweep. The report loses the
    friendly names harvested from Report.wer and the report-folder tally.

.PARAMETER Quiet
    Suppress console output; still writes the files and sets the exit code.
    One line survives -Quiet: the notice that -OutputDir could not be written
    to and the files went to %TEMP% instead, because it says where they are.

.PARAMETER Open
    Open the generated HTML report when finished.

.OUTPUTS
    Exit code 0 = no crashes or hangs in the effective window,
    1 = crashes or hangs present and the trend is flat or improving, or the
        counters could not be read at all - zero events read is not the same
        fact as zero events, so it is never reported as an all-clear,
    2 = worsening (recent half of the window much worse than the first half, a
        single app with 10 or more events, or a stability index well below its
        own window average),
    4 = no usable data (both event counters failed and the reliability class
        was unavailable),
    5 = no writable output folder.

    Note that powershell.exe -File also returns 1 when a parameter fails
    validation, which is indistinguishable from verdict 1. A launcher that acts
    on these codes should validate its arguments before invoking the script.

.EXAMPLE
    .\Get-CrashTrend.ps1 -Open

.EXAMPLE
    .\Get-CrashTrend.ps1 -Days 90 -Top 25

.EXAMPLE
    # Plot the machine's stability against its crash count
    Import-Csv "$env:LOCALAPPDATA\GetCrashTrend\StabilityDaily.csv" |
        Select-Object Date, SystemStabilityIndex, CrashCount
#>
[CmdletBinding()]
param(
    # Requested window; the effective window shrinks to the Application log's
    # oldest record and the report says so.
    [ValidateRange(7, 365)]
    [int]$Days = 30,

    [ValidateRange(1, 50)]
    [int]$Top = 15,

    # Upper bound on any single Get-WinEvent query, so a pathological log
    # cannot balloon the run.
    [ValidateRange(500, 100000)]
    [int]$MaxEvents = 5000,

    [ValidateRange(0, 100000)]
    [int]$HistoryLimit = 5000,

    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'GetCrashTrend'),

    # Test seam: freeze "now" so the window halves land on known data. DontShow
    # keeps it out of the GUI launcher's generated form and out of tab
    # completion: the only correct value is the default.
    [Parameter(DontShow)]
    [datetime]$Now = (Get-Date),

    [switch]$SkipWerStore,
    [switch]$Quiet,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
function Get-Prop {
    # Strict-mode-safe property read.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Initialize-OutputDir {
    <#
        A redirected or locked-down LOCALAPPDATA, or an -OutputDir nobody can
        write to, should not stop a diagnostic from running - it should put the
        files somewhere that works and say where.
    #>
    param([string]$Preferred, [string]$ToolName)

    $fallback = ''
    if ($env:TEMP) { $fallback = Join-Path $env:TEMP $ToolName }

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
                Write-Host "Cannot write to '$Preferred' - using '$candidate' instead." -ForegroundColor Yellow
            }
            return $candidate
        }
        catch { }
    }
    throw "No writable output folder (tried '$Preferred' and '$fallback')."
}

function Write-Line {
    param([string]$Message, [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Message -ForegroundColor $Colour }
    else         { Write-Host $Message }
}

function Write-Step {
    param([string]$Message)
    Write-Line "  $Message"
}

function Format-Invariant {
    # This machine's regional format is dd-MM-yyyy; nothing written to a file
    # or compared as a string may depend on that.
    param([datetime]$Value, [string]$Format = 'yyyy-MM-dd HH:mm:ss')
    $Value.ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-HtmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-RawHtml {
    <#
        Marks a cell whose value is already HTML and must not be escaped again.
        A typed marker cannot be produced by accident the way a string prefix
        convention can.
    #>
    param([string]$Html)
    [pscustomobject]@{ PSTypeName = 'CrashTrend.RawHtml'; Html = $Html }
}

function New-KeyValueTable {
    param([hashtable]$Pairs, [string[]]$Order)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table class="kv">')
    foreach ($k in $Order) {
        if (-not $Pairs.ContainsKey($k)) { continue }
        $v = $Pairs[$k]
        if ($null -eq $v -or [string]$v -eq '') { $v = '-' }
        if ($v.PSObject.TypeNames -contains 'CrashTrend.RawHtml') {
            [void]$sb.Append(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlText $k), [string]$v.Html))
        }
        else {
            [void]$sb.Append(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlText $k), (ConvertTo-HtmlText $v)))
        }
    }
    [void]$sb.Append('</table>')
    $sb.ToString()
}

function New-DataTable {
    <#
        Rows in, HTML table out. Columns are the property names of the first row
        unless -Columns is given. A value built by New-RawHtml is emitted
        unescaped; everything else is HTML-encoded.
    #>
    param([psobject[]]$Rows, [string[]]$Columns, [string]$Empty = 'Nothing to show.')

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return ('<p class="empty">{0}</p>' -f (ConvertTo-HtmlText $Empty))
    }
    if (-not $Columns) {
        $Columns = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="scroll"><table class="data"><thead><tr>')
    foreach ($c in $Columns) { [void]$sb.Append('<th>' + (ConvertTo-HtmlText $c) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')

    foreach ($row in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($c in $Columns) {
            $val = Get-Prop $row $c ''
            if ($null -ne $val -and $val.PSObject.TypeNames -contains 'CrashTrend.RawHtml') {
                [void]$sb.Append('<td>' + [string]$val.Html + '</td>')
            }
            else {
                [void]$sb.Append('<td>' + (ConvertTo-HtmlText ([string]$val)) + '</td>')
            }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    $sb.ToString()
}

function New-Bar {
    param([double]$Percent, [string]$Class = '')
    $p = [math]::Max(0, [math]::Min(100, $Percent))
    New-RawHtml ('<div class="bar ' + $Class + '"><span style="width:' + ([math]::Round($p, 1)) + '%"></span></div>')
}

$Sections = New-Object System.Collections.Generic.List[string]
function Add-Section {
    param([string]$Title, [string]$Body, [string]$Note = '')
    $noteHtml = ''
    if ($Note) { $noteHtml = '<p class="note">' + (ConvertTo-HtmlText $Note) + '</p>' }
    $Sections.Add('<section><h2>' + (ConvertTo-HtmlText $Title) + '</h2>' + $noteHtml + $Body + '</section>')
}

# Degradation is reported, never fatal: a missing source costs the report a
# column, not the run.
$Health = New-Object System.Collections.Generic.List[psobject]
function Add-Health {
    param([string]$Source, [ValidateSet('ok','degraded','skipped')][string]$State, [string]$Detail)
    $Health.Add([pscustomobject]@{ Source = $Source; State = $State; Detail = $Detail })
}

function Add-HistoryRow {
    <#
        Export-Csv -Append refuses outright when the file's header does not
        match the object being appended, so the first run after this script
        gains or loses a column would throw away the whole reading at the last
        step. The stale file is archived under a dated name instead, which keeps
        the earlier readings rather than discarding them.
    #>
    param([psobject[]]$Rows, [string]$Path, [int]$MaxRows)

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }
    $header = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })

    if (Test-Path -LiteralPath $Path) {
        $existing = @()
        try {
            $firstLine = @(Get-Content -LiteralPath $Path -TotalCount 1)
            if ($firstLine.Count -gt 0 -and $firstLine[0]) {
                # Column names here are plain identifiers, so splitting on the
                # comma is safe and avoids parsing the whole file to read them.
                $existing = @($firstLine[0].Split(',') | ForEach-Object { $_.Trim('"') })
            }
        }
        catch { }

        if ($existing.Count -gt 0 -and ($existing -join '|') -ne ($header -join '|')) {
            # Get-Date -Format honours the current culture's calendar, which
            # names this file 2569 on a Thai machine and breaks the
            # sort-by-name ordering the archives exist to have.
            $archive = Join-Path (Split-Path -Parent $Path) `
                (('{0}_upto_{1}.csv' -f [System.IO.Path]::GetFileNameWithoutExtension($Path), (Format-Invariant (Get-Date) 'yyyyMMdd-HHmmss')))
            Move-Item -LiteralPath $Path -Destination $archive -Force
            Write-Step ("history columns changed - earlier rows kept as {0}" -f (Split-Path -Leaf $archive))
        }
    }

    $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 -Append

    if ($MaxRows -gt 0) {
        $lines = @(Get-Content -LiteralPath $Path)
        if ($lines.Count -gt ($MaxRows + 1)) {
            $keep = @($lines[0]) + @($lines[($lines.Count - $MaxRows)..($lines.Count - 1)])
            $keep | Set-Content -LiteralPath $Path -Encoding utf8
        }
    }
}

function Merge-DailyCsv {
    <#
        The daily table is not append-only: today's row changes every time the
        script runs, because today is not over. Skipping a date that is already
        present would freeze today's counts at whatever the first run of the day
        saw, so every computed day replaces its stored twin and the file is
        rewritten in date order.

        Header drift is handled the same way as the append-only history: the old
        file is archived, not merged into and not deleted.

        -PreserveWhenBlank names the columns where "this run has no number for
        that day" is written as an empty field. A blank there keeps whatever an
        earlier run stored, because a counter that could not be read this time
        must not erase what a working run already measured.
    #>
    param([psobject[]]$Rows, [string]$Path, [int]$MaxRows, [string]$KeyColumn = 'Date',
          [string[]]$PreserveWhenBlank = @())

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }
    $header = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })

    $kept = New-Object System.Collections.Specialized.OrderedDictionary

    if (Test-Path -LiteralPath $Path) {
        $existingHeader = @()
        try {
            $firstLine = @(Get-Content -LiteralPath $Path -TotalCount 1)
            if ($firstLine.Count -gt 0 -and $firstLine[0]) {
                $existingHeader = @($firstLine[0].Split(',') | ForEach-Object { $_.Trim('"') })
            }
        }
        catch { }

        if ($existingHeader.Count -gt 0 -and ($existingHeader -join '|') -ne ($header -join '|')) {
            $archive = Join-Path (Split-Path -Parent $Path) `
                (('{0}_upto_{1}.csv' -f [System.IO.Path]::GetFileNameWithoutExtension($Path), (Format-Invariant (Get-Date) 'yyyyMMdd-HHmmss')))
            Move-Item -LiteralPath $Path -Destination $archive -Force
            Write-Step ("daily columns changed - earlier rows kept as {0}" -f (Split-Path -Leaf $archive))
        }
        else {
            try {
                foreach ($r in @(Import-Csv -LiteralPath $Path)) {
                    $k = [string](Get-Prop $r $KeyColumn '')
                    if ($k) { $kept[$k] = $r }
                }
            }
            catch { }
        }
    }

    foreach ($r in $Rows) {
        $k = [string](Get-Prop $r $KeyColumn '')
        if (-not $k) { continue }
        if ($PreserveWhenBlank.Count -gt 0 -and $kept.Contains($k)) {
            $stored = $kept[$k]
            foreach ($col in $PreserveWhenBlank) {
                if ([string](Get-Prop $r $col '') -ne '') { continue }
                $old = [string](Get-Prop $stored $col '')
                if ($old -ne '') { $r.$col = $old }
            }
        }
        $kept[$k] = $r
    }

    $ordered = @($kept.Keys | Sort-Object | ForEach-Object { $kept[$_] })
    if ($MaxRows -gt 0 -and $ordered.Count -gt $MaxRows) {
        $ordered = @($ordered[($ordered.Count - $MaxRows)..($ordered.Count - 1)])
    }
    $ordered | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8
}

# ------------------------------------------------------- event-record helpers
function Get-EventField {
    <#
        Event property counts are not a contract. A 1000 record has 15 elements
        here and fewer on some builds, and under strict mode reaching past the
        end throws and takes the whole query's worth of records with it.
    #>
    param($Record, [int]$Index)
    try {
        $props = $Record.Properties
        if ($null -eq $props -or $props.Count -le $Index) { return '' }
        $v = $props[$Index].Value
        if ($null -eq $v) { return '' }
        return ([string]$v).Trim()
    }
    catch { return '' }
}

function Clear-Sentinel {
    <#
        Firmware and WER both write "no value" as a magic number rather than an
        empty field: 4294967295 for an unknown hang termination time, 0.0.0.0
        for an unknown version, 00000000 for an unknown timestamp. Printing them
        as fact is worse than printing nothing.
    #>
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $v = $Value.Trim()
    if ($v -eq '4294967295' -or $v -eq '0.0.0.0' -or $v -eq '00000000' -or $v -eq '0') { return '' }
    $v
}

function Get-LeafName {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim().Trim('"')
    $i = $p.LastIndexOfAny([char[]]@('\', '/'))
    if ($i -ge 0 -and $i -lt ($p.Length - 1)) { return $p.Substring($i + 1) }
    $p
}

function ConvertTo-LocalPath {
    # WER writes its own paths in \\?\ extended-length form; nothing else in
    # this script wants to see that prefix.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    ($Path.Trim() -replace '^\\\\\?\\', '')
}

function Get-EventsSafe {
    <#
        Three outcomes have to stay distinct. Zero matching events is a normal,
        happy result and Get-WinEvent signals it by throwing
        NoMatchingEventsFound; an unregistered provider or an unreadable log is
        a genuine degradation; anything else is reported as one too.
    #>
    param([string]$Provider, [int[]]$EventId, [datetime]$Start, [datetime]$End, [int]$Max)

    $filter = @{
        LogName      = 'Application'
        ProviderName = $Provider
        Id           = $EventId
        StartTime    = $Start
        EndTime      = $End
    }
    try {
        $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents $Max -ErrorAction Stop)
        return [pscustomobject]@{ Ok = $true; Events = $events; Detail = '' }
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') {
            return [pscustomobject]@{ Ok = $true; Events = @(); Detail = '' }
        }
        return [pscustomobject]@{ Ok = $false; Events = @(); Detail = $_.Exception.Message }
    }
}

# -------------------------------------------------------------- phase 0: init
$RunStart = Get-Date
try {
    $OutputDir = Initialize-OutputDir -Preferred $OutputDir -ToolName 'GetCrashTrend'
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 5
}

$ReportFile   = Join-Path $OutputDir 'CrashTrend.html'
$DailyCsv     = Join-Path $OutputDir 'StabilityDaily.csv'
$RunHistoryCsv= Join-Path $OutputDir 'RunHistory.csv'

Write-Line ''
Write-Line "=== Crash trend : $(Format-Invariant $RunStart) ==="
Write-Line ''

$requestedFloor = $Now.AddDays(-$Days)
$oldestRecord   = $null
try {
    $oldestEvent = @(Get-WinEvent -LogName 'Application' -Oldest -MaxEvents 1 -ErrorAction Stop)
    if ($oldestEvent.Count -gt 0) { $oldestRecord = [datetime]$oldestEvent[0].TimeCreated }
}
catch { }

$effectiveFloor = $requestedFloor
$windowTruncated = $false
$windowEmpty     = $false
if ($null -ne $oldestRecord -and $oldestRecord -gt $requestedFloor) {
    $effectiveFloor  = $oldestRecord
    $windowTruncated = $true
}
# A clock moved backwards - or a -Now that predates every record in the log -
# would otherwise produce a negative window, which puts every event in the
# "recent" half and makes the shrunk-window message a lie.
if ($effectiveFloor -ge $Now) {
    $effectiveFloor  = $Now
    $windowTruncated = $false
    $windowEmpty     = $true
}

$windowSpan     = $Now - $effectiveFloor
$effectiveDays  = [math]::Round([math]::Max(0.0, $windowSpan.TotalDays), 1)
$halfDays       = [math]::Round($effectiveDays / 2.0, 1)
$midpoint       = $effectiveFloor.AddTicks([long]($windowSpan.Ticks / 2))

Write-Step ("window {0} to {1} ({2} days)" -f (Format-Invariant $effectiveFloor 'yyyy-MM-dd'), (Format-Invariant $Now 'yyyy-MM-dd'), $effectiveDays)
if ($windowEmpty) {
    Write-Step ("the Application log holds nothing at or before {0}, so the window is empty" -f (Format-Invariant $Now 'yyyy-MM-dd'))
}
elseif ($windowTruncated) {
    Write-Step ("the Application log only reaches back to {0}, so '-Days {1}' is really {2}" -f (Format-Invariant $effectiveFloor 'yyyy-MM-dd'), $Days, $effectiveDays)
}

# ------------------------------------------------------------- phase 1: count
Write-Step 'Counting application crashes and hangs...'

$HostProcesses = @('dllhost.exe', 'rundll32.exe', 'svchost.exe', 'taskhostw.exe')
$Records = New-Object System.Collections.Generic.List[psobject]

$crashQuery = Get-EventsSafe -Provider 'Application Error' -EventId @(1000) -Start $effectiveFloor -End $Now -Max $MaxEvents
if ($crashQuery.Ok) {
    foreach ($e in $crashQuery.Events) {
        $appPath  = Get-EventField $e 10
        $exeName  = Get-LeafName $appPath
        if (-not $exeName) { $exeName = Get-EventField $e 0 }
        if (-not $exeName) { continue }

        $module      = Get-EventField $e 3
        $moduleUnloaded = $false
        if ($module -and $module.ToLowerInvariant().EndsWith('_unloaded')) {
            $module = $module.Substring(0, $module.Length - 9)
            $moduleUnloaded = $true
        }

        $Records.Add([pscustomobject]@{
            Kind            = 'Crash'
            TimeCreated     = [datetime]$e.TimeCreated
            ExeName         = $exeName.ToLowerInvariant()
            ExeDisplay      = $exeName
            ExePath         = $appPath
            AppVersion      = Clear-Sentinel (Get-EventField $e 1)
            ModuleName      = $module
            ModulePath      = Get-EventField $e 11
            ModuleUnloaded  = $moduleUnloaded
            ExceptionCode   = Clear-Sentinel (Get-EventField $e 6)
            PackageFullName = Get-EventField $e 13
            ReportId        = (Get-EventField $e 12).ToLowerInvariant()
            IsHostProcess   = ($HostProcesses -contains $exeName.ToLowerInvariant())
            HangType        = ''
            DotNet          = ''
            Bucket          = ''
            WerFolder       = ''
        })
    }
}
else {
    Add-Health 'Application Error 1000 (crash counter)' 'degraded' $crashQuery.Detail
    Write-Step "! the crash counter could not be read: $($crashQuery.Detail)"
}

$hangQuery = Get-EventsSafe -Provider 'Application Hang' -EventId @(1002) -Start $effectiveFloor -End $Now -Max $MaxEvents
if ($hangQuery.Ok) {
    foreach ($e in $hangQuery.Events) {
        # Keyed on the executable path, not on the package: a dllhost.exe hang
        # on this machine carried a Visual Studio Code package name, and taking
        # that as the identity would file the surrogate's hang under VS Code.
        $exePath = Get-EventField $e 5
        $exeName = Get-LeafName $exePath
        if (-not $exeName) { $exeName = Get-EventField $e 0 }
        if (-not $exeName) { continue }

        $Records.Add([pscustomobject]@{
            Kind            = 'Hang'
            TimeCreated     = [datetime]$e.TimeCreated
            ExeName         = $exeName.ToLowerInvariant()
            ExeDisplay      = $exeName
            ExePath         = $exePath
            AppVersion      = Clear-Sentinel (Get-EventField $e 1)
            ModuleName      = ''
            ModulePath      = ''
            ModuleUnloaded  = $false
            ExceptionCode   = ''
            PackageFullName = Get-EventField $e 7
            ReportId        = (Get-EventField $e 6).ToLowerInvariant()
            IsHostProcess   = ($HostProcesses -contains $exeName.ToLowerInvariant())
            HangType        = Get-EventField $e 9
            DotNet          = ''
            Bucket          = ''
            WerFolder       = ''
        })
    }
}
else {
    Add-Health 'Application Hang 1002 (hang counter)' 'degraded' $hangQuery.Detail
    Write-Step "! the hang counter could not be read: $($hangQuery.Detail)"
}

if ($crashQuery.Ok) {
    $note = ''
    if ($crashQuery.Events.Count -ge $MaxEvents) { $note = "query hit the -MaxEvents cap of $MaxEvents; older crashes in the window were not read" }
    Add-Health 'Application Error 1000 (crash counter)' 'ok' $note
}
if ($hangQuery.Ok) {
    $note = ''
    if ($hangQuery.Events.Count -ge $MaxEvents) { $note = "query hit the -MaxEvents cap of $MaxEvents; older hangs in the window were not read" }
    Add-Health 'Application Hang 1002 (hang counter)' 'ok' $note
}

$CrashTotal = @($Records | Where-Object { $_.Kind -eq 'Crash' }).Count
$HangTotal  = @($Records | Where-Object { $_.Kind -eq 'Hang' }).Count

# "Nothing was counted" and "nothing could be counted" produce an identical
# empty $Records, and every number, verdict and stored row downstream depends
# on telling them apart. A counter that threw has no reading at all - not a
# reading of zero - and an all-clear must never be built on one.
$crashCounterOk = [bool]$crashQuery.Ok
$hangCounterOk  = [bool]$hangQuery.Ok
$countersOk     = $crashCounterOk -and $hangCounterOk

# A query that hit -MaxEvents read the newest events and stopped, so the oldest
# days of the window were cut off mid-count. Those days have no reading either;
# only the days after the oldest event that came back were seen in full.
$crashCapped = $crashCounterOk -and ($crashQuery.Events.Count -ge $MaxEvents)
$hangCapped  = $hangCounterOk  -and ($hangQuery.Events.Count -ge $MaxEvents)

$crashCountedFromKey = ''
$hangCountedFromKey  = ''
if ($crashCapped) {
    $t = @($Records | Where-Object { $_.Kind -eq 'Crash' } | ForEach-Object { $_.TimeCreated } | Sort-Object)
    if ($t.Count -gt 0) { $crashCountedFromKey = Format-Invariant $t[0] 'yyyy-MM-dd' }
}
if ($hangCapped) {
    $t = @($Records | Where-Object { $_.Kind -eq 'Hang' } | ForEach-Object { $_.TimeCreated } | Sort-Object)
    if ($t.Count -gt 0) { $hangCountedFromKey = Format-Invariant $t[0] 'yyyy-MM-dd' }
}

Write-Step ("{0} crash(es), {1} hang(s)" -f $CrashTotal, $HangTotal)
if (-not $countersOk) {
    Write-Step '! those totals are what could be read, not what happened'
}

# ------------------------------------------------------------ phase 2: enrich
Write-Step 'Reading Windows Error Reporting and .NET detail...'

# .NET Runtime 1026 is a companion to a 1000, not another crash: the runtime
# logs the managed exception and the OS logs the process death a moment later.
# Counting both would double every managed crash on the machine.
$dotnetOrphans = 0
$dotnetQuery = Get-EventsSafe -Provider '.NET Runtime' -EventId @(1026) -Start $effectiveFloor -End $Now -Max $MaxEvents
if ($dotnetQuery.Ok) {
    foreach ($e in $dotnetQuery.Events) {
        $blob = Get-EventField $e 0
        if (-not $blob) { continue }
        $exe  = ''
        $code = ''
        if ($blob -match '(?m)^Application:\s*(.+?)\s*$')      { $exe  = $Matches[1] }
        if ($blob -match 'exception code\s+([0-9a-fA-F]+)')    { $code = $Matches[1] }
        if (-not $exe) { continue }

        $leaf  = (Get-LeafName $exe).ToLowerInvariant()
        $when  = [datetime]$e.TimeCreated
        $match = @($Records | Where-Object {
            $_.Kind -eq 'Crash' -and $_.ExeName -eq $leaf -and
            [math]::Abs(($_.TimeCreated - $when).TotalSeconds) -le 5
        })
        if ($match.Count -eq 0) { $dotnetOrphans++; continue }
        foreach ($m in $match) {
            $m.DotNet = 'unhandled .NET exception'
            if ($code) { $m.DotNet = "unhandled .NET exception (code $code)" }
        }
    }
    Add-Health '.NET Runtime 1026 (enrichment)' 'ok' ''
}
else {
    Add-Health '.NET Runtime 1026 (enrichment)' 'degraded' $dotnetQuery.Detail
}

# 1001 is enrichment and nothing else. On this machine 3,770 of them were
# logged in 30 days, of which 3,709 were BlueScreen and LiveKernelEvent records
# re-logged on every upload retry - a league table built on 1001 would put
# "BlueScreen" at the top and call it a crashing application.
$WerEventNames = @('APPCRASH', 'BEX', 'BEX64', 'AppHangB1', 'MoAppCrash', 'MoAppHang')
$werByReportId = @{}
$werRawCount   = 0
$werKeptCount  = 0

$werQuery = Get-EventsSafe -Provider 'Windows Error Reporting' -EventId @(1001) -Start $effectiveFloor -End $Now -Max $MaxEvents
if ($werQuery.Ok) {
    $werRawCount = $werQuery.Events.Count
    foreach ($e in $werQuery.Events) {
        $eventName = Get-EventField $e 2
        if ($WerEventNames -notcontains $eventName) { continue }
        $reportId = (Get-EventField $e 19).ToLowerInvariant()
        if (-not $reportId) { continue }
        # Same report, re-logged on each upload attempt; the first one seen
        # (the newest) is the whole record.
        if ($werByReportId.ContainsKey($reportId)) { continue }

        # P4 is the faulting module on a crash report and a hang signature on a
        # hang report, so it is only read as a module for the crash event names.
        # StackHash_xxxx and PCH_... are WER saying it could not resolve a
        # module, not the name of one, and a value with no extension is a hash
        # rather than a file.
        $p4 = ''
        if (@('APPCRASH', 'BEX', 'BEX64', 'MoAppCrash') -contains $eventName) {
            $p4 = Get-EventField $e 8
            if ($p4 -like 'StackHash_*' -or $p4 -like 'PCH_*' -or $p4 -notlike '*.*') { $p4 = '' }
        }

        $werByReportId[$reportId] = [pscustomobject]@{
            EventName = $eventName
            Bucket    = Get-EventField $e 21
            StorePath = ConvertTo-LocalPath (Get-EventField $e 16)
            P4        = $p4
        }
        $werKeptCount++
    }

    $noise = $werRawCount - $werKeptCount
    $detail = 'no Error Reporting events in the window'
    if ($werRawCount -gt 0) {
        $detail = "{0} of {1} events in the window were application reports; the other {2} are kernel and upload-retry records and were ignored" -f $werKeptCount, $werRawCount, $noise
    }
    if ($werQuery.Events.Count -ge $MaxEvents) {
        $detail += ". The query hit the -MaxEvents cap of $MaxEvents, so some enrichment may be missing"
    }
    Add-Health 'Windows Error Reporting 1001 (enrichment only)' 'ok' $detail
    Write-Step ("{0} WER application report(s) kept out of {1} events ({2} kernel or retry records ignored)" -f $werKeptCount, $werRawCount, $noise)
}
else {
    Add-Health 'Windows Error Reporting 1001 (enrichment only)' 'degraded' $werQuery.Detail
}

foreach ($r in $Records) {
    if (-not $r.ReportId) { continue }
    if (-not $werByReportId.ContainsKey($r.ReportId)) { continue }
    $w = $werByReportId[$r.ReportId]
    $r.Bucket    = $w.Bucket
    $r.WerFolder = $w.StorePath
    if (-not $r.ModuleName -and $w.P4) {
        $module = $w.P4
        if ($module.ToLowerInvariant().EndsWith('_unloaded')) {
            $module = $module.Substring(0, $module.Length - 9)
            $r.ModuleUnloaded = $true
        }
        $r.ModuleName = $module
    }
}

# ------------------------------------------------------- phase 3: WER folders
$werRootsSeen   = 0
$werDirsTotal   = 0
$werDirsWindow  = 0
$werDirsRead    = 0
$werDirsDenied  = 0
$werNameByExe   = @{}

if ($SkipWerStore) {
    Add-Health 'WER report folders' 'skipped' '-SkipWerStore was given'
}
else {
    Write-Step 'Sweeping the Windows Error Reporting folders...'

    $werRoots = @(
        'C:\ProgramData\Microsoft\Windows\WER\ReportArchive'
        'C:\ProgramData\Microsoft\Windows\WER\ReportQueue'
    )
    if ($env:ProgramData) {
        $werRoots = @(
            (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportArchive')
            (Join-Path $env:ProgramData 'Microsoft\Windows\WER\ReportQueue')
        )
    }
    if ($env:LOCALAPPDATA) {
        $werRoots += (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportArchive')
        $werRoots += (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportQueue')
    }

    # Names and timestamps only at this stage. Touching contents is what costs
    # time and what trips the ACLs, so the window filter happens first.
    $candidates = New-Object System.Collections.Generic.List[psobject]
    foreach ($root in $werRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $werRootsSeen++
        try {
            foreach ($d in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
                $werDirsTotal++
                if ($d.LastWriteTime -lt $effectiveFloor -or $d.LastWriteTime -gt $Now) { continue }
                $werDirsWindow++
                $candidates.Add($d)
            }
        }
        catch { }
    }

    # A folder named by a 1001 we already kept is worth reading first; the cap
    # then spends what is left of its budget on the rest of the window.
    $joinable = @{}
    foreach ($w in $werByReportId.Values) {
        if ($w.StorePath) { $joinable[(Get-LeafName $w.StorePath).ToLowerInvariant()] = $true }
    }
    $ordered = @(
        @($candidates | Where-Object { $joinable.ContainsKey($_.Name.ToLowerInvariant()) })
        @($candidates | Where-Object { -not $joinable.ContainsKey($_.Name.ToLowerInvariant()) })
    )

    $budget = 200
    foreach ($d in $ordered) {
        if ($budget -le 0) { break }
        $budget--
        $werPath = Join-Path $d.FullName 'Report.wer'
        try {
            $fi = Get-Item -LiteralPath $werPath -ErrorAction Stop
            # A Report.wer is tens of kilobytes; anything past a few megabytes
            # is not one and is not worth reading into memory.
            if ($fi.Length -gt 4MB) { continue }
            # Report.wer is UTF-16 LE with a byte-order mark here, but the BOM
            # is what says so - the reader is told to detect rather than assume.
            $reader = New-Object System.IO.StreamReader($fi.FullName, [System.Text.Encoding]::UTF8, $true)
            try   { $text = $reader.ReadToEnd() }
            finally { $reader.Dispose() }
            $werDirsRead++
        }
        catch {
            $werDirsDenied++
            continue
        }

        $appName = ''
        $appPath = ''
        $sigNames = @{}
        foreach ($line in ($text -split "`r?`n")) {
            $eq = $line.IndexOf('=')
            if ($eq -lt 1) { continue }
            $key = $line.Substring(0, $eq)
            $val = $line.Substring($eq + 1)
            if     ($key -eq 'AppName') { $appName = $val }
            elseif ($key -eq 'AppPath') { $appPath = $val }
            elseif ($key -match '^Sig\[(\d+)\]\.Name$')  { $sigNames[('n' + $Matches[1])] = $val }
            elseif ($key -match '^Sig\[(\d+)\]\.Value$') { $sigNames[('v' + $Matches[1])] = $val }
        }

        # Sig entries are paired by name text, never by index: the index that
        # holds "Application Name" is not fixed across event types.
        $sigAppName = ''
        foreach ($k in @($sigNames.Keys)) {
            if ($k -notlike 'n*') { continue }
            if ([string]$sigNames[$k] -ne 'Application Name') { continue }
            $vk = 'v' + $k.Substring(1)
            if ($sigNames.ContainsKey($vk)) { $sigAppName = [string]$sigNames[$vk] }
        }

        $exeLeaf = (Get-LeafName $appPath).ToLowerInvariant()
        if (-not $exeLeaf) { $exeLeaf = (Get-LeafName $sigAppName).ToLowerInvariant() }
        if (-not $exeLeaf -or -not $appName) { continue }
        # Windows falls back to the bare file name when it has no product name,
        # which is not a friendlier name than the one already in hand.
        if ($appName.ToLowerInvariant() -eq $exeLeaf) { continue }
        if (-not $werNameByExe.ContainsKey($exeLeaf)) { $werNameByExe[$exeLeaf] = $appName }
    }

    $detail = "{0} report folder(s) on disk, {1} inside the window; {2} read, {3} belong to other accounts (normal for a standard user)" -f `
        $werDirsTotal, $werDirsWindow, $werDirsRead, $werDirsDenied
    Add-Health 'WER report folders' 'ok' $detail
    Write-Step ("{0} report folder(s) in the window - {1} readable, {2} owned by other accounts" -f $werDirsWindow, $werDirsRead, $werDirsDenied)
}

# ----------------------------------------------------- phase 4: friendly names
Write-Step 'Resolving application names...'

$uninstallEntries = New-Object System.Collections.Generic.List[psobject]
foreach ($hive in @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')) {

    if (-not (Test-Path -LiteralPath $hive)) { continue }
    try { $keys = @(Get-ChildItem -LiteralPath $hive -ErrorAction Stop) } catch { continue }

    foreach ($k in $keys) {
        $values = $null
        try { $values = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction Stop } catch { continue }

        $displayName = [string](Get-Prop $values 'DisplayName' '')
        if (-not $displayName.Trim()) { continue }

        $icon = [string](Get-Prop $values 'DisplayIcon' '')
        $iconLeaf = ''
        if ($icon) {
            # DisplayIcon is "<path>,<index>" as often as it is a bare path, and
            # a .ico file says nothing about which executable this is.
            $iconPath = ($icon -replace ',\s*-?\d+\s*$', '').Trim().Trim('"')
            if ($iconPath -and -not $iconPath.ToLowerInvariant().EndsWith('.ico')) {
                $iconLeaf = (Get-LeafName $iconPath).ToLowerInvariant()
            }
        }

        # Riot's installers write InstallLocation with forward slashes, so a
        # plain prefix comparison against the backslash path in the event never
        # matches unless both sides are normalised first.
        $location = [string](Get-Prop $values 'InstallLocation' '')
        $locPrefix = ''
        if ($location.Trim()) {
            $locPrefix = $location.Trim().Trim('"').Replace('/', '\').TrimEnd('\').ToLowerInvariant()
        }

        $uninstallEntries.Add([pscustomobject]@{
            DisplayName = $displayName.Trim()
            IconLeaf    = $iconLeaf
            LocPrefix   = $locPrefix
        })
    }
}

function Resolve-FriendlyName {
    <#
        First tier that answers wins. The package tier is skipped for the
        generic Windows host processes: the package name on a dllhost.exe hang
        belongs to whatever asked for the surrogate, not to the thing that hung.
    #>
    param([string]$ExeName, [string]$ExeDisplay, [string]$ExePath, [string]$PackageFullName, [bool]$IsHostProcess)

    if (-not $IsHostProcess -and $PackageFullName) {
        $segment = ($PackageFullName -split '_')[0]
        if ($segment) { return $segment }
    }

    if ($uninstallEntries.Count -gt 0) {
        foreach ($u in $uninstallEntries) {
            if ($u.IconLeaf -and $u.IconLeaf -eq $ExeName) { return $u.DisplayName }
        }
        if ($ExePath) {
            $normalised = $ExePath.Replace('/', '\').ToLowerInvariant()
            $best = ''
            $bestLen = 0
            foreach ($u in $uninstallEntries) {
                if (-not $u.LocPrefix) { continue }
                if (-not $normalised.StartsWith($u.LocPrefix + '\')) { continue }
                # The deepest matching install location is the specific product
                # rather than the vendor folder above it.
                if ($u.LocPrefix.Length -gt $bestLen) { $best = $u.DisplayName; $bestLen = $u.LocPrefix.Length }
            }
            if ($best) { return $best }
        }
    }

    if ($werNameByExe.ContainsKey($ExeName)) { return [string]$werNameByExe[$ExeName] }

    if ($ExePath) {
        try {
            $vi = (Get-Item -LiteralPath $ExePath -ErrorAction Stop).VersionInfo
            $desc = [string](Get-Prop $vi 'FileDescription' '')
            if ($desc.Trim() -and $desc.ToLowerInvariant() -ne $ExeName) { return $desc.Trim() }
            $product = [string](Get-Prop $vi 'ProductName' '')
            if ($product.Trim()) { return $product.Trim() }
        }
        catch { }
    }

    $ExeDisplay
}

# --------------------------------------------------------- phase 5: aggregate
Write-Step 'Grouping by application...'

$SystemDirs = @('c:\windows\system32\', 'c:\windows\syswow64\', 'c:\windows\winsxs\')

function Get-ModuleVerdict {
    <#
        Whose fault the module is, in the only three categories worth telling a
        non-technical owner apart. An unloaded module is always the third: a DLL
        that was gone by the time the fault was recorded is characteristic of an
        injected plugin, codec or overlay rather than of the app's own code.
    #>
    param([string]$ModulePath, [string]$AppPath, [bool]$Unloaded)

    if ($Unloaded) { return 'a third-party DLL loaded into the app - often a plugin, codec, or overlay' }
    if (-not $ModulePath) { return 'a module Windows could not place' }

    $m = $ModulePath.Replace('/', '\').ToLowerInvariant()
    foreach ($s in $SystemDirs) {
        if ($m.StartsWith($s)) { return 'a Windows component' }
    }
    if ($AppPath) {
        $appDir = ''
        try { $appDir = ([System.IO.Path]::GetDirectoryName($AppPath.Replace('/', '\'))).ToLowerInvariant() } catch { }
        if ($appDir -and $m.StartsWith($appDir + '\')) { return 'the app itself' }
    }
    'a third-party DLL loaded into the app - often a plugin, codec, or overlay'
}

$Apps = New-Object System.Collections.Generic.List[psobject]
foreach ($group in @($Records | Group-Object -Property ExeName)) {
    $rows      = @($group.Group)
    $crashes   = @($rows | Where-Object { $_.Kind -eq 'Crash' })
    $hangs     = @($rows | Where-Object { $_.Kind -eq 'Hang' })
    $recent    = @($rows | Where-Object { $_.TimeCreated -ge $midpoint })
    $earlier   = @($rows | Where-Object { $_.TimeCreated -lt $midpoint })
    $times     = @($rows | ForEach-Object { $_.TimeCreated } | Sort-Object)

    $exeDisplay = [string]$rows[0].ExeDisplay
    $exePath    = ''
    $paths      = @($rows | ForEach-Object { $_.ExePath } | Where-Object { $_ } | Sort-Object -Unique)
    if ($paths.Count -gt 0) { $exePath = [string]$paths[0] }

    $package = ''
    $packages = @($rows | ForEach-Object { $_.PackageFullName } | Where-Object { $_ } | Sort-Object -Unique)
    if ($packages.Count -gt 0) { $package = [string]$packages[0] }

    $isHost = $false
    foreach ($r in $rows) { if ($r.IsHostProcess) { $isHost = $true } }

    $friendly = Resolve-FriendlyName -ExeName ([string]$group.Name) -ExeDisplay $exeDisplay `
        -ExePath $exePath -PackageFullName $package -IsHostProcess $isHost

    # Culprit: one module has to dominate before it is worth naming. A module
    # seen once is a coincidence, not a pattern.
    $culpritModule  = ''
    $culpritHits    = 0
    $culpritVerdict = ''
    $moduleGroups = @($crashes | Where-Object { $_.ModuleName } | Group-Object -Property ModuleName | Sort-Object Count -Descending)
    if ($moduleGroups.Count -gt 0 -and $crashes.Count -gt 0) {
        $topModule = $moduleGroups[0]
        $share = 1.0 * $topModule.Count / $crashes.Count
        if ($topModule.Count -ge 2 -and $share -ge 0.60) {
            $sample = @($topModule.Group)[0]
            $culpritModule  = [string]$topModule.Name
            $culpritHits    = [int]$topModule.Count
            $culpritVerdict = Get-ModuleVerdict -ModulePath ([string]$sample.ModulePath) -AppPath $exePath -Unloaded ([bool]$sample.ModuleUnloaded)
        }
    }

    $trend = 'steady'
    if ($recent.Count -gt $earlier.Count)     { $trend = 'rising' }
    elseif ($recent.Count -lt $earlier.Count) { $trend = 'falling' }

    $Apps.Add([pscustomobject]@{
        ExeName        = [string]$group.Name
        ExeDisplay     = $exeDisplay
        Friendly       = $friendly
        ExePath        = $exePath
        AllPaths       = $paths
        Package        = $package
        IsHostProcess  = $isHost
        Crashes        = $crashes.Count
        Hangs          = $hangs.Count
        Total          = $rows.Count
        Recent         = $recent.Count
        Earlier        = $earlier.Count
        Trend          = $trend
        FirstSeen      = $times[0]
        LastSeen       = $times[$times.Count - 1]
        CulpritModule  = $culpritModule
        CulpritHits    = $culpritHits
        CulpritVerdict = $culpritVerdict
        Codes          = @($crashes | ForEach-Object { $_.ExceptionCode } | Where-Object { $_ } | Sort-Object -Unique)
        Buckets        = @($rows | ForEach-Object { $_.Bucket } | Where-Object { $_ } | Sort-Object -Unique)
        HangTypes      = @($hangs | ForEach-Object { $_.HangType } | Where-Object { $_ } | Sort-Object -Unique)
        DotNet         = @($crashes | ForEach-Object { $_.DotNet } | Where-Object { $_ } | Sort-Object -Unique)
        Records        = $rows
    })
}

$AppsRanked = @($Apps | Sort-Object -Property `
    @{ Expression = 'Total';    Descending = $true }, `
    @{ Expression = 'LastSeen'; Descending = $true }, `
    @{ Expression = 'ExeName';  Descending = $false })

$RecentHalf = @($Records | Where-Object { $_.TimeCreated -ge $midpoint }).Count
$EarlierHalf= @($Records | Where-Object { $_.TimeCreated -lt $midpoint }).Count

# --------------------------------------------------------- phase 6: stability
Write-Step 'Reading the reliability history...'

$stabilityByDay = @{}
$stabilityOk    = $false
$stabilityDetail= ''
try {
    # The only CIM call in the script, and the only realistic way for it to
    # hang; a broken WMI repository blocks indefinitely without a timeout.
    $metrics = @(Get-CimInstance -Namespace 'root\cimv2' -ClassName 'Win32_ReliabilityStabilityMetrics' `
        -OperationTimeoutSec 15 -ErrorAction Stop)

    if ($metrics.Count -eq 0) {
        # An empty result is what a machine with the Reliability Analysis
        # Component disabled looks like. That is a missing source, not an error.
        $stabilityDetail = 'the class returned no rows - Reliability Monitor is probably turned off on this machine'
        Add-Health 'Win32_ReliabilityStabilityMetrics' 'degraded' $stabilityDetail
    }
    else {
        foreach ($m in $metrics) {
            $when = Get-Prop $m 'TimeGenerated'
            $ssi  = Get-Prop $m 'SystemStabilityIndex'
            if ($null -eq $when -or $null -eq $ssi) { continue }
            $when = [datetime]$when
            $key  = Format-Invariant $when 'yyyy-MM-dd'
            # Hourly samples; the last one of a day is that day's standing.
            if (-not $stabilityByDay.ContainsKey($key) -or $stabilityByDay[$key].When -lt $when) {
                $stabilityByDay[$key] = [pscustomobject]@{ When = $when; Index = [double]$ssi }
            }
        }
        $stabilityOk = $stabilityByDay.Count -gt 0
        $stabilityDetail = "{0} hourly sample(s) covering {1} day(s)" -f $metrics.Count, $stabilityByDay.Count
        Add-Health 'Win32_ReliabilityStabilityMetrics' 'ok' $stabilityDetail
    }
}
catch {
    $stabilityDetail = $_.Exception.Message
    Add-Health 'Win32_ReliabilityStabilityMetrics' 'degraded' $stabilityDetail
    Write-Step "! the reliability index could not be read: $stabilityDetail"
}

# Both counters dead and no reliability data means there is nothing to report
# on, which is the one fatal outcome that is not about the output folder.
if (-not $crashQuery.Ok -and -not $hangQuery.Ok -and -not $stabilityOk) {
    Write-Line ''
    Write-Line 'No usable data: neither event counter could be read and the reliability class is unavailable.' 'Red'
    exit 4
}

$crashesByDay = @{}
$hangsByDay   = @{}
foreach ($r in $Records) {
    $key = Format-Invariant $r.TimeCreated 'yyyy-MM-dd'
    if ($r.Kind -eq 'Crash') {
        if ($crashesByDay.ContainsKey($key)) { $crashesByDay[$key] = $crashesByDay[$key] + 1 } else { $crashesByDay[$key] = 1 }
    }
    else {
        if ($hangsByDay.ContainsKey($key)) { $hangsByDay[$key] = $hangsByDay[$key] + 1 } else { $hangsByDay[$key] = 1 }
    }
}

$dayKeys = @(@($stabilityByDay.Keys) + @($crashesByDay.Keys) + @($hangsByDay.Keys) | Sort-Object -Unique)
$floorKey = Format-Invariant $effectiveFloor 'yyyy-MM-dd'
$nowKey   = Format-Invariant $Now 'yyyy-MM-dd'
$dayKeys  = @($dayKeys | Where-Object { $_ -ge $floorKey -and $_ -le $nowKey })

$DailyRows = New-Object System.Collections.Generic.List[psobject]
foreach ($k in $dayKeys) {
    # A day is only counted where the counter worked and the query reached back
    # past it; anywhere else the field is blank, which the merge below reads as
    # "no reading this run" and leaves the stored one alone.
    $crashKnown = ($crashCounterOk -and $k -gt $crashCountedFromKey)
    $hangKnown  = ($hangCounterOk  -and $k -gt $hangCountedFromKey)
    $haveIndex  = $stabilityByDay.ContainsKey($k)
    if (-not $haveIndex -and -not $crashKnown -and -not $hangKnown) { continue }

    # Export-Csv serialises a [double] through the current culture, so a German
    # or French machine writes "6,652" into a file every reader parses as
    # invariant - a 0-to-10 index that comes back as 6652.
    $idx = ''
    if ($haveIndex) {
        $idx = ([math]::Round($stabilityByDay[$k].Index, 3)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    $c = ''
    if ($crashKnown) {
        $c = 0
        if ($crashesByDay.ContainsKey($k)) { $c = $crashesByDay[$k] }
    }
    $h = ''
    if ($hangKnown) {
        $h = 0
        if ($hangsByDay.ContainsKey($k)) { $h = $hangsByDay[$k] }
    }
    $DailyRows.Add([pscustomobject]@{
        Date                 = $k
        SystemStabilityIndex = $idx
        CrashCount           = $c
        HangCount            = $h
    })
}

$ssiValues = @($DailyRows | Where-Object { $_.SystemStabilityIndex -ne '' } | ForEach-Object { [double]$_.SystemStabilityIndex })
$ssiLatest = $null
$ssiAverage= $null
$ssiMin    = $null
if ($ssiValues.Count -gt 0) {
    $ssiLatest  = [math]::Round($ssiValues[$ssiValues.Count - 1], 2)
    $ssiAverage = [math]::Round((($ssiValues | Measure-Object -Average).Average), 2)
    $ssiMin     = [math]::Round((($ssiValues | Measure-Object -Minimum).Minimum), 2)
}

# ------------------------------------------------------ phase 7: the verdict
$EventTotal = $Records.Count
$topApp     = $null
if ($AppsRanked.Count -gt 0) { $topApp = $AppsRanked[0] }

$worseReasons = New-Object System.Collections.Generic.List[string]
if ($EventTotal -gt 0 -and $RecentHalf -ge [math]::Max(3, 2 * $EarlierHalf)) {
    $worseReasons.Add(("{0} of the {1} events fell in the most recent half of the window (the first half had {2})" -f $RecentHalf, $EventTotal, $EarlierHalf))
}
if ($null -ne $ssiLatest -and $null -ne $ssiAverage -and $ssiLatest -le ($ssiAverage - 2.0)) {
    $worseReasons.Add(("the stability index has fallen to {0}, well below its {1} average for the window" -f $ssiLatest, $ssiAverage))
}
if ($null -ne $topApp -and $topApp.Total -ge 10) {
    $worseReasons.Add(("{0} alone accounts for {1} events" -f $topApp.Friendly, $topApp.Total))
}

# What a failed counter costs the numbers above, said once and reused by the
# console, the headline and the report.
$counterNote = ''
if (-not $crashCounterOk -and -not $hangCounterOk) {
    $counterNote = 'Neither counter could be read this run - Application Error 1000 and Application Hang 1002 both failed - so no crash and no hang was countable.'
}
elseif (-not $crashCounterOk) {
    $counterNote = 'The crash counter (Application Error 1000) could not be read this run, so crashes are missing from every count here.'
}
elseif (-not $hangCounterOk) {
    $counterNote = 'The hang counter (Application Hang 1002) could not be read this run, so hangs are missing from every count here.'
}
if ($crashCapped -or $hangCapped) {
    $capNote = 'A counter query hit the -MaxEvents cap of {0}, so the oldest days of the window are counted short.' -f $MaxEvents
    if ($counterNote) { $counterNote += ' ' + $capNote } else { $counterNote = $capNote }
}

$ExitCode = 1
$Verdict  = 'Crashes and hangs are present, and the trend is flat or improving.'
if ($EventTotal -eq 0 -and -not $countersOk) {
    # Exit 0 states a fact - zero crashes and zero hangs in the window - that a
    # run which could not read the counters has not established. Calling this
    # an all-clear is the one answer the tool must never give.
    $Verdict = 'The crash and hang counters could not be read, so nothing can be said about the applications on this machine.'
    if ($worseReasons.Count -gt 0) {
        $ExitCode = 2
        $Verdict  = 'The crash and hang counters could not be read, and what could still be read is getting worse.'
    }
}
elseif ($EventTotal -eq 0) {
    $ExitCode = 0
    $Verdict  = 'Nothing crashed or hung.'
}
elseif ($worseReasons.Count -gt 0) {
    $ExitCode = 2
    $Verdict  = 'Worsening.'
}

# The headline sentence, in the words someone would use out loud.
$headline = ''
if ($windowEmpty) {
    $headline = 'The Application log holds no records at or before {0}, so there is nothing here to measure.' -f (Format-Invariant $Now 'yyyy-MM-dd HH:mm')
}
elseif ($EventTotal -eq 0 -and -not $countersOk) {
    $headline = "No crash or hang could be counted over the last $effectiveDays days, which is not the same as none having happened."
}
elseif ($EventTotal -eq 0) {
    $headline = "No application crashes or hangs were recorded in the last $effectiveDays days."
}
else {
    $verb = ''
    if ($topApp.Crashes -gt 0 -and $topApp.Hangs -gt 0) {
        $verb = 'crashed {0} time(s) and hung {1} time(s)' -f $topApp.Crashes, $topApp.Hangs
    }
    elseif ($topApp.Crashes -gt 0) {
        $verb = 'crashed {0} time(s)' -f $topApp.Crashes
    }
    else {
        $verb = 'hung {0} time(s)' -f $topApp.Hangs
    }

    $headline = '{0} {1} in the last {2} days' -f $topApp.Friendly, $verb, $effectiveDays
    if ($topApp.Recent -gt 0 -and $halfDays -gt 0) {
        $headline += ' - {0} of them in the last {1} days -' -f $topApp.Recent, $halfDays
    }
    if ($topApp.CulpritModule) {
        $headline += ' and {0} of {1} crashes point at {2} ({3}).' -f `
            $topApp.CulpritHits, $topApp.Crashes, $topApp.CulpritVerdict, $topApp.CulpritModule
    }
    else {
        $headline += '.'
    }

    if ($EventTotal -gt $topApp.Total) {
        $headline += ' {0} other application(s) also faulted in the window.' -f ($Apps.Count - 1)
    }
}

if ($counterNote) { $headline = $headline.TrimEnd() + ' ' + $counterNote }

if ($null -ne $ssiLatest) {
    if ($ssiLatest -ge 8) {
        $headline += ' The rest of the PC looks stable (stability index {0} out of 10).' -f $ssiLatest
    }
    else {
        $headline += ' Windows rates this machine''s overall stability at {0} out of 10, against a {1} average for the window.' -f $ssiLatest, $ssiAverage
    }
}

# ------------------------------------------------------------- console output
Write-Line ''
$verdictColour = 'Green'
if ($ExitCode -eq 1) { $verdictColour = 'Yellow' }
if ($ExitCode -eq 2) { $verdictColour = 'Red' }

Write-Line 'Verdict'
Write-Line '-------'
if (-not $Quiet) { Write-Host ("  " + $Verdict) -ForegroundColor $verdictColour }
Write-Line ("  " + $headline)
foreach ($reason in $worseReasons) { Write-Line ("  why: {0}" -f $reason) 'Yellow' }

if ($AppsRanked.Count -gt 0) {
    Write-Line ''
    Write-Line 'League table'
    Write-Line '------------'
    Write-Line ('  {0,-30} {1,-28} {2,7} {3,6} {4,8}  {5}' -f 'Application', 'Executable', 'Crashes', 'Hangs', 'Trend', 'Last seen')
    foreach ($a in @($AppsRanked | Select-Object -First $Top)) {
        Write-Line ('  {0,-30} {1,-28} {2,7} {3,6} {4,8}  {5}' -f `
            ($a.Friendly.PadRight(30).Substring(0, 30)),
            ($a.ExeDisplay.PadRight(28).Substring(0, 28)),
            $a.Crashes, $a.Hangs, $a.Trend, (Format-Invariant $a.LastSeen 'yyyy-MM-dd HH:mm'))
    }
}

# ------------------------------------------------------------------ the CSVs
$csvOk = $true
try {
    Merge-DailyCsv -Rows @($DailyRows) -Path $DailyCsv -MaxRows $HistoryLimit `
        -PreserveWhenBlank @('SystemStabilityIndex', 'CrashCount', 'HangCount')
}
catch {
    $csvOk = $false
    Write-Step "! could not update $DailyCsv : $($_.Exception.Message)"
}

$topAppName = ''
if ($null -ne $topApp) { $topAppName = $topApp.Friendly }
# A counter that failed contributes no total; a stored 0 would read as a
# measurement forever after.
$crashesCell = $CrashTotal
if (-not $crashCounterOk) { $crashesCell = '' }
$hangsCell = $HangTotal
if (-not $hangCounterOk) { $hangsCell = '' }
try {
    Add-HistoryRow -Rows @([pscustomobject]@{
        RunTimeUtc     = Format-Invariant ($Now.ToUniversalTime())
        WindowDays     = $effectiveDays.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        EffectiveFloor = Format-Invariant $effectiveFloor
        Crashes        = $crashesCell
        Hangs          = $hangsCell
        TopApp         = $topAppName
        Verdict        = $Verdict
        ExitCode       = $ExitCode
    }) -Path $RunHistoryCsv -MaxRows $HistoryLimit
}
catch {
    $csvOk = $false
    Write-Step "! could not update $RunHistoryCsv : $($_.Exception.Message)"
}

# --------------------------------------------------------- the HTML sections
Write-Step 'Rendering HTML...'

# 2. machine trend
$crashText = '{0} crash(es)' -f $CrashTotal
if (-not $crashCounterOk) { $crashText = 'crashes not counted' }
$hangText = '{0} hang(s)' -f $HangTotal
if (-not $hangCounterOk) { $hangText = 'hangs not counted' }

$trendPairs = @{
    'Window'              = '{0} to {1} ({2} days)' -f (Format-Invariant $effectiveFloor 'yyyy-MM-dd HH:mm'), (Format-Invariant $Now 'yyyy-MM-dd HH:mm'), $effectiveDays
    'Crashes / hangs'     = '{0}, {1}, {2} application(s) involved' -f $crashText, $hangText, $Apps.Count
    'First half'          = '{0} event(s) before {1}' -f $EarlierHalf, (Format-Invariant $midpoint 'yyyy-MM-dd HH:mm')
    'Second half'         = '{0} event(s) since {1}' -f $RecentHalf, (Format-Invariant $midpoint 'yyyy-MM-dd HH:mm')
    'Stability index now' = $(if ($null -ne $ssiLatest) { '{0} out of 10' -f $ssiLatest } else { 'not available' })
    'Stability average'   = $(if ($null -ne $ssiAverage) { '{0} (lowest day {1})' -f $ssiAverage, $ssiMin } else { 'not available' })
}
$trendBody = New-KeyValueTable -Pairs $trendPairs -Order @(
    'Window', 'Crashes / hangs', 'First half', 'Second half', 'Stability index now', 'Stability average')

$dailyDisplay = @(
    @($DailyRows) | Select-Object -Last 45 | ForEach-Object {
        $cell = '<div class="idx"><span>no sample</span></div>'
        if ($_.SystemStabilityIndex -ne '') {
            $value = [double]$_.SystemStabilityIndex
            $cls = ''
            if     ($value -lt 4) { $cls = 'danger' }
            elseif ($value -lt 7) { $cls = 'warn' }
            $cell = '<div class="idx"><span>{0:N2}</span>{1}</div>' -f `
                $value, (New-Bar -Percent (10.0 * $value) -Class $cls).Html
        }
        # An empty cell in a column of numbers reads as a zero; a day nobody
        # counted has to say so.
        $crashCell = $_.CrashCount
        if ([string]$crashCell -eq '') { $crashCell = 'not counted' }
        $hangCell = $_.HangCount
        if ([string]$hangCell -eq '') { $hangCell = 'not counted' }
        [pscustomobject]@{
            Date              = $_.Date
            'Stability index' = New-RawHtml $cell
            Crashes           = $crashCell
            Hangs             = $hangCell
        }
    }
)
$trendBody += New-DataTable -Rows $dailyDisplay -Empty 'No daily data for this window.'

Add-Section 'The machine as a whole' $trendBody `
    'The stability index is Windows'' own Reliability Monitor score, 0 to 10. It drops on a crash and climbs back over the days that follow, so a low number means recent trouble rather than old trouble.'

# 3. league table
$leagueRows = @(
    @($AppsRanked | Select-Object -First $Top) | ForEach-Object {
        $arrow = '&rarr; steady'
        if ($_.Trend -eq 'rising')  { $arrow = '&uarr; rising' }
        if ($_.Trend -eq 'falling') { $arrow = '&darr; falling' }
        [pscustomobject]@{
            Application = $_.Friendly
            Executable  = $_.ExeDisplay
            Crashes     = $_.Crashes
            Hangs       = $_.Hangs
            'First seen'= Format-Invariant $_.FirstSeen 'yyyy-MM-dd HH:mm'
            'Last seen' = Format-Invariant $_.LastSeen 'yyyy-MM-dd HH:mm'
            Trend       = New-RawHtml ('<span class="trend ' + $_.Trend + '">' + $arrow + '</span>')
        }
    }
)
$leagueNote = 'Counted from Application Error (1000) and Application Hang (1002) records only.'
if ($AppsRanked.Count -gt $Top) {
    $leagueNote += ' {0} further application(s) faulted less often and are not listed.' -f ($AppsRanked.Count - $Top)
}
Add-Section 'Which applications faulted' (New-DataTable -Rows $leagueRows -Empty 'No application crashed or hung in this window.') $leagueNote

# 4. per-app detail
$detailHtml = New-Object System.Text.StringBuilder
if ($AppsRanked.Count -eq 0) {
    [void]$detailHtml.Append('<p class="empty">Nothing to detail.</p>')
}
foreach ($a in @($AppsRanked | Select-Object -First $Top)) {
    [void]$detailHtml.Append('<div class="app">')
    [void]$detailHtml.Append(('<h3>{0} <span class="exe">{1}</span></h3>' -f (ConvertTo-HtmlText $a.Friendly), (ConvertTo-HtmlText $a.ExeDisplay)))

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('{0} crash(es) and {1} hang(s), {2} in the second half of the window.' -f $a.Crashes, $a.Hangs, $a.Recent))

    if ($a.CulpritModule) {
        $lines.Add(('{0} of {1} crashes fault in <strong>{2}</strong> - {3}.' -f `
            $a.CulpritHits, $a.Crashes, (ConvertTo-HtmlText $a.CulpritModule), (ConvertTo-HtmlText $a.CulpritVerdict)))
    }
    elseif ($a.Crashes -gt 0) {
        $lines.Add('No single module accounts for most of these crashes, so there is no obvious culprit to name.')
    }

    if ($a.IsHostProcess) {
        $pkgText = 'no package was recorded'
        if ($a.Package) { $pkgText = 'the package recorded alongside it was ' + (ConvertTo-HtmlText $a.Package) }
        $lines.Add(('This is a generic Windows host process - the code that actually failed was loaded into it, not written as it. For reference, {0}.' -f $pkgText))
    }
    elseif ($a.Package) {
        $lines.Add(('Installed as the Store package <code>{0}</code>.' -f (ConvertTo-HtmlText $a.Package)))
    }

    if (@($a.Codes).Count -gt 0) {
        $lines.Add(('Exception code(s): <code>{0}</code>.' -f (ConvertTo-HtmlText ((@($a.Codes)) -join ', '))))
    }
    if (@($a.DotNet).Count -gt 0) {
        $lines.Add(('The .NET runtime also logged: {0}.' -f (ConvertTo-HtmlText ((@($a.DotNet)) -join '; '))))
    }
    if (@($a.HangTypes).Count -gt 0) {
        $lines.Add(('Hang type(s) reported by Windows: {0}.' -f (ConvertTo-HtmlText ((@($a.HangTypes)) -join '; '))))
    }
    if (@($a.Buckets).Count -gt 0) {
        $lines.Add(('Error Reporting bucket(s): <code>{0}</code>.' -f (ConvertTo-HtmlText ((@($a.Buckets)) -join ', '))))
    }
    if (@($a.AllPaths).Count -gt 1) {
        # Two installs of the same executable share a league row; the paths are
        # the only thing that says they are not the same copy.
        $lines.Add(('Seen at {0} different paths: {1}.' -f @($a.AllPaths).Count, (ConvertTo-HtmlText ((@($a.AllPaths)) -join '  |  '))))
    }
    elseif ($a.ExePath) {
        $lines.Add(('Path: <code>{0}</code>.' -f (ConvertTo-HtmlText $a.ExePath)))
    }

    [void]$detailHtml.Append('<ul class="facts">')
    foreach ($l in $lines) { [void]$detailHtml.Append('<li>' + $l + '</li>') }
    [void]$detailHtml.Append('</ul>')

    $occurrences = @(
        @($a.Records | Sort-Object TimeCreated -Descending | Select-Object -First 12) | ForEach-Object {
            [pscustomobject]@{
                When   = Format-Invariant $_.TimeCreated
                What   = $_.Kind
                Module = $(if ($_.ModuleUnloaded -and $_.ModuleName) { $_.ModuleName + ' (unloaded)' } else { $_.ModuleName })
                Code   = $_.ExceptionCode
                Detail = $(if ($_.HangType) { $_.HangType } else { $_.DotNet })
            }
        }
    )
    [void]$detailHtml.Append((New-DataTable -Rows $occurrences -Empty 'No occurrences recorded.'))
    [void]$detailHtml.Append('</div>')
}
Add-Section 'What happened, application by application' $detailHtml.ToString()

# 5. data-source health
$healthRows = @(
    @($Health) | ForEach-Object {
        $state = 'read'
        if ($_.State -eq 'degraded') { $state = 'unavailable' }
        if ($_.State -eq 'skipped')  { $state = 'skipped' }
        [pscustomobject]@{
            Source = $_.Source
            State  = $state
            Notes  = $_.Detail
        }
    }
)
$healthNote = 'Nothing here changes the verdict. A source that could not be read costs the report a column, not its conclusion.'
if ($dotnetOrphans -gt 0) {
    $healthNote += ' {0} .NET runtime record(s) had no matching crash within five seconds and were left uncounted rather than counted twice.' -f $dotnetOrphans
}
Add-Section 'Where these numbers came from' (New-DataTable -Rows $healthRows) $healthNote

# ------------------------------------------------------------- assemble page
$css = @'
:root {
  --bg:#f6f7f9; --card:#ffffff; --ink:#1b1f24; --muted:#5b6572;
  --line:#e2e6ea; --accent:#2f6fed; --ok:#1a7f45; --warn:#a5670a; --crit:#b3261e;
  --bar:#dfe4ea; --barfill:#2f6fed;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg:#14171c; --card:#1c2027; --ink:#e6e9ee; --muted:#9aa4b2;
    --line:#2a2f38; --accent:#79a5ff; --ok:#5ed08c; --warn:#e5b567; --crit:#ff8a80;
    --bar:#2a2f38; --barfill:#79a5ff;
  }
}
* { box-sizing:border-box; }
body {
  margin:0; padding:32px 20px 64px; background:var(--bg); color:var(--ink);
  font:15px/1.55 "Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,sans-serif;
}
.wrap { max-width:1100px; margin:0 auto; }
h1 { font-size:26px; margin:0 0 4px; letter-spacing:-0.02em; }
.sub { color:var(--muted); margin:0 0 28px; font-size:14px; }
section {
  background:var(--card); border:1px solid var(--line); border-radius:12px;
  padding:20px 22px; margin-bottom:18px;
}
h2 { font-size:16px; margin:0 0 14px; letter-spacing:-0.01em; }
h3 { font-size:15px; margin:0 0 8px; letter-spacing:-0.01em; }
h3 .exe { color:var(--muted); font-weight:400; font-size:13px; margin-left:8px; }
.note { color:var(--muted); font-size:13px; margin:-6px 0 14px; }
.empty { color:var(--muted); font-style:italic; margin:0; }
code { font:12.5px/1.4 "Cascadia Mono",Consolas,monospace; background:var(--bar); padding:1px 5px; border-radius:4px; }
table { border-collapse:collapse; width:100%; font-size:14px; }
.kv th {
  text-align:left; font-weight:600; color:var(--muted); padding:5px 16px 5px 0;
  white-space:nowrap; vertical-align:top; width:170px; font-size:13px;
}
.kv td { padding:5px 0; word-break:break-word; }
.scroll { overflow-x:auto; }
.data th {
  text-align:left; font-size:12px; text-transform:uppercase; letter-spacing:0.04em;
  color:var(--muted); border-bottom:1px solid var(--line); padding:8px 12px 8px 0; white-space:nowrap;
}
.data td { padding:8px 12px 8px 0; border-bottom:1px solid var(--line); vertical-align:middle; }
.data tr:last-child td { border-bottom:none; }
.bar { background:var(--bar); border-radius:999px; height:8px; width:110px; overflow:hidden; }
.bar span { display:block; height:100%; background:var(--barfill); }
.bar.warn span { background:var(--warn); }
.bar.danger span { background:var(--crit); }
.idx { display:flex; align-items:center; gap:10px; }
.idx span { min-width:60px; color:var(--muted); font-size:13px; }
.trend { font-weight:600; font-size:13px; white-space:nowrap; }
.trend.rising { color:var(--crit); }
.trend.falling { color:var(--ok); }
.trend.steady { color:var(--muted); }
.verdict {
  background:var(--card); border:1px solid var(--line); border-radius:12px;
  padding:18px 22px; margin-bottom:18px;
}
.verdict.calm  { border-left:4px solid var(--ok); }
.verdict.noted { border-left:4px solid var(--warn); }
.verdict.worse { border-left:4px solid var(--crit); }
.verdict .tag {
  font-size:11px; text-transform:uppercase; letter-spacing:0.06em; font-weight:700;
  padding:2px 8px; border-radius:999px; background:var(--bar); color:var(--muted);
}
.verdict.calm .tag  { color:var(--ok); }
.verdict.noted .tag { color:var(--warn); }
.verdict.worse .tag { color:var(--crit); }
.verdict p { margin:12px 0 0; }
.verdict ul { margin:10px 0 0; padding-left:20px; color:var(--muted); font-size:13px; }
.app { border-top:1px solid var(--line); padding:18px 0 4px; }
.app:first-child { border-top:none; padding-top:0; }
ul.facts { margin:0 0 12px; padding-left:20px; }
ul.facts li { margin:3px 0; }
footer { color:var(--muted); font-size:12px; text-align:center; margin-top:28px; }
'@

$verdictClass = 'noted'
$verdictTag   = 'Faults recorded'
if ($ExitCode -eq 0) { $verdictClass = 'calm';  $verdictTag = 'All clear' }
if ($EventTotal -eq 0 -and -not $countersOk) { $verdictTag = 'Counters unreadable' }
if ($ExitCode -eq 2) { $verdictClass = 'worse'; $verdictTag = 'Getting worse' }

$verdictHtml = New-Object System.Text.StringBuilder
[void]$verdictHtml.Append('<div class="verdict ' + $verdictClass + '">')
[void]$verdictHtml.Append('<span class="tag">' + (ConvertTo-HtmlText $verdictTag) + '</span>')
[void]$verdictHtml.Append('<p><strong>' + (ConvertTo-HtmlText $Verdict) + '</strong></p>')
[void]$verdictHtml.Append('<p>' + (ConvertTo-HtmlText $headline) + '</p>')
if ($worseReasons.Count -gt 0) {
    [void]$verdictHtml.Append('<ul>')
    foreach ($reason in $worseReasons) { [void]$verdictHtml.Append('<li>' + (ConvertTo-HtmlText $reason) + '</li>') }
    [void]$verdictHtml.Append('</ul>')
}
[void]$verdictHtml.Append('</div>')

$footerNote = 'Generated by Get-CrashTrend.ps1 - read-only, nothing on this machine was changed.'
if ($windowEmpty) {
    $footerNote = ("The Application log holds no records at or before {0}, so this window is empty. " -f `
        (Format-Invariant $Now 'yyyy-MM-dd HH:mm')) + $footerNote
}
elseif ($windowTruncated) {
    $footerNote = ("The Application log only reaches back to {0}, so '-Days {1}' is really {2} days. " -f `
        (Format-Invariant $effectiveFloor 'yyyy-MM-dd'), $Days, $effectiveDays) + $footerNote
}

$html = New-Object System.Collections.Generic.List[string]
$html.Add('<!doctype html>')
$html.Add('<html lang="en"><head><meta charset="utf-8">')
$html.Add('<meta name="viewport" content="width=device-width,initial-scale=1">')
$html.Add(('<title>Crash trend - {0}</title>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add('<style>' + $css + '</style></head><body><div class="wrap">')
$html.Add(('<h1>{0}</h1>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add(('<p class="sub">Application crashes and hangs, {0} to {1} &middot; report written {2}</p>' -f `
    (ConvertTo-HtmlText (Format-Invariant $effectiveFloor 'yyyy-MM-dd')),
    (ConvertTo-HtmlText (Format-Invariant $Now 'yyyy-MM-dd')),
    (ConvertTo-HtmlText (Format-Invariant $RunStart))))
$html.Add($verdictHtml.ToString())
foreach ($s in $Sections) { $html.Add($s) }
$html.Add('<footer>' + (ConvertTo-HtmlText $footerNote) + '</footer>')
$html.Add('</div></body></html>')

$htmlOk = $true
try {
    ($html -join [Environment]::NewLine) | Set-Content -LiteralPath $ReportFile -Encoding utf8
}
catch {
    $htmlOk = $false
    Write-Step "! could not write $ReportFile : $($_.Exception.Message)"
}

# ------------------------------------------------------------------ wrap up
Write-Line ''
if ($htmlOk) { Write-Line "Report  : $ReportFile" }
if ($csvOk) {
    Write-Line "Daily   : $DailyCsv"
    Write-Line "History : $RunHistoryCsv"
}
Write-Line ''

if ($Open -and $htmlOk) {
    # A machine with no handler registered for .html throws here, which would
    # otherwise lose the exit code the caller actually asked for.
    try { Start-Process -FilePath $ReportFile -ErrorAction Stop }
    catch { Write-Step "! could not open the report: $($_.Exception.Message)" }
}

exit $ExitCode
