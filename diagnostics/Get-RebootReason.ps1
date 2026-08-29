<#
.SYNOPSIS
    Why this PC restarted, shut down, bluescreened, slept and woke up - a plain
    English timeline of the last -Days days, plus the list of things that are
    allowed to wake it on their own.

.DESCRIPTION
    Read-only. Everything comes from the System event log, powercfg, the task
    scheduler and two registry values; nothing on this machine is changed.

    Windows already knows why it restarted. It just does not say so anywhere a
    person can read. This assembles the answer from the eleven places it is
    scattered across and writes one HTML page: every boot session in the window,
    how each one ended, and a sentence for each that a non-technical person can
    act on.

    The second section answers the question people actually ask - "why does it
    turn itself on at three in the morning" - by naming the devices that are
    armed to wake it and the scheduled tasks that are allowed to, and saying
    where in Windows each one is turned off. It never turns anything off itself.

    Verdicts come from codes and typed fields only, never from the rendered
    event text: that text is localized, and on a non-English Windows every
    string match would silently stop matching.

.PARAMETER Days
    How far back the timeline reaches. Default 14. The System log on a small
    SSD rarely retains more than a few weeks, so a large number here is a
    request, not a promise - the report says where the log actually begins.

.PARAMETER OutputDir
    Where the report and the history CSV land.
    Defaults to $env:LOCALAPPDATA\GetRebootReason. If that cannot be written
    to, the tool falls back to %TEMP%\GetRebootReason and says so.

.PARAMETER HistoryLimit
    Rows kept in RebootReason_history.csv. Default 5000; 0 keeps everything.
    When the column set changes the old file is archived, never discarded.

.PARAMETER MaxEventsPerQuery
    Per-provider ceiling on how many events are read. Default 2000. A log that
    hits the ceiling is reported as truncated rather than queried again.

.PARAMETER Now
    The instant the window is measured back from. Defaults to the current time;
    exists so a run can be pinned to a fixed moment and reproduced.

.PARAMETER Quiet
    Suppress console output; still writes the files and sets the exit code.
    One line survives it, in every tool in this collection: if -OutputDir
    cannot be written to, the fallback folder is announced even under -Quiet,
    because a scheduled caller looking for its report in the folder it asked
    for would otherwise find nothing and be told nothing.

.PARAMETER Open
    Open the generated HTML report when finished.

.OUTPUTS
    Exit code 0 = report written, no bluescreens or unexpected power losses.
    1 = report written, at least one bluescreen or unexpected power loss.
    2 = report written but a data source was unavailable (wins over 0 and 1).
    3 = the System event log could not be read at all; no report written.
    4 = neither -OutputDir nor the %TEMP% fallback could be written to.
    5 = unexpected error.

.EXAMPLE
    .\Get-RebootReason.ps1 -Open

.EXAMPLE
    .\Get-RebootReason.ps1 -Days 45

.EXAMPLE
    # Sessions that did not end cleanly, across every run so far
    Import-Csv "$env:LOCALAPPDATA\GetRebootReason\RebootReason_history.csv" |
        Where-Object { $_.EndVerdict -notlike 'Clean*' }
#>
[CmdletBinding()]
param(
    # How far back the timeline reaches. Event logs on small SSDs rarely
    # retain more than a few weeks of System log anyway.
    [ValidateRange(1, 365)]
    [int]$Days = 14,

    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'GetRebootReason'),

    # Rolling CSV row cap; 0 = keep all rows.
    [ValidateRange(0, 100000)]
    [int]$HistoryLimit = 5000,

    # Per-provider Get-WinEvent cap so a churning log cannot stall the run.
    [ValidateRange(50, 20000)]
    [int]$MaxEventsPerQuery = 2000,

    # Injectable clock; nothing below reads the wall clock directly. DontShow
    # keeps it out of the GUI launcher's generated form and out of tab
    # completion: it is a test seam, and the only correct value is the default.
    [Parameter(DontShow)]
    [datetime]$Now = (Get-Date),

    [switch]$Quiet,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ScriptVersion = '1.0.0'
$ToolName      = 'GetRebootReason'
$Invariant     = [System.Globalization.CultureInfo]::InvariantCulture

# A shutdown burst is several User32 1074 records seconds apart, and a session
# end is only ever attributed to 1074s inside this many seconds of it.
$ShutdownBurstSeconds = 120
# Off and back within this long reads as a restart rather than a shutdown.
$RestartGapSeconds    = 180
# A crash record is written by the *next* boot; beyond this it is not about the
# session that just ended.
$CrashEvidenceSeconds = 1800
# Update events this far either side of a system-initiated shutdown are treated
# as the reason for it. Inferred correlation, not something Windows states.
$UpdateWindowSeconds  = 1800
# Two boot anchors closer together than this are the same boot seen twice.
$BootDedupeSeconds    = 120
# Modern Standby cycles number in the dozens per day. Only sleeps at least this
# long reach the timeline; the rest are counted and summarised.
$SleepNoticeMinutes   = 20

# Anything that goes wrong after this point still has to produce an exit code a
# scheduled caller can read, rather than a stack trace and a 1.
trap {
    if (-not $Quiet) {
        $Host.UI.WriteErrorLine("Get-RebootReason: unexpected error - $($_.Exception.Message)")
    }
    exit 5
}

# ------------------------------------------------------------------- plumbing
function Get-Prop {
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

function Add-HistoryRow {
    <#
        Export-Csv -Append refuses outright when the file's header does not
        match the object being appended, so the first run after this script
        gains or loses a column would throw away the whole reading at the last
        step. The stale file is archived under a dated name instead, which
        keeps the earlier readings rather than discarding them.
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
            $archive = Join-Path (Split-Path -Parent $Path) `
                (('{0}_upto_{1}.csv' -f [System.IO.Path]::GetFileNameWithoutExtension($Path), (Get-Date -Format 'yyyy-MM-dd_HHmmss')))
            Move-Item -LiteralPath $Path -Destination $archive -Force
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

function Invoke-WithTimeout {
    <#
        Runs a scriptblock in its own runspace and gives up on it after
        -Seconds.

        Get-ScheduledTask talks to the Schedule service, which can be wedged,
        and the cmdlet takes no timeout of its own. A diagnostic that hangs
        forever is worse than one that says it could not read something.

        An abandoned runspace is asked to stop but never waited on: waiting is
        precisely the thing that would hang. It goes away with the process.
    #>
    param([scriptblock]$Script, [int]$Seconds = 30)

    $ps = $null
    try {
        $ps = [powershell]::Create()
        [void]$ps.AddScript($Script)
        $handle = $ps.BeginInvoke()

        if ($handle.AsyncWaitHandle.WaitOne($Seconds * 1000)) {
            $output = @()
            try { $output = @($ps.EndInvoke($handle)) } catch { }
            try { $ps.Dispose() } catch { }
            return [pscustomobject]@{ Completed = $true; Output = $output }
        }

        try { [void]$ps.BeginStop($null, $null) } catch { }
        return [pscustomobject]@{ Completed = $false; Output = @() }
    }
    catch {
        if ($ps) { try { $ps.Dispose() } catch { } }
        return [pscustomobject]@{ Completed = $true; Output = @() }
    }
}

function Write-Line {
    param([string]$Message, [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Message -ForegroundColor $Colour }
    else         { Write-Host $Message }
}

# --------------------------------------------------------------- time helpers
function ConvertTo-Utc {
    # An event's TimeCreated is local with a Local kind; everything downstream
    # sorts and subtracts in UTC so a DST boundary inside the window cannot
    # reorder the timeline.
    param($Value)
    if ($null -eq $Value) { return $null }
    $dt = [datetime]$Value
    if ($dt.Kind -eq [System.DateTimeKind]::Utc) { return $dt }
    if ($dt.Kind -eq [System.DateTimeKind]::Unspecified) {
        return [datetime]::SpecifyKind($dt, [System.DateTimeKind]::Utc)
    }
    $dt.ToUniversalTime()
}

function ConvertTo-UtcFromIso {
    <#
        Kernel-General writes ISO-8601 with a Z suffix, which is the one date
        format in this whole report that is not at the mercy of the machine's
        regional settings. Parsed with the invariant culture so an en-IN or
        Thai-calendar box reads it the same way.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $parsed = [datetime]::MinValue
    # RoundtripKind alone: it honours the trailing Z and cannot be combined
    # with AdjustToUniversal, which throws rather than being ignored.
    $styles = [System.Globalization.DateTimeStyles]::RoundtripKind
    if ([datetime]::TryParse($Text, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) {
        if ($parsed.Kind -eq [System.DateTimeKind]::Local) { return $parsed.ToUniversalTime() }
        # These fields are documented UTC, so a missing Z is a formatting quirk
        # rather than a local time.
        return [datetime]::SpecifyKind($parsed, [System.DateTimeKind]::Utc)
    }
    $null
}

function Format-UtcStamp {
    param($Utc)
    if ($null -eq $Utc) { return '' }
    ([datetime]$Utc).ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-LocalStamp {
    # Displayed times are local, formatted invariantly: this machine's regional
    # format is dd-MM-yyyy, which has already produced ambiguous dates in this
    # collection.
    param($Utc, [string]$Format = 'yyyy-MM-dd HH:mm:ss')
    if ($null -eq $Utc) { return '' }
    ([datetime]$Utc).ToLocalTime().ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-Duration {
    <#
        Every component is truncated, never rounded. [int] on a double rounds
        in PowerShell rather than truncating, and '{0:N0}' rounds too, so
        1h 54m 18s used to render as "2h 54m" and 59m 40s as "60 min" - an
        hour of invented uptime in half the rows of a table that prints the two
        timestamps either side of it.

        Each branch reads the component the branch above has already bounded,
        which cannot round by construction.
    #>
    param($Span)
    if ($null -eq $Span) { return '' }
    $ts = [timespan]$Span
    if ($ts.TotalSeconds -lt 0)  { return '' }
    if ($ts.TotalMinutes -lt 1)  { return ('{0} s' -f $ts.Seconds) }
    if ($ts.TotalHours -lt 1)    { return ('{0} min' -f $ts.Minutes) }
    if ($ts.TotalDays -lt 1)     { return ('{0}h {1:00}m' -f $ts.Hours, $ts.Minutes) }
    '{0}d {1}h {2:00}m' -f $ts.Days, $ts.Hours, $ts.Minutes
}

# ---------------------------------------------------------------- event access
function Get-EventFields {
    <#
        EventData read by name, through XPath that ignores the schema
        namespace.

        Two reasons never to walk this with the PowerShell XML adapter: an
        empty element such as <Data Name="WakeSourceText"/> has no '#text'
        child, so reading it throws under strict mode rather than returning an
        empty string; and on <Data Name="Reason"> the adapter's .Name is the
        attribute, not the element name, which is a coincidence rather than a
        contract. GetAttribute and InnerText are neither.
    #>
    param($Event)
    $fields = @{}
    if ($null -eq $Event) { return $fields }
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($Event.ToXml())
        foreach ($node in @($doc.SelectNodes("/*[local-name()='Event']/*[local-name()='EventData']/*[local-name()='Data']"))) {
            $name = $node.GetAttribute('Name')
            if ([string]::IsNullOrEmpty($name)) { continue }
            $text = ''
            if ($null -ne $node.InnerText) { $text = [string]$node.InnerText }
            $fields[$name] = $text
        }
    }
    catch { }
    $fields
}

function Get-Field {
    # Hashtable indexing returns $null for an absent key even under strict
    # mode, but dot access on the same key throws - so every read goes through
    # here rather than through whichever form was quicker to type.
    param([hashtable]$Fields, [string]$Name, $Default = '')
    if ($null -eq $Fields -or -not $Fields.ContainsKey($Name)) { return $Default }
    $v = $Fields[$Name]
    if ($null -eq $v) { return $Default }
    $v
}

function Get-EventProp {
    <#
        Positional read for the EventLog provider, whose 6005-6013 records have
        no field names at all. Layouts are stable only per provider and id, so
        the count is checked before every index rather than trusted.
    #>
    param($Event, [int]$Index, $Default = $null)
    if ($null -eq $Event) { return $Default }
    $props = Get-Prop $Event 'Properties'
    if ($null -eq $props -or $props.Count -le $Index) { return $Default }
    $v = $props[$Index].Value
    if ($null -eq $v) { return $Default }
    $v
}

function ConvertTo-LongOrNull {
    param($Text)
    if ($null -eq $Text) { return $null }
    $parsed = [long]0
    if ([long]::TryParse(([string]$Text).Trim(), [ref]$parsed)) { return $parsed }
    $null
}

function ConvertTo-UInt64OrNull {
    # 507 carries counters that have wrapped past Int64 - NonResiliencyTimeInUs
    # comes back as 18446744073637813636 on this machine - so anything that may
    # be one is parsed as unsigned and never as [long].
    param($Text)
    if ($null -eq $Text) { return $null }
    $parsed = [uint64]0
    if ([uint64]::TryParse(([string]$Text).Trim(), [ref]$parsed)) { return $parsed }
    $null
}

function ConvertFrom-HexOrNull {
    param($Text)
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return $null }
    $s = ([string]$Text).Trim()
    if ($s.StartsWith('0x') -or $s.StartsWith('0X')) { $s = $s.Substring(2) }
    try { return [Convert]::ToInt64($s, 16) }
    catch { return $null }
}

function ConvertFrom-SystemTimeBlob {
    <#
        EventLog 6008 carries the crash instant as two SYSTEMTIME structures in
        one Byte[64]: bytes 0-15 local, 16-31 UTC. The rendered date and time
        strings sitting next to it in the same record are localized and salted
        with U+200E direction marks, so the bytes are the only honest source.

        Short blobs turn up on older builds; the length check degrades to "we
        do not know when" rather than throwing away the whole event.
    #>
    param($Bytes, [int]$Offset = 16)
    if ($null -eq $Bytes) { return $null }
    $b = $Bytes -as [byte[]]
    if ($null -eq $b -or $b.Length -lt ($Offset + 16)) { return $null }
    try {
        $year  = [int][BitConverter]::ToUInt16($b, $Offset)
        $month = [int][BitConverter]::ToUInt16($b, $Offset + 2)
        $day   = [int][BitConverter]::ToUInt16($b, $Offset + 6)
        $hour  = [int][BitConverter]::ToUInt16($b, $Offset + 8)
        $min   = [int][BitConverter]::ToUInt16($b, $Offset + 10)
        $sec   = [int][BitConverter]::ToUInt16($b, $Offset + 12)
        $ms    = [int][BitConverter]::ToUInt16($b, $Offset + 14)

        if ($year -lt 1980 -or $year -gt 2200) { return $null }
        if ($month -lt 1 -or $month -gt 12)    { return $null }
        if ($day -lt 1 -or $day -gt 31)        { return $null }
        if ($hour -gt 23 -or $min -gt 59 -or $sec -gt 59 -or $ms -gt 999) { return $null }

        return (New-Object datetime ($year, $month, $day, $hour, $min, $sec, $ms, ([System.DateTimeKind]::Utc)))
    }
    catch { return $null }
}

$AccountCache = @{}
function Test-RealUserAccount {
    <#
        Decides whether a shutdown was asked for by a person.

        The account name in a 1074 is localized - SYSTEM is NT-AUTORITAET\SYSTEM
        on German Windows - so it is resolved to a SID and compared against the
        three well-known service accounts instead of being string-matched. An
        account that no longer resolves falls back to "not one of ours".
    #>
    param([string]$Account)
    if ([string]::IsNullOrWhiteSpace($Account)) { return $false }
    if ($AccountCache.ContainsKey($Account)) { return [bool]$AccountCache[$Account] }

    $result = $false
    try {
        $sid = (New-Object System.Security.Principal.NTAccount($Account)).Translate(
            [System.Security.Principal.SecurityIdentifier]).Value
        # LocalSystem, LocalService, NetworkService. Never a person.
        $result = -not ($sid -eq 'S-1-5-18' -or $sid -eq 'S-1-5-19' -or $sid -eq 'S-1-5-20')
    }
    catch {
        # Deleted account, or a domain that cannot be reached. A name carrying
        # this machine's own name is still most likely a local person.
        $lower = $Account.ToLowerInvariant()
        $me    = ''
        if ($env:COMPUTERNAME) { $me = $env:COMPUTERNAME.ToLowerInvariant() }
        $result = ($me -ne '' -and $lower.StartsWith($me + '\'))
    }
    $AccountCache[$Account] = $result
    $result
}

# ------------------------------------------------------------ bugcheck naming
# Only the stop codes a home or small-office machine actually hits, each with
# the sentence a non-technical owner needs. Anything not listed is still shown
# by its hex code - an unknown code is reported as unknown, not guessed at.
$StopCodeNames = @{
    '0x0000000A' = @('IRQL_NOT_LESS_OR_EQUAL',              'A driver touched memory it should not have. Usually a driver bug.')
    '0x00000019' = @('BAD_POOL_HEADER',                     'Kernel memory was corrupted, most often by a driver.')
    '0x0000001A' = @('MEMORY_MANAGEMENT',                    'Memory handling went wrong. Faulty RAM is a common cause - worth a memory test.')
    '0x0000001E' = @('KMODE_EXCEPTION_NOT_HANDLED',          'A driver hit an error it had no plan for.')
    '0x00000024' = @('NTFS_FILE_SYSTEM',                     'A problem reading or writing the NTFS file system. Check the disk.')
    '0x0000003B' = @('SYSTEM_SERVICE_EXCEPTION',             'A system call failed inside the kernel, usually via a driver.')
    '0x00000050' = @('PAGE_FAULT_IN_NONPAGED_AREA',          'Something asked for memory that was not there. Faulty RAM or a bad driver.')
    '0x0000007A' = @('KERNEL_DATA_INPAGE_ERROR',             'Windows could not read a page back from disk. Check the drive health.')
    '0x0000007B' = @('INACCESSIBLE_BOOT_DEVICE',             'Windows could not reach the disk it boots from.')
    '0x0000007E' = @('SYSTEM_THREAD_EXCEPTION_NOT_HANDLED',  'A system thread crashed, nearly always inside a driver.')
    '0x0000007F' = @('UNEXPECTED_KERNEL_MODE_TRAP',          'The CPU reported a fault the kernel could not recover from. Often hardware.')
    '0x0000009F' = @('DRIVER_POWER_STATE_FAILURE',           'A driver did not finish going to sleep or waking up in time.')
    '0x000000C2' = @('BAD_POOL_CALLER',                      'A driver misused kernel memory.')
    '0x000000D1' = @('DRIVER_IRQL_NOT_LESS_OR_EQUAL',        'A driver touched memory at the wrong time. Update or roll back drivers.')
    '0x000000EF' = @('CRITICAL_PROCESS_DIED',                'A process Windows cannot run without stopped.')
    '0x000000F4' = @('CRITICAL_OBJECT_TERMINATION',          'A critical system object ended unexpectedly.')
    '0x00000101' = @('CLOCK_WATCHDOG_TIMEOUT',               'A CPU core stopped responding to the others.')
    '0x00000116' = @('VIDEO_TDR_ERROR',                      'The graphics driver stopped responding and could not be reset.')
    '0x00000117' = @('VIDEO_TDR_TIMEOUT_DETECTED',           'The graphics driver stopped responding.')
    '0x00000124' = @('WHEA_UNCORRECTABLE_ERROR',             'The hardware itself reported an error it could not correct. Take this one seriously.')
    '0x00000133' = @('DPC_WATCHDOG_VIOLATION',               'Something in the kernel ran far too long without yielding. Usually a storage or chipset driver.')
    '0x00000139' = @('KERNEL_SECURITY_CHECK_FAILURE',        'A kernel safety check failed - corrupted data structures.')
    '0x0000019C' = @('WIN32K_POWER_WATCHDOG_TIMEOUT',        'The display stack did not finish a sleep or wake transition in time.')
    '0x000001CA' = @('SYNTHETIC_WATCHDOG_TIMEOUT',           'A watchdog timer fired because part of the system stopped responding.')
}

function Get-StopCodeHex {
    param($Value)
    if ($null -eq $Value) { return '' }
    '0x{0:x8}' -f ([long]$Value)
}

function Get-StopCodeName {
    param([string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { return '' }
    $key = $Hex.ToLowerInvariant()
    foreach ($k in $StopCodeNames.Keys) {
        if ($k.ToLowerInvariant() -eq $key) { return [string]$StopCodeNames[$k][0] }
    }
    ''
}

function Get-StopCodeMeaning {
    param([string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { return '' }
    $key = $Hex.ToLowerInvariant()
    foreach ($k in $StopCodeNames.Keys) {
        if ($k.ToLowerInvariant() -eq $key) { return [string]$StopCodeNames[$k][1] }
    }
    'This stop code is not in the short list this tool carries. Searching for it by name will turn up what it means.'
}

# ------------------------------------------------------------------ HTML bits
function ConvertTo-HtmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-RawHtml {
    <#
        Marks a cell whose value is already HTML and must not be escaped again.
        A typed marker rather than a string prefix, so no genuine value - a
        device name, a task path, a stop code - can ever be mistaken for one.
    #>
    param([string]$Html)
    [pscustomobject]@{ PSTypeName = 'Snapshot.RawHtml'; Html = $Html }
}

function New-KeyValueTable {
    param([hashtable]$Pairs, [string[]]$Order)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table class="kv">')
    foreach ($k in $Order) {
        if (-not $Pairs.ContainsKey($k)) { continue }
        $v = $Pairs[$k]
        if ($null -eq $v -or [string]$v -eq '') { $v = '-' }
        [void]$sb.Append(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlText $k), (ConvertTo-HtmlText $v)))
    }
    [void]$sb.Append('</table>')
    $sb.ToString()
}

function New-DataTable {
    param([psobject[]]$Rows, [string[]]$Columns, [string]$EmptyText = 'Nothing to show.')

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return ('<p class="empty">{0}</p>' -f (ConvertTo-HtmlText $EmptyText))
    }
    if (-not $Columns) {
        $Columns = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="scroll"><table class="data"><thead><tr>')
    foreach ($c in $Columns) { [void]$sb.Append('<th>' + (ConvertTo-HtmlText $c) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')

    foreach ($row in $Rows) {
        $cls = [string](Get-Prop $row '_RowClass' '')
        if ($cls) { [void]$sb.Append('<tr class="' + $cls + '">') }
        else      { [void]$sb.Append('<tr>') }
        foreach ($c in $Columns) {
            $val = Get-Prop $row $c ''
            if ($null -ne $val -and $val.PSObject.TypeNames -contains 'Snapshot.RawHtml') {
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

$Attention = New-Object System.Collections.Generic.List[psobject]
$Sections  = New-Object System.Collections.Generic.List[string]

function Add-Attention {
    param(
        [ValidateSet('Critical','Warning','Info')][string]$Level,
        [string]$Title,
        [string]$Detail
    )
    $Attention.Add([pscustomobject]@{ Level = $Level; Title = $Title; Detail = $Detail })
}

function Add-Section {
    param([string]$Title, [string]$Body, [string]$Note = '')
    $noteHtml = ''
    if ($Note) { $noteHtml = '<p class="note">' + (ConvertTo-HtmlText $Note) + '</p>' }
    $Sections.Add('<section><h2>' + (ConvertTo-HtmlText $Title) + '</h2>' + $noteHtml + $Body + '</section>')
}

# ============================================================ 1. environment
$OutputDirResolved = ''
try {
    $OutputDirResolved = Initialize-OutputDir -Preferred $OutputDir -ToolName $ToolName
}
catch {
    if (-not $Quiet) { $Host.UI.WriteErrorLine("Get-RebootReason: $($_.Exception.Message)") }
    exit 4
}
$OutputDir = $OutputDirResolved

$ReportFile = Join-Path $OutputDir 'RebootReason.html'
$HistoryCsv = Join-Path $OutputDir 'RebootReason_history.csv'

# -Now given as a string binds with Kind = Unspecified. Get-WinEvent reads such
# a value as local time, while ConvertTo-Utc reads it as UTC - so a run pinned
# with -Now '2026-08-29 10:00:00' harvested up to 10:00 local but stamped every
# displayed time, and the CSV's RunId key, an offset away from it. It is a
# wall-clock instant everywhere it is used, so it is stamped as one here.
if ($Now.Kind -eq [System.DateTimeKind]::Unspecified) {
    $Now = [datetime]::SpecifyKind($Now, [System.DateTimeKind]::Local)
}

$NowUtc      = ConvertTo-Utc $Now
$WindowStart = $Now.AddDays(-$Days)
$WindowStartUtc = ConvertTo-Utc $WindowStart
$RunId       = Format-UtcStamp $NowUtc

$DegradedNotes = New-Object System.Collections.Generic.List[string]
$TruncNotes    = New-Object System.Collections.Generic.List[string]
$Degraded      = $false

Write-Line ''
Write-Line "=== Why this PC restarted : last $Days day(s) ==="
Write-Line ''

# ============================================================== 2. harvest
function Get-SystemEvents {
    <#
        One wrapped query per provider.

        Get-WinEvent throws NoMatchingEventsFound when nothing matches, which
        under $ErrorActionPreference = 'Stop' is indistinguishable at the call
        site from "the log is unreadable" - and those two answers deserve very
        different reports. The error id is matched by prefix; its Message is
        localized and would stop matching on a non-English Windows.

        Returns an object rather than an array because a function returning @()
        unrolls to $null on the way out, and $null.Count throws under strict
        mode - which is how this exact wrapper failed the first time.
    #>
    param([string]$Provider, [int[]]$Ids, [datetime]$Since, [datetime]$Until, [int]$Max)

    # EndTime as well as StartTime: with only a start, the query runs to the
    # wall clock and a -Now pinned to last week still returns everything since,
    # which makes a "reproducible" run neither reproducible nor a window.
    $filter = @{ LogName = 'System'; ProviderName = $Provider; StartTime = $Since; EndTime = $Until }
    if ($null -ne $Ids -and $Ids.Count -gt 0) { $filter['Id'] = $Ids }

    try {
        $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents $Max -ErrorAction Stop)
        return [pscustomobject]@{ Ok = $true; Events = $events; Error = ''; Truncated = ($events.Count -ge $Max) }
    }
    catch {
        if ([string]$_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') {
            return [pscustomobject]@{ Ok = $true; Events = @(); Error = ''; Truncated = $false }
        }
        return [pscustomobject]@{ Ok = $false; Events = @(); Error = [string]$_.Exception.Message; Truncated = $false }
    }
}

$Records = New-Object System.Collections.Generic.List[psobject]

function Add-Record {
    param([string]$Kind, $TimeUtc, [int]$EventId, [hashtable]$Data)
    if ($null -eq $TimeUtc) { return }
    if ($null -eq $Data) { $Data = @{} }
    $Records.Add([pscustomobject]@{
        Kind    = $Kind
        TimeUtc = [datetime]$TimeUtc
        EventId = $EventId
        Data    = $Data
    })
}

function Select-Records {
    # Callers wrap this in @(): a function returning an empty array unrolls to
    # $null on the way out, and $null.Count throws under strict mode.
    param([string[]]$Kind)
    @($Records | Where-Object { $Kind -contains $_.Kind } | Sort-Object TimeUtc)
}

Write-Line '  Reading the System log...'

# Nine queries, batched by provider. Each result is normalised into a flat
# record immediately, so every phase after this one is free of event objects,
# XML and property-index arithmetic - and can be fed hand-built fixtures.
$sources = @(
    @{ Key = 'User32';       Provider = 'User32';                                     Ids = @(1074) }
    @{ Key = 'EventLog';     Provider = 'EventLog';                                   Ids = @(6005, 6006, 6008, 6009, 6013) }
    @{ Key = 'KernelPower';  Provider = 'Microsoft-Windows-Kernel-Power';              Ids = @(41, 42, 107, 506, 507) }
    @{ Key = 'Wer';          Provider = 'Microsoft-Windows-WER-SystemErrorReporting';  Ids = @(1001) }
    @{ Key = 'Update';       Provider = 'Microsoft-Windows-WindowsUpdateClient';       Ids = @(19, 20, 22, 43, 44) }
    @{ Key = 'WakeReport';   Provider = 'Microsoft-Windows-Power-Troubleshooter';      Ids = @(1) }
    @{ Key = 'KernelBoot';   Provider = 'Microsoft-Windows-Kernel-Boot';               Ids = @(27) }
    @{ Key = 'KernelGeneral';Provider = 'Microsoft-Windows-Kernel-General';            Ids = @(12, 13) }
    @{ Key = 'Whea';         Provider = 'Microsoft-Windows-WHEA-Logger';               Ids = @(17, 18, 19, 20, 46, 47) }
)

$harvest = @{}
$failedSources = 0
$firstError = ''
foreach ($src in $sources) {
    $res = Get-SystemEvents -Provider $src.Provider -Ids $src.Ids -Since $WindowStart -Until $Now -Max $MaxEventsPerQuery
    $harvest[$src.Key] = $res

    if (-not $res.Ok) {
        $failedSources++
        if (-not $firstError) { $firstError = $res.Error }
        $DegradedNotes.Add(("{0} could not be read ({1}). Anything that source would have explained is missing from this report." -f $src.Provider, $res.Error))
    }
    elseif ($res.Truncated) {
        $TruncNotes.Add(("{0} hit the {1}-event ceiling, so the oldest part of the window is missing from that source." -f $src.Provider, $MaxEventsPerQuery))
    }
}

# Every provider failing is not nine coincidences - it is the System log itself
# being unreadable, and there is no timeline to write. One or two failing is a
# thinner report, which is still worth having.
if ($failedSources -eq $sources.Count) {
    if (-not $Quiet) {
        $Host.UI.WriteErrorLine("Get-RebootReason: the System event log could not be read - $firstError")
    }
    exit 3
}
if ($failedSources -gt 0) { $Degraded = $true }

# --- User32 1074: who asked for a shutdown -----------------------------------
foreach ($e in @($harvest['User32'].Events)) {
    $process = [string](Get-EventProp $e 0 '')
    # The process name carries an optional " (HOSTNAME)" suffix that is noise
    # in a report and would break any grouping done on the name.
    $suffix = ' ({0})' -f $env:COMPUTERNAME
    if ($process.EndsWith($suffix)) { $process = $process.Substring(0, $process.Length - $suffix.Length) }
    $process = $process.Trim()

    Add-Record 'ShutdownRequest' (ConvertTo-Utc $e.TimeCreated) 1074 @{
        Process     = $process
        ReasonTitle = [string](Get-EventProp $e 2 '')
        ReasonCode  = ConvertFrom-HexOrNull (Get-EventProp $e 3 '')
        TypeText    = [string](Get-EventProp $e 4 '')
        Comment     = [string](Get-EventProp $e 5 '')
        Account     = [string](Get-EventProp $e 6 '')
    }
}

# --- EventLog 6005/6006/6008/6013: unnamed, positional -----------------------
foreach ($e in @($harvest['EventLog'].Events)) {
    $when = ConvertTo-Utc $e.TimeCreated
    switch ([int]$e.Id) {
        6005 { Add-Record 'LogStart' $when 6005 @{} }
        6006 { Add-Record 'LogStop'  $when 6006 @{} }
        6008 {
            Add-Record 'DirtyShutdown' $when 6008 @{
                CrashTimeUtc = ConvertFrom-SystemTimeBlob (Get-EventProp $e 7 $null) 16
            }
        }
        6013 {
            Add-Record 'UptimeReport' $when 6013 @{
                UptimeSeconds = ConvertTo-LongOrNull (Get-EventProp $e 4 $null)
            }
        }
    }
}

# --- Kernel-Power ------------------------------------------------------------
foreach ($e in @($harvest['KernelPower'].Events)) {
    $when   = ConvertTo-Utc $e.TimeCreated
    $fields = Get-EventFields $e
    switch ([int]$e.Id) {
        41 {
            # BugcheckCode here is decimal; the same number reaches WER as hex.
            # Reading either one in the other's base is the classic way to
            # report 0x133 as 0x307 and confuse everyone.
            Add-Record 'UncleanBoot' $when 41 @{
                BugcheckCode         = ConvertTo-LongOrNull (Get-Field $fields 'BugcheckCode' '0')
                BugcheckParam1       = [string](Get-Field $fields 'BugcheckParameter1' '')
                BugcheckParam2       = [string](Get-Field $fields 'BugcheckParameter2' '')
                BugcheckParam3       = [string](Get-Field $fields 'BugcheckParameter3' '')
                BugcheckParam4       = [string](Get-Field $fields 'BugcheckParameter4' '')
                PowerButtonTimestamp = ConvertTo-UInt64OrNull (Get-Field $fields 'PowerButtonTimestamp' '0')
            }
        }
        42 {
            Add-Record 'SleepEnter' $when 42 @{
                Family      = 'classic'
                Reason      = [string](Get-Field $fields 'Reason' '')
                TargetState = [string](Get-Field $fields 'TargetState' '')
            }
        }
        107 {
            Add-Record 'SleepExit' $when 107 @{
                Family            = 'classic'
                WakeFromState     = [string](Get-Field $fields 'WakeFromState' '')
                ProgrammedWakeAc  = ConvertTo-UtcFromIso (Get-Field $fields 'ProgrammedWakeTimeAc' '')
                ProgrammedWakeDc  = ConvertTo-UtcFromIso (Get-Field $fields 'ProgrammedWakeTimeDc' '')
                WakeRequesterAc   = [string](Get-Field $fields 'WakeRequesterTypeAc' '')
                WakeRequesterDc   = [string](Get-Field $fields 'WakeRequesterTypeDc' '')
            }
        }
        506 {
            Add-Record 'SleepEnter' $when 506 @{
                Family = 'modern'
                Reason = [string](Get-Field $fields 'Reason' '')
                BootId = [string](Get-Field $fields 'BootId' '')
            }
        }
        507 {
            $durUs = ConvertTo-UInt64OrNull (Get-Field $fields 'DurationInUs' '0')
            $seconds = $null
            # Durations past a year are a wrapped unsigned counter, not a fact.
            if ($null -ne $durUs -and $durUs -lt 31536000000000) { $seconds = [double]$durUs / 1000000.0 }
            Add-Record 'SleepExit' $when 507 @{
                Family         = 'modern'
                Reason         = [string](Get-Field $fields 'Reason' '')
                DurationSecond = $seconds
                SleepEntered   = ([string](Get-Field $fields 'SleepEntered' '')).ToLowerInvariant()
                BootId         = [string](Get-Field $fields 'BootId' '')
            }
        }
    }
}

# --- WER 1001: the stop code in the form people quote ------------------------
foreach ($e in @($harvest['Wer'].Events)) {
    $text = [string](Get-EventProp $e 0 '')
    $hex  = ''
    # The first token is the stop code; the four bugcheck parameters after it
    # are not useful to a person and are not carried forward.
    if ($text -match '^(0x[0-9a-fA-F]{8})') { $hex = $Matches[1].ToLowerInvariant() }
    Add-Record 'Bluescreen' (ConvertTo-Utc $e.TimeCreated) 1001 @{
        StopCodeHex  = $hex
        StopCodeText = $text
        MinidumpPath = [string](Get-EventProp $e 1 '')
    }
}

# --- Windows Update ----------------------------------------------------------
foreach ($e in @($harvest['Update'].Events)) {
    $id = [int]$e.Id
    # Id 20 alone puts an Int32 errorCode at [0] and pushes the title to [1].
    # Reading [0] as the title on that one id is the layout bug this whole
    # provider is famous for.
    $titleIndex = 0
    $errorCode  = $null
    if ($id -eq 20) {
        $titleIndex = 1
        $errorCode  = Get-EventProp $e 0 $null
    }
    $title = [string](Get-EventProp $e $titleIndex '')
    $kb = ''
    if ($title -match '(KB\d{6,})') { $kb = $Matches[1] }

    Add-Record 'Update' (ConvertTo-Utc $e.TimeCreated) $id @{
        Title     = $title
        Kb        = $kb
        ErrorCode = $errorCode
    }
}

# --- Power-Troubleshooter 1: the only record that names a wake source --------
foreach ($e in @($harvest['WakeReport'].Events)) {
    $fields = Get-EventFields $e
    Add-Record 'WakeReport' (ConvertTo-Utc $e.TimeCreated) 1 @{
        SleepTimeUtc    = ConvertTo-UtcFromIso (Get-Field $fields 'SleepTime' '')
        WakeTimeUtc     = ConvertTo-UtcFromIso (Get-Field $fields 'WakeTime' '')
        WakeSourceType  = ConvertTo-LongOrNull (Get-Field $fields 'WakeSourceType' '0')
        WakeSourceText  = [string](Get-Field $fields 'WakeSourceText' '')
        WakeTimerOwner  = [string](Get-Field $fields 'WakeTimerOwner' '')
        WakeTimerContext= [string](Get-Field $fields 'WakeTimerContext' '')
    }
}

# --- Kernel-Boot 27: cold boot, fast-startup resume or hibernate resume ------
foreach ($e in @($harvest['KernelBoot'].Events)) {
    $fields = Get-EventFields $e
    Add-Record 'BootType' (ConvertTo-Utc $e.TimeCreated) 27 @{
        BootType = ConvertTo-LongOrNull (Get-Field $fields 'BootType' '')
    }
}

# --- Kernel-General 12/13: typed UTC, the cleanest anchors in the log --------
foreach ($e in @($harvest['KernelGeneral'].Events)) {
    $fields = Get-EventFields $e
    switch ([int]$e.Id) {
        12 {
            $t = ConvertTo-UtcFromIso (Get-Field $fields 'StartTime' '')
            if ($null -eq $t) { $t = ConvertTo-Utc $e.TimeCreated }
            Add-Record 'SessionStart' $t 12 @{ BootMode = [string](Get-Field $fields 'BootMode' '') }
        }
        13 {
            $t = ConvertTo-UtcFromIso (Get-Field $fields 'StopTime' '')
            if ($null -eq $t) { $t = ConvertTo-Utc $e.TimeCreated }
            Add-Record 'SessionStop' $t 13 @{}
        }
    }
}

# --- WHEA: annotation only. Absent is the healthy answer. --------------------
foreach ($e in @($harvest['Whea'].Events)) {
    Add-Record 'HardwareError' (ConvertTo-Utc $e.TimeCreated) ([int]$e.Id) @{}
}

$logStartRecords = @(Select-Records @('LogStart'))

# Where the System log actually begins. This has to be asked of the log itself
# and not of the harvest above: every query up there is bounded by StartTime,
# so the oldest record they return is never older than the window, and a quiet
# stretch at the start of it - a laptop that was simply switched off overnight
# - is indistinguishable from a log that has been trimmed. Reading that as
# truncation told this machine its log stopped at 45 days when it holds 99.
$LogFloorUtc = $null
try {
    $oldest = @(Get-WinEvent -LogName 'System' -Oldest -MaxEvents 1 -ErrorAction Stop)
    if ($oldest.Count -gt 0) { $LogFloorUtc = ConvertTo-Utc $oldest[0].TimeCreated }
}
catch { }

Write-Line ("  {0} events normalised from {1} source(s)." -f $Records.Count, $sources.Count)

# ========================================================== 3. boot sessions
# Kernel-General 12 is primary: its StartTime is a typed UTC value rather than
# a log-write timestamp. EventLog 6005 is the fallback for a machine or an
# older build where 12 is absent, and is skipped where a 12 already covers it.
$sessionStarts = @(Select-Records @('SessionStart'))
$bootAnchors = New-Object System.Collections.Generic.List[psobject]
foreach ($r in $sessionStarts) {
    $bootAnchors.Add([pscustomobject]@{ TimeUtc = $r.TimeUtc; Source = 'Kernel-General 12' })
}
foreach ($r in $logStartRecords) {
    $covered = $false
    foreach ($a in $bootAnchors) {
        if ([math]::Abs(($r.TimeUtc - $a.TimeUtc).TotalSeconds) -le $BootDedupeSeconds) { $covered = $true; break }
    }
    if (-not $covered) {
        $bootAnchors.Add([pscustomobject]@{ TimeUtc = $r.TimeUtc; Source = 'EventLog 6005' })
    }
}
$bootAnchors = @($bootAnchors | Sort-Object TimeUtc)

$bootTypeRecords = @(Select-Records @('BootType'))
$sleepEnters = @(Select-Records @('SleepEnter'))
$sleepExits = @(Select-Records @('SleepExit'))
$wakeReports = @(Select-Records @('WakeReport'))
$shutdownReqs = @(Select-Records @('ShutdownRequest'))
$sessionStops = @(Select-Records @('SessionStop'))
$logStops = @(Select-Records @('LogStop'))
$bluescreens = @(Select-Records @('Bluescreen'))
$uncleanBoots = @(Select-Records @('UncleanBoot'))
$dirtyShutdowns = @(Select-Records @('DirtyShutdown'))
$updates = @(Select-Records @('Update'))
$hardwareErrors = @(Select-Records @('HardwareError'))

$Sessions = New-Object System.Collections.Generic.List[psobject]

# A crash record is written by the boot that follows the crash, so the oldest
# one in the window can describe a session that started before the window did.
# That session is represented rather than invented: it has no boot time.
$Sessions.Add([pscustomobject]@{
    BootTimeUtc  = $null
    BootSource   = ''
    BootType     = $null
    EndTimeUtc   = $null
    NextBootUtc  = $null
    IsPreWindow  = $true
    Evidence     = New-Object System.Collections.Generic.List[psobject]
    Verdict      = ''
    Detail       = ''
    Category     = ''
    StopCodeHex  = ''
    Kb           = ''
    Wakes        = 0
    Unexplained  = 0
    ShortSleeps  = 0
})

for ($i = 0; $i -lt $bootAnchors.Count; $i++) {
    $boot = $bootAnchors[$i].TimeUtc
    $next = $null
    if ($i -lt ($bootAnchors.Count - 1)) { $next = $bootAnchors[$i + 1].TimeUtc }

    # A Kernel-Boot 27 close to the boot instant describes that boot. Hibernate
    # resumes also raise a 27 but no Kernel-General 12, so a loose window here
    # would label a cold boot with a resume's boot type.
    $bootType = $null
    foreach ($bt in $bootTypeRecords) {
        if ([math]::Abs(($bt.TimeUtc - $boot).TotalSeconds) -le $BootDedupeSeconds) {
            $bootType = Get-Field $bt.Data 'BootType' $null
            break
        }
    }

    $Sessions.Add([pscustomobject]@{
        BootTimeUtc  = $boot
        BootSource   = $bootAnchors[$i].Source
        BootType     = $bootType
        EndTimeUtc   = $null
        NextBootUtc  = $next
        IsPreWindow  = $false
        Evidence     = New-Object System.Collections.Generic.List[psobject]
        Verdict      = ''
        Detail       = ''
        Category     = ''
        StopCodeHex  = ''
        Kb           = ''
        Wakes        = 0
        Unexplained  = 0
        ShortSleeps  = 0
    })
}

function Get-SessionIndexAt {
    <#
        Which session was running at a given instant. -1 means the instant is
        older than every boot in view and there is no pre-window placeholder to
        put it in - which happens once the placeholder has been dropped, so the
        answer has to be "nowhere" rather than "session zero".
    #>
    param($TimeUtc)
    $idx = -1
    for ($i = 0; $i -lt $Sessions.Count; $i++) {
        if ($null -eq $Sessions[$i].BootTimeUtc) { $idx = $i; continue }
        if ($TimeUtc -ge $Sessions[$i].BootTimeUtc) { $idx = $i } else { break }
    }
    $idx
}

# Clean stop evidence belongs to the session it happened in.
foreach ($stop in $sessionStops) {
    $i = Get-SessionIndexAt $stop.TimeUtc
    if ($i -le 0) { continue }
    $Sessions[$i].EndTimeUtc = $stop.TimeUtc
    $Sessions[$i].Evidence.Add($stop)
}
foreach ($stop in $logStops) {
    $i = Get-SessionIndexAt $stop.TimeUtc
    if ($i -le 0) { continue }
    if ($null -eq $Sessions[$i].EndTimeUtc) { $Sessions[$i].EndTimeUtc = $stop.TimeUtc }
    $Sessions[$i].Evidence.Add($stop)
}

# Crash evidence is logged by the boot that follows the crash, so it belongs to
# the session before the one it appears in.
foreach ($rec in @($uncleanBoots + $dirtyShutdowns + $bluescreens)) {
    $i = Get-SessionIndexAt $rec.TimeUtc
    if ($i -lt 0) { continue }
    # Never $host: that is the PowerShell host object, and shadowing it breaks
    # the error reporting this script does through $Host.UI.
    $holder = $Sessions[$i]
    if (-not $holder.IsPreWindow -and $null -ne $holder.BootTimeUtc) {
        if (($rec.TimeUtc - $holder.BootTimeUtc).TotalSeconds -gt $CrashEvidenceSeconds) {
            # Far too late in the session to be about the previous one. Leave
            # it where it is rather than mis-attributing it.
            $holder.Evidence.Add($rec)
            continue
        }
    }
    $target = $i - 1
    if ($target -lt 0) { $target = 0 }
    $Sessions[$target].Evidence.Add($rec)
}

# ==================================================== 4. end-of-session verdict
function Get-UpdateNear {
    <#
        Inferred correlation, not a fact Windows records: a system-initiated
        restart with an update installed either side of it is reported as an
        update restart. Defender's signature package (KB2267602) installs
        several times a day and would otherwise claim every restart, so it is
        never allowed to be the answer.
    #>
    param($TimeUtc)
    $best = $null
    $bestGap = [double]::MaxValue
    foreach ($u in $updates) {
        if (@(19, 22, 43) -notcontains $u.EventId) { continue }
        $title = [string](Get-Field $u.Data 'Title' '')
        if ($title -like '*KB2267602*') { continue }
        $gap = [math]::Abs(($u.TimeUtc - $TimeUtc).TotalSeconds)
        if ($gap -le $UpdateWindowSeconds -and $gap -lt $bestGap) { $bestGap = $gap; $best = $u }
    }
    $best
}

function Get-ShutdownBurst {
    # One human shutdown emits two to four 1074s seconds apart. Grouping only
    # those just before this session's stop keeps two real shutdowns minutes
    # apart from merging into one.
    param($EndUtc)
    if ($null -eq $EndUtc) { return @() }
    @($shutdownReqs | Where-Object {
        $delta = ($EndUtc - $_.TimeUtc).TotalSeconds
        $delta -ge -5 -and $delta -le $ShutdownBurstSeconds
    } | Sort-Object TimeUtc)
}

function Resolve-SessionEnd {
    <#
        The verdict engine. Consumes only normalised records, so every branch
        below can be exercised with hand-built fixtures on any machine.

        Priority is deliberate: a bluescreen outranks a power button, which
        outranks a bare power loss, which outranks the clean-shutdown path.
        Windows writes several of these for one event and the most specific
        one is the true answer.
    #>
    param($Session)

    $ev        = @($Session.Evidence)
    $screens   = @($ev | Where-Object { $_.Kind -eq 'Bluescreen' })
    $unclean   = @($ev | Where-Object { $_.Kind -eq 'UncleanBoot' })
    $dirty     = @($ev | Where-Object { $_.Kind -eq 'DirtyShutdown' })
    $cleanStop = @($ev | Where-Object { $_.Kind -eq 'SessionStop' -or $_.Kind -eq 'LogStop' })

    # (a) Bluescreen. WER carries the stop code as hex; Kernel-Power 41 carries
    #     the same number in decimal. Either one is proof.
    $stopHex = ''
    $dump    = ''
    foreach ($s in $screens) {
        $h = [string](Get-Field $s.Data 'StopCodeHex' '')
        if ($h) { $stopHex = $h }
        $d = [string](Get-Field $s.Data 'MinidumpPath' '')
        if ($d) { $dump = $d }
    }
    if (-not $stopHex) {
        foreach ($u in $unclean) {
            $code = Get-Field $u.Data 'BugcheckCode' $null
            if ($null -ne $code -and [long]$code -ne 0) { $stopHex = Get-StopCodeHex $code; break }
        }
    }
    if ($stopHex) {
        $name = Get-StopCodeName $stopHex
        $label = $stopHex
        if ($name) { $label = '{0} {1}' -f $stopHex, $name }
        $detail = 'It bluescreened ({0}). {1}' -f $label, (Get-StopCodeMeaning $stopHex)
        if ($dump) { $detail = '{0} Crash dump: {1}' -f $detail, $dump }
        return [pscustomobject]@{
            Verdict = 'Bluescreen'; Detail = $detail; Category = 'Bluescreen'
            StopCodeHex = $stopHex; Kb = ''
        }
    }

    # (b) Power button held down. The timestamp is a FILETIME; only whether it
    #     is set matters here.
    foreach ($u in $unclean) {
        $pbt = Get-Field $u.Data 'PowerButtonTimestamp' $null
        if ($null -ne $pbt -and [uint64]$pbt -ne 0) {
            return [pscustomobject]@{
                Verdict = 'Power button'
                Detail  = 'Someone held the power button down until it switched off. Windows had no chance to shut down properly.'
                Category = 'PowerButton'; StopCodeHex = ''; Kb = ''
            }
        }
    }

    # (c) Power simply stopped. A 6008 knows when, to the second.
    if ($unclean.Count -gt 0 -or $dirty.Count -gt 0) {
        $when = $null
        foreach ($d in $dirty) {
            $t = Get-Field $d.Data 'CrashTimeUtc' $null
            if ($null -ne $t) { $when = $t; break }
        }
        $detail = 'Power was cut, the battery ran flat, or it was forced off. Windows did not get to shut down.'
        if (-not $script:HasBattery) {
            $detail = 'Power was cut or it was forced off. Windows did not get to shut down.'
        }
        if ($null -ne $when) {
            $detail = '{0} It stopped at {1}.' -f $detail, (Format-LocalStamp $when)
        }
        return [pscustomobject]@{
            Verdict = 'Unexpected power loss'; Detail = $detail; Category = 'PowerLoss'
            StopCodeHex = ''; Kb = ''
        }
    }

    # (d) A clean stop. Who asked for it, and was it off for long?
    if ($cleanStop.Count -gt 0 -and $null -ne $Session.EndTimeUtc) {
        $burst = @(Get-ShutdownBurst $Session.EndTimeUtc)

        $wentBackOn = $false
        if ($null -ne $Session.NextBootUtc) {
            $gap = ($Session.NextBootUtc - $Session.EndTimeUtc).TotalSeconds
            $wentBackOn = ($gap -ge 0 -and $gap -le $RestartGapSeconds)
        }
        # Decided by how long it stayed off, never by the 1074's type string -
        # that string is localized and says "power off" for both.
        $word = 'shut it down'
        if ($wentBackOn) { $word = 'restarted it' }

        $person = $null
        foreach ($b in $burst) {
            if (Test-RealUserAccount ([string](Get-Field $b.Data 'Account' ''))) { $person = $b; break }
        }

        if ($null -ne $person) {
            $proc = [string](Get-Field $person.Data 'Process' '')
            $detail = 'You {0}.' -f $word
            if ($proc) { $detail = '{0} The request came from {1}.' -f $detail, $proc }
            $verdict = 'Clean - by you'
            if ($wentBackOn) { $verdict = 'Clean restart - by you' }
            return [pscustomobject]@{ Verdict = $verdict; Detail = $detail; Category = 'CleanUser'; StopCodeHex = ''; Kb = '' }
        }

        if ($burst.Count -gt 0) {
            $upd = Get-UpdateNear $Session.EndTimeUtc
            if ($null -ne $upd) {
                $kb = [string](Get-Field $upd.Data 'Kb' '')
                $title = [string](Get-Field $upd.Data 'Title' '')
                if ($kb) {
                    $detail = 'Windows Update {0} to finish installing {1}.' -f $word, $kb
                }
                else {
                    $detail = 'Windows finished installing an update around this restart, so it very likely {0}.' -f $word
                }
                if ($title) { $detail = '{0} ({1})' -f $detail, $title }
                return [pscustomobject]@{ Verdict = 'Clean - Windows Update'; Detail = $detail; Category = 'CleanUpdate'; StopCodeHex = ''; Kb = $kb }
            }
            return [pscustomobject]@{
                Verdict = 'Clean - by Windows'
                Detail  = 'Windows itself asked for this one - a scheduled restart, a maintenance task, or an installer that needed one. No update was recorded near it.'
                Category = 'CleanSystem'; StopCodeHex = ''; Kb = ''
            }
        }

        return [pscustomobject]@{
            Verdict = 'Clean - unattributed'
            Detail  = 'It shut down properly, but nothing recorded who asked for it. A lid close, a low battery hibernate or a power-plan action all look like this.'
            Category = 'CleanUnknown'; StopCodeHex = ''; Kb = ''
        }
    }

    # No end evidence at all. The newest session is simply still running.
    if ($null -eq $Session.NextBootUtc -and -not $Session.IsPreWindow) {
        return [pscustomobject]@{
            Verdict = 'Still running'; Detail = 'This is the session you are in now.'
            Category = 'Running'; StopCodeHex = ''; Kb = ''
        }
    }

    [pscustomobject]@{
        Verdict = 'Unknown'
        Detail  = 'Nothing in the log says how this one ended. The records that would have said so have most likely been overwritten.'
        Category = 'Unknown'; StopCodeHex = ''; Kb = ''
    }
}

$script:HasBattery = $false
try { $script:HasBattery = (@(Get-CimInstance -ClassName Win32_Battery -OperationTimeoutSec 20 -ErrorAction Stop).Count -gt 0) }
catch { }

foreach ($s in $Sessions) {
    $r = Resolve-SessionEnd $s
    $s.Verdict     = $r.Verdict
    $s.Detail      = $r.Detail
    $s.Category    = $r.Category
    $s.StopCodeHex = $r.StopCodeHex
    $s.Kb          = $r.Kb

    # A session that ended badly did not end when the event log service
    # stopped. A stray EventLog 6006 partway through a session - the log
    # service being restarted, a servicing operation - is picked up as clean
    # stop evidence above and would otherwise date the crash to it: the one
    # bluescreen on this machine has a 6006 at 13:32 and crashed at 15:31.
    # So for these three verdicts the end instant is re-derived from the crash
    # itself: the 6008's decoded time, and failing that the boot that followed,
    # which is within seconds of a bugcheck. This also fills in the end time
    # and uptime of every unclean session, which used to be left blank.
    if (@('Bluescreen', 'PowerLoss', 'PowerButton') -contains $s.Category) {
        $crashAt = $null
        foreach ($ev in @($s.Evidence)) {
            $ct = Get-Field $ev.Data 'CrashTimeUtc' $null
            if ($null -ne $ct -and ($null -eq $crashAt -or $ct -lt $crashAt)) { $crashAt = $ct }
        }
        if ($null -eq $crashAt) { $crashAt = $s.NextBootUtc }
        # An instant before the boot is describing some other session.
        if ($null -ne $crashAt -and $null -ne $s.BootTimeUtc -and $crashAt -lt $s.BootTimeUtc) {
            $crashAt = $null
        }
        $s.EndTimeUtc = $crashAt
    }

    # Fast startup turns a shutdown into a hibernate, which is why a machine
    # that was "shut down" can come back with everything still broken.
    if ($s.Category -like 'Clean*' -and $null -ne $s.NextBootUtc) {
        foreach ($bt in $bootTypeRecords) {
            if ([math]::Abs(($bt.TimeUtc - $s.NextBootUtc).TotalSeconds) -le $BootDedupeSeconds) {
                $t = Get-Field $bt.Data 'BootType' $null
                if ($null -ne $t -and [long]$t -eq 1) {
                    $s.Detail = '{0} Fast startup was on, so Windows hibernated rather than really shutting down.' -f $s.Detail
                }
                break
            }
        }
    }
}

# Drop the placeholder if nothing before the window needed explaining.
if ($Sessions.Count -gt 0 -and $Sessions[0].IsPreWindow -and @($Sessions[0].Evidence).Count -eq 0) {
    $Sessions.RemoveAt(0)
}

# ================================================ 5. sleep and wake interleave
# Both event families are live on this hardware, so a single sleep can raise a
# 42 and a 506 half a minute apart. Modern Standby is preferred where the two
# overlap: it is the transition the machine actually performed.
$mergedSleeps = New-Object System.Collections.Generic.List[psobject]
foreach ($s in $sleepEnters) {
    if ((Get-Field $s.Data 'Family' '') -eq 'classic') {
        $dupe = $false
        foreach ($m in $sleepEnters) {
            if ((Get-Field $m.Data 'Family' '') -ne 'modern') { continue }
            if ([math]::Abs(($m.TimeUtc - $s.TimeUtc).TotalSeconds) -le 60) { $dupe = $true; break }
        }
        if ($dupe) { continue }
    }
    $mergedSleeps.Add($s)
}

# A wake is counted where Windows wrote a wake report for it. A classic 107
# without one is still a wake and is counted too; a 507 is a Modern Standby
# screen-off cycle and is not, or the count would run to thousands.
$mergedWakes = New-Object System.Collections.Generic.List[psobject]
foreach ($w in $wakeReports) { $mergedWakes.Add($w) }
foreach ($x in $sleepExits) {
    if ((Get-Field $x.Data 'Family' '') -ne 'classic') { continue }
    $covered = $false
    foreach ($w in $wakeReports) {
        if ([math]::Abs(($w.TimeUtc - $x.TimeUtc).TotalSeconds) -le 120) { $covered = $true; break }
    }
    if (-not $covered) { $mergedWakes.Add($x) }
}
$mergedWakes = @($mergedWakes | Sort-Object TimeUtc)

# WakeSourceType 0 and 7 are what this machine actually produces; 1, 3 and 5
# are from the documented enumeration and have not been seen here. Anything
# else is reported by number rather than guessed at.
function Get-WakeExplanation {
    param($Record)
    if ($Record.EventId -ne 1) {
        return [pscustomobject]@{ Text = 'It woke up. Windows did not write a wake report for this one.'; Explained = $false }
    }
    $type  = Get-Field $Record.Data 'WakeSourceType' $null
    $text  = [string](Get-Field $Record.Data 'WakeSourceText' '')
    $owner = [string](Get-Field $Record.Data 'WakeTimerOwner' '')

    if ($text)  { return [pscustomobject]@{ Text = ('Woken by a device: {0}.' -f $text); Explained = $true } }
    if ($owner) { return [pscustomobject]@{ Text = ('Woken by a scheduled wake timer belonging to {0}.' -f $owner); Explained = $true } }

    if ($null -ne $type) {
        switch ([long]$type) {
            0 { return [pscustomobject]@{ Text = 'Windows did not record what woke it. This is the usual answer, and it is not a fault.'; Explained = $false } }
            1 { return [pscustomobject]@{ Text = 'Woken by the power button.'; Explained = $true } }
            3 { return [pscustomobject]@{ Text = 'Woken by a wake timer, but the owner was not recorded.'; Explained = $true } }
            5 { return [pscustomobject]@{ Text = 'Woken by a device, but the name was not recorded.'; Explained = $true } }
            7 { return [pscustomobject]@{ Text = 'Woken by a timer set by a driver that uses the older wake interface. Windows records no name for these.'; Explained = $false } }
            default { return [pscustomobject]@{ Text = ('Woken by source type {0}, which this tool has no name for.' -f $type); Explained = $false } }
        }
    }
    [pscustomobject]@{ Text = 'It woke up; the record does not say what from.'; Explained = $false }
}

$WakeCount = 0
$WakeExplained = 0
foreach ($w in $mergedWakes) {
    $WakeCount++
    if ((Get-WakeExplanation $w).Explained) { $WakeExplained++ }
    $i = Get-SessionIndexAt $w.TimeUtc
    if ($i -ge 0 -and $i -lt $Sessions.Count) {
        $Sessions[$i].Wakes++
        if (-not (Get-WakeExplanation $w).Explained) { $Sessions[$i].Unexplained++ }
    }
}

# Long sleeps reach the timeline; the short Modern Standby cycles are counted
# against their session and summarised instead.
$NotableSleeps = New-Object System.Collections.Generic.List[psobject]
$ShortSleepTotal = 0
foreach ($x in $sleepExits) {
    $secs = Get-Field $x.Data 'DurationSecond' $null
    if ($null -eq $secs) { continue }
    if ([double]$secs -ge ($SleepNoticeMinutes * 60)) { $NotableSleeps.Add($x) }
    else {
        $ShortSleepTotal++
        $i = Get-SessionIndexAt $x.TimeUtc
        if ($i -ge 0 -and $i -lt $Sessions.Count) { $Sessions[$i].ShortSleeps++ }
    }
}

# ========================================================= 6. wake forensics
Write-Line '  Reading wake settings...'

function Invoke-Powercfg {
    <#
        $ErrorActionPreference = 'Stop' plus a redirected stderr turns anything
        powercfg writes to the error stream into a terminating
        NativeCommandError - verified here with both 2>&1 and 2>$null. Relaxing
        the preference for the duration of the call keeps the text, and
        $LASTEXITCODE stays the signal for whether it worked.
    #>
    param([string[]]$Arguments)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out  = & powercfg.exe @Arguments 2>&1
        $code = $LASTEXITCODE
        $lines = @()
        foreach ($line in @($out)) {
            $s = [string]$line
            if (-not [string]::IsNullOrWhiteSpace($s)) { $lines += $s.Trim() }
        }
        return [pscustomobject]@{ Code = $code; Lines = $lines }
    }
    catch {
        return [pscustomobject]@{ Code = -1; Lines = @() }
    }
    finally {
        $ErrorActionPreference = $previous
    }
}

$wakeArmed = @()
$pc = Invoke-Powercfg @('/devicequery', 'wake_armed')
if ($pc.Code -eq 0) {
    # "NONE" is powercfg's own answer for an empty list and is not a device.
    $wakeArmed = @($pc.Lines | Where-Object { $_.ToUpperInvariant() -ne 'NONE' })
}
else {
    $Degraded = $true
    $DegradedNotes.Add('powercfg could not list the devices that are allowed to wake this PC.')
}

$sleepStates = @()
$pcA = Invoke-Powercfg @('/a')
if ($pcA.Code -eq 0) { $sleepStates = @($pcA.Lines) }

$lastWake = @()
$pcW = Invoke-Powercfg @('/lastwake')
if ($pcW.Code -eq 0) { $lastWake = @($pcW.Lines) }

# The only thing in this report that needs administrator rights. Expected to
# fail on every target machine, so it is a note rather than a degraded source.
$wakeTimers = @()
$wakeTimersDenied = $false
$pcT = Invoke-Powercfg @('/waketimers')
if ($pcT.Code -eq 0) { $wakeTimers = @($pcT.Lines) } else { $wakeTimersDenied = $true }

$scheduledWake = @()
$taskSourceOk  = $true
$taskResult = Invoke-WithTimeout -Seconds 30 -Script {
    try {
        Get-ScheduledTask -ErrorAction Stop | ForEach-Object {
            $task = $_
            $wake = $false
            try { $wake = [bool]$task.Settings.WakeToRun } catch { }
            if (-not $wake) { return }
            $next = $null
            try { $next = (Get-ScheduledTaskInfo -InputObject $task -ErrorAction Stop).NextRunTime } catch { }
            [pscustomobject]@{
                TaskPath = [string]$task.TaskPath
                TaskName = [string]$task.TaskName
                State    = [string]$task.State
                NextRun  = $next
            }
        }
    }
    catch { }
}
if (-not $taskResult.Completed) {
    $taskSourceOk = $false
    $Degraded = $true
    $DegradedNotes.Add('The Task Scheduler did not answer within 30 seconds, so scheduled tasks allowed to wake this PC are not listed.')
}
else {
    $scheduledWake = @($taskResult.Output)
    if ($scheduledWake.Count -eq 0) {
        # An empty result is ambiguous: no such tasks, or no ScheduledTasks
        # module. The module is present on every supported build, so absence is
        # treated as the honest "none found" unless the run timed out.
        $scheduledWake = @()
    }
}

$hiberboot = $null
try {
    $k = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -ErrorAction Stop
    $hiberboot = Get-Prop $k 'HiberbootEnabled'
}
catch { }

# Whether hibernate is switched on, as a number rather than as a line of
# powercfg's localized prose. "powercfg /h off" writes HibernateEnabled = 0; a
# machine that has never been told either way carries no such value at all and
# follows HibernateEnabledDefault - which is exactly this machine, where
# hibernate works and HibernateEnabled is absent, so reading only the first of
# the two reports a working feature as missing.
$hibernateFlag = $null
try {
    $k = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -ErrorAction Stop
    $hibernateFlag = Get-Prop $k 'HibernateEnabled'
    if ($null -eq $hibernateFlag) { $hibernateFlag = Get-Prop $k 'HibernateEnabledDefault' }
}
catch { }

$lastCleanShutdownUtc = $null
try {
    $k = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Windows' -ErrorAction Stop
    $raw = Get-Prop $k 'ShutdownTime'
    $bytes = $raw -as [byte[]]
    # An 8-byte little-endian FILETIME. Any other length is not one.
    if ($null -ne $bytes -and $bytes.Length -eq 8) {
        $ft = [BitConverter]::ToInt64($bytes, 0)
        if ($ft -gt 0) { $lastCleanShutdownUtc = [datetime]::FromFileTimeUtc($ft) }
    }
}
catch { }

# ================================================================ 7. tallies
$RealSessions = @($Sessions | Where-Object { -not $_.IsPreWindow })

$BluescreenCount  = @($Sessions | Where-Object { $_.Category -eq 'Bluescreen' }).Count
# Held-down power button and cut power both end a session dirtily, but they are
# different problems with different fixes and are never added together.
$PowerLossCount   = @($Sessions | Where-Object { $_.Category -eq 'PowerLoss' }).Count
$PowerButtonCount = @($Sessions | Where-Object { $_.Category -eq 'PowerButton' }).Count
$ByUserCount      = @($Sessions | Where-Object { $_.Category -eq 'CleanUser' }).Count
$ByUpdateCount    = @($Sessions | Where-Object { $_.Category -eq 'CleanUpdate' }).Count
$CycleCount       = @($Sessions | Where-Object { $_.Category -ne 'Running' }).Count

$HibernateResumes = 0
foreach ($bt in $bootTypeRecords) {
    $t = Get-Field $bt.Data 'BootType' $null
    if ($null -ne $t -and [long]$t -eq 2) { $HibernateResumes++ }
}

$SummaryLine = ('In the last {0} day(s): {1} time(s) this PC went off and came back ({2} by you, {3} by Windows Update), {4} bluescreen(s), {5} sudden power loss(es), {6} wake(s) ({7} of them explained).' -f `
    $Days, $CycleCount, $ByUserCount, $ByUpdateCount, $BluescreenCount, $PowerLossCount, $WakeCount, $WakeExplained)
if ($PowerButtonCount -gt 0) {
    $SummaryLine = '{0} A further {1} ended with the power button held down.' -f $SummaryLine, $PowerButtonCount
}
# A source that hit the ceiling was read from the newest event backwards, so
# the oldest part of its window is missing and every count above it is a floor
# rather than a total. Saying so here puts the caveat on the console summary
# and the top of the page at once.
if ($TruncNotes.Count -gt 0) {
    $SummaryLine = '{0} Those are the least it can have been: at least one source hit the {1}-event ceiling, so the oldest part of the window was not read.' -f $SummaryLine, $MaxEventsPerQuery
}

if ($BluescreenCount -gt 0) {
    $codes = @($Sessions | Where-Object { $_.Category -eq 'Bluescreen' -and $_.StopCodeHex } |
        ForEach-Object { $_.StopCodeHex } | Select-Object -Unique)
    Add-Attention 'Critical' 'Bluescreens' ("{0} crash(es) in this window. Stop code(s): {1}." -f $BluescreenCount, ($codes -join ', '))
}
if ($PowerLossCount -gt 0) {
    Add-Attention 'Warning' 'Unexpected power loss' ("{0} time(s) the power stopped without Windows shutting down. Repeated ones point at the battery, the charger or a loose plug." -f $PowerLossCount)
}
if ($PowerButtonCount -gt 0) {
    Add-Attention 'Info' 'Power button held down' ("{0} time(s) the machine was switched off by holding the power button. That is normally a sign it had already stopped responding." -f $PowerButtonCount)
}
if ($hardwareErrors.Count -gt 0) {
    Add-Attention 'Critical' 'Hardware errors logged' ("{0} WHEA hardware-error event(s) in this window. These come from the hardware itself, not from Windows." -f $hardwareErrors.Count)
}
$armedTasks = @($scheduledWake | Where-Object { ([string]$_.State).ToLowerInvariant() -ne 'disabled' })
if ($armedTasks.Count -gt 0) {
    Add-Attention 'Info' 'Tasks may wake this PC' ("{0} scheduled task(s) are both enabled and allowed to wake it. They are listed below with where to turn that off." -f $armedTasks.Count)
}
if ($wakeArmed.Count -gt 0 -and $WakeCount -gt 0 -and $WakeExplained -eq 0) {
    Add-Attention 'Info' 'Wakes with no recorded cause' 'Every wake in this window came back without a source. The armed devices below are the usual suspects.'
}
if ($TruncNotes.Count -gt 0) {
    Add-Attention 'Warning' 'Part of the window was not read' ("A source hit the {0}-event ceiling, so the oldest part of the window is missing from it and the counts here are lower bounds. Run with a smaller -Days or a larger -MaxEventsPerQuery for a complete answer." -f $MaxEventsPerQuery)
}
if ($Degraded) {
    Add-Attention 'Warning' 'Part of this report is missing' 'At least one source could not be read. See "What could not be read" at the bottom.'
}

# ================================================================= 8. render
Write-Line '  Rendering...'

# --- verdict ---------------------------------------------------------------
$verdictHtml = '<p class="lede">' + (ConvertTo-HtmlText $SummaryLine) + '</p>'
# Only worth saying when the log itself stops inside the window - and only when
# the harvest was complete, because a truncated one cannot support a claim
# about retention any more than it can support the counts above it.
# The hour of slack is presentation, not doubt: a log that falls short of the
# window by minutes has cost the reader nothing worth a paragraph.
if ($TruncNotes.Count -eq 0 -and $null -ne $LogFloorUtc -and ($LogFloorUtc - $WindowStartUtc).TotalHours -gt 1) {
    $verdictHtml += '<p class="note">' + (ConvertTo-HtmlText (
        'The System log only reaches back to {0}, which is less than the {1} days asked for. Anything older has been overwritten and cannot be recovered.' -f (Format-LocalStamp $LogFloorUtc), $Days)) + '</p>'
}
Add-Section 'The short answer' $verdictHtml

# --- wake forensics --------------------------------------------------------
$wakeBody = New-Object System.Text.StringBuilder
[void]$wakeBody.Append('<p class="lede">This tool never changes a setting. Everything below is where <em>you</em> would turn it off.</p>')

[void]$wakeBody.Append('<h3>Devices allowed to wake it</h3>')
if ($wakeArmed.Count -gt 0) {
    $rows = @($wakeArmed | ForEach-Object {
        [pscustomobject]@{
            Device = $_
            'Where to turn it off' = 'Device Manager, find this device, Properties, Power Management tab, clear "Allow this device to wake the computer".'
        }
    })
    [void]$wakeBody.Append((New-DataTable -Rows $rows))
}
else {
    [void]$wakeBody.Append('<p class="empty">No device is armed to wake this PC.</p>')
}

[void]$wakeBody.Append('<h3>Scheduled tasks allowed to wake it</h3>')
if (-not $taskSourceOk) {
    [void]$wakeBody.Append('<p class="empty">The Task Scheduler did not answer in time, so this list could not be built.</p>')
}
elseif ($scheduledWake.Count -gt 0) {
    $rows = @($scheduledWake | ForEach-Object {
        $state = [string]$_.State
        $armed = 'No - the task is disabled'
        if ($state.ToLowerInvariant() -ne 'disabled') { $armed = 'YES' }
        $next = ''
        if ($null -ne $_.NextRun) { $next = Format-LocalStamp (ConvertTo-Utc $_.NextRun) }
        [pscustomobject]@{
            Task    = ('{0}{1}' -f $_.TaskPath, $_.TaskName)
            State   = $state
            'Can wake it' = $armed
            'Next run' = $next
            'Where to turn it off' = 'Task Scheduler, open this task, Properties, Conditions tab, clear "Wake the computer to run this task".'
        }
    })
    [void]$wakeBody.Append((New-DataTable -Rows $rows))
    [void]$wakeBody.Append('<p class="note">' + (ConvertTo-HtmlText 'A disabled task still shows a next run time. It cannot wake anything while it is disabled - the "Can wake it" column is the one that matters.') + '</p>')
}
else {
    [void]$wakeBody.Append('<p class="empty">No scheduled task is set to wake this PC.</p>')
}

[void]$wakeBody.Append('<h3>Wake timers</h3>')
if ($wakeTimersDenied) {
    [void]$wakeBody.Append('<p class="empty">' + (ConvertTo-HtmlText 'Active wake timers need administrator rights to list - everything else in this report is complete. To see them, open an administrator Command Prompt and run: powercfg /waketimers') + '</p>')
}
elseif ($wakeTimers.Count -gt 0) {
    [void]$wakeBody.Append('<pre class="raw">' + (ConvertTo-HtmlText ($wakeTimers -join [Environment]::NewLine)) + '</pre>')
    [void]$wakeBody.Append('<p class="note">' + (ConvertTo-HtmlText 'To stop wake timers entirely: Control Panel, Power Options, Change plan settings, Change advanced power settings, Sleep, Allow wake timers, Disable.') + '</p>')
}
else {
    [void]$wakeBody.Append('<p class="empty">No active wake timers.</p>')
}

if ($lastWake.Count -gt 0) {
    [void]$wakeBody.Append('<h3>What Windows says woke it last</h3>')
    [void]$wakeBody.Append('<pre class="raw">' + (ConvertTo-HtmlText ($lastWake -join [Environment]::NewLine)) + '</pre>')
}

Add-Section 'Why your PC turns on by itself' $wakeBody.ToString()

# --- sessions --------------------------------------------------------------
$sessionRows = New-Object System.Collections.Generic.List[psobject]
foreach ($s in @($Sessions | Sort-Object @{ Expression = { if ($null -eq $_.BootTimeUtc) { [datetime]::MinValue } else { $_.BootTimeUtc } }; Descending = $true })) {
    $bootText = '(started before this window)'
    if ($null -ne $s.BootTimeUtc) { $bootText = Format-LocalStamp $s.BootTimeUtc }

    $typeText = ''
    if ($null -ne $s.BootType) {
        switch ([long]$s.BootType) {
            0 { $typeText = 'Cold boot' }
            1 { $typeText = 'Fast startup resume' }
            2 { $typeText = 'Hibernate resume' }
            default { $typeText = ('Boot type {0}' -f $s.BootType) }
        }
    }

    $upText = ''
    if ($null -ne $s.BootTimeUtc) {
        $endForSpan = $s.EndTimeUtc
        if ($null -eq $endForSpan -and $s.Category -eq 'Running') { $endForSpan = $NowUtc }
        if ($null -ne $endForSpan) { $upText = Format-Duration ($endForSpan - $s.BootTimeUtc) }
    }

    $cls = ''
    if ($s.Category -eq 'Bluescreen') { $cls = 'bad' }
    elseif ($s.Category -eq 'PowerLoss' -or $s.Category -eq 'PowerButton') { $cls = 'warnrow' }

    $sessionRows.Add([pscustomobject]@{
        'Turned on'  = $bootText
        'How'        = $typeText
        'Up for'     = $upText
        'Ended'      = $(if ($null -ne $s.EndTimeUtc) { Format-LocalStamp $s.EndTimeUtc } else { '' })
        'Verdict'    = $s.Verdict
        'What happened' = $s.Detail
        '_RowClass'  = $cls
    })
}
Add-Section 'Every time it started and stopped' `
    (New-DataTable -Rows @($sessionRows) -Columns @('Turned on','How','Up for','Ended','Verdict','What happened') -EmptyText 'No boot was recorded in this window.') `
    'Newest first. Times are this machine''s local time.'

# --- timeline --------------------------------------------------------------
$timeline = New-Object System.Collections.Generic.List[psobject]
function Add-Timeline {
    param($TimeUtc, [string]$What, [string]$Text, [string]$Class = '')
    if ($null -eq $TimeUtc) { return }
    $timeline.Add([pscustomobject]@{ TimeUtc = [datetime]$TimeUtc; What = $What; Text = $Text; Class = $Class })
}

foreach ($s in $Sessions) {
    if ($null -ne $s.BootTimeUtc) {
        $how = 'Started up'
        if ($null -ne $s.BootType -and [long]$s.BootType -eq 2) { $how = 'Resumed from hibernation' }
        elseif ($null -ne $s.BootType -and [long]$s.BootType -eq 1) { $how = 'Resumed from fast startup' }
        Add-Timeline $s.BootTimeUtc $how 'Windows started a new session.'
    }
    if ($null -ne $s.EndTimeUtc -or $s.Category -eq 'Bluescreen' -or $s.Category -eq 'PowerLoss' -or $s.Category -eq 'PowerButton') {
        $when = $s.EndTimeUtc
        if ($null -eq $when) {
            foreach ($ev in @($s.Evidence)) {
                $ct = Get-Field $ev.Data 'CrashTimeUtc' $null
                if ($null -ne $ct) { $when = $ct; break }
            }
        }
        if ($null -eq $when -and $null -ne $s.NextBootUtc) { $when = $s.NextBootUtc }
        $cls = ''
        if ($s.Category -eq 'Bluescreen') { $cls = 'bad' }
        elseif ($s.Category -eq 'PowerLoss' -or $s.Category -eq 'PowerButton') { $cls = 'warnrow' }
        Add-Timeline $when $s.Verdict $s.Detail $cls
    }
}

# A hibernate resume with no boot event of its own is a real event in the
# user's day and would otherwise vanish between two sessions.
foreach ($bt in $bootTypeRecords) {
    $t = Get-Field $bt.Data 'BootType' $null
    if ($null -eq $t -or [long]$t -ne 2) { continue }
    $matched = $false
    foreach ($s in $Sessions) {
        if ($null -eq $s.BootTimeUtc) { continue }
        if ([math]::Abs(($bt.TimeUtc - $s.BootTimeUtc).TotalSeconds) -le $BootDedupeSeconds) { $matched = $true; break }
    }
    if (-not $matched) {
        Add-Timeline $bt.TimeUtc 'Resumed from hibernation' 'It came back from hibernation rather than starting fresh.'
    }
}

foreach ($w in $mergedWakes) {
    $exp = Get-WakeExplanation $w
    Add-Timeline $w.TimeUtc 'Woke up' $exp.Text
}

foreach ($x in $NotableSleeps) {
    $secs = Get-Field $x.Data 'DurationSecond' $null
    $span = [timespan]::FromSeconds([double]$secs)
    Add-Timeline $x.TimeUtc 'Was asleep' ('It had been asleep for {0}.' -f (Format-Duration $span))
}

foreach ($h in $hardwareErrors) {
    Add-Timeline $h.TimeUtc 'Hardware error' ('The hardware reported an error (WHEA event {0}). Repeated ones mean a component is failing.' -f $h.EventId) 'bad'
}

$timelineHtml = New-Object System.Text.StringBuilder
$ordered = @($timeline | Sort-Object TimeUtc -Descending)
if ($ordered.Count -eq 0) {
    [void]$timelineHtml.Append('<p class="empty">Nothing happened in this window that the log recorded.</p>')
}
else {
    $currentDay = ''
    foreach ($t in $ordered) {
        $day = Format-LocalStamp $t.TimeUtc 'yyyy-MM-dd'
        if ($day -ne $currentDay) {
            if ($currentDay -ne '') { [void]$timelineHtml.Append('</ul>') }
            $heading = ([datetime]$t.TimeUtc).ToLocalTime().ToString('dddd, d MMMM yyyy', $Invariant)
            [void]$timelineHtml.Append('<h3 class="day">' + (ConvertTo-HtmlText $heading) + '</h3><ul class="tl">')
            $currentDay = $day
        }
        $cls = ''
        if ($t.Class) { $cls = ' class="' + $t.Class + '"' }
        [void]$timelineHtml.Append(('<li{0}><span class="t">{1}</span><span class="w">{2}</span><span class="d">{3}</span></li>' -f `
            $cls,
            (ConvertTo-HtmlText (Format-LocalStamp $t.TimeUtc 'HH:mm:ss')),
            (ConvertTo-HtmlText $t.What),
            (ConvertTo-HtmlText $t.Text)))
    }
    [void]$timelineHtml.Append('</ul>')
}

$timelineNote = 'Newest day first.'
if ($ShortSleepTotal -gt 0) {
    $timelineNote = ('{0} {1} short Modern Standby cycle(s) - screen off for under {2} minutes - are counted but not listed, or they would bury everything else.' -f `
        $timelineNote, $ShortSleepTotal, $SleepNoticeMinutes)
}
Add-Section 'What happened, day by day' $timelineHtml.ToString() $timelineNote

# --- machine facts ---------------------------------------------------------
# Neither of these comes from powercfg /a any more. That output lists its state
# names twice, once under an available heading and once under a not-available
# one, and both headings are localized - so a name matched anywhere in the text
# says nothing about which list it was in. This machine prints "Standby (S0 Low
# Power Idle) Network Connected" under the not-available heading, and a machine
# with hibernate switched off still prints the word "Hibernate" there. The raw
# output is shown below as powercfg's own words; the two rows are decided from
# typed sources.

# Modern Standby is not reported as a capability but as something the machine
# was seen doing: Kernel-Power 506/507 are its sleep and wake pair, 42/107 the
# classic S3 one.
$allSleepRecords = @(@($sleepEnters) + @($sleepExits))
$modernSleepSeen  = @($allSleepRecords | Where-Object { (Get-Field $_.Data 'Family' '') -eq 'modern' }).Count
$classicSleepSeen = @($allSleepRecords | Where-Object { (Get-Field $_.Data 'Family' '') -eq 'classic' }).Count

$modernStandby = 'No sleep recorded in this window'
if ($modernSleepSeen -gt 0)      { $modernStandby = 'Yes - it slept this way in this window' }
elseif ($classicSleepSeen -gt 0) { $modernStandby = 'No - it used the older S3 sleep instead' }

$hibernateAvailable = 'Could not be read'
if ($null -ne $hibernateFlag) {
    if ([int]$hibernateFlag -eq 0) { $hibernateAvailable = 'No' } else { $hibernateAvailable = 'Yes' }
}
# A resume from hibernation settles it whatever the registry says.
if ($HibernateResumes -gt 0) { $hibernateAvailable = 'Yes' }

$fastStartup = 'Could not be read'
if ($null -ne $hiberboot) {
    if ([int]$hiberboot -eq 0) { $fastStartup = 'Off' } else { $fastStartup = 'On' }
}

$factPairs = @{
    'Computer name'        = $env:COMPUTERNAME
    'Window'               = ('{0} to {1}' -f (Format-LocalStamp $WindowStartUtc 'yyyy-MM-dd HH:mm'), (Format-LocalStamp $NowUtc 'yyyy-MM-dd HH:mm'))
    'Boot sessions seen'   = $RealSessions.Count
    'Modern Standby'       = $modernStandby
    'Hibernate available'  = $hibernateAvailable
    'Fast startup'         = $fastStartup
    'Hibernate resumes'    = $HibernateResumes
    'Battery present'      = $(if ($script:HasBattery) { 'Yes' } else { 'No' })
    'Last clean shutdown'  = $(if ($null -ne $lastCleanShutdownUtc) { Format-LocalStamp $lastCleanShutdownUtc } else { 'not recorded' })
}
$factsBody = New-Object System.Text.StringBuilder
[void]$factsBody.Append((New-KeyValueTable -Pairs $factPairs -Order @(
    'Computer name','Window','Boot sessions seen','Modern Standby','Hibernate available',
    'Fast startup','Hibernate resumes','Battery present','Last clean shutdown')))
if ($sleepStates.Count -gt 0) {
    [void]$factsBody.Append('<h3>Sleep states, in powercfg''s own words</h3>')
    [void]$factsBody.Append('<pre class="raw">' + (ConvertTo-HtmlText ($sleepStates -join [Environment]::NewLine)) + '</pre>')
    [void]$factsBody.Append('<p class="note">' + (ConvertTo-HtmlText 'Shown as printed, not interpreted: powercfg lists the same state names under both an available and a not-available heading, so reading a verdict out of this text is how a PC with hibernate switched off gets reported as having it.') + '</p>')
}
Add-Section 'This machine' $factsBody.ToString() `
    'Fast startup makes "Shut down" hibernate instead. When it is on, a shutdown does not clear a problem the way a restart does.'

# --- degraded --------------------------------------------------------------
$notesHtml = New-Object System.Text.StringBuilder
if ($DegradedNotes.Count -eq 0 -and $TruncNotes.Count -eq 0 -and -not $wakeTimersDenied) {
    [void]$notesHtml.Append('<p class="empty">Every source this report uses was readable.</p>')
}
else {
    [void]$notesHtml.Append('<ul class="plain">')
    foreach ($n in $DegradedNotes) { [void]$notesHtml.Append('<li>' + (ConvertTo-HtmlText $n) + '</li>') }
    foreach ($n in $TruncNotes)    { [void]$notesHtml.Append('<li>' + (ConvertTo-HtmlText $n) + '</li>') }
    if ($wakeTimersDenied) {
        [void]$notesHtml.Append('<li>' + (ConvertTo-HtmlText 'Active wake timers need administrator rights to list. This is expected and does not affect anything else in the report.') + '</li>')
    }
    [void]$notesHtml.Append('</ul>')
}
Add-Section 'What could not be read' $notesHtml.ToString()

# --- fine print ------------------------------------------------------------
$finePrint = @'
<ul class="plain">
<li>This tool is read-only. It changes no setting, deletes nothing, and writes only its own report and history file.</li>
<li>Verdicts come from numeric codes and typed fields, never from the wording of an event, so they read the same on a Windows in any language.</li>
<li>"Windows Update restarted it" is a correlation, not something Windows states outright: a system-initiated restart with an update installed within half an hour of it. Defender signature updates are excluded, because they install several times a day and would otherwise claim every restart.</li>
<li>A wake with no recorded source is the normal case, not a fault. Windows only names a source when a driver bothers to supply one.</li>
<li>Crash records are written by the boot that follows the crash, so they are attributed to the session before the one they appear in.</li>
</ul>
'@
Add-Section 'Fine print' $finePrint

# --- page ------------------------------------------------------------------
$attentionHtml = ''
if ($Attention.Count -eq 0) {
    $attentionHtml = '<div class="ok">No bluescreens and no unexpected power losses in this window.</div>'
}
else {
    $rank = @{ 'Critical' = 0; 'Warning' = 1; 'Info' = 2 }
    $sorted = @($Attention | Sort-Object @{ Expression = { $rank[$_.Level] } })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<ul class="attention">')
    foreach ($a in $sorted) {
        [void]$sb.Append(('<li class="{0}"><span class="tag">{1}</span><strong>{2}</strong><span class="detail">{3}</span></li>' -f `
            $a.Level.ToLowerInvariant(), (ConvertTo-HtmlText $a.Level), (ConvertTo-HtmlText $a.Title), (ConvertTo-HtmlText $a.Detail)))
    }
    [void]$sb.Append('</ul>')
    $attentionHtml = $sb.ToString()
}

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
h3 { font-size:14px; margin:20px 0 10px; letter-spacing:-0.01em; }
h3:first-of-type { margin-top:6px; }
.note { color:var(--muted); font-size:13px; margin:-6px 0 14px; }
.empty { color:var(--muted); font-style:italic; margin:0; }
.lede { margin:0 0 12px; font-size:15px; }
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
.data td { padding:8px 12px 8px 0; border-bottom:1px solid var(--line); vertical-align:top; }
.data tr:last-child td { border-bottom:none; }
.data tr.bad td:first-child { box-shadow:inset 3px 0 0 var(--crit); padding-left:9px; }
.data tr.warnrow td:first-child { box-shadow:inset 3px 0 0 var(--warn); padding-left:9px; }
.ok {
  background:var(--card); border:1px solid var(--line); border-left:4px solid var(--ok);
  border-radius:12px; padding:16px 20px; margin-bottom:18px; color:var(--ok); font-weight:600;
}
ul.attention { list-style:none; margin:0 0 18px; padding:0; }
ul.attention li {
  background:var(--card); border:1px solid var(--line); border-radius:12px;
  padding:14px 18px; margin-bottom:10px; display:flex; gap:12px; align-items:baseline; flex-wrap:wrap;
}
ul.attention li.critical { border-left:4px solid var(--crit); }
ul.attention li.warning  { border-left:4px solid var(--warn); }
ul.attention li.info     { border-left:4px solid var(--accent); }
.tag {
  font-size:11px; text-transform:uppercase; letter-spacing:0.06em; font-weight:700;
  padding:2px 8px; border-radius:999px; background:var(--bar); color:var(--muted);
}
li.critical .tag { color:var(--crit); }
li.warning .tag  { color:var(--warn); }
.detail { color:var(--muted); flex:1 1 100%; font-size:13px; }
h3.day {
  color:var(--muted); font-size:12px; text-transform:uppercase; letter-spacing:0.06em;
  border-bottom:1px solid var(--line); padding-bottom:6px; margin:22px 0 8px;
}
ul.tl { list-style:none; margin:0; padding:0; font-size:14px; }
ul.tl li { display:flex; gap:12px; padding:6px 0 6px 9px; align-items:baseline; flex-wrap:wrap; }
ul.tl li.bad { box-shadow:inset 3px 0 0 var(--crit); }
ul.tl li.warnrow { box-shadow:inset 3px 0 0 var(--warn); }
ul.tl .t { color:var(--muted); font-variant-numeric:tabular-nums; white-space:nowrap; width:70px; flex:none; }
ul.tl .w { font-weight:600; white-space:nowrap; }
ul.tl .d { color:var(--muted); flex:1 1 260px; font-size:13px; }
ul.plain { margin:0; padding-left:18px; color:var(--muted); font-size:13px; }
ul.plain li { margin-bottom:6px; }
pre.raw {
  background:var(--bg); border:1px solid var(--line); border-radius:8px; padding:12px 14px;
  overflow-x:auto; font-size:13px; margin:0 0 10px; white-space:pre-wrap; word-break:break-word;
}
footer { color:var(--muted); font-size:12px; text-align:center; margin-top:28px; }
'@

$html = New-Object System.Collections.Generic.List[string]
$html.Add('<!doctype html>')
$html.Add('<html lang="en"><head><meta charset="utf-8">')
$html.Add('<meta name="viewport" content="width=device-width,initial-scale=1">')
$html.Add(('<title>Why this PC restarted - {0}</title>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add('<style>' + $css + '</style></head><body><div class="wrap">')
$html.Add(('<h1>{0}</h1>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add(('<p class="sub">Restart, crash and wake history &middot; last {0} day(s) &middot; generated {1}</p>' -f `
    $Days, (ConvertTo-HtmlText (Format-LocalStamp $NowUtc))))
$html.Add('<h2 style="font-size:15px;margin:0 0 12px;">Needs attention</h2>')
$html.Add($attentionHtml)
foreach ($s in $Sections) { $html.Add($s) }
$html.Add('<footer>Generated by Get-RebootReason.ps1 - read-only, nothing on this machine was changed.</footer>')
$html.Add('</div></body></html>')

$reportWritten = $true
try {
    ($html -join [Environment]::NewLine) | Set-Content -LiteralPath $ReportFile -Encoding utf8
}
catch {
    $reportWritten = $false
    Write-Line "  ! Could not write $ReportFile : $($_.Exception.Message)" 'Yellow'
}

# ============================================================== 9. history CSV
$historyRows = New-Object System.Collections.Generic.List[psobject]
foreach ($s in @($Sessions | Sort-Object @{ Expression = { if ($null -eq $_.BootTimeUtc) { [datetime]::MinValue } else { $_.BootTimeUtc } } })) {
    $hours = ''
    if ($null -ne $s.BootTimeUtc) {
        $endForSpan = $s.EndTimeUtc
        if ($null -eq $endForSpan -and $s.Category -eq 'Running') { $endForSpan = $NowUtc }
        if ($null -ne $endForSpan) { $hours = [math]::Round((($endForSpan - $s.BootTimeUtc).TotalHours), 2) }
    }
    $typeText = ''
    if ($null -ne $s.BootType) { $typeText = [string]$s.BootType }

    $historyRows.Add([pscustomobject]@{
        RunId                = $RunId
        ComputerName         = $env:COMPUTERNAME
        BootTimeUtc          = Format-UtcStamp $s.BootTimeUtc
        BootType             = $typeText
        BootVerdict          = $(if ($s.IsPreWindow) { 'Before window' } else { 'Booted' })
        EndTimeUtc           = Format-UtcStamp $s.EndTimeUtc
        EndVerdict           = $s.Verdict
        EndDetail            = $s.Detail
        StopCodeHex          = $s.StopCodeHex
        KbNumber             = $s.Kb
        WakeCount            = $s.Wakes
        UnexplainedWakeCount = $s.Unexplained
        SessionHours         = $hours
        Degraded             = $Degraded
        ScriptVersion        = $ScriptVersion
    })
}

$historyOk = $true
try { Add-HistoryRow -Rows @($historyRows) -Path $HistoryCsv -MaxRows $HistoryLimit }
catch {
    $historyOk = $false
    Write-Line "  ! Could not update $HistoryCsv : $($_.Exception.Message)" 'Yellow'
}

# ================================================================ 10. console
if (-not $Quiet) {
    Write-Host ''
    Write-Host $SummaryLine
    Write-Host ''

    $recent = @($Sessions |
        Where-Object { $_.Category -eq 'Bluescreen' -or $_.Category -eq 'PowerLoss' -or $_.Category -eq 'PowerButton' } |
        Sort-Object @{ Expression = { if ($null -eq $_.BootTimeUtc) { [datetime]::MinValue } else { $_.BootTimeUtc } }; Descending = $true })

    if ($recent.Count -gt 0) {
        Write-Host 'Sessions that did not end cleanly'
        Write-Host '---------------------------------'
        foreach ($s in @($recent | Select-Object -First 10)) {
            $c = 'Yellow'
            if ($s.Category -eq 'Bluescreen') { $c = 'Red' }
            $when = '(before this window)'
            if ($null -ne $s.BootTimeUtc) { $when = Format-LocalStamp $s.BootTimeUtc }
            Write-Host ('  session from {0}' -f $when)
            Write-Host ('    {0}' -f $s.Detail) -ForegroundColor $c
        }
        if ($recent.Count -gt 10) {
            Write-Host ('  ... and {0} more in the report.' -f ($recent.Count - 10))
        }
        Write-Host ''
    }

    if ($wakeArmed.Count -gt 0 -or $scheduledWake.Count -gt 0) {
        Write-Host 'Allowed to wake this PC'
        Write-Host '-----------------------'
        foreach ($d in $wakeArmed)   { Write-Host ('  device : {0}' -f $d) }
        foreach ($t in $armedTasks)  { Write-Host ('  task   : {0}{1}' -f $t.TaskPath, $t.TaskName) }
        if ($armedTasks.Count -eq 0 -and $scheduledWake.Count -gt 0) {
            Write-Host ('  {0} task(s) are set to wake it, but all of them are disabled.' -f $scheduledWake.Count)
        }
        Write-Host '  The report says where to turn each of these off.'
        Write-Host ''
    }

    if ($Attention.Count -gt 0) {
        Write-Host 'Needs attention'
        Write-Host '---------------'
        foreach ($a in $Attention) {
            $c = 'Yellow'
            if ($a.Level -eq 'Critical') { $c = 'Red' }
            if ($a.Level -eq 'Info')     { $c = 'Cyan' }
            Write-Host ('  [{0}] {1} - {2}' -f $a.Level, $a.Title, $a.Detail) -ForegroundColor $c
        }
        Write-Host ''
    }

    if ($reportWritten) { Write-Host "Report  : $ReportFile" }
    if ($historyOk)     { Write-Host "History : $HistoryCsv" }
}

if ($Open -and $reportWritten -and (Test-Path -LiteralPath $ReportFile)) {
    # A machine with no handler registered for .html throws here, which would
    # otherwise lose the exit code the caller actually asked for.
    try { Start-Process -FilePath $ReportFile -ErrorAction Stop }
    catch { Write-Line "  ! Could not open the report: $($_.Exception.Message)" 'Yellow' }
}

# Degraded wins: a caller acting on "all clear" deserves to know the report was
# not complete before it deserves to know the machine looked fine.
if ($Degraded) { exit 2 }
if ($BluescreenCount -gt 0 -or $PowerLossCount -gt 0 -or $PowerButtonCount -gt 0) { exit 1 }
exit 0
