<#
.SYNOPSIS
    Traffic-light answer to "are my files actually backed up?" for Desktop,
    Documents, Pictures, Downloads and any extra folders you name - based on
    evidence found on this machine, not on what a settings page claims.

.DESCRIPTION
    Read-only. Nothing outside this tool's own output folder is written, no
    service or process is started or stopped, and no file under a sync root is
    ever opened - reading a cloud-only placeholder asks the sync provider to
    download it, which is exactly the failure this tool exists to detect.

    The tool never certifies a backup. Every verdict is phrased as "signs of" or
    "no signs of", because the only thing a script on this side of the wire can
    see is evidence: a folder sitting inside a sync root, a File History
    inclusion list, a provider process, a timestamp. Whether the other end of
    that pipe still holds a readable copy of your files is not knowable from
    here.

    The case worth knowing about is the zombie sync root. OneDrive can be signed
    out for a year and still leave: %OneDrive% set, the sync folder on disk,
    ClientEverSignedIn = 1, and folders full of files whose contents live only
    in the cloud. Explorer shows names and sizes; opening one fails. Every naive
    signal says "backed up", every signal that matters says "not syncing". A
    folder in that state is reported RED, and deliberately called worse than no
    backup at all.

    Mechanisms checked: OneDrive (account subtree, process, SyncRootManager,
    SyncEngines mounts, sync-root marker files, placeholder attribute bits),
    File History (configuration XML, fhsvc, catalog, its event channels),
    Dropbox and Google Drive (config, process, mounted volumes). System Restore
    and Windows' settings backup are reported as context only - neither one
    backs up your files, and neither can change a folder's light.

    A folder can be covered by more than one mechanism. The light is the best of
    the coverage evidence, except that a placeholder hazard overrides everything:
    a second mechanism does not make an unreadable file readable.

.PARAMETER Path
    Extra folders to assess alongside Desktop, Documents, Pictures and Downloads.
    A path that does not exist is reported as an error row and sets exit code 3;
    the rest of the run continues.

.PARAMETER StaleDays
    Evidence older than this many days turns a GREEN into a YELLOW marked stale.
    Default 14.

.PARAMETER SampleLimit
    Per-folder cap on the recursive file walk used for the recency reading and
    the cloud-only census. Default 2000. The cap is not politeness: walking a
    sync root whose provider is not running took 23 seconds for 59 files on the
    machine this was built against, because every placeholder waits for a
    provider that never answers.

.PARAMETER HistoryLimit
    Rows kept in BackupReality_history.csv. Default 5000; 0 keeps everything.
    When the column set changes the old file is archived, not appended to.

.PARAMETER OutputDir
    Where the report and history CSV land. Defaults to
    $env:LOCALAPPDATA\Test-BackupReality. If that cannot be written to, the tool
    falls back to %TEMP%\Test-BackupReality and says so.

.PARAMETER Now
    The clock to measure staleness against. Injectable so the staleness maths can
    be exercised without waiting weeks: -Now (Get-Date).AddDays(400) makes every
    piece of evidence stale.

.PARAMETER Quiet
    Suppress console output; still writes the files and sets the exit code.

.PARAMETER Open
    Open the generated HTML report when finished.

.OUTPUTS
    Exit code 0 = every assessed location is GREEN (signs of active backup,
    fresher than -StaleDays).
    1 = at least one YELLOW (signs present but stale, weak, or not fully
    verifiable), no RED.
    2 = at least one RED: no signs of any backup, or inside a zombie sync root
    whose files may be unreachable placeholders.
    3 = invocation error: a -Path entry does not exist, or not one known folder
    could be resolved. The report is still written for whatever did resolve.
    4 = internal failure, or no writable output folder.
    Higher codes win: an invocation error is reported as 3 even when a folder is
    also RED, because the caller asked for something the tool could not do.

.EXAMPLE
    .\Test-BackupReality.ps1 -Open

.EXAMPLE
    # Assess a project folder alongside the four known folders
    .\Test-BackupReality.ps1 -Path 'D:\Work','D:\Photos'

.EXAMPLE
    # Prove the staleness maths without waiting a year
    .\Test-BackupReality.ps1 -Now (Get-Date).AddDays(400) -Quiet
#>
[CmdletBinding()]
param(
    # Extra folders to assess alongside Desktop/Documents/Pictures/Downloads
    [Parameter(Position = 0)]
    [string[]]$Path = @(),
    # Evidence older than this many days turns a GREEN into YELLOW ("stale")
    [ValidateRange(1, 365)]
    [int]$StaleDays = 14,
    # Per-folder cap on recursive file sampling (recency + cloud-only census).
    # Bounded because a cold sync-root walk took 23s for 59 files on the probe machine.
    [ValidateRange(100, 100000)]
    [int]$SampleLimit = 2000,
    # Rolling CSV row cap; 0 = keep all
    [ValidateRange(0, 1000000)]
    [int]$HistoryLimit = 5000,
    # Survives a null LOCALAPPDATA: binding null into Join-Path dies before the body runs
    [string]$OutputDir = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'Test-BackupReality'),
    # Injectable clock so staleness math is testable without waiting weeks.
    # DontShow keeps it out of the GUI launcher's generated form and out of tab
    # completion: it is a test seam, and the only correct value is the default.
    [Parameter(DontShow)]
    [datetime]$Now = (Get-Date),
    [switch]$Quiet,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ToolVersion = '1.0.0'

# A diagnostic that dies should still say why and return its documented
# internal-failure code rather than a bare 1. This prints even under -Quiet:
# a silent exit 4 is indistinguishable from a hung machine.
trap {
    Write-Host ''
    Write-Host ("Test-BackupReality failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
    if ($_.InvocationInfo) {
        Write-Host ("  at {0}" -f $_.InvocationInfo.PositionMessage.Trim()) -ForegroundColor DarkGray
    }
    exit 4
}

# ------------------------------------------------------------------- plumbing
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

function Get-Prop {
    # Strict-mode-safe property read.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
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

$RunStart = Get-Date
$OutputDir = Initialize-OutputDir -Preferred $OutputDir -ToolName 'Test-BackupReality'
$ReportFile = Join-Path $OutputDir 'BackupReality.html'
$HistoryCsv = Join-Path $OutputDir 'BackupReality_history.csv'

$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$Epoch     = New-Object System.DateTime(1970, 1, 1, 0, 0, 0, ([System.DateTimeKind]::Utc))

# All staleness arithmetic happens in UTC so a machine that changes time zone
# between runs does not produce a jump in the history CSV.
$NowUtc = $Now
if ($Now.Kind -ne [System.DateTimeKind]::Utc) { $NowUtc = $Now.ToUniversalTime() }

$Attention = New-Object System.Collections.Generic.List[psobject]
$Sections  = New-Object System.Collections.Generic.List[string]
$AdminSkips = New-Object System.Collections.Generic.List[psobject]

function Write-Line {
    param([string]$Message, [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Message -ForegroundColor $Colour }
    else         { Write-Host $Message }
}

function Write-Step {
    param([string]$Message)
    if (-not $Quiet) { Write-Host "  $Message" }
}

function Add-AdminSkip {
    # Status is a parameter because "needs admin" is a claim about why a check
    # did not run, and it is false when the run was elevated and the check
    # failed anyway.
    param([string]$Check, [string]$Consequence, [string]$Status = 'skipped: needs admin')
    $AdminSkips.Add([pscustomobject]@{ Check = $Check; Status = $Status; Consequence = $Consequence })
}

# --------------------------------------------------------------- formatting
function Format-Utc {
    <#
        Every timestamp this tool prints or stores is invariant-formatted UTC.
        The machine this was built on is en-IN (dd-MM-yyyy); a CSV written in
        that format and read back on a us-EN machine silently swaps day and
        month for the first twelve days of every month.
    #>
    param($When)
    if ($null -eq $When) { return '' }
    $d = [datetime]$When
    if ($d.Kind -ne [System.DateTimeKind]::Utc) { $d = $d.ToUniversalTime() }
    $d.ToString('yyyy-MM-dd HH:mm:ss', $Invariant) + ' UTC'
}

function Format-IsoUtc {
    param($When)
    if ($null -eq $When) { return '' }
    $d = [datetime]$When
    if ($d.Kind -ne [System.DateTimeKind]::Utc) { $d = $d.ToUniversalTime() }
    $d.ToString('u', $Invariant)
}

function ConvertTo-UtcDate {
    <#
        CIM dates arrive in two shapes in the same report: a real DateTime for a
        property the class declares as datetime, and a DMTF string like
        '20250614103000.000000+330' for one it declares as a string. Both are
        accepted; anything else is no date rather than a wrong one.
    #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return ([datetime]$Value).ToUniversalTime() }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return [System.Management.ManagementDateTimeConverter]::ToDateTime($text).ToUniversalTime() } catch { }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, $Invariant, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    $null
}

function Get-AgeDays {
    param($WhenUtc)
    if ($null -eq $WhenUtc) { return $null }
    ($NowUtc - ([datetime]$WhenUtc)).TotalDays
}

function Format-AgeText {
    param($WhenUtc)
    if ($null -eq $WhenUtc) { return 'no date' }
    $days = Get-AgeDays $WhenUtc
    # A clock that has been wound back (or -Now in the past) produces negative
    # ages. Printing "-38 days ago" invites the reader to trust a number that
    # is telling them their clock is wrong.
    if ($days -lt -1) { return ('{0} (in the future)' -f (Format-Utc $WhenUtc)) }
    if ($days -lt 1)  { return ('{0} (today)' -f (Format-Utc $WhenUtc)) }
    ('{0} ({1:N0} days ago)' -f (Format-Utc $WhenUtc), $days)
}

function Format-Size {
    param([double]$Bytes)
    if     ($Bytes -ge 1TB) { '{0:N2} TB' -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { '{0:N1} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N0} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -gt 0)   { '{0:N0} KB' -f ($Bytes / 1KB) }
    else                    { '0' }
}

function Join-Readable {
    # "A", "A and B", "A, B and C" - the report is read by people, not parsers.
    param([string[]]$Items)
    $list = @($Items | Where-Object { $_ })
    if ($list.Count -eq 0) { return '' }
    if ($list.Count -eq 1) { return $list[0] }
    if ($list.Count -eq 2) { return ($list[0] + ' and ' + $list[1]) }
    (($list[0..($list.Count - 2)]) -join ', ') + ' and ' + $list[-1]
}

# ----------------------------------------------------------------- registry
function Read-RegistryKey {
    <#
        One snapshot of a key - does it exist, its value names and data, its
        subkey names - with the handle closed afterwards.

        Subkeys are always enumerated rather than inferred from the key's
        existence. HKCU:\Software\SyncEngines\Providers exists on a machine with
        no sync engines registered at all, so Test-Path answers True to the
        question "is anything syncing here?" and is wrong.

        -RawValues suppresses the expansion Get-ItemProperty performs on
        REG_EXPAND_SZ, which is the only way to tell a literal path apart from
        one that reads %USERPROFILE%\Desktop and happens to expand to the same
        place.
    #>
    param(
        [ValidateSet('HKCU', 'HKLM')][string]$Hive,
        [string]$SubKey,
        [switch]$RawValues
    )

    $result = [pscustomobject]@{
        Exists  = $false
        Denied  = $false
        Values  = @{}
        Kinds   = @{}
        SubKeys = @()
    }

    $key = $null
    try {
        if ($Hive -eq 'HKCU') { $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($SubKey) }
        else                  { $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($SubKey) }
        if ($null -eq $key) { return $result }

        $result.Exists  = $true
        $result.SubKeys = @($key.GetSubKeyNames())

        foreach ($n in @($key.GetValueNames())) {
            try {
                if ($RawValues) {
                    $result.Values[$n] = $key.GetValue($n, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                }
                else {
                    $result.Values[$n] = $key.GetValue($n, $null)
                }
                $result.Kinds[$n] = [string]$key.GetValueKind($n)
            }
            catch { }
        }
    }
    catch [System.Security.SecurityException] { $result.Denied = $true }
    catch { }
    finally { if ($null -ne $key) { try { $key.Close() } catch { } } }

    $result
}

function Get-RegString {
    param($Snapshot, [string]$Name, [string]$Default = '')
    if ($null -eq $Snapshot) { return $Default }
    if (-not $Snapshot.Values.ContainsKey($Name)) { return $Default }
    $v = $Snapshot.Values[$Name]
    if ($null -eq $v) { return $Default }
    [string]$v
}

function Get-RegNumber {
    param($Snapshot, [string]$Name, $Default = $null)
    if ($null -eq $Snapshot) { return $Default }
    if (-not $Snapshot.Values.ContainsKey($Name)) { return $Default }
    $v = $Snapshot.Values[$Name]
    if ($null -eq $v) { return $Default }
    $parsed = [long]0
    if ([long]::TryParse(([string]$v).Trim(), [ref]$parsed)) { return $parsed }
    $Default
}

# --------------------------------------------------------------------- paths
function ConvertTo-NormalPath {
    <#
        One spelling per folder so containment can be a string compare.
        Ordinal-invariant lowercase, not ToLower(): a Turkish locale maps I to a
        dotless i and 'C:\Users\Ilhan' stops matching itself.

        A relative path is resolved against the caller's PowerShell location.
        PowerShell never syncs [Environment]::CurrentDirectory with
        Set-Location, and GetFullPath uses that stale process directory: from a
        session sitting in one folder, -Path 'probe' would be assessed in a
        different tree and still reported under the name the caller typed.

        Anything that cannot become an absolute filesystem path comes back
        empty rather than as an invented one. An environment variable that is
        not set survives expansion as its own literal %NAME%, and a location on
        a non-filesystem provider resolves to HKEY_LOCAL_MACHINE\...; anchoring
        either of those to the process directory produces a confident-looking
        path to a folder nobody named.
    #>
    param([string]$RawPath)
    if ([string]::IsNullOrWhiteSpace($RawPath)) { return '' }
    $p = $RawPath
    try { $p = [System.Environment]::ExpandEnvironmentVariables($p) } catch { }
    if ($p -match '%[^%\\]+%') { return '' }

    # Rooted paths are left for GetFullPath: IsPathRooted already covers UNC,
    # and the provider is only needed to answer "relative to what".
    if (-not [System.IO.Path]::IsPathRooted($p)) {
        $viaProvider = ''
        try { $viaProvider = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) } catch { }
        if ([string]::IsNullOrWhiteSpace($viaProvider) -or -not [System.IO.Path]::IsPathRooted($viaProvider)) { return '' }
        $p = $viaProvider
    }

    try { $p = [System.IO.Path]::GetFullPath($p) } catch { }
    if (-not [System.IO.Path]::IsPathRooted($p)) { return '' }
    $p = $p.TrimEnd([char]'\', [char]'/')
    # GetFullPath('C:\') is 'C:\' and trimming it leaves 'C:' which is a drive
    # cursor, not a folder; put the separator back.
    if ($p -match '^[A-Za-z]:$') { $p = $p + '\' }
    $p
}

function Get-PathKey {
    param([string]$NormalPath)
    if ([string]::IsNullOrWhiteSpace($NormalPath)) { return '' }
    $NormalPath.ToLowerInvariant()
}

function Test-PathInside {
    # True when $Child is $Root or lives underneath it.
    param([string]$Child, [string]$Root)
    if ([string]::IsNullOrWhiteSpace($Child) -or [string]::IsNullOrWhiteSpace($Root)) { return $false }
    $c = Get-PathKey $Child
    $r = Get-PathKey $Root
    if ($c -eq $r) { return $true }
    if (-not $r.EndsWith('\')) { $r = $r + '\' }
    $c.StartsWith($r, [System.StringComparison]::Ordinal)
}

# ------------------------------------------------------------------ sampling
function Get-FolderSample {
    <#
        Metadata only. Nothing here opens a file: Get-Content on a cloud-only
        placeholder asks the sync provider to download it, and with the provider
        stopped that either hangs or fails with an error that reads like a disk
        fault. Attributes, LastWriteTimeUtc and Length are all served from the
        directory entry and leave a placeholder dehydrated.

        The walk is capped and stops early. On a sync root whose provider is not
        running each entry waits on a provider that never answers - 23 seconds
        for 59 files on the machine this was written against.
    #>
    param([string]$FolderPath, [int]$Limit)

    $result = [pscustomobject]@{
        Ok             = $false
        FileCount      = 0
        Truncated      = $false
        NewestUtc      = $null
        CloudOnly      = 0
        CloudOnlyBytes = [long]0
        Offline        = 0
        ElapsedSeconds = 0.0
        Note           = ''
    }

    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        # SilentlyContinue on the enumeration so one unreadable subtree costs a
        # subtree, not the whole sample.
        $files = @(Get-ChildItem -LiteralPath $FolderPath -Recurse -Force -File -ErrorAction SilentlyContinue |
                    Select-Object -First ($Limit + 1))
        $result.Ok = $true

        if ($files.Count -gt $Limit) {
            $result.Truncated = $true
            $files = @($files | Select-Object -First $Limit)
        }
        $result.FileCount = $files.Count

        $newest = $null
        foreach ($f in $files) {
            # A DirectoryInfo's intrinsic Length is 1, not a byte count. -File
            # should make that impossible; the guard costs nothing and the
            # alternative is a silently wrong "bytes at risk" figure.
            if (-not ($f -is [System.IO.FileInfo])) { continue }

            $attr = 0
            try { $attr = [int]$f.Attributes } catch { }

            $len = [long]0
            try { $len = [long]$f.Length } catch { }

            # 0x400000 is FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS - the bit that
            # means "this name is real, the contents are not here". Windows
            # PowerShell 5.1 has no enum name for it, so .ToString() on
            # Attributes returns a bare integer and any string match on it is a
            # coincidence waiting to happen. Test the bit.
            if (($attr -band 0x400000) -ne 0) {
                $result.CloudOnly++
                $result.CloudOnlyBytes += $len
            }
            if (($attr -band 0x1000) -ne 0) { $result.Offline++ }

            $t = $null
            try { $t = $f.LastWriteTimeUtc } catch { }
            if ($null -ne $t -and ($null -eq $newest -or $t -gt $newest)) { $newest = $t }
        }
        $result.NewestUtc = $newest
    }
    catch {
        $result.Note = $_.Exception.Message
    }
    finally {
        $watch.Stop()
        $result.ElapsedSeconds = [math]::Round($watch.Elapsed.TotalSeconds, 1)
    }

    $result
}

# -------------------------------------------------------------- event logs
function Get-ChannelInfo {
    <#
        Get-WinEvent raises a bare System.Exception for "no such channel",
        "channel is empty" and "you are not allowed to read this one" alike;
        only the inner exception separates them. On this machine the -ListLog
        form wraps UnauthorizedAccessException while a direct read throws it
        unwrapped, so both shapes are tested.

        RecordCount is read first so an empty channel is skipped without ever
        provoking the throw. A channel's LastWriteTime is not evidence: the
        Backup channel here has a recent LastWriteTime and zero records.
    #>
    param([string]$LogName)

    $info = [pscustomobject]@{
        Name        = $LogName
        Exists      = $false
        Denied      = $false
        RecordCount = 0
        Note        = ''
    }

    try {
        $log = Get-WinEvent -ListLog $LogName -ErrorAction Stop
        $info.Exists = $true
        $rc = Get-Prop $log 'RecordCount' 0
        if ($null -ne $rc) { $info.RecordCount = [int]$rc }
    }
    catch {
        $ex = $_.Exception
        $denied = $false
        if ($ex -is [System.UnauthorizedAccessException]) { $denied = $true }
        elseif ($null -ne $ex.InnerException -and $ex.InnerException -is [System.UnauthorizedAccessException]) { $denied = $true }

        if ($denied) {
            $info.Exists = $true
            $info.Denied = $true
        }
        else {
            $info.Note = $ex.Message
        }
    }

    $info
}

function Get-ChannelEvent {
    param([string]$LogName, [int[]]$Id = @(), [int]$MaxEvents = 50)
    try {
        $filter = @{ LogName = $LogName }
        if ($Id.Count -gt 0) { $filter['Id'] = $Id }
        return @(Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEvents -ErrorAction Stop)
    }
    catch { return @() }
}

function Get-EventDataMap {
    <#
        Named EventData fields, read through XPath rather than the XML type
        adapter: under strict mode $xml.Event.EventData throws outright on an
        event that has no EventData block, and plenty do.
    #>
    param($EventRecord)
    $map = @{}
    if ($null -eq $EventRecord) { return $map }
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml($EventRecord.ToXml())
        $index = 0
        foreach ($d in @($doc.SelectNodes("//*[local-name()='EventData']/*[local-name()='Data']"))) {
            $name = ''
            $attr = $d.Attributes['Name']
            if ($null -ne $attr) { $name = [string]$attr.Value }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = ('[{0}]' -f $index) }
            $map[$name] = [string]$d.InnerText
            $index++
        }
    }
    catch { }
    $map
}

function Get-ConfigPathList {
    <#
        Pulls the paths out of one of File History's folder lists.

        The element names inside FolderInclusionList are not stable across
        Windows builds - UserFolder in one, Folder wrapping FullPath in another -
        so every leaf element under the list whose text looks like a path is
        taken, rather than one hard-coded child name a later build renames.
    #>
    param($Doc, [string]$ListName)

    $out = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Doc) { return @($out) }
    foreach ($n in @($Doc.SelectNodes("//*[local-name()='$ListName']//*"))) {
        if (@($n.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }).Count -gt 0) { continue }
        $t = [string]$n.InnerText
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        $t = $t.Trim()
        if ($t -match '^[A-Za-z]:\\' -or $t.StartsWith('\\')) { $out.Add($t) }
    }
    @($out)
}

# ---------------------------------------------------------------- reachability
function Test-TargetReachable {
    <#
        Test-Path against a powered-off NAS or an unplugged external drive
        blocks on SMB's own timeout, which is tens of seconds and not adjustable
        from here. This is the one check worth the cost of a background job.
    #>
    param([string]$TargetPath, [int]$TimeoutSeconds = 5)

    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return 'unknown' }

    $job = $null
    try {
        $job = Start-Job -ScriptBlock { param($p) Test-Path -LiteralPath $p } -ArgumentList $TargetPath
        $finished = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $finished) { return 'timeout' }
        $out = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
        if ($out.Count -gt 0 -and $out[0] -eq $true) { return 'reachable' }
        return 'missing'
    }
    catch { return 'unknown' }
    finally {
        if ($null -ne $job) { try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue } catch { } }
    }
}

# --------------------------------------------------------------------- HTML
function ConvertTo-HtmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function New-RawHtml {
    # A typed marker for a cell that is already markup. A string prefix would
    # mean any real value starting with those characters escapes the encoder.
    param([string]$Html)
    [pscustomobject]@{ PSTypeName = 'BackupReality.RawHtml'; Html = $Html }
}

function Add-Attention {
    param(
        [ValidateSet('Critical', 'Warning', 'Info')][string]$Level,
        [string]$Title,
        [string]$Detail
    )
    # Sort-Object in Windows PowerShell is not stable and has no -Stable switch,
    # so two Critical items would come out in an arbitrary order and the report
    # would shuffle between identical runs. The insertion index is the tiebreak.
    $Attention.Add([pscustomobject]@{ Level = $Level; Title = $Title; Detail = $Detail; Index = $Attention.Count })
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
    param([psobject[]]$Rows, [string[]]$Columns)

    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return '<p class="empty">Nothing to show.</p>'
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
            if ($null -ne $val -and $val.PSObject.TypeNames -contains 'BackupReality.RawHtml') {
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

function Add-Section {
    param([string]$Title, [string]$Body, [string]$Note = '')
    $noteHtml = ''
    if ($Note) { $noteHtml = '<p class="note">' + (ConvertTo-HtmlText $Note) + '</p>' }
    $Sections.Add('<section><h2>' + (ConvertTo-HtmlText $Title) + '</h2>' + $noteHtml + $Body + '</section>')
}

function New-LightPill {
    param([string]$Light)
    $class = 'red'
    if     ($Light -eq 'GREEN')  { $class = 'green' }
    elseif ($Light -eq 'YELLOW') { $class = 'amber' }
    New-RawHtml ('<span class="pill ' + $class + '">' + (ConvertTo-HtmlText $Light) + '</span>')
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host "=== Backup reality check : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) ==="
    Write-Host ''
}

$IsElevated = $false
try {
    $principal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
    $IsElevated = $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}
catch { }

# ================================================== 1. resolve the folders
Write-Step 'Resolving known folders...'

$InvocationError = $false
$FolderErrors = New-Object System.Collections.Generic.List[psobject]

$shellFolders = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -RawValues

# The Downloads folder has no [Environment+SpecialFolder] member, which is why
# it is addressed by its known-folder GUID here and nowhere else.
$knownFolderMap = @(
    [pscustomobject]@{ Label = 'Desktop';   ValueName = 'Desktop';                                  SpecialFolder = 'Desktop';     ProfileLeaf = 'Desktop' }
    [pscustomobject]@{ Label = 'Documents'; ValueName = 'Personal';                                 SpecialFolder = 'MyDocuments'; ProfileLeaf = 'Documents' }
    [pscustomobject]@{ Label = 'Pictures';  ValueName = 'My Pictures';                              SpecialFolder = 'MyPictures';  ProfileLeaf = 'Pictures' }
    [pscustomobject]@{ Label = 'Downloads'; ValueName = '{374DE290-123F-4565-9164-39C4925E467B}';   SpecialFolder = '';            ProfileLeaf = 'Downloads' }
)

$Targets  = New-Object System.Collections.Generic.List[psobject]
$seenKeys = @{}

function Add-Target {
    param(
        [string]$Label,
        [string]$ResolvedPath,
        [string]$Kind,
        [string]$Source,
        [string]$Confidence,
        [string]$RawValue = '',
        [string]$ValueKind = ''
    )
    $norm = ConvertTo-NormalPath $ResolvedPath
    if ([string]::IsNullOrWhiteSpace($norm)) { return $null }
    $key = Get-PathKey $norm
    if ($seenKeys.ContainsKey($key)) {
        # A -Path that names a folder already on the list, or a KFM layout where
        # two known folders collapse onto one place, should be one row.
        $existing = $seenKeys[$key]
        if ($existing.Label -ne $Label -and $Kind -ne 'Sync root') {
            $existing.Aliases.Add($Label) | Out-Null
        }
        return $null
    }
    $exists = $false
    try { $exists = Test-Path -LiteralPath $norm -PathType Container } catch { }

    $t = [pscustomobject]@{
        Label      = $Label
        Path       = $norm
        Key        = $key
        Kind       = $Kind
        Source     = $Source
        Confidence = $Confidence
        RawValue   = $RawValue
        ValueKind  = $ValueKind
        Exists     = $exists
        Aliases    = (New-Object System.Collections.Generic.List[string])
    }
    $seenKeys[$key] = $t
    $Targets.Add($t)
    $t
}

$knownResolved = 0

# The profile root is wanted even when the environment does not carry it: a
# scheduled task, a service account or a deliberately scrubbed shell can run
# with no USERPROFILE at all, and the shell API still answers correctly.
$profileRoot   = ''
$profileSource = ''
if ($env:USERPROFILE) {
    $profileRoot   = $env:USERPROFILE
    $profileSource = '%USERPROFILE%'
}
else {
    try { $profileRoot = [System.Environment]::GetFolderPath('UserProfile') } catch { }
    if ($profileRoot) { $profileSource = 'GetFolderPath(UserProfile)' }
}

foreach ($kf in $knownFolderMap) {
    $raw       = ''
    $kind      = ''
    $resolved  = ''
    $source    = ''
    $confidence = 'high'
    $registryNote = ''

    if ($shellFolders.Exists -and $shellFolders.Values.ContainsKey($kf.ValueName)) {
        $raw  = [string]$shellFolders.Values[$kf.ValueName]
        if ($shellFolders.Kinds.ContainsKey($kf.ValueName)) { $kind = [string]$shellFolders.Kinds[$kf.ValueName] }
        $resolved = $raw
        $source   = 'User Shell Folders'

        # These values are stored as %USERPROFILE%\Desktop and expanded by
        # whoever reads them. In a process whose environment has no
        # USERPROFILE the variable expands to itself, and carrying that
        # forward would report all four known folders as missing at an
        # address that has never existed on any machine.
        if ([string]::IsNullOrWhiteSpace((ConvertTo-NormalPath $resolved))) {
            $registryNote = 'the registry value ' + $raw + ' does not expand to a path in this process'
            $resolved = ''
            $source   = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolved) -and $kf.SpecialFolder) {
        try { $resolved = [System.Environment]::GetFolderPath($kf.SpecialFolder) } catch { }
        if ($resolved) { $source = 'GetFolderPath'; $confidence = 'medium' }
    }

    if ([string]::IsNullOrWhiteSpace($resolved) -and $profileRoot) {
        $resolved = Join-Path $profileRoot $kf.ProfileLeaf
        $source = $profileSource + ' fallback'
        $confidence = 'low'
    }

    if ($registryNote -and $source) {
        $source = $source + ', because ' + $registryNote
    }

    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $problem = 'could not be resolved from the registry or the environment'
        if ($registryNote) { $problem = $problem + ' - ' + $registryNote }
        $FolderErrors.Add([pscustomobject]@{ Label = $kf.Label; Path = ''; Problem = $problem })
        continue
    }

    # The registry and the shell API disagreeing is the signature of a
    # half-finished Known Folder Move, and it is worth saying out loud rather
    # than silently preferring one of them.
    if ($kf.SpecialFolder -and $source -eq 'User Shell Folders') {
        $viaApi = ''
        try { $viaApi = [System.Environment]::GetFolderPath($kf.SpecialFolder) } catch { }
        if ($viaApi -and (Get-PathKey (ConvertTo-NormalPath $viaApi)) -ne (Get-PathKey (ConvertTo-NormalPath $resolved))) {
            $confidence = 'medium'
            $source = 'User Shell Folders (disagrees with GetFolderPath: ' + $viaApi + ')'
        }
    }

    $t = Add-Target -Label $kf.Label -ResolvedPath $resolved -Kind 'Known folder' -Source $source -Confidence $confidence -RawValue $raw -ValueKind $kind
    if ($null -ne $t) {
        $knownResolved++
        if (-not $t.Exists) {
            $FolderErrors.Add([pscustomobject]@{ Label = $kf.Label; Path = $t.Path; Problem = 'the resolved path does not exist on disk' })
        }
    }
    else {
        # Add-Target returns nothing for two different reasons: this path is
        # already on the list under another label - a Known Folder Move can
        # collapse two of them onto one place, which is still a folder
        # resolved - or the path could not be made absolute, which is not.
        $dupeKey = Get-PathKey (ConvertTo-NormalPath $resolved)
        if ($dupeKey -and $seenKeys.ContainsKey($dupeKey)) { $knownResolved++ }
        else {
            $FolderErrors.Add([pscustomobject]@{ Label = $kf.Label; Path = $resolved; Problem = 'the value found for this folder is not a usable path' })
        }
    }
}

foreach ($p in @($Path)) {
    if ([string]::IsNullOrWhiteSpace($p)) { continue }
    $norm = ConvertTo-NormalPath $p
    if ([string]::IsNullOrWhiteSpace($norm)) {
        $FolderErrors.Add([pscustomobject]@{
            Label   = $p
            Path    = ''
            Problem = 'this -Path entry could not be turned into an absolute folder path: it names a drive that is not there, an environment variable that is not set on this machine, or a location outside the filesystem'
        })
        $InvocationError = $true
        continue
    }
    $exists = $false
    try { $exists = Test-Path -LiteralPath $norm -PathType Container } catch { }
    if (-not $exists) {
        $FolderErrors.Add([pscustomobject]@{ Label = $p; Path = $norm; Problem = 'this -Path entry does not exist' })
        $InvocationError = $true
        continue
    }
    [void](Add-Target -Label (Split-Path -Leaf $norm) -ResolvedPath $norm -Kind 'Extra path' -Source '-Path' -Confidence 'high')
}

if ($knownResolved -eq 0) {
    $InvocationError = $true
    Write-Line '  ! Not one known folder could be resolved.' 'Yellow'
}

# ================================================ 2. machine-wide mechanisms
Write-Step 'Taking the mechanism census...'

$SyncRoots = New-Object System.Collections.Generic.List[psobject]

function Add-SyncRoot {
    param(
        [string]$Provider,
        [string]$RootPath,
        [string]$State,
        [string]$Account = '',
        [string]$Discovery = '',
        $LastEvidenceUtc = $null,
        [string]$Confidence = 'medium'
    )
    $norm = ConvertTo-NormalPath $RootPath
    if ([string]::IsNullOrWhiteSpace($norm)) { return }
    $key = Get-PathKey $norm
    foreach ($existing in $SyncRoots) {
        if ($existing.Key -eq $key) {
            if ($Discovery -and $existing.Discovery -notlike ('*' + $Discovery + '*')) {
                $existing.Discovery = $existing.Discovery + ', ' + $Discovery
            }
            return
        }
    }
    $exists = $false
    try { $exists = Test-Path -LiteralPath $norm -PathType Container } catch { }

    $hasMarker = $false
    if ($exists) {
        try { $hasMarker = Test-Path -LiteralPath (Join-Path $norm '.849C9593-D756-4E56-8D6E-42412F2A707B') } catch { }
    }

    $SyncRoots.Add([pscustomobject]@{
        Provider        = $Provider
        Path            = $norm
        Key             = $key
        State           = $State
        Account         = $Account
        Discovery       = $Discovery
        LastEvidenceUtc = $LastEvidenceUtc
        Confidence      = $Confidence
        Exists          = $exists
        HasMarker       = $hasMarker
    })
}

# ---- OneDrive
$odRoot     = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\Microsoft\OneDrive'
$odAccounts = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\Microsoft\OneDrive\Accounts'

$odEverSignedIn = Get-RegNumber $odRoot 'ClientEverSignedIn' 0
$odLastUpdateUtc = $null
$odLastUpdateRaw = Get-RegNumber $odAccounts 'LastUpdate' $null
if ($null -ne $odLastUpdateRaw -and $odLastUpdateRaw -gt 0) {
    # QWord of unix seconds. Anything past year 3000 is a different unit
    # (milliseconds, or a FILETIME landing here by mistake) and is not worth
    # printing as a date.
    if ($odLastUpdateRaw -lt 32503680000) {
        try { $odLastUpdateUtc = $Epoch.AddSeconds([double]$odLastUpdateRaw) } catch { }
    }
}

$odProcessCount = 0
try { $odProcessCount = @(Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue).Count } catch { }

$syncRootManager = Read-RegistryKey -Hive 'HKLM' -SubKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SyncRootManager'
$srmNames = @($syncRootManager.SubKeys)
$srmOneDrive = @($srmNames | Where-Object { $_.ToLowerInvariant().StartsWith('onedrive') })

$syncEngines = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\SyncEngines\Providers'
$syncEngineMounts = New-Object System.Collections.Generic.List[psobject]
foreach ($engine in @($syncEngines.SubKeys)) {
    $engineKey = Read-RegistryKey -Hive 'HKCU' -SubKey ('Software\SyncEngines\Providers\' + $engine)
    foreach ($id in @($engineKey.SubKeys)) {
        $mountKey = Read-RegistryKey -Hive 'HKCU' -SubKey ('Software\SyncEngines\Providers\' + $engine + '\' + $id)
        $syncEngineMounts.Add([pscustomobject]@{
            Engine       = $engine
            Id           = $id
            MountPoint   = (Get-RegString $mountKey 'MountPoint')
            UrlNamespace = (Get-RegString $mountKey 'UrlNamespace')
            LibraryType  = (Get-RegString $mountKey 'LibraryType')
        })
    }
}

$odAccountRows = New-Object System.Collections.Generic.List[psobject]
foreach ($accName in @($odAccounts.SubKeys)) {
    $acc = Read-RegistryKey -Hive 'HKCU' -SubKey ('Software\Microsoft\OneDrive\Accounts\' + $accName)

    # Value names under an account key are inferred, not documented. Rather than
    # demanding UserFolder exists, anything named like a folder whose value
    # points at a real directory is accepted.
    $userFolder = ''
    foreach ($vn in @($acc.Values.Keys)) {
        if ([string]::IsNullOrWhiteSpace($vn)) { continue }
        if ($vn.ToLowerInvariant() -notlike '*folder*') { continue }
        $candidate = Get-RegString $acc $vn
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $normCandidate = ConvertTo-NormalPath $candidate
        $isDir = $false
        try { $isDir = Test-Path -LiteralPath $normCandidate -PathType Container } catch { }
        if ($isDir) { $userFolder = $normCandidate; break }
    }

    $kfmFlags = @()
    foreach ($vn in @($acc.Values.Keys)) {
        if ([string]::IsNullOrWhiteSpace($vn)) { continue }
        if ($vn.ToLowerInvariant() -like '*kfm*') { $kfmFlags += ('{0}={1}' -f $vn, (Get-RegString $acc $vn)) }
    }

    $odAccountRows.Add([pscustomobject]@{
        Account    = $accName
        UserEmail  = (Get-RegString $acc 'UserEmail')
        UserFolder = $userFolder
        KfmFlags   = ($kfmFlags -join '; ')
        ValueCount = @($acc.Values.Keys).Count
    })
}

# The whole point of this tool. An account subtree with entries plus either a
# running client or a registered sync root is the only combination that means
# "OneDrive is syncing". A sync folder on disk, %OneDrive% set and
# ClientEverSignedIn = 1 are all true on a machine that signed out a year ago.
$oneDriveState = 'absent'
$oneDriveWhy   = ''
$odHasAccounts = (@($odAccounts.SubKeys).Count -gt 0)
$odHasRegistration = ($odProcessCount -gt 0 -or $srmOneDrive.Count -gt 0)

$odCandidateRoots = New-Object System.Collections.Generic.List[string]
foreach ($v in @($env:OneDrive, $env:OneDriveConsumer, $env:OneDriveCommercial)) {
    if ($v) { $odCandidateRoots.Add($v) }
}
if ($env:USERPROFILE) {
    $odCandidateRoots.Add((Join-Path $env:USERPROFILE 'OneDrive'))
    try {
        Get-ChildItem -LiteralPath $env:USERPROFILE -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name.ToLowerInvariant().StartsWith('onedrive') } |
            ForEach-Object { $odCandidateRoots.Add($_.FullName) }
    }
    catch { }
}
foreach ($r in $odAccountRows) { if ($r.UserFolder) { $odCandidateRoots.Add($r.UserFolder) } }
foreach ($m in $syncEngineMounts) { if ($m.MountPoint) { $odCandidateRoots.Add($m.MountPoint) } }

$odRootsOnDisk = New-Object System.Collections.Generic.List[string]
$odSeen = @{}
foreach ($c in $odCandidateRoots) {
    $norm = ConvertTo-NormalPath $c
    if ([string]::IsNullOrWhiteSpace($norm)) { continue }
    $k = Get-PathKey $norm
    if ($odSeen.ContainsKey($k)) { continue }
    $odSeen[$k] = $true
    $onDisk = $false
    try { $onDisk = Test-Path -LiteralPath $norm -PathType Container } catch { }
    if ($onDisk) { $odRootsOnDisk.Add($norm) }
}

if ($odHasAccounts -and $odHasRegistration) {
    $oneDriveState = 'active'
    $oneDriveWhy = '{0} account subkey(s) present and {1}' -f @($odAccounts.SubKeys).Count, $(if ($odProcessCount -gt 0) { 'the OneDrive client is running' } else { 'a sync root is registered with Explorer' })
}
elseif ($odHasAccounts) {
    $oneDriveState = 'stopped'
    $oneDriveWhy = '{0} account subkey(s) present, but the OneDrive client is not running and Explorer has no registered sync root' -f @($odAccounts.SubKeys).Count
}
elseif ($odRootsOnDisk.Count -gt 0 -or $odEverSignedIn -eq 1) {
    $oneDriveState = 'zombie'
    $oneDriveWhy = 'no account subkeys under HKCU:\Software\Microsoft\OneDrive\Accounts, no OneDrive process and no registered sync root, yet the sync folder is still on disk'
    if ($odEverSignedIn -eq 1) { $oneDriveWhy = $oneDriveWhy + ' and ClientEverSignedIn is 1' }
}
else {
    $oneDriveWhy = 'no OneDrive registry key, folder or process found'
}

foreach ($r in $odRootsOnDisk) {
    $acct = ''
    foreach ($a in $odAccountRows) { if ($a.UserFolder -and (Get-PathKey $a.UserFolder) -eq (Get-PathKey $r)) { $acct = $a.Account } }
    Add-SyncRoot -Provider 'OneDrive' -RootPath $r -State $oneDriveState -Account $acct `
        -Discovery 'folder on disk' -LastEvidenceUtc $odLastUpdateUtc -Confidence 'high'
}

# ---- Dropbox
$dropboxProcCount = 0
try { $dropboxProcCount = @(Get-Process -Name 'Dropbox' -ErrorAction SilentlyContinue).Count } catch { }
$dropboxReg = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\Dropbox'
$dropboxInfoPaths = @()
foreach ($base in @($env:APPDATA, $env:LOCALAPPDATA)) {
    if ($base) { $dropboxInfoPaths += (Join-Path $base 'Dropbox\info.json') }
}
$dropboxAccounts = New-Object System.Collections.Generic.List[psobject]
$dropboxNote = ''
$dropboxInfoFound = $false
foreach ($ip in $dropboxInfoPaths) {
    $found = $false
    try { $found = Test-Path -LiteralPath $ip -PathType Leaf } catch { }
    if (-not $found) { continue }
    $dropboxInfoFound = $true
    try {
        $json = Get-Content -LiteralPath $ip -Raw -ErrorAction Stop | ConvertFrom-Json
        foreach ($prop in @($json.PSObject.Properties)) {
            $dropboxAccounts.Add([pscustomobject]@{
                Kind = $prop.Name
                Path = [string](Get-Prop $prop.Value 'path' '')
                File = $ip
            })
        }
    }
    catch {
        # A truncated or half-written info.json is weak evidence that Dropbox
        # was here, not a reason to stop the report.
        $dropboxNote = 'info.json is present but could not be parsed: ' + $_.Exception.Message
    }
}

$dropboxState = 'absent'
if ($dropboxInfoFound -or $dropboxReg.Exists -or $dropboxProcCount -gt 0) {
    if ($dropboxProcCount -gt 0) { $dropboxState = 'active' } else { $dropboxState = 'stopped' }
}
foreach ($a in $dropboxAccounts) {
    if ([string]::IsNullOrWhiteSpace($a.Path)) { continue }
    $state = $dropboxState
    $onDisk = $false
    try { $onDisk = Test-Path -LiteralPath (ConvertTo-NormalPath $a.Path) -PathType Container } catch { }
    if (-not $onDisk) { $state = 'missing' }
    Add-SyncRoot -Provider 'Dropbox' -RootPath $a.Path -State $state -Account $a.Kind `
        -Discovery 'info.json' -LastEvidenceUtc $(if ($dropboxProcCount -gt 0) { $NowUtc } else { $null }) -Confidence 'medium'
}

# ---- Google Drive
$gdriveProcCount = 0
try { $gdriveProcCount = @(Get-Process -Name 'GoogleDriveFS' -ErrorAction SilentlyContinue).Count } catch { }
$gdriveReg = Read-RegistryKey -Hive 'HKCU' -SubKey 'Software\Google\DriveFS'
$gdriveDataDir = ''
if ($env:LOCALAPPDATA) { $gdriveDataDir = Join-Path $env:LOCALAPPDATA 'Google\DriveFS' }
$gdriveDataFound = $false
if ($gdriveDataDir) { try { $gdriveDataFound = Test-Path -LiteralPath $gdriveDataDir -PathType Container } catch { } }

$gdriveVolumes = @()
try {
    # Drive File Stream mounts as a virtual volume under whatever letter is
    # free, so the letter is not a constant. The label is; the file system is
    # not NTFS and filtering on it would drop the volume entirely.
    $gdriveVolumes = @(Get-CimInstance -ClassName Win32_LogicalDisk -OperationTimeoutSec 20 -ErrorAction Stop |
        Where-Object { ([string](Get-Prop $_ 'VolumeName' '')).ToLowerInvariant() -like '*google drive*' })
}
catch { }

$gdriveState = 'absent'
if ($gdriveDataFound -or $gdriveReg.Exists -or $gdriveProcCount -gt 0 -or $gdriveVolumes.Count -gt 0) {
    if ($gdriveProcCount -gt 0) { $gdriveState = 'active' } else { $gdriveState = 'stopped' }
}
foreach ($v in $gdriveVolumes) {
    $letter = [string](Get-Prop $v 'DeviceID' '')
    if ($letter) {
        Add-SyncRoot -Provider 'Google Drive' -RootPath ($letter + '\') -State $gdriveState `
            -Account ([string](Get-Prop $v 'VolumeName' '')) -Discovery 'mounted volume' `
            -LastEvidenceUtc $(if ($gdriveProcCount -gt 0) { $NowUtc } else { $null }) -Confidence 'low'
    }
}

# ---- File History
Write-Step 'Reading File History configuration...'

$fhBase = ''
if ($env:LOCALAPPDATA) { $fhBase = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\FileHistory' }
$fhConfigDir = ''
if ($fhBase) { $fhConfigDir = Join-Path $fhBase 'Configuration' }
$fhDataDir = ''
if ($fhBase) { $fhDataDir = Join-Path $fhBase 'Data' }

$fh = [pscustomobject]@{
    Configured    = $false
    ConfigFile    = ''
    ConfigUtc     = $null
    Inclusions    = @()
    Exclusions    = @()
    TargetName    = ''
    TargetUrl     = ''
    TargetType    = ''
    TargetStatus  = 'unknown'
    CatalogUtc    = $null
    ServiceStatus = 'not present'
    ServiceStart  = ''
    ParseNote     = ''
}

try {
    $svc = Get-Service -Name 'fhsvc' -ErrorAction Stop
    # Matched by service Name. DisplayName is localized and matching on it
    # breaks on every non-English Windows.
    $fh.ServiceStatus = [string]$svc.Status
    $fh.ServiceStart  = [string]$svc.StartType
}
catch { }

$fhConfigDirExists = $false
if ($fhConfigDir) { try { $fhConfigDirExists = Test-Path -LiteralPath $fhConfigDir -PathType Container } catch { } }

if ($fhConfigDirExists) {
    $newestConfig = $null
    try {
        $newestConfig = Get-ChildItem -LiteralPath $fhConfigDir -Filter 'Config*.xml' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    }
    catch { }

    if ($null -ne $newestConfig) {
        $fh.ConfigFile = $newestConfig.FullName
        $fh.ConfigUtc  = $newestConfig.LastWriteTimeUtc
        try {
            $doc = New-Object System.Xml.XmlDocument
            $doc.Load($newestConfig.FullName)

            $fh.Inclusions = Get-ConfigPathList -Doc $doc -ListName 'FolderInclusionList'
            $fh.Exclusions = Get-ConfigPathList -Doc $doc -ListName 'FolderExclusionList'

            foreach ($pair in @(
                @('TargetName', 'TargetName'),
                @('TargetUrl', 'TargetUrl'),
                @('TargetDriveType', 'TargetType')
            )) {
                $node = $doc.SelectSingleNode("//*[local-name()='Target']//*[local-name()='" + $pair[0] + "']")
                if ($null -ne $node) {
                    $value = ([string]$node.InnerText).Trim()
                    $fh.($pair[1]) = $value
                }
            }

            $fh.Configured = (@($fh.Inclusions).Count -gt 0 -or $fh.TargetUrl -ne '')
        }
        catch {
            $fh.ParseNote = 'the configuration XML is present but could not be parsed: ' + $_.Exception.Message
        }
    }
}

if ($fh.TargetUrl) {
    $fh.TargetStatus = Test-TargetReachable -TargetPath $fh.TargetUrl -TimeoutSeconds 5
}

$fhDataDirExists = $false
if ($fhDataDir) { try { $fhDataDirExists = Test-Path -LiteralPath $fhDataDir -PathType Container } catch { } }
if ($fhDataDirExists) {
    try {
        $catalog = Get-ChildItem -LiteralPath $fhDataDir -Filter 'Catalog*.edb' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
        if ($null -ne $catalog) { $fh.CatalogUtc = $catalog.LastWriteTimeUtc }
    }
    catch { }
}

$fhCoreChannel = Get-ChannelInfo 'Microsoft-Windows-FileHistory-Core/WHC'
$fhCoreNewestUtc = $null
if ($fhCoreChannel.Exists -and -not $fhCoreChannel.Denied -and $fhCoreChannel.RecordCount -gt 0) {
    $events = @(Get-ChannelEvent -LogName $fhCoreChannel.Name -MaxEvents 200)
    if ($events.Count -gt 0) {
        $t = Get-Prop $events[0] 'TimeCreated'
        if ($null -ne $t) { $fhCoreNewestUtc = ([datetime]$t).ToUniversalTime() }
    }
}

$fhEngineChannel = Get-ChannelInfo 'Microsoft-Windows-FileHistory-Engine/BackupLog'
$fhEngineNewestUtc = $null
if ($fhEngineChannel.Denied) {
    Add-AdminSkip 'Microsoft-Windows-FileHistory-Engine/BackupLog' 'the per-run File History backup log could not be read, so "last successful run" rests on the catalog file time and the Core channel'
}
elseif ($fhEngineChannel.Exists -and $fhEngineChannel.RecordCount -gt 0) {
    $events = @(Get-ChannelEvent -LogName $fhEngineChannel.Name -MaxEvents 200)
    if ($events.Count -gt 0) {
        $t = Get-Prop $events[0] 'TimeCreated'
        if ($null -ne $t) { $fhEngineNewestUtc = ([datetime]$t).ToUniversalTime() }
    }
}

# Best available "File History actually ran" timestamp, and which source it came
# from. The source matters: a per-run log entry says a run happened and how it
# ended, while the catalog file's timestamp only says something touched it -
# a failed run touches it too.
$fhLastRunUtc = $null
$fhEvidenceSource = ''
foreach ($candidate in @(
    @($fhEngineNewestUtc, 'the per-run File History log'),
    @($fhCoreNewestUtc,   'the File History Core event channel'),
    @($fh.CatalogUtc,     'the File History catalog file')
)) {
    if ($null -eq $candidate[0]) { continue }
    if ($null -eq $fhLastRunUtc -or $candidate[0] -gt $fhLastRunUtc) {
        $fhLastRunUtc     = $candidate[0]
        $fhEvidenceSource = [string]$candidate[1]
    }
}

# ---- machine context (never changes a folder's light)
Write-Step 'Reading machine context...'

$srReg = Read-RegistryKey -Hive 'HKLM' -SubKey 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
$srInterval = Get-RegNumber $srReg 'RPSessionInterval' $null
$srInitDone = Get-RegNumber $srReg 'SRInitDone' $null
$srLastMaintUtc = $null
# The value name really is misspelled in the registry. Correcting it here reads
# it back as absent.
$srLastMaintRaw = Get-RegNumber $srReg 'LastMainenanceTaskRunTimeStamp' $null
if ($null -ne $srLastMaintRaw -and $srLastMaintRaw -gt 0) {
    try { $srLastMaintUtc = [datetime]::FromFileTimeUtc([int64]$srLastMaintRaw) } catch { }
}

# Both probes below are attempted on every run rather than assumed from the
# elevation check. A row saying a check was "skipped: needs admin" tells the
# reader that running as administrator would buy them that answer, and the
# only way to be entitled to say it is to have asked and been refused. Neither
# call is expensive when it fails: without rights the restore-point query
# answers "Access denied" and the shadow-copy query answers "Initialization
# failure", both in well under a second.
$skipStatus = 'skipped: needs admin'
if ($IsElevated) { $skipStatus = 'attempted, failed' }

$srPointCount = $null
$srNewestUtc  = $null
try {
    $srPoints = @(Get-CimInstance -Namespace 'root/default' -ClassName 'SystemRestore' -OperationTimeoutSec 20 -ErrorAction Stop)
    $srPointCount = $srPoints.Count
    foreach ($rp in $srPoints) {
        $when = ConvertTo-UtcDate (Get-Prop $rp 'CreationTime')
        if ($null -ne $when -and ($null -eq $srNewestUtc -or $when -gt $srNewestUtc)) { $srNewestUtc = $when }
    }
}
catch {
    Add-AdminSkip 'CIM root/default:SystemRestore' `
        ('the list of restore points could not be read ("{0}"), so the registry values above are all this run has on System Restore' -f $_.Exception.Message.Trim()) `
        -Status $skipStatus
}

$shadowCount     = $null
$shadowNewestUtc = $null
try {
    $shadows = @(Get-CimInstance -ClassName 'Win32_ShadowCopy' -OperationTimeoutSec 20 -ErrorAction Stop)
    $shadowCount = $shadows.Count
    foreach ($sc in $shadows) {
        $when = ConvertTo-UtcDate (Get-Prop $sc 'InstallDate')
        if ($null -ne $when -and ($null -eq $shadowNewestUtc -or $when -gt $shadowNewestUtc)) { $shadowNewestUtc = $when }
    }
}
catch {
    Add-AdminSkip 'Win32_ShadowCopy' `
        ('shadow copies could not be enumerated ("{0}" is what this class answers without rights, however much it reads like a fault), so whether Previous Versions holds an older copy of these folders is unknown' -f $_.Exception.Message.Trim()) `
        -Status $skipStatus
}

$settingsChannel = Get-ChannelInfo 'Microsoft-Windows-UserSettingsBackup-Orchestrator/Operational'
$settingsLastUtc = $null
$settingsResult  = ''
if ($settingsChannel.Exists -and -not $settingsChannel.Denied -and $settingsChannel.RecordCount -gt 0) {
    $events = @(Get-ChannelEvent -LogName $settingsChannel.Name -Id @(10) -MaxEvents 1)
    if ($events.Count -gt 0) {
        $t = Get-Prop $events[0] 'TimeCreated'
        if ($null -ne $t) { $settingsLastUtc = ([datetime]$t).ToUniversalTime() }
        $map = Get-EventDataMap $events[0]
        if ($map.ContainsKey('OverallResult')) { $settingsResult = [string]$map['OverallResult'] }
        if ([string]::IsNullOrWhiteSpace($settingsResult)) { $settingsResult = 'not stated in the event' }
    }
}

$backupChannel = Get-ChannelInfo 'Microsoft-Windows-Backup'

# ============================================== 3. assess each folder
Write-Step 'Assessing folders...'

# A folder carrying the sync-root marker file is being managed by some sync
# client whether or not the census recognised that client. Walking up from each
# assessed folder catches a root nothing else revealed: a second tenant, a root
# moved off the profile, a provider this tool has never heard of.
foreach ($target in @($Targets)) {
    if (-not $target.Exists) { continue }
    $probe = $target.Path
    $steps = 0
    while ($probe -and $steps -lt 64) {
        $steps++
        $markerFound = $false
        try { $markerFound = Test-Path -LiteralPath (Join-Path $probe '.849C9593-D756-4E56-8D6E-42412F2A707B') -PathType Leaf } catch { }
        if ($markerFound) {
            $alreadyKnown = $false
            foreach ($r in $SyncRoots) { if ($r.Key -eq (Get-PathKey $probe)) { $alreadyKnown = $true; break } }
            if (-not $alreadyKnown) {
                # A marker with no live provider behind it is the zombie shape,
                # whether or not OneDrive was ever installed here.
                $markerState = $oneDriveState
                if ($markerState -eq 'absent') { $markerState = 'zombie' }
                Add-SyncRoot -Provider 'OneDrive-like marker' -RootPath $probe -State $markerState `
                    -Discovery 'sync-root marker file' -LastEvidenceUtc $odLastUpdateUtc -Confidence 'low'
            }
            break
        }
        $parent = Split-Path -Parent $probe
        if (-not $parent -or $parent -eq $probe) { break }
        $probe = $parent
    }
}

# A sync root nobody's assessed folder lives inside is still a folder full of
# the user's files, and on a signed-out machine it is the one place the
# placeholder hazard actually bites. It gets its own row.
foreach ($root in $SyncRoots) {
    if (-not $root.Exists) { continue }
    $covered = $false
    foreach ($t in $Targets) {
        if (Test-PathInside -Child $t.Path -Root $root.Path) { $covered = $true; break }
    }
    if ($covered) { continue }
    [void](Add-Target -Label ($root.Provider + ' sync root') -ResolvedPath $root.Path -Kind 'Sync root' `
        -Source ($root.Provider + ' - ' + $root.Discovery) -Confidence $root.Confidence)
}

$lightRank = @{ 'GREEN' = 0; 'YELLOW' = 1; 'RED' = 2 }

$Assessments = New-Object System.Collections.Generic.List[psobject]

foreach ($target in $Targets) {
    if (-not $target.Exists) {
        $Assessments.Add([pscustomobject]@{
            Target          = $target
            Light           = 'RED'
            Mechanism       = 'folder missing'
            Detail          = 'This folder does not exist on disk, so nothing here could be assessed.'
            LastEvidenceUtc = $null
            Sample          = $null
            Notes           = @()
        })
        continue
    }

    $coverage = New-Object System.Collections.Generic.List[psobject]
    $hazards  = New-Object System.Collections.Generic.List[psobject]
    $notes    = New-Object System.Collections.Generic.List[string]

    # ---- (a) containment in a sync root
    foreach ($root in $SyncRoots) {
        if (-not $root.Exists) { continue }
        if (-not (Test-PathInside -Child $target.Path -Root $root.Path)) { continue }

        $who = $root.Provider
        if ($root.Account) { $who = $who + ' (' + $root.Account + ')' }

        # The sync root itself is one of the assessed locations, so the sentence
        # has to work for "this folder IS the root" as well as "this folder is
        # somewhere under it".
        $placement = ('{0} physically lives inside the {1} sync root at {2}' -f $target.Label, $who, $root.Path)
        if ((Get-PathKey $target.Path) -eq $root.Key) {
            $placement = ('This is the {0} sync root at {1}' -f $who, $root.Path)
        }

        switch ($root.State) {
            'active' {
                $coverage.Add([pscustomobject]@{
                    Mechanism = $root.Provider + ' sync root'
                    Light     = 'GREEN'
                    WhenUtc   = $root.LastEvidenceUtc
                    Detail    = ('{0}, and the provider shows signs of running.' -f $placement)
                    Order     = 0
                })
            }
            'stopped' {
                $coverage.Add([pscustomobject]@{
                    Mechanism = $root.Provider + ' sync root (provider not running)'
                    Light     = 'YELLOW'
                    WhenUtc   = $root.LastEvidenceUtc
                    Detail    = ('{0}, but the provider is not running, so nothing is being uploaded right now.' -f $placement)
                    Order     = 0
                })
            }
            'missing' {
                $coverage.Add([pscustomobject]@{
                    Mechanism = $root.Provider + ' (target missing)'
                    Light     = 'YELLOW'
                    WhenUtc   = $root.LastEvidenceUtc
                    Detail    = ('{0} is configured for {1}, but the configured folder is not on this machine.' -f $target.Label, $who)
                    Order     = 0
                })
            }
            'zombie' {
                $hazards.Add([pscustomobject]@{
                    Mechanism = $root.Provider + ' sync root, signed out'
                    Light     = 'RED'
                    WhenUtc   = $root.LastEvidenceUtc
                    Detail    = ('{0}, and that sync is dead: files here may be cloud-only placeholders that cannot be opened - this is worse than no backup. Last account activity {1}.' -f `
                                    $placement, (Format-AgeText $root.LastEvidenceUtc))
                    Order     = 0
                })
            }
            default {
                $coverage.Add([pscustomobject]@{
                    Mechanism = $root.Provider + ' sync root (state unclear)'
                    Light     = 'YELLOW'
                    WhenUtc   = $root.LastEvidenceUtc
                    Detail    = ('{0}, but whether it is syncing could not be established.' -f $placement)
                    Order     = 0
                })
            }
        }
    }

    # A folder of the same name sitting inside a dead sync root is the exact
    # trap this tool exists for: it looks like the backup and is not.
    foreach ($root in $SyncRoots) {
        if (-not $root.Exists -or $root.State -ne 'zombie') { continue }
        if (Test-PathInside -Child $target.Path -Root $root.Path) { continue }
        if ($target.Kind -ne 'Known folder') { continue }
        $shadow = Join-Path $root.Path $target.Label
        $shadowExists = $false
        try { $shadowExists = Test-Path -LiteralPath $shadow -PathType Container } catch { }
        if ($shadowExists) {
            $notes.Add(("A folder named '{0}' also exists inside the dead {1} sync root at {2}. That is not this folder and it is not a backup of it." -f $target.Label, $root.Provider, $root.Path))
        }
    }

    # ---- (b) File History inclusion
    if ($fh.Configured) {
        $included = $false
        $excluded = $false
        foreach ($inc in @($fh.Inclusions)) {
            if (Test-PathInside -Child $target.Path -Root (ConvertTo-NormalPath $inc)) { $included = $true; break }
        }
        foreach ($exc in @($fh.Exclusions)) {
            if (Test-PathInside -Child $target.Path -Root (ConvertTo-NormalPath $exc)) { $excluded = $true; break }
        }

        if ($included -and -not $excluded) {
            $fhLight  = 'YELLOW'
            $fhDetail = ''
            if ($fh.TargetStatus -eq 'timeout') {
                $fhDetail = ('File History lists {0}, but its target {1} did not answer within 5 seconds - it may be a drive that is not plugged in or a NAS that is off.' -f $target.Label, $fh.TargetUrl)
            }
            elseif ($fh.TargetStatus -eq 'missing') {
                $fhDetail = ('File History lists {0}, but its target {1} is not reachable from this machine.' -f $target.Label, $fh.TargetUrl)
            }
            elseif ($null -eq $fhLastRunUtc) {
                $fhDetail = ('File History is configured and lists {0}, but nothing on this machine shows that it has ever run.' -f $target.Label)
            }
            else {
                $fhLight  = 'GREEN'
                $fhDetail = ('File History is configured, lists {0}, and {1} was last written {2}.' -f `
                                $target.Label, $fhEvidenceSource, (Format-AgeText $fhLastRunUtc))
            }
            $coverage.Add([pscustomobject]@{
                Mechanism = 'File History'
                Light     = $fhLight
                WhenUtc   = $fhLastRunUtc
                Detail    = $fhDetail
                Order     = 1
            })
        }
        elseif ($excluded) {
            $notes.Add('File History is configured on this machine but this folder is on its exclusion list.')
        }
    }

    # ---- (c) bounded sample
    if ($target.Kind -eq 'Sync root') {
        Write-Step ("Scanning {0} - a stopped sync provider makes this slow..." -f $target.Path)
    }
    $sample = Get-FolderSample -FolderPath $target.Path -Limit $SampleLimit

    if ($sample.Ok -and $sample.FileCount -eq 0) {
        $notes.Add('No files were found in this folder, so there is nothing here to lose.')
    }
    if ($sample.CloudOnly -gt 0) {
        $notes.Add(("{0} of the {1} files sampled are cloud-only placeholders ({2} of names with no local contents)." -f `
            $sample.CloudOnly, $sample.FileCount, (Format-Size $sample.CloudOnlyBytes)))
    }
    elseif ($sample.Offline -gt 0) {
        # Offline without the cloud-only bit is the older shape: something has
        # moved the contents elsewhere without claiming to be a sync provider.
        $notes.Add(("{0} of the {1} files sampled are marked offline, so their contents may not be on this disk." -f `
            $sample.Offline, $sample.FileCount))
    }
    if ($sample.ElapsedSeconds -ge 5) {
        $notes.Add(("Walking this folder took {0} seconds - the usual cause is a sync provider that is not running, leaving every file waiting on an answer that never comes." -f $sample.ElapsedSeconds))
    }
    if ($sample.Truncated) {
        $notes.Add(("The walk stopped at the -SampleLimit of {0} files; counts above describe the sample, not the whole folder." -f $SampleLimit))
    }

    # ---- decide
    # Coverage takes the best light available - a stale second mechanism should
    # not drag down a working first one. A hazard is not coverage and cannot be
    # cancelled by one: a second backup does not make an unreadable file
    # readable.
    $light     = 'RED'
    $mechanism = 'none found'
    $detail    = ''
    $whenUtc   = $null

    if ($hazards.Count -gt 0) {
        $worst = @($hazards | Sort-Object @{ Expression = { $_.Order } })[0]
        $light     = 'RED'
        $mechanism = $worst.Mechanism
        $detail    = $worst.Detail
        $whenUtc   = $worst.WhenUtc
    }
    elseif ($coverage.Count -gt 0) {
        $ordered = @($coverage | Sort-Object @{ Expression = { $lightRank[$_.Light] } }, @{ Expression = { $_.Order } })
        $best = $ordered[0]
        $light     = $best.Light
        $mechanism = $best.Mechanism
        $detail    = $best.Detail
        $whenUtc   = $best.WhenUtc

        if ($light -eq 'GREEN') {
            $age = Get-AgeDays $whenUtc
            if ($null -eq $age) {
                $light  = 'YELLOW'
                $detail = $detail + ' No date could be put on that evidence, so how recent it is cannot be shown.'
            }
            elseif ($age -lt -1) {
                $light  = 'YELLOW'
                $detail = $detail + ' The date on that evidence is in the future, so this machine''s clock cannot be trusted.'
            }
            elseif ($age -gt $StaleDays) {
                $light  = 'YELLOW'
                $detail = $detail + (' That is older than the {0}-day freshness limit, so it is reported as stale.' -f $StaleDays)
            }
        }
    }
    else {
        $detail = ('Nothing on this machine shows any sign of {0} being backed up: it is not inside any sync root, and File History is not covering it.' -f $target.Label)
    }

    # Only an admin skip that bears on the deciding evidence downgrades the
    # light. A folder inside a running OneDrive is not made less certain by an
    # unreadable File History log, and downgrading on any admin skip at all
    # would put GREEN out of reach of every non-elevated run.
    #
    # The catalog file is the weak case: its timestamp moves when File History
    # runs and fails as readily as when it runs and succeeds, and the log that
    # tells the two apart is the one that needs administrator rights.
    if ($light -eq 'GREEN' -and $mechanism -eq 'File History' -and
        $fhEngineChannel.Denied -and $fhEvidenceSource -eq 'the File History catalog file') {
        $light  = 'YELLOW'
        $detail = $detail + ' That timestamp moves whether a run succeeds or fails, and the per-run log that would tell them apart needs administrator rights, so this run cannot say the last backup finished.'
    }

    $Assessments.Add([pscustomobject]@{
        Target          = $target
        Light           = $light
        Mechanism       = $mechanism
        Detail          = $detail
        LastEvidenceUtc = $whenUtc
        Sample          = $sample
        Notes           = @($notes)
    })
}

# ============================================== 4. verdict and exit code
$reds    = @($Assessments | Where-Object { $_.Light -eq 'RED' })
$yellows = @($Assessments | Where-Object { $_.Light -eq 'YELLOW' })

$ExitCode = 0
if     ($yellows.Count -gt 0) { $ExitCode = 1 }
if     ($reds.Count    -gt 0) { $ExitCode = 2 }
if     ($InvocationError)     { $ExitCode = 3 }

$folderAssessments = @($Assessments | Where-Object { $_.Target.Kind -ne 'Sync root' })
# The count and the named folders have to add up to the same set. Counting only
# GREEN while naming only RED leaves a YELLOW folder in neither: the banner
# would say nothing is backed up while the table below it showed a mechanism,
# and would name three folders out of four as bad, which invites the reader to
# assume the fourth is fine.
$folderCovered = @($folderAssessments | Where-Object { $_.Light -ne 'RED' })
$folderWeak    = @($folderAssessments | Where-Object { $_.Light -eq 'YELLOW' })
$folderBad     = @($folderAssessments | Where-Object { $_.Light -eq 'RED' })

$folderWord = 'folders'
if ($folderAssessments.Count -eq 1) { $folderWord = 'folder' }

$bannerLines = New-Object System.Collections.Generic.List[string]
$bannerLines.Add(('{0} of {1} {2} show signs of backup.' -f $folderCovered.Count, $folderAssessments.Count, $folderWord))
if ($folderWeak.Count -gt 0) {
    $bannerLines.Add((Join-Readable @($folderWeak | ForEach-Object { $_.Target.Label })) + ': signs are there, but stale, weak or not verifiable from here.')
}
if ($folderBad.Count -gt 0) {
    $bannerLines.Add((Join-Readable @($folderBad | ForEach-Object { $_.Target.Label })) + ': no signs of any backup you can rely on.')
}
$bannerLines.Add('This tool reports evidence; it cannot certify that a backup works.')
$bannerText = $bannerLines -join ' '

$bannerClass = 'good'
if ($yellows.Count -gt 0) { $bannerClass = 'warn' }
if ($reds.Count -gt 0)    { $bannerClass = 'bad' }

foreach ($a in $reds) {
    Add-Attention 'Critical' ($a.Target.Label + ': ' + $a.Mechanism) $a.Detail
}
foreach ($a in $yellows) {
    Add-Attention 'Warning' ($a.Target.Label + ': ' + $a.Mechanism) $a.Detail
}
foreach ($e in $FolderErrors) {
    Add-Attention 'Warning' ('Could not assess ' + $e.Label) ($e.Problem + ' (' + $e.Path + ')')
}
# A provider configured against a folder that is not on this machine is the
# unplugged-external-drive case. It is not any one folder's problem, so it is
# reported here rather than as a light.
foreach ($root in $SyncRoots) {
    if ($root.Exists) { continue }
    Add-Attention 'Warning' ($root.Provider + ' points at a folder that is not here') `
        ("{0} names {1} as its sync folder, but there is nothing at that path - a drive that is not plugged in, or a root that has been moved." -f $root.Provider, $root.Path)
}
if ($oneDriveState -eq 'zombie') {
    Add-Attention 'Critical' 'OneDrive is signed out but its folder is still here' `
        ('OneDrive last showed account activity {0}. Files inside {1} may be cloud-only placeholders: the names and sizes are on disk, the contents are not, and with no signed-in client there is nothing to fetch them.' -f `
            (Format-AgeText $odLastUpdateUtc), (Join-Readable @($odRootsOnDisk)))
}

# ============================================== 5. history CSV
$historyRows = New-Object System.Collections.Generic.List[psobject]
foreach ($a in $Assessments) {
    $filesSampled  = ''
    $cloudOnly     = ''
    $truncated     = ''
    if ($null -ne $a.Sample) {
        $filesSampled = $a.Sample.FileCount
        $cloudOnly    = $a.Sample.CloudOnly
        $truncated    = $a.Sample.Truncated
    }
    $historyRows.Add([pscustomobject]@{
        TimestampUtc    = (Format-IsoUtc $NowUtc)
        Folder          = $a.Target.Label
        ResolvedPath    = $a.Target.Path
        Light           = $a.Light
        Mechanism       = $a.Mechanism
        LastEvidenceUtc = (Format-IsoUtc $a.LastEvidenceUtc)
        EvidenceDetail  = $a.Detail
        FilesSampled    = $filesSampled
        CloudOnlyCount  = $cloudOnly
        SampleTruncated = $truncated
        ExitCode        = $ExitCode
        ToolVersion     = $ToolVersion
    })
}

$historyOk = $true
try { Add-HistoryRow -Rows @($historyRows) -Path $HistoryCsv -MaxRows $HistoryLimit }
catch {
    $historyOk = $false
    Write-Line ("  ! Could not update {0}: {1}" -f $HistoryCsv, $_.Exception.Message) 'Yellow'
}

# ============================================== 6. build the report
Write-Step 'Rendering the report...'

$folderRows = New-Object System.Collections.Generic.List[psobject]
foreach ($a in $Assessments) {
    $filesSampled = '-'
    $cloudOnly    = '-'
    $newest       = '-'
    if ($null -ne $a.Sample -and $a.Sample.Ok) {
        $filesSampled = [string]$a.Sample.FileCount
        if ($a.Sample.Truncated) { $filesSampled = $filesSampled + '+' }
        $cloudOnly = [string]$a.Sample.CloudOnly
        if ($null -ne $a.Sample.NewestUtc) { $newest = (Format-Utc $a.Sample.NewestUtc) }
    }
    $locationLabel = $a.Target.Label
    if (@($a.Target.Aliases).Count -gt 0) {
        $locationLabel = $locationLabel + ' (also ' + (Join-Readable @($a.Target.Aliases)) + ')'
    }
    $folderRows.Add([pscustomobject]@{
        'Location'      = $locationLabel
        'Light'         = (New-LightPill $a.Light)
        'Mechanism'     = $a.Mechanism
        'Last evidence' = $(if ($null -eq $a.LastEvidenceUtc) { 'none' } else { Format-Utc $a.LastEvidenceUtc })
        'Files sampled' = $filesSampled
        'Cloud-only'    = $cloudOnly
        'Newest file'   = $newest
        'Path'          = $a.Target.Path
    })
}
Add-Section 'Folders' (New-DataTable -Rows @($folderRows)) `
    'A light describes the evidence found, not a guarantee. "Cloud-only" counts files whose contents are not on this disk.'

$evidenceRows = New-Object System.Collections.Generic.List[psobject]
foreach ($a in $Assessments) {
    $why = $a.Detail
    foreach ($n in @($a.Notes)) { $why = $why + ' ' + $n }
    # The raw registry value is worth showing next to the resolved path: a
    # Known Folder Move rewrites it to point into the sync folder, and that
    # rewrite is invisible once the value has been expanded.
    $resolvedFrom = $a.Target.Source + ' (confidence: ' + $a.Target.Confidence + ')'
    if ($a.Target.RawValue) {
        $resolvedFrom = $resolvedFrom + ', from the registry value ' + $a.Target.RawValue
        if ($a.Target.ValueKind) { $resolvedFrom = $resolvedFrom + ' [' + $a.Target.ValueKind + ']' }
    }
    $evidenceRows.Add([pscustomobject]@{
        'Location' = $a.Target.Label
        'Light'    = (New-LightPill $a.Light)
        'Evidence' = $why
        'How the path was resolved' = $resolvedFrom
    })
}
Add-Section 'What the evidence actually says' (New-DataTable -Rows @($evidenceRows))

if ($FolderErrors.Count -gt 0) {
    $errRows = @($FolderErrors | ForEach-Object {
        [pscustomobject]@{ 'Folder' = $_.Label; 'Path' = $_.Path; 'Problem' = $_.Problem }
    })
    Add-Section 'Folders that could not be assessed' (New-DataTable -Rows $errRows)
}

# ---- mechanism detail
$markerRoots = @($SyncRoots | Where-Object { $_.HasMarker } | ForEach-Object { $_.Path })

$odPairs = @{
    'State'                  = $oneDriveState
    'Why'                    = $oneDriveWhy
    'Account subkeys'        = @($odAccounts.SubKeys).Count
    'ClientEverSignedIn'     = $(if ($null -eq $odEverSignedIn) { 'not set' } else { $odEverSignedIn })
    'Last account activity'  = $(if ($null -eq $odLastUpdateUtc) { 'no timestamp recorded' } else { Format-AgeText $odLastUpdateUtc })
    'OneDrive processes'     = $odProcessCount
    'Explorer sync roots'    = ('{0} registered ({1} for OneDrive)' -f $srmNames.Count, $srmOneDrive.Count)
    'Sync engine mounts'     = ('{0} registered under HKCU:\Software\SyncEngines\Providers' -f $syncEngineMounts.Count)
    'Sync folders on disk'   = $(if ($odRootsOnDisk.Count -eq 0) { 'none' } else { (Join-Readable @($odRootsOnDisk)) })
    'Sync-root marker file'  = $(if ($markerRoots.Count -eq 0) { 'not found at any discovered root' } else { ('present at ' + (Join-Readable @($markerRoots))) })
    'Confidence'             = 'High. Account subtree, process list and Explorer registration are all read directly; the meaning of individual per-account values is inferred.'
}
$odBody = New-KeyValueTable -Pairs $odPairs -Order @(
    'State', 'Why', 'Account subkeys', 'ClientEverSignedIn', 'Last account activity',
    'OneDrive processes', 'Explorer sync roots', 'Sync engine mounts', 'Sync folders on disk',
    'Sync-root marker file', 'Confidence')
if ($odAccountRows.Count -gt 0) {
    $odBody = $odBody + (New-DataTable -Rows @($odAccountRows | ForEach-Object {
        [pscustomobject]@{
            'Account'      = $_.Account
            'Email'        = $_.UserEmail
            'Sync folder'  = $_.UserFolder
            'KFM values'   = $_.KfmFlags
            'Values in key'= $_.ValueCount
        }
    }))
}
if ($syncEngineMounts.Count -gt 0) {
    $odBody = $odBody + (New-DataTable -Rows @($syncEngineMounts | ForEach-Object {
        [pscustomobject]@{
            'Engine'        = $_.Engine
            'Mount point'   = $_.MountPoint
            'Url namespace' = $_.UrlNamespace
            'Library type'  = $_.LibraryType
        }
    }))
}
$odNote = 'A OneDrive folder on disk, %OneDrive% in the environment and ClientEverSignedIn = 1 all survive signing out. None of them is evidence that anything is syncing.'
Add-Section 'OneDrive' $odBody $odNote

$fhInclusionText = 'none listed'
if (@($fh.Inclusions).Count -gt 0) { $fhInclusionText = (@($fh.Inclusions) -join '; ') }
$fhExclusionText = 'none listed'
if (@($fh.Exclusions).Count -gt 0) { $fhExclusionText = (@($fh.Exclusions) -join '; ') }

$fhTargetText = 'no target configured'
if ($fh.TargetUrl -or $fh.TargetName) {
    $fhTargetText = ('{0} {1} - {2}' -f $fh.TargetName, $fh.TargetUrl, $fh.TargetStatus).Trim()
}

$fhPairs = @{
    'Configured'      = $(if ($fh.Configured) { 'yes' } else { 'no configuration file found' })
    'Configuration'   = $(if ($fh.ConfigFile) { $fh.ConfigFile + ' (written ' + (Format-Utc $fh.ConfigUtc) + ')' } else { 'no Config*.xml under ' + $fhConfigDir })
    'Included folders'= $fhInclusionText
    'Excluded folders'= $fhExclusionText
    'Target'          = $fhTargetText
    'Catalog'         = $(if ($null -eq $fh.CatalogUtc) { 'no catalog database' } else { Format-AgeText $fh.CatalogUtc })
    'fhsvc service'   = ('{0}, start type {1}' -f $fh.ServiceStatus, $fh.ServiceStart)
    'Core/WHC channel'= ('{0} records' -f $fhCoreChannel.RecordCount)
    'Engine/BackupLog'= $(if ($fhEngineChannel.Denied) { 'skipped: needs admin' } else { ('{0} records' -f $fhEngineChannel.RecordCount) })
    'Last run evidence' = $(if ($null -eq $fhLastRunUtc) { 'nothing on this machine shows File History has ever run' } else { ('{0}, from {1}' -f (Format-AgeText $fhLastRunUtc), $fhEvidenceSource) })
    'Confidence'      = 'Medium. The absence of a configuration is certain; the configuration XML schema varies between Windows builds, so inclusion lists are read by shape rather than by fixed element names.'
}
if ($fh.ParseNote) { $fhPairs['Parse note'] = $fh.ParseNote }
Add-Section 'File History' (New-KeyValueTable -Pairs $fhPairs -Order @(
    'Configured', 'Configuration', 'Included folders', 'Excluded folders', 'Target', 'Catalog',
    'fhsvc service', 'Core/WHC channel', 'Engine/BackupLog', 'Last run evidence', 'Parse note', 'Confidence')) `
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\FileHistory exists on machines that have never used File History. Only a configuration file counts.'

$cloudPairs = @{
    'Dropbox'      = ('{0} - info.json {1}, registry key {2}, {3} process(es)' -f $dropboxState, $(if ($dropboxInfoFound) { 'found' } else { 'not found' }), $(if ($dropboxReg.Exists) { 'present' } else { 'absent' }), $dropboxProcCount)
    'Dropbox roots'= $(if ($dropboxAccounts.Count -eq 0) { 'none' } else { (@($dropboxAccounts | ForEach-Object { $_.Kind + ': ' + $_.Path + ' (from ' + $_.File + ')' }) -join '; ') })
    'Google Drive' = ('{0} - DriveFS data {1}, registry key {2}, {3} process(es), {4} mounted volume(s)' -f $gdriveState, $(if ($gdriveDataFound) { 'found' } else { 'not found' }), $(if ($gdriveReg.Exists) { 'present' } else { 'absent' }), $gdriveProcCount, $gdriveVolumes.Count)
    'Confidence'   = 'Low to medium. Absence is reliable. Presence detection relies on installer layouts that change between versions; a Google Drive volume is matched by its label, which the user can rename.'
}
if ($dropboxNote) { $cloudPairs['Dropbox note'] = $dropboxNote }
Add-Section 'Other cloud providers' (New-KeyValueTable -Pairs $cloudPairs -Order @('Dropbox', 'Dropbox roots', 'Dropbox note', 'Google Drive', 'Confidence'))

# ---- machine context
$srText = 'not initialised'
if ($srInitDone -eq 1) { $srText = 'initialised' }
if ($null -ne $srInterval) {
    if ($srInterval -eq 0) { $srText = $srText + ', restore point interval is 0 (automatic points are off)' }
    else                   { $srText = $srText + (', restore point interval {0}' -f $srInterval) }
}

$settingsText = 'no settings-backup events on this machine'
if ($settingsChannel.Exists -and $settingsChannel.RecordCount -gt 0) {
    if ($null -ne $settingsLastUtc) {
        $settingsText = ('Windows Backup logged "{0}" {1}, but that protects settings, apps and preferences - not your files.' -f `
            $settingsResult, (Format-AgeText $settingsLastUtc))
    }
    else {
        $settingsText = ('The settings-backup channel holds {0} records but none of them is a run result.' -f $settingsChannel.RecordCount)
    }
}

$srPointText = 'could not be read - see the skipped checks below'
if ($null -ne $srPointCount) {
    if     ($srPointCount -eq 0)     { $srPointText = 'none on this machine' }
    elseif ($null -eq $srNewestUtc)  { $srPointText = ('{0} restore point(s), none of them carrying a readable date' -f $srPointCount) }
    else                             { $srPointText = ('{0} restore point(s), newest {1}' -f $srPointCount, (Format-AgeText $srNewestUtc)) }
}

$shadowText = 'could not be read - see the skipped checks below'
if ($null -ne $shadowCount) {
    if     ($shadowCount -eq 0)         { $shadowText = 'none on this machine, so Previous Versions has nothing to offer' }
    elseif ($null -eq $shadowNewestUtc) { $shadowText = ('{0} shadow copy/copies, none of them carrying a readable date' -f $shadowCount) }
    else                                { $shadowText = ('{0} shadow copy/copies, newest {1}' -f $shadowCount, (Format-AgeText $shadowNewestUtc)) }
}

$contextPairs = @{
    'Windows Backup (settings)' = $settingsText
    'System Restore'            = $srText
    'Restore points'            = $srPointText
    'System Restore maintenance'= $(if ($null -eq $srLastMaintUtc) { 'no timestamp recorded' } else { Format-AgeText $srLastMaintUtc })
    'Shadow copies'             = $shadowText
    'Windows Backup channel'    = ('{0} records' -f $backupChannel.RecordCount)
    'Running as administrator'  = $(if ($IsElevated) { 'yes' } else { 'no' })
}
$contextBody = New-KeyValueTable -Pairs $contextPairs -Order @(
    'Windows Backup (settings)', 'System Restore', 'Restore points', 'System Restore maintenance',
    'Shadow copies', 'Windows Backup channel', 'Running as administrator')
if ($AdminSkips.Count -gt 0) {
    $contextBody = $contextBody + (New-DataTable -Rows @($AdminSkips))
}
Add-Section 'Machine context' $contextBody `
    ('Nothing in this section can raise or lower a folder''s light. System Restore keeps system state, not documents, and Windows'' settings backup keeps preferences, not files. ' +
     'Windows Server Backup is out of scope: wbadmin answers a query it refuses with exit code 0 and explains the refusal only in a translated sentence, so there is nothing there a script can read without guessing.')

# ---- history trend
if ($historyOk -and (Test-Path -LiteralPath $HistoryCsv)) {
    try {
        $history = @(Import-Csv -LiteralPath $HistoryCsv)
        $runs = @{}
        $order = New-Object System.Collections.Generic.List[string]
        foreach ($row in $history) {
            $ts = [string](Get-Prop $row 'TimestampUtc' '')
            if (-not $ts) { continue }
            if (-not $runs.ContainsKey($ts)) {
                $runs[$ts] = [pscustomobject]@{ 'Run (UTC)' = $ts; 'Green' = 0; 'Yellow' = 0; 'Red' = 0; 'Exit code' = [string](Get-Prop $row 'ExitCode' '') }
                $order.Add($ts)
            }
            switch ([string](Get-Prop $row 'Light' '')) {
                'GREEN'  { $runs[$ts].Green++ }
                'YELLOW' { $runs[$ts].Yellow++ }
                'RED'    { $runs[$ts].Red++ }
            }
        }
        $recent = @($order | Select-Object -Last 12 | ForEach-Object { $runs[$_] })
        if ($recent.Count -gt 0) {
            Add-Section 'History' (New-DataTable -Rows $recent) `
                ('One row per run, newest last, from {0}.' -f (Split-Path -Leaf $HistoryCsv))
        }
    }
    catch {
        Add-Section 'History' ('<p class="empty">The history file could not be read: ' + (ConvertTo-HtmlText $_.Exception.Message) + '</p>')
    }
}

# ---- methodology
$methodology = @(
    '<ul class="method">',
    '<li>Every verdict is worded as <strong>signs of</strong> or <strong>no signs of</strong> a backup. Nothing here certifies that a backup exists, is complete, or can be restored. Only a test restore proves that.</li>',
    '<li><strong>GREEN</strong>: this folder sits inside a sync root whose provider shows signs of running, or File History lists it and shows a run inside the freshness limit.</li>',
    '<li><strong>YELLOW</strong>: signs are present but stale, weak, pointing at something unreachable, or a check that would have settled it needed administrator rights.</li>',
    '<li><strong>RED</strong>: no signs of any backup at all, or the folder is inside a sync root that has been signed out - where the files on disk may be names without contents.</li>',
    '<li>Files are never opened. Only directory metadata is read, because reading a cloud-only placeholder asks the provider to download it.</li>',
    '<li>Nothing outside this tool''s own output folder is written, no service or process is started or stopped, and the registry is only read.</li>',
    '</ul>'
) -join ''
Add-Section 'How to read this' $methodology

# ---- attention panel
$attentionHtml = ''
if ($Attention.Count -eq 0) {
    $attentionHtml = '<div class="ok">Nothing needs attention. Every folder assessed shows signs of a current backup.</div>'
}
else {
    $rank = @{ 'Critical' = 0; 'Warning' = 1; 'Info' = 2 }
    $sorted = @($Attention | Sort-Object @{ Expression = { $rank[$_.Level] } }, @{ Expression = { $_.Index } })
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
.note { color:var(--muted); font-size:13px; margin:-6px 0 14px; }
.empty { color:var(--muted); font-style:italic; margin:0; }
table { border-collapse:collapse; width:100%; font-size:14px; }
.kv th {
  text-align:left; font-weight:600; color:var(--muted); padding:5px 16px 5px 0;
  white-space:nowrap; vertical-align:top; width:200px; font-size:13px;
}
.kv td { padding:5px 0; word-break:break-word; }
.kv + .data, .kv + .scroll { margin-top:16px; }
.scroll { overflow-x:auto; }
.data th {
  text-align:left; font-size:12px; text-transform:uppercase; letter-spacing:0.04em;
  color:var(--muted); border-bottom:1px solid var(--line); padding:8px 12px 8px 0; white-space:nowrap;
}
.data td { padding:8px 12px 8px 0; border-bottom:1px solid var(--line); vertical-align:top; }
.data tr:last-child td { border-bottom:none; }
.pill {
  display:inline-block; font-size:11px; font-weight:700; letter-spacing:0.06em;
  padding:3px 10px; border-radius:999px; background:var(--bar); color:var(--muted);
}
.pill.green { color:var(--ok); }
.pill.amber { color:var(--warn); }
.pill.red   { color:var(--crit); }
.verdict {
  background:var(--card); border:1px solid var(--line); border-radius:12px;
  padding:18px 22px; margin-bottom:18px; font-size:16px; line-height:1.5;
}
.verdict.good { border-left:4px solid var(--ok); }
.verdict.warn { border-left:4px solid var(--warn); }
.verdict.bad  { border-left:4px solid var(--crit); }
.verdict strong { display:block; margin-bottom:4px; }
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
ul.method { margin:0; padding-left:20px; }
ul.method li { margin-bottom:8px; }
footer { color:var(--muted); font-size:12px; text-align:center; margin-top:28px; }
'@

$html = New-Object System.Collections.Generic.List[string]
$html.Add('<!doctype html>')
$html.Add('<html lang="en"><head><meta charset="utf-8">')
$html.Add('<meta name="viewport" content="width=device-width,initial-scale=1">')
$html.Add(('<title>Backup reality - {0}</title>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add('<style>' + $css + '</style></head><body><div class="wrap">')
$html.Add('<h1>Is anything actually backed up?</h1>')
$html.Add(('<p class="sub">{0} &middot; checked {1} &middot; evidence older than {2} days counts as stale</p>' -f `
    (ConvertTo-HtmlText $env:COMPUTERNAME), (ConvertTo-HtmlText (Format-Utc $NowUtc)), $StaleDays))
$html.Add(('<div class="verdict {0}"><strong>{1}</strong></div>' -f $bannerClass, (ConvertTo-HtmlText $bannerText)))
$html.Add('<h2 style="font-size:15px;margin:0 0 12px;">Needs attention</h2>')
$html.Add($attentionHtml)
foreach ($s in $Sections) { $html.Add($s) }
$html.Add(('<footer>Generated by Test-BackupReality.ps1 {0} - read-only, nothing on this machine was changed.</footer>' -f (ConvertTo-HtmlText $ToolVersion)))
$html.Add('</div></body></html>')

($html -join [Environment]::NewLine) | Set-Content -LiteralPath $ReportFile -Encoding utf8

# ============================================== 7. console summary
if (-not $Quiet) {
    Write-Host ''
    Write-Host 'Backup reality'
    Write-Host '--------------'
    foreach ($a in $Assessments) {
        $colour = 'Red'
        if ($a.Light -eq 'GREEN')  { $colour = 'Green' }
        if ($a.Light -eq 'YELLOW') { $colour = 'Yellow' }
        Write-Host ('  {0,-24} ' -f $a.Target.Label) -NoNewline
        Write-Host ('{0,-6}' -f $a.Light) -ForegroundColor $colour -NoNewline
        Write-Host ('  {0}' -f $a.Mechanism)
    }

    Write-Host ''
    foreach ($line in $bannerLines) { Write-Host ('  ' + $line) }

    if ($reds.Count -gt 0) {
        Write-Host ''
        Write-Host '  What is wrong'
        Write-Host '  -------------'
        foreach ($a in $reds) {
            Write-Host ('    {0}: {1}' -f $a.Target.Label, $a.Detail) -ForegroundColor Red
        }
    }
    if ($FolderErrors.Count -gt 0) {
        Write-Host ''
        foreach ($e in $FolderErrors) {
            Write-Host ('  ! {0}: {1}' -f $e.Label, $e.Problem) -ForegroundColor Yellow
        }
    }
    if ($AdminSkips.Count -gt 0) {
        Write-Host ''
        Write-Host ('  {0} check(s) were asked for and refused, most of them for want of administrator rights; see the report.' -f $AdminSkips.Count) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host "Report  : $ReportFile"
    Write-Host "History : $HistoryCsv"
}

if ($Open) {
    # A machine with no handler registered for .html throws here, which would
    # otherwise lose the exit code the caller actually asked for.
    try { Start-Process -FilePath $ReportFile -ErrorAction Stop }
    catch { Write-Line ("  Could not open the report: {0}" -f $_.Exception.Message) 'Yellow' }
}

exit $ExitCode
