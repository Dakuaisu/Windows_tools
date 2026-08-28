<#
.SYNOPSIS
    Battery health in numbers: design capacity vs what the cells actually hold
    now, wear percentage, cycle count and live charge/discharge rate - plus
    Windows' own battery report HTML.

.DESCRIPTION
    Read-only. Wraps powercfg /batteryreport (both the HTML you can read and
    the XML you can parse), and cross-checks it against the live WMI battery
    classes, because the two disagree often enough to be worth seeing side by
    side.

    The number that matters is Health: full-charge capacity as a percentage of
    design capacity. Below ~80% is where a laptop starts feeling like it needs
    to stay plugged in; below ~60% the battery is effectively done.

    Every run appends a row to battery-history.csv, so after a few months you
    can see the wear curve rather than a single meaningless snapshot. That is
    the whole reason to run this on a schedule.

    Machines with two packs are read as a set: capacities are summed, health is
    computed from those totals, the cycle count is the highest of the packs, and
    each pack also gets its own line so one failing cell is still visible.

    Machines with no battery (desktops, VMs) are detected and exit cleanly.

.PARAMETER Days
    How many days of history powercfg should include in the HTML report.
    Default 14. Windows keeps at most about 60 days.

.PARAMETER OutputDir
    Where the reports and history CSV land.
    Defaults to $env:LOCALAPPDATA\BatteryReport. If that cannot be written to,
    the tool falls back to %TEMP%\BatteryReport and says so.

.PARAMETER HistoryLimit
    Rows kept in battery-history.csv. Default 5000; 0 keeps everything. When
    the column set changes the old file is archived rather than deleted.

.PARAMETER Open
    Open the generated HTML report when finished.

.PARAMETER Quiet
    Suppress console output; still writes files and sets the exit code.

.OUTPUTS
    Exit code 0 = battery healthy, 1 = worn (health below -WarnHealthPercent),
    2 = no battery present, 3 = could not read battery data.

.PARAMETER WarnHealthPercent
    Health below this is reported as worn. Default 80.

.EXAMPLE
    .\Get-BatteryReport.ps1 -Open

.EXAMPLE
    .\Get-BatteryReport.ps1 -Days 30

.EXAMPLE
    # Track wear over time
    Import-Csv "$env:LOCALAPPDATA\BatteryReport\battery-history.csv" |
        Select-Object Timestamp, HealthPercent, CycleCount
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 60)]
    [int]$Days = 14,

    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'BatteryReport'),
    [ValidateRange(1, 100)]
    [int]$WarnHealthPercent = 80,
    [ValidateRange(0, 1000000)]
    [int]$HistoryLimit = 5000,
    [switch]$Open,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

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

$RunStart = Get-Date
$Stamp    = $RunStart.ToString('yyyy-MM-dd_HHmmss')

$OutputDir = Initialize-OutputDir -Preferred $OutputDir -ToolName 'BatteryReport'

$HtmlReport = Join-Path $OutputDir "battery-report_$Stamp.html"
$XmlReport  = Join-Path $OutputDir "battery-report_$Stamp.xml"
$HistoryCsv = Join-Path $OutputDir 'battery-history.csv'

function Write-Line {
    param([string]$Message, [string]$Colour = '')
    if ($Quiet) { return }
    if ($Colour) { Write-Host $Message -ForegroundColor $Colour }
    else         { Write-Host $Message }
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Get-XmlValue {
    # Namespace-agnostic single-value read.
    param($Node, [string]$LocalName)
    if ($null -eq $Node) { return $null }
    $found = $Node.SelectSingleNode("*[local-name()='$LocalName']")
    if ($null -eq $found) { return $null }
    $t = [string]$found.InnerText
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    $t.Trim()
}

function Format-MilliwattHours {
    param($Mwh)
    if ($null -eq $Mwh -or $Mwh -le 0) { return 'n/a' }
    '{0:N0} mWh ({1:N2} Wh)' -f $Mwh, ($Mwh / 1000)
}

function ConvertTo-LongOrNull {
    <#
        powercfg's XML is machine-generated, but the numbers in it come from
        battery firmware: blanks, dashes and values too big for Int64 all turn
        up in the wild. A bare [long] cast on one of those throws and takes the
        whole XML parse down with it, discarding the good values next to the
        bad one.
    #>
    param($Text)
    if ($null -eq $Text) { return $null }
    $parsed = [long]0
    if ([long]::TryParse(([string]$Text).Trim(), [ref]$parsed)) { return $parsed }
    $null
}

function Add-HistoryRow {
    <#
        Export-Csv -Append refuses outright when the file's header does not
        match the object being appended, so the first run after this script
        gains or loses a column would throw away the whole reading at the last
        step. The stale file is archived under a dated name instead, which
        keeps the earlier readings rather than discarding them.

        Trimming happens here too: battery-history.csv had no bound at all.
    #>
    param([psobject]$Row, [string]$Path, [int]$MaxRows)

    $header = @($Row.PSObject.Properties | ForEach-Object { $_.Name })

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
            Write-Line ("  History columns changed - earlier readings kept as {0}" -f (Split-Path -Leaf $archive)) 'Yellow'
        }
    }

    $Row | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 -Append

    if ($MaxRows -gt 0) {
        $lines = @(Get-Content -LiteralPath $Path)
        if ($lines.Count -gt ($MaxRows + 1)) {
            $keep = @($lines[0]) + @($lines[($lines.Count - $MaxRows)..($lines.Count - 1)])
            $keep | Set-Content -LiteralPath $Path -Encoding utf8
        }
    }
}

Write-Line ''
Write-Line "=== Battery report : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) ==="
Write-Line ''

# ------------------------------------------------------- 1. is there a battery
$win32Batteries = @()
try {
    # -OperationTimeoutSec keeps a wedged WMI provider from hanging the run.
    $win32Batteries = @(Get-CimInstance -ClassName Win32_Battery -OperationTimeoutSec 20 -ErrorAction Stop)
}
catch { }

if ($win32Batteries.Count -eq 0) {
    Write-Line 'No battery detected - this looks like a desktop or a VM.' 'Yellow'
    Write-Line 'Nothing to report.'
    exit 2
}
$win32Battery = $win32Batteries[0]

# ------------------------------------------------------- 2. powercfg reports
$powercfgOk = $true

function Invoke-Powercfg {
    param([string[]]$Arguments)
    # $ErrorActionPreference = 'Stop' plus 2>&1 turns anything powercfg writes
    # to stderr into a terminating NativeCommandError, so the catch below used
    # to swallow the real message and report a bare RemoteException instead.
    # Relaxing the preference for the duration of the call keeps the actual
    # output, and $LASTEXITCODE stays the signal for whether it worked.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & powercfg.exe @Arguments 2>&1
        $text = ($out | Out-String)
        if ($LASTEXITCODE -ne 0) {
            $text = "powercfg exited with code $LASTEXITCODE`n$text"
        }
        return $text
    }
    catch {
        return "powercfg failed: $($_.Exception.Message)"
    }
    finally {
        $ErrorActionPreference = $previous
    }
}

Write-Line '  Generating powercfg battery report...'
$htmlOut = Invoke-Powercfg @('/batteryreport', '/output', $HtmlReport, '/duration', "$Days")
$xmlOut  = Invoke-Powercfg @('/batteryreport', '/output', $XmlReport,  '/duration', "$Days", '/xml')

if (-not (Test-Path -LiteralPath $HtmlReport)) {
    $powercfgOk = $false
    Write-Line "  ! powercfg produced no HTML report." 'Yellow'
    Write-Line ("    $($htmlOut.Trim())") 'DarkGray'
}

# ------------------------------------------------------------ 3. parse the XML
# One <Battery> element per pack. Machines with two packs - ThinkPads with a
# bridge battery, Surface Books, some gaming laptops - were previously read as
# whichever pack happened to come first, so their capacity and health were
# reported at roughly half the true figure.
$xmlBatteries = New-Object System.Collections.Generic.List[psobject]

if (Test-Path -LiteralPath $XmlReport) {
    try {
        $xml = New-Object System.Xml.XmlDocument
        $xml.Load($XmlReport)

        foreach ($batteryNode in @($xml.SelectNodes("//*[local-name()='Battery']"))) {
            $xmlBatteries.Add([pscustomobject]@{
                Id           = [string](Get-XmlValue $batteryNode 'Id')
                Design       = ConvertTo-LongOrNull (Get-XmlValue $batteryNode 'DesignCapacity')
                FullCharge   = ConvertTo-LongOrNull (Get-XmlValue $batteryNode 'FullChargeCapacity')
                Cycles       = ConvertTo-LongOrNull (Get-XmlValue $batteryNode 'CycleCount')
                Chemistry    = [string](Get-XmlValue $batteryNode 'Chemistry')
                Manufacturer = [string](Get-XmlValue $batteryNode 'Manufacturer')
                Serial       = [string](Get-XmlValue $batteryNode 'SerialNumber')
            })
        }
    }
    catch {
        Write-Line "  ! Could not parse the battery XML: $($_.Exception.Message)" 'Yellow'
    }
}

$xmlChemistry = ''
$xmlManuf     = ''
$xmlSerial    = ''
if ($xmlBatteries.Count -gt 0) {
    $xmlChemistry = [string]$xmlBatteries[0].Chemistry
    $xmlManuf     = [string]$xmlBatteries[0].Manufacturer
    $xmlSerial    = [string]$xmlBatteries[0].Serial
}

# --------------------------------------------------------- 4. live WMI values
# root\wmi is where the firmware numbers live. Any of these can be absent
# depending on how honest the OEM's ACPI implementation is.
$wmiCharging = $null
$wmiRate     = $null
$wmiVoltage  = $null
$wmiRemaining= $null

function Get-WmiBatteryValues {
    # All instances, not just the first: one per pack, same as the XML.
    param([string]$Class, [string]$Property)
    try {
        return @(
            Get-CimInstance -Namespace 'root\wmi' -ClassName $Class -OperationTimeoutSec 20 -ErrorAction Stop |
                ForEach-Object { Get-Prop $_ $Property } |
                Where-Object { $null -ne $_ }
        )
    }
    catch { return @() }
}

$wmiDesigns = @(Get-WmiBatteryValues 'BatteryStaticData'          'DesignedCapacity')
$wmiFulls   = @(Get-WmiBatteryValues 'BatteryFullChargedCapacity' 'FullChargedCapacity')
$wmiCycleValues = @(Get-WmiBatteryValues 'BatteryCycleCount'      'CycleCount')

try {
    $status = Get-CimInstance -Namespace 'root\wmi' -ClassName 'BatteryStatus' -OperationTimeoutSec 20 -ErrorAction Stop |
        Select-Object -First 1
    if ($status) {
        $wmiCharging  = [bool](Get-Prop $status 'Charging' $false)
        $wmiVoltage   = Get-Prop $status 'Voltage'
        $wmiRemaining = Get-Prop $status 'RemainingCapacity'
        $charge    = Get-Prop $status 'ChargeRate' 0
        $discharge = Get-Prop $status 'DischargeRate' 0
        if ($charge)    { $wmiRate = [int]$charge }
        if ($discharge) { $wmiRate = -[int]$discharge }
    }
}
catch { }

# ------------------------------------------------------------ 5. reconcile
# Prefer the powercfg XML - it is the same source the OS trusts - and fall
# back to the live WMI classes when powercfg came up empty. Capacities are
# summed across packs, because health is a property of the whole pack set.
function Measure-Sum {
    param($Values)
    $list = @($Values | Where-Object { $null -ne $_ -and $_ -gt 0 })
    if ($list.Count -eq 0) { return $null }
    $total = [long]0
    foreach ($v in $list) { $total += [long]$v }
    $total
}

$design     = Measure-Sum @($xmlBatteries | ForEach-Object { $_.Design })
$fullCharge = Measure-Sum @($xmlBatteries | ForEach-Object { $_.FullCharge })
$cycleList  = @($xmlBatteries | ForEach-Object { $_.Cycles } | Where-Object { $null -ne $_ -and $_ -gt 0 })
$source     = 'powercfg XML'

if ($null -eq $design)     { $design     = Measure-Sum $wmiDesigns; $source = 'WMI (root\wmi)' }
if ($null -eq $fullCharge) { $fullCharge = Measure-Sum $wmiFulls;   $source = 'WMI (root\wmi)' }
if ($cycleList.Count -eq 0) {
    $cycleList = @($wmiCycleValues | Where-Object { $null -ne $_ -and $_ -gt 0 })
}

# Cycles are per pack and do not add up to anything meaningful; the most worn
# pack is the one that matters.
$cycles = $null
if ($cycleList.Count -gt 0) {
    $cycles = [long](($cycleList | Measure-Object -Maximum).Maximum)
}

$batteryCount = $xmlBatteries.Count
if ($batteryCount -eq 0) { $batteryCount = $win32Batteries.Count }

$health = $null
$wear   = $null
if ($design -and $fullCharge -and $design -gt 0) {
    $health = [math]::Round(100.0 * $fullCharge / $design, 1)
    $wear   = [math]::Round(100.0 - $health, 1)
}

$chargePercent = Get-Prop $win32Battery 'EstimatedChargeRemaining'
$runtimeMin    = Get-Prop $win32Battery 'EstimatedRunTime'
# 71582788 is the "unknown / on AC" sentinel Win32_Battery returns. Some
# firmware returns other implausible values for the same thing, so anything
# past a week is treated as "not reported" rather than printed as fact.
if ($null -ne $runtimeMin -and $runtimeMin -ge 10080) { $runtimeMin = $null }

$chemistry = $xmlChemistry
if (-not $chemistry) { $chemistry = [string](Get-Prop $win32Battery 'Chemistry' '') }

# ---------------------------------------------------------------- 6. output
$verdict = 'Unknown'
$colour  = 'Yellow'
if ($null -ne $health) {
    if     ($health -ge 90)                  { $verdict = 'Excellent'; $colour = 'Green' }
    elseif ($health -ge $WarnHealthPercent)  { $verdict = 'Good';      $colour = 'Green' }
    elseif ($health -ge 60)                  { $verdict = 'Worn';      $colour = 'Yellow' }
    else                                     { $verdict = 'Failing';   $colour = 'Red' }
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('Battery health')
$lines.Add('--------------')
$lines.Add(('  Device            : {0}' -f [string](Get-Prop $win32Battery 'Name' 'unknown')))
if ($xmlManuf)  { $lines.Add(('  Manufacturer      : {0}' -f $xmlManuf)) }
if ($xmlSerial) { $lines.Add(('  Serial            : {0}' -f $xmlSerial)) }
if ($chemistry) { $lines.Add(('  Chemistry         : {0}' -f $chemistry)) }
$lines.Add(('  Data source       : {0}' -f $source))
if ($batteryCount -gt 1) {
    $lines.Add(('  Battery packs     : {0} (capacities below are the total)' -f $batteryCount))
}
$lines.Add('')
$lines.Add(('  Design capacity   : {0}' -f (Format-MilliwattHours $design)))
$lines.Add(('  Full charge now   : {0}' -f (Format-MilliwattHours $fullCharge)))

if ($null -ne $health) {
    $barWidth = 40
    $filled   = [int][math]::Round($barWidth * [math]::Min($health, 100) / 100)
    $bar      = ('#' * $filled) + ('.' * ($barWidth - $filled))
    $lines.Add(('  Health            : {0}%  [{1}]' -f $health, $bar))
    $lines.Add(('  Wear              : {0}%' -f $wear))
}
else {
    $lines.Add('  Health            : could not be determined')
}

if ($null -ne $cycles -and $cycles -gt 0) {
    $label = '  Cycle count       : {0}'
    if ($batteryCount -gt 1) { $label = '  Cycle count       : {0} (highest of {1} packs)' }
    $lines.Add(($label -f $cycles, $batteryCount))
}
else {
    $lines.Add('  Cycle count       : not reported by this firmware')
}

# With more than one pack the totals above hide a single failing cell, so each
# pack gets its own line as well.
if ($batteryCount -gt 1 -and $xmlBatteries.Count -gt 1) {
    $lines.Add('')
    $lines.Add('  Per pack')
    foreach ($b in $xmlBatteries) {
        $packHealth = 'health unknown'
        if ($b.Design -and $b.FullCharge -and $b.Design -gt 0) {
            $packHealth = '{0}% health' -f [math]::Round(100.0 * $b.FullCharge / $b.Design, 1)
        }
        $packName = $b.Id
        if (-not $packName) { $packName = '(unnamed)' }
        $lines.Add(('    {0,-24} {1}  ->  {2}' -f $packName, (Format-MilliwattHours $b.FullCharge), $packHealth))
    }
}

$lines.Add('')
$lines.Add('Right now')
$lines.Add('---------')
if ($null -ne $chargePercent) { $lines.Add(('  Charge            : {0}%' -f $chargePercent)) }
$state = ''
if ($wmiCharging) {
    $state = 'charging'
}
else {
    # Win32_Battery reports 2 for "AC connected", which is how a laptop sitting
    # at 100% on the charger looks: not charging, but not on battery either.
    $acStatus = Get-Prop $win32Battery 'BatteryStatus' 0
    if ($acStatus -eq 2) { $state = 'on AC, not charging' }
    else                 { $state = 'on battery' }
}
$lines.Add(('  State             : {0}' -f $state))
if ($null -ne $wmiRate -and $wmiRate -ne 0) {
    $lines.Add(('  Power flow        : {0:N0} mW' -f $wmiRate))
}
if ($null -ne $wmiVoltage -and $wmiVoltage -gt 0) {
    $lines.Add(('  Voltage           : {0:N2} V' -f ($wmiVoltage / 1000)))
}
if ($null -ne $runtimeMin) {
    $lines.Add(('  Estimated runtime : {0} min' -f $runtimeMin))
}

$lines.Add('')
$lines.Add(('  VERDICT           : {0}' -f $verdict))

$text = $lines -join [Environment]::NewLine
Write-Line ''
if (-not $Quiet) {
    foreach ($l in $lines) {
        if ($l -like '*VERDICT*') { Write-Host $l -ForegroundColor $colour }
        else                      { Write-Host $l }
    }
}

# ------------------------------------------------------------- 7. history row
$historyRow = [pscustomobject]@{
    Timestamp        = $RunStart.ToString('yyyy-MM-dd HH:mm:ss')
    DesignCapacityMwh= $design
    FullChargeMwh    = $fullCharge
    HealthPercent    = $health
    WearPercent      = $wear
    CycleCount       = $cycles
    ChargePercent    = $chargePercent
    Charging         = $wmiCharging
    Verdict          = $verdict
    Source           = $source
}
$historyOk = $true
try { Add-HistoryRow -Row $historyRow -Path $HistoryCsv -MaxRows $HistoryLimit }
catch {
    $historyOk = $false
    Write-Line "  ! Could not update $HistoryCsv : $($_.Exception.Message)" 'Yellow'
}

# Show the trend if we have enough history to be interesting. A history file
# written by an older version of this script can be missing HealthPercent
# entirely, and strict mode turns that from $null into a thrown error - so the
# column is looked up rather than dereferenced.
if (-not $Quiet -and $historyOk -and (Test-Path -LiteralPath $HistoryCsv)) {
    try {
        $history = @(Import-Csv -LiteralPath $HistoryCsv |
            Where-Object { (Get-Prop $_ 'HealthPercent') -and (Get-Prop $_ 'Timestamp') })
        if ($history.Count -ge 2) {
            $first     = $history[0]
            $last      = $history[-1]
            $firstPct  = 0.0
            $lastPct   = 0.0
            if ([double]::TryParse([string](Get-Prop $first 'HealthPercent'), [ref]$firstPct) -and
                [double]::TryParse([string](Get-Prop $last  'HealthPercent'), [ref]$lastPct)) {
                $delta = [math]::Round($lastPct - $firstPct, 1)
                $sign  = ''
                if ($delta -ge 0) { $sign = '+' }
                Write-Host ''
                Write-Host 'Trend'
                Write-Host '-----'
                Write-Host ("  {0} readings since {1}" -f $history.Count, (Get-Prop $first 'Timestamp'))
                Write-Host ("  health {0}% -> {1}%  ({2}{3} points)" -f $firstPct, $lastPct, $sign, $delta)
            }
        }
    }
    catch {
        Write-Line "  ! Could not read the history trend: $($_.Exception.Message)" 'Yellow'
    }
}

# Keep the 12 most recent report pairs.
Get-ChildItem -LiteralPath $OutputDir -Filter 'battery-report_*' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 24 |
    Remove-Item -Force -ErrorAction SilentlyContinue

Write-Line ''
if ($powercfgOk) { Write-Line "HTML report : $HtmlReport" }
Write-Line "History     : $HistoryCsv"

if ($Open -and (Test-Path -LiteralPath $HtmlReport)) {
    # A machine with no handler registered for .html throws here, which would
    # otherwise lose the exit code the caller actually asked for.
    try { Start-Process -FilePath $HtmlReport -ErrorAction Stop }
    catch { Write-Line "  ! Could not open the report: $($_.Exception.Message)" 'Yellow' }
}

if ($null -eq $health)                    { exit 3 }
if ($health -lt $WarnHealthPercent)       { exit 1 }
exit 0
