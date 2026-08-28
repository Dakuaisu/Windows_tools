<#
.SYNOPSIS
    One self-contained HTML page describing this machine right now: hardware,
    disks and SMART health, memory, network, battery, pending reboots, driver
    problems, Defender state and the biggest memory consumers.

.DESCRIPTION
    Read-only. Everything comes from CIM/WMI, the registry and a few cmdlets;
    nothing is changed. The output is a single HTML file with no external
    assets, so you can mail it, attach it to a support ticket, or keep one per
    month and diff them when the machine starts behaving oddly.

    A "Needs attention" panel is built first and pinned to the top: low disk
    space, unhealthy drives, pending reboots, devices in an error state, stale
    Windows Update, disabled real-time protection. If that panel is empty the
    machine is, by these checks, fine.

    Battery details are delegated to Get-BatteryReport.ps1 when it sits in the
    same folder; otherwise a short inline summary is used.

.PARAMETER OutputDir
    Where the snapshot lands. Defaults to $env:LOCALAPPDATA\SystemSnapshot. If
    that cannot be written to, the tool falls back to %TEMP%\SystemSnapshot and
    says so.

.PARAMETER ProbeTimeoutSeconds
    How long to wait on the two checks that can hang on a broken machine -
    Defender status and device enumeration - before giving up on that check and
    carrying on with the rest of the report. Default 30.

.PARAMETER TopProcesses
    How many processes to list, by working set. Default 12. Zero to skip.

.PARAMETER LowDiskPercent
    Free space below this percentage raises an attention item. Default 10.

.PARAMETER Open
    Open the report when finished.

.PARAMETER Quiet
    Suppress console output; still writes the file and sets the exit code.

.OUTPUTS
    Exit code 0 = nothing needs attention, 1 = at least one attention item.

.EXAMPLE
    .\Get-SystemSnapshot.ps1 -Open

.EXAMPLE
    .\Get-SystemSnapshot.ps1 -Quiet -TopProcesses 0
#>
[CmdletBinding()]
param(
    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir      = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'SystemSnapshot'),
    [ValidateRange(0, 500)]
    [int]   $TopProcesses   = 12,
    [ValidateRange(0, 99)]
    [int]   $LowDiskPercent = 10,
    [ValidateRange(5, 600)]
    [int]   $ProbeTimeoutSeconds = 30,
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

$OutputDir  = Initialize-OutputDir -Preferred $OutputDir -ToolName 'SystemSnapshot'
$ReportFile = Join-Path $OutputDir "snapshot_$Stamp.html"

$Attention = New-Object System.Collections.Generic.List[psobject]
$Sections  = New-Object System.Collections.Generic.List[string]

function Write-Step {
    param([string]$Message)
    if (-not $Quiet) { Write-Host "  $Message" }
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Get-Cim {
    # Every CIM class here is optional on some SKU or other; never let one
    # missing class take down the whole report. -OperationTimeoutSec bounds a
    # wedged WMI provider, which is otherwise an unbounded wait.
    param([string]$Class, [string]$Namespace = 'root\cimv2')
    try { return @(Get-CimInstance -Namespace $Namespace -ClassName $Class -OperationTimeoutSec 30 -ErrorAction Stop) }
    catch { return @() }
}

function Invoke-WithTimeout {
    <#
        Runs a scriptblock in its own runspace and gives up on it after
        -Seconds.

        Get-MpComputerStatus and Get-PnpDevice both talk to services that can be
        wedged - a half-broken Defender install, a device enumeration stuck on a
        failing USB controller - and neither cmdlet takes a timeout. A
        diagnostic that hangs forever is worse than one that says it could not
        read something.

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

function Get-SafeValue {
    <#
        Reading .CPU, .Threads or .StartTime on a process that has just exited -
        or one this session is not allowed to open - throws, and inside a
        pipeline with $ErrorActionPreference = 'Stop' that takes down the whole
        section rather than the one row.
    #>
    param([scriptblock]$Get, $Default = '')
    try {
        $v = & $Get
        if ($null -eq $v) { return $Default }
        return $v
    }
    catch { return $Default }
}

function Format-Size {
    param([double]$Bytes)
    if     ($Bytes -ge 1TB) { '{0:N2} TB' -f ($Bytes / 1TB) }
    elseif ($Bytes -ge 1GB) { '{0:N1} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N0} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -gt 0)   { '{0:N0} KB' -f ($Bytes / 1KB) }
    else                    { '-' }
}

function ConvertTo-HtmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Add-Attention {
    param(
        [ValidateSet('Critical','Warning','Info')][string]$Level,
        [string]$Title,
        [string]$Detail
    )
    $Attention.Add([pscustomobject]@{ Level = $Level; Title = $Title; Detail = $Detail })
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

function New-RawHtml {
    <#
        Marks a cell whose value is already HTML and must not be escaped again.

        This used to be a 'RAW:' string prefix, which is a footgun: any genuine
        value that happened to start with those four characters - a volume
        label, a device name, a process name - would have been written into the
        page unescaped. A typed marker cannot be produced by accident.
    #>
    param([string]$Html)
    [pscustomobject]@{ PSTypeName = 'Snapshot.RawHtml'; Html = $Html }
}

function New-DataTable {
    <#
        Rows in, HTML table out. Columns are the property names of the first
        row unless -Columns is given. A value built by New-RawHtml is emitted
        unescaped, which is how the disk-usage bars get in; everything else is
        HTML-encoded.
    #>
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

function Add-Section {
    param([string]$Title, [string]$Body, [string]$Note = '')
    $noteHtml = ''
    if ($Note) { $noteHtml = '<p class="note">' + (ConvertTo-HtmlText $Note) + '</p>' }
    $Sections.Add('<section><h2>' + (ConvertTo-HtmlText $Title) + '</h2>' + $noteHtml + $Body + '</section>')
}

function New-Bar {
    param([double]$Percent, [string]$Class = '')
    $p = [math]::Max(0, [math]::Min(100, $Percent))
    New-RawHtml ('<div class="bar ' + $Class + '"><span style="width:' + ([math]::Round($p, 1)) + '%"></span></div>')
}

if (-not $Quiet) {
    Write-Host ''
    Write-Host "=== System snapshot : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) ==="
    Write-Host ''
}

# ------------------------------------------------------------- 1. the machine
Write-Step 'Reading machine identity...'

$os      = @(Get-Cim 'Win32_OperatingSystem')      | Select-Object -First 1
$cs      = @(Get-Cim 'Win32_ComputerSystem')       | Select-Object -First 1
$bios    = @(Get-Cim 'Win32_BIOS')                 | Select-Object -First 1
$cpus    = @(Get-Cim 'Win32_Processor')
$baseb   = @(Get-Cim 'Win32_BaseBoard')            | Select-Object -First 1

$uptime      = ''
$lastBoot    = Get-Prop $os 'LastBootUpTime'
$uptimeSpan  = $null
if ($lastBoot) {
    $uptimeSpan = $RunStart - [datetime]$lastBoot
    $uptime = '{0}d {1}h {2}m' -f $uptimeSpan.Days, $uptimeSpan.Hours, $uptimeSpan.Minutes
}

$displayVersion = ''
try {
    $cv = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    $displayVersion = [string](Get-Prop $cv 'DisplayVersion' (Get-Prop $cv 'ReleaseId' ''))
    $ubr = Get-Prop $cv 'UBR'
    if ($ubr) { $displayVersion = ('{0} (build {1}.{2})' -f $displayVersion, (Get-Prop $cv 'CurrentBuild' ''), $ubr) }
}
catch { }

$cpuName = ''
$cpuInfo = ''
if ($cpus.Count -gt 0) {
    $cpuName = [string](Get-Prop $cpus[0] 'Name' '')
    $cpuInfo = '{0} core(s), {1} logical' -f `
        (Get-Prop $cpus[0] 'NumberOfCores' '?'), (Get-Prop $cpus[0] 'NumberOfLogicalProcessors' '?')
}

$ramTotal = Get-Prop $cs 'TotalPhysicalMemory' 0
$ramFree  = (Get-Prop $os 'FreePhysicalMemory' 0) * 1KB
$ramUsedPct = 0
if ($ramTotal -gt 0) { $ramUsedPct = [math]::Round(100.0 * ($ramTotal - $ramFree) / $ramTotal, 1) }

$machinePairs = @{
    'Computer name'   = $env:COMPUTERNAME
    'User'            = "$env:USERDOMAIN\$env:USERNAME"
    'Manufacturer'    = Get-Prop $cs 'Manufacturer' ''
    'Model'           = Get-Prop $cs 'Model' ''
    'Motherboard'     = ('{0} {1}' -f (Get-Prop $baseb 'Manufacturer' ''), (Get-Prop $baseb 'Product' '')).Trim()
    'Serial'          = Get-Prop $bios 'SerialNumber' ''
    'BIOS'            = ('{0} {1}' -f (Get-Prop $bios 'SMBIOSBIOSVersion' ''), (Get-Prop $bios 'ReleaseDate' '')).Trim()
    'OS'              = Get-Prop $os 'Caption' ''
    'Version'         = $displayVersion
    'Architecture'    = Get-Prop $os 'OSArchitecture' ''
    'Installed'       = Get-Prop $os 'InstallDate' ''
    'Last boot'       = $lastBoot
    'Uptime'          = $uptime
    'CPU'             = $cpuName
    'CPU topology'    = $cpuInfo
    'Memory'          = ('{0} total, {1} free ({2}% in use)' -f (Format-Size $ramTotal), (Format-Size $ramFree), $ramUsedPct)
    'PowerShell'      = $PSVersionTable.PSVersion.ToString()
}

Add-Section 'Machine' (New-KeyValueTable -Pairs $machinePairs -Order @(
    'Computer name','User','Manufacturer','Model','Motherboard','Serial','BIOS',
    'OS','Version','Architecture','Installed','Last boot','Uptime',
    'CPU','CPU topology','Memory','PowerShell'))

if ($uptimeSpan -and $uptimeSpan.TotalDays -gt 14) {
    Add-Attention 'Info' 'Long uptime' ("This machine has been up for {0:N0} days. Updates and drivers often need a reboot to finish applying." -f $uptimeSpan.TotalDays)
}
if ($ramUsedPct -ge 90) {
    Add-Attention 'Warning' 'Memory pressure' "$ramUsedPct% of physical memory is in use."
}

# --------------------------------------------------------------- 2. memory
Write-Step 'Reading memory modules...'
$dimms = @(
    Get-Cim 'Win32_PhysicalMemory' | ForEach-Object {
        [pscustomobject]@{
            Slot         = [string](Get-Prop $_ 'DeviceLocator' '')
            Size         = Format-Size (Get-Prop $_ 'Capacity' 0)
            'Speed MT/s' = [string](Get-Prop $_ 'ConfiguredClockSpeed' (Get-Prop $_ 'Speed' ''))
            Manufacturer = [string](Get-Prop $_ 'Manufacturer' '')
            Part         = [string](Get-Prop $_ 'PartNumber' '').Trim()
        }
    }
)
if ($dimms.Count -gt 0) {
    Add-Section 'Memory modules' (New-DataTable -Rows $dimms)
}

# ---------------------------------------------------------------- 3. storage
Write-Step 'Reading disks and volumes...'

$physical = @()
try {
    # Built first, then judged. Raising attention items inside the pipeline
    # meant a Get-PhysicalDisk that failed halfway left those items behind and
    # the catch below then re-enumerated the same disks from scratch.
    $physical = @(
        Get-PhysicalDisk -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{
                Disk       = [string](Get-Prop $_ 'FriendlyName' '')
                Media      = [string](Get-Prop $_ 'MediaType' '')
                Bus        = [string](Get-Prop $_ 'BusType' '')
                Size       = Format-Size (Get-Prop $_ 'Size' 0)
                Health     = [string](Get-Prop $_ 'HealthStatus' 'Unknown')
                Usage      = [string](Get-Prop $_ 'Usage' '')
                Serial     = [string](Get-Prop $_ 'SerialNumber' '').Trim()
            }
        }
    )
    foreach ($d in $physical) {
        if ($d.Health -and $d.Health -ne 'Healthy') {
            $name = $d.Disk
            if (-not $name) { $name = 'disk' }
            Add-Attention 'Critical' 'Drive health' "$name reports health status '$($d.Health)'. Back up now and check SMART data."
        }
    }
}
catch {
    # Storage cmdlets are absent on some editions; fall back to the old class.
    $physical = @(
        Get-Cim 'Win32_DiskDrive' | ForEach-Object {
            [pscustomobject]@{
                Disk   = [string](Get-Prop $_ 'Model' '')
                Media  = [string](Get-Prop $_ 'MediaType' '')
                Bus    = [string](Get-Prop $_ 'InterfaceType' '')
                Size   = Format-Size (Get-Prop $_ 'Size' 0)
                Health = [string](Get-Prop $_ 'Status' '')
                Usage  = ''
                Serial = [string](Get-Prop $_ 'SerialNumber' '').Trim()
            }
        }
    )
}
if ($physical.Count -gt 0) {
    Add-Section 'Physical disks' (New-DataTable -Rows $physical)
}

$volumes = @(
    Get-Cim 'Win32_LogicalDisk' |
        Where-Object { (Get-Prop $_ 'DriveType' 0) -eq 3 } |
        ForEach-Object {
            $size = [double](Get-Prop $_ 'Size' 0)
            $free = [double](Get-Prop $_ 'FreeSpace' 0)
            $freePct = 0.0
            if ($size -gt 0) { $freePct = [math]::Round(100.0 * $free / $size, 1) }

            $barClass = ''
            if ($freePct -lt $LowDiskPercent) { $barClass = 'danger' }
            elseif ($freePct -lt 20)          { $barClass = 'warn' }

            if ($size -gt 0 -and $freePct -lt $LowDiskPercent) {
                Add-Attention 'Warning' 'Low disk space' ("$([string](Get-Prop $_ 'DeviceID' '')) has only $freePct% free ($(Format-Size $free)).")
            }

            [pscustomobject]@{
                Drive     = [string](Get-Prop $_ 'DeviceID' '')
                Label     = [string](Get-Prop $_ 'VolumeName' '')
                FileSystem= [string](Get-Prop $_ 'FileSystem' '')
                Size      = Format-Size $size
                Free      = Format-Size $free
                'Free %'  = "$freePct%"
                Used      = New-Bar -Percent (100 - $freePct) -Class $barClass
            }
        }
)
if ($volumes.Count -gt 0) {
    Add-Section 'Volumes' (New-DataTable -Rows $volumes)
}

# ---------------------------------------------------------------- 4. graphics
Write-Step 'Reading graphics adapters...'
$gpus = @(
    Get-Cim 'Win32_VideoController' | ForEach-Object {
        [pscustomobject]@{
            Adapter    = [string](Get-Prop $_ 'Name' '')
            Driver     = [string](Get-Prop $_ 'DriverVersion' '')
            'Driver date' = [string](Get-Prop $_ 'DriverDate' '')
            Resolution = ('{0} x {1}' -f (Get-Prop $_ 'CurrentHorizontalResolution' '?'), (Get-Prop $_ 'CurrentVerticalResolution' '?'))
            VRAM       = Format-Size ([double](Get-Prop $_ 'AdapterRAM' 0))
        }
    }
)
if ($gpus.Count -gt 0) {
    Add-Section 'Graphics' (New-DataTable -Rows $gpus) 'VRAM is what the driver reports through WMI and is capped at 4 GB for older adapters - treat a suspiciously round 4 GB as "unknown".'
}

# ---------------------------------------------------------------- 5. network
Write-Step 'Reading network configuration...'
$net = @()
try {
    $net = @(
        Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
            $adapter = $_
            $ip = ''
            $gw = ''
            $dns = ''
            try {
                $cfg = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -ErrorAction Stop
                $ip  = (@(Get-Prop $cfg 'IPv4Address' @()) | ForEach-Object { $_.IPAddress }) -join ', '
                $gw  = (@(Get-Prop $cfg 'IPv4DefaultGateway' @()) | ForEach-Object { $_.NextHop }) -join ', '
                $dnsObj = Get-Prop $cfg 'DNSServer'
                if ($dnsObj) { $dns = (@($dnsObj | ForEach-Object { $_.ServerAddresses }) -join ', ') }
            }
            catch { }

            [pscustomobject]@{
                Adapter = [string]$adapter.Name
                Type    = [string](Get-Prop $adapter 'MediaType' '')
                Speed   = [string](Get-Prop $adapter 'LinkSpeed' '')
                MAC     = [string](Get-Prop $adapter 'MacAddress' '')
                IPv4    = $ip
                Gateway = $gw
                DNS     = $dns
            }
        }
    )
}
catch {
    $net = @(
        Get-Cim 'Win32_NetworkAdapterConfiguration' | Where-Object { (Get-Prop $_ 'IPEnabled' $false) } | ForEach-Object {
            [pscustomobject]@{
                Adapter = [string](Get-Prop $_ 'Description' '')
                Type    = ''
                Speed   = ''
                MAC     = [string](Get-Prop $_ 'MACAddress' '')
                IPv4    = (@(Get-Prop $_ 'IPAddress' @()) -join ', ')
                Gateway = (@(Get-Prop $_ 'DefaultIPGateway' @()) -join ', ')
                DNS     = (@(Get-Prop $_ 'DNSServerSearchOrder' @()) -join ', ')
            }
        }
    )
}
if ($net.Count -gt 0) {
    Add-Section 'Network' (New-DataTable -Rows $net)
}

# ---------------------------------------------------------------- 6. battery
Write-Step 'Reading battery...'
$batteries = @(Get-Cim 'Win32_Battery')
$battery   = $batteries | Select-Object -First 1
if ($battery) {
    # Summed across packs: a machine with two batteries was previously read as
    # whichever pack came first, halving both capacities and the health figure.
    function Measure-BatteryTotal {
        param([string]$Class, [string]$Property)
        try {
            $values = @(
                Get-CimInstance -Namespace 'root\wmi' -ClassName $Class -OperationTimeoutSec 20 -ErrorAction Stop |
                    ForEach-Object { Get-Prop $_ $Property } |
                    Where-Object { $null -ne $_ -and $_ -gt 0 }
            )
            if ($values.Count -eq 0) { return $null }
            $total = [long]0
            foreach ($v in $values) { $total += [long]$v }
            return $total
        }
        catch { return $null }
    }

    $design = Measure-BatteryTotal 'BatteryStaticData'          'DesignedCapacity'
    $full   = Measure-BatteryTotal 'BatteryFullChargedCapacity' 'FullChargedCapacity'

    $health = $null
    if ($design -and $full -and $design -gt 0) {
        $health = [math]::Round(100.0 * $full / $design, 1)
        if ($health -lt 70) {
            Add-Attention 'Warning' 'Battery worn' "Full-charge capacity is $health% of design capacity."
        }
    }

    $charge = Get-Prop $battery 'EstimatedChargeRemaining' 0
    $healthText = 'unknown'
    if ($null -ne $health) { $healthText = "$health%" }

    $deviceName = [string](Get-Prop $battery 'Name' '')
    if ($batteries.Count -gt 1) {
        $deviceName = '{0} (+{1} more pack(s); figures below are the total)' -f $deviceName, ($batteries.Count - 1)
    }

    $batteryPairs = @{
        'Device'         = $deviceName
        'Charge'         = "$charge%"
        'Design capacity'= $(if ($design) { '{0:N0} mWh' -f $design } else { 'not reported' })
        'Full charge now'= $(if ($full)   { '{0:N0} mWh' -f $full }   else { 'not reported' })
        'Health'         = $healthText
    }
    $body = New-KeyValueTable -Pairs $batteryPairs -Order @('Device','Charge','Design capacity','Full charge now','Health')

    $sibling = Join-Path $PSScriptRoot 'Get-BatteryReport.ps1'
    $note = ''
    if (Test-Path -LiteralPath $sibling) {
        $note = 'Run Get-BatteryReport.ps1 in this folder for cycle count, wear trend and the full powercfg report.'
    }
    Add-Section 'Battery' $body $note
}

# ------------------------------------------------------- 7. health checks
Write-Step 'Running health checks...'

# --- pending reboot ---
$rebootReasons = @()
$rebootKeys = @(
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason = 'Component Based Servicing' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason = 'Windows Update' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'; Reason = 'Pending package install' }
)
foreach ($rk in $rebootKeys) {
    if (Test-Path -LiteralPath $rk.Path) { $rebootReasons += $rk.Reason }
}
try {
    $sm = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
    $pending = Get-Prop $sm 'PendingFileRenameOperations'
    if ($pending -and @($pending).Count -gt 0) { $rebootReasons += 'Pending file rename operations' }
}
catch { }

if ($rebootReasons.Count -gt 0) {
    Add-Attention 'Warning' 'Reboot pending' ('Reason(s): ' + ($rebootReasons -join ', '))
}

# --- problem devices ---
$badDevices = @()
$pnp = Invoke-WithTimeout -Seconds $ProbeTimeoutSeconds -Script {
    try {
        Get-PnpDevice -ErrorAction Stop |
            Where-Object { $_.Status -eq 'Error' -or $_.Status -eq 'Degraded' } |
            Select-Object FriendlyName, Name, Class, Status, InstanceId
    }
    catch { }
}
if (-not $pnp.Completed) {
    Write-Step '! device enumeration timed out - skipping that check'
}
else {
    $badDevices = @(
        $pnp.Output | ForEach-Object {
            [pscustomobject]@{
                Device = [string](Get-Prop $_ 'FriendlyName' (Get-Prop $_ 'Name' ''))
                Class  = [string](Get-Prop $_ 'Class' '')
                Status = [string](Get-Prop $_ 'Status' '')
                Id     = [string](Get-Prop $_ 'InstanceId' '')
            }
        }
    )
}
if ($badDevices.Count -gt 0) {
    Add-Attention 'Warning' 'Devices in an error state' ("$($badDevices.Count) device(s) are reporting a problem - see the Devices section.")
    Add-Section 'Devices with problems' (New-DataTable -Rows $badDevices)
}

# --- defender ---
$mpResult = Invoke-WithTimeout -Seconds $ProbeTimeoutSeconds -Script {
    try {
        Get-MpComputerStatus -ErrorAction Stop |
            Select-Object RealTimeProtectionEnabled, AntivirusSignatureAge, AntivirusEnabled,
                          AntivirusSignatureVersion, QuickScanEndTime, FullScanEndTime, IsTamperProtected
    }
    catch { }
}
if (-not $mpResult.Completed) {
    Write-Step '! Defender status timed out - skipping that check'
}
$mp = @($mpResult.Output) | Select-Object -First 1

try {
    if ($null -eq $mp) { throw 'no Defender status' }
    $rtp = [bool](Get-Prop $mp 'RealTimeProtectionEnabled' $false)
    if (-not $rtp) {
        Add-Attention 'Critical' 'Real-time protection off' 'Microsoft Defender real-time protection is disabled.'
    }
    $sigAge = Get-Prop $mp 'AntivirusSignatureAge' 0
    if ($sigAge -ge 7) {
        Add-Attention 'Warning' 'Stale antivirus signatures' "Definitions are $sigAge days old."
    }

    $defPairs = @{
        'Real-time protection' = $(if ($rtp) { 'On' } else { 'OFF' })
        'Antivirus enabled'    = [string](Get-Prop $mp 'AntivirusEnabled' '')
        'Signature age'        = "$sigAge day(s)"
        'Signature version'    = [string](Get-Prop $mp 'AntivirusSignatureVersion' '')
        'Last quick scan'      = [string](Get-Prop $mp 'QuickScanEndTime' 'never')
        'Last full scan'       = [string](Get-Prop $mp 'FullScanEndTime' 'never')
        'Tamper protection'    = [string](Get-Prop $mp 'IsTamperProtected' '')
    }
    Add-Section 'Microsoft Defender' (New-KeyValueTable -Pairs $defPairs -Order @(
        'Real-time protection','Antivirus enabled','Signature age','Signature version',
        'Last quick scan','Last full scan','Tamper protection'))
}
catch {
    # Third-party AV installed, or the Defender module is unavailable.
}

# --- windows update ---
# Get-HotFix is one of the slower calls in this script, so it is made once and
# both the table and the staleness check read the same result.
$hotfixes = @()
try {
    $allHotfixes = @(Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending)

    $hotfixes = @(
        $allHotfixes | Select-Object -First 10 | ForEach-Object {
            [pscustomobject]@{
                Update      = [string](Get-Prop $_ 'HotFixID' '')
                Description = [string](Get-Prop $_ 'Description' '')
                Installed   = [string](Get-Prop $_ 'InstalledOn' '')
            }
        }
    )

    $newest = @($allHotfixes | Where-Object { Get-Prop $_ 'InstalledOn' } | Select-Object -First 1)
    if ($newest.Count -gt 0) {
        $age = ($RunStart - [datetime](Get-Prop $newest[0] 'InstalledOn')).TotalDays
        if ($age -gt 60) {
            Add-Attention 'Warning' 'Windows Update looks stale' ("The most recent update was installed {0:N0} days ago." -f $age)
        }
    }
}
catch { }
if ($hotfixes.Count -gt 0) {
    Add-Section 'Recent updates' (New-DataTable -Rows $hotfixes)
}

# --- startup item count ---
$startupCount = 0
foreach ($p in @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run')) {
    if (-not (Test-Path -LiteralPath $p)) { continue }
    try {
        $item = Get-ItemProperty -LiteralPath $p -ErrorAction Stop
        $startupCount += @($item.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }).Count
    }
    catch { }
}
foreach ($f in @([System.Environment]::GetFolderPath('Startup'), [System.Environment]::GetFolderPath('CommonStartup'))) {
    if ($f -and (Test-Path -LiteralPath $f)) {
        $startupCount += @(Get-ChildItem -LiteralPath $f -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ne 'desktop.ini' }).Count
    }
}
if ($startupCount -gt 15) {
    Add-Attention 'Info' 'Many startup items' "$startupCount entries launch at logon. Run Get-StartupImpact.ps1 for the breakdown."
}

# --- temperature (usually unavailable, so never a failure) ---
$temps = @()
try {
    $temps = @(
        Get-CimInstance -Namespace 'root\wmi' -ClassName 'MSAcpi_ThermalZoneTemperature' -ErrorAction Stop | ForEach-Object {
            $deci = [double](Get-Prop $_ 'CurrentTemperature' 0)
            [pscustomobject]@{
                Zone      = [string](Get-Prop $_ 'InstanceName' '')
                Celsius   = [math]::Round(($deci / 10) - 273.15, 1)
            }
        }
    )
}
catch { }
if ($temps.Count -gt 0) {
    Add-Section 'Thermal zones' (New-DataTable -Rows $temps) 'Most consumer firmware does not expose ACPI thermal zones; an empty or constant value here is normal.'
}

# ------------------------------------------------------------- 8. processes
if ($TopProcesses -gt 0) {
    Write-Step 'Reading top processes...'
    # Every one of these properties can throw: on a protected process the
    # session cannot open, or on one that exits between the enumeration and the
    # read. Each is fetched defensively so one bad row does not lose the table.
    $procs = @(
        Get-Process -ErrorAction SilentlyContinue |
            Sort-Object WorkingSet64 -Descending |
            Select-Object -First $TopProcesses |
            ForEach-Object {
                $p = $_
                [pscustomobject]@{
                    Process   = Get-SafeValue { $p.ProcessName } ''
                    PID       = Get-SafeValue { $p.Id } ''
                    Memory    = Format-Size (Get-SafeValue { $p.WorkingSet64 } 0)
                    'CPU (s)' = Get-SafeValue { [math]::Round($p.CPU, 1) } ''
                    Threads   = Get-SafeValue { $p.Threads.Count } ''
                    Started   = Get-SafeValue { $p.StartTime.ToString('yyyy-MM-dd HH:mm') } ''
                }
            }
    )
    Add-Section "Top $TopProcesses processes by memory" (New-DataTable -Rows $procs)
}

# --------------------------------------------------------- 9. build the page
Write-Step 'Rendering HTML...'

$attentionHtml = ''
if ($Attention.Count -eq 0) {
    $attentionHtml = '<div class="ok">Nothing needs attention. Every check in this report passed.</div>'
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
.note { color:var(--muted); font-size:13px; margin:-6px 0 14px; }
.empty { color:var(--muted); font-style:italic; margin:0; }
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
.bar { background:var(--bar); border-radius:999px; height:8px; width:150px; overflow:hidden; }
.bar span { display:block; height:100%; background:var(--barfill); }
.bar.warn span { background:var(--warn); }
.bar.danger span { background:var(--crit); }
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
footer { color:var(--muted); font-size:12px; text-align:center; margin-top:28px; }
'@

$html = New-Object System.Collections.Generic.List[string]
$html.Add('<!doctype html>')
$html.Add('<html lang="en"><head><meta charset="utf-8">')
$html.Add('<meta name="viewport" content="width=device-width,initial-scale=1">')
$html.Add(('<title>System snapshot - {0}</title>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add('<style>' + $css + '</style></head><body><div class="wrap">')
$html.Add(('<h1>{0}</h1>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
$html.Add(('<p class="sub">System snapshot taken {0} &middot; {1}</p>' -f `
    (ConvertTo-HtmlText $RunStart.ToString('dddd, d MMMM yyyy HH:mm:ss')),
    (ConvertTo-HtmlText ([string](Get-Prop $os 'Caption' 'Windows')))))
$html.Add('<h2 style="font-size:15px;margin:0 0 12px;">Needs attention</h2>')
$html.Add($attentionHtml)
foreach ($s in $Sections) { $html.Add($s) }
$html.Add('<footer>Generated by Get-SystemSnapshot.ps1 - read-only, nothing on this machine was changed.</footer>')
$html.Add('</div></body></html>')

($html -join [Environment]::NewLine) | Set-Content -LiteralPath $ReportFile -Encoding utf8

# Keep the 20 most recent snapshots.
Get-ChildItem -LiteralPath $OutputDir -Filter 'snapshot_*.html' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 20 |
    Remove-Item -Force -ErrorAction SilentlyContinue

# ------------------------------------------------------------------ wrap up
if (-not $Quiet) {
    Write-Host ''
    if ($Attention.Count -eq 0) {
        Write-Host '  Nothing needs attention.' -ForegroundColor Green
    }
    else {
        Write-Host '  Needs attention:'
        foreach ($a in $Attention) {
            $c = 'Yellow'
            if ($a.Level -eq 'Critical') { $c = 'Red' }
            if ($a.Level -eq 'Info')     { $c = 'Cyan' }
            Write-Host ('    [{0}] {1} - {2}' -f $a.Level, $a.Title, $a.Detail) -ForegroundColor $c
        }
    }
    Write-Host ''
    Write-Host "Report : $ReportFile"
}

if ($Open) {
    # A machine with no handler registered for .html throws here, which would
    # otherwise lose the exit code the caller actually asked for.
    try { Start-Process -FilePath $ReportFile -ErrorAction Stop }
    catch {
        if (-not $Quiet) { Write-Host "  Could not open the report: $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}

if ($Attention.Count -gt 0) { exit 1 }
exit 0
