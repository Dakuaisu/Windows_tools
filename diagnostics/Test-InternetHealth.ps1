<#
.SYNOPSIS
    Measures connection health - gateway, DNS, latency, jitter, packet loss and
    HTTP time-to-first-byte - and appends one row per run to a CSV history you
    can point at when your ISP says "we see no problem on our end".

.DESCRIPTION
    Read-only apart from its own log files. Runs four probe groups:

      1. LAN      - ping the default gateway. Separates "my wifi is bad" from
                    "the internet is bad", which is the single most useful
                    distinction when something feels slow.
      2. DNS      - times name resolution against several hostnames. Slow DNS
                    feels exactly like slow internet but has a different fix.
      3. Latency  - ICMP to public anchors; reports avg/min/max, jitter
                    (stddev of round-trips) and packet loss.
      4. HTTP     - time to first byte for real HTTPS endpoints, which is what
                    a browser actually waits on.

    Throughput is only measured with -SpeedTest, because it downloads real
    data and you do not want that on a metered connection or on every run of a
    scheduled task.

    Uses System.Net.NetworkInformation.Ping rather than Test-Connection so the
    numbers mean the same thing on Windows PowerShell 5.1 and PowerShell 7.

.PARAMETER Count
    Number of full passes to run. Default 1. Use with -IntervalSeconds to
    watch a connection over time.

.PARAMETER IntervalSeconds
    Seconds to wait between passes. Default 60.

.PARAMETER PingCount
    ICMP echoes per target per pass. Default 10.

.PARAMETER Targets
    Public IPs to ping. Default Cloudflare, Google and Quad9 resolvers.

.PARAMETER HttpTargets
    URLs to time. Default a small, fast, well-distributed set.

.PARAMETER SpeedTest
    Also measure download throughput. Off by default; downloads ~25 MB.

.PARAMETER SpeedTestBytes
    How many bytes to pull for the throughput sample. Default 25000000.

.PARAMETER LatencyWarnMs
    Average round-trip above this is reported as degraded. Default 120.

.PARAMETER LossWarnPercent
    Packet loss at or above this is reported as degraded. Default 2.

.PARAMETER SkipCaptivePortalCheck
    Skip the hotel/airport sign-in-page check. That check asks a plain-HTTP
    endpoint with a known response body; anything else answering means
    something is intercepting the connection. It exists because every HTTPS
    target here fails behind a portal, which looks identical to a DNS fault.

.PARAMETER HistoryLimit
    Rows kept in history.csv. Default 5000; 0 keeps everything. When the
    column set changes the old file is archived rather than deleted.

.PARAMETER Quiet
    Suppress per-probe console output; still writes history and sets the exit
    code. Intended for scheduled tasks.

.PARAMETER LogDir
    Where history.csv and the run log live.
    Defaults to $env:LOCALAPPDATA\InternetHealth. If that cannot be written to,
    the tool falls back to %TEMP%\InternetHealth and says so.

.OUTPUTS
    Exit code 0 = healthy, 1 = degraded, 2 = no usable connection. A captive
    portal counts as 2: nothing reaches the internet until you sign in.

.EXAMPLE
    .\Test-InternetHealth.ps1

.EXAMPLE
    .\Test-InternetHealth.ps1 -SpeedTest

.EXAMPLE
    # Watch for an hour, one pass a minute, then look at history.csv
    .\Test-InternetHealth.ps1 -Count 60 -IntervalSeconds 60 -Quiet

.EXAMPLE
    # Chart the history afterwards
    Import-Csv "$env:LOCALAPPDATA\InternetHealth\history.csv" | Select-Object -Last 50
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 100000)]
    [int]   $Count           = 1,
    [ValidateRange(0, 86400)]
    [int]   $IntervalSeconds = 60,
    [ValidateRange(1, 1000)]
    [int]   $PingCount       = 10,
    [string[]]$Targets       = @('1.1.1.1', '8.8.8.8', '9.9.9.9'),
    [string[]]$HttpTargets   = @('https://www.cloudflare.com', 'https://www.google.com', 'https://github.com'),
    [string[]]$DnsNames      = @('www.microsoft.com', 'github.com', 'cloudflare.com'),
    [switch]$SpeedTest,
    [long]  $SpeedTestBytes  = 25000000,
    [string]$SpeedTestUrl    = 'https://speed.cloudflare.com/__down?bytes={0}',
    [int]   $LatencyWarnMs   = 120,
    [double]$LossWarnPercent = 2,
    [switch]$SkipCaptivePortalCheck,
    [ValidateRange(0, 1000000)]
    [int]   $HistoryLimit    = 5000,
    [switch]$Quiet,
    # $env:LOCALAPPDATA is normally set, but binding a null into Join-Path fails
    # before the script body ever runs, which is a baffling way for a tool to
    # die on a stripped or service-account environment.
    [string]$LogDir          = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'InternetHealth')
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

$LogDir = Initialize-OutputDir -Preferred $LogDir -ToolName 'InternetHealth'

$HistoryCsv = Join-Path $LogDir 'history.csv'
$LogFile    = Join-Path $LogDir 'internet-health.log'
$Transcript = New-Object System.Collections.Generic.List[string]

# The probe Windows' own network status indicator uses. A captive portal has to
# intercept it to work at all, which is what makes it a reliable detector.
$CaptivePortalUrl    = 'http://www.msftconnecttest.com/connecttest.txt'
$CaptivePortalExpect = 'Microsoft Connect Test'

# 5.1 negotiates TLS 1.0/1.1 by default, which several of the HTTP targets now
# refuse. Add to whatever is already configured rather than replacing it, so a
# machine policy that enabled something else keeps it.
try {
    $wanted = [Net.SecurityProtocolType]::Tls12
    if ([enum]::GetNames([Net.SecurityProtocolType]) -contains 'Tls13') {
        $wanted = $wanted -bor [Net.SecurityProtocolType]::Tls13
    }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $wanted
}
catch { }

function Write-Probe {
    param(
        [string]$Message,
        [ValidateSet('OK','WARN','FAIL','INFO')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $Transcript.Add($line)
    if ($Quiet) { return }

    switch ($Level) {
        'OK'   { Write-Host ('  [ ok ] ' + $Message) -ForegroundColor Green }
        'WARN' { Write-Host ('  [warn] ' + $Message) -ForegroundColor Yellow }
        'FAIL' { Write-Host ('  [FAIL] ' + $Message) -ForegroundColor Red }
        default { Write-Host ('         ' + $Message) }
    }
}

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Get-StdDev {
    param([double[]]$Values)
    if ($null -eq $Values -or $Values.Count -lt 2) { return 0.0 }
    $mean = ($Values | Measure-Object -Average).Average
    $sum  = 0.0
    foreach ($v in $Values) { $sum += [math]::Pow($v - $mean, 2) }
    [math]::Sqrt($sum / ($Values.Count - 1))
}

function Get-DefaultGateway {
    # Get-NetRoute is the clean path; fall back to WMI on older/odd stacks.
    try {
        $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
            Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
            Sort-Object RouteMetric |
            Select-Object -First 1
        if ($route) { return [string]$route.NextHop }
    }
    catch { }

    try {
        $cfg = Get-CimInstance Win32_NetworkAdapterConfiguration -ErrorAction Stop |
            Where-Object { $_.IPEnabled -and $_.DefaultIPGateway } |
            Select-Object -First 1
        if ($cfg) { return [string]@($cfg.DefaultIPGateway)[0] }
    }
    catch { }

    $null
}

function Invoke-PingProbe {
    <#
        Returns avg/min/max/jitter/loss for one target, or $null if the target
        never answered at all.
    #>
    param([string]$Target, [int]$Echoes, [int]$TimeoutMs = 2000)

    $ping    = New-Object System.Net.NetworkInformation.Ping
    $times   = New-Object System.Collections.Generic.List[double]
    $sent    = 0
    $lost    = 0
    $payload = New-Object byte[] 32

    try {
        for ($i = 0; $i -lt $Echoes; $i++) {
            $sent++
            try {
                $reply = $ping.Send($Target, $TimeoutMs, $payload)
                if ($reply.Status -eq 'Success') { $times.Add([double]$reply.RoundtripTime) }
                else                             { $lost++ }
            }
            catch { $lost++ }
            Start-Sleep -Milliseconds 120
        }
    }
    finally {
        $ping.Dispose()
    }

    $lossPct = 0.0
    if ($sent -gt 0) { $lossPct = [math]::Round(100.0 * $lost / $sent, 1) }

    if ($times.Count -eq 0) {
        return [pscustomobject]@{
            Target = $Target; Sent = $sent; Received = 0; LossPercent = 100.0
            AvgMs = $null; MinMs = $null; MaxMs = $null; JitterMs = $null
        }
    }

    $stats = $times | Measure-Object -Average -Minimum -Maximum
    [pscustomobject]@{
        Target      = $Target
        Sent        = $sent
        Received    = $times.Count
        LossPercent = $lossPct
        AvgMs       = [math]::Round($stats.Average, 1)
        MinMs       = [math]::Round($stats.Minimum, 1)
        MaxMs       = [math]::Round($stats.Maximum, 1)
        JitterMs    = [math]::Round((Get-StdDev $times.ToArray()), 1)
    }
}

function Invoke-DnsProbe {
    param([string]$Name)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ok = $false
    $addr = ''

    try {
        # -DnsOnly/-NoHostsFile keeps the cache and hosts file out of the timing.
        $res = Resolve-DnsName -Name $Name -Type A -DnsOnly -NoHostsFile -QuickTimeout -ErrorAction Stop
        $first = @($res | Where-Object { (Get-Prop $_ 'IPAddress') }) | Select-Object -First 1
        if ($first) { $ok = $true; $addr = [string]$first.IPAddress }
    }
    catch {
        # Resolve-DnsName is absent on some SKUs; fall back to the .NET resolver.
        try {
            $ips = [System.Net.Dns]::GetHostAddresses($Name)
            if ($ips -and $ips.Count -gt 0) { $ok = $true; $addr = [string]$ips[0] }
        }
        catch { }
    }

    $sw.Stop()
    [pscustomobject]@{
        Name      = $Name
        Success   = $ok
        Address   = $addr
        ElapsedMs = [math]::Round($sw.Elapsed.TotalMilliseconds, 1)
    }
}

function Invoke-HttpProbe {
    <#
        Times to first byte: the stopwatch stops when response headers arrive,
        not when the body finishes downloading.
    #>
    param([string]$Url, [int]$TimeoutMs = 10000)

    $sw   = [System.Diagnostics.Stopwatch]::StartNew()
    $resp = $null
    $code = ''
    $ok   = $false
    $err  = ''

    try {
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Method            = 'GET'
        $req.Timeout           = $TimeoutMs
        $req.ReadWriteTimeout  = $TimeoutMs
        $req.AllowAutoRedirect = $true
        $req.UserAgent         = 'Test-InternetHealth/1.0'
        $resp = $req.GetResponse()
        $sw.Stop()
        $code = [string][int]$resp.StatusCode
        $ok   = $true
    }
    catch [System.Net.WebException] {
        $sw.Stop()
        $err = $_.Exception.Message
        $r = $_.Exception.Response
        if ($r) { $code = [string][int]$r.StatusCode; $ok = $true }   # a 404 still proves connectivity
    }
    catch {
        $sw.Stop()
        $err = $_.Exception.Message
    }
    finally {
        if ($resp) { $resp.Close() }
    }

    [pscustomobject]@{
        Url        = $Url
        Success    = $ok
        StatusCode = $code
        TtfbMs     = [math]::Round($sw.Elapsed.TotalMilliseconds, 1)
        Error      = $err
    }
}

function Test-CaptivePortal {
    <#
        Every -HttpTargets entry is HTTPS, and a portal cannot forge a valid
        certificate for those - so behind a portal they simply fail, and the
        script blames DNS or the proxy. This asks a plain-HTTP endpoint whose
        exact response body is known instead: a redirect, or any other body,
        means something is sitting in the middle answering for it.

        Failing to reach it at all is NOT reported as a portal. That is an
        ordinary outage and the other probes already cover it - claiming a
        portal on no evidence would send people hunting for a login page that
        does not exist.
    #>
    param([string]$Url, [string]$Expect, [int]$TimeoutMs = 8000)

    $resp = $null
    try {
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Method            = 'GET'
        $req.Timeout           = $TimeoutMs
        $req.ReadWriteTimeout  = $TimeoutMs
        $req.AllowAutoRedirect = $false            # the redirect *is* the signal
        $req.UserAgent         = 'Test-InternetHealth/1.0'
        $resp = $req.GetResponse()

        $code = [int]$resp.StatusCode
        if ($code -ge 300 -and $code -lt 400) {
            return [pscustomobject]@{ Checked = $true; Portal = $true; Detail = "HTTP $code redirect to a sign-in page" }
        }

        $body = ''
        $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
        try { $body = $reader.ReadToEnd() } finally { $reader.Dispose() }

        if ($body.Trim() -eq $Expect) {
            return [pscustomobject]@{ Checked = $true; Portal = $false; Detail = 'connectivity probe answered correctly' }
        }
        return [pscustomobject]@{ Checked = $true; Portal = $true; Detail = "HTTP $code but the body was not the expected sentinel" }
    }
    catch {
        return [pscustomobject]@{ Checked = $false; Portal = $false; Detail = $_.Exception.Message }
    }
    finally {
        if ($resp) { $resp.Close() }
    }
}

function Invoke-SpeedProbe {
    param([string]$UrlTemplate, [long]$Bytes)

    $url = $UrlTemplate
    if ($UrlTemplate -like '*{0}*') { $url = $UrlTemplate -f $Bytes }

    $resp   = $null
    $stream = $null
    try {
        $req = [System.Net.HttpWebRequest]::Create($url)
        $req.Method    = 'GET'
        $req.Timeout   = 20000
        $req.ReadWriteTimeout = 60000
        $req.UserAgent = 'Test-InternetHealth/1.0'

        $sw     = [System.Diagnostics.Stopwatch]::StartNew()
        $resp   = $req.GetResponse()
        $stream = $resp.GetResponseStream()

        # Both caps matter: ReadWriteTimeout only bounds a single read, so a
        # server that dribbles bytes forever - or a -SpeedTestUrl pointing at
        # something far larger than -SpeedTestBytes - would otherwise download
        # without end on someone's metered connection.
        $maxSeconds = 120
        $buffer = New-Object byte[] 81920
        $total  = [long]0
        while ($true) {
            $read = $stream.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $total += $read
            if ($total -ge $Bytes) { break }
            if ($sw.Elapsed.TotalSeconds -ge $maxSeconds) { break }
        }
        $sw.Stop()

        $seconds = $sw.Elapsed.TotalSeconds
        $mbps    = 0.0
        if ($seconds -gt 0) { $mbps = [math]::Round(($total * 8) / $seconds / 1000000, 2) }

        return [pscustomobject]@{
            Success   = $true
            Bytes     = $total
            Seconds   = [math]::Round($seconds, 2)
            Mbps      = $mbps
            Error     = ''
        }
    }
    catch {
        return [pscustomobject]@{
            Success = $false; Bytes = 0; Seconds = 0; Mbps = 0.0; Error = $_.Exception.Message
        }
    }
    finally {
        if ($stream) { $stream.Dispose() }
        if ($resp)   { $resp.Close() }
    }
}

# ------------------------------------------------------------------ one pass
function Invoke-Pass {
    param([int]$PassNumber)

    $passStart = Get-Date
    if (-not $Quiet) {
        Write-Host ''
        Write-Host ("=== Pass {0}/{1} : {2} ===" -f $PassNumber, $Count, $passStart.ToString('HH:mm:ss'))
    }

    $issues = New-Object System.Collections.Generic.List[string]

    # --- LAN -----------------------------------------------------------------
    $gateway    = Get-DefaultGateway
    $gatewayMs  = $null
    $gatewayUp  = $false
    if ($gateway) {
        $g = Invoke-PingProbe -Target $gateway -Echoes ([math]::Max(3, [int]($PingCount / 3)))
        $gatewayMs = $g.AvgMs
        $gatewayUp = ($g.Received -gt 0)
        if ($gatewayUp) {
            Write-Probe ("gateway {0}  avg {1} ms  loss {2}%" -f $gateway, $g.AvgMs, $g.LossPercent) 'OK'
            if ($g.LossPercent -ge $LossWarnPercent) {
                $issues.Add("gateway packet loss $($g.LossPercent)% - suspect wifi/cable, not the ISP")
            }
        }
        else {
            Write-Probe ("gateway {0} unreachable" -f $gateway) 'FAIL'
            $issues.Add('default gateway unreachable - local network problem')
        }
    }
    else {
        Write-Probe 'no default gateway found - not connected to any network' 'FAIL'
        $issues.Add('no default gateway')
    }

    # --- DNS -----------------------------------------------------------------
    $dnsResults = @(foreach ($n in $DnsNames) { Invoke-DnsProbe -Name $n })
    $dnsOk      = @($dnsResults | Where-Object { $_.Success })
    $dnsAvgMs   = $null
    if ($dnsOk.Count -gt 0) {
        $dnsAvgMs = [math]::Round((($dnsOk | Measure-Object -Property ElapsedMs -Average).Average), 1)
    }

    foreach ($d in $dnsResults) {
        if ($d.Success) {
            $lvl = 'OK'
            if ($d.ElapsedMs -gt 300) { $lvl = 'WARN' }
            Write-Probe ("dns {0,-22} {1,7} ms  -> {2}" -f $d.Name, $d.ElapsedMs, $d.Address) $lvl
        }
        else {
            Write-Probe ("dns {0,-22} FAILED" -f $d.Name) 'FAIL'
        }
    }
    if ($dnsOk.Count -eq 0)                    { $issues.Add('DNS resolution failing for every test name') }
    elseif ($dnsAvgMs -and $dnsAvgMs -gt 300)  { $issues.Add("DNS slow (avg $dnsAvgMs ms) - consider changing resolver") }

    # --- latency -------------------------------------------------------------
    $pingResults = @(foreach ($t in $Targets) { Invoke-PingProbe -Target $t -Echoes $PingCount })
    $reachable   = @($pingResults | Where-Object { $_.Received -gt 0 })

    foreach ($p in $pingResults) {
        if ($p.Received -eq 0) {
            Write-Probe ("ping {0,-16} no reply" -f $p.Target) 'FAIL'
            continue
        }
        $lvl = 'OK'
        if ($p.AvgMs -gt $LatencyWarnMs -or $p.LossPercent -ge $LossWarnPercent) { $lvl = 'WARN' }
        Write-Probe ("ping {0,-16} avg {1,6} ms  min {2,6}  max {3,6}  jitter {4,5}  loss {5}%" -f `
            $p.Target, $p.AvgMs, $p.MinMs, $p.MaxMs, $p.JitterMs, $p.LossPercent) $lvl
    }

    $avgLatency = $null
    $avgLoss    = 100.0
    $avgJitter  = $null
    if ($reachable.Count -gt 0) {
        $avgLatency = [math]::Round((($reachable | Measure-Object -Property AvgMs    -Average).Average), 1)
        $avgJitter  = [math]::Round((($reachable | Measure-Object -Property JitterMs -Average).Average), 1)
        $avgLoss    = [math]::Round((($pingResults | Measure-Object -Property LossPercent -Average).Average), 1)

        if ($avgLatency -gt $LatencyWarnMs)  { $issues.Add("high latency (avg $avgLatency ms)") }
        if ($avgLoss -ge $LossWarnPercent)   { $issues.Add("packet loss (avg $avgLoss%)") }
        if ($avgJitter -gt 30)               { $issues.Add("high jitter ($avgJitter ms) - calls and games will stutter") }
    }
    else {
        $issues.Add('no public host answered ICMP')
    }

    # --- HTTP ----------------------------------------------------------------
    $httpResults = @(foreach ($u in $HttpTargets) { Invoke-HttpProbe -Url $u })
    $httpOk      = @($httpResults | Where-Object { $_.Success })
    $avgTtfb     = $null
    if ($httpOk.Count -gt 0) {
        $avgTtfb = [math]::Round((($httpOk | Measure-Object -Property TtfbMs -Average).Average), 1)
    }

    foreach ($h in $httpResults) {
        if ($h.Success) {
            $lvl = 'OK'
            if ($h.TtfbMs -gt 1000) { $lvl = 'WARN' }
            Write-Probe ("http {0,-34} {1}  ttfb {2} ms" -f $h.Url, $h.StatusCode, $h.TtfbMs) $lvl
        }
        else {
            Write-Probe ("http {0,-34} FAILED : {1}" -f $h.Url, $h.Error) 'FAIL'
        }
    }
    if ($httpOk.Count -eq 0) {
        $issues.Add('no HTTPS endpoint responded - DNS, proxy or captive portal problem')
    }
    elseif ($avgTtfb -gt 1500) {
        $issues.Add("slow HTTPS response (avg TTFB $avgTtfb ms)")
    }

    # ICMP blocked but HTTP fine is a normal, healthy configuration.
    if ($reachable.Count -eq 0 -and $httpOk.Count -gt 0) {
        Write-Probe 'ICMP appears blocked upstream, but HTTPS works - this is fine.' 'INFO'
        $issues.Remove('no public host answered ICMP') | Out-Null
    }

    # --- captive portal ------------------------------------------------------
    $portal = $false
    if (-not $SkipCaptivePortalCheck) {
        $cp = Test-CaptivePortal -Url $CaptivePortalUrl -Expect $CaptivePortalExpect
        if ($cp.Checked -and $cp.Portal) {
            $portal = $true
            Write-Probe ("captive portal detected - {0}" -f $cp.Detail) 'FAIL'
            $issues.Add('captive portal: this network wants you to sign in before it will carry traffic')
        }
        elseif ($cp.Checked) {
            Write-Probe ("captive portal check: none - {0}" -f $cp.Detail) 'OK'
        }
    }

    # --- throughput ----------------------------------------------------------
    $mbps = $null
    if ($SpeedTest) {
        Write-Probe ("downloading {0:N0} MB for a throughput sample..." -f ($SpeedTestBytes / 1MB)) 'INFO'
        $s = Invoke-SpeedProbe -UrlTemplate $SpeedTestUrl -Bytes $SpeedTestBytes
        if ($s.Success) {
            $mbps = $s.Mbps
            Write-Probe ("download {0} Mbps  ({1:N1} MB in {2}s)" -f $s.Mbps, ($s.Bytes / 1MB), $s.Seconds) 'OK'
        }
        else {
            Write-Probe ("speed test failed: {0}" -f $s.Error) 'WARN'
        }
    }

    # --- verdict -------------------------------------------------------------
    # A portal is 'Down', not 'Degraded': nothing on this connection reaches
    # the internet until somebody signs in, which is exactly what exit code 2
    # already means.
    $status = 'Healthy'
    if ($portal)                                          { $status = 'Down' }
    elseif ($httpOk.Count -eq 0 -and $reachable.Count -eq 0) { $status = 'Down' }
    elseif ($issues.Count -gt 0)                          { $status = 'Degraded' }

    [pscustomobject]@{
        Timestamp    = $passStart.ToString('yyyy-MM-dd HH:mm:ss')
        Status       = $status
        Gateway      = $gateway
        GatewayMs    = $gatewayMs
        GatewayUp    = $gatewayUp
        DnsAvgMs     = $dnsAvgMs
        DnsFailures  = ($dnsResults.Count - $dnsOk.Count)
        LatencyAvgMs = $avgLatency
        JitterMs     = $avgJitter
        LossPercent  = $avgLoss
        HttpAvgTtfbMs= $avgTtfb
        HttpFailures = ($httpResults.Count - $httpOk.Count)
        DownloadMbps = $mbps
        Issues       = ($issues -join '; ')
    }
}

# ---------------------------------------------------------------------- main
function Add-HistoryRow {
    <#
        Export-Csv -Append refuses outright when the file's header does not
        match the object being appended - so the first run after this script
        gains or loses a column would throw away a whole pass of work at the
        very last step. The stale file is archived under a dated name instead,
        which also keeps the old readings.

        Trimming happens here too: history.csv was the one file in this
        collection that grew without any bound at all.
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
            Write-Host ("  History columns changed - earlier readings kept as {0}" -f (Split-Path -Leaf $archive)) -ForegroundColor Yellow
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

$results = New-Object System.Collections.Generic.List[psobject]

for ($pass = 1; $pass -le $Count; $pass++) {
    $result = Invoke-Pass -PassNumber $pass
    $results.Add($result)

    # Append immediately, so a long -Count run still leaves data if interrupted.
    # A history the script cannot write is worth a warning, not a lost run.
    try { Add-HistoryRow -Row $result -Path $HistoryCsv -MaxRows $HistoryLimit }
    catch { Write-Host "  ! Could not update $HistoryCsv : $($_.Exception.Message)" -ForegroundColor Yellow }

    if (-not $Quiet) {
        $colour = 'Green'
        if ($result.Status -eq 'Degraded') { $colour = 'Yellow' }
        if ($result.Status -eq 'Down')     { $colour = 'Red' }
        Write-Host ''
        Write-Host ("  VERDICT: {0}" -f $result.Status) -ForegroundColor $colour
        if ($result.Issues) {
            foreach ($i in ($result.Issues -split '; ')) { Write-Host "    - $i" -ForegroundColor $colour }
        }
    }

    if ($pass -lt $Count) {
        if (-not $Quiet) { Write-Host ("  sleeping {0}s..." -f $IntervalSeconds) }
        Start-Sleep -Seconds $IntervalSeconds
    }
}

# Rolling log, trimmed to the last 5000 lines.
try {
    if ($Transcript.Count -gt 0) {
        Add-Content -LiteralPath $LogFile -Value $Transcript.ToArray() -Encoding utf8
    }
    $logLines = @(Get-Content -LiteralPath $LogFile -ErrorAction SilentlyContinue)
    if ($logLines.Count -gt 5000) {
        $logLines[-5000..-1] | Set-Content -LiteralPath $LogFile -Encoding utf8
    }
}
catch { Write-Host "  ! Could not update $LogFile : $($_.Exception.Message)" -ForegroundColor Yellow }

# ------------------------------------------------------------------- summary
$down     = @($results | Where-Object { $_.Status -eq 'Down' })
$degraded = @($results | Where-Object { $_.Status -eq 'Degraded' })

if (-not $Quiet) {
    Write-Host ''
    Write-Host '=== Summary ==='
    Write-Host ("  passes   : {0}" -f $results.Count)
    Write-Host ("  healthy  : {0}" -f @($results | Where-Object { $_.Status -eq 'Healthy' }).Count)
    Write-Host ("  degraded : {0}" -f $degraded.Count)
    Write-Host ("  down     : {0}" -f $down.Count)
    Write-Host ''
    Write-Host "History : $HistoryCsv"
    Write-Host "Log     : $LogFile"
}

if ($down.Count -gt 0)     { exit 2 }
if ($degraded.Count -gt 0) { exit 1 }
exit 0
