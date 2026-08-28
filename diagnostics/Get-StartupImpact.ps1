<#
.SYNOPSIS
    One table of everything that launches at boot or logon: Run keys, Startup
    folders, Winlogon, scheduled tasks and auto-start services - with
    publisher, signature status and a heuristic impact estimate.

.DESCRIPTION
    Read-only. Nothing is disabled or removed; the script only reads the
    registry, the Startup folders, the task scheduler and the service list, and
    tells you what it found so you can prune by hand.

    Task Manager only shows you Run keys and Startup folders. Half of what
    actually slows a boot is a scheduled task or an auto-start service, which
    is why those are folded into the same table here - along with the policy
    Run keys and a non-default Winlogon Userinit/Shell, neither of which Task
    Manager shows at all.

    Active Setup is deliberately NOT swept. It does run once per user at logon,
    but on any real machine it is dozens of rows of Microsoft component
    registration with nothing actionable in them, and burying four real
    findings under that is worse than not listing it.

    IMPORTANT: the Impact column is this script's own heuristic (binary size,
    signature, install location, delayed-start flag). It is NOT the number
    Task Manager shows - Windows derives that from boot tracing that is not
    exposed to scripts. Treat it as a sorting hint, not a measurement.

    A "Suspicious" flag is raised for entries that are unsigned AND living in
    a user-writable location (Temp, AppData, ProgramData), or whose target
    executable no longer exists. That combination is common for both malware
    and for leftovers from software you uninstalled years ago.

.PARAMETER IncludeServices
    Include Automatic-start services. On by default; -IncludeServices:$false
    to drop them.

.PARAMETER IncludeTasks
    Include logon/boot scheduled tasks. On by default.

.PARAMETER All
    Do not filter out Microsoft-published services and \Microsoft\ scheduled
    tasks. Adds several hundred rows of OS plumbing.

.PARAMETER OnlyEnabled
    Hide entries that Windows has already disabled.

.PARAMETER SkipSignatureCheck
    Skip Authenticode verification. Much faster, but loses the Signed and
    Suspicious columns.

.PARAMETER OutputDir
    Where the report and CSV land.
    Defaults to $env:LOCALAPPDATA\StartupImpact. If that cannot be written to,
    the tool falls back to %TEMP%\StartupImpact and says so.

.OUTPUTS
    Exit code 0 = nothing suspicious, 1 = at least one suspicious entry.

.EXAMPLE
    .\Get-StartupImpact.ps1

.EXAMPLE
    .\Get-StartupImpact.ps1 -OnlyEnabled -SkipSignatureCheck

.EXAMPLE
    .\Get-StartupImpact.ps1 -All -OutputDir C:\Temp\startup
#>
[CmdletBinding()]
param(
    [switch]$IncludeServices = $true,
    [switch]$IncludeTasks    = $true,
    [switch]$All,
    [switch]$OnlyEnabled,
    [switch]$SkipSignatureCheck,
    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir = (Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }) 'StartupImpact')
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

$OutputDir = Initialize-OutputDir -Preferred $OutputDir -ToolName 'StartupImpact'

$ReportFile = Join-Path $OutputDir "startup-report_$Stamp.txt"
$ReportCsv  = Join-Path $OutputDir "startup-report_$Stamp.csv"

$Entries      = New-Object System.Collections.Generic.List[psobject]
$FactsCache   = @{}
$UserWritable = @(
    $env:TEMP, $env:TMP,
    (Join-Path $env:USERPROFILE 'AppData'),
    $env:ProgramData,
    (Join-Path $env:USERPROFILE 'Downloads')
) | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') }

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

$ExeExtensions = @('.exe', '.com', '.bat', '.cmd', '.scr', '.pif')

# Interpreters and stubs that launch something else. Reporting one of these as
# the target is worse than useless here: they are all Microsoft-signed and sit
# in System32, so whatever they run inherits a clean bill of health it did not
# earn.
$WrapperNames = @(
    'cmd.exe', 'cmd', 'powershell.exe', 'powershell', 'pwsh.exe', 'pwsh',
    'wscript.exe', 'wscript', 'cscript.exe', 'cscript', 'mshta.exe', 'mshta',
    'explorer.exe', 'explorer'
)

function Test-LeafPath {
    # [System.IO.File]::Exists never throws - not on a path with illegal
    # characters, not on a disconnected UNC share - where Test-Path can.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try { return [System.IO.File]::Exists($Path) }
    catch { return $false }
}

function Resolve-CommandToken {
    <#
        Bare name -> full path via PATH. -CommandType Application matters:
        plain Get-Command happily returns the 'sc' *alias* for 'sc start x',
        whose Source is empty, and the caller then reported the literal string
        'sc' as though it were a file path.
    #>
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    if ($Token.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) { return $null }
    try {
        $c = @(Get-Command -Name $Token -CommandType Application -ErrorAction SilentlyContinue) |
            Select-Object -First 1
        if ($c) {
            $src = [string](Get-Prop $c 'Source' '')
            if ($src) { return $src }
        }
    }
    catch { }
    $null
}

function Get-ExecutableToken {
    <#
        Splits a command line into the executable and its arguments.

        The hard case is an unquoted path containing spaces and no extension -
        'C:\Program Files\Some App\launcher -silent'. That is indistinguishable
        from a command plus arguments until you ask the file system, which is
        exactly what Windows itself does: try the longest prefix first and work
        backwards. The old code split on the first space and reported
        'C:\Program'.
    #>
    param([string]$CommandLine)

    $cmd = ([string]$CommandLine).Trim()
    # Scheduled task actions often store an already-quoted Execute value, which
    # then gets quoted again when the command line is rebuilt.
    $cmd = ($cmd -replace '"{2,}', '"').Trim()
    if (-not $cmd) { return $null }

    if ($cmd.StartsWith('"')) {
        $end = $cmd.IndexOf('"', 1)
        if ($end -gt 1) {
            return [pscustomobject]@{ Token = $cmd.Substring(1, $end - 1); Rest = $cmd.Substring($end + 1).Trim() }
        }
        return [pscustomobject]@{ Token = $cmd.Trim('"'); Rest = '' }
    }

    $expanded = [System.Environment]::ExpandEnvironmentVariables($cmd)
    $parts    = @($expanded -split '\s+' | Where-Object { $_ })
    if ($parts.Count -eq 0) { return $null }

    # Bounded: a long command line is not going to be an unquoted path.
    $maxJoin = [math]::Min($parts.Count, 8)
    for ($take = $maxJoin; $take -ge 1; $take--) {
        $candidate = ($parts[0..($take - 1)] -join ' ')
        $hit = ''
        if (Test-LeafPath $candidate) { $hit = $candidate }
        else {
            foreach ($ext in $ExeExtensions) {
                if (Test-LeafPath ($candidate + $ext)) { $hit = $candidate + $ext; break }
            }
        }
        if ($hit) {
            $rest = ''
            if ($take -lt $parts.Count) { $rest = ($parts[$take..($parts.Count - 1)] -join ' ') }
            return [pscustomobject]@{ Token = $hit; Rest = $rest }
        }
    }

    # Nothing on disk matched. The first token may still be a bare command name
    # that only exists on PATH - 'cmd /c ...', 'rundll32.exe ...'. Resolving it
    # here rather than in the regex below keeps the arguments separate, which is
    # what lets the wrapper unwrapping in Get-ExecutablePath see what the
    # interpreter was actually told to run.
    if (Resolve-CommandToken $parts[0]) {
        $rest = ''
        if ($parts.Count -gt 1) { $rest = ($parts[1..($parts.Count - 1)] -join ' ') }
        return [pscustomobject]@{ Token = $parts[0]; Rest = $rest }
    }

    # A broken entry, then. Fall back to the first token that ends in an
    # executable extension, and finally to the first word.
    $m = [regex]::Match($cmd, '^(.*?\.(?:exe|com|bat|cmd|scr|pif|dll|ps1))(?:\s|$)', 'IgnoreCase')
    if ($m.Success) {
        return [pscustomobject]@{ Token = $m.Groups[1].Value; Rest = $cmd.Substring($m.Groups[1].Value.Length).Trim() }
    }
    $first = ($cmd -split '\s+')[0]
    return [pscustomobject]@{ Token = $first; Rest = $cmd.Substring($first.Length).Trim() }
}

function Get-WrappedTarget {
    <#
        What did an interpreter get told to run? Its own switches vary far too
        much to model (/c /k /s, -NoProfile, -WindowStyle Hidden, -File), so
        rather than parsing them this looks through the arguments for the first
        thing that is genuinely a file on disk. If nothing is, the wrapper
        itself stays the answer, which is the honest one.
    #>
    param([string]$Arguments)

    $a = ([string]$Arguments).Trim()
    if (-not $a) { return $null }
    $a = [System.Environment]::ExpandEnvironmentVariables($a)

    # Quoted candidates first: they are unambiguous.
    foreach ($m in [regex]::Matches($a, '"([^"]+)"')) {
        $c = $m.Groups[1].Value.Trim()
        if (Test-LeafPath $c) { return $c }
    }

    $tokens = @($a -split '\s+' | Where-Object { $_ })
    if ($tokens.Count -eq 0) { return $null }

    for ($start = 0; $start -lt $tokens.Count; $start++) {
        if ($tokens[$start] -match '^[-/]') { continue }
        $last = [math]::Min($tokens.Count - 1, $start + 7)
        for ($end = $last; $end -ge $start; $end--) {
            $c = ($tokens[$start..$end] -join ' ').Trim('"')
            if (Test-LeafPath $c) { return $c }
            foreach ($ext in $ExeExtensions) {
                if (Test-LeafPath ($c + $ext)) { return $c + $ext }
            }
        }
    }
    $null
}

function Get-ExecutablePath {
    <#
        Pulls the target executable out of a command line:
          "C:\Program Files\App\app.exe" --minimized  ->  C:\Program Files\App\app.exe
          C:\Program Files\Some App\launcher -silent  ->  C:\Program Files\Some App\launcher
          cmd /c "C:\Users\Public\x.exe"              ->  C:\Users\Public\x.exe
          rundll32.exe "C:\...\hook.dll",Start        ->  C:\...\hook.dll
          %SystemRoot%\system32\notepad.exe           ->  C:\Windows\system32\notepad.exe

        Wrappers are unwrapped one level. That matters for the Suspicious
        column: cmd.exe and rundll32.exe are Microsoft-signed and live in
        System32, so an entry that launches an unsigned binary from AppData
        *through* one of them used to score as harmless.
    #>
    param([string]$CommandLine, [int]$Depth = 0)

    if ([string]::IsNullOrWhiteSpace($CommandLine)) { return $null }

    $split = Get-ExecutableToken $CommandLine
    if ($null -eq $split) { return $null }

    $token = [System.Environment]::ExpandEnvironmentVariables([string]$split.Token).Trim('"', ' ')
    if ([string]::IsNullOrWhiteSpace($token)) { return $null }

    $rest = [string]$split.Rest
    $leaf = $token
    try { $leaf = [System.IO.Path]::GetFileName($token) } catch { }
    if (-not $leaf) { $leaf = $token }

    if ($Depth -lt 2 -and $rest) {
        if ($leaf -match '^(?:rundll32|regsvr32)(?:\.exe)?$') {
            # The module is the payload; the stub is not. rundll32 takes
            # "<dll>",Entry and regsvr32 takes switches before the path.
            $dll = ($rest -split ',')[0].Trim()
            $dll = ($dll -replace '^(?:[-/]\S+\s+)+', '').Trim().Trim('"')
            if ($dll) {
                $inner = Get-ExecutablePath -CommandLine $dll -Depth ($Depth + 1)
                if ($inner -and (Test-LeafPath $inner)) { return $inner }
            }
        }
        elseif ($WrapperNames -contains $leaf.ToLowerInvariant()) {
            $inner = Get-WrappedTarget -Arguments $rest
            if ($inner -and (Test-LeafPath $inner)) { return $inner }
        }
    }

    if (Test-LeafPath $token) { return $token }

    $resolved = Resolve-CommandToken $token
    if ($resolved) { return $resolved }

    $token
}

function Get-FileFacts {
    <#
        Cached by path. The services sweep asks for a binary's facts once to
        decide whether to filter it out and Add-Entry then asks again for the
        same binary - that is a second Get-Item and a second Authenticode
        verification each time, and Authenticode can go to the network for a
        revocation check.
    #>
    param([string]$Path)

    $key = ''
    if (-not [string]::IsNullOrWhiteSpace($Path)) { $key = $Path.ToLowerInvariant() }
    if ($key -and $FactsCache.ContainsKey($key)) { return $FactsCache[$key] }

    $facts = [pscustomobject]@{
        Exists    = $false
        SizeBytes = [long]0
        Publisher = ''
        Signed    = $null
        Product   = ''
    }
    if (-not (Test-LeafPath $Path)) {
        if ($key) { $FactsCache[$key] = $facts }
        return $facts
    }

    $facts.Exists = $true
    try {
        $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
        $facts.SizeBytes = $fi.Length
        $vi = $fi.VersionInfo
        if ($vi) {
            $facts.Publisher = [string](Get-Prop $vi 'CompanyName' '')
            $facts.Product   = [string](Get-Prop $vi 'ProductName' '')
        }
    }
    catch { }

    if (-not $SkipSignatureCheck) {
        $sig = $null
        try { $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop } catch { }
        if ($sig) {
            $facts.Signed = ($sig.Status -eq 'Valid')
            $subject = [string](Get-Prop (Get-Prop $sig 'SignerCertificate') 'Subject' '')
            if ($subject -and -not $facts.Publisher) {
                $m = [regex]::Match($subject, 'CN=([^,]+)')
                if ($m.Success) { $facts.Publisher = $m.Groups[1].Value.Trim('"', ' ') }
            }
        }
    }

    if ($key) { $FactsCache[$key] = $facts }
    $facts
}

function Test-UserWritablePath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    foreach ($root in $UserWritable) {
        if ($Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    $false
}

function Get-ImpactScore {
    <#
        Heuristic only - see the .DESCRIPTION note. Bigger binaries, unsigned
        code and user-writable install locations all correlate with slow,
        unmanaged startup items. Delayed-start services cost nothing at boot.
    #>
    param([long]$SizeBytes, $Signed, [string]$Path, [bool]$Delayed)

    $score = 0
    if     ($SizeBytes -ge 100MB) { $score += 3 }
    elseif ($SizeBytes -ge 25MB)  { $score += 2 }
    elseif ($SizeBytes -ge 5MB)   { $score += 1 }

    if ($Signed -eq $false)              { $score += 1 }
    if (Test-UserWritablePath $Path)     { $score += 1 }
    if ($Delayed)                        { $score -= 2 }

    if ($score -ge 4) { return 'High' }
    if ($score -ge 2) { return 'Medium' }
    'Low'
}

function Format-Size {
    param([long]$Bytes)
    if     ($Bytes -ge 1GB) { '{0:N2} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N1} MB' -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { '{0:N0} KB' -f ($Bytes / 1KB) }
    elseif ($Bytes -gt 0)   { "$Bytes B" }
    else                    { '' }
}

function Add-Entry {
    param(
        [string]$Source,
        [string]$Name,
        [string]$Command,
        [string]$Enabled = 'Enabled',
        [bool]  $Delayed = $false,
        [string]$Location = ''
    )

    $exe   = Get-ExecutablePath $Command
    $facts = Get-FileFacts $exe

    $suspicious = $false
    if ($exe) {
        if (-not $facts.Exists) { $suspicious = $true }
        elseif ((-not $SkipSignatureCheck) -and ($facts.Signed -eq $false) -and (Test-UserWritablePath $exe)) {
            $suspicious = $true
        }
    }

    $Entries.Add([pscustomobject]@{
        Source     = $Source
        Name       = $Name
        Enabled    = $Enabled
        Impact     = Get-ImpactScore -SizeBytes $facts.SizeBytes -Signed $facts.Signed -Path $exe -Delayed $Delayed
        Publisher  = $facts.Publisher
        Signed     = $facts.Signed
        Suspicious = $suspicious
        SizeBytes  = $facts.SizeBytes
        Size       = Format-Size $facts.SizeBytes
        Target     = $exe
        Command    = $Command
        Location   = $Location
    })
}

# --------------------------------------------------- Explorer approval state
# StartupApproved records whether the user disabled an entry in Task Manager.
# The value is a 12-byte blob whose first byte carries the state: 02 and 06
# mean enabled, 03 and 07 mean disabled - the low bit is the disable flag.
#
# Bytes 4-11 hold the time it was disabled, and that is deliberately NOT used
# as the signal: this machine has a 'Docker Desktop' entry sitting at
# 03 00 00 00 00 00 00 00 00 00 00 00 - disabled, with an all-zero timestamp -
# which a timestamp-based rule would report as enabled.
#
# An entry with no value here has never been touched in Task Manager, so it is
# enabled. Buckets are keyed by hive as well as by key: 'Discord' can exist in
# both HKCU\Run and HKLM\Run, and merging them let one hive's disabled state
# silently mask the other's. RunOnce and the policy Run keys get no bucket at
# all - Task Manager does not manage them, so there is nothing to look up.
$Approval = @{}

function Read-ApprovalKey {
    param([string]$KeyPath, [string]$Bucket)
    if (-not (Test-Path -LiteralPath $KeyPath)) { return }
    $item = $null
    try { $item = Get-ItemProperty -LiteralPath $KeyPath -ErrorAction Stop } catch { return }

    foreach ($prop in $item.PSObject.Properties) {
        if ($prop.Name -like 'PS*') { continue }
        $bytes = $prop.Value
        if ($bytes -isnot [byte[]] -or $bytes.Length -eq 0) { continue }
        $state = 'Enabled'
        if (($bytes[0] -band 1) -eq 1) { $state = 'Disabled' }
        $Approval["$Bucket|$($prop.Name)"] = $state
    }
}

$approvalRoot = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved'
Read-ApprovalKey "HKCU:\$approvalRoot\Run"           'HKCU:Run'
Read-ApprovalKey "HKCU:\$approvalRoot\Run32"         'HKCU:Run32'
Read-ApprovalKey "HKCU:\$approvalRoot\StartupFolder" 'HKCU:StartupFolder'
Read-ApprovalKey "HKLM:\$approvalRoot\Run"           'HKLM:Run'
Read-ApprovalKey "HKLM:\$approvalRoot\Run32"         'HKLM:Run32'
Read-ApprovalKey "HKLM:\$approvalRoot\StartupFolder" 'HKLM:StartupFolder'

function Get-ApprovalState {
    param([string]$Bucket, [string]$Name)
    if (-not $Bucket) { return 'Enabled' }
    $k = "$Bucket|$Name"
    if ($Approval.ContainsKey($k)) { return $Approval[$k] }
    'Enabled'
}

Write-Host ''
Write-Host "=== Startup inventory : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) ==="
Write-Host ''

# -------------------------------------------------------------- 1. Run keys
Write-Host '  Reading Run / RunOnce keys...'

# Policies\Explorer\Run is a real autostart location that Task Manager does not
# show at all, which is exactly the gap this tool exists to close.
$runKeys = @(
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run';                      Label = 'Run (HKLM)';           Approval = 'HKLM:Run' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce';                  Label = 'RunOnce (HKLM)';       Approval = '' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run';          Label = 'Run (HKLM 32-bit)';    Approval = 'HKLM:Run32' }
    @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce';      Label = 'RunOnce (HKLM 32-bit)';Approval = '' }
    @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run';                      Label = 'Run (HKCU)';           Approval = 'HKCU:Run' }
    @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce';                  Label = 'RunOnce (HKCU)';       Approval = '' }
    @{ Path = 'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run';          Label = 'Run (HKCU 32-bit)';    Approval = 'HKCU:Run32' }
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run';    Label = 'Policy Run (HKLM)';    Approval = '' }
    @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run';    Label = 'Policy Run (HKCU)';    Approval = '' }
)

foreach ($rk in $runKeys) {
    if (-not (Test-Path -LiteralPath $rk.Path)) { continue }
    $item = $null
    try { $item = Get-ItemProperty -LiteralPath $rk.Path -ErrorAction Stop } catch { continue }

    foreach ($prop in $item.PSObject.Properties) {
        if ($prop.Name -like 'PS*') { continue }
        if ([string]::IsNullOrWhiteSpace([string]$prop.Value)) { continue }

        Add-Entry -Source $rk.Label `
                  -Name $prop.Name `
                  -Command ([string]$prop.Value) `
                  -Enabled (Get-ApprovalState $rk.Approval $prop.Name) `
                  -Location $rk.Path
    }
}

# ---------------------------------------------------------- 1b. Winlogon
# Userinit and Shell run before anything else at logon, which makes them a
# classic place to hang a launcher. Only NON-default values are listed: on a
# healthy machine these read 'userinit.exe,' and 'explorer.exe', and printing
# those every run would be noise that trains people to ignore the section.
Write-Host '  Reading Winlogon Userinit / Shell...'

function Test-WinlogonDefault {
    param([string]$Name, [string]$Value)
    $v = [System.Environment]::ExpandEnvironmentVariables(([string]$Value).Trim().Trim('"')).ToLowerInvariant()
    $root = ([string]$env:SystemRoot).ToLowerInvariant().TrimEnd('\')
    if ($Name -eq 'Userinit') { return @('userinit.exe', "$root\system32\userinit.exe") -contains $v }
    if ($Name -eq 'Shell')    { return @('explorer.exe', "$root\explorer.exe") -contains $v }
    $false
}

foreach ($wl in @(
    @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'; Label = 'Winlogon (HKLM)' }
    @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'; Label = 'Winlogon (HKCU)' })) {

    if (-not (Test-Path -LiteralPath $wl.Path)) { continue }
    $item = $null
    try { $item = Get-ItemProperty -LiteralPath $wl.Path -ErrorAction Stop } catch { continue }

    foreach ($valueName in @('Userinit', 'Shell')) {
        $raw = [string](Get-Prop $item $valueName '')
        if (-not $raw) { continue }

        # Userinit is a comma-separated list; each element launches separately.
        foreach ($piece in @($raw -split ',')) {
            $p = $piece.Trim()
            if (-not $p) { continue }
            if (Test-WinlogonDefault -Name $valueName -Value $p) { continue }
            Add-Entry -Source $wl.Label `
                      -Name ('{0} (non-default)' -f $valueName) `
                      -Command $p `
                      -Location $wl.Path
        }
    }
}

# -------------------------------------------------------- 2. Startup folders
Write-Host '  Reading Startup folders...'

$startupFolders = @(
    @{ Path = [System.Environment]::GetFolderPath('Startup');       Label = 'Startup folder (user)';      Approval = 'HKCU:StartupFolder' }
    @{ Path = [System.Environment]::GetFolderPath('CommonStartup'); Label = 'Startup folder (all users)'; Approval = 'HKLM:StartupFolder' }
)

$shell = $null
try { $shell = New-Object -ComObject WScript.Shell } catch { }

foreach ($sf in $startupFolders) {
    if (-not $sf.Path -or -not (Test-Path -LiteralPath $sf.Path)) { continue }

    foreach ($file in @(Get-ChildItem -LiteralPath $sf.Path -File -Force -ErrorAction SilentlyContinue)) {
        if ($file.Name -ieq 'desktop.ini') { continue }

        $command = $file.FullName
        if ($file.Extension -ieq '.lnk' -and $shell) {
            try {
                $lnk    = $shell.CreateShortcut($file.FullName)
                $target = [string]$lnk.TargetPath
                if ($target) {
                    $command = $target
                    if ($lnk.Arguments) { $command = '"{0}" {1}' -f $target, $lnk.Arguments }
                }
            }
            catch { }
        }

        Add-Entry -Source $sf.Label `
                  -Name $file.BaseName `
                  -Command $command `
                  -Enabled (Get-ApprovalState $sf.Approval $file.Name) `
                  -Location $sf.Path
    }
}

if ($shell) {
    # WScript.Shell is a COM object; without this the RCW lives until the
    # runspace goes away, which for a scheduled run means holding a handle on
    # every shortcut it opened.
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch { }
    $shell = $null
}

# ------------------------------------------------------- 3. scheduled tasks
if ($IncludeTasks) {
    Write-Host '  Reading logon/boot scheduled tasks...'
    try {
        $tasks = @(Get-ScheduledTask -ErrorAction Stop)

        foreach ($task in $tasks) {
            $path = [string](Get-Prop $task 'TaskPath' '\')
            if (-not $All -and $path -like '\Microsoft\*') { continue }

            $triggers = @(Get-Prop $task 'Triggers' @())
            $kinds    = @()
            foreach ($t in $triggers) {
                $cls = ''
                try { $cls = [string]$t.CimClass.CimClassName } catch { }
                if ($cls -match 'LogonTrigger') { $kinds += 'Logon' }
                if ($cls -match 'BootTrigger')  { $kinds += 'Boot' }
            }
            if ($kinds.Count -eq 0) { continue }

            # Rebuild the command line from the task's first Execute action.
            $command = ''
            foreach ($a in @(Get-Prop $task 'Actions' @())) {
                # Execute may or may not already be quoted; normalise before
                # re-quoting, or the path parser sees "" and gives up.
                $exe = ([string](Get-Prop $a 'Execute' '')).Trim().Trim('"').Trim()
                if (-not $exe) { continue }
                $argStr = [string](Get-Prop $a 'Arguments' '')
                $command = $exe
                if ($argStr) { $command = '"{0}" {1}' -f $exe, $argStr }
                break
            }
            if (-not $command) { continue }

            $state = 'Enabled'
            if ([string](Get-Prop $task 'State' '') -eq 'Disabled') { $state = 'Disabled' }

            Add-Entry -Source ('Task ({0})' -f (($kinds | Select-Object -Unique) -join '+')) `
                      -Name $task.TaskName `
                      -Command $command `
                      -Enabled $state `
                      -Location $path
        }
    }
    catch {
        Write-Host "  ! Could not enumerate scheduled tasks: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ------------------------------------------------------------- 4. services
if ($IncludeServices) {
    Write-Host '  Reading Automatic-start services...'
    try {
        $services = @(Get-CimInstance Win32_Service -ErrorAction Stop | Where-Object { $_.StartMode -eq 'Auto' })

        foreach ($svc in $services) {
            $pathName = [string](Get-Prop $svc 'PathName' '')
            if (-not $pathName) { continue }

            # Delayed auto-start services do not compete for the boot window.
            $delayed = $false
            $svcKey  = "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)"
            if (Test-Path -LiteralPath $svcKey) {
                try {
                    $d = Get-ItemProperty -LiteralPath $svcKey -Name 'DelayedAutostart' -ErrorAction Stop
                    if ((Get-Prop $d 'DelayedAutostart' 0) -eq 1) { $delayed = $true }
                }
                catch { }
            }

            $exe   = Get-ExecutablePath $pathName
            $facts = Get-FileFacts $exe
            if (-not $All -and $facts.Publisher -match 'Microsoft') { continue }

            $state = 'Enabled'
            if ($delayed) { $state = 'Enabled (delayed)' }

            Add-Entry -Source 'Service (auto)' `
                      -Name ('{0} [{1}]' -f $svc.DisplayName, $svc.Name) `
                      -Command $pathName `
                      -Enabled $state `
                      -Delayed $delayed `
                      -Location 'Services'
        }
    }
    catch {
        Write-Host "  ! Could not enumerate services: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# --------------------------------------------------------------- the report
$final = @($Entries)
if ($OnlyEnabled) { $final = @($final | Where-Object { $_.Enabled -notlike 'Disabled*' }) }

$impactRank = @{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }
$final = @(
    $final | Sort-Object `
        @{ Expression = { $impactRank[$_.Impact] } },
        @{ Expression = { -not $_.Suspicious } },
        @{ Expression = { $_.SizeBytes }; Descending = $true },
        Name
)

$suspicious = @($final | Where-Object { $_.Suspicious })
$disabled   = @($final | Where-Object { $_.Enabled -like 'Disabled*' })
$missing    = @($final | Where-Object { $_.Target -and -not (Test-LeafPath $_.Target) })

$report = New-Object System.Collections.Generic.List[string]
$report.Add('Startup impact report')
$report.Add("Generated : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss'))")
$report.Add("Machine   : $env:COMPUTERNAME")
$report.Add('')
$report.Add('SUMMARY')
$report.Add("  total entries    : $($final.Count)")
$report.Add("  high impact      : $(@($final | Where-Object { $_.Impact -eq 'High' }).Count)")
$report.Add("  disabled already : $($disabled.Count)")
$report.Add("  suspicious       : $($suspicious.Count)")
$report.Add("  broken targets   : $($missing.Count)")
$report.Add('')
$report.Add('NOTE: Impact is a heuristic from binary size, signature, install location')
$report.Add('      and delayed-start flag. It is not Task Manager''s measured figure.')
$report.Add('')

$byGroup = $final | Group-Object Source | Sort-Object Name
$report.Add('BY SOURCE')
foreach ($g in $byGroup) {
    $report.Add(('    {0,-28} {1}' -f $g.Name, $g.Count))
}
$report.Add('')

$report.Add(('{0,-8} {1,-18} {2,-28} {3,-30} {4,9}  {5}' -f 'IMPACT','STATE','NAME','PUBLISHER','SIZE','TARGET'))
$report.Add('-' * 148)
foreach ($e in $final) {
    $flag = ' '
    if ($e.Suspicious) { $flag = '!' }
    $name = $e.Name
    if ($name.Length -gt 27) { $name = $name.Substring(0, 24) + '...' }
    $pub = $e.Publisher
    if ($pub.Length -gt 29) { $pub = $pub.Substring(0, 26) + '...' }

    $report.Add(('{0}{1,-7} {2,-18} {3,-28} {4,-30} {5,9}  {6}' -f `
        $flag, $e.Impact, $e.Enabled, $name, $pub, $e.Size, $e.Target))
}

if ($suspicious.Count -gt 0) {
    $report.Add('')
    $report.Add('SUSPICIOUS (unsigned in a user-writable path, or target missing)')
    $report.Add('-' * 140)
    foreach ($s in $suspicious) {
        $report.Add(('  [{0}] {1}' -f $s.Source, $s.Name))
        $report.Add(('        {0}' -f $s.Command))
    }
    $report.Add('')
    $report.Add('  These are often just leftovers from uninstalled software. Verify')
    $report.Add('  before deleting anything - check the publisher and the file date.')
}

$reportText = $report -join [Environment]::NewLine
$reportText | Set-Content -LiteralPath $ReportFile -Encoding utf8
$final | Select-Object Source, Name, Enabled, Impact, Publisher, Signed, Suspicious, Size, Target, Command, Location |
    Export-Csv -LiteralPath $ReportCsv -NoTypeInformation -Encoding utf8

# Keep the 12 most recent report pairs.
Get-ChildItem -LiteralPath $OutputDir -Filter 'startup-report_*' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 24 |
    Remove-Item -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host $reportText
Write-Host ''
Write-Host "Report : $ReportFile"
Write-Host "CSV    : $ReportCsv"

if ($suspicious.Count -gt 0) { exit 1 }
exit 0
