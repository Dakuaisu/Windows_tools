<#
.SYNOPSIS
    Inventories everything installed on this machine and emits a restore.ps1
    you can run on a fresh Windows install to get most of it back.

.DESCRIPTION
    Read-only. Nothing is installed, uninstalled or modified - the script only
    reads registry / CLI output and writes files into an output folder.

    Sources collected:
      * Registry uninstall keys (HKLM 64/32-bit + HKCU) - the authoritative
        "Programs and Features" list, including things winget never heard of.
      * winget export        -> a real winget-import-able JSON
      * Microsoft Store apps (Appx)
      * VS Code / Insiders / Cursor extensions
      * npm global packages, pip packages
      * PowerShell modules installed from a repository
      * Chocolatey / Scoop, if present

    The generated restore.ps1 is deliberately conservative: winget import and
    the per-ecosystem installs are live commands, while anything with no
    reliable automated install path is emitted as a commented checklist so you
    review it rather than trust it.

.PARAMETER OutputDir
    Parent folder for the dated inventory folder.
    Defaults to $env:USERPROFILE\Documents\SoftwareInventory. If that cannot be
    written to, the tool falls back to %TEMP%\SoftwareInventory and says so.

.PARAMETER Sources
    Which collectors to run. Default: all of them.

.PARAMETER ToolTimeoutSeconds
    How long to wait on any one external tool - winget, an editor CLI, npm,
    pip, choco, scoop, the Appx enumeration - before giving up on that
    collector and moving on. Default 120. None of these take a timeout of
    their own, and any of them can hang.

.PARAMETER IncludeVersions
    Pin exact versions in the winget export. Off by default - pinned versions
    frequently fail to restore because that exact build is no longer offered.

.PARAMETER Open
    Open the output folder in Explorer when finished.

.EXAMPLE
    .\Export-InstalledSoftware.ps1

.EXAMPLE
    .\Export-InstalledSoftware.ps1 -Sources Registry,Winget,VSCode -Open

.EXAMPLE
    # On the new machine:
    .\restore.ps1 -WhatIf      # preview
    .\restore.ps1              # go
#>
[CmdletBinding()]
param(
    # Binding a null into Join-Path fails before the script body runs, which is
    # a baffling way to die on a stripped or service-account environment.
    [string]$OutputDir = (Join-Path $(if ($env:USERPROFILE) { $env:USERPROFILE } else { [System.IO.Path]::GetTempPath() }) 'Documents\SoftwareInventory'),

    [ValidateSet('Registry','Winget','Store','VSCode','Npm','Pip','PSModules','Choco','Scoop')]
    [string[]]$Sources = @('Registry','Winget','Store','VSCode','Npm','Pip','PSModules','Choco','Scoop'),

    [ValidateRange(5, 3600)]
    [int]$ToolTimeoutSeconds = 120,

    [switch]$IncludeVersions,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
function Initialize-OutputDir {
    <#
        A redirected profile, or an -OutputDir nobody can write to, should not
        stop the inventory from running - it should put the files somewhere
        that works and say where.
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
$OutDir   = Initialize-OutputDir `
    -Preferred (Join-Path $OutputDir ('Inventory_{0}' -f $RunStart.ToString('yyyy-MM-dd'))) `
    -ToolName 'SoftwareInventory'

$Warnings = New-Object System.Collections.Generic.List[string]

function Write-Step {
    param([string]$Message)
    Write-Host ('  {0}' -f $Message)
}

function Write-Warn {
    param([string]$Message)
    $Warnings.Add($Message)
    Write-Host ('  ! {0}' -f $Message) -ForegroundColor Yellow
}

function Test-Tool {
    param([string]$Name)
    $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Resolve-Winget {
    <#
        winget ships as a Store "app execution alias". Interactive shells pick
        it up from PATH, but plenty of non-interactive hosts do not, so fall
        back to the fixed alias location before giving up on it.
    #>
    $cmd = Get-Command 'winget' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd) {
        $src = $cmd.PSObject.Properties['Source']
        if ($src -and $src.Value) { return [string]$src.Value }
    }
    $alias = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path -LiteralPath $alias) { return $alias }
    $null
}

function Get-Prop {
    # Strict-mode-safe property read.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function ConvertFrom-InstallDate {
    param($Raw)
    if (-not $Raw) { return $null }
    $s = ([string]$Raw).Trim()
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, 'yyyyMMdd', $null, 'None', [ref]$parsed)) { return $parsed }
    if ([datetime]::TryParse($s, [ref]$parsed)) { return $parsed }
    return $null
}

function Test-Source {
    param([string]$Name)
    $Sources -contains $Name
}

function Invoke-ExternalCommand {
    <#
        Runs an external program with a hard timeout and hands back its stdout.

        Every collector here shells out to something that can hang: winget
        pausing on a source refresh, an editor CLI waiting on a wedged instance,
        scoop doing a git fetch over a dead link. The `&` call operator has no
        timeout at all, so one stuck tool used to hang the whole inventory
        indefinitely with no way to tell what it was waiting on.

        Both streams are drained asynchronously - a program that fills the
        stderr pipe while we block reading stdout would otherwise deadlock.

        .cmd / .bat shims (code, npm, scoop) cannot be launched by CreateProcess
        directly, so those go through cmd.exe.
    #>
    param(
        [string]$FilePath,
        [string[]]$Arguments = @(),
        [int]$TimeoutSeconds = 120
    )

    $result = [pscustomobject]@{
        Success = $false; TimedOut = $false; ExitCode = -1; StdOut = ''; StdErr = ''
    }

    $exe = $FilePath
    if (-not [System.IO.File]::Exists($exe)) {
        $found = @(Get-Command -Name $FilePath -CommandType Application -ErrorAction SilentlyContinue) |
            Select-Object -First 1
        if ($found) { $exe = [string](Get-Prop $found 'Source' $FilePath) }
    }

    $argLine = (@($Arguments | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { [string]$_ }
    }) -join ' ')

    if ($exe -match '\.(cmd|bat)$') {
        # /s /c with the whole command wrapped in one more pair of quotes is the
        # documented form that survives a path containing spaces.
        $argLine = '/d /s /c ""{0}" {1}"' -f $exe, $argLine
        $exe     = Join-Path $env:SystemRoot 'System32\cmd.exe'
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $exe
    $psi.Arguments              = $argLine
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true

    $proc = $null
    try {
        $proc    = [System.Diagnostics.Process]::Start($psi)
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()

        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            $result.TimedOut = $true
            try { $proc.Kill() } catch { }
            return $result
        }

        $result.ExitCode = $proc.ExitCode
        try { $result.StdOut = [string]$outTask.Result } catch { }
        try { $result.StdErr = [string]$errTask.Result } catch { }
        $result.Success = ($proc.ExitCode -eq 0)
        return $result
    }
    catch {
        $result.StdErr = $_.Exception.Message
        return $result
    }
    finally {
        if ($proc) { try { $proc.Dispose() } catch { } }
    }
}

function ConvertTo-Lines {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    @($Text -split "`r?`n" | Where-Object { $_ -and $_.Trim() })
}

function Invoke-CmdletWithTimeout {
    <#
        Get-AppxPackage is a cmdlet, not a process, so it cannot be killed the
        way Invoke-ExternalCommand kills one - but on a machine with a damaged
        package store it can block for a very long time. Running it in its own
        runspace lets the inventory give up on it and carry on. The abandoned
        runspace is asked to stop but never waited on; waiting is the thing
        that would hang.
    #>
    param([scriptblock]$Script, [int]$Seconds = 120)

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

function ConvertTo-CommentText {
    <#
        A DisplayName straight out of the registry can contain a newline. In
        the generated restore.ps1's manual checklist that would break out of
        the leading '# ' and land the rest of the name on its own line as
        executable PowerShell.
    #>
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    (([string]$Text) -replace "[`r`n`t]+", ' ').Trim()
}

function Test-SafeCommandArgument {
    # Emitted bare into a generated command line, so anything that could be
    # read as a separator, a quote or a redirection is refused outright.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    ($Text -match '^[A-Za-z0-9._@/+-]+$')
}

Write-Host ''
Write-Host "=== Software inventory : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) ==="
Write-Host "Output: $OutDir"
Write-Host ''

# --------------------------------------------------------- 1. registry: apps
$Programs = @()

if (Test-Source 'Registry') {
    Write-Step 'Reading uninstall registry keys...'
    $collected = New-Object System.Collections.Generic.List[psobject]

    $uninstallRoots = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Scope = 'Machine (64-bit)' }
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'Machine (32-bit)' }
        @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Scope = 'User' }
    )

    foreach ($root in $uninstallRoots) {
        if (-not (Test-Path -LiteralPath $root.Path)) { continue }

        foreach ($sub in @(Get-ChildItem -LiteralPath $root.Path -ErrorAction SilentlyContinue)) {
            $key = $null
            try { $key = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction Stop } catch { continue }

            $name = Get-Prop $key 'DisplayName'
            if ([string]::IsNullOrWhiteSpace([string]$name)) { continue }

            # Skip OS plumbing: hidden components, patches, and child entries.
            if ((Get-Prop $key 'SystemComponent' 0) -eq 1) { continue }
            if (Get-Prop $key 'ParentKeyName')             { continue }
            if ([string](Get-Prop $key 'ReleaseType' '') -match 'Update|Hotfix|Security') { continue }

            $sizeKb = Get-Prop $key 'EstimatedSize' 0
            $sizeMb = $null
            if ($sizeKb) { $sizeMb = [math]::Round($sizeKb / 1024, 1) }

            $collected.Add([pscustomobject]@{
                Name            = [string]$name
                Version         = [string](Get-Prop $key 'DisplayVersion' '')
                Publisher       = [string](Get-Prop $key 'Publisher' '')
                Scope           = $root.Scope
                InstallDate     = ConvertFrom-InstallDate (Get-Prop $key 'InstallDate')
                InstallLocation = [string](Get-Prop $key 'InstallLocation' '')
                SizeMB          = $sizeMb
                UninstallString = [string](Get-Prop $key 'UninstallString' '')
            })
        }
    }

    # The same app registered in several hives shows up once.
    $Programs = @($collected | Sort-Object Name, Version -Unique | Sort-Object Name)

    $csv = Join-Path $OutDir 'installed-programs.csv'
    $Programs | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding utf8
    Write-Step ('  {0} program(s) -> installed-programs.csv' -f $Programs.Count)
}

# ---------------------------------------------------------------- 2. winget
$WingetApps   = @()
$WingetExport = Join-Path $OutDir 'winget-export.json'

if (Test-Source 'Winget') {
    $wingetExe = Resolve-Winget
    if (-not $wingetExe) {
        Write-Warn 'winget not found - skipping winget export.'
    }
    else {
        Write-Step "Running winget export ($wingetExe)..."
        $wgArgs = @('export', '--output', $WingetExport, '--accept-source-agreements', '--disable-interactivity')
        if ($IncludeVersions) { $wgArgs += '--include-versions' }

        $wg = Invoke-ExternalCommand -FilePath $wingetExe -Arguments $wgArgs -TimeoutSeconds $ToolTimeoutSeconds
        if ($wg.TimedOut) {
            Write-Warn "winget export timed out after $ToolTimeoutSeconds seconds and was stopped."
        }
        elseif (-not $wg.Success) {
            $detail = ConvertTo-CommentText ($wg.StdErr + ' ' + $wg.StdOut)
            if ($detail.Length -gt 200) { $detail = $detail.Substring(0, 200) + '...' }
            Write-Warn "winget export exited with code $($wg.ExitCode). $detail"
        }

        if (Test-Path -LiteralPath $WingetExport) {
            try {
                $json = Get-Content -LiteralPath $WingetExport -Raw | ConvertFrom-Json
                foreach ($src in @(Get-Prop $json 'Sources' @())) {
                    foreach ($pkg in @(Get-Prop $src 'Packages' @())) {
                        $WingetApps += [pscustomobject]@{
                            PackageIdentifier = [string](Get-Prop $pkg 'PackageIdentifier' '')
                            Version           = [string](Get-Prop $pkg 'Version' '')
                        }
                    }
                }
                Write-Step ('  {0} winget package(s) -> winget-export.json' -f $WingetApps.Count)
            }
            catch {
                Write-Warn "winget export produced unreadable JSON: $($_.Exception.Message)"
            }
        }
        else {
            Write-Warn 'winget export produced no file.'
        }
    }
}

# ----------------------------------------------------------- 3. store / appx
$StoreApps = @()
if (Test-Source 'Store') {
    Write-Step 'Enumerating Microsoft Store packages...'
    try {
        $appx = Invoke-CmdletWithTimeout -Seconds $ToolTimeoutSeconds -Script {
            try {
                Get-AppxPackage -ErrorAction Stop |
                    Where-Object { -not $_.IsFramework -and $_.SignatureKind -ne 'System' } |
                    Select-Object Name, Publisher, Version
            }
            catch { }
        }
        if (-not $appx.Completed) {
            Write-Warn "Appx enumeration timed out after $ToolTimeoutSeconds seconds - store-apps.csv not written."
        }
        else {
            $StoreApps = @(
                $appx.Output | ForEach-Object {
                    [pscustomobject]@{
                        Name      = [string](Get-Prop $_ 'Name' '')
                        Publisher = [string](Get-Prop $_ 'Publisher' '')
                        Version   = [string](Get-Prop $_ 'Version' '')
                    }
                } | Sort-Object Name
            )
            $StoreApps | Export-Csv -LiteralPath (Join-Path $OutDir 'store-apps.csv') -NoTypeInformation -Encoding utf8
            Write-Step ('  {0} store package(s) -> store-apps.csv' -f $StoreApps.Count)
        }
    }
    catch {
        Write-Warn "Could not enumerate Appx packages: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------- 4. editor extensions
$Editors = @()
if (Test-Source 'VSCode') {
    foreach ($editor in @('code', 'code-insiders', 'cursor', 'windsurf')) {
        if (-not (Test-Tool $editor)) { continue }
        Write-Step "Listing $editor extensions..."
        try {
            $r = Invoke-ExternalCommand -FilePath $editor -Arguments @('--list-extensions') -TimeoutSeconds $ToolTimeoutSeconds
            if ($r.TimedOut) {
                Write-Warn "$editor --list-extensions timed out after $ToolTimeoutSeconds seconds and was stopped."
                continue
            }
            $ext = @(ConvertTo-Lines $r.StdOut | ForEach-Object { $_.Trim() } | Sort-Object)
            if ($ext.Count -gt 0) {
                $Editors += [pscustomobject]@{ Command = $editor; Extensions = $ext }
                $ext | Set-Content -LiteralPath (Join-Path $OutDir "$editor-extensions.txt") -Encoding utf8
                Write-Step ('  {0} extension(s) -> {1}-extensions.txt' -f $ext.Count, $editor)
            }
        }
        catch {
            Write-Warn "Failed to list $editor extensions: $($_.Exception.Message)"
        }
    }
}

# ------------------------------------------------------------- 5. npm globals
$NpmPackages = @()
if (Test-Source 'Npm') {
    if (Test-Tool 'npm') {
        Write-Step 'Listing global npm packages...'
        try {
            # npm exits non-zero when it has anything to complain about but
            # still prints usable JSON, so the exit code is not the test here.
            $r = Invoke-ExternalCommand -FilePath 'npm' -Arguments @('ls','-g','--depth=0','--json') -TimeoutSeconds $ToolTimeoutSeconds
            if ($r.TimedOut) { throw "npm ls timed out after $ToolTimeoutSeconds seconds and was stopped." }
            $obj  = $r.StdOut | ConvertFrom-Json
            $deps = Get-Prop $obj 'dependencies'
            if ($deps) {
                $NpmPackages = @(
                    $deps.PSObject.Properties |
                        Where-Object { $_.Name -ne 'npm' } |
                        ForEach-Object {
                            [pscustomobject]@{ Name = $_.Name; Version = [string](Get-Prop $_.Value 'version' '') }
                        } | Sort-Object Name
                )
            }
            ($NpmPackages | Format-Table -AutoSize | Out-String) |
                Set-Content -LiteralPath (Join-Path $OutDir 'npm-global.txt') -Encoding utf8
            Write-Step ('  {0} global npm package(s) -> npm-global.txt' -f $NpmPackages.Count)
        }
        catch {
            Write-Warn "npm listing failed: $($_.Exception.Message)"
        }
    }
    else { Write-Step 'npm not found - skipped.' }
}

# ------------------------------------------------------------ 6. pip packages
$PipPackages = @()
if (Test-Source 'Pip') {
    $pipCmd = $null
    if     (Test-Tool 'pip')    { $pipCmd = @('pip') }
    elseif (Test-Tool 'python') { $pipCmd = @('python', '-m', 'pip') }
    elseif (Test-Tool 'py')     { $pipCmd = @('py', '-m', 'pip') }

    if ($pipCmd) {
        Write-Step 'Listing pip packages...'
        try {
            $exe  = $pipCmd[0]
            $rest = @($pipCmd | Select-Object -Skip 1) + @('list', '--format=json', '--disable-pip-version-check')
            $r    = Invoke-ExternalCommand -FilePath $exe -Arguments $rest -TimeoutSeconds $ToolTimeoutSeconds
            if ($r.TimedOut) { throw "pip list timed out after $ToolTimeoutSeconds seconds and was stopped." }
            $PipPackages = @($r.StdOut | ConvertFrom-Json | Sort-Object name)
            ($PipPackages | ForEach-Object { '{0}=={1}' -f $_.name, $_.version }) |
                Set-Content -LiteralPath (Join-Path $OutDir 'pip-requirements.txt') -Encoding utf8
            Write-Step ('  {0} pip package(s) -> pip-requirements.txt' -f $PipPackages.Count)
        }
        catch {
            Write-Warn "pip listing failed: $($_.Exception.Message)"
        }
    }
    else { Write-Step 'pip not found - skipped.' }
}

# ------------------------------------------------------- 7. powershell modules
$PSModules = @()
if (Test-Source 'PSModules') {
    Write-Step 'Listing gallery-installed PowerShell modules...'
    try {
        $PSModules = @(
            Get-InstalledModule -ErrorAction SilentlyContinue |
                Select-Object Name, @{n='Version';e={[string]$_.Version}}, Repository |
                Sort-Object Name
        )
        ($PSModules | Format-Table -AutoSize | Out-String) |
            Set-Content -LiteralPath (Join-Path $OutDir 'powershell-modules.txt') -Encoding utf8
        Write-Step ('  {0} module(s) -> powershell-modules.txt' -f $PSModules.Count)
    }
    catch {
        Write-Warn "Get-InstalledModule failed: $($_.Exception.Message)"
    }
}

# ----------------------------------------------------------- 8. choco / scoop
$ChocoPackages = @()
if ((Test-Source 'Choco') -and (Test-Tool 'choco')) {
    Write-Step 'Listing Chocolatey packages...'
    try {
        # The flag has to match the major version, in both directions:
        #   v1  `choco list` searches the REMOTE feed - without --local-only it
        #       would return thousands of packages that are not installed.
        #   v2  `choco list` is local by default and --local-only is deprecated;
        #       2.7 still tolerates it, but it is on its way out.
        $chocoArgs = @('list', '--limit-output')
        $ver = Invoke-ExternalCommand -FilePath 'choco' -Arguments @('--version') -TimeoutSeconds $ToolTimeoutSeconds
        $major = 0
        if ($ver.Success) {
            $m = [regex]::Match((ConvertTo-Lines $ver.StdOut | Select-Object -Last 1), '^\s*(\d+)')
            if ($m.Success) { $major = [int]$m.Groups[1].Value }
        }
        if ($major -gt 0 -and $major -lt 2) { $chocoArgs += '--local-only' }

        $r = Invoke-ExternalCommand -FilePath 'choco' -Arguments $chocoArgs -TimeoutSeconds $ToolTimeoutSeconds
        if ($r.TimedOut) { throw "choco list timed out after $ToolTimeoutSeconds seconds and was stopped." }

        $ChocoPackages = @(
            ConvertTo-Lines $r.StdOut |
                Where-Object { $_ -match '\|' } |
                ForEach-Object {
                    $parts = $_.Split('|')
                    [pscustomobject]@{ Name = $parts[0].Trim(); Version = $parts[1].Trim() }
                }
        )
        Write-Step ('  {0} choco package(s) (choco v{1})' -f $ChocoPackages.Count, $major)
    }
    catch { Write-Warn "choco listing failed: $($_.Exception.Message)" }
}

$ScoopApps = @()
if ((Test-Source 'Scoop') -and (Test-Tool 'scoop')) {
    Write-Step 'Exporting Scoop apps...'
    try {
        $r = Invoke-ExternalCommand -FilePath 'scoop' -Arguments @('export') -TimeoutSeconds $ToolTimeoutSeconds
        if ($r.TimedOut) { throw "scoop export timed out after $ToolTimeoutSeconds seconds and was stopped." }

        # Older scoop printed a plain list here rather than JSON. Writing that
        # to a file called scoop-export.json and telling people to `scoop
        # import` it would send them somewhere unhelpful.
        $isJson = $false
        try { $null = $r.StdOut | ConvertFrom-Json; $isJson = $true } catch { }

        if ($isJson) {
            $r.StdOut | Set-Content -LiteralPath (Join-Path $OutDir 'scoop-export.json') -Encoding utf8
            Write-Step '  -> scoop-export.json'
        }
        else {
            $r.StdOut | Set-Content -LiteralPath (Join-Path $OutDir 'scoop-apps.txt') -Encoding utf8
            Write-Warn 'scoop export did not return JSON (older scoop) - wrote scoop-apps.txt instead.'
        }

        $listed = Invoke-ExternalCommand -FilePath 'scoop' -Arguments @('list') -TimeoutSeconds $ToolTimeoutSeconds
        $ScoopApps = @(ConvertTo-Lines $listed.StdOut)
    }
    catch { Write-Warn "scoop export failed: $($_.Exception.Message)" }
}

# --------------------------------------------------------- 9. combined json
$osCaption = ''
try { $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption } catch { }

$inventory = [pscustomobject]@{
    Generated      = $RunStart.ToString('o')
    ComputerName   = $env:COMPUTERNAME
    UserName       = $env:USERNAME
    OS             = $osCaption
    OSVersion      = [string][System.Environment]::OSVersion.Version
    Programs       = @($Programs)
    WingetPackages = @($WingetApps)
    StoreApps      = @($StoreApps)
    Editors        = @($Editors)
    NpmGlobal      = @($NpmPackages)
    Pip            = @($PipPackages)
    PSModules      = @($PSModules)
    Chocolatey     = @($ChocoPackages)
    Scoop          = @($ScoopApps)
}
$inventory | ConvertTo-Json -Depth 6 |
    Set-Content -LiteralPath (Join-Path $OutDir 'inventory.json') -Encoding utf8

# --------------------------------------------------------- 10. restore script
Write-Step 'Generating restore.ps1...'

$r = New-Object System.Collections.Generic.List[string]
$r.Add('<#')
$r.Add('.SYNOPSIS')
$r.Add('    Rebuilds the software set captured by Export-InstalledSoftware.ps1.')
$r.Add('')
$r.Add('.DESCRIPTION')
$r.Add("    Generated $($RunStart.ToString('yyyy-MM-dd HH:mm:ss')) from $env:COMPUTERNAME.")
$r.Add('    Run with -WhatIf first. Every step is optional and independent - comment')
$r.Add('    out what you do not want before running for real.')
$r.Add('')
$r.Add('.EXAMPLE')
$r.Add('    .\restore.ps1 -WhatIf')
$r.Add('#>')
$r.Add('[CmdletBinding(SupportsShouldProcess)]')
$r.Add('param(')
$r.Add('    [switch]$SkipWinget,')
$r.Add('    [switch]$SkipEditors,')
$r.Add('    [switch]$SkipNpm,')
$r.Add('    [switch]$SkipPip,')
$r.Add('    [switch]$SkipModules')
$r.Add(')')
$r.Add('')
$r.Add('$ErrorActionPreference = ''Continue''')
$r.Add('$here = Split-Path -Parent $MyInvocation.MyCommand.Path')
$r.Add('')
$r.Add('function Invoke-Step {')
$r.Add('    param([string]$Label, [scriptblock]$Action)')
$r.Add('    Write-Host ""')
$r.Add('    Write-Host "==> $Label" -ForegroundColor Cyan')
$r.Add('    if ($PSCmdlet.ShouldProcess($Label)) { & $Action }')
$r.Add('    else { Write-Host "    (WhatIf) skipped" }')
$r.Add('}')
$r.Add('')

if ($WingetApps.Count -gt 0) {
    $r.Add('# --------------------------------------------------------------- winget')
    $r.Add(('# {0} package(s) captured.' -f $WingetApps.Count))
    $r.Add('if (-not $SkipWinget) {')
    $r.Add('    Invoke-Step "winget import" {')
    $r.Add('        $f = Join-Path $here "winget-export.json"')
    $r.Add('        if (Test-Path -LiteralPath $f) {')
    $r.Add('            winget import --import-file $f --accept-package-agreements --accept-source-agreements --ignore-unavailable --ignore-versions')
    $r.Add('        } else { Write-Warning "winget-export.json not found next to this script." }')
    $r.Add('    }')
    $r.Add('}')
    $r.Add('')
}

foreach ($ed in $Editors) {
    $r.Add(('# ---------------------------------------------------- {0} extensions' -f $ed.Command))
    $r.Add('if (-not $SkipEditors) {')
    $r.Add(("    Invoke-Step '{0} extensions' {{" -f $ed.Command))
    $r.Add(('        if (Get-Command {0} -ErrorAction SilentlyContinue) {{' -f $ed.Command))
    # Everything below is emitted bare into a generated command line, so
    # anything that is not a plain identifier becomes a comment rather than
    # something restore.ps1 would execute.
    foreach ($e in $ed.Extensions) {
        if (Test-SafeCommandArgument $e) {
            $r.Add(('            {0} --install-extension {1} --force' -f $ed.Command, $e))
        }
        else {
            $r.Add(('            # skipped, unexpected characters in id: {0}' -f (ConvertTo-CommentText $e)))
        }
    }
    $r.Add(('        }} else {{ Write-Warning ''{0} not on PATH.'' }}' -f $ed.Command))
    $r.Add('    }')
    $r.Add('}')
    $r.Add('')
}

$npmSafe    = @($NpmPackages | Where-Object { Test-SafeCommandArgument ([string]$_.Name) })
$npmUnsafe  = @($NpmPackages | Where-Object { -not (Test-SafeCommandArgument ([string]$_.Name)) })
if ($npmSafe.Count -gt 0) {
    $r.Add('# ------------------------------------------------------ npm global packages')
    $r.Add('if (-not $SkipNpm) {')
    $r.Add('    Invoke-Step "npm globals" {')
    $r.Add('        if (Get-Command npm -ErrorAction SilentlyContinue) {')
    $r.Add(('            npm install -g {0}' -f (($npmSafe | ForEach-Object { $_.Name }) -join ' ')))
    $r.Add('        } else { Write-Warning "npm not on PATH." }')
    $r.Add('    }')
    $r.Add('}')
    foreach ($u in $npmUnsafe) {
        $r.Add(('# skipped, unexpected characters in package name: {0}' -f (ConvertTo-CommentText ([string]$u.Name))))
    }
    $r.Add('')
}

if ($PipPackages.Count -gt 0) {
    $r.Add('# ---------------------------------------------------------- pip packages')
    $r.Add('if (-not $SkipPip) {')
    $r.Add('    Invoke-Step "pip packages" {')
    $r.Add('        $req = Join-Path $here "pip-requirements.txt"')
    $r.Add('        if ((Get-Command pip -ErrorAction SilentlyContinue) -and (Test-Path -LiteralPath $req)) {')
    $r.Add('            pip install -r $req')
    $r.Add('        } else { Write-Warning "pip or pip-requirements.txt missing." }')
    $r.Add('    }')
    $r.Add('}')
    $r.Add('')
}

if ($PSModules.Count -gt 0) {
    $r.Add('# ----------------------------------------------------- powershell modules')
    $r.Add('if (-not $SkipModules) {')
    $r.Add('    Invoke-Step "PowerShell modules" {')
    foreach ($m in $PSModules) {
        $repo = 'PSGallery'
        if ($m.Repository) { $repo = $m.Repository }
        if ((Test-SafeCommandArgument ([string]$m.Name)) -and (Test-SafeCommandArgument ([string]$repo))) {
            $r.Add(("        Install-Module -Name '{0}' -Repository '{1}' -Scope CurrentUser -Force -AllowClobber -ErrorAction Continue" -f $m.Name, $repo))
        }
        else {
            $r.Add(('        # skipped, unexpected characters: {0} from {1}' -f (ConvertTo-CommentText ([string]$m.Name)), (ConvertTo-CommentText ([string]$repo))))
        }
    }
    $r.Add('    }')
    $r.Add('}')
    $r.Add('')
}

if ($ChocoPackages.Count -gt 0) {
    $r.Add('# ------------------------------------------------------------- chocolatey')
    $r.Add('# Requires an elevated shell. Uncomment to use.')
    foreach ($c in $ChocoPackages) {
        $r.Add(('# choco install {0} -y' -f (ConvertTo-CommentText ([string]$c.Name))))
    }
    $r.Add('')
}

if ($ScoopApps.Count -gt 0) {
    $r.Add('# ------------------------------------------------------------------ scoop')
    $r.Add('# scoop import (Join-Path $here "scoop-export.json")')
    $r.Add('')
}

# Registry apps winget cannot obviously map back to a package id: review-only.
function ConvertTo-MatchKey {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    (([string]$Text).ToLowerInvariant() -replace '[^a-z0-9]', '')
}

# Each winget id is Publisher.Product[.More]; all three forms are useful.
$wingetKeys = New-Object System.Collections.Generic.List[psobject]
foreach ($app in $WingetApps) {
    $id = [string]$app.PackageIdentifier
    if (-not $id) { continue }
    $segments = @($id -split '\.' | Where-Object { $_ })
    if ($segments.Count -eq 0) { continue }
    # Mid matters because plenty of ids end in a version or architecture rather
    # than the product: PostgreSQL.PostgreSQL.16, EclipseAdoptium.Temurin.21.JDK,
    # Microsoft.DotNet.SDK.7. For those the second segment is the real name.
    $mid = ''
    if ($segments.Count -ge 2) { $mid = ConvertTo-MatchKey $segments[1] }

    $wingetKeys.Add([pscustomobject]@{
        Full = ConvertTo-MatchKey ($segments -join '')
        Head = ConvertTo-MatchKey $segments[0]
        Mid  = $mid
        Tail = ConvertTo-MatchKey $segments[-1]
    })
}

function Test-CoveredByWinget {
    <#
        Is this Programs-and-Features entry plausibly one of the winget
        packages, and therefore already handled by `winget import` above?

        The old rule was "does the flattened display name CONTAIN the last
        segment of any winget id". On a real machine that quietly hid a third
        of the checklist behind coincidences: 'Microsoft .NET SDK 7.0.410 (x64)'
        matched the tail 'x64' from Microsoft.VCRedist.2015+.x64, 'Riot Client'
        matched 'cli' from GitHub.cli, and 'GitHub Desktop' matched 'git' from
        Git.Git. None of those would be restored by winget, and every one of
        them was omitted from the list of things to install by hand.

        So: an exact match against a segment or the whole id is trusted; a
        leading match is trusted when the thing matched is long enough to mean
        something; and anything looser has to be corroborated by the publisher.
        The failure mode is now over-listing, which costs a moment's reading,
        rather than under-listing, which costs software you did not notice was
        gone.
    #>
    param([string]$Name, [string]$Publisher)

    $n = ConvertTo-MatchKey $Name
    $p = ConvertTo-MatchKey $Publisher
    if (-not $n) { return $false }

    foreach ($k in $wingetKeys) {
        if (-not $k.Tail) { continue }

        $publisherAgrees = ($p -and $k.Head -and ($p.StartsWith($k.Head) -or $k.Head.StartsWith($p)))

        foreach ($name in @($k.Tail, $k.Mid)) {
            if (-not $name) { continue }
            if ($n -eq $name)                                     { return $true }
            if ($name.Length -ge 6 -and $n.StartsWith($name))     { return $true }
            if ($name.Length -ge 5 -and $publisherAgrees -and $n.Contains($name)) { return $true }
        }

        if ($n -eq $k.Full)                                  { return $true }
        if ($p -and ($p + $n) -eq $k.Full)                   { return $true }
        if ($k.Full.Length -ge 8 -and $n.StartsWith($k.Full)) { return $true }
    }
    $false
}

$manual  = New-Object System.Collections.Generic.List[string]
$covered = 0
foreach ($p in $Programs) {
    if (Test-CoveredByWinget -Name ([string]$p.Name) -Publisher ([string]$p.Publisher)) {
        $covered++
        continue
    }
    $ver = ''
    $pub = ''
    if ($p.Version)   { $ver = ' ({0})'  -f (ConvertTo-CommentText ([string]$p.Version)) }
    if ($p.Publisher) { $pub = ' - {0}'  -f (ConvertTo-CommentText ([string]$p.Publisher)) }
    # ConvertTo-CommentText on every field: a DisplayName containing a newline
    # would otherwise end the '# ' comment and put the rest on its own line as
    # executable PowerShell.
    $manual.Add(('#   {0}{1}{2}' -f (ConvertTo-CommentText ([string]$p.Name)), $ver, $pub))
}

$r.Add('# ============================================================================')
$r.Add('# MANUAL CHECKLIST - installed on the source machine, no reliable automated')
$r.Add('# restore path found. Review and install by hand whatever you still want.')
$r.Add(('# {0} of {1} registry program(s) listed; {2} looked like winget packages' -f $manual.Count, $Programs.Count, $covered))
$r.Add('# already covered by the import above. That mapping is a guess - if')
$r.Add('# something you wanted is missing after a restore, check here first.')
$r.Add('# ============================================================================')
foreach ($line in $manual) { $r.Add($line) }
$r.Add('')
$r.Add('Write-Host ""')
$r.Add('Write-Host "Restore pass complete. Review the manual checklist at the bottom of this script." -ForegroundColor Green')

$restorePath = Join-Path $OutDir 'restore.ps1'
($r -join [Environment]::NewLine) | Set-Content -LiteralPath $restorePath -Encoding utf8

# -------------------------------------------------------------------- readme
$warnBlock = '  (none)'
if ($Warnings.Count -gt 0) {
    $warnBlock = (($Warnings | ForEach-Object { "  * $_" }) -join [Environment]::NewLine)
}

$wingetLine = 'winget-export.json       (not produced - winget was unavailable)'
if (Test-Path -LiteralPath $WingetExport) {
    $wingetLine = 'winget-export.json       Feed to: winget import --import-file winget-export.json'
}

$readme = @"
Software inventory
==================
Machine   : $env:COMPUTERNAME
User      : $env:USERNAME
Generated : $($RunStart.ToString('yyyy-MM-dd HH:mm:ss'))

Files
-----
inventory.json           Everything, machine-readable.
installed-programs.csv   Programs & Features list.
$wingetLine
store-apps.csv           Store/Appx packages - reinstall from the Store by hand.
*-extensions.txt         Editor extension ids.
pip-requirements.txt     Feed to: pip install -r pip-requirements.txt
npm-global.txt           Global npm packages.
powershell-modules.txt   Gallery modules.
restore.ps1              Generated restore script. Run with -WhatIf first.

Counts
------
Programs         : $($Programs.Count)
winget packages  : $($WingetApps.Count)
Store apps       : $($StoreApps.Count)
Editors          : $($Editors.Count)
npm globals      : $($NpmPackages.Count)
pip packages     : $($PipPackages.Count)
PS modules       : $($PSModules.Count)
choco packages   : $($ChocoPackages.Count)

Warnings
--------
$warnBlock
"@

$readme | Set-Content -LiteralPath (Join-Path $OutDir 'README.txt') -Encoding utf8

# ------------------------------------------------------------------- wrap up
Write-Host ''
Write-Host $readme
Write-Host ''
Write-Host "Output folder : $OutDir"
Write-Host "Restore script: $restorePath"

if ($Open) {
    # -ArgumentList, quoted: the default output path sits under Documents and
    # would otherwise reach Explorer as several separate arguments.
    try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $OutDir) -ErrorAction Stop }
    catch { Write-Warn "Could not open the output folder: $($_.Exception.Message)" }
}
exit 0
