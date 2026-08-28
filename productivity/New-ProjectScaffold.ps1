<#
.SYNOPSIS
    Creates a new project folder with a README, a type-appropriate .gitignore,
    an .editorconfig, a CLAUDE.md, an initialised git repo and the starter
    files for the stack you picked - then opens it.

.DESCRIPTION
    The point is to remove the ten minutes of boilerplate between "I have an
    idea" and "I am writing the actual code", without generating a framework's
    worth of files you then have to read and delete.

    Non-destructive: an existing file is never overwritten. If the target
    folder already has contents, the script stops unless you pass -Force, and
    even then it only adds files that are not already there.

    Project types:
      blank       README, .gitignore, .editorconfig, CLAUDE.md and nothing else
      node        package.json, src/index.js, npm-flavoured .gitignore
      python      pyproject.toml, src/<package>/__init__.py, tests/, optional venv
      powershell  module manifest + .psm1, Public/, Private/, Tests/
      static      index.html, styles.css, script.js
      dotnet      common files first, then delegates to `dotnet new console`.
                  Run without --force, so a folder that already holds a
                  project is reported and left alone rather than overwritten.

.PARAMETER Name
    Project name. Becomes the folder name and seeds package/module names.

.PARAMETER Root
    Where the project folder is created.
    Defaults to $env:USERPROFILE\Documents\projects.

.PARAMETER Type
    Project type. Default blank.

.PARAMETER Description
    One-line description used in the README and in package metadata.

.PARAMETER License
    Add a LICENSE file. MIT or None. Default None.

.PARAMETER Author
    Name used in package metadata and the licence. Defaults to the git
    user.name setting, then to the Windows user name.

.PARAMETER CreateVenv
    Python only: also create a .venv virtual environment.

.PARAMETER NoGit
    Skip git init entirely.

.PARAMETER NoCommit
    Initialise the repo but do not make the initial commit.

.PARAMETER OpenWith
    What to launch when finished: code, claude, explorer, or none.
    Default code when VS Code is on PATH, otherwise none.

.PARAMETER Force
    Allow scaffolding into a folder that already has files in it.

.PARAMETER DryRun
    Show what would be created, touch nothing.

.OUTPUTS
    Exit code 0 = the project folder was scaffolded, 1 = nothing was created
    (the folder already had contents and -Force was not given, the name is not
    usable as a Windows folder, or -Root could not be written to). Optional
    steps that fail - dotnet, git, venv - are reported as warnings and still
    exit 0, because the folder itself was scaffolded.

.EXAMPLE
    .\New-ProjectScaffold.ps1 -Name my-thing -DryRun

.EXAMPLE
    .\New-ProjectScaffold.ps1 -Name scraper -Type python -CreateVenv -License MIT

.EXAMPLE
    .\New-ProjectScaffold.ps1 -Name PSToolbox -Type powershell -OpenWith claude
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidatePattern('^[^\\/:*?"<>|]+$')]
    [string]$Name,

    [string]$Root = (Join-Path $env:USERPROFILE 'Documents\projects'),

    [ValidateSet('blank','node','python','powershell','static','dotnet')]
    [string]$Type = 'blank',

    [string]$Description = '',

    [ValidateSet('None','MIT')]
    [string]$License = 'None',

    [string]$Author = '',

    [switch]$CreateVenv,
    [switch]$NoGit,
    [switch]$NoCommit,

    [ValidateSet('code','claude','explorer','none')]
    [string]$OpenWith = '',

    [switch]$Force,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
# ValidatePattern already blocks path separators, but it still lets through
# names Windows will not accept as a folder - and '..' would silently scaffold
# into the PARENT of -Root, which is never what anyone meant.
$reservedNames = @(
    'CON','PRN','AUX','NUL',
    'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
    'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
)
if ($Name -eq '.' -or $Name -eq '..') {
    throw "-Name '$Name' would resolve to a folder outside -Root. Pick a real project name."
}
if ($Name.EndsWith('.') -or $Name.EndsWith(' ')) {
    throw "-Name '$Name' ends with a dot or a space, which Windows will not keep as a folder name."
}
if ($reservedNames -contains ($Name -split '\.')[0].Trim().ToUpperInvariant()) {
    throw "-Name '$Name' is a reserved Windows device name."
}

$ProjectPath = Join-Path $Root $Name
$Created     = New-Object System.Collections.Generic.List[string]
$Skipped     = New-Object System.Collections.Generic.List[string]

function Test-Tool {
    param([string]$Tool)
    $null -ne (Get-Command $Tool -ErrorAction SilentlyContinue)
}

function Write-Action {
    param([string]$Message, [string]$Colour = '')
    $prefix = '  '
    if ($DryRun) { $prefix = '  [dry-run] ' }
    if ($Colour) { Write-Host ($prefix + $Message) -ForegroundColor $Colour }
    else         { Write-Host ($prefix + $Message) }
}

function New-ProjectDirectory {
    param([string]$Relative)
    $full = Join-Path $ProjectPath $Relative
    if (Test-Path -LiteralPath $full) { return }
    if (-not $DryRun) { New-Item -ItemType Directory -Path $full -Force | Out-Null }
    $Created.Add($Relative.TrimEnd('\') + '\')
}

function New-ProjectFile {
    <#
        Writes a file only if it does not already exist. Never overwrites -
        that is the whole safety guarantee of -Force.
    #>
    param([string]$Relative, [string]$Content)

    $full = Join-Path $ProjectPath $Relative
    if (Test-Path -LiteralPath $full) {
        $Skipped.Add($Relative)
        return
    }

    $parent = Split-Path -Parent $full
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        if (-not $DryRun) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    }

    if (-not $DryRun) {
        # No BOM: git, node and python all read these files and none of them
        # want one.
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($full, $Content, $utf8)
    }
    $Created.Add(($Relative -replace '/', '\'))
}

function Get-SlugName {
    param([string]$Raw)
    $s = ($Raw.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
    # A name made entirely of punctuation ('+++') would slug to nothing, and an
    # empty "name" field makes package.json and `dotnet new -n` both fail.
    if (-not $s) { $s = 'project' }
    $s
}

function Get-PythonPackageName {
    param([string]$Raw)
    $p = ($Raw.ToLowerInvariant() -replace '[^a-z0-9]+', '_').Trim('_')
    if (-not $p)         { $p = 'package' }
    if ($p -match '^\d') { $p = '_' + $p }
    $p
}

# ---- escaping ---------------------------------------------------------------
# The generated files are hand-assembled rather than run through ConvertTo-Json
# so they stay readable and diffable. That makes escaping this script's job: a
# -Description with a quote in it, or an -Author called O'Brien, otherwise
# produces a package.json npm cannot parse and a .psd1 PowerShell cannot load.

function ConvertTo-JsonText {
    # Also used for TOML basic strings, which escape the same way.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $s = $Text -replace '\\', '\\'
    $s = $s -replace '"', '\"'
    $s = $s -replace "`r", ''
    $s = $s -replace "`n", ' '
    $s = $s -replace "`t", ' '
    $s
}

function ConvertTo-PsLiteral {
    # Single-quoted PowerShell strings escape a quote by doubling it.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    (($Text -replace "`r", '') -replace "`n", ' ') -replace "'", "''"
}

function ConvertTo-JsLiteral {
    # For a single-quoted JavaScript string.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $s = $Text -replace '\\', '\\'
    $s = $s -replace "'", "\'"
    (($s -replace "`r", '') -replace "`n", ' ')
}

function ConvertTo-HtmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    [System.Net.WebUtility]::HtmlEncode($Text)
}

# ------------------------------------------------------------------ defaults
if (-not $Author) {
    if (Test-Tool 'git') {
        try {
            $gitUser = (& git config --global user.name) 2>$null
            if ($gitUser) { $Author = ([string]$gitUser).Trim() }
        }
        catch { }
    }
    if (-not $Author) { $Author = $env:USERNAME }
}

if (-not $Description) { $Description = "$Name" }

if (-not $OpenWith) {
    if (Test-Tool 'code') { $OpenWith = 'code' } else { $OpenWith = 'none' }
}

$Slug    = Get-SlugName $Name
$PyPkg   = Get-PythonPackageName $Name
$Year    = (Get-Date).Year
$Today   = (Get-Date).ToString('yyyy-MM-dd')

# Pre-escaped once, so each template below picks the form its own file format
# needs instead of interpolating raw user text into JSON, TOML, HTML or code.
$NameJson  = ConvertTo-JsonText  $Name
$NameJs    = ConvertTo-JsLiteral $Name
$NameHtml  = ConvertTo-HtmlText  $Name
$DescJson  = ConvertTo-JsonText  $Description
$DescHtml  = ConvertTo-HtmlText  $Description
$DescPs    = ConvertTo-PsLiteral $Description
$AuthJson  = ConvertTo-JsonText  $Author
$AuthPs    = ConvertTo-PsLiteral $Author

# --------------------------------------------------------------- target check
Write-Host ''
Write-Host "=== New $Type project: $Name ==="
Write-Host "  Location : $ProjectPath"
Write-Host ''

if (Test-Path -LiteralPath $ProjectPath) {
    $existing = @(Get-ChildItem -LiteralPath $ProjectPath -Force -ErrorAction SilentlyContinue)
    if ($existing.Count -gt 0 -and -not $Force) {
        Write-Host "  $ProjectPath already exists and is not empty." -ForegroundColor Red
        Write-Host '  Pass -Force to add the missing files to it. Existing files are never overwritten.' -ForegroundColor Red
        Write-Host ''
        exit 1
    }
}
else {
    if (-not $DryRun) {
        try { New-Item -ItemType Directory -Path $ProjectPath -Force -ErrorAction Stop | Out-Null }
        catch {
            Write-Host "  Could not create $ProjectPath" -ForegroundColor Red
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
            Write-Host '  Pick a -Root you can write to.' -ForegroundColor Red
            Write-Host ''
            exit 1
        }
    }
    $Created.Add('.\')
}

# ------------------------------------------------------------ common content
$gitignoreCommon = @'
# OS
Thumbs.db
ehthumbs.db
Desktop.ini
$RECYCLE.BIN/
.DS_Store

# Editors
.vscode/*
!.vscode/settings.json
!.vscode/extensions.json
.idea/
*.swp
*~

# Logs & scratch
*.log
.tmp/
tmp/
scratch/
'@

$gitignoreByType = @{
    'node' = @'

# Node
node_modules/
npm-debug.log*
yarn-debug.log*
yarn-error.log*
pnpm-debug.log*
.pnpm-store/
dist/
build/
coverage/
.env
.env.local
*.tsbuildinfo
'@
    'python' = @'

# Python
__pycache__/
*.py[cod]
*$py.class
.venv/
venv/
env/
.eggs/
*.egg-info/
dist/
build/
.pytest_cache/
.mypy_cache/
.ruff_cache/
.coverage
htmlcov/
.env
'@
    'powershell' = @'

# PowerShell
*.psd1.bak
TestResults/
output/
'@
    'dotnet' = @'

# .NET
bin/
obj/
*.user
*.suo
.vs/
TestResults/
'@
    'static' = @'

# Static site
dist/
node_modules/
.cache/
'@
    'blank' = ''
}

$editorConfig = @'
root = true

[*]
charset = utf-8
end_of_line = lf
insert_final_newline = true
trim_trailing_whitespace = true
indent_style = space
indent_size = 4

[*.{js,jsx,ts,tsx,json,yml,yaml,html,css,scss}]
indent_size = 2

[*.md]
trim_trailing_whitespace = false

[*.{ps1,psm1,psd1}]
indent_size = 4
end_of_line = crlf

[Makefile]
indent_style = tab
'@

$readme = @"
# $Name

$Description

## Status

Scaffolded $Today. Nothing works yet.

## Getting started

``````
# fill this in once there is something to run
``````

## Layout

| Path | What it is |
| ---- | ---------- |
|      |            |

## Notes

"@

$claudeMd = @"
# $Name

$Description

## What this project is

<!-- One paragraph. What problem does this solve, and for whom? -->

## Layout

<!-- Where the important code lives, so Claude does not have to guess. -->

## Conventions

<!-- Anything non-obvious: naming, error handling, testing style, formatting. -->

## Commands

``````
# build:
# test:
# run:
``````

## Gotchas

<!-- The things that would waste an hour if you did not know them. -->
"@

$mitLicense = @"
MIT License

Copyright (c) $Year $Author

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"@

# ------------------------------------------------------------- write commons
$ignoreExtra = ''
if ($gitignoreByType.ContainsKey($Type)) { $ignoreExtra = $gitignoreByType[$Type] }

New-ProjectFile 'README.md'      $readme
New-ProjectFile '.gitignore'     ($gitignoreCommon + $ignoreExtra + "`n")
New-ProjectFile '.editorconfig'  $editorConfig
New-ProjectFile 'CLAUDE.md'      $claudeMd

if ($License -eq 'MIT') { New-ProjectFile 'LICENSE' $mitLicense }

# -------------------------------------------------------- type-specific files
switch ($Type) {

    'node' {
        $licenseField = 'UNLICENSED'
        if ($License -eq 'MIT') { $licenseField = 'MIT' }

        $pkg = @"
{
  "name": "$Slug",
  "version": "0.1.0",
  "description": "$DescJson",
  "type": "module",
  "main": "src/index.js",
  "scripts": {
    "start": "node src/index.js",
    "test": "node --test"
  },
  "keywords": [],
  "author": "$AuthJson",
  "license": "$licenseField"
}
"@
        New-ProjectFile 'package.json' $pkg
        New-ProjectFile 'src/index.js' @"
export function main() {
  console.log('$NameJs');
}

main();
"@
        New-ProjectFile 'test/index.test.js' @"
import { test } from 'node:test';
import assert from 'node:assert/strict';

test('placeholder', () => {
  assert.ok(true);
});
"@
    }

    'python' {
        $licenseBlock = ''
        if ($License -eq 'MIT') { $licenseBlock = "license = { text = `"MIT`" }`n" }

        New-ProjectFile 'pyproject.toml' @"
[project]
name = "$Slug"
version = "0.1.0"
description = "$DescJson"
authors = [{ name = "$AuthJson" }]
requires-python = ">=3.10"
dependencies = []
$licenseBlock
[project.optional-dependencies]
dev = ["pytest", "ruff"]

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[tool.setuptools.packages.find]
where = ["src"]

[tool.ruff]
line-length = 100

[tool.pytest.ini_options]
testpaths = ["tests"]
"@
        # Python triple-quoted strings process backslash escapes, so the JSON
        # escaping is the right shape here too.
        New-ProjectFile "src/$PyPkg/__init__.py" @"
"""$DescJson"""

__version__ = "0.1.0"
"@
        New-ProjectFile "src/$PyPkg/__main__.py" @"
def main() -> None:
    print("$NameJson")


if __name__ == "__main__":
    main()
"@
        New-ProjectFile 'tests/test_placeholder.py' @"
def test_placeholder():
    assert True
"@
    }

    'powershell' {
        # A module name has to survive being a file name, a function suffix and
        # an identifier in the manifest, so anything else is stripped - and a
        # name that strips to nothing falls back rather than emitting 'Get-'.
        $moduleName = $Name -replace '[^A-Za-z0-9_.-]', ''
        if (-not $moduleName -or $moduleName -notmatch '^[A-Za-z_]') { $moduleName = 'Module' + $moduleName }
        New-ProjectDirectory 'Public'
        New-ProjectDirectory 'Private'
        New-ProjectDirectory 'Tests'

        New-ProjectFile "$moduleName.psm1" @"
# Dot-source every function, then export only what lives in Public\.
`$public  = @(Get-ChildItem -Path `$PSScriptRoot\Public\*.ps1  -ErrorAction SilentlyContinue)
`$private = @(Get-ChildItem -Path `$PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

foreach (`$file in @(`$public + `$private)) {
    try { . `$file.FullName }
    catch { Write-Error "Failed to import `$(`$file.FullName): `$_" }
}

Export-ModuleMember -Function `$public.BaseName
"@

        New-ProjectFile "Public/Get-$moduleName.ps1" @"
function Get-$moduleName {
    <#
    .SYNOPSIS
        Placeholder.
    #>
    [CmdletBinding()]
    param()

    '$moduleName'
}
"@

        New-ProjectFile "Tests/$moduleName.Tests.ps1" @"
BeforeAll {
    Import-Module (Join-Path `$PSScriptRoot '..' '$moduleName.psm1') -Force
}

Describe 'Get-$moduleName' {
    It 'returns the module name' {
        Get-$moduleName | Should -Be '$moduleName'
    }
}
"@

        # A hand-written manifest beats New-ModuleManifest here: it stays
        # readable and diffable instead of being 90% commented-out defaults.
        New-ProjectFile "$moduleName.psd1" @"
@{
    RootModule        = '$moduleName.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '$([guid]::NewGuid().ToString())'
    Author            = '$AuthPs'
    Description       = '$DescPs'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-$moduleName')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData = @{
        PSData = @{
            Tags = @()
        }
    }
}
"@
    }

    'static' {
        New-ProjectFile 'index.html' @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$NameHtml</title>
<link rel="stylesheet" href="styles.css">
</head>
<body>
  <main>
    <h1>$NameHtml</h1>
    <p>$DescHtml</p>
  </main>
  <script src="script.js"></script>
</body>
</html>
"@
        New-ProjectFile 'styles.css' @"
:root {
  --bg: #ffffff;
  --ink: #1b1f24;
  --accent: #2f6fed;
}
@media (prefers-color-scheme: dark) {
  :root { --bg: #14171c; --ink: #e6e9ee; --accent: #79a5ff; }
}
* { box-sizing: border-box; }
body {
  margin: 0;
  padding: 48px 24px;
  background: var(--bg);
  color: var(--ink);
  font: 16px/1.6 system-ui, -apple-system, "Segoe UI", sans-serif;
}
main { max-width: 720px; margin: 0 auto; }
h1 { letter-spacing: -0.02em; }
a { color: var(--accent); }
"@
        New-ProjectFile 'script.js' @"
console.log('$NameJs');
"@
    }

    'dotnet' {
        if (Test-Tool 'dotnet') {
            if ($DryRun) {
                Write-Action "would run: dotnet new console -o `"$ProjectPath`" -n $Slug"
            }
            else {
                Write-Action 'running dotnet new console...'
                # No --force. This script promises never to overwrite an
                # existing file, and --force is precisely the flag that lets
                # `dotnet new` do it - so an existing Program.cs or .csproj is
                # reported and left alone, exactly like every other file here.
                & dotnet new console -o $ProjectPath -n $Slug | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    $Created.Add('(dotnet new console output)')
                }
                else {
                    Write-Action "dotnet new console failed (exit $LASTEXITCODE) - nothing was overwritten" 'Yellow'
                    Write-Action 'usually means this folder already holds a project; run dotnet new by hand if you want it replaced' 'DarkGray'
                }
            }
        }
        else {
            Write-Action 'dotnet SDK not found - skipped project generation' 'Yellow'
        }
    }
}

# ---------------------------------------------------------------- python venv
if ($Type -eq 'python' -and $CreateVenv) {
    $py = $null
    if     (Test-Tool 'py')     { $py = 'py' }
    elseif (Test-Tool 'python') { $py = 'python' }

    if (-not $py) {
        Write-Action 'python not found - skipped venv creation' 'Yellow'
    }
    elseif ($DryRun) {
        Write-Action "would create a virtual environment at $ProjectPath\.venv"
    }
    else {
        Write-Action 'creating virtual environment (.venv)...'
        try {
            # A native command that fails does not throw, so the exit code is
            # the only thing that distinguishes "made a venv" from "printed an
            # error and made nothing".
            & $py -m venv (Join-Path $ProjectPath '.venv')
            if ($LASTEXITCODE -eq 0) { $Created.Add('.venv\') }
            else { Write-Action "venv creation failed (exit $LASTEXITCODE)" 'Yellow' }
        }
        catch {
            Write-Action "venv creation failed: $($_.Exception.Message)" 'Yellow'
        }
    }
}

# ----------------------------------------------------------------------- git
if (-not $NoGit) {
    if (-not (Test-Tool 'git')) {
        Write-Action 'git not found on PATH - skipped repo initialisation' 'Yellow'
    }
    elseif (Test-Path -LiteralPath (Join-Path $ProjectPath '.git')) {
        Write-Action 'git repo already exists here - left alone'
    }
    elseif ($DryRun) {
        Write-Action 'would run: git init, git add -A, git commit'
    }
    else {
        Push-Location $ProjectPath
        try {
            # `git init -b main` on git < 2.28 prints to stderr and returns a
            # non-zero exit code - it does not throw, so catch{} never fires and
            # the old code reported success on a repo that was never created.
            # $LASTEXITCODE is the only reliable signal here.
            $initOk = $false
            & git init -b main 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) {
                $initOk = $true
            }
            else {
                & git init | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    & git symbolic-ref HEAD refs/heads/main | Out-Null
                    $initOk = ($LASTEXITCODE -eq 0)
                }
            }

            if ($initOk) {
                Write-Action 'git repo initialised (branch: main)'
                if (-not $NoCommit) {
                    & git add -A | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Write-Action "git add failed (exit $LASTEXITCODE) - nothing committed" 'Yellow'
                    }
                    else {
                        & git commit -m 'Initial scaffold' --quiet
                        if ($LASTEXITCODE -eq 0) { Write-Action 'initial commit created' }
                        else { Write-Action 'nothing committed - check git user.name / user.email' 'Yellow' }
                    }
                }
            }
            else {
                Write-Action 'git init failed - the folder was still scaffolded' 'Yellow'
            }
        }
        finally {
            Pop-Location
        }
    }
}

# ------------------------------------------------------------------- summary
Write-Host ''
Write-Host "  Created ($($Created.Count)):"
foreach ($c in ($Created | Sort-Object)) { Write-Host "    + $c" -ForegroundColor Green }

if ($Skipped.Count -gt 0) {
    Write-Host ''
    Write-Host "  Already existed, left untouched ($($Skipped.Count)):"
    foreach ($s in $Skipped) { Write-Host "    = $s" -ForegroundColor DarkGray }
}

Write-Host ''
Write-Host "  Project: $ProjectPath"

if ($DryRun) {
    Write-Host ''
    Write-Host '  DRY RUN - nothing was created. Re-run without -DryRun to build it.' -ForegroundColor Cyan
    exit 0
}

# ---------------------------------------------------------------------- open
switch ($OpenWith) {
    'code' {
        if (Test-Tool 'code') { & code $ProjectPath }
        else { Write-Action 'VS Code not on PATH - not opening' 'Yellow' }
    }
    'claude' {
        if (Test-Tool 'claude') {
            Write-Host ''
            Write-Host '  Starting Claude Code in the new project...'
            Start-Process -FilePath 'cmd.exe' -ArgumentList '/s', '/k', "pushd `"$ProjectPath`" && claude"
        }
        else { Write-Action 'claude not on PATH - not opening' 'Yellow' }
    }
    'explorer' {
        # -ArgumentList, quoted: a project path under "Documents\my projects"
        # would otherwise reach Explorer as two separate arguments.
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $ProjectPath)
    }
    default    { }
}

Write-Host ''
exit 0
