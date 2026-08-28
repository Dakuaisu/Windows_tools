<#
.SYNOPSIS
    A single window for every tool in this folder: pick one on the left, fill
    in the form on the right, watch it run.

.DESCRIPTION
    A launcher, and deliberately nothing more. It knows how to find scripts and
    how to read their parameters; it knows nothing about what any individual
    tool does.

    That is the whole design. Every control on the right-hand side is generated
    from the target script's own param() block and comment-based help, read
    through the PowerShell parser. Add a parameter to a tool tomorrow and its
    form grows a control on the next launch, with no change here. Nothing is
    hardcoded per tool, and no tool has to know this window exists.

    The dependency points one way only, which is what keeps the promise the
    README makes: every script in this folder still runs on its own, copied to
    a machine that has never seen this file.

    Widgets are chosen from the parameter's own declaration:

      [switch]                     a checkbox
      [ValidateSet]                a dropdown, or a multi-select list for arrays
      [int] / [double]             a number box, with any [ValidateRange] shown
      [string[]]                   a box taking one value per line
      anything named *Path/*Dir    a text box with a Browse button

    Scripts with more than one parameter set - Rename-Bulk, Set-FocusMode - get
    a Mode selector, and only the parameters valid for that mode are shown.

    Tools run as a separate PowerShell process with their output streamed into
    the pane at the bottom, so the window stays responsive and what you see is
    exactly what the console would have shown. The .cmd tools in
    DownloadsJanitor open their own window instead, because they are interactive
    and ask questions.

.PARAMETER ToolRoot
    Folder to scan for tools. Defaults to the folder holding this script.

.PARAMETER Scan
    Subfolders to look in. Defaults to the four this repository uses.

.EXAMPLE
    .\Toolbox.ps1

.EXAMPLE
    # Point it at a copy of the tools kept somewhere else
    .\Toolbox.ps1 -ToolRoot D:\tools
#>
[CmdletBinding()]
param(
    [string]$ToolRoot = $PSScriptRoot,
    [string[]]$Scan   = @('diagnostics', 'backup', 'productivity', 'DownloadsJanitor')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ------------------------------------------------------------------- plumbing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

if (-not $ToolRoot) { $ToolRoot = (Get-Location).Path }
if (-not (Test-Path -LiteralPath $ToolRoot)) { throw "Tool folder not found: $ToolRoot" }

$PsExe = Join-Path $PSHOME 'powershell.exe'

function Get-Prop {
    # Strict-mode-safe property read.
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    $p.Value
}

function Get-FriendlyName {
    # 'Get-SystemSnapshot' -> 'Get System Snapshot'
    param([string]$BaseName)
    $s = $BaseName -replace '-', ' '
    ($s -creplace '(?<=[a-z0-9])(?=[A-Z])', ' ')
}

# --------------------------------------------------------------- 1. discovery
function ConvertFrom-DefaultValue {
    <#
        The parser hands back the default as source text, not a value:
        "(Join-Path $(if ($env:LOCALAPPDATA) {...}) 'BatteryReport')". Showing
        that in a text box would be useless, so it is evaluated to get the same
        string the script itself would compute. These are the repository's own
        scripts and are about to be executed anyway, so evaluating their
        defaults adds no exposure that running them does not already carry -
        but anything that fails falls back to the raw text rather than throwing.
    #>
    param($DefaultAst)

    if ($null -eq $DefaultAst) { return '' }
    $text = $DefaultAst.Extent.Text
    try {
        $value = Invoke-Expression $text
        if ($null -eq $value)          { return '' }
        if ($value -is [bool])         { return $value }
        if ($value -is [array])        { return (($value | ForEach-Object { [string]$_ }) -join [Environment]::NewLine) }
        return [string]$value
    }
    catch { return $text }
}

function Get-ToolDefinition {
    <#
        Everything the form builder needs, read out of the script itself.
        Returns $null for anything that does not parse, so a broken file in the
        folder cannot take the window down with it.
    #>
    param([System.IO.FileInfo]$File)

    $errs = $null
    $toks = $null
    $ast  = $null
    try { $ast = [System.Management.Automation.Language.Parser]::ParseFile($File.FullName, [ref]$toks, [ref]$errs) }
    catch { return $null }
    if ($null -eq $ast -or $errs.Count -gt 0) { return $null }

    $help     = $ast.GetHelpContent()
    $synopsis = ''
    $helpParams = @{}
    if ($help) {
        if ($help.Synopsis) {
            $synopsis = (($help.Synopsis -split "`r?`n" | Where-Object { $_.Trim() }) -join ' ').Trim()
        }
        if ($help.Parameters) { $helpParams = $help.Parameters }
    }

    # DefaultParameterSetName off the [CmdletBinding()] attribute, so the Mode
    # selector opens on the same set the script itself would have picked.
    $defaultSet = ''
    if ($ast.ParamBlock) {
        foreach ($attr in $ast.ParamBlock.Attributes) {
            if ((Get-Prop $attr.TypeName 'Name' '') -ne 'CmdletBinding') { continue }
            foreach ($na in $attr.NamedArguments) {
                if ($na.ArgumentName -eq 'DefaultParameterSetName' -and
                    $na.Argument -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $defaultSet = $na.Argument.Value
                }
            }
        }
    }

    $params = New-Object System.Collections.Generic.List[psobject]
    $sets   = New-Object System.Collections.Generic.List[string]

    if ($ast.ParamBlock) {
        foreach ($p in $ast.ParamBlock.Parameters) {
            $name = $p.Name.VariablePath.UserPath
            $type = 'String'
            if ($p.StaticType) { $type = $p.StaticType.Name }

            $validSet  = New-Object System.Collections.Generic.List[string]
            $rangeText = ''
            $mySets    = New-Object System.Collections.Generic.List[string]
            $mandatory = $false

            foreach ($a in $p.Attributes) {
                $tn = [string](Get-Prop (Get-Prop $a 'TypeName') 'Name' '')

                if ($tn -eq 'ValidateSet') {
                    foreach ($arg in $a.PositionalArguments) {
                        if ($arg -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                            $validSet.Add($arg.Value)
                        }
                    }
                }
                elseif ($tn -eq 'ValidateRange') {
                    $rangeText = (@($a.PositionalArguments | ForEach-Object { $_.Extent.Text }) -join ' to ')
                }
                elseif ($tn -eq 'Parameter') {
                    foreach ($na in $a.NamedArguments) {
                        if ($na.ArgumentName -eq 'ParameterSetName' -and
                            $na.Argument -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                            $mySets.Add($na.Argument.Value)
                            if (-not $sets.Contains($na.Argument.Value)) { $sets.Add($na.Argument.Value) }
                        }
                        if ($na.ArgumentName -eq 'Mandatory' -and $na.Argument.Extent.Text -match '\$true') {
                            $mandatory = $true
                        }
                    }
                }
            }

            $isArray  = ($type -like '*[[]]*') -or ($p.StaticType -and $p.StaticType.IsArray)
            $isSwitch = ($type -eq 'SwitchParameter')

            # The widget follows from the declaration; nothing here is per-tool.
            $widget = 'Text'
            if     ($isSwitch)                          { $widget = 'Check' }
            elseif ($validSet.Count -gt 0 -and $isArray) { $widget = 'MultiList' }
            elseif ($validSet.Count -gt 0)               { $widget = 'Combo' }
            elseif ($isArray)                            { $widget = 'Lines' }
            elseif ($type -match 'Int|Double|Single|Decimal') { $widget = 'Number' }
            if (-not $isSwitch -and -not $isArray -and $name -match 'Path$|Dir$|Root$|File$') { $widget = 'Folder' }

            $desc = ''
            $key  = $name.ToUpperInvariant()
            if ($helpParams -and $helpParams.ContainsKey($key)) {
                $desc = (($helpParams[$key] -split "`r?`n" | Where-Object { $_.Trim() }) -join ' ').Trim()
            }

            $params.Add([pscustomobject]@{
                Name      = $name
                Type      = $type
                Widget    = $widget
                Choices   = $validSet.ToArray()
                Range     = $rangeText
                Default   = (ConvertFrom-DefaultValue $p.DefaultValue)
                Help      = $desc
                Sets      = $mySets.ToArray()
                Mandatory = $mandatory
                IsSwitch  = $isSwitch
                IsArray   = $isArray
            })
        }
    }

    [pscustomobject]@{
        Name        = $File.BaseName
        Display     = (Get-FriendlyName $File.BaseName)
        Path        = $File.FullName
        Folder      = Split-Path -Leaf (Split-Path -Parent $File.FullName)
        Synopsis    = $synopsis
        Params      = $params
        Sets        = $sets.ToArray()
        DefaultSet  = $defaultSet
        Kind        = 'Script'
    }
}

function Get-CmdDefinition {
    <#
        The DownloadsJanitor launchers. They take no parameters and they ask
        the user questions, so they get their own console window rather than
        being piped into the output pane where nobody could answer them.
    #>
    param([System.IO.FileInfo]$File)

    [pscustomobject]@{
        Name       = $File.BaseName
        Display    = $File.BaseName
        Path       = $File.FullName
        Folder     = Split-Path -Leaf (Split-Path -Parent $File.FullName)
        Synopsis   = 'Opens in its own window - it asks before it changes anything.'
        Params     = (New-Object System.Collections.Generic.List[psobject])
        Sets       = @()
        DefaultSet = ''
        Kind       = 'Cmd'
    }
}

$Tools = New-Object System.Collections.Generic.List[psobject]
foreach ($sub in $Scan) {
    $dir = Join-Path $ToolRoot $sub
    if (-not (Test-Path -LiteralPath $dir)) { continue }

    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter *.ps1 -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $def = Get-ToolDefinition -File $f
        if ($def) { $Tools.Add($def) }
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter *.cmd -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $Tools.Add((Get-CmdDefinition -File $f))
    }
}

if ($Tools.Count -eq 0) {
    throw "No tools found under $ToolRoot. Expected subfolders: $($Scan -join ', ')"
}

# ------------------------------------------------------------ 2. the argument
function ConvertTo-PsArgument {
    <#
        Builds the command line for one parameter. Everything is quoted for
        PowerShell rather than for cmd, because the tool is invoked through
        -Command; a folder called "My Stuff's Backup" has to survive this.
    #>
    param($Param, $Value)

    $name = $Param.Name

    if ($Param.IsSwitch) {
        # A switch that defaults to $true has to be turned off explicitly,
        # which is why this cannot just omit it.
        $defaultOn = ($Param.Default -is [bool] -and $Param.Default)
        if ($Value) { if (-not $defaultOn) { return "-$name" } else { return '' } }
        if ($defaultOn) { return "-${name}:`$false" }
        return ''
    }

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    if ($Param.IsArray) {
        $items = @($text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($items.Count -eq 0) { return '' }
        $quoted = ($items | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ','
        return "-$name $quoted"
    }

    if ($Param.Widget -eq 'Number') {
        return "-$name $text"
    }

    "-$name '" + ($text -replace "'", "''") + "'"
}

# ------------------------------------------------------------------ 3. the UI
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Tools" Height="760" Width="1140" WindowStartupLocation="CenterScreen"
        Background="#F6F7F9">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="14,6"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="MinWidth" Value="96"/>
    </Style>
    <Style TargetType="TextBlock"><Setter Property="Foreground" Value="#1B1F24"/></Style>
  </Window.Resources>
  <Grid Margin="12">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="270"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <Border Grid.Column="0" Background="White" BorderBrush="#E2E6EA" BorderThickness="1" CornerRadius="8" Margin="0,0,10,0">
      <DockPanel>
        <TextBlock DockPanel.Dock="Top" Text="Tools" FontSize="15" FontWeight="SemiBold" Margin="14,12,14,8"/>
        <ListBox x:Name="ToolList" BorderThickness="0" Margin="6,0,6,8" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
          <ListBox.ItemTemplate>
            <DataTemplate>
              <StackPanel Margin="6,5">
                <TextBlock Text="{Binding Display}" FontWeight="SemiBold"/>
                <TextBlock Text="{Binding Folder}" FontSize="11" Foreground="#5B6572"/>
              </StackPanel>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </DockPanel>
    </Border>

    <Grid Grid.Column="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="230"/>
      </Grid.RowDefinitions>

      <Border Grid.Row="0" Background="White" BorderBrush="#E2E6EA" BorderThickness="1" CornerRadius="8" Padding="16,12">
        <StackPanel>
          <TextBlock x:Name="TitleText" FontSize="18" FontWeight="SemiBold"/>
          <TextBlock x:Name="SynopsisText" TextWrapping="Wrap" Foreground="#5B6572" Margin="0,4,0,0"/>
        </StackPanel>
      </Border>

      <Border Grid.Row="1" Background="White" BorderBrush="#E2E6EA" BorderThickness="1" CornerRadius="8" Margin="0,10,0,0">
        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="16,12">
          <StackPanel x:Name="FormPanel"/>
        </ScrollViewer>
      </Border>

      <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,10,0,0">
        <Button x:Name="RunButton" Content="Run" FontWeight="SemiBold"/>
        <Button x:Name="StopButton" Content="Stop" IsEnabled="False"/>
        <Button x:Name="ReportButton" Content="Open last report" MinWidth="130"/>
        <Button x:Name="HelpButton" Content="Full help"/>
        <CheckBox x:Name="ElevateBox" Content="Run as administrator" VerticalAlignment="Center" Margin="10,0,0,0"/>
        <TextBlock x:Name="StatusText" VerticalAlignment="Center" Margin="16,0,0,0" Foreground="#5B6572"/>
      </StackPanel>

      <Border Grid.Row="3" Background="#14171C" BorderBrush="#E2E6EA" BorderThickness="1" CornerRadius="8" Margin="0,10,0,0">
        <DockPanel>
          <TextBlock DockPanel.Dock="Top" Text="Output" Foreground="#9AA4B2" FontSize="11" Margin="12,8,12,4"/>
          <ScrollViewer x:Name="OutScroll" VerticalScrollBarVisibility="Auto" Margin="0,0,0,8">
            <TextBox x:Name="OutputBox" Background="Transparent" Foreground="#E6E9EE" BorderThickness="0"
                     FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="NoWrap"
                     Margin="12,0" VerticalAlignment="Stretch"/>
          </ScrollViewer>
        </DockPanel>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

$ToolList     = $Window.FindName('ToolList')
$TitleText    = $Window.FindName('TitleText')
$SynopsisText = $Window.FindName('SynopsisText')
$FormPanel    = $Window.FindName('FormPanel')
$RunButton    = $Window.FindName('RunButton')
$StopButton   = $Window.FindName('StopButton')
$ReportButton = $Window.FindName('ReportButton')
$HelpButton   = $Window.FindName('HelpButton')
$ElevateBox   = $Window.FindName('ElevateBox')
$StatusText   = $Window.FindName('StatusText')
$OutputBox    = $Window.FindName('OutputBox')
$OutScroll    = $Window.FindName('OutScroll')

# Shown in a dropdown whose declared default is not one of its own choices,
# meaning the script computes the value itself and should be left to do so.
$Script:DefaultChoiceLabel = '(script default)'

# Mutable state shared with the event handlers. A hashtable rather than plain
# variables because a scriptblock handler cannot assign to an enclosing scope.
#
# Controls and Rows are only ever reached through ContainsKey() and the
# indexer, never through dot notation: a tool with a parameter called Count or
# Keys - Test-InternetHealth has -Count - would otherwise shadow the
# hashtable's own members and return a text box where a number was expected.
$State = @{
    Controls   = @{}        # parameter name -> its control
    Rows       = @{}        # parameter name -> the row holding it, for Mode filtering
    ModeCombo  = $null
    Tool       = $null
    Process    = $null
    OutQueue   = $null
    ErrQueue   = $null
    Timer      = $null
    Handlers   = @()
}

function Write-Output-Line {
    param([string]$Text)
    $OutputBox.AppendText($Text + [Environment]::NewLine)
    $OutScroll.ScrollToEnd()
}

# --------------------------------------------------------- 4. form generation
function Add-FormRow {
    param([string]$Label, $Control, [string]$Hint, [string[]]$Sets)

    $row = New-Object System.Windows.Controls.StackPanel
    $row.Margin = '0,0,0,14'

    $lab = New-Object System.Windows.Controls.TextBlock
    $lab.Text       = $Label
    $lab.FontWeight = 'SemiBold'
    $lab.Margin     = '0,0,0,3'
    [void]$row.Children.Add($lab)

    if ($Hint) {
        $h = New-Object System.Windows.Controls.TextBlock
        $h.Text         = $Hint
        $h.FontSize     = 11
        $h.Foreground   = '#5B6572'
        $h.TextWrapping = 'Wrap'
        $h.Margin       = '0,0,0,4'
        [void]$row.Children.Add($h)
    }

    [void]$row.Children.Add($Control)
    [void]$FormPanel.Children.Add($row)
    $row
}

function New-ParameterControl {
    param($Param)

    switch ($Param.Widget) {

        'Check' {
            $c = New-Object System.Windows.Controls.CheckBox
            $c.Content = 'enabled'
            if ($Param.Default -is [bool]) { $c.IsChecked = $Param.Default }
            return $c
        }

        'Combo' {
            $c = New-Object System.Windows.Controls.ComboBox

            # A declared default outside the ValidateSet means the script works
            # the value out at runtime - New-ProjectScaffold's -OpenWith
            # defaults to '' and then picks 'code' only if VS Code is on PATH.
            # Selecting the first choice here would quietly make that decision
            # for it, so the placeholder is offered instead and the parameter is
            # left off the command line entirely.
            $declared = [string]$Param.Default
            if ($declared -and ($Param.Choices -notcontains $declared)) { $declared = '' }
            if (-not $declared) { [void]$c.Items.Add($Script:DefaultChoiceLabel) }

            foreach ($choice in $Param.Choices) { [void]$c.Items.Add($choice) }

            if ($declared) { $c.SelectedItem = $declared } else { $c.SelectedIndex = 0 }
            return $c
        }

        'MultiList' {
            $c = New-Object System.Windows.Controls.ListBox
            $c.SelectionMode = 'Multiple'
            $c.MaxHeight     = 132
            $defaults = @([string]$Param.Default -split "`r?`n" | ForEach-Object { $_.Trim() })
            foreach ($choice in $Param.Choices) {
                [void]$c.Items.Add($choice)
                if ($defaults -contains $choice) { [void]$c.SelectedItems.Add($choice) }
            }
            return $c
        }

        'Lines' {
            $c = New-Object System.Windows.Controls.TextBox
            $c.AcceptsReturn = $true
            $c.MinHeight     = 64
            $c.MaxHeight     = 120
            $c.VerticalScrollBarVisibility = 'Auto'
            $c.Text = [string]$Param.Default
            return $c
        }

        'Folder' {
            $g = New-Object System.Windows.Controls.DockPanel
            $b = New-Object System.Windows.Controls.Button
            $b.Content = 'Browse'
            $b.MinWidth = 74
            $b.Margin  = '6,0,0,0'
            [System.Windows.Controls.DockPanel]::SetDock($b, 'Right')
            $t = New-Object System.Windows.Controls.TextBox
            $t.Text = [string]$Param.Default
            [void]$g.Children.Add($b)
            [void]$g.Children.Add($t)

            $b.Add_Click({
                $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
                if ($t.Text -and (Test-Path -LiteralPath $t.Text)) { $dlg.SelectedPath = $t.Text }
                if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $t.Text = $dlg.SelectedPath }
            }.GetNewClosure())

            # The text box is what the reader reads, not the wrapper.
            $g | Add-Member -NotePropertyName 'ValueBox' -NotePropertyValue $t -Force
            return $g
        }

        default {
            $c = New-Object System.Windows.Controls.TextBox
            $c.Text = [string]$Param.Default
            return $c
        }
    }
}

function Read-ParameterControl {
    param($Param, $Control)

    switch ($Param.Widget) {
        'Check'     { return [bool]$Control.IsChecked }
        'Combo'     {
            $v = [string]$Control.SelectedItem
            # The placeholder means "say nothing and let the script decide".
            if ($v -eq $Script:DefaultChoiceLabel) { return '' }
            return $v
        }
        'MultiList' { return ((@($Control.SelectedItems) | ForEach-Object { [string]$_ }) -join [Environment]::NewLine) }
        'Folder'    { return [string]$Control.ValueBox.Text }
        default     { return [string]$Control.Text }
    }
}

function Update-ModeVisibility {
    <#
        With a Mode selector present, a parameter is shown only if it belongs to
        the chosen set or to no set at all. Without this, Rename-Bulk would
        offer -Pattern and -Template side by side and PowerShell would reject
        whatever you filled in.
    #>
    if ($null -eq $State.ModeCombo) { return }
    $mode = [string]$State.ModeCombo.SelectedItem

    foreach ($p in $State.Tool.Params) {
        if (-not $State.Rows.ContainsKey($p.Name)) { continue }
        $row = $State.Rows[$p.Name]
        $visible = ($p.Sets.Count -eq 0) -or ($p.Sets -contains $mode)
        if ($visible) { $row.Visibility = 'Visible' } else { $row.Visibility = 'Collapsed' }
    }
}

function Show-Tool {
    param($Tool)

    $State.Tool      = $Tool
    $State.Controls  = @{}
    $State.Rows      = @{}
    $State.ModeCombo = $null
    $FormPanel.Children.Clear()

    $TitleText.Text    = $Tool.Name
    $SynopsisText.Text = $Tool.Synopsis

    $ReportButton.IsEnabled = ($Tool.Kind -eq 'Script')
    $HelpButton.IsEnabled   = ($Tool.Kind -eq 'Script')

    if ($Tool.Kind -eq 'Cmd') {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text         = 'This one takes no options. Run opens it in its own window.'
        $t.Foreground   = '#5B6572'
        $t.TextWrapping = 'Wrap'
        [void]$FormPanel.Children.Add($t)
        return
    }

    if ($Tool.Sets.Count -gt 1) {
        $combo = New-Object System.Windows.Controls.ComboBox
        foreach ($s in $Tool.Sets) { [void]$combo.Items.Add($s) }
        if ($Tool.DefaultSet -and $combo.Items.Contains($Tool.DefaultSet)) { $combo.SelectedItem = $Tool.DefaultSet }
        else { $combo.SelectedIndex = 0 }
        $combo.Add_SelectionChanged({ Update-ModeVisibility })
        $State.ModeCombo = $combo
        [void](Add-FormRow -Label 'Mode' -Control $combo -Hint 'This script has more than one way to be called.')
    }

    foreach ($p in $Tool.Params) {
        $control = New-ParameterControl -Param $p

        $hint = $p.Help
        if ($p.Range)     { $hint = ($hint + "  (" + $p.Range + ")").Trim() }
        if ($p.Mandatory) { $hint = ('Required. ' + $hint).Trim() }

        $label = $p.Name
        if ($p.Mandatory) { $label = $p.Name + ' *' }

        $row = Add-FormRow -Label $label -Control $control -Hint $hint -Sets $p.Sets
        $State.Controls[$p.Name] = $control
        $State.Rows[$p.Name]     = $row
    }

    Update-ModeVisibility
}

# ------------------------------------------------------------- 5. the running
function Stop-RunningTool {
    if ($State.Process) {
        try { if (-not $State.Process.HasExited) { $State.Process.Kill() } } catch { }
    }
}

function Complete-Run {
    param([string]$Message)

    if ($State.Timer) { $State.Timer.Stop(); $State.Timer = $null }
    foreach ($h in $State.Handlers) { try { Unregister-Event -SubscriptionId $h.Id -ErrorAction SilentlyContinue } catch { } }
    $State.Handlers = @()
    if ($State.Process) { try { $State.Process.Dispose() } catch { } }
    $State.Process = $null

    $RunButton.IsEnabled  = $true
    $StopButton.IsEnabled = $false
    $StatusText.Text      = $Message
}

function Get-ToolCommandLine {
    <#
        Turns the current state of the form into the command PowerShell will be
        asked to run. Kept out of the click handler so it can be exercised
        without a window: a launcher whose only testable path is "click it and
        see" is a launcher nobody can trust.

        Returns an object with either Command or Error set.
    #>
    $tool = $State.Tool
    if ($null -eq $tool) { return [pscustomobject]@{ Command = ''; Error = 'No tool selected.' } }

    $argParts = New-Object System.Collections.Generic.List[string]
    $mode = ''
    if ($State.ModeCombo) { $mode = [string]$State.ModeCombo.SelectedItem }

    foreach ($p in $tool.Params) {
        # Only parameters valid in the chosen mode, or PowerShell rejects the call.
        if ($mode -and $p.Sets.Count -gt 0 -and ($p.Sets -notcontains $mode)) { continue }
        if (-not $State.Controls.ContainsKey($p.Name)) { continue }

        $value = Read-ParameterControl -Param $p -Control $State.Controls[$p.Name]

        if ($p.Mandatory -and -not $p.IsSwitch -and [string]::IsNullOrWhiteSpace([string]$value)) {
            return [pscustomobject]@{ Command = ''; Error = ($p.Name + ' is required.') }
        }

        # Unchanged values are left off the command line entirely, so the script
        # applies its own default rather than being handed a copy of it. That
        # matters for the paths: the default shown in the box is the resolved
        # one, and echoing it back would pin a value the script means to compute.
        $isDefault = $false
        if ($p.IsSwitch) { $isDefault = ($value -eq ($p.Default -is [bool] -and $p.Default)) }
        else             { $isDefault = ([string]$value -eq [string]$p.Default) }
        if ($isDefault) { continue }

        $piece = ConvertTo-PsArgument -Param $p -Value $value
        if ($piece) { $argParts.Add($piece) }
    }

    $command = "& '" + ($tool.Path -replace "'", "''") + "'"
    if ($argParts.Count -gt 0) { $command += ' ' + ($argParts -join ' ') }
    [pscustomobject]@{ Command = $command; Error = '' }
}

function Invoke-Tool {
    if ($null -eq $State.Tool) { return }
    $tool = $State.Tool

    # --- the .cmd launchers get their own console; they ask questions ---
    if ($tool.Kind -eq 'Cmd') {
        try {
            Start-Process -FilePath $tool.Path -WorkingDirectory (Split-Path -Parent $tool.Path)
            $StatusText.Text = 'Opened in its own window.'
        }
        catch { Write-Output-Line ("Could not start: " + $_.Exception.Message) }
        return
    }

    $built = Get-ToolCommandLine
    if ($built.Error) {
        Write-Output-Line $built.Error
        $StatusText.Text = $built.Error
        return
    }
    $command = $built.Command

    $OutputBox.Clear()
    Write-Output-Line ('> ' + $command)
    Write-Output-Line ''

    # --- elevated runs cannot be captured, so they get a real console ---
    if ($ElevateBox.IsChecked) {
        try {
            Start-Process -FilePath $PsExe -Verb RunAs -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-Command', $command)
            $StatusText.Text = 'Running elevated in its own window.'
            Write-Output-Line 'Started in an elevated window - output cannot be captured across the UAC boundary.'
        }
        catch { Write-Output-Line ('Elevation refused or failed: ' + $_.Exception.Message) }
        return
    }

    # --- normal run: separate process, output streamed into the pane ---
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $PsExe
    $psi.Arguments              = '-NoProfile -ExecutionPolicy Bypass -Command "' + ($command -replace '"', '\"') + '"'
    $psi.UseShellExecute        = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.CreateNoWindow         = $true
    $psi.WorkingDirectory       = $ToolRoot

    # One queue per stream, not one shared between both. Two -Action handlers
    # feeding a single queue run concurrently and genuinely interleave: a test
    # emitting 300 numbered lines came back with nine of them out of sequence,
    # which in a report means a heading landing below its own table. Each queue
    # now has exactly one writer, which keeps it ordered, and stderr is drained
    # after stdout on each tick.
    $outQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $errQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $State.OutQueue = $outQueue
    $State.ErrQueue = $errQueue

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo           = $psi
    $proc.EnableRaisingEvents = $true

    # The stream events fire on a worker thread, which must not touch a WPF
    # control. They only enqueue; a timer on the UI thread does the drawing.
    $sink = { if ($null -ne $EventArgs.Data) { $Event.MessageData.Enqueue($EventArgs.Data) } }
    $State.Handlers = @(
        (Register-ObjectEvent -InputObject $proc -EventName OutputDataReceived -Action $sink -MessageData $outQueue)
        (Register-ObjectEvent -InputObject $proc -EventName ErrorDataReceived  -Action $sink -MessageData $errQueue)
    )

    try {
        [void]$proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()
    }
    catch {
        Write-Output-Line ('Could not start PowerShell: ' + $_.Exception.Message)
        Complete-Run 'Failed to start.'
        return
    }

    $State.Process        = $proc
    $RunButton.IsEnabled  = $false
    $StopButton.IsEnabled = $true
    $StatusText.Text      = 'Running...'

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(120)
    $timer.Add_Tick({
        $line = ''
        while ($State.OutQueue.TryDequeue([ref]$line)) { Write-Output-Line $line }
        while ($State.ErrQueue.TryDequeue([ref]$line)) { Write-Output-Line $line }

        if ($State.Process -and $State.Process.HasExited) {
            # One last drain: lines can still be in flight when the process ends.
            Start-Sleep -Milliseconds 120
            while ($State.OutQueue.TryDequeue([ref]$line)) { Write-Output-Line $line }
            while ($State.ErrQueue.TryDequeue([ref]$line)) { Write-Output-Line $line }
            $code = $State.Process.ExitCode
            Write-Output-Line ''
            Write-Output-Line ("--- finished, exit code {0} ---" -f $code)
            Complete-Run ("Finished. Exit code {0}." -f $code)
        }
    })
    $State.Timer = $timer
    $timer.Start()
}

# ------------------------------------------------------------- 6. wiring up
function Open-LastReport {
    <#
        Opens the newest file the selected tool wrote. The output folder comes
        from whichever parameter the script called OutputDir or LogDir, read
        straight off the form, so this needs no table of tools either.
    #>
    if ($null -eq $State.Tool) { return }

    $dir = ''
    foreach ($p in $State.Tool.Params) {
        if ($p.Name -notin @('OutputDir', 'LogDir')) { continue }
        if (-not $State.Controls.ContainsKey($p.Name)) { continue }
        $dir = [string](Read-ParameterControl -Param $p -Control $State.Controls[$p.Name])
        break
    }

    if (-not $dir -or -not (Test-Path -LiteralPath $dir)) {
        $StatusText.Text = 'No output folder yet - run it first.'
        return
    }

    $newest = @(Get-ChildItem -LiteralPath $dir -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    try {
        if ($newest.Count -gt 0) { Start-Process -FilePath $newest[0].FullName }
        else { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $dir) }
    }
    catch { $StatusText.Text = 'Could not open it: ' + $_.Exception.Message }
}

$ToolList.ItemsSource = $Tools
$ToolList.Add_SelectionChanged({
    if ($ToolList.SelectedItem) { Show-Tool $ToolList.SelectedItem }
})

$RunButton.Add_Click({ Invoke-Tool })
$StopButton.Add_Click({ Stop-RunningTool; $StatusText.Text = 'Stopping...' })
$ReportButton.Add_Click({ Open-LastReport })
$HelpButton.Add_Click({
    if ($null -eq $State.Tool -or $State.Tool.Kind -ne 'Script') { return }
    Start-Process -FilePath $PsExe -ArgumentList @(
        '-NoProfile', '-NoExit', '-Command',
        ("Get-Help '" + ($State.Tool.Path -replace "'", "''") + "' -Full | Out-Host"))
})

$Window.Add_Closing({ Stop-RunningTool })

$ToolList.SelectedIndex = 0
$StatusText.Text = ("{0} tools found in {1}" -f $Tools.Count, $ToolRoot)

[void]$Window.ShowDialog()
