#requires -Version 5.1
<#
.SYNOPSIS
    Simple WPF disk space analyzer ("TreeSize Light") in a single PowerShell file.

.DESCRIPTION
    - Folder selection
    - Scan in a background runspace so the WPF interface remains responsive
    - Scan can be cancelled
    - Tree view of subfolders
    - Size bar for each folder relative to its parent folder
    - Live status with scanned files, bytes, current path, and skipped entries
    - Reparse points/junctions are skipped to avoid loops and double counting

    Note:
    An exact scan percentage cannot be determined meaningfully without first performing a
    complete counting pass. The global progress bar is therefore indeterminate during scanning.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# WPF requires an STA thread.
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    if ($PSCommandPath) {
        $exe = (Get-Process -Id $PID).Path
        $args = @(
            '-NoProfile'
            '-STA'
            '-ExecutionPolicy', 'Bypass'
            '-File', ('"{0}"' -f $PSCommandPath)
        )
        Start-Process -FilePath $exe -ArgumentList $args | Out-Null
        exit
    }
    throw 'This script must run in an STA thread. Start it, for example, with powershell.exe -STA -File .\StorageTree.en.ps1'
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="StorageTree - PowerShell Disk Space Analysis"
    Width="1050"
    Height="720"
    MinWidth="780"
    MinHeight="520"
    WindowStartupLocation="CenterScreen">

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="10"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Padding="10" BorderBrush="#D8D8D8" BorderThickness="1" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Grid.Column="0"
                           Text="Folder:"
                           VerticalAlignment="Center"
                           Margin="0,0,8,0"/>

                <TextBox x:Name="PathBox"
                         Grid.Column="1"
                         MinWidth="300"
                         VerticalContentAlignment="Center"
                         Margin="0,0,8,0"/>

                <Button x:Name="BrowseButton"
                        Grid.Column="2"
                        Content="Browse…"
                        Padding="12,5"
                        Margin="0,0,8,0"/>

                <Button x:Name="ScanButton"
                        Grid.Column="3"
                        Content="Scan"
                        Padding="16,5"
                        Margin="0,0,8,0"/>

                <Button x:Name="CancelButton"
                        Grid.Column="4"
                        Content="Cancel"
                        Padding="12,5"
                        IsEnabled="False"/>
            </Grid>
        </Border>

        <Border Grid.Row="2" BorderBrush="#D8D8D8" BorderThickness="1" CornerRadius="4">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="10,8,10,6">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="140"/>
                        <ColumnDefinition Width="110"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Text="Folder"
                               FontWeight="SemiBold"/>

                    <TextBlock Grid.Column="1"
                               Text="Size"
                               FontWeight="SemiBold"
                               HorizontalAlignment="Right"
                               Margin="0,0,12,0"/>

                    <TextBlock Grid.Column="2"
                               Text="Share"
                               FontWeight="SemiBold"
                               HorizontalAlignment="Center"/>
                </Grid>

                <TreeView x:Name="FolderTree"
                          Grid.Row="1"
                          Margin="6"
                          BorderThickness="0"
                          VirtualizingStackPanel.IsVirtualizing="True"
                          VirtualizingStackPanel.VirtualizationMode="Recycling"/>
            </Grid>
        </Border>

        <Border Grid.Row="4" Padding="10" BorderBrush="#D8D8D8" BorderThickness="1" CornerRadius="4">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="4"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <ProgressBar x:Name="ScanProgress"
                             Grid.Row="0"
                             Height="16"
                             Minimum="0"
                             Maximum="100"
                             Value="0"/>

                <TextBlock x:Name="StatusText"
                           Grid.Row="2"
                           Text="Ready."
                           TextTrimming="CharacterEllipsis"/>

                <TextBlock x:Name="SummaryText"
                           Grid.Row="4"
                           Foreground="#666666"
                           Text=""/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$PathBox      = $window.FindName('PathBox')
$BrowseButton = $window.FindName('BrowseButton')
$ScanButton   = $window.FindName('ScanButton')
$CancelButton = $window.FindName('CancelButton')
$FolderTree   = $window.FindName('FolderTree')
$ScanProgress = $window.FindName('ScanProgress')
$StatusText   = $window.FindName('StatusText')
$SummaryText  = $window.FindName('SummaryText')

$PathBox.Text = [Environment]::GetFolderPath('UserProfile')

function Format-Size {
    param([long]$Bytes)

    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return ('{0:N0} B' -f $Bytes)
}

function Get-NodeProperty {
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)][string]$Name
    )
    return $Node.PSObject.Properties[$Name].Value
}

function New-TreeHeader {
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)][long]$ParentSize
    )

    $nodeSize = [long](Get-NodeProperty $Node 'Size')
    $nodeName = [string](Get-NodeProperty $Node 'Name')

    $percent = if ($ParentSize -gt 0) {
        [Math]::Min(100.0, ($nodeSize / [double]$ParentSize) * 100.0)
    } else {
        0.0
    }

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = '1,2,1,2'
    $grid.MinWidth = 650

    $c1 = New-Object System.Windows.Controls.ColumnDefinition
    $c1.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
    $c2 = New-Object System.Windows.Controls.ColumnDefinition
    $c2.Width = New-Object System.Windows.GridLength(140)
    $c3 = New-Object System.Windows.Controls.ColumnDefinition
    $c3.Width = New-Object System.Windows.GridLength(110)

    [void]$grid.ColumnDefinitions.Add($c1)
    [void]$grid.ColumnDefinitions.Add($c2)
    [void]$grid.ColumnDefinitions.Add($c3)

    $nameText = New-Object System.Windows.Controls.TextBlock
    $nameText.Text = $nodeName
    $nameText.VerticalAlignment = 'Center'
    $nameText.TextTrimming = 'CharacterEllipsis'
    [System.Windows.Controls.Grid]::SetColumn($nameText, 0)

    $sizeText = New-Object System.Windows.Controls.TextBlock
    $sizeText.Text = Format-Size $nodeSize
    $sizeText.VerticalAlignment = 'Center'
    $sizeText.HorizontalAlignment = 'Right'
    $sizeText.Margin = '0,0,12,0'
    [System.Windows.Controls.Grid]::SetColumn($sizeText, 1)

    $barHost = New-Object System.Windows.Controls.Grid
    [System.Windows.Controls.Grid]::SetColumn($barHost, 2)

    $bar = New-Object System.Windows.Controls.ProgressBar
    $bar.Minimum = 0
    $bar.Maximum = 100
    $bar.Value = $percent
    $bar.Height = 16
    $bar.Width = 92
    $bar.HorizontalAlignment = 'Center'
    $bar.VerticalAlignment = 'Center'
    $bar.ToolTip = ('{0:N1} % of parent folder' -f $percent)

    $percentText = New-Object System.Windows.Controls.TextBlock
    $percentText.Text = ('{0:N0}%' -f $percent)
    $percentText.HorizontalAlignment = 'Center'
    $percentText.VerticalAlignment = 'Center'
    $percentText.FontSize = 10
    $percentText.IsHitTestVisible = $false

    [void]$barHost.Children.Add($bar)
    [void]$barHost.Children.Add($percentText)

    [void]$grid.Children.Add($nameText)
    [void]$grid.Children.Add($sizeText)
    [void]$grid.Children.Add($barHost)

    return $grid
}

function Add-LazyChildren {
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TreeViewItem]$Item
    )

    if ($Item.Items.Count -ne 1) { return }

    $first = $Item.Items[0]
    if (-not ($first -is [System.Windows.Controls.TreeViewItem])) { return }
    if ($first.Tag -ne '__PLACEHOLDER__') { return }

    $Item.Items.Clear()

    $node = $Item.Tag
    $nodeSize = [long](Get-NodeProperty $node 'Size')
    $children = @(Get-NodeProperty $node 'Children')

    foreach ($child in $children) {
        [void]$Item.Items.Add((New-TreeItem -Node $child -ParentSize $nodeSize))
    }
}

function New-TreeItem {
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)][long]$ParentSize
    )

    $item = New-Object System.Windows.Controls.TreeViewItem
    $item.Tag = $Node
    $item.Header = New-TreeHeader -Node $Node -ParentSize $ParentSize
    $item.ToolTip = [string](Get-NodeProperty $Node 'Path')

    $children = @(Get-NodeProperty $Node 'Children')
    if ($children.Count -gt 0) {
        $placeholder = New-Object System.Windows.Controls.TreeViewItem
        $placeholder.Header = '…'
        $placeholder.Tag = '__PLACEHOLDER__'
        [void]$item.Items.Add($placeholder)

        $item.Add_Expanded({
            param($sender, $eventArgs)
            Add-LazyChildren -Item $sender
        })
    }

    return $item
}

$scanScript = {
    param(
        [string]$RootPath,
        [hashtable]$State
    )

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    function Note-Skip {
        param([string]$Message)
        $State.Skipped = [long]$State.Skipped + 1
        $State.LastError = $Message
    }

    function Scan-Folder {
        param(
            [string]$Path,
            [bool]$IsRoot = $false
        )

        if ($State.Cancel) { return $null }

        $State.CurrentPath = $Path

        try {
            $dirInfo = [System.IO.DirectoryInfo]::new($Path)
        }
        catch {
            Note-Skip $_.Exception.Message
            return $null
        }

        if (-not $IsRoot) {
            try {
                if (($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                    $State.ReparseSkipped = [long]$State.ReparseSkipped + 1
                    return $null
                }
            }
            catch {
                Note-Skip $_.Exception.Message
                return $null
            }
        }

        $totalSize = [long]0
        $directFiles = [long]0
        $totalFiles = [long]0
        $totalFolders = [long]0
        $children = New-Object 'System.Collections.Generic.List[object]'

        # Files directly in this folder
        try {
            foreach ($file in $dirInfo.EnumerateFiles('*', [System.IO.SearchOption]::TopDirectoryOnly)) {
                if ($State.Cancel) { break }

                try {
                    $length = [long]$file.Length
                    $totalSize += $length
                    $directFiles++
                    $totalFiles++

                    $State.Files = [long]$State.Files + 1
                    $State.Bytes = [long]$State.Bytes + $length

                    if (([long]$State.Files % 40) -eq 0) {
                        $State.CurrentPath = $file.FullName
                    }
                }
                catch {
                    Note-Skip $_.Exception.Message
                }
            }
        }
        catch {
            Note-Skip $_.Exception.Message
        }

        # Subfolders
        if (-not $State.Cancel) {
            try {
                foreach ($subDir in $dirInfo.EnumerateDirectories('*', [System.IO.SearchOption]::TopDirectoryOnly)) {
                    if ($State.Cancel) { break }

                    try {
                        if (($subDir.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                            $State.ReparseSkipped = [long]$State.ReparseSkipped + 1
                            continue
                        }
                    }
                    catch {
                        Note-Skip $_.Exception.Message
                        continue
                    }

                    $State.Folders = [long]$State.Folders + 1
                    $child = Scan-Folder -Path $subDir.FullName

                    if ($null -ne $child) {
                        $children.Add($child) | Out-Null
                        $totalSize += [long]$child.Size
                        $totalFiles += [long]$child.TotalFiles
                        $totalFolders += 1 + [long]$child.TotalFolders
                    }
                }
            }
            catch {
                Note-Skip $_.Exception.Message
            }
        }

        $sortedChildren = @($children | Sort-Object -Property Size -Descending)

        [PSCustomObject]@{
            Name         = $dirInfo.Name
            Path         = $dirInfo.FullName
            Size         = $totalSize
            DirectFiles  = $directFiles
            TotalFiles   = $totalFiles
            TotalFolders = $totalFolders
            Children     = $sortedChildren
        }
    }

    try {
        $root = Scan-Folder -Path $RootPath -IsRoot $true
        if ($null -ne $root) {
            $root
        }
    }
    catch {
        $State.FatalError = $_.Exception.ToString()
    }
    finally {
        $State.Done = $true
    }
}

$script:ScanState = $null
$script:ScanPowerShell = $null
$script:ScanRunspace = $null
$script:ScanAsync = $null
$script:ScanTimer = $null

function Set-ScanningUi {
    param([bool]$Scanning)

    $BrowseButton.IsEnabled = -not $Scanning
    $ScanButton.IsEnabled = -not $Scanning
    $PathBox.IsEnabled = -not $Scanning
    $CancelButton.IsEnabled = $Scanning

    $ScanProgress.IsIndeterminate = $Scanning
    if (-not $Scanning) {
        $ScanProgress.Value = 0
    }
}

function Stop-ScanResources {
    if ($script:ScanTimer) {
        $script:ScanTimer.Stop()
        $script:ScanTimer = $null
    }

    if ($script:ScanPowerShell) {
        $script:ScanPowerShell.Dispose()
        $script:ScanPowerShell = $null
    }

    if ($script:ScanRunspace) {
        try { $script:ScanRunspace.Close() } catch {}
        $script:ScanRunspace.Dispose()
        $script:ScanRunspace = $null
    }

    $script:ScanAsync = $null
}

function Complete-Scan {
    try {
        $results = @($script:ScanPowerShell.EndInvoke($script:ScanAsync))
        $root = $results | Select-Object -First 1

        $FolderTree.Items.Clear()

        if ($null -ne $root) {
            $rootSize = [long](Get-NodeProperty $root 'Size')
            $rootItem = New-TreeItem -Node $root -ParentSize $rootSize

            # Expand the root immediately; deeper levels remain lazy-loaded.
            Add-LazyChildren -Item $rootItem

            [void]$FolderTree.Items.Add($rootItem)
            $rootItem.IsExpanded = $true

            $cancelled = [bool]$script:ScanState.Cancel
            if ($cancelled) {
                $StatusText.Text = 'Scan cancelled – partial results are shown.'
            }
            else {
                $StatusText.Text = 'Scan completed.'
            }

            $SummaryText.Text = (
                '{0} • {1:N0} files • {2:N0} subfolders • {3:N0} skipped • {4:N0} reparse points skipped' -f
                    (Format-Size $rootSize),
                    [long](Get-NodeProperty $root 'TotalFiles'),
                    [long](Get-NodeProperty $root 'TotalFolders'),
                    [long]$script:ScanState.Skipped,
                    [long]$script:ScanState.ReparseSkipped
            )
        }
        elseif ($script:ScanState.FatalError) {
            $StatusText.Text = 'Scan failed.'
            $SummaryText.Text = [string]$script:ScanState.FatalError
        }
        else {
            $StatusText.Text = 'No data found.'
            $SummaryText.Text = ''
        }
    }
    catch {
        $StatusText.Text = 'Error while completing the scan.'
        $SummaryText.Text = $_.Exception.Message
    }
    finally {
        Set-ScanningUi $false
        Stop-ScanResources
    }
}

function Start-Scan {
    $path = $PathBox.Text.Trim()

    if (-not [System.IO.Directory]::Exists($path)) {
        [System.Windows.MessageBox]::Show(
            $window,
            'The specified folder does not exist.',
            'Invalid folder',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }

    Stop-ScanResources
    $FolderTree.Items.Clear()
    $SummaryText.Text = ''
    $StatusText.Text = 'Starting scan…'

    $script:ScanState = [hashtable]::Synchronized(@{
        Cancel         = $false
        Done           = $false
        Files          = [long]0
        Folders        = [long]0
        Bytes          = [long]0
        Skipped        = [long]0
        ReparseSkipped = [long]0
        CurrentPath    = $path
        LastError      = ''
        FatalError     = ''
    })

    $script:ScanRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:ScanRunspace.ApartmentState = [Threading.ApartmentState]::MTA
    $script:ScanRunspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $script:ScanRunspace.Open()

    $script:ScanPowerShell = [System.Management.Automation.PowerShell]::Create()
    $script:ScanPowerShell.Runspace = $script:ScanRunspace

    [void]$script:ScanPowerShell.AddScript($scanScript.ToString())
    [void]$script:ScanPowerShell.AddArgument($path)
    [void]$script:ScanPowerShell.AddArgument($script:ScanState)

    $script:ScanAsync = $script:ScanPowerShell.BeginInvoke()

    Set-ScanningUi $true

    $script:ScanTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:ScanTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $script:ScanTimer.Add_Tick({
        if ($null -eq $script:ScanState) { return }

        $files = [long]$script:ScanState.Files
        $folders = [long]$script:ScanState.Folders
        $bytes = [long]$script:ScanState.Bytes
        $current = [string]$script:ScanState.CurrentPath
        $skipped = [long]$script:ScanState.Skipped

        $StatusText.Text = (
            'Scanning… {0:N0} files, {1:N0} folders, {2} read, {3:N0} skipped — {4}' -f
                $files, $folders, (Format-Size $bytes), $skipped, $current
        )

        if ($script:ScanAsync -and $script:ScanAsync.IsCompleted) {
            $script:ScanTimer.Stop()
            Complete-Scan
        }
    })

    $script:ScanTimer.Start()
}

$BrowseButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select a folder for disk space analysis'
    $dialog.ShowNewFolderButton = $false

    if ([System.IO.Directory]::Exists($PathBox.Text)) {
        $dialog.SelectedPath = $PathBox.Text
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $PathBox.Text = $dialog.SelectedPath
    }

    $dialog.Dispose()
})

$ScanButton.Add_Click({
    Start-Scan
})

$CancelButton.Add_Click({
    if ($script:ScanState) {
        $script:ScanState.Cancel = $true
        $CancelButton.IsEnabled = $false
        $StatusText.Text = 'Cancellation requested…'
    }
})

$PathBox.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and $ScanButton.IsEnabled) {
        Start-Scan
    }
})

$FolderTree.Add_MouseDoubleClick({
    param($sender, $eventArgs)

    $item = $FolderTree.SelectedItem
    if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Tag -and $item.Tag -ne '__PLACEHOLDER__') {
        try {
            $path = [string](Get-NodeProperty $item.Tag 'Path')
            if ([System.IO.Directory]::Exists($path)) {
                Start-Process explorer.exe -ArgumentList ('"{0}"' -f $path)
            }
        }
        catch {}
    }
})

$window.Add_Closing({
    if ($script:ScanState) {
        $script:ScanState.Cancel = $true
    }

    if ($script:ScanPowerShell -and $script:ScanAsync -and -not $script:ScanAsync.IsCompleted) {
        try { $script:ScanPowerShell.Stop() } catch {}
    }

    Stop-ScanResources
})

[void]$window.ShowDialog()
