#requires -Version 5.1
<#
.SYNOPSIS
    WPF disk space analyzer with an interactive SequoiaView-style treemap.

.DESCRIPTION
    - Folder selection
    - Scanning in a background runspace so the WPF interface remains responsive
    - Scan cancellation
    - Tree view of subfolders
    - Size bar for each folder relative to its parent folder
    - Explorer-like file view next to the folder tree
    - Top 200 largest files in the entire scan
    - Squarified treemap for the global top files or the selected folder
    - Colors by file type, tooltips, synchronized selection, and opening by double-click
    - Configurable rectangle limit and aggregate area for responsive rendering
    - Right-click files to open them, reveal them in Explorer, copy their path, or recycle them
    - Additional Windows shell actions in a submenu
    - Live status showing scanned files, bytes, the current path, and skipped entries
    - Reparse points/junctions are skipped to avoid loops and double counting

    Note:
    An exact scan percentage cannot be determined meaningfully without first performing a
    complete counting pass. The global progress bar is therefore indeterminate during a scan.

    Treemap data and limitations:
    The treemap uses the same data as the file list: either the global top 200 files from the
    latest scan or all files located directly in the selected folder. It does not perform an
    additional recursive scan. The selectable limit (default: 500 rectangles) only restricts
    WPF rendering; when there are more files, one proportional rectangle aggregates the rest.
    Zero-byte files have no visible area. The display uses logical file sizes; NTFS compression,
    sparse files, and allocated disk space can differ.
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
    throw 'This script must run in an STA thread. For example, start it with powershell.exe -STA -File .\WinDirStat3.en.ps1'
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="WinDirStat 3.0 - PowerShell Storage Analysis with Treemap"
    Width="1250"
    Height="760"
    MinWidth="900"
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
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="11*" MinWidth="360"/>
                    <ColumnDefinition Width="6"/>
                    <ColumnDefinition Width="10*" MinWidth="360"/>
                </Grid.ColumnDefinitions>

                <!-- Folder tree -->
                <Grid Grid.Column="0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="10,8,10,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="110"/>
                            <ColumnDefinition Width="95"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Text="Folder" FontWeight="SemiBold"/>
                        <TextBlock Grid.Column="1" Text="Size" FontWeight="SemiBold"
                                   HorizontalAlignment="Right" Margin="0,0,12,0"/>
                        <TextBlock Grid.Column="2" Text="Share" FontWeight="SemiBold"
                                   HorizontalAlignment="Center"/>
                    </Grid>

                    <TreeView x:Name="FolderTree"
                              Grid.Row="1"
                              Margin="6"
                              BorderThickness="0"
                              VirtualizingStackPanel.IsVirtualizing="True"
                              VirtualizingStackPanel.VirtualizationMode="Recycling"/>
                </Grid>

                <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch"
                              VerticalAlignment="Stretch" Background="#E6E6E6"/>

                <!-- File view -->
                <Grid Grid.Column="2">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="10,7,8,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="FilesHeading" Text="Largest Files"
                                   FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <Button x:Name="ShowTopFilesButton" Grid.Column="1"
                                Content="Top 200 in Scan" Padding="9,3"/>
                    </Grid>

                    <TextBlock x:Name="FilesPathText" Grid.Row="1" Margin="10,0,8,5"
                               Foreground="#666666" Text="The largest files will be shown here after the scan."
                               TextTrimming="CharacterEllipsis"/>

                    <TabControl x:Name="FileViews" Grid.Row="2" Margin="6">
                        <TabItem Header="File List">
                            <DataGrid x:Name="FileGrid"
                                      AutoGenerateColumns="False" IsReadOnly="True"
                                      CanUserAddRows="False" CanUserDeleteRows="False"
                                      CanUserReorderColumns="True" CanUserResizeColumns="True"
                                      CanUserSortColumns="False" SelectionMode="Single"
                                      HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                      AlternatingRowBackground="#FAFAFA"
                                      VirtualizingPanel.IsVirtualizing="True"
                                      VirtualizingPanel.VirtualizationMode="Recycling">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                                    <DataGridTextColumn Header="Size" Binding="{Binding SizeText}" Width="90"/>
                                    <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="85"/>
                                    <DataGridTextColumn Header="Modified" Binding="{Binding ModifiedText}" Width="125"/>
                                    <DataGridTextColumn Header="Path" Binding="{Binding Directory}" Width="2*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </TabItem>

                        <TabItem x:Name="TreemapTab" Header="Treemap">
                            <Grid Margin="2">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>

                                <DockPanel Grid.Row="0" Margin="5,3,5,6" LastChildFill="True">
                                    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                                        <TextBlock Text="Max. rectangles:"
                                                   VerticalAlignment="Center" Margin="10,0,5,0"/>
                                        <ComboBox x:Name="TreemapLimitBox" Width="72"
                                                  SelectedIndex="2" VerticalContentAlignment="Center">
                                            <ComboBoxItem Content="100"/>
                                            <ComboBoxItem Content="250"/>
                                            <ComboBoxItem Content="500"/>
                                            <ComboBoxItem Content="750"/>
                                            <ComboBoxItem Content="1000"/>
                                        </ComboBox>
                                    </StackPanel>
                                    <TextBlock x:Name="TreemapInfoText"
                                               Foreground="#666666"
                                               VerticalAlignment="Center"
                                               Text="The treemap will be available here after a scan."
                                               TextTrimming="CharacterEllipsis"/>
                                </DockPanel>

                                <Border Grid.Row="1" BorderBrush="#50555B" BorderThickness="1"
                                        Background="#171A1E" SnapsToDevicePixels="True">
                                    <Grid x:Name="TreemapHost" Background="#171A1E"
                                          ClipToBounds="True" MinHeight="180">
                                        <Canvas x:Name="TreemapCanvas" Background="#171A1E"
                                                ClipToBounds="True"/>
                                        <TextBlock x:Name="TreemapEmptyText"
                                                   Text="No files to display."
                                                   Foreground="#D6D9DC" FontSize="14"
                                                   HorizontalAlignment="Center"
                                                   VerticalAlignment="Center"
                                                   TextAlignment="Center"
                                                   IsHitTestVisible="False"/>
                                    </Grid>
                                </Border>
                            </Grid>
                        </TabItem>
                    </TabControl>
                </Grid>
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
$FolderTree        = $window.FindName('FolderTree')
$FileGrid          = $window.FindName('FileGrid')
$FileViews         = $window.FindName('FileViews')
$TreemapTab        = $window.FindName('TreemapTab')
$TreemapHost       = $window.FindName('TreemapHost')
$TreemapCanvas     = $window.FindName('TreemapCanvas')
$TreemapEmptyText  = $window.FindName('TreemapEmptyText')
$TreemapInfoText   = $window.FindName('TreemapInfoText')
$TreemapLimitBox   = $window.FindName('TreemapLimitBox')
$FilesHeading      = $window.FindName('FilesHeading')
$FilesPathText     = $window.FindName('FilesPathText')
$ShowTopFilesButton = $window.FindName('ShowTopFilesButton')
$ScanProgress      = $window.FindName('ScanProgress')
$StatusText   = $window.FindName('StatusText')
$SummaryText  = $window.FindName('SummaryText')

$PathBox.Text = [Environment]::GetFolderPath('UserProfile')

$script:CurrentFileRows = @()
$script:TreemapRenderTimer = $null
$script:TreemapElementsByPath = @{}
$script:SelectedTreemapElement = $null
$script:TreemapBrushCache = @{}

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


function New-FileRow {
    param(
        [Parameter(Mandatory)][string]$FullName,
        [Parameter(Mandatory)][long]$Length,
        [Parameter(Mandatory)][datetime]$LastWriteTime
    )

    $name = [System.IO.Path]::GetFileName($FullName)
    $directory = [System.IO.Path]::GetDirectoryName($FullName)
    $extension = [System.IO.Path]::GetExtension($FullName)
    $type = if ([string]::IsNullOrWhiteSpace($extension)) {
        'File'
    }
    else {
        $extension.TrimStart('.').ToUpperInvariant()
    }

    [PSCustomObject]@{
        Name         = $name
        Length       = $Length
        SizeText     = Format-Size $Length
        Type         = $type
        LastWriteTime = $LastWriteTime
        ModifiedText = $LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        Directory    = $directory
        FullName     = $FullName
    }
}

function Get-TreemapLimit {
    $limit = 500

    if ($null -ne $TreemapLimitBox -and $null -ne $TreemapLimitBox.SelectedItem) {
        $parsed = 0
        if ([int]::TryParse(
                [string]$TreemapLimitBox.SelectedItem.Content,
                [ref]$parsed
            )) {
            $limit = $parsed
        }
    }

    return [Math]::Max(25, [Math]::Min(2000, $limit))
}

function Test-TreemapAggregateRow {
    param($Row)

    return (
        $null -ne $Row -and
        $null -ne $Row.PSObject.Properties['IsAggregate'] -and
        [bool]$Row.IsAggregate
    )
}

function Get-TreemapDisplayRows {
    param(
        [object[]]$Rows,
        [int]$Limit
    )

    $positiveRows = New-Object 'System.Collections.Generic.List[object]'
    foreach ($row in @($Rows)) {
        if ($null -ne $row -and [long]$row.Length -gt 0) {
            $positiveRows.Add($row) | Out-Null
        }
    }

    if ($positiveRows.Count -le $Limit) {
        return $positiveRows.ToArray()
    }

    # One slot is reserved for the aggregate area. This keeps the total area
    # proportional without creating thousands of WPF elements for very large folders.
    $individualCount = [Math]::Max(1, $Limit - 1)
    $displayRows = New-Object 'System.Collections.Generic.List[object]'

    for ($index = 0; $index -lt $individualCount; $index++) {
        $displayRows.Add($positiveRows[$index]) | Out-Null
    }

    $remainingBytes = [long]0
    for ($index = $individualCount; $index -lt $positiveRows.Count; $index++) {
        $remainingBytes += [long]$positiveRows[$index].Length
    }

    $remainingCount = $positiveRows.Count - $individualCount
    $displayRows.Add([PSCustomObject]@{
        Name          = ('Remaining {0:N0} files' -f $remainingCount)
        Length        = $remainingBytes
        SizeText      = Format-Size $remainingBytes
        Type          = 'Weitere'
        LastWriteTime = [datetime]::MinValue
        ModifiedText  = ''
        Directory     = ''
        FullName      = $null
        IsAggregate   = $true
        FileCount     = $remainingCount
    }) | Out-Null

    return $displayRows.ToArray()
}

function Get-TreemapWorstAspect {
    param(
        [object[]]$Entries,
        [double]$ShortSide
    )

    if ($Entries.Count -eq 0 -or $ShortSide -le 0) {
        return [double]::PositiveInfinity
    }

    $sum = [double]0
    $minimum = [double]::PositiveInfinity
    $maximum = [double]0

    foreach ($entry in $Entries) {
        $area = [double]$entry.Area
        $sum += $area
        if ($area -lt $minimum) { $minimum = $area }
        if ($area -gt $maximum) { $maximum = $area }
    }

    if ($sum -le 0 -or $minimum -le 0) {
        return [double]::PositiveInfinity
    }

    $sideSquared = $ShortSide * $ShortSide
    $sumSquared = $sum * $sum

    return [Math]::Max(
        ($sideSquared * $maximum) / $sumSquared,
        $sumSquared / ($sideSquared * $minimum)
    )
}

function Add-TreemapStrip {
    param(
        [object[]]$Entries,
        $Bounds,
        [System.Collections.Generic.List[object]]$Result
    )

    if ($Entries.Count -eq 0 -or $Bounds.Width -le 0 -or $Bounds.Height -le 0) {
        return
    }

    $stripArea = [double]0
    foreach ($entry in $Entries) {
        $stripArea += [double]$entry.Area
    }

    if ($Bounds.Width -ge $Bounds.Height) {
        # Wide remaining area: place a vertical column on the left.
        $stripWidth = [Math]::Min([double]$Bounds.Width, $stripArea / [double]$Bounds.Height)
        if ($stripWidth -le 0) { return }

        $cursorY = [double]$Bounds.Y
        for ($index = 0; $index -lt $Entries.Count; $index++) {
            $itemHeight = if ($index -eq $Entries.Count - 1) {
                ([double]$Bounds.Y + [double]$Bounds.Height) - $cursorY
            }
            else {
                [double]$Entries[$index].Area / $stripWidth
            }

            $Result.Add([PSCustomObject]@{
                Item   = $Entries[$index].Item
                X      = [double]$Bounds.X
                Y      = $cursorY
                Width  = $stripWidth
                Height = [Math]::Max(0.0, $itemHeight)
            }) | Out-Null
            $cursorY += $itemHeight
        }

        $Bounds.X = [double]$Bounds.X + $stripWidth
        $Bounds.Width = [Math]::Max(0.0, [double]$Bounds.Width - $stripWidth)
    }
    else {
        # Tall remaining area: place a horizontal row at the top.
        $stripHeight = [Math]::Min([double]$Bounds.Height, $stripArea / [double]$Bounds.Width)
        if ($stripHeight -le 0) { return }

        $cursorX = [double]$Bounds.X
        for ($index = 0; $index -lt $Entries.Count; $index++) {
            $itemWidth = if ($index -eq $Entries.Count - 1) {
                ([double]$Bounds.X + [double]$Bounds.Width) - $cursorX
            }
            else {
                [double]$Entries[$index].Area / $stripHeight
            }

            $Result.Add([PSCustomObject]@{
                Item   = $Entries[$index].Item
                X      = $cursorX
                Y      = [double]$Bounds.Y
                Width  = [Math]::Max(0.0, $itemWidth)
                Height = $stripHeight
            }) | Out-Null
            $cursorX += $itemWidth
        }

        $Bounds.Y = [double]$Bounds.Y + $stripHeight
        $Bounds.Height = [Math]::Max(0.0, [double]$Bounds.Height - $stripHeight)
    }
}

function Get-SquarifiedTreemapLayout {
    param(
        [object[]]$Items,
        [double]$Width,
        [double]$Height
    )

    $weightedItems = New-Object 'System.Collections.Generic.List[object]'
    $total = [double]0

    foreach ($item in @($Items)) {
        $value = [double][long]$item.Length
        if ($value -gt 0) {
            $entry = [PSCustomObject]@{
                Item  = $item
                Value = $value
                Area  = [double]0
            }
            $weightedItems.Add($entry) | Out-Null
            $total += $value
        }
    }

    if ($weightedItems.Count -eq 0 -or $total -le 0 -or $Width -le 0 -or $Height -le 0) {
        return @()
    }

    $availableArea = $Width * $Height
    foreach ($entry in $weightedItems) {
        $entry.Area = ([double]$entry.Value / $total) * $availableArea
    }

    $bounds = [PSCustomObject]@{
        X      = [double]0
        Y      = [double]0
        Width  = $Width
        Height = $Height
    }
    $result = New-Object 'System.Collections.Generic.List[object]'
    $row = New-Object 'System.Collections.Generic.List[object]'
    $itemIndex = 0

    while ($itemIndex -lt $weightedItems.Count) {
        $next = $weightedItems[$itemIndex]
        if ($row.Count -eq 0) {
            $row.Add($next) | Out-Null
            $itemIndex++
            continue
        }

        $candidate = New-Object 'System.Collections.Generic.List[object]'
        foreach ($entry in $row) {
            $candidate.Add($entry) | Out-Null
        }
        $candidate.Add($next) | Out-Null

        $shortSide = [Math]::Min([double]$bounds.Width, [double]$bounds.Height)
        $currentWorst = Get-TreemapWorstAspect -Entries $row.ToArray() -ShortSide $shortSide
        $candidateWorst = Get-TreemapWorstAspect -Entries $candidate.ToArray() -ShortSide $shortSide

        if ($candidateWorst -le $currentWorst) {
            $row.Add($next) | Out-Null
            $itemIndex++
        }
        else {
            Add-TreemapStrip -Entries $row.ToArray() -Bounds $bounds -Result $result
            $row.Clear()
        }
    }

    if ($row.Count -gt 0) {
        Add-TreemapStrip -Entries $row.ToArray() -Bounds $bounds -Result $result
    }

    return $result.ToArray()
}

function Get-TreemapBrush {
    param($Row)

    $type = ([string]$Row.Type).ToUpperInvariant()
    $category = if (Test-TreemapAggregateRow $Row) {
        'Weitere'
    }
    elseif (@('JPG', 'JPEG', 'PNG', 'GIF', 'BMP', 'WEBP', 'TIFF', 'TIF', 'SVG', 'HEIC') -contains $type) {
        'Bild'
    }
    elseif (@('MP4', 'MKV', 'AVI', 'MOV', 'WMV', 'WEBM', 'M4V', 'MPEG', 'MPG') -contains $type) {
        'Video'
    }
    elseif (@('MP3', 'WAV', 'FLAC', 'AAC', 'M4A', 'OGG', 'WMA') -contains $type) {
        'Audio'
    }
    elseif (@('ZIP', '7Z', 'RAR', 'TAR', 'GZ', 'BZ2', 'XZ', 'CAB', 'ISO') -contains $type) {
        'Archiv'
    }
    elseif (@('DOC', 'DOCX', 'PDF', 'XLS', 'XLSX', 'PPT', 'PPTX', 'ODT', 'ODS', 'TXT', 'RTF', 'CSV') -contains $type) {
        'Dokument'
    }
    elseif (@('PS1', 'PSM1', 'PSD1', 'CS', 'VB', 'JS', 'TS', 'HTML', 'HTM', 'CSS', 'XML', 'JSON', 'YAML', 'YML', 'PY', 'JAVA', 'CPP', 'C', 'H', 'SQL') -contains $type) {
        'Quelltext'
    }
    elseif (@('EXE', 'DLL', 'MSI', 'SYS', 'COM', 'BAT', 'CMD', 'SCR') -contains $type) {
        'Programm'
    }
    else {
        'Sonstige'
    }

    if ($script:TreemapBrushCache.ContainsKey($category)) {
        return $script:TreemapBrushCache[$category]
    }

    $colors = @{
        Bild      = '#C96D21'
        Video     = '#326D9A'
        Audio     = '#3C873E'
        Archiv    = '#B64645'
        Dokument  = '#765395'
        Quelltext = '#25827E'
        Programm  = '#B74F78'
        Sonstige  = '#64727D'
        Weitere   = '#4C4F53'
    }

    $baseColor = [System.Windows.Media.ColorConverter]::ConvertFromString([string]$colors[$category])
    $lightColor = [System.Windows.Media.Color]::FromRgb(
        [byte][Math]::Min(255, [int]$baseColor.R + 38),
        [byte][Math]::Min(255, [int]$baseColor.G + 38),
        [byte][Math]::Min(255, [int]$baseColor.B + 38)
    )
    $darkColor = [System.Windows.Media.Color]::FromRgb(
        [byte][Math]::Max(0, [int]$baseColor.R - 24),
        [byte][Math]::Max(0, [int]$baseColor.G - 24),
        [byte][Math]::Max(0, [int]$baseColor.B - 24)
    )

    # The diagonal gradient imitates SequoiaView's raised "cushions" while remaining
    # considerably less expensive than a custom pixel renderer.
    $brush = New-Object System.Windows.Media.LinearGradientBrush
    $brush.StartPoint = [System.Windows.Point]::new(0, 0)
    $brush.EndPoint = [System.Windows.Point]::new(1, 1)

    $lightStop = New-Object System.Windows.Media.GradientStop
    $lightStop.Color = $lightColor
    $lightStop.Offset = 0
    $darkStop = New-Object System.Windows.Media.GradientStop
    $darkStop.Color = $darkColor
    $darkStop.Offset = 1
    [void]$brush.GradientStops.Add($lightStop)
    [void]$brush.GradientStops.Add($darkStop)

    if ($brush.CanFreeze) { $brush.Freeze() }
    $script:TreemapBrushCache[$category] = $brush
    return $brush
}

function Set-TreemapSelection {
    param(
        $Row,
        [System.Windows.Controls.Border]$Element = $null
    )

    if ($null -ne $script:SelectedTreemapElement) {
        $script:SelectedTreemapElement.BorderBrush = [System.Windows.Media.Brushes]::White
        $script:SelectedTreemapElement.BorderThickness = [System.Windows.Thickness]::new(1)
        [System.Windows.Controls.Panel]::SetZIndex($script:SelectedTreemapElement, 0)
    }

    $script:SelectedTreemapElement = $null

    if ($null -eq $Element -and $null -ne $Row -and $Row.PSObject.Properties['FullName']) {
        $path = [string]$Row.FullName
        if (-not [string]::IsNullOrWhiteSpace($path) -and
            $script:TreemapElementsByPath.ContainsKey($path)) {
            $Element = $script:TreemapElementsByPath[$path]
        }
    }

    if ($null -ne $Element) {
        $Element.BorderBrush = [System.Windows.Media.Brushes]::Gold
        $Element.BorderThickness = [System.Windows.Thickness]::new(3)
        [System.Windows.Controls.Panel]::SetZIndex($Element, 10)
        $script:SelectedTreemapElement = $Element
    }
}

function Select-TreemapRow {
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)][System.Windows.Controls.Border]$Element
    )

    $FileGrid.SelectedItem = $Row
    try { $FileGrid.ScrollIntoView($Row) } catch {}
    Set-TreemapSelection -Row $Row -Element $Element

    if ($Row.PSObject.Properties['FullName']) {
        $StatusText.Text = [string]$Row.FullName
    }
}

function Open-TreemapFile {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if ([System.IO.File]::Exists($Path)) {
            Start-Process -FilePath $Path
        }
        else {
            $StatusText.Text = 'The selected file no longer exists.'
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Could not open file:`n{0}" -f $_.Exception.Message),
            'Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
}

function Show-Treemap {
    $TreemapCanvas.Children.Clear()
    $script:TreemapElementsByPath = @{}
    $script:SelectedTreemapElement = $null

    $width = [double]$TreemapHost.ActualWidth
    $height = [double]$TreemapHost.ActualHeight
    if ($width -lt 20 -or $height -lt 20) {
        $TreemapEmptyText.Text = 'Preparing treemap…'
        $TreemapEmptyText.Visibility = [System.Windows.Visibility]::Visible
        return
    }

    $positiveCount = 0
    $totalBytes = [long]0
    foreach ($row in @($script:CurrentFileRows)) {
        if ($null -ne $row -and [long]$row.Length -gt 0) {
            $positiveCount++
            $totalBytes += [long]$row.Length
        }
    }

    if ($positiveCount -eq 0 -or $totalBytes -le 0) {
        $TreemapInfoText.Text = 'No files larger than 0 bytes.'
        $TreemapEmptyText.Text = 'No files to display.'
        $TreemapEmptyText.Visibility = [System.Windows.Visibility]::Visible
        return
    }

    $limit = Get-TreemapLimit
    $displayRows = @(Get-TreemapDisplayRows -Rows $script:CurrentFileRows -Limit $limit)
    $layout = @(Get-SquarifiedTreemapLayout -Items $displayRows -Width $width -Height $height)

    $individualCount = if ($positiveCount -gt $limit) { $limit - 1 } else { $positiveCount }
    if ($positiveCount -gt $limit) {
        $TreemapInfoText.Text = (
            '{0:N0} files • {1} • {2:N0} largest shown individually, remainder aggregated' -f
                $positiveCount, (Format-Size $totalBytes), $individualCount
        )
    }
    else {
        $TreemapInfoText.Text = (
            '{0:N0} files • {1} • area represents file size' -f
                $positiveCount, (Format-Size $totalBytes)
        )
    }

    foreach ($rectangle in $layout) {
        $row = $rectangle.Item
        $isAggregate = Test-TreemapAggregateRow $row
        $innerWidth = [Math]::Max(0.25, [double]$rectangle.Width - 1.2)
        $innerHeight = [Math]::Max(0.25, [double]$rectangle.Height - 1.2)

        $border = New-Object System.Windows.Controls.Border
        $border.Width = $innerWidth
        $border.Height = $innerHeight
        $border.Background = Get-TreemapBrush $row
        $border.BorderBrush = [System.Windows.Media.Brushes]::White
        $border.BorderThickness = [System.Windows.Thickness]::new(1)
        $border.Opacity = 0.96
        $border.SnapsToDevicePixels = $true

        [System.Windows.Controls.Canvas]::SetLeft($border, [double]$rectangle.X + 0.6)
        [System.Windows.Controls.Canvas]::SetTop($border, [double]$rectangle.Y + 0.6)

        if ($isAggregate) {
            $border.ToolTip = (
                "{0}`n{1}`nAggregate area used to limit rendering" -f
                    $row.Name, $row.SizeText
            )
        }
        else {
            $border.Tag = $row
            $border.Cursor = [System.Windows.Input.Cursors]::Hand
            $border.ToolTip = (
                "{0}`n{1} • {2}`n{3}" -f
                    $row.Name, $row.SizeText, $row.Type, $row.FullName
            )
            [System.Windows.Controls.ToolTipService]::SetInitialShowDelay($border, 120)
            [System.Windows.Controls.ToolTipService]::SetShowDuration($border, 60000)

            $border.Add_MouseLeftButtonDown({
                param($sender, $eventArgs)

                $clickedRow = $sender.Tag
                if ($null -eq $clickedRow) { return }

                Select-TreemapRow -Row $clickedRow -Element $sender
                if ($eventArgs.ClickCount -ge 2 -and $clickedRow.FullName) {
                    Open-TreemapFile -Path ([string]$clickedRow.FullName)
                    $eventArgs.Handled = $true
                }
            })

            $script:TreemapElementsByPath[[string]$row.FullName] = $border
        }

        if ($innerWidth -ge 62 -and $innerHeight -ge 34) {
            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = ("{0}`n{1}" -f $row.Name, $row.SizeText)
            $label.Foreground = [System.Windows.Media.Brushes]::White
            $label.FontSize = 11
            $label.FontWeight = [System.Windows.FontWeights]::SemiBold
            $label.Margin = [System.Windows.Thickness]::new(4, 2, 3, 2)
            $label.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
            $label.IsHitTestVisible = $false
            $border.Child = $label
        }

        $TreemapCanvas.Children.Add($border) | Out-Null
    }

    $TreemapEmptyText.Visibility = [System.Windows.Visibility]::Collapsed
    if ($null -ne $FileGrid.SelectedItem) {
        Set-TreemapSelection -Row $FileGrid.SelectedItem
    }
}

function Request-TreemapRender {
    if ($null -eq $script:TreemapRenderTimer) {
        $script:TreemapRenderTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:TreemapRenderTimer.Interval = [TimeSpan]::FromMilliseconds(140)
        $script:TreemapRenderTimer.Add_Tick({
            $script:TreemapRenderTimer.Stop()
            Show-Treemap
        })
    }

    $script:TreemapRenderTimer.Stop()
    $script:TreemapRenderTimer.Start()
}

function Show-FileRows {
    param(
        [object[]]$Rows,
        [string]$Heading,
        [string]$SubText
    )

    $script:CurrentFileRows = @($Rows)
    $FileGrid.ItemsSource = $null
    $FileGrid.ItemsSource = $script:CurrentFileRows
    $FilesHeading.Text = $Heading
    $FilesPathText.Text = $SubText
    Request-TreemapRender
}

function Show-FilesForFolder {
    param([Parameter(Mandatory)][string]$Path)

    $script:CurrentFilesMode = 'Folder'
    $script:CurrentFolderPath = $Path

    if (-not [System.IO.Directory]::Exists($Path)) {
        Show-FileRows -Rows @() -Heading 'Files' -SubText 'The folder is no longer available.'
        return
    }

    try {
        $dir = [System.IO.DirectoryInfo]::new($Path)
        $rows = @(
            @(
                foreach ($file in $dir.EnumerateFiles('*', [System.IO.SearchOption]::TopDirectoryOnly)) {
                    try {
                        New-FileRow -FullName $file.FullName -Length ([long]$file.Length) -LastWriteTime $file.LastWriteTime
                    }
                    catch {}
                }
            ) | Sort-Object -Property Length -Descending
        )

        Show-FileRows -Rows $rows `
            -Heading ('Files ({0:N0})' -f $rows.Count) `
            -SubText ('{0} — sorted by size descending; only files directly in this folder' -f $Path)
    }
    catch {
        Show-FileRows -Rows @() -Heading 'Files' -SubText ('Cannot read folder: {0}' -f $_.Exception.Message)
    }
}

function Show-GlobalTopFiles {
    $script:CurrentFilesMode = 'Top'
    $script:CurrentFolderPath = $null

    $rows = @(
        foreach ($file in $script:GlobalTopFiles) {
            New-FileRow -FullName ([string]$file.FullName) `
                        -Length ([long]$file.Length) `
                        -LastWriteTime ([datetime]$file.LastWriteTime)
        }
    )

    Show-FileRows -Rows $rows `
        -Heading ('Top {0:N0} largest files' -f $rows.Count) `
        -SubText 'Entire scanned area — sorted by size descending'
}

function Get-SelectedFileRow {
    $row = $FileGrid.SelectedItem
    if ($null -eq $row) { return $null }
    if (-not $row.PSObject.Properties['FullName']) { return $null }
    return $row
}

function Get-SelectedFolderPath {
    $item = $FolderTree.SelectedItem
    if ($item -is [System.Windows.Controls.TreeViewItem] -and
        $item.Tag -and $item.Tag -ne '__PLACEHOLDER__') {
        try {
            return [string](Get-NodeProperty $item.Tag 'Path')
        }
        catch {}
    }
    return $null
}

function Open-FolderInExplorer {
    param([Parameter(Mandatory)][string]$Path)

    if ([System.IO.Directory]::Exists($Path)) {
        Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $Path)
    }
}

function Show-FileInExplorer {
    param([Parameter(Mandatory)][string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        Start-Process -FilePath explorer.exe -ArgumentList ('/select,"{0}"' -f $Path)
    }
}

function Copy-PathToClipboard {
    param([Parameter(Mandatory)][string]$Path)

    try {
        [System.Windows.Clipboard]::SetText($Path)
        $StatusText.Text = 'Path copied to the clipboard.'
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Could not copy path:`n{0}" -f $_.Exception.Message),
            'Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
}

function Refresh-CurrentFileView {
    if ($script:CurrentFilesMode -eq 'Folder' -and
        -not [string]::IsNullOrWhiteSpace([string]$script:CurrentFolderPath)) {
        Show-FilesForFolder -Path ([string]$script:CurrentFolderPath)
    }
    else {
        Show-GlobalTopFiles
    }
}

function Remove-FileToRecycleBin {
    param([Parameter(Mandatory)][string]$Path)

    if (-not [System.IO.File]::Exists($Path)) {
        [System.Windows.MessageBox]::Show(
            $window,
            'The file no longer exists.',
            'File Not Found',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        Refresh-CurrentFileView
        return
    }

    try {
        # Use the Recycle Bin for standard Windows deletion behavior. AllDialogs lets
        # Windows display confirmation and error dialogs.
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $Path,
            [Microsoft.VisualBasic.FileIO.UIOption]::AllDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin,
            [Microsoft.VisualBasic.FileIO.UICancelOption]::DoNothing
        )

        if (-not [System.IO.File]::Exists($Path)) {
            $script:GlobalTopFiles = @(
                $script:GlobalTopFiles | Where-Object {
                    [string]$_.FullName -ne $Path
                }
            )

            Refresh-CurrentFileView
            $StatusText.Text = 'File moved to the Recycle Bin.'
            $SummaryText.Text = 'Note: The folder sizes on the left still reflect the previous scan. Scan again to update them.'
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Could not delete file:`n{0}" -f $_.Exception.Message),
            'Deletion Failed',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
}

function Get-WindowsShellVerbs {
    param([Parameter(Mandatory)][string]$Path)

    $result = New-Object 'System.Collections.Generic.List[object]'

    try {
        $directory = [System.IO.Path]::GetDirectoryName($Path)
        $leaf = [System.IO.Path]::GetFileName($Path)

        if ([string]::IsNullOrWhiteSpace($directory) -or
            [string]::IsNullOrWhiteSpace($leaf)) {
            return $result.ToArray()
        }

        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($directory)
        if ($null -eq $folder) { return $result.ToArray() }

        $item = $folder.ParseName($leaf)
        if ($null -eq $item) { return $result.ToArray() }

        foreach ($verb in $item.Verbs()) {
            $rawName = [string]$verb.Name
            $displayName = ($rawName -replace '&', '').Trim()

            if (-not [string]::IsNullOrWhiteSpace($displayName)) {
                $result.Add([PSCustomObject]@{
                    DisplayName = $displayName
                    VerbName    = $rawName
                }) | Out-Null
            }
        }
    }
    catch {
        # Convenience behavior: the core menu remains usable even without shell verbs.
    }

    return $result.ToArray()
}

function Invoke-WindowsShellVerb {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$VerbName
    )

    try {
        if (-not [System.IO.File]::Exists($Path)) { return }

        $directory = [System.IO.Path]::GetDirectoryName($Path)
        $leaf = [System.IO.Path]::GetFileName($Path)

        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace($directory)
        if ($null -eq $folder) { return }

        $item = $folder.ParseName($leaf)
        if ($null -eq $item) { return }

        foreach ($verb in $item.Verbs()) {
            if ([string]$verb.Name -eq $VerbName) {
                $verb.DoIt()
                return
            }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Could not execute Windows action:`n{0}" -f $_.Exception.Message),
            'Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
}

function Select-DataGridRowUnderMouse {
    param([Parameter(Mandatory)]$OriginalSource)

    $element = $OriginalSource -as [System.Windows.DependencyObject]
    while ($null -ne $element -and
           -not ($element -is [System.Windows.Controls.DataGridRow])) {
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }

    if ($element -is [System.Windows.Controls.DataGridRow]) {
        $FileGrid.SelectedItem = $element.Item
        $element.IsSelected = $true
        [void]$element.Focus()
    }
    else {
        # Important: right-clicking empty space must not accidentally retain the
        # previously selected file as the deletion target.
        $FileGrid.SelectedItem = $null
    }
}

function Select-TreeItemUnderMouse {
    param([Parameter(Mandatory)]$OriginalSource)

    $element = $OriginalSource -as [System.Windows.DependencyObject]
    while ($null -ne $element -and
           -not ($element -is [System.Windows.Controls.TreeViewItem])) {
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }

    if ($element -is [System.Windows.Controls.TreeViewItem]) {
        $element.IsSelected = $true
        [void]$element.Focus()
    }
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
    $grid.MinWidth = 480

    $c1 = New-Object System.Windows.Controls.ColumnDefinition
    $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $c2 = New-Object System.Windows.Controls.ColumnDefinition
    $c2.Width = [System.Windows.GridLength]::new(110)
    $c3 = New-Object System.Windows.Controls.ColumnDefinition
    $c3.Width = [System.Windows.GridLength]::new(95)

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
    $bar.ToolTip = ('{0:N1}% of parent folder' -f $percent)

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


    $TopLimit = 200
    $TopFilesBySize = [System.Collections.Generic.SortedDictionary[long,object]]::new()
    $script:TopFileCount = 0

    function Get-SmallestTopSize {
        if ($TopFilesBySize.Count -eq 0) { return [long]0 }
        $enumerator = $TopFilesBySize.Keys.GetEnumerator()
        try {
            if ($enumerator.MoveNext()) { return [long]$enumerator.Current }
        }
        finally {
            if ($enumerator -is [System.IDisposable]) { $enumerator.Dispose() }
        }
        return [long]0
    }

    function Add-TopFile {
        param([System.IO.FileInfo]$File, [long]$Length)

        $smallest = Get-SmallestTopSize
        if ($script:TopFileCount -ge $TopLimit -and $Length -le $smallest) { return }

        $entry = [PSCustomObject]@{
            Name          = $File.Name
            FullName      = $File.FullName
            Directory     = $File.DirectoryName
            Length        = $Length
            Extension     = $File.Extension
            LastWriteTime = $File.LastWriteTime
        }

        if (-not $TopFilesBySize.ContainsKey($Length)) {
            $TopFilesBySize.Add($Length, [System.Collections.Generic.Queue[object]]::new())
        }
        $TopFilesBySize[$Length].Enqueue($entry)
        $script:TopFileCount++

        while ($script:TopFileCount -gt $TopLimit) {
            $minSize = Get-SmallestTopSize
            $queue = $TopFilesBySize[$minSize]
            [void]$queue.Dequeue()
            $script:TopFileCount--
            if ($queue.Count -eq 0) {
                [void]$TopFilesBySize.Remove($minSize)
            }
        }
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
                    Add-TopFile -File $file -Length $length

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
            $topFiles = New-Object 'System.Collections.Generic.List[object]'
            foreach ($size in @($TopFilesBySize.Keys | Sort-Object -Descending)) {
                foreach ($entry in $TopFilesBySize[$size]) {
                    $topFiles.Add($entry) | Out-Null
                }
            }

            [PSCustomObject]@{
                Root     = $root
                # PowerShell 5.1 can fail on @($genericList) with a
                # PSToObjectArrayBinder error. ToArray() works around it.
                TopFiles = $topFiles.ToArray()
            }
        }
    }
    catch {
        $lineNumber = $_.InvocationInfo.ScriptLineNumber
        $lineText = $_.InvocationInfo.Line
        $State.FatalError = (
            "{0}`n`nPowerShell line: {1}`n{2}" -f
                $_.Exception.ToString(), $lineNumber, $lineText
        )
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
$script:GlobalTopFiles = @()
$script:CurrentFilesMode = 'Top'
$script:CurrentFolderPath = $null

function Set-ScanningUi {
    param([bool]$Scanning)

    $BrowseButton.IsEnabled = -not $Scanning
    $ScanButton.IsEnabled = -not $Scanning
    $PathBox.IsEnabled = -not $Scanning
    $CancelButton.IsEnabled = $Scanning
    $ShowTopFilesButton.IsEnabled = -not $Scanning

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
        $scanResult = $results | Select-Object -First 1
        $root = if ($null -ne $scanResult) { $scanResult.Root } else { $null }
        $script:GlobalTopFiles = if ($null -ne $scanResult) { $scanResult.TopFiles } else { @() }

        $FolderTree.Items.Clear()
        $FileGrid.ItemsSource = $null

        if ($null -ne $root) {
            $rootSize = [long](Get-NodeProperty $root 'Size')
            $rootItem = New-TreeItem -Node $root -ParentSize $rootSize

            # Expand the root immediately; deeper levels remain lazy-loaded.
            Add-LazyChildren -Item $rootItem

            [void]$FolderTree.Items.Add($rootItem)
            $rootItem.IsExpanded = $true

            $cancelled = [bool]$script:ScanState.Cancel
            if ($cancelled) {
                $StatusText.Text = 'Scan cancelled — displaying partial results.'
            }
            else {
                $StatusText.Text = 'Scan complete.'
            }

            $SummaryText.Text = (
                '{0} • {1:N0} files • {2:N0} subfolders • {3:N0} skipped • {4:N0} reparse points skipped' -f
                    (Format-Size $rootSize),
                    [long](Get-NodeProperty $root 'TotalFiles'),
                    [long](Get-NodeProperty $root 'TotalFolders'),
                    [long]$script:ScanState.Skipped,
                    [long]$script:ScanState.ReparseSkipped
            )


            Show-GlobalTopFiles
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
            'Invalid Folder',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }

    Stop-ScanResources
    $FolderTree.Items.Clear()
    $FileGrid.ItemsSource = $null
    $script:CurrentFileRows = @()
    $TreemapCanvas.Children.Clear()
    $script:TreemapElementsByPath = @{}
    $script:SelectedTreemapElement = $null
    $TreemapInfoText.Text = 'Scan in progress…'
    $TreemapEmptyText.Text = 'The treemap will be built after the scan.'
    $TreemapEmptyText.Visibility = [System.Windows.Visibility]::Visible
    $script:GlobalTopFiles = @()
    $script:CurrentFilesMode = 'Top'
    $script:CurrentFolderPath = $null
    $FilesHeading.Text = 'Largest Files'
    $FilesPathText.Text = 'Scan in progress…'
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
    $dialog.Description = 'Select a folder for storage analysis'
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

$TreemapHost.Add_SizeChanged({
    Request-TreemapRender
})

$TreemapLimitBox.Add_SelectionChanged({
    Request-TreemapRender
})

$FileViews.Add_SelectionChanged({
    param($sender, $eventArgs)

    # SelectionChanged is a routed event and also reaches the TabControl from the
    # DataGrid/ComboBox. Handle only an actual tab change.
    if ($eventArgs.OriginalSource -eq $FileViews -and
        $FileViews.SelectedItem -eq $TreemapTab) {
        Request-TreemapRender
    }
})

$FileGrid.Add_SelectionChanged({
    Set-TreemapSelection -Row $FileGrid.SelectedItem
})

$FolderTree.Add_SelectedItemChanged({
    param($sender, $eventArgs)

    $item = $FolderTree.SelectedItem
    if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Tag -and $item.Tag -ne '__PLACEHOLDER__') {
        try {
            $path = [string](Get-NodeProperty $item.Tag 'Path')
            Show-FilesForFolder -Path $path
        }
        catch {}
    }
})

$ShowTopFilesButton.Add_Click({
    Show-GlobalTopFiles
})

$FileGrid.Add_MouseDoubleClick({
    $row = $FileGrid.SelectedItem
    if ($null -ne $row -and $row.FullName) {
        try {
            if ([System.IO.File]::Exists([string]$row.FullName)) {
                Start-Process -FilePath ([string]$row.FullName)
            }
        }
        catch {
            [System.Windows.MessageBox]::Show(
                $window,
                ("Could not open file:`n{0}" -f $_.Exception.Message),
                'Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
        }
    }
})

$FolderTree.Add_MouseDoubleClick({
    param($sender, $eventArgs)

    $item = $FolderTree.SelectedItem
    if ($item -is [System.Windows.Controls.TreeViewItem] -and $item.Tag -and $item.Tag -ne '__PLACEHOLDER__') {
        try {
            $path = [string](Get-NodeProperty $item.Tag 'Path')
            Open-FolderInExplorer -Path $path
        }
        catch {}
    }
})

# ------------------------------------------------------------
# Explorer-like context menus
# ------------------------------------------------------------

$fileContextMenu = New-Object System.Windows.Controls.ContextMenu

$fileOpenMenuItem = New-Object System.Windows.Controls.MenuItem
$fileOpenMenuItem.Header = 'Open'
[void]$fileContextMenu.Items.Add($fileOpenMenuItem)

$fileShowExplorerMenuItem = New-Object System.Windows.Controls.MenuItem
$fileShowExplorerMenuItem.Header = 'Show in Explorer'
[void]$fileContextMenu.Items.Add($fileShowExplorerMenuItem)

$fileCopyPathMenuItem = New-Object System.Windows.Controls.MenuItem
$fileCopyPathMenuItem.Header = 'Copy Path'
[void]$fileContextMenu.Items.Add($fileCopyPathMenuItem)

[void]$fileContextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

$fileDeleteMenuItem = New-Object System.Windows.Controls.MenuItem
$fileDeleteMenuItem.Header = 'Delete (Recycle Bin)'
[void]$fileContextMenu.Items.Add($fileDeleteMenuItem)

[void]$fileContextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

$fileWindowsActionsMenuItem = New-Object System.Windows.Controls.MenuItem
$fileWindowsActionsMenuItem.Header = 'More Windows Actions'
[void]$fileContextMenu.Items.Add($fileWindowsActionsMenuItem)

$FileGrid.ContextMenu = $fileContextMenu

$folderContextMenu = New-Object System.Windows.Controls.ContextMenu

$folderOpenExplorerMenuItem = New-Object System.Windows.Controls.MenuItem
$folderOpenExplorerMenuItem.Header = 'Open in Explorer'
[void]$folderContextMenu.Items.Add($folderOpenExplorerMenuItem)

$folderCopyPathMenuItem = New-Object System.Windows.Controls.MenuItem
$folderCopyPathMenuItem.Header = 'Copy Path'
[void]$folderContextMenu.Items.Add($folderCopyPathMenuItem)

$FolderTree.ContextMenu = $folderContextMenu

# A right-click first selects the item directly under the pointer.
$FileGrid.Add_PreviewMouseRightButtonDown({
    param($sender, $eventArgs)
    Select-DataGridRowUnderMouse -OriginalSource $eventArgs.OriginalSource
})

$FolderTree.Add_PreviewMouseRightButtonDown({
    param($sender, $eventArgs)
    Select-TreeItemUnderMouse -OriginalSource $eventArgs.OriginalSource
})

$FileGrid.Add_PreviewKeyDown({
    param($sender, $eventArgs)

    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Delete) {
        $row = Get-SelectedFileRow
        if ($null -ne $row -and $row.FullName) {
            Remove-FileToRecycleBin -Path ([string]$row.FullName)
            $eventArgs.Handled = $true
        }
    }
})

$fileContextMenu.Add_Opened({
    $row = Get-SelectedFileRow
    $hasFile = ($null -ne $row -and
                $row.FullName -and
                [System.IO.File]::Exists([string]$row.FullName))

    $fileOpenMenuItem.IsEnabled = $hasFile
    $fileShowExplorerMenuItem.IsEnabled = $hasFile
    $fileCopyPathMenuItem.IsEnabled = ($null -ne $row -and $row.FullName)
    $fileDeleteMenuItem.IsEnabled = $hasFile

    $fileWindowsActionsMenuItem.Items.Clear()

    if ($hasFile) {
        $verbs = @(Get-WindowsShellVerbs -Path ([string]$row.FullName))

        foreach ($verb in $verbs) {
            $menuItem = New-Object System.Windows.Controls.MenuItem
            $menuItem.Header = [string]$verb.DisplayName
            $menuItem.Tag = [string]$verb.VerbName
            $menuItem.Add_Click({
                param($sender, $eventArgs)

                $selected = Get-SelectedFileRow
                if ($null -ne $selected -and $selected.FullName) {
                    Invoke-WindowsShellVerb `
                        -Path ([string]$selected.FullName) `
                        -VerbName ([string]$sender.Tag)
                }
            })
            [void]$fileWindowsActionsMenuItem.Items.Add($menuItem)
        }
    }

    $fileWindowsActionsMenuItem.IsEnabled = ($fileWindowsActionsMenuItem.Items.Count -gt 0)
})

$folderContextMenu.Add_Opened({
    $path = Get-SelectedFolderPath
    $exists = (-not [string]::IsNullOrWhiteSpace([string]$path) -and
               [System.IO.Directory]::Exists([string]$path))
    $folderOpenExplorerMenuItem.IsEnabled = $exists
    $folderCopyPathMenuItem.IsEnabled = (-not [string]::IsNullOrWhiteSpace([string]$path))
})

$fileOpenMenuItem.Add_Click({
    $row = Get-SelectedFileRow
    if ($null -ne $row -and $row.FullName) {
        try {
            Start-Process -FilePath ([string]$row.FullName)
        }
        catch {
            [System.Windows.MessageBox]::Show(
                $window,
                ("Could not open file:`n{0}" -f $_.Exception.Message),
                'Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
        }
    }
})

$fileShowExplorerMenuItem.Add_Click({
    $row = Get-SelectedFileRow
    if ($null -ne $row -and $row.FullName) {
        Show-FileInExplorer -Path ([string]$row.FullName)
    }
})

$fileCopyPathMenuItem.Add_Click({
    $row = Get-SelectedFileRow
    if ($null -ne $row -and $row.FullName) {
        Copy-PathToClipboard -Path ([string]$row.FullName)
    }
})

$fileDeleteMenuItem.Add_Click({
    $row = Get-SelectedFileRow
    if ($null -ne $row -and $row.FullName) {
        Remove-FileToRecycleBin -Path ([string]$row.FullName)
    }
})

$folderOpenExplorerMenuItem.Add_Click({
    $path = Get-SelectedFolderPath
    if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
        Open-FolderInExplorer -Path ([string]$path)
    }
})

$folderCopyPathMenuItem.Add_Click({
    $path = Get-SelectedFolderPath
    if (-not [string]::IsNullOrWhiteSpace([string]$path)) {
        Copy-PathToClipboard -Path ([string]$path)
    }
})

$window.Add_Closing({
    if ($script:TreemapRenderTimer) {
        $script:TreemapRenderTimer.Stop()
    }

    if ($script:ScanState) {
        $script:ScanState.Cancel = $true
    }

    if ($script:ScanPowerShell -and $script:ScanAsync -and -not $script:ScanAsync.IsCompleted) {
        try { $script:ScanPowerShell.Stop() } catch {}
    }

    Stop-ScanResources
})

[void]$window.ShowDialog()
