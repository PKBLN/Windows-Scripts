#requires -Version 5.1
<#
.SYNOPSIS
    Einfacher WPF-Speicherplatz-Analyzer ("TreeSize Light") in einer einzelnen PowerShell-Datei.

.DESCRIPTION
    - Ordnerauswahl
    - Scan im Hintergrund-Runspace, damit die WPF-Oberfläche responsiv bleibt
    - Abbrechen möglich
    - Baumansicht der Unterordner
    - Größenbalken pro Ordner relativ zum jeweiligen Elternordner
    - Explorer-artige Dateiansicht neben dem Ordnerbaum
    - Top 200 der größten Dateien im gesamten Scan
    - Rechtsklick auf Dateien: Öffnen, Explorer, Pfad kopieren, Papierkorb
    - Zusätzliche Windows-Shell-Aktionen als Untermenü
    - Live-Status mit gescannten Dateien, Bytes, aktuellem Pfad und übersprungenen Einträgen
    - Reparse Points/Junctions werden ausgelassen, um Schleifen/Doppelzählungen zu vermeiden

    Hinweis:
    Ein exakter Prozentfortschritt des Scans ist ohne vorgeschalteten vollständigen Zähllauf
    nicht sinnvoll möglich. Daher ist der globale Scanbalken während des Scans indeterminiert.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# WPF benötigt einen STA-Thread.
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
    throw 'Dieses Skript muss in einem STA-Thread laufen. Starte es z. B. mit powershell.exe -STA -File .\StorageTree.ps1'
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
    Title="StorageTree 2.2 - PowerShell Speicheranalyse"
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
                           Text="Ordner:"
                           VerticalAlignment="Center"
                           Margin="0,0,8,0"/>

                <TextBox x:Name="PathBox"
                         Grid.Column="1"
                         MinWidth="300"
                         VerticalContentAlignment="Center"
                         Margin="0,0,8,0"/>

                <Button x:Name="BrowseButton"
                        Grid.Column="2"
                        Content="Durchsuchen…"
                        Padding="12,5"
                        Margin="0,0,8,0"/>

                <Button x:Name="ScanButton"
                        Grid.Column="3"
                        Content="Scannen"
                        Padding="16,5"
                        Margin="0,0,8,0"/>

                <Button x:Name="CancelButton"
                        Grid.Column="4"
                        Content="Abbrechen"
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

                <!-- Ordnerbaum -->
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

                        <TextBlock Text="Ordner" FontWeight="SemiBold"/>
                        <TextBlock Grid.Column="1" Text="Größe" FontWeight="SemiBold"
                                   HorizontalAlignment="Right" Margin="0,0,12,0"/>
                        <TextBlock Grid.Column="2" Text="Anteil" FontWeight="SemiBold"
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

                <!-- Dateiansicht -->
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
                        <TextBlock x:Name="FilesHeading" Text="Größte Dateien"
                                   FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <Button x:Name="ShowTopFilesButton" Grid.Column="1"
                                Content="Top 200 im Scan" Padding="9,3"/>
                    </Grid>

                    <TextBlock x:Name="FilesPathText" Grid.Row="1" Margin="10,0,8,5"
                               Foreground="#666666" Text="Nach dem Scan werden hier die größten Dateien angezeigt."
                               TextTrimming="CharacterEllipsis"/>

                    <DataGrid x:Name="FileGrid" Grid.Row="2" Margin="6"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              CanUserAddRows="False" CanUserDeleteRows="False"
                              CanUserReorderColumns="True" CanUserResizeColumns="True"
                              CanUserSortColumns="False" SelectionMode="Single"
                              HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                              AlternatingRowBackground="#FAFAFA">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="Größe" Binding="{Binding SizeText}" Width="90"/>
                            <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="85"/>
                            <DataGridTextColumn Header="Geändert" Binding="{Binding ModifiedText}" Width="125"/>
                            <DataGridTextColumn Header="Pfad" Binding="{Binding Directory}" Width="2*"/>
                        </DataGrid.Columns>
                    </DataGrid>
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
                           Text="Bereit."
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
$FilesHeading      = $window.FindName('FilesHeading')
$FilesPathText     = $window.FindName('FilesPathText')
$ShowTopFilesButton = $window.FindName('ShowTopFilesButton')
$ScanProgress      = $window.FindName('ScanProgress')
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
        'Datei'
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

function Show-FileRows {
    param(
        [object[]]$Rows,
        [string]$Heading,
        [string]$SubText
    )

    $FileGrid.ItemsSource = $null
    $FileGrid.ItemsSource = @($Rows)
    $FilesHeading.Text = $Heading
    $FilesPathText.Text = $SubText
}

function Show-FilesForFolder {
    param([Parameter(Mandatory)][string]$Path)

    $script:CurrentFilesMode = 'Folder'
    $script:CurrentFolderPath = $Path

    if (-not [System.IO.Directory]::Exists($Path)) {
        Show-FileRows -Rows @() -Heading 'Dateien' -SubText 'Ordner ist nicht mehr verfügbar.'
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
            -Heading ('Dateien ({0:N0})' -f $rows.Count) `
            -SubText ('{0} — nach Größe absteigend, nur Dateien direkt in diesem Ordner' -f $Path)
    }
    catch {
        Show-FileRows -Rows @() -Heading 'Dateien' -SubText ('Kann Ordner nicht lesen: {0}' -f $_.Exception.Message)
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
        -Heading ('Top {0:N0} größte Dateien' -f $rows.Count) `
        -SubText 'Gesamter gescannter Bereich — nach Größe absteigend'
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
        $StatusText.Text = 'Pfad in die Zwischenablage kopiert.'
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Pfad konnte nicht kopiert werden:`n{0}" -f $_.Exception.Message),
            'Fehler',
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
            'Die Datei existiert nicht mehr.',
            'Datei nicht gefunden',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        Refresh-CurrentFileView
        return
    }

    try {
        # Windows-typisches Löschen über den Papierkorb. AllDialogs lässt Windows
        # Bestätigungen und Fehlerdialoge anzeigen.
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
            $StatusText.Text = 'Datei in den Papierkorb verschoben.'
            $SummaryText.Text = 'Hinweis: Die Ordnergrößen links entsprechen noch dem letzten Scan. Für aktuelle Werte bitte neu scannen.'
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            $window,
            ("Datei konnte nicht gelöscht werden:`n{0}" -f $_.Exception.Message),
            'Löschen fehlgeschlagen',
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
        # Komfortfunktion: Das Kernmenü bleibt auch ohne Shell-Verben nutzbar.
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
            ("Windows-Aktion konnte nicht ausgeführt werden:`n{0}" -f $_.Exception.Message),
            'Fehler',
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
        # Wichtig: Rechtsklick auf Leerraum darf nicht versehentlich die zuvor
        # ausgewählte Datei als Ziel für "Löschen" behalten.
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
    $bar.ToolTip = ('{0:N1} % des Elternordners' -f $percent)

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

        # Dateien direkt in diesem Ordner
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

        # Unterordner
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
                # PowerShell 5.1 kann bei @($genericList) mit einem
                # PSToObjectArrayBinder-Fehler abbrechen. ToArray() umgeht das.
                TopFiles = $topFiles.ToArray()
            }
        }
    }
    catch {
        $lineNumber = $_.InvocationInfo.ScriptLineNumber
        $lineText = $_.InvocationInfo.Line
        $State.FatalError = (
            "{0}`n`nPowerShell-Zeile: {1}`n{2}" -f
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

            # Root direkt auffächern, die tieferen Ebenen bleiben lazy.
            Add-LazyChildren -Item $rootItem

            [void]$FolderTree.Items.Add($rootItem)
            $rootItem.IsExpanded = $true

            $cancelled = [bool]$script:ScanState.Cancel
            if ($cancelled) {
                $StatusText.Text = 'Scan abgebrochen – Teilergebnis wird angezeigt.'
            }
            else {
                $StatusText.Text = 'Scan abgeschlossen.'
            }

            $SummaryText.Text = (
                '{0} • {1:N0} Dateien • {2:N0} Unterordner • {3:N0} übersprungen • {4:N0} Reparse Points ausgelassen' -f
                    (Format-Size $rootSize),
                    [long](Get-NodeProperty $root 'TotalFiles'),
                    [long](Get-NodeProperty $root 'TotalFolders'),
                    [long]$script:ScanState.Skipped,
                    [long]$script:ScanState.ReparseSkipped
            )


            Show-GlobalTopFiles
        }
        elseif ($script:ScanState.FatalError) {
            $StatusText.Text = 'Scan fehlgeschlagen.'
            $SummaryText.Text = [string]$script:ScanState.FatalError
        }
        else {
            $StatusText.Text = 'Keine Daten gefunden.'
            $SummaryText.Text = ''
        }
    }
    catch {
        $StatusText.Text = 'Fehler beim Abschließen des Scans.'
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
            'Der angegebene Ordner existiert nicht.',
            'Ungültiger Ordner',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }

    Stop-ScanResources
    $FolderTree.Items.Clear()
    $FileGrid.ItemsSource = $null
    $script:GlobalTopFiles = @()
    $script:CurrentFilesMode = 'Top'
    $script:CurrentFolderPath = $null
    $FilesHeading.Text = 'Größte Dateien'
    $FilesPathText.Text = 'Scan läuft…'
    $SummaryText.Text = ''
    $StatusText.Text = 'Scan wird gestartet…'

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
            'Scanne… {0:N0} Dateien, {1:N0} Ordner, {2} gelesen, {3:N0} übersprungen — {4}' -f
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
    $dialog.Description = 'Ordner für die Speicheranalyse auswählen'
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
        $StatusText.Text = 'Abbruch angefordert…'
    }
})

$PathBox.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Key -eq [System.Windows.Input.Key]::Enter -and $ScanButton.IsEnabled) {
        Start-Scan
    }
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
                ("Datei konnte nicht geöffnet werden:`n{0}" -f $_.Exception.Message),
                'Fehler',
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
# Explorer-artige Kontextmenüs
# ------------------------------------------------------------

$fileContextMenu = New-Object System.Windows.Controls.ContextMenu

$fileOpenMenuItem = New-Object System.Windows.Controls.MenuItem
$fileOpenMenuItem.Header = 'Öffnen'
[void]$fileContextMenu.Items.Add($fileOpenMenuItem)

$fileShowExplorerMenuItem = New-Object System.Windows.Controls.MenuItem
$fileShowExplorerMenuItem.Header = 'Im Explorer anzeigen'
[void]$fileContextMenu.Items.Add($fileShowExplorerMenuItem)

$fileCopyPathMenuItem = New-Object System.Windows.Controls.MenuItem
$fileCopyPathMenuItem.Header = 'Pfad kopieren'
[void]$fileContextMenu.Items.Add($fileCopyPathMenuItem)

[void]$fileContextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

$fileDeleteMenuItem = New-Object System.Windows.Controls.MenuItem
$fileDeleteMenuItem.Header = 'Löschen (Papierkorb)'
[void]$fileContextMenu.Items.Add($fileDeleteMenuItem)

[void]$fileContextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

$fileWindowsActionsMenuItem = New-Object System.Windows.Controls.MenuItem
$fileWindowsActionsMenuItem.Header = 'Weitere Windows-Aktionen'
[void]$fileContextMenu.Items.Add($fileWindowsActionsMenuItem)

$FileGrid.ContextMenu = $fileContextMenu

$folderContextMenu = New-Object System.Windows.Controls.ContextMenu

$folderOpenExplorerMenuItem = New-Object System.Windows.Controls.MenuItem
$folderOpenExplorerMenuItem.Header = 'Im Explorer öffnen'
[void]$folderContextMenu.Items.Add($folderOpenExplorerMenuItem)

$folderCopyPathMenuItem = New-Object System.Windows.Controls.MenuItem
$folderCopyPathMenuItem.Header = 'Pfad kopieren'
[void]$folderContextMenu.Items.Add($folderCopyPathMenuItem)

$FolderTree.ContextMenu = $folderContextMenu

# Rechtsklick wählt zuerst genau das Element unter dem Mauszeiger aus.
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
                ("Datei konnte nicht geöffnet werden:`n{0}" -f $_.Exception.Message),
                'Fehler',
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
    if ($script:ScanState) {
        $script:ScanState.Cancel = $true
    }

    if ($script:ScanPowerShell -and $script:ScanAsync -and -not $script:ScanAsync.IsCompleted) {
        try { $script:ScanPowerShell.Stop() } catch {}
    }

    Stop-ScanResources
})

[void]$window.ShowDialog()
