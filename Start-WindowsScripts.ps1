#requires -Version 5.1
<#
.SYNOPSIS
    Bilingual launcher for the Windows-Scripts collection.

.DESCRIPTION
    Lets the user choose a tool and its language, removes the Internet zone
    marker from the selected script with Unblock-File, and launches it in a
    separate Windows PowerShell process.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:CurrentEntries = @()

$toolDefinitions = @(
    [PSCustomObject]@{
        Key                = 'WinDirStat3'
        GermanScript       = 'WinDirStat3.ps1'
        EnglishScript      = 'WinDirStat3.en.ps1'
        GermanName         = 'WinDirStat 3 – Treemap'
        EnglishName        = 'WinDirStat 3 – Treemap'
        GermanDescription  = 'Die umfangreichste Speicheranalyse: Ordnerbaum, Dateiliste, Explorer-Aktionen und eine SequoiaView-artige Treemap.'
        EnglishDescription = 'The most comprehensive storage analyzer: folder tree, file list, Explorer actions, and a SequoiaView-style treemap.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'WinDirStat2'
        GermanScript       = 'WinDirStat2.ps1'
        EnglishScript      = 'WinDirStat2.en.ps1'
        GermanName         = 'WinDirStat 2 – Dateiaktionen'
        EnglishName        = 'WinDirStat 2 – File actions'
        GermanDescription  = 'Speicheranalyse mit Ordnerbaum, Dateiliste, Top 200, Kontextmenüs und Löschen über den Papierkorb.'
        EnglishDescription = 'Storage analyzer with folder tree, file list, top 200, context menus, and Recycle Bin deletion.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'WinDirStat'
        GermanScript       = 'WinDirStat.ps1'
        EnglishScript      = 'WinDirStat.en.ps1'
        GermanName         = 'WinDirStat 1 – Dateiansicht'
        EnglishName        = 'WinDirStat 1 – File view'
        GermanDescription  = 'Speicheranalyse mit Ordnerbaum, direkter Dateiansicht und den 200 größten Dateien des Scans.'
        EnglishDescription = 'Storage analyzer with a folder tree, direct file view, and the 200 largest files in the scan.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'StorageTree'
        GermanScript       = 'StorageTree.ps1'
        EnglishScript      = 'StorageTree.en.ps1'
        GermanName         = 'StorageTree – kompakte Analyse'
        EnglishName        = 'StorageTree – compact analyzer'
        GermanDescription  = 'Schlanke grafische Speicheranalyse mit Größenbaum und Live-Status.'
        EnglishDescription = 'Compact graphical storage analyzer with a size tree and live status.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'Search'
        GermanScript       = 'suche.ps1'
        EnglishScript      = 'suche.en.ps1'
        GermanName         = 'Datei-Inhalt-Suche'
        EnglishName        = 'File content search'
        GermanDescription  = 'Durchsucht passende Dateien unterhalb des Skriptordners zeilenweise nach einem Text.'
        EnglishDescription = 'Searches matching files below the script folder for literal text, line by line.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'CdxSearch'
        GermanScript       = 'CdxSearchGui.ps1'
        EnglishScript      = 'CdxSearchGui.en.ps1'
        GermanName         = 'Archive.org CDX-Suche'
        EnglishName        = 'Archive.org CDX search'
        GermanDescription  = 'Sucht über den Archive.org-CDX-Dienst nach archivierten Dateien einer Domain und erzeugt anklickbare Wayback-Links.'
        EnglishDescription = 'Searches the Archive.org CDX service for archived files from a domain and produces clickable Wayback links.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'EncodingDoctor'
        GermanScript       = 'Encoding-Doctor.de.ps1'
        EnglishScript      = 'Encoding-Doctor.ps1'
        GermanName         = 'Encoding Doctor – Kodierung prüfen'
        EnglishName        = 'Encoding Doctor – check encoding'
        GermanDescription  = 'Prüft Textdateien rekursiv, ergänzt für PowerShell bei Bedarf einen UTF-8-BOM und kann typische Mojibake-Schäden mit Sicherung reparieren.'
        EnglishDescription = 'Recursively checks text files, adds a UTF-8 BOM to PowerShell sources when needed, and can repair common mojibake damage with backups.'
        Console             = $false
    }
    [PSCustomObject]@{
        Key                = 'ProfileSize'
        GermanScript       = 'ProfileSize.ps1'
        EnglishScript      = 'ProfileSize.en.ps1'
        GermanName         = 'Profilordner-Größen'
        EnglishName        = 'Profile folder sizes'
        GermanDescription  = 'Gibt die Größen der direkten Unterordner des aktuellen Benutzerprofils in einer Konsole aus.'
        EnglishDescription = "Prints the sizes of the current user profile's direct subfolders in a console."
        Console             = $true
    }
)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Windows Scripts – Starter / Launcher'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(820, 510)
$form.MinimumSize = New-Object System.Drawing.Size(850, 540)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(18, 16)
$titleLabel.Size = New-Object System.Drawing.Size(520, 32)
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
$form.Controls.Add($titleLabel)

$languageLabel = New-Object System.Windows.Forms.Label
$languageLabel.Location = New-Object System.Drawing.Point(580, 17)
$languageLabel.Size = New-Object System.Drawing.Size(90, 22)
$languageLabel.TextAlign = 'MiddleRight'
$languageLabel.Anchor = 'Top,Right'
$form.Controls.Add($languageLabel)

$languageBox = New-Object System.Windows.Forms.ComboBox
$languageBox.Location = New-Object System.Drawing.Point(678, 16)
$languageBox.Size = New-Object System.Drawing.Size(124, 25)
$languageBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$languageBox.Anchor = 'Top,Right'
[void]$languageBox.Items.Add('Deutsch')
[void]$languageBox.Items.Add('English')
$form.Controls.Add($languageBox)

$introLabel = New-Object System.Windows.Forms.Label
$introLabel.Location = New-Object System.Drawing.Point(20, 55)
$introLabel.Size = New-Object System.Drawing.Size(780, 38)
$introLabel.Anchor = 'Top,Left,Right'
$form.Controls.Add($introLabel)

$toolList = New-Object System.Windows.Forms.ListBox
$toolList.Location = New-Object System.Drawing.Point(20, 103)
$toolList.Size = New-Object System.Drawing.Size(330, 310)
$toolList.Anchor = 'Top,Bottom,Left'
$toolList.DisplayMember = 'DisplayName'
$toolList.IntegralHeight = $false
$form.Controls.Add($toolList)

$descriptionGroup = New-Object System.Windows.Forms.GroupBox
$descriptionGroup.Location = New-Object System.Drawing.Point(370, 103)
$descriptionGroup.Size = New-Object System.Drawing.Size(430, 310)
$descriptionGroup.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($descriptionGroup)

$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Location = New-Object System.Drawing.Point(16, 28)
$descriptionLabel.Size = New-Object System.Drawing.Size(398, 94)
$descriptionLabel.Anchor = 'Top,Left,Right'
$descriptionGroup.Controls.Add($descriptionLabel)

$fileCaptionLabel = New-Object System.Windows.Forms.Label
$fileCaptionLabel.Location = New-Object System.Drawing.Point(16, 142)
$fileCaptionLabel.Size = New-Object System.Drawing.Size(398, 22)
$fileCaptionLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$descriptionGroup.Controls.Add($fileCaptionLabel)

$fileLabel = New-Object System.Windows.Forms.TextBox
$fileLabel.Location = New-Object System.Drawing.Point(16, 168)
$fileLabel.Size = New-Object System.Drawing.Size(398, 24)
$fileLabel.Anchor = 'Top,Left,Right'
$fileLabel.ReadOnly = $true
$descriptionGroup.Controls.Add($fileLabel)

$unblockInfoLabel = New-Object System.Windows.Forms.Label
$unblockInfoLabel.Location = New-Object System.Drawing.Point(16, 213)
$unblockInfoLabel.Size = New-Object System.Drawing.Size(398, 62)
$unblockInfoLabel.Anchor = 'Top,Left,Right'
$descriptionGroup.Controls.Add($unblockInfoLabel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 425)
$statusLabel.Size = New-Object System.Drawing.Size(500, 50)
$statusLabel.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($statusLabel)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(594, 438)
$startButton.Size = New-Object System.Drawing.Size(100, 34)
$startButton.Anchor = 'Bottom,Right'
$form.AcceptButton = $startButton
$form.Controls.Add($startButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(702, 438)
$closeButton.Size = New-Object System.Drawing.Size(98, 34)
$closeButton.Anchor = 'Bottom,Right'
$form.CancelButton = $closeButton
$form.Controls.Add($closeButton)

function Get-UiLanguage {
    if ($languageBox.SelectedIndex -eq 1) { return 'en' }
    return 'de'
}

function Update-SelectedTool {
    $entry = $toolList.SelectedItem
    if ($null -eq $entry) {
        $descriptionLabel.Text = ''
        $fileLabel.Text = ''
        $startButton.Enabled = $false
        return
    }

    $descriptionLabel.Text = [string]$entry.Description
    $fileLabel.Text = [string]$entry.ScriptFile
    $startButton.Enabled = $true
}

function Update-UiLanguage {
    $language = Get-UiLanguage
    $selectedKey = if ($null -ne $toolList.SelectedItem) {
        [string]$toolList.SelectedItem.Key
    }
    else {
        'WinDirStat3'
    }

    $toolList.BeginUpdate()
    try {
        $toolList.Items.Clear()
        $script:CurrentEntries = @(
            foreach ($definition in $toolDefinitions) {
                if ($language -eq 'de') {
                    [PSCustomObject]@{
                        Key         = $definition.Key
                        DisplayName = $definition.GermanName
                        Description = $definition.GermanDescription
                        ScriptFile  = $definition.GermanScript
                        Console     = $definition.Console
                    }
                }
                else {
                    [PSCustomObject]@{
                        Key         = $definition.Key
                        DisplayName = $definition.EnglishName
                        Description = $definition.EnglishDescription
                        ScriptFile  = $definition.EnglishScript
                        Console     = $definition.Console
                    }
                }
            }
        )

        $selectedIndex = 0
        for ($index = 0; $index -lt $script:CurrentEntries.Count; $index++) {
            [void]$toolList.Items.Add($script:CurrentEntries[$index])
            if ($script:CurrentEntries[$index].Key -eq $selectedKey) {
                $selectedIndex = $index
            }
        }
        $toolList.SelectedIndex = $selectedIndex
    }
    finally {
        $toolList.EndUpdate()
    }

    if ($language -eq 'de') {
        $titleLabel.Text = 'Windows-Skripte starten'
        $languageLabel.Text = 'Sprache:'
        $introLabel.Text = 'Werkzeug auswählen. Vor jedem Start entfernt der Starter mit Unblock-File eine vorhandene Internet-Markierung von der ausgewählten Skriptdatei.'
        $descriptionGroup.Text = 'Beschreibung'
        $fileCaptionLabel.Text = 'Gestartete Datei'
        $unblockInfoLabel.Text = 'Unblock-File ändert keine Ausführungsrichtlinie und erteilt keine Administratorrechte. Der Starter verwendet Bypass nur im neu gestarteten Prozess.'
        $startButton.Text = 'Starten'
        $closeButton.Text = 'Schließen'
        $statusLabel.Text = 'Bereit.'
    }
    else {
        $titleLabel.Text = 'Launch Windows scripts'
        $languageLabel.Text = 'Language:'
        $introLabel.Text = 'Select a tool. Before each launch, the launcher uses Unblock-File to remove an Internet zone marker from the selected script file.'
        $descriptionGroup.Text = 'Description'
        $fileCaptionLabel.Text = 'Script to launch'
        $unblockInfoLabel.Text = 'Unblock-File does not change the execution policy or grant administrator rights. The launcher uses Bypass only in the newly started process.'
        $startButton.Text = 'Launch'
        $closeButton.Text = 'Close'
        $statusLabel.Text = 'Ready.'
    }

    Update-SelectedTool
}

function Start-SelectedTool {
    $entry = $toolList.SelectedItem
    if ($null -eq $entry) { return }

    $language = Get-UiLanguage
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath ([string]$entry.ScriptFile)
    $encodingDoctorDryRun = $false

    if (-not [System.IO.File]::Exists($scriptPath)) {
        $message = if ($language -eq 'de') {
            "Die Skriptdatei wurde nicht gefunden:`n$scriptPath"
        }
        else {
            "The script file was not found:`n$scriptPath"
        }
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $message,
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    if ([string]$entry.Key -eq 'EncodingDoctor') {
        $choice = if ($language -eq 'de') {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Der Encoding Doctor untersucht den gesamten Repository-Ordner rekursiv.`n`nJa: Reparaturmodus mit Sicherungskopien`nNein: nur prüfen (DryRun)`nAbbrechen: nicht starten",
                'Encoding Doctor',
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Encoding Doctor scans the complete repository folder recursively.`n`nYes: repair mode with backups`nNo: inspect only (DryRun)`nCancel: do not launch",
                'Encoding Doctor',
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2
            )
        }

        if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) {
            return
        }
        $encodingDoctorDryRun = ($choice -eq [System.Windows.Forms.DialogResult]::No)
    }

    try {
        Unblock-File -LiteralPath $scriptPath -ErrorAction Stop

        $windowsPowerShell = Join-Path $PSHOME 'powershell.exe'
        if (-not [System.IO.File]::Exists($windowsPowerShell)) {
            $windowsPowerShell = (Get-Command powershell.exe -ErrorAction Stop).Source
        }

        $arguments = New-Object 'System.Collections.Generic.List[string]'
        [void]$arguments.Add('-NoProfile')
        [void]$arguments.Add('-ExecutionPolicy')
        [void]$arguments.Add('Bypass')
        [void]$arguments.Add('-STA')
        if ([bool]$entry.Console) {
            [void]$arguments.Add('-NoExit')
        }
        [void]$arguments.Add('-File')
        [void]$arguments.Add(('"{0}"' -f $scriptPath))
        if ($encodingDoctorDryRun) {
            [void]$arguments.Add('-DryRun')
        }

        Start-Process -FilePath $windowsPowerShell `
                      -ArgumentList $arguments.ToArray() `
                      -WorkingDirectory $PSScriptRoot | Out-Null

        $statusLabel.Text = if ($language -eq 'de') {
            'Internet-Markierung entfernt (falls vorhanden) und Skript gestartet.'
        }
        else {
            'Internet zone marker removed (if present) and script launched.'
        }
    }
    catch {
        $message = if ($language -eq 'de') {
            "Das Skript konnte nicht freigegeben oder gestartet werden:`n$($_.Exception.Message)"
        }
        else {
            "The script could not be unblocked or launched:`n$($_.Exception.Message)"
        }
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $message,
            $form.Text,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

$languageBox.Add_SelectedIndexChanged({ Update-UiLanguage })
$toolList.Add_SelectedIndexChanged({ Update-SelectedTool })
$toolList.Add_DoubleClick({ Start-SelectedTool })
$startButton.Add_Click({ Start-SelectedTool })
$closeButton.Add_Click({ $form.Close() })

$languageBox.SelectedIndex = if ([Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName -eq 'de') { 0 } else { 1 }

[void]$form.ShowDialog()
