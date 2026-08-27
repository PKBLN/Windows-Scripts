#Requires -Version 5.1
<#
.SYNOPSIS
    GUI zum Durchsuchen des archive.org CDX-Servers nach Dateien einer Domain.
.DESCRIPTION
    Eingabe: Domain/URL, Dateitypen (Wildcard oder Regex), Zeitraum via DatePicker, MatchType.
    Ausgabe: RichTextBox mit klickbaren Wayback-Download-Links (id_-Suffix = Originaldatei).
    Bei jeder neuen Suche wird das Ergebnisfeld vollstaendig geleert.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ---------- Hilfsfunktionen ----------

function ConvertTo-ExtensionRegex {
    param([string]$Pattern)

    if ([string]::IsNullOrWhiteSpace($Pattern)) { return $null }

    $p = $Pattern.Trim()

    # Falls der Nutzer bereits eine vollstaendige Regex eingibt (erkennbar an ^, $, (, |, \)
    if ($p -match '[\^\$\(\)\|\\]') {
        return $p
    }

    # Mehrere Endungen per Komma/Semikolon getrennt, "*.exe", ".exe" oder nur "exe" erlaubt
    $exts = $p -split '[,;]' | ForEach-Object {
        $e = $_.Trim()
        $e = $e -replace '^\*', ''      # "*.exe" -> ".exe"
        $e = $e -replace '^\.', ''      # ".exe"  -> "exe"
        $e = $e.Trim()
        if ($e) { [regex]::Escape($e) }
    } | Where-Object { $_ }

    if (-not $exts) { return $null }

    $alt = $exts -join '|'
    return "(?i).*\.($alt)(\?.*)?`$"
}

function ConvertTo-CleanDomain {
    param([string]$Value)
    $v = $Value.Trim()
    $v = $v -replace '^[a-zA-Z]+://', ''   # Schema entfernen, z.B. https://
    $v = $v.TrimEnd('/')
    return $v
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Build-CdxUri {
    param(
        [string]$Domain,
        [string]$MatchType,
        [string]$FilterRegex,
        [string]$From,
        [string]$To
    )

    # UriBuilder garantiert eine strukturell gueltige URI - Host/Schema sind fest verdrahtet,
    # nur der Query-Teil wird dynamisch (und korrekt escaped) zusammengesetzt.
    $uriBuilder = [System.UriBuilder]::new("https://web.archive.org/cdx/search/cdx")

    $params = [System.Collections.Generic.List[string]]::new()
    $params.Add("url=" + [uri]::EscapeDataString($Domain))
    $params.Add("matchType=" + [uri]::EscapeDataString($MatchType))
    $params.Add("collapse=urlkey")
    $params.Add("fl=timestamp,original,statuscode,length")
    $params.Add("output=json")

    if ($FilterRegex) {
        $params.Add("filter=" + [uri]::EscapeDataString("original:$FilterRegex"))
    }
    $params.Add("filter=" + [uri]::EscapeDataString("statuscode:200"))

    if ($From) { $params.Add("from=" + [uri]::EscapeDataString($From)) }
    if ($To)   { $params.Add("to=" + [uri]::EscapeDataString($To)) }

    $uriBuilder.Query = ($params -join "&")
    return $uriBuilder.Uri
}

function Invoke-CdxSearch {
    param(
        [string]$Domain,
        [string]$MatchType,
        [string]$FilterRegex,
        [string]$From,
        [string]$To
    )

    $cleanDomain = ConvertTo-CleanDomain -Value $Domain
    $uri = Build-CdxUri -Domain $cleanDomain -MatchType $MatchType -FilterRegex $FilterRegex -From $From -To $To

    $resp = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 120 -UseBasicParsing
    return $resp
}

# ---------- GUI-Aufbau ----------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Archive.org CDX Suche"
$form.Size = New-Object System.Drawing.Size(780, 640)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size

$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Text = "Domain / URL:"
$labelDomain.Location = New-Object System.Drawing.Point(12, 15)
$labelDomain.AutoSize = $true

$textDomain = New-Object System.Windows.Forms.TextBox
$textDomain.Location = New-Object System.Drawing.Point(140, 12)
$textDomain.Size = New-Object System.Drawing.Size(400, 22)
$textDomain.Text = "powerbasic.com"

$labelMatch = New-Object System.Windows.Forms.Label
$labelMatch.Text = "Suchbereich:"
$labelMatch.Location = New-Object System.Drawing.Point(560, 15)
$labelMatch.AutoSize = $true

$comboMatch = New-Object System.Windows.Forms.ComboBox
$comboMatch.Location = New-Object System.Drawing.Point(650, 12)
$comboMatch.Size = New-Object System.Drawing.Size(105, 22)
$comboMatch.DropDownStyle = "DropDownList"
[void]$comboMatch.Items.AddRange(@("domain", "prefix", "host", "exact"))
$comboMatch.SelectedItem = "domain"

$labelExt = New-Object System.Windows.Forms.Label
$labelExt.Text = "Dateitypen (z.B. exe oder *.exe,*.zip):"
$labelExt.Location = New-Object System.Drawing.Point(12, 48)
$labelExt.AutoSize = $true

$textExt = New-Object System.Windows.Forms.TextBox
$textExt.Location = New-Object System.Drawing.Point(240, 45)
$textExt.Size = New-Object System.Drawing.Size(300, 22)
$textExt.Text = "exe"

# --- Von-Datum: DateTimePicker mit Checkbox (aktiv/inaktiv) ---
$labelFrom = New-Object System.Windows.Forms.Label
$labelFrom.Text = "Von:"
$labelFrom.Location = New-Object System.Drawing.Point(12, 82)
$labelFrom.AutoSize = $true

$dtpFrom = New-Object System.Windows.Forms.DateTimePicker
$dtpFrom.Location = New-Object System.Drawing.Point(140, 79)
$dtpFrom.Size = New-Object System.Drawing.Size(150, 22)
$dtpFrom.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpFrom.CustomFormat = "yyyy-MM-dd"
$dtpFrom.ShowCheckBox = $true
$dtpFrom.Checked = $false
$dtpFrom.Value = (Get-Date "2000-01-01")
$dtpFrom.MinDate = (Get-Date "1996-01-01")
$dtpFrom.MaxDate = (Get-Date)

# --- Bis-Datum: DateTimePicker mit Checkbox (aktiv/inaktiv) ---
$labelTo = New-Object System.Windows.Forms.Label
$labelTo.Text = "Bis:"
$labelTo.Location = New-Object System.Drawing.Point(300, 82)
$labelTo.AutoSize = $true

$dtpTo = New-Object System.Windows.Forms.DateTimePicker
$dtpTo.Location = New-Object System.Drawing.Point(340, 79)
$dtpTo.Size = New-Object System.Drawing.Size(150, 22)
$dtpTo.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpTo.CustomFormat = "yyyy-MM-dd"
$dtpTo.ShowCheckBox = $true
$dtpTo.Checked = $false
$dtpTo.Value = (Get-Date)
$dtpTo.MinDate = (Get-Date "1996-01-01")
$dtpTo.MaxDate = (Get-Date)

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "Suchen"
$buttonSearch.Location = New-Object System.Drawing.Point(560, 78)
$buttonSearch.Size = New-Object System.Drawing.Size(90, 26)

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "Leeren"
$buttonClear.Location = New-Object System.Drawing.Point(655, 78)
$buttonClear.Size = New-Object System.Drawing.Size(100, 26)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Bereit."
$statusLabel.Location = New-Object System.Drawing.Point(12, 118)
$statusLabel.AutoSize = $true
$statusLabel.ForeColor = [System.Drawing.Color]::DimGray

$richResult = New-Object System.Windows.Forms.RichTextBox
$richResult.Location = New-Object System.Drawing.Point(12, 145)
$richResult.Size = New-Object System.Drawing.Size(756, 450)
$richResult.Anchor = "Top,Bottom,Left,Right"
$richResult.ReadOnly = $true
$richResult.DetectUrls = $true
$richResult.Font = New-Object System.Drawing.Font("Consolas", 9)
$richResult.WordWrap = $false
$richResult.ScrollBars = "Both"

$richResult.Add_LinkClicked({
    param($sender, $e)
    Start-Process $e.LinkText
})

$form.Controls.AddRange(@(
    $labelDomain, $textDomain, $labelMatch, $comboMatch,
    $labelExt, $textExt,
    $labelFrom, $dtpFrom, $labelTo, $dtpTo,
    $buttonSearch, $buttonClear,
    $statusLabel, $richResult
))

# ---------- Event: Suche ----------

$buttonSearch.Add_Click({
    # Ergebnisfeld bei jeder neuen Suche zuerst leeren
    $richResult.Clear()
    $statusLabel.Text = "Suche laeuft..."
    $statusLabel.ForeColor = [System.Drawing.Color]::DimGray
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $domain = $textDomain.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $statusLabel.Text = "Bitte eine Domain/URL angeben."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    $filterRegex = ConvertTo-ExtensionRegex -Pattern $textExt.Text
    $matchType = $comboMatch.SelectedItem.ToString()

    # Datum nur uebernehmen, wenn die Checkbox am DateTimePicker aktiviert ist
    $from = if ($dtpFrom.Checked) { $dtpFrom.Value.ToString("yyyyMMdd") } else { $null }
    $to   = if ($dtpTo.Checked)   { $dtpTo.Value.ToString("yyyyMMdd") }   else { $null }

    if ($from -and $to -and $dtpFrom.Value -gt $dtpTo.Value) {
        $statusLabel.Text = "'Von' liegt nach 'Bis' - bitte Zeitraum pruefen."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    try {
        $rows = Invoke-CdxSearch -Domain $domain -MatchType $matchType -FilterRegex $filterRegex -From $from -To $to

        if (-not $rows -or $rows.Count -le 1) {
            $richResult.AppendText("Keine Treffer gefunden.`r`n")
            $statusLabel.Text = "0 Treffer."
            $statusLabel.ForeColor = [System.Drawing.Color]::DimGray
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }

        $dataRows = $rows | Select-Object -Skip 1
        $count = 0

        foreach ($row in $dataRows) {
            $timestamp = $row[0]
            $original  = $row[1]
            $status    = $row[2]
            $length    = 0
            [long]::TryParse($row[3], [ref]$length) | Out-Null

            $dt = [datetime]::MinValue
            try {
                $dt = [datetime]::ParseExact($timestamp, "yyyyMMddHHmmss", $null)
            } catch {}

            $downloadLink = "https://web.archive.org/web/${timestamp}id_/${original}"
            $sizeText = Format-Bytes -Bytes $length

            $richResult.AppendText("[$($dt.ToString('yyyy-MM-dd HH:mm'))] ($sizeText, HTTP $status)`r`n")
            $richResult.AppendText("$downloadLink`r`n`r`n")
            $count++
        }

        $statusLabel.Text = "$count Treffer gefunden."
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $richResult.SelectionStart = 0
        $richResult.ScrollToCaret()
    }
    catch {
        $richResult.AppendText("Fehler bei der Abfrage:`r`n$($_.Exception.Message)`r`n")
        $statusLabel.Text = "Fehler."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ---------- Event: Leeren ----------

$buttonClear.Add_Click({
    $richResult.Clear()
    $statusLabel.Text = "Bereit."
    $statusLabel.ForeColor = [System.Drawing.Color]::DimGray
})

# ---------- Start ----------

[void]$form.ShowDialog()
