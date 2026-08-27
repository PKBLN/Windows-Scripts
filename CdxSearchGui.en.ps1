#Requires -Version 5.1
<#
.SYNOPSIS
    GUI for searching the archive.org CDX server for files from a domain.
.DESCRIPTION
    Input: domain/URL, file types (wildcard or regex), date range via DatePicker, match type.
    Output: RichTextBox with clickable Wayback download links (id_ suffix = original file).
    The results field is cleared completely before every new search.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ---------- Helper functions ----------

function ConvertTo-ExtensionRegex {
    param([string]$Pattern)

    if ([string]::IsNullOrWhiteSpace($Pattern)) { return $null }

    $p = $Pattern.Trim()

    # If the user has already entered a complete regex (identified by ^, $, (, |, \)
    if ($p -match '[\^\$\(\)\|\\]') {
        return $p
    }

    # Multiple extensions separated by commas/semicolons; "*.exe", ".exe", or just "exe" are allowed
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
    $v = $v -replace '^[a-zA-Z]+://', ''   # Remove the scheme, e.g. https://
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

    # UriBuilder guarantees a structurally valid URI; the host and scheme are fixed,
    # and only the query portion is assembled dynamically (and escaped correctly).
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

# ---------- GUI setup ----------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Archive.org CDX Search"
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
$labelMatch.Text = "Search scope:"
$labelMatch.Location = New-Object System.Drawing.Point(560, 15)
$labelMatch.AutoSize = $true

$comboMatch = New-Object System.Windows.Forms.ComboBox
$comboMatch.Location = New-Object System.Drawing.Point(650, 12)
$comboMatch.Size = New-Object System.Drawing.Size(105, 22)
$comboMatch.DropDownStyle = "DropDownList"
[void]$comboMatch.Items.AddRange(@("domain", "prefix", "host", "exact"))
$comboMatch.SelectedItem = "domain"

$labelExt = New-Object System.Windows.Forms.Label
$labelExt.Text = "File types (e.g. exe or *.exe,*.zip):"
$labelExt.Location = New-Object System.Drawing.Point(12, 48)
$labelExt.AutoSize = $true

$textExt = New-Object System.Windows.Forms.TextBox
$textExt.Location = New-Object System.Drawing.Point(240, 45)
$textExt.Size = New-Object System.Drawing.Size(300, 22)
$textExt.Text = "exe"

# --- From date: DateTimePicker with check box (enabled/disabled) ---
$labelFrom = New-Object System.Windows.Forms.Label
$labelFrom.Text = "From:"
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

# --- To date: DateTimePicker with check box (enabled/disabled) ---
$labelTo = New-Object System.Windows.Forms.Label
$labelTo.Text = "To:"
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
$buttonSearch.Text = "Search"
$buttonSearch.Location = New-Object System.Drawing.Point(560, 78)
$buttonSearch.Size = New-Object System.Drawing.Size(90, 26)

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "Clear"
$buttonClear.Location = New-Object System.Drawing.Point(655, 78)
$buttonClear.Size = New-Object System.Drawing.Size(100, 26)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready."
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

# ---------- Event: Search ----------

$buttonSearch.Add_Click({
    # Clear the results field before every new search
    $richResult.Clear()
    $statusLabel.Text = "Searching..."
    $statusLabel.ForeColor = [System.Drawing.Color]::DimGray
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    $domain = $textDomain.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($domain)) {
        $statusLabel.Text = "Please enter a domain/URL."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    $filterRegex = ConvertTo-ExtensionRegex -Pattern $textExt.Text
    $matchType = $comboMatch.SelectedItem.ToString()

    # Use a date only if the corresponding DateTimePicker check box is enabled
    $from = if ($dtpFrom.Checked) { $dtpFrom.Value.ToString("yyyyMMdd") } else { $null }
    $to   = if ($dtpTo.Checked)   { $dtpTo.Value.ToString("yyyyMMdd") }   else { $null }

    if ($from -and $to -and $dtpFrom.Value -gt $dtpTo.Value) {
        $statusLabel.Text = "'From' is later than 'To' - please check the date range."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }

    try {
        $rows = Invoke-CdxSearch -Domain $domain -MatchType $matchType -FilterRegex $filterRegex -From $from -To $to

        if (-not $rows -or $rows.Count -le 1) {
            $richResult.AppendText("No matches found.`r`n")
            $statusLabel.Text = "0 matches."
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

        $statusLabel.Text = "$count matches found."
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $richResult.SelectionStart = 0
        $richResult.ScrollToCaret()
    }
    catch {
        $richResult.AppendText("Query failed:`r`n$($_.Exception.Message)`r`n")
        $statusLabel.Text = "Error."
        $statusLabel.ForeColor = [System.Drawing.Color]::Firebrick
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ---------- Event: Clear ----------

$buttonClear.Add_Click({
    $richResult.Clear()
    $statusLabel.Text = "Ready."
    $statusLabel.ForeColor = [System.Drawing.Color]::DimGray
})

# ---------- Start ----------

[void]$form.ShowDialog()
