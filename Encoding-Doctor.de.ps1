param(
    [switch]$DryRun
)

# Encoding-Doctor.de.ps1
# Reine ASCII-Quelldatei, damit dieses Skript sich nicht selbst durch
# eine Verwechslung von ANSI und UTF-8 beschaedigen kann.
#
# Hauptzweck:
#   * Windows PowerShell 5.1 liest UTF-8-Dateien mit den Endungen
#     .ps1/.psm1/.psd1 OHNE BOM als aktive ANSI-Codepage. Eine vollkommen
#     korrekte UTF-8-Quelldatei kann deshalb zur Laufzeit fehlerhafte
#     Umlaute ANZEIGEN.
#   * Bei PowerShell-Dateien fuegt dieses Skript bei Bedarf eine UTF-8-BOM
#     hinzu. Die urspruenglichen UTF-8-Bytes bleiben Byte fuer Byte erhalten;
#     nur EF BB BF wird vorangestellt.
#   * Es erkennt und repariert ausserdem tatsaechlichen Mojibake-Text in
#     gaengigen Textdateien.
#
# Von jeder geaenderten Datei wird zuerst eine Sicherung erstellt.

$ErrorActionPreference = 'Stop'

$Root = $PSScriptRoot
$Recurse = $true
$MaxFileSizeMB = 25
$MaxRepairPasses = 4

# Nur dann auf $true setzen, wenn ausdruecklich jede gueltige UTF-8-Textdatei
# eine BOM erhalten soll. Normalerweise sollte dieser Wert $false bleiben.
$AddBomToAllUtf8Text = $false

# PowerShell 5.1 profitiert bei Quelldateien von einer UTF-8-BOM.
$PowerShellExtensions = @(
    '.ps1', '.psm1', '.psd1'
)

# Andere Text-/Codedateien werden auf TATSAECHLICHEN Mojibake-Inhalt geprueft.
# Sie erhalten NICHT automatisch eine BOM, sofern AddBomToAllUtf8Text nicht
# auf true gesetzt ist.
$TextExtensions = @(
    '.ps1', '.psm1', '.psd1',
    '.txt', '.log', '.md', '.csv', '.tsv',
    '.json', '.xml', '.xaml', '.config',
    '.ini', '.cfg', '.conf', '.properties',
    '.html', '.htm', '.css',
    '.js', '.mjs', '.cjs', '.ts',
    '.cs', '.vb', '.fs',
    '.java', '.kt',
    '.c', '.h', '.cpp', '.hpp',
    '.py', '.rb', '.php',
    '.sql',
    '.yaml', '.yml',
    '.cmd', '.bat',
    '.reg'
)

$Stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$BackupRoot = Join-Path $Root ('_Encoding_Backup_' + $Stamp)

$Utf8StrictNoBom = New-Object System.Text.UTF8Encoding($false, $true)
$Utf8NoBom       = New-Object System.Text.UTF8Encoding($false)
$Utf8WithBom     = New-Object System.Text.UTF8Encoding($true)

$Cp1252Strict = [System.Text.Encoding]::GetEncoding(
    1252,
    [System.Text.EncoderFallback]::ExceptionFallback,
    [System.Text.DecoderFallback]::ExceptionFallback
)

$Utf16LE = New-Object System.Text.UnicodeEncoding($false, $true, $true)
$Utf16BE = New-Object System.Text.UnicodeEncoding($true, $true, $true)

function Test-StartsWithBytes {
    param(
        [byte[]]$Bytes,
        [byte[]]$Prefix
    )

    if ($Bytes.Length -lt $Prefix.Length) {
        return $false
    }

    for ($i = 0; $i -lt $Prefix.Length; $i++) {
        if ($Bytes[$i] -ne $Prefix[$i]) {
            return $false
        }
    }

    return $true
}

function Test-HasNonAscii {
    param([byte[]]$Bytes)

    foreach ($b in $Bytes) {
        if ($b -ge 0x80) {
            return $true
        }
    }

    return $false
}

function Test-LooksBinary {
    param([byte[]]$Bytes)

    # NUL-Bytes weisen normalerweise auf Binaerdaten hin. Eine Ausnahme
    # bildet UTF-16 mit BOM, das vor Nutzung dieser Funktion erkannt wird.
    $limit = [Math]::Min($Bytes.Length, 8192)

    for ($i = 0; $i -lt $limit; $i++) {
        if ($Bytes[$i] -eq 0) {
            return $true
        }
    }

    return $false
}

function Get-EncodingInfo {
    param([byte[]]$Bytes)

    $utf8Bom = [byte[]](0xEF, 0xBB, 0xBF)
    $utf16LeBom = [byte[]](0xFF, 0xFE)
    $utf16BeBom = [byte[]](0xFE, 0xFF)

    if (Test-StartsWithBytes $Bytes $utf8Bom) {
        $text = [System.Text.Encoding]::UTF8.GetString(
            $Bytes,
            3,
            $Bytes.Length - 3
        )

        return [PSCustomObject]@{
            Kind = 'UTF8-BOM'
            Text = $text
            HasBom = $true
            IsUtf8 = $true
        }
    }

    if (Test-StartsWithBytes $Bytes $utf16LeBom) {
        $text = [System.Text.Encoding]::Unicode.GetString(
            $Bytes,
            2,
            $Bytes.Length - 2
        )

        return [PSCustomObject]@{
            Kind = 'UTF16-LE-BOM'
            Text = $text
            HasBom = $true
            IsUtf8 = $false
        }
    }

    if (Test-StartsWithBytes $Bytes $utf16BeBom) {
        $text = [System.Text.Encoding]::BigEndianUnicode.GetString(
            $Bytes,
            2,
            $Bytes.Length - 2
        )

        return [PSCustomObject]@{
            Kind = 'UTF16-BE-BOM'
            Text = $text
            HasBom = $true
            IsUtf8 = $false
        }
    }

    if (Test-LooksBinary $Bytes) {
        return [PSCustomObject]@{
            Kind = 'BINARY'
            Text = $null
            HasBom = $false
            IsUtf8 = $false
        }
    }

    try {
        $text = $Utf8StrictNoBom.GetString($Bytes)

        return [PSCustomObject]@{
            Kind = 'UTF8-NOBOM'
            Text = $text
            HasBom = $false
            IsUtf8 = $true
        }
    }
    catch {
        # Hier wird die Codierung lediglich als CP1252 BEZEICHNET. Bei
        # Nicht-PowerShell-Dateien vermeiden wir eine automatische
        # Konvertierung, da sich eine unbekannte Legacy-Codierung nicht
        # mit Sicherheit bestimmen laesst.
        try {
            $text = $Cp1252Strict.GetString($Bytes)
        }
        catch {
            $text = $null
        }

        return [PSCustomObject]@{
            Kind = 'LEGACY/UNKNOWN'
            Text = $text
            HasBom = $false
            IsUtf8 = $false
        }
    }
}

function Get-MojibakeScore {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return 0
    }

    # Typische Zeichen am Anfang von als CP1252 interpretiertem UTF-8-Zeichensalat:
    # U+00C2, U+00C3, U+00E2, U+00EF, U+00F0
    # Zusaetzlich haeufige Anzeichen eines zweiten fehlerhaften Durchlaufs:
    # U+0192 und U+0178
    # U+FFFD ist das Unicode-Ersatzzeichen.
    $codePoints = @(
        0x00C2,
        0x00C3,
        0x00E2,
        0x00EF,
        0x00F0,
        0x0192,
        0x0178,
        0xFFFD
    )

    $score = 0

    foreach ($cp in $codePoints) {
        $needle = [char]$cp
        $index = 0

        while ($true) {
            $index = $Text.IndexOf($needle, $index)

            if ($index -lt 0) {
                break
            }

            $score++
            $index++
        }
    }

    return $score
}

function Repair-Mojibake {
    param([string]$Text)

    $current = $Text

    for ($pass = 1; $pass -le $MaxRepairPasses; $pass++) {
        $oldScore = Get-MojibakeScore $current

        if ($oldScore -eq 0) {
            break
        }

        try {
            # Klassische Reparatur:
            # Unicode-Text mit als CP1252 interpretiertem UTF-8-Zeichensalat
            # -> urspruengliche Bytes -> diese Bytes als UTF-8 decodieren.
            $bytes = $Cp1252Strict.GetBytes($current)
            $candidate = $Utf8StrictNoBom.GetString($bytes)
        }
        catch {
            break
        }

        $newScore = Get-MojibakeScore $candidate

        if ($newScore -lt $oldScore) {
            $current = $candidate
        }
        else {
            break
        }
    }

    return $current
}

function Backup-File {
    param(
        [System.IO.FileInfo]$File
    )

    $relative = $File.FullName.Substring($Root.Length)
    $relative = $relative.TrimStart([char[]]@('\', '/'))

    $destination = Join-Path $BackupRoot $relative
    $destinationDir = Split-Path -Parent $destination

    if (-not (Test-Path -LiteralPath $destinationDir)) {
        New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
    }

    Copy-Item -LiteralPath $File.FullName -Destination $destination -Force
}

function Write-Utf8WithBom {
    param(
        [string]$Path,
        [string]$Text
    )

    [System.IO.File]::WriteAllText($Path, $Text, $Utf8WithBom)
}

function Write-Utf8NoBom {
    param(
        [string]$Path,
        [string]$Text
    )

    [System.IO.File]::WriteAllText($Path, $Text, $Utf8NoBom)
}

function Add-Utf8BomWithoutReencoding {
    param(
        [string]$Path,
        [byte[]]$OriginalBytes
    )

    $newBytes = New-Object byte[] ($OriginalBytes.Length + 3)
    $newBytes[0] = 0xEF
    $newBytes[1] = 0xBB
    $newBytes[2] = 0xBF

    [Array]::Copy(
        $OriginalBytes,
        0,
        $newBytes,
        3,
        $OriginalBytes.Length
    )

    [System.IO.File]::WriteAllBytes($Path, $newBytes)
}

if ($Recurse) {
    $files = Get-ChildItem -LiteralPath $Root -File -Recurse
}
else {
    $files = Get-ChildItem -LiteralPath $Root -File
}

$thisScript = $MyInvocation.MyCommand.Path
$maxBytes = $MaxFileSizeMB * 1MB

$files = @(
    $files | Where-Object {
        $ext = $_.Extension.ToLowerInvariant()
        $isText = $TextExtensions -contains $ext
        $notSelf = $_.FullName -ne $thisScript
        $sizeOk = $_.Length -le $maxBytes

        $notOldBackup =
            ($_.FullName.IndexOf('\_Encoding_Backup_', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) -and
            ($_.FullName.IndexOf('\_Mojibake_Backup_', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) -and
            ($_.FullName.IndexOf('/_Encoding_Backup_', [System.StringComparison]::OrdinalIgnoreCase) -lt 0) -and
            ($_.FullName.IndexOf('/_Mojibake_Backup_', [System.StringComparison]::OrdinalIgnoreCase) -lt 0)

        $isText -and $notSelf -and $sizeOk -and $notOldBackup
    }
)

$stats = [ordered]@{
    Checked = 0
    BomAdded = 0
    MojibakeFixed = 0
    LegacyPsConverted = 0
    AlreadyOk = 0
    LegacySkipped = 0
    BinarySkipped = 0
    Errors = 0
}

Write-Host ''
Write-Host 'Encoding Doctor'
Write-Host '==============='
Write-Host ('Stammverzeichnis : ' + $Root)
Write-Host ('Dateien           : ' + $files.Count)
Write-Host ('Testlauf          : ' + [bool]$DryRun)
Write-Host ''

foreach ($file in $files) {
    $stats.Checked++

    try {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        $info = Get-EncodingInfo $bytes
        $ext = $file.Extension.ToLowerInvariant()
        $isPowerShell = $PowerShellExtensions -contains $ext
        $relative = $file.FullName.Substring($Root.Length).TrimStart([char[]]@('\', '/'))

        if ($info.Kind -eq 'BINARY') {
            Write-Host ('[BINAER UEBERSPRUNGEN] ' + $relative)
            $stats.BinarySkipped++
            continue
        }

        # ----------------------------------------------------------
        # FALL 1:
        # Korrekte PowerShell-Quelldatei in UTF-8 ohne BOM.
        #
        # Dies ist der WICHTIGSTE Fall fuer Windows PowerShell 5.1.
        # Der Inhalt kann auf dem Datentraeger vollkommen korrekt sein
        # und zur Laufzeit trotzdem fehlerhafte Umlaute anzeigen, weil
        # PowerShell 5.1 ANSI annimmt.
        # ----------------------------------------------------------
        if ($isPowerShell -and $info.Kind -eq 'UTF8-NOBOM') {
            $fixedText = Repair-Mojibake $info.Text
            $contentActuallyBroken = ($fixedText -cne $info.Text)

            if (-not $DryRun) {
                Backup-File $file

                if ($contentActuallyBroken) {
                    Write-Utf8WithBom $file.FullName $fixedText
                }
                else {
                    # Sicherste Reparatur: jedes urspruengliche Byte bleibt
                    # erhalten und nur die UTF-8-BOM wird vorangestellt.
                    Add-Utf8BomWithoutReencoding $file.FullName $bytes
                }
            }

            if ($contentActuallyBroken) {
                Write-Host ('[MOJIBAKE REPARIERT + BOM] ' + $relative)
                $stats.MojibakeFixed++
            }
            else {
                if (Test-HasNonAscii $bytes) {
                    Write-Host ('[UTF8-BOM HINZUGEFUEGT] ' + $relative)
                }
                else {
                    Write-Host ('[UTF8-BOM ASCII HINZUGEFUEGT] ' + $relative)
                }

                $stats.BomAdded++
            }

            continue
        }

        # ----------------------------------------------------------
        # FALL 2:
        # PowerShell-Quelldatei in einer alten/unbekannten 8-Bit-Codierung.
        # Nur bei PS-Dateien wird wahrscheinliches Windows-1252 als
        # UTF-8 mit BOM normalisiert.
        # ----------------------------------------------------------
        if ($isPowerShell -and $info.Kind -eq 'LEGACY/UNKNOWN' -and $null -ne $info.Text) {
            $fixedText = Repair-Mojibake $info.Text

            if (-not $DryRun) {
                Backup-File $file
                Write-Utf8WithBom $file.FullName $fixedText
            }

            Write-Host ('[PS -> UTF8-BOM KONVERTIERT] ' + $relative)
            $stats.LegacyPsConverted++

            if ($fixedText -cne $info.Text) {
                $stats.MojibakeFixed++
            }

            continue
        }

        # ----------------------------------------------------------
        # FALL 3:
        # UTF-8-Text, dessen TEXT SELBST bereits Mojibake enthaelt.
        # Dies gilt fuer PowerShell und andere gaengige Text-/Codedateien.
        # ----------------------------------------------------------
        if ($info.IsUtf8 -and $null -ne $info.Text) {
            $fixedText = Repair-Mojibake $info.Text

            if ($fixedText -cne $info.Text) {
                if (-not $DryRun) {
                    Backup-File $file

                    if ($isPowerShell -or $info.Kind -eq 'UTF8-BOM' -or $AddBomToAllUtf8Text) {
                        Write-Utf8WithBom $file.FullName $fixedText
                    }
                    else {
                        Write-Utf8NoBom $file.FullName $fixedText
                    }
                }

                Write-Host ('[MOJIBAKE REPARIERT] ' + $relative)
                $stats.MojibakeFixed++
                continue
            }

            if ($AddBomToAllUtf8Text -and $info.Kind -eq 'UTF8-NOBOM') {
                if (-not $DryRun) {
                    Backup-File $file
                    Add-Utf8BomWithoutReencoding $file.FullName $bytes
                }

                Write-Host ('[UTF8-BOM HINZUGEFUEGT] ' + $relative)
                $stats.BomAdded++
                continue
            }

            $stats.AlreadyOk++
            continue
        }

        # ----------------------------------------------------------
        # FALL 4:
        # UTF-16 ist durch die BOM bereits eindeutig. Nicht veraendern.
        # ----------------------------------------------------------
        if ($info.Kind -like 'UTF16-*') {
            $stats.AlreadyOk++
            continue
        }

        # ----------------------------------------------------------
        # FALL 5:
        # Legacy-/unbekannte Datei ausserhalb von PowerShell. Melden,
        # aber keine Vermutung anstellen.
        # ----------------------------------------------------------
        Write-Host ('[LEGACY/UNBEKANNT UEBERSPRUNGEN] ' + $relative)
        $stats.LegacySkipped++
    }
    catch {
        Write-Warning ('[FEHLER] ' + $file.FullName + ' : ' + $_.Exception.Message)
        $stats.Errors++
    }
}

Write-Host ''
Write-Host 'Zusammenfassung'
Write-Host '---------------'
Write-Host ('Geprueft                  : ' + $stats.Checked)
Write-Host ('UTF-8-BOM hinzugefuegt    : ' + $stats.BomAdded)
Write-Host ('Mojibake repariert        : ' + $stats.MojibakeFixed)
Write-Host ('Legacy-PS konvertiert     : ' + $stats.LegacyPsConverted)
Write-Host ('Bereits in Ordnung        : ' + $stats.AlreadyOk)
Write-Host ('Legacy uebersprungen      : ' + $stats.LegacySkipped)
Write-Host ('Binaer uebersprungen      : ' + $stats.BinarySkipped)
Write-Host ('Fehler                    : ' + $stats.Errors)

if (-not $DryRun -and
    ($stats.BomAdded + $stats.MojibakeFixed + $stats.LegacyPsConverted) -gt 0) {
    Write-Host ('Sicherung                 : ' + $BackupRoot)
}

Write-Host ''
if ($DryRun) {
    Write-Host 'Nur Testlauf. Es wurden keine Dateien geaendert.'
}
else {
    Write-Host 'Fertig.'
}

Read-Host 'Zum Beenden die Eingabetaste druecken'
