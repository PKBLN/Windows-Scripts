param(
    [switch]$DryRun
)

# Encoding-Doctor.ps1
# ASCII-only source file so this script cannot damage itself through
# ANSI/UTF-8 confusion.
#
# Main purpose:
#   * Windows PowerShell 5.1 reads UTF-8 .ps1/.psm1/.psd1 files WITHOUT
#     a BOM as the active ANSI code page. A perfectly correct UTF-8 source
#     file can therefore DISPLAY broken umlauts at runtime.
#   * For PowerShell files, this script adds a UTF-8 BOM when required.
#     The original UTF-8 bytes are preserved byte-for-byte; only EF BB BF
#     is prepended.
#   * It also detects and repairs actual mojibake text in common text files.
#
# Every changed file is backed up first.

$ErrorActionPreference = 'Stop'

$Root = $PSScriptRoot
$Recurse = $true
$MaxFileSizeMB = 25
$MaxRepairPasses = 4

# Set this to $true ONLY if you explicitly want a BOM added to every valid
# UTF-8 text file. Normally this should stay false.
$AddBomToAllUtf8Text = $false

# PowerShell 5.1 benefits from UTF-8 BOM for source files.
$PowerShellExtensions = @(
    '.ps1', '.psm1', '.psd1'
)

# Other text/code files are scanned for ACTUAL mojibake content.
# They do NOT automatically receive a BOM unless AddBomToAllUtf8Text is true.
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

    # NUL bytes usually indicate binary data, except BOM-marked UTF-16
    # which is detected before this function is used.
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
        # We only LABEL this as CP1252 here. For non-PowerShell files we
        # avoid automatic conversion because an unknown legacy encoding
        # cannot be identified with certainty.
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

    # Typical characters at the start of UTF-8-as-CP1252 garbage:
    # U+00C2, U+00C3, U+00E2, U+00EF, U+00F0
    # Plus common signs of a second broken round-trip:
    # U+0192 and U+0178
    # U+FFFD is the Unicode replacement character.
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
            # Classic repair:
            # Unicode text containing CP1252-decoded UTF-8 garbage
            # -> original bytes -> decode those bytes as UTF-8.
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
Write-Host ('Root    : ' + $Root)
Write-Host ('Files   : ' + $files.Count)
Write-Host ('Dry run : ' + [bool]$DryRun)
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
            Write-Host ('[SKIP BINARY] ' + $relative)
            $stats.BinarySkipped++
            continue
        }

        # ----------------------------------------------------------
        # CASE 1:
        # Correct UTF-8 without BOM PowerShell source.
        #
        # This is the MOST IMPORTANT case for Windows PowerShell 5.1.
        # The content can be completely correct on disk and still appear
        # as broken umlauts at runtime because PS 5.1 assumes ANSI.
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
                    # Safest possible fix: preserve every original byte and
                    # only prepend the UTF-8 BOM.
                    Add-Utf8BomWithoutReencoding $file.FullName $bytes
                }
            }

            if ($contentActuallyBroken) {
                Write-Host ('[FIX MOJIBAKE + BOM] ' + $relative)
                $stats.MojibakeFixed++
            }
            else {
                if (Test-HasNonAscii $bytes) {
                    Write-Host ('[ADD UTF8 BOM] ' + $relative)
                }
                else {
                    Write-Host ('[ADD UTF8 BOM ASCII] ' + $relative)
                }

                $stats.BomAdded++
            }

            continue
        }

        # ----------------------------------------------------------
        # CASE 2:
        # PowerShell source in an old/unknown 8-bit encoding.
        # For PS files only, normalize probable Windows-1252 to UTF-8 BOM.
        # ----------------------------------------------------------
        if ($isPowerShell -and $info.Kind -eq 'LEGACY/UNKNOWN' -and $null -ne $info.Text) {
            $fixedText = Repair-Mojibake $info.Text

            if (-not $DryRun) {
                Backup-File $file
                Write-Utf8WithBom $file.FullName $fixedText
            }

            Write-Host ('[CONVERT PS -> UTF8 BOM] ' + $relative)
            $stats.LegacyPsConverted++

            if ($fixedText -cne $info.Text) {
                $stats.MojibakeFixed++
            }

            continue
        }

        # ----------------------------------------------------------
        # CASE 3:
        # UTF-8 text where the TEXT ITSELF is already mojibake.
        # This applies to PowerShell and other common text/code files.
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

                Write-Host ('[FIX MOJIBAKE] ' + $relative)
                $stats.MojibakeFixed++
                continue
            }

            if ($AddBomToAllUtf8Text -and $info.Kind -eq 'UTF8-NOBOM') {
                if (-not $DryRun) {
                    Backup-File $file
                    Add-Utf8BomWithoutReencoding $file.FullName $bytes
                }

                Write-Host ('[ADD UTF8 BOM] ' + $relative)
                $stats.BomAdded++
                continue
            }

            $stats.AlreadyOk++
            continue
        }

        # ----------------------------------------------------------
        # CASE 4:
        # UTF-16 is already self-describing via BOM. Leave untouched.
        # ----------------------------------------------------------
        if ($info.Kind -like 'UTF16-*') {
            $stats.AlreadyOk++
            continue
        }

        # ----------------------------------------------------------
        # CASE 5:
        # Non-PowerShell legacy/unknown file. Report but do not guess.
        # ----------------------------------------------------------
        Write-Host ('[SKIP LEGACY/UNKNOWN] ' + $relative)
        $stats.LegacySkipped++
    }
    catch {
        Write-Warning ('[ERROR] ' + $file.FullName + ' : ' + $_.Exception.Message)
        $stats.Errors++
    }
}

Write-Host ''
Write-Host 'Summary'
Write-Host '-------'
Write-Host ('Checked             : ' + $stats.Checked)
Write-Host ('UTF-8 BOM added     : ' + $stats.BomAdded)
Write-Host ('Mojibake fixed      : ' + $stats.MojibakeFixed)
Write-Host ('Legacy PS converted : ' + $stats.LegacyPsConverted)
Write-Host ('Already OK          : ' + $stats.AlreadyOk)
Write-Host ('Legacy skipped      : ' + $stats.LegacySkipped)
Write-Host ('Binary skipped      : ' + $stats.BinarySkipped)
Write-Host ('Errors              : ' + $stats.Errors)

if (-not $DryRun -and
    ($stats.BomAdded + $stats.MojibakeFixed + $stats.LegacyPsConverted) -gt 0) {
    Write-Host ('Backup              : ' + $BackupRoot)
}

Write-Host ''
if ($DryRun) {
    Write-Host 'Dry run only. No files were changed.'
}
else {
    Write-Host 'Finished.'
}

Read-Host 'Press Enter to exit'
