# --------------------------------------------------------------
#  Datei-Inhalt Suche Pro
#  - case-insensitive Dateifilter
#  - case-insensitive Inhaltssuche
#  - Live-Anzeige der aktuellen Datei
#  - zuverlässiger Abbruch
#  - UI bleibt während der Suche bedienbar
# --------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ==============================================================
# GUI
# ==============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Datei-Inhalt Suche Pro'
$form.Size          = New-Object System.Drawing.Size(860, 680)
$form.StartPosition = 'CenterScreen'


# ---- Label & TextBox: Dateifilter ----------------------------

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text     = 'Dateifilter (z.B. *.txt oder *config*):'
$lblFile.Location = New-Object System.Drawing.Point(10, 20)
$lblFile.AutoSize = $true
$form.Controls.Add($lblFile)

$txtFileFilter = New-Object System.Windows.Forms.TextBox
$txtFileFilter.Text     = '*'
$txtFileFilter.Location = New-Object System.Drawing.Point(220, 17)
$txtFileFilter.Size     = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($txtFileFilter)


# ---- Label & TextBox: Suchbegriff ----------------------------

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Suchbegriff im Inhalt:'
$lblSearch.Location = New-Object System.Drawing.Point(10, 50)
$lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(220, 47)
$txtSearch.Size     = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($txtSearch)


# ---- Button: Suchen ------------------------------------------

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text     = 'Suchen'
$btnSearch.Location = New-Object System.Drawing.Point(540, 15)
$btnSearch.Size     = New-Object System.Drawing.Size(100, 60)
$form.Controls.Add($btnSearch)


# ---- Button: Abbrechen ---------------------------------------

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = 'Abbrechen'
$btnCancel.Location = New-Object System.Drawing.Point(650, 15)
$btnCancel.Size     = New-Object System.Drawing.Size(100, 60)
$btnCancel.Enabled  = $false
$form.Controls.Add($btnCancel)


# ---- Label: Aktuelle Datei -----------------------------------

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Text     = 'Aktuelle Datei: -'
$lblCurrent.Location = New-Object System.Drawing.Point(10, 80)
$lblCurrent.AutoSize = $true
$form.Controls.Add($lblCurrent)


# ---- ListBox: Ergebnisliste ----------------------------------

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 110)
$listBox.Size     = New-Object System.Drawing.Size(820, 500)
$listBox.HorizontalScrollbar = $true
$form.Controls.Add($listBox)


# ==============================================================
# Status / globale Variablen
# ==============================================================

$script:searchRunspace    = $null
$script:searchPowerShell  = $null
$script:searchAsync       = $null
$script:cancelSource      = $null
$script:resultQueue       = $null
$script:timer             = $null

$script:matchCount        = 0
$script:cancelMessageShown = $false
$script:fatalError        = $false


# ==============================================================
# Hilfsfunktion: Ressourcen freigeben
# ==============================================================

function Close-SearchResources {
    param(
        [switch]$ForceStop
    )

    if ($script:timer) {
        try {
            $script:timer.Stop()
            $script:timer.Dispose()
        }
        catch {}

        $script:timer = $null
    }

    if ($ForceStop -and $script:searchPowerShell) {
        try {
            $script:searchPowerShell.Stop()
        }
        catch {}
    }

    if ($script:searchPowerShell) {
        try {
            $script:searchPowerShell.Dispose()
        }
        catch {}

        $script:searchPowerShell = $null
    }

    if ($script:searchRunspace) {
        try {
            $script:searchRunspace.Dispose()
        }
        catch {}

        $script:searchRunspace = $null
    }

    if ($script:cancelSource) {
        try {
            $script:cancelSource.Dispose()
        }
        catch {}

        $script:cancelSource = $null
    }

    $script:searchAsync = $null
    $script:resultQueue = $null
}


# ==============================================================
# Hilfsfunktion: Meldungen des Workers verarbeiten
#
# WICHTIG:
# Diese Funktion wird nur vom WinForms-UI-Thread aufgerufen.
# Daher dürfen hier Controls direkt verändert werden.
# ==============================================================

function Read-WorkerMessages {

    if (-not $script:resultQueue) {
        return
    }

    $message = $null

    while ($script:resultQueue.TryDequeue([ref]$message)) {

        switch ($message.Type) {

            'Progress' {

                $lblCurrent.Text =
                    "Aktuelle Datei: $($message.Text)"
            }


            'Match' {

                $script:matchCount++

                [void]$listBox.Items.Add($message.Text)

                # automatisch zum neuesten Treffer scrollen
                if ($listBox.Items.Count -gt 0) {
                    $listBox.TopIndex =
                        $listBox.Items.Count - 1
                }
            }


            'Cancelled' {

                if (-not $script:cancelMessageShown) {

                    [void]$listBox.Items.Add(
                        'Suche wurde vom Benutzer abgebrochen.'
                    )

                    $script:cancelMessageShown = $true
                }
            }


            'FatalError' {

                [void]$listBox.Items.Add(
                    "Fehler: $($message.Text)"
                )

                $script:fatalError = $true
            }
        }

        $message = $null
    }
}


# ==============================================================
# Abbrechen
# ==============================================================

$btnCancel.Add_Click({

    if (
        $script:cancelSource -and
        -not $script:cancelSource.IsCancellationRequested
    ) {

        $script:cancelSource.Cancel()

        $btnCancel.Enabled = $false

        [void]$listBox.Items.Add(
            'Abbruch wird angefordert ...'
        )
    }
})


# ==============================================================
# Suche
# ==============================================================

$btnSearch.Add_Click({

    # ----------------------------------------------------------
    # Eingaben
    # ----------------------------------------------------------

    $searchTerm = $txtSearch.Text.Trim()
    $fileFilter = $txtFileFilter.Text.Trim()


    if ([string]::IsNullOrWhiteSpace($searchTerm)) {

        [System.Windows.Forms.MessageBox]::Show(
            'Bitte einen Suchbegriff eingeben!',
            'Datei-Inhalt Suche',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        return
    }


    # Leerer Dateifilter = alle Dateien
    if ([string]::IsNullOrWhiteSpace($fileFilter)) {
        $fileFilter = '*'
    }


    # ----------------------------------------------------------
    # UI vorbereiten
    # ----------------------------------------------------------

    $listBox.Items.Clear()

    $lblCurrent.Text = 'Aktuelle Datei: -'

    $btnSearch.Enabled = $false
    $btnCancel.Enabled = $true

    $script:matchCount         = 0
    $script:cancelMessageShown = $false
    $script:fatalError         = $false


    # ----------------------------------------------------------
    # Suchpfad bestimmen
    # ----------------------------------------------------------

    if ($PSScriptRoot) {
        $searchPath = $PSScriptRoot
    }
    else {
        $searchPath = (Get-Location).Path
    }

    $searchPath = [System.IO.Path]::GetFullPath($searchPath)


    # ----------------------------------------------------------
    # Gemeinsame Objekte für UI und Hintergrund-Runspace
    # ----------------------------------------------------------

    $script:cancelSource =
        [System.Threading.CancellationTokenSource]::new()

    $script:resultQueue =
        [System.Collections.Concurrent.ConcurrentQueue[object]]::new()


    $cancelToken = $script:cancelSource.Token
    $queue       = $script:resultQueue


    # ----------------------------------------------------------
    # Hintergrund-Runspace
    # ----------------------------------------------------------

    $script:searchRunspace =
        [RunspaceFactory]::CreateRunspace()

    $script:searchRunspace.Open()


    $script:searchPowerShell =
        [PowerShell]::Create()

    $script:searchPowerShell.Runspace =
        $script:searchRunspace


    # ==========================================================
    # Worker-Script
    # ==========================================================

    $searchScript = {

        param(
            $path,
            $filter,
            $term,
            $cancelToken,
            $queue
        )


        # ------------------------------------------------------
        # Wildcard-Filter explizit case-insensitive
        # ------------------------------------------------------

        $wildcard =
            [System.Management.Automation.WildcardPattern]::new(
                $filter,
                [System.Management.Automation.WildcardOptions]::IgnoreCase
            )


        # Präfix für relative Pfade
        $pathPrefix = $path

        if (
            -not $pathPrefix.EndsWith(
                [System.IO.Path]::DirectorySeparatorChar.ToString()
            )
        ) {
            $pathPrefix +=
                [System.IO.Path]::DirectorySeparatorChar
        }


        try {

            # --------------------------------------------------
            # Dateien STREAMEND verarbeiten.
            #
            # Anders als:
            #
            #   $files = Get-ChildItem ...
            #
            # wird hier nicht erst die komplette Dateiliste
            # aufgebaut.
            # --------------------------------------------------

            Get-ChildItem `
                -Path $path `
                -Recurse `
                -File `
                -ErrorAction SilentlyContinue |

                ForEach-Object {

                    # ------------------------------------------
                    # Abbruch bereits bei Dateisuche prüfen
                    # ------------------------------------------

                    if ($cancelToken.IsCancellationRequested) {
                        throw [System.OperationCanceledException]::new()
                    }


                    $file = $_


                    # ------------------------------------------
                    # Dateifilter
                    # ------------------------------------------

                    if ($wildcard.IsMatch($file.Name)) {

                        # aktuelle Datei an UI melden
                        [void]$queue.Enqueue(
                            [PSCustomObject]@{
                                Type = 'Progress'
                                Text = $file.FullName
                            }
                        )


                        # --------------------------------------
                        # Datei zeilenweise lesen
                        #
                        # Dadurch kann auch während einer großen
                        # Datei abgebrochen werden.
                        # --------------------------------------

                        $reader = $null

                        try {

                            $reader =
                                [System.IO.StreamReader]::new(
                                    $file.FullName,
                                    $true
                                )


                            $lineNumber = 0


                            while ($true) {

                                if (
                                    $cancelToken.IsCancellationRequested
                                ) {
                                    throw [System.OperationCanceledException]::new()
                                }


                                $line = $reader.ReadLine()


                                if ($null -eq $line) {
                                    break
                                }


                                $lineNumber++


                                # --------------------------------
                                # Literale, case-insensitive Suche
                                # --------------------------------

                                if (
                                    $line.IndexOf(
                                        $term,
                                        [System.StringComparison]::OrdinalIgnoreCase
                                    ) -ge 0
                                ) {

                                    # relativer Dateiname
                                    if (
                                        $file.FullName.StartsWith(
                                            $pathPrefix,
                                            [System.StringComparison]::OrdinalIgnoreCase
                                        )
                                    ) {
                                        $relative =
                                            $file.FullName.Substring(
                                                $pathPrefix.Length
                                            )
                                    }
                                    else {
                                        $relative = $file.Name
                                    }


                                    $text =
                                        "Datei: $relative | " +
                                        "Zeile ${lineNumber}: " +
                                        $line.Trim()


                                    # Treffer in Queue legen
                                    [void]$queue.Enqueue(
                                        [PSCustomObject]@{
                                            Type = 'Match'
                                            Text = $text
                                        }
                                    )
                                }
                            }
                        }

                        catch [System.OperationCanceledException] {

                            # muss zum äußeren Catch weitergegeben
                            # werden
                            throw
                        }

                        catch {

                            # Einzelne Dateien dürfen unlesbar sein
                            # (gesperrt, binär, keine Berechtigung ...)
                            #
                            # Wie im ursprünglichen Skript:
                            # solche Dateien still überspringen.
                        }

                        finally {

                            if ($reader) {
                                $reader.Dispose()
                            }
                        }
                    }
                }


            # Sicherheitshalber noch einmal prüfen
            if ($cancelToken.IsCancellationRequested) {
                throw [System.OperationCanceledException]::new()
            }
        }


        # ------------------------------------------------------
        # Benutzerabbruch
        # ------------------------------------------------------

        catch [System.OperationCanceledException] {

            [void]$queue.Enqueue(
                [PSCustomObject]@{
                    Type = 'Cancelled'
                    Text = $null
                }
            )
        }


        # ------------------------------------------------------
        # Unerwarteter Fehler
        # ------------------------------------------------------

        catch {

            [void]$queue.Enqueue(
                [PSCustomObject]@{
                    Type = 'FatalError'
                    Text = $_.Exception.Message
                }
            )
        }
    }


    # ----------------------------------------------------------
    # Argumente an Worker übergeben
    # ----------------------------------------------------------

    [void]$script:searchPowerShell.AddScript($searchScript)

    [void]$script:searchPowerShell.AddArgument($searchPath)
    [void]$script:searchPowerShell.AddArgument($fileFilter)
    [void]$script:searchPowerShell.AddArgument($searchTerm)
    [void]$script:searchPowerShell.AddArgument($cancelToken)
    [void]$script:searchPowerShell.AddArgument($queue)


    # ----------------------------------------------------------
    # Asynchron starten
    # ----------------------------------------------------------

    try {

        $script:searchAsync =
            $script:searchPowerShell.BeginInvoke()
    }

    catch {

        [void]$listBox.Items.Add(
            "Suche konnte nicht gestartet werden: " +
            $_.Exception.Message
        )

        $btnSearch.Enabled = $true
        $btnCancel.Enabled = $false

        Close-SearchResources

        return
    }


    # ==========================================================
    # UI-Timer
    #
    # Der Timer läuft IM UI-Thread.
    # Er liest nur die threadsichere Queue aus.
    # ==========================================================

    $script:timer =
        New-Object System.Windows.Forms.Timer

    $script:timer.Interval = 100


    $script:timer.Add_Tick({

        # neue Meldungen anzeigen
        Read-WorkerMessages


        # ------------------------------------------------------
        # Ist der Worker fertig?
        # ------------------------------------------------------

        if (
            $script:searchAsync -and
            $script:searchAsync.IsCompleted
        ) {

            # eventuell unmittelbar vor Ende erzeugte Meldungen
            # noch abholen
            Read-WorkerMessages


            # EndInvoke nach abgeschlossener Operation blockiert
            # nicht mehr und sorgt für sauberes Beenden.
            try {

                $null =
                    $script:searchPowerShell.EndInvoke(
                        $script:searchAsync
                    )
            }

            catch {

                if (
                    -not $script:cancelSource.IsCancellationRequested
                ) {

                    [void]$listBox.Items.Add(
                        "Fehler beim Beenden der Suche: " +
                        $_.Exception.Message
                    )

                    $script:fatalError = $true
                }
            }


            # Nach EndInvoke noch einmal Queue leeren
            Read-WorkerMessages


            $wasCancelled =
                $script:cancelSource.IsCancellationRequested


            # Falls der Worker das Cancel-Signal nicht mehr
            # in die Queue legen konnte:
            if (
                $wasCancelled -and
                -not $script:cancelMessageShown
            ) {

                [void]$listBox.Items.Add(
                    'Suche wurde vom Benutzer abgebrochen.'
                )

                $script:cancelMessageShown = $true
            }


            # Keine Treffer
            if (
                -not $wasCancelled -and
                -not $script:fatalError -and
                $script:matchCount -eq 0
            ) {

                [void]$listBox.Items.Add(
                    'Keine Treffer gefunden.'
                )
            }


            # optionales Ende bei Treffern
            if (
                -not $wasCancelled -and
                -not $script:fatalError -and
                $script:matchCount -gt 0
            ) {

                [void]$listBox.Items.Add(
                    "---- Suche beendet: " +
                    "$($script:matchCount) Treffer ----"
                )
            }


            # UI freigeben
            $btnSearch.Enabled = $true
            $btnCancel.Enabled = $false

            $lblCurrent.Text = 'Aktuelle Datei: -'


            # Ressourcen freigeben
            Close-SearchResources
        }
    })


    $script:timer.Start()
})


# ==============================================================
# Fenster schließen
#
# Falls während einer Suche geschlossen wird, Worker ebenfalls
# stoppen und Ressourcen aufräumen.
# ==============================================================

$form.Add_FormClosing({

    if ($script:cancelSource) {

        try {
            $script:cancelSource.Cancel()
        }
        catch {}
    }

    Close-SearchResources -ForceStop
})


# ==============================================================
# UI starten
# ==============================================================

[void]$form.ShowDialog()