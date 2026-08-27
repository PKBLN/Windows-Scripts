# --------------------------------------------------------------
#  File Content Search Pro
#  - case-insensitive file filter
#  - case-insensitive content search
#  - live display of the current file
#  - reliable cancellation
#  - UI remains responsive during the search
# --------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# ==============================================================
# GUI
# ==============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'File Content Search Pro'
$form.Size          = New-Object System.Drawing.Size(860, 680)
$form.StartPosition = 'CenterScreen'


# ---- Label & TextBox: File filter ----------------------------

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text     = 'File filter (e.g. *.txt or *config*):'
$lblFile.Location = New-Object System.Drawing.Point(10, 20)
$lblFile.AutoSize = $true
$form.Controls.Add($lblFile)

$txtFileFilter = New-Object System.Windows.Forms.TextBox
$txtFileFilter.Text     = '*'
$txtFileFilter.Location = New-Object System.Drawing.Point(220, 17)
$txtFileFilter.Size     = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($txtFileFilter)


# ---- Label & TextBox: Search term ----------------------------

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Search term in contents:'
$lblSearch.Location = New-Object System.Drawing.Point(10, 50)
$lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(220, 47)
$txtSearch.Size     = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($txtSearch)


# ---- Button: Search ------------------------------------------

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text     = 'Search'
$btnSearch.Location = New-Object System.Drawing.Point(540, 15)
$btnSearch.Size     = New-Object System.Drawing.Size(100, 60)
$form.Controls.Add($btnSearch)


# ---- Button: Cancel ---------------------------------------

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text     = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(650, 15)
$btnCancel.Size     = New-Object System.Drawing.Size(100, 60)
$btnCancel.Enabled  = $false
$form.Controls.Add($btnCancel)


# ---- Label: Current file -----------------------------------

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Text     = 'Current file: -'
$lblCurrent.Location = New-Object System.Drawing.Point(10, 80)
$lblCurrent.AutoSize = $true
$form.Controls.Add($lblCurrent)


# ---- ListBox: Results ----------------------------------

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 110)
$listBox.Size     = New-Object System.Drawing.Size(820, 500)
$listBox.HorizontalScrollbar = $true
$form.Controls.Add($listBox)


# ==============================================================
# State / global variables
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
# Helper function: release resources
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
# Helper function: process worker messages
#
# IMPORTANT:
# This function is called only from the WinForms UI thread.
# Controls can therefore be modified directly here.
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
                    "Current file: $($message.Text)"
            }


            'Match' {

                $script:matchCount++

                [void]$listBox.Items.Add($message.Text)

                # Automatically scroll to the latest match
                if ($listBox.Items.Count -gt 0) {
                    $listBox.TopIndex =
                        $listBox.Items.Count - 1
                }
            }


            'Cancelled' {

                if (-not $script:cancelMessageShown) {

                    [void]$listBox.Items.Add(
                        'The search was cancelled by the user.'
                    )

                    $script:cancelMessageShown = $true
                }
            }


            'FatalError' {

                [void]$listBox.Items.Add(
                    "Error: $($message.Text)"
                )

                $script:fatalError = $true
            }
        }

        $message = $null
    }
}


# ==============================================================
# Cancel
# ==============================================================

$btnCancel.Add_Click({

    if (
        $script:cancelSource -and
        -not $script:cancelSource.IsCancellationRequested
    ) {

        $script:cancelSource.Cancel()

        $btnCancel.Enabled = $false

        [void]$listBox.Items.Add(
            'Requesting cancellation ...'
        )
    }
})


# ==============================================================
# Search
# ==============================================================

$btnSearch.Add_Click({

    # ----------------------------------------------------------
    # Input
    # ----------------------------------------------------------

    $searchTerm = $txtSearch.Text.Trim()
    $fileFilter = $txtFileFilter.Text.Trim()


    if ([string]::IsNullOrWhiteSpace($searchTerm)) {

        [System.Windows.Forms.MessageBox]::Show(
            'Please enter a search term!',
            'File Content Search',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        return
    }


    # Empty file filter = all files
    if ([string]::IsNullOrWhiteSpace($fileFilter)) {
        $fileFilter = '*'
    }


    # ----------------------------------------------------------
    # Prepare the UI
    # ----------------------------------------------------------

    $listBox.Items.Clear()

    $lblCurrent.Text = 'Current file: -'

    $btnSearch.Enabled = $false
    $btnCancel.Enabled = $true

    $script:matchCount         = 0
    $script:cancelMessageShown = $false
    $script:fatalError         = $false


    # ----------------------------------------------------------
    # Determine the search path
    # ----------------------------------------------------------

    if ($PSScriptRoot) {
        $searchPath = $PSScriptRoot
    }
    else {
        $searchPath = (Get-Location).Path
    }

    $searchPath = [System.IO.Path]::GetFullPath($searchPath)


    # ----------------------------------------------------------
    # Shared objects for the UI and background runspace
    # ----------------------------------------------------------

    $script:cancelSource =
        [System.Threading.CancellationTokenSource]::new()

    $script:resultQueue =
        [System.Collections.Concurrent.ConcurrentQueue[object]]::new()


    $cancelToken = $script:cancelSource.Token
    $queue       = $script:resultQueue


    # ----------------------------------------------------------
    # Background runspace
    # ----------------------------------------------------------

    $script:searchRunspace =
        [RunspaceFactory]::CreateRunspace()

    $script:searchRunspace.Open()


    $script:searchPowerShell =
        [PowerShell]::Create()

    $script:searchPowerShell.Runspace =
        $script:searchRunspace


    # ==========================================================
    # Worker script
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
        # Explicitly case-insensitive wildcard filter
        # ------------------------------------------------------

        $wildcard =
            [System.Management.Automation.WildcardPattern]::new(
                $filter,
                [System.Management.Automation.WildcardOptions]::IgnoreCase
            )


        # Prefix for relative paths
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
            # Process files AS A STREAM.
            #
            # Unlike:
            #
            #   $files = Get-ChildItem ...
            #
            # the complete file list is not built first
            # in this implementation.
            # --------------------------------------------------

            Get-ChildItem `
                -Path $path `
                -Recurse `
                -File `
                -ErrorAction SilentlyContinue |

                ForEach-Object {

                    # ------------------------------------------
                    # Check for cancellation while enumerating files
                    # ------------------------------------------

                    if ($cancelToken.IsCancellationRequested) {
                        throw [System.OperationCanceledException]::new()
                    }


                    $file = $_


                    # ------------------------------------------
                    # File filter
                    # ------------------------------------------

                    if ($wildcard.IsMatch($file.Name)) {

                        # Report the current file to the UI
                        [void]$queue.Enqueue(
                            [PSCustomObject]@{
                                Type = 'Progress'
                                Text = $file.FullName
                            }
                        )


                        # --------------------------------------
                        # Read the file line by line
                        #
                        # This allows cancellation even while reading
                        # a large file.
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
                                # Literal, case-insensitive search
                                # --------------------------------

                                if (
                                    $line.IndexOf(
                                        $term,
                                        [System.StringComparison]::OrdinalIgnoreCase
                                    ) -ge 0
                                ) {

                                    # Relative file name
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
                                        "File: $relative | " +
                                        "Line ${lineNumber}: " +
                                        $line.Trim()


                                    # Add the match to the queue
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

                            # Must be passed on to the outer catch
                            # block
                            throw
                        }

                        catch {

                            # Individual files may be unreadable
                            # (locked, binary, insufficient permissions, etc.)
                            #
                            # As in the original script:
                            # silently skip such files.
                        }

                        finally {

                            if ($reader) {
                                $reader.Dispose()
                            }
                        }
                    }
                }


            # Check once more to be safe
            if ($cancelToken.IsCancellationRequested) {
                throw [System.OperationCanceledException]::new()
            }
        }


        # ------------------------------------------------------
        # User cancellation
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
        # Unexpected error
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
    # Pass arguments to the worker
    # ----------------------------------------------------------

    [void]$script:searchPowerShell.AddScript($searchScript)

    [void]$script:searchPowerShell.AddArgument($searchPath)
    [void]$script:searchPowerShell.AddArgument($fileFilter)
    [void]$script:searchPowerShell.AddArgument($searchTerm)
    [void]$script:searchPowerShell.AddArgument($cancelToken)
    [void]$script:searchPowerShell.AddArgument($queue)


    # ----------------------------------------------------------
    # Start asynchronously
    # ----------------------------------------------------------

    try {

        $script:searchAsync =
            $script:searchPowerShell.BeginInvoke()
    }

    catch {

        [void]$listBox.Items.Add(
            "The search could not be started: " +
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
    # The timer runs IN the UI thread.
    # It only reads the thread-safe queue.
    # ==========================================================

    $script:timer =
        New-Object System.Windows.Forms.Timer

    $script:timer.Interval = 100


    $script:timer.Add_Tick({

        # Display new messages
        Read-WorkerMessages


        # ------------------------------------------------------
        # Has the worker finished?
        # ------------------------------------------------------

        if (
            $script:searchAsync -and
            $script:searchAsync.IsCompleted
        ) {

            # Retrieve any messages generated immediately before
            # completion
            Read-WorkerMessages


            # EndInvoke no longer blocks after the operation has
            # completed and ensures a clean shutdown.
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
                        "Error while completing the search: " +
                        $_.Exception.Message
                    )

                    $script:fatalError = $true
                }
            }


            # Drain the queue once more after EndInvoke
            Read-WorkerMessages


            $wasCancelled =
                $script:cancelSource.IsCancellationRequested


            # In case the worker could no longer add the cancellation
            # signal to the queue:
            if (
                $wasCancelled -and
                -not $script:cancelMessageShown
            ) {

                [void]$listBox.Items.Add(
                    'The search was cancelled by the user.'
                )

                $script:cancelMessageShown = $true
            }


            # No matches
            if (
                -not $wasCancelled -and
                -not $script:fatalError -and
                $script:matchCount -eq 0
            ) {

                [void]$listBox.Items.Add(
                    'No matches found.'
                )
            }


            # Optional completion message when matches were found
            if (
                -not $wasCancelled -and
                -not $script:fatalError -and
                $script:matchCount -gt 0
            ) {

                [void]$listBox.Items.Add(
                    "---- Search completed: " +
                    "$($script:matchCount) matches ----"
                )
            }


            # Re-enable the UI
            $btnSearch.Enabled = $true
            $btnCancel.Enabled = $false

            $lblCurrent.Text = 'Current file: -'


            # Release resources
            Close-SearchResources
        }
    })


    $script:timer.Start()
})


# ==============================================================
# Close the window
#
# If the window is closed during a search, also stop the worker
# and release its resources.
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
# Start the UI
# ==============================================================

[void]$form.ShowDialog()
