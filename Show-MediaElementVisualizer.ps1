<#
.SYNOPSIS
    A unified 3D media visualizer that combines multiple display styles into a single script.
.DESCRIPTION
    This script launches a comprehensive GUI to select image and video files, and then allows the user
    to choose from a variety of 3D visualization styles. Once a style is selected, the script renders
    the media onto the corresponding 3D objects in a full-screen WPF window.

    Available styles include:
    - Rotating Cube, Sphere, Star, and Funnel
    - Multi-object scenes like Floating Cubes, Floating Stars, and Butterfly Effect
    - Complex geometries like Carousels, Vortexes, Wagon Wheels, and various Funnels

    This script uses the built-in Windows MediaElement for video playback, so video format support is
    limited to codecs installed on the local system (e.g., MP4, WMV, AVI). It centralizes all common
    functionality, such as the UI, media handling, and error reporting, into a single, manageable file.
.EXAMPLE
    PS C:\> .\Show-MediaElementVisualizer.ps1

    Launches the media and style selection GUI. After selecting files and a display style,
    clicking "Play" will launch the chosen 3D visualization.
.NOTES
    Name:           Show-MediaElementVisualizer.ps1
    Version:        1.0.0, 12/29/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- Script Metadata (for Script Launcher) ---
$ExternalButtonName = "Unified 3D Visualizer`n(MediaElement)"
$ScriptDescription = "A single script that combines all MediaElement-based 3D visualizers, selectable via a GUI."
$RequiredExecutables = @() # No external executables needed

# This must be outside the loop to persist across "Redo" clicks for proper cleanup.
$animationLoop = $null

# --- Main Application Loop ---
while ($true) {
    # CRITICAL: Cleanup from the previous run. If an animation loop was attached,
    # it must be removed before starting a new visualization to prevent conflicts.
    if ($animationLoop) {
        [System.Windows.Media.CompositionTarget]::remove_Rendering($animationLoop)
        $animationLoop = $null
    }
    #region --- File and Style Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Unified MediaElement Visualizer - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 800)
    $SelectForm.StartPosition = "CenterScreen"

    # --- File Selection Controls ---
    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $AddBubblesCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Add Bubbles (Aquarium)"; AutoSize = $true; Location = '150, 70'; Checked = $true; Visible = $false }
    $AddWaterCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Add Water (Aquarium)"; AutoSize = $true; Location = '320, 70'; Checked = $true; Visible = $false }
    $NightSkyCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Dark Night Sky"; AutoSize = $true; Location = '150, 70'; Checked = $true; Visible = $false }
    $TwinkleCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Twinkling Stars"; AutoSize = $true; Location = '320, 70'; Checked = $true; Visible = $false }

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{ Location = '10, 95'; Size = '760, 200'; Anchor = 'Top, Left, Right'; AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65 }
    
    $SelectForm.Controls.AddRange(@($BrowseButton, $FolderPathTextBox, $RecursiveCheckBox, $TransparentCheckbox, $AddBubblesCheckbox, $AddWaterCheckbox, $NightSkyCheckbox, $TwinkleCheckbox, $SelectAllCheckbox, $DataGridView))

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true }
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Play Selected"; Location = '600, 40'; Size = '170, 30' }
    $SelectForm.Controls.Add($PlayButton)

    # --- Display Style Selection ---
    $StyleGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Display Style"; Location = '10, 305'; Size = '760, 195'; Anchor = 'Top, Left, Right' }
    $SelectForm.Controls.Add($StyleGroupBox)
    
    $visualizationStyles = @(
        "Aquarium", "Butterfly Effect", "Carousel", "Concentric Funnel", "Curved Vortex", "Faceted Sphere (Multi)",
        "Faceted Sphere (Single)", "Floating Cubes", "Floating Spheres", "Floating Stars", "Funnel (Multi)", "Funnel (Single)", "Media Flow Funnel", "Pie 3D",
        "Pinwheel", "Pulsing Star", "Reindeer Sleigh", "Roller Coaster", "Rotating Cube",
        "Rotating Star", "Scrolling Horizontal", "Scrolling Vertical", "Sphere", "Sphere Vortex", "Wagon Wheel"
    ) | Sort-Object | Get-Unique

    $colWidth = 180; $rowHeight = 25; $x = 15; $y = 20
    foreach ($styleName in $visualizationStyles) {
        $rb = New-Object System.Windows.Forms.RadioButton -Property @{
            Text = $styleName; Location = "$x, $y"; AutoSize = $true; Tag = $styleName.Replace(" ", "").Replace("(", "").Replace(")", "")
        }
        $rb.Add_CheckedChanged({
            param($sender, $e)
            if ($sender.Checked) {
                $isAquarium = ($sender.Tag -eq "Aquarium")
                $isReindeer = ($sender.Tag -eq "ReindeerSleigh")

                $AddBubblesCheckbox.Visible = $isAquarium
                $AddWaterCheckbox.Visible = $isAquarium
                $NightSkyCheckbox.Visible = $isReindeer
                $TwinkleCheckbox.Visible = $isReindeer
            }
        })
        if ($x -eq 15 -and $y -eq 20) { $rb.Checked = $true } # Default to first
        $StyleGroupBox.Controls.Add($rb)
        $x += $colWidth
        if ($x + $colWidth -gt $StyleGroupBox.Width) { $x = 15; $y += $rowHeight }
    }

    # --- Text Overlay Controls ---
    $TextGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 510'; Size = '760, 250'; Anchor = 'Top, Left, Right' }
    $SelectForm.Controls.Add($TextGroupBox)

    $TextOptionsGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Options"; Location = '10, 20'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $TextOptionsGroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $TextGroupBox.Controls.Add($TextOptionsGroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '140, 20'; Size = '455, 210'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center' }
    $TextGroupBox.Controls.Add($TextBox)

    $FilenamePreviewLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Filename.mp4"
        Location = '140, 20'
        Size = '455, 210' # Match the TextBox size
        Visible = $false
        TextAlign = 'MiddleCenter'
    }
    $TextGroupBox.Controls.Add($FilenamePreviewLabel)

    $FontControlsGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Style"; Location = '600, 20'; Size = '150, 210'; Visible = $false }
    $TextGroupBox.Controls.Add($FontControlsGroupBox)

    $CurrentColor = New-Object System.Windows.Forms.Label -Property @{ Text = "Color:"; Location = '10, 25'; AutoSize = $true }
    $ColorExample = New-Object System.Windows.Forms.Label -Property @{ Text = "    "; Location = '50, 25'; AutoSize = $true; BackColor = [System.Drawing.Color]::Black; BorderStyle = 'FixedSingle' }
    $SelectColorButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change"; Location = '80, 20'; Size = '60, 25' }
    $SizeLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Size:"; AutoSize = $true; Location = '10, 57' }
    $NumericUpDown = New-Object System.Windows.Forms.NumericUpDown -Property @{ Location = '50, 55'; Size = '50, 20'; Minimum = 8; Maximum = 72; Value = 24 }
    $FontButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change Font"; Location = '10, 90'; Size = '130, 25' }
    $ItalicCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Italic"; Location = '10, 130'; Size = '60, 20'; Checked = $false }
    $BoldCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Bold"; Location = '80, 130'; Size = '60, 20'; Checked = $true }
    
    $FontControlsGroupBox.Controls.AddRange(@($CurrentColor, $ColorExample, $SelectColorButton, $SizeLabel, $NumericUpDown, $FontButton, $ItalicCheckbox, $BoldCheckbox))

    # --- Form Logic ---
    $formState = @{ TextColor = [System.Drawing.Color]::Black; FontFamily = "Arial" }
    
    $textOverlayEvent = {
        $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked
        $isCustomText = $RadioButton3.Checked
        $isFilename = $RadioButton2.Checked
        $TextBox.Visible = $isCustomText
        $FontControlsGroupBox.Visible = $isTextVisible
        $FilenamePreviewLabel.Visible = $isFilename
    }
    $RadioButton1.Add_Click($textOverlayEvent)
    $RadioButton2.Add_Click($textOverlayEvent)
    $RadioButton3.Add_Click($textOverlayEvent)

    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.TextColor = $colorDialog.Color
            $ColorExample.BackColor = $formState.TextColor
            $TextBox.ForeColor = $formState.TextColor
            $FilenamePreviewLabel.ForeColor = $formState.TextColor
        }
    })

    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
            $FilenamePreviewLabel.Font = $newFont
        } catch {
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style)
            $FilenamePreviewLabel.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }

    $FontButton.Add_Click({
        $fontDialog = New-Object System.Windows.Forms.FontDialog
        try {
            $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value)
        } catch { $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12) }

        if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.FontFamily = $fontDialog.Font.Name
            $FontButton.Text = $formState.FontFamily
            $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
            $BoldCheckbox.Checked = $fontDialog.Font.Bold
            $ItalicCheckbox.Checked = $fontDialog.Font.Italic
            & $updateTextBoxFont
        }
    })
    
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont)
    $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont)
    $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)
    & $updateTextBoxFont

    $SelectAllCheckbox.Add_CheckedChanged({
        $isChecked = $SelectAllCheckbox.Checked
        foreach ($row in $DataGridView.Rows) { $row.Cells["Select"].Value = $isChecked }
        $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    })

    $BrowseButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SelectedPath = $FolderBrowser.SelectedPath
            $FolderPathTextBox.Text = $SelectedPath
            $DataGridView.Rows.Clear()

            $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
            $VideoExtensions = "*.mp4", "*.m4v", "*.wmv", "*.avi", "*.mpg", "*.mpeg"
            $AllowedExtensions = $ImageExtensions + $VideoExtensions
            
            $gciParams = @{ File = $true; Include = $AllowedExtensions }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true } 
            else { $gciParams.Path = Join-Path $SelectedPath "*" }

            Get-ChildItem @gciParams | ForEach-Object { $DataGridView.Rows.Add($false, $_.Name, $_.FullName) | Out-Null }
            $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $filePath = $DataGridView.Rows[$e.RowIndex].Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                try { Start-Process $filePath } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", "OK", "Error") }
            }
        }
    })

    $PlayButton.Add_Click({
        $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
        if ($selectedFiles.Count -gt 0) {
            $formState.SelectedFiles = [System.Collections.ArrayList]::new($selectedFiles)
            $formState.UseTransparentEffect = $TransparentCheckbox.Checked
            $formState.AddBubbles = $AddBubblesCheckbox.Checked
            $formState.AddWater = $AddWaterCheckbox.Checked
            $formState.NightSky = $NightSkyCheckbox.Checked
            $formState.TwinklingStars = $TwinkleCheckbox.Checked
            
            $selectedStyleRB = $StyleGroupBox.Controls | Where-Object { $_.Checked }
            $formState.VisualizationStyle = if ($selectedStyleRB) { $selectedStyleRB.Tag } else { "RotatingCube" }

            switch ($true) {
                { $RadioButton1.Checked } { $formState.RbSelection = "Hidden"; break }
                { $RadioButton2.Checked } { $formState.RbSelection = "Filename"; break }
                { $RadioButton3.Checked } { $formState.RbSelection = "Custom"; break }
            }
            $formState.CustomText = $TextBox.Text
            $formState.FontSize = $NumericUpDown.Value
            $formState.IsBold = $BoldCheckbox.Checked
            $formState.IsItalic = $ItalicCheckbox.Checked
            $SelectForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $null = $SelectForm.ShowDialog()
    $SelectForm.Dispose()

    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break 
    }
    #endregion --- File and Style Selection Form ---

    # --- Central State and Media Handling ---
    $SyncHash = [hashtable]::Synchronized(@{
        SelectedFiles        = $formState.SelectedFiles
        UseTransparentEffect = $formState.UseTransparentEffect
        AddBubbles           = $formState.AddBubbles
        AddWater             = $formState.AddWater
        NightSky             = $formState.NightSky
        TwinklingStars       = $formState.TwinklingStars
        VisualizationStyle   = $formState.VisualizationStyle
        CurrentIndex         = -1
        BadMediaFiles        = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        PlayerStates         = [hashtable]::Synchronized(@{})
        Paused               = $false
        ControlsHidden       = $false
        RedoClicked          = $false
        LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
        SpeedMultiplier      = 1.0
        # Text Overlay Settings
        RbSelection          = $formState.RbSelection
        CustomText           = $formState.CustomText
        TextColor            = $formState.TextColor
        FontSize             = $formState.FontSize
        FontFamily           = $formState.FontFamily
        IsBold               = $formState.IsBold
        IsItalic             = $formState.IsItalic
    })

    $globalIndexLock = New-Object object
    function Get-NextMediaIndex {
        [System.Threading.Monitor]::Enter($globalIndexLock)
        try {
            $fileCount = $SyncHash.SelectedFiles.Count
            if ($fileCount -eq 0) { return -1 }

            # Prevent an infinite loop if all files are blacklisted.
            if ($SyncHash.BadMediaFiles.Count -ge $fileCount) {
                Write-Warning "All available media files have failed. No more media to display."
                return -1
            }

            $startIndex = ($SyncHash.CurrentIndex + 1) % $fileCount
            $currentIndex = $startIndex
            
            do {
                $filePath = $SyncHash.SelectedFiles[$currentIndex]
                if (-not $SyncHash.BadMediaFiles.Contains($filePath)) {
                    $SyncHash.CurrentIndex = $currentIndex
                    return $currentIndex
                }
                $currentIndex = ($currentIndex + 1) % $fileCount
            } while ($currentIndex -ne $startIndex)

            return -1 # All files are blacklisted
        } finally {
            [System.Threading.Monitor]::Exit($globalIndexLock)
        }
    }

    function Handle-MediaFailure {
        param([string]$PlayerKey, [string]$Reason)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $playerState -or $playerState.IsFailed) { return }
        $playerState.IsFailed = $true
        # Stop the stopwatch immediately to prevent it from running into the next media's lifecycle.
        $playerState.PlaybackStopwatch.Stop()
        # Also stop the image timer, if it exists, to prevent orphaned timers.
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }

        # 1. Log the failure immediately. This must happen before any dispatched UI work.
        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for player '$PlayerKey' in style '$($SyncHash.VisualizationStyle)' (File: '$fileName'). Reason: $Reason. Attempting to replace."

        # 2. Add the bad file to the blacklist.
        if ($playerState.CurrentSource -and -not $SyncHash.BadMediaFiles.Contains($playerState.CurrentSource.LocalPath)) {
            $SyncHash.BadMediaFiles.Add($playerState.CurrentSource.LocalPath) | Out-Null
        }

        # 3. Dispatch UI cleanup to the UI thread.
        $SyncHash.Window.Dispatcher.Invoke([action]{
            if ($playerState.CurrentMediaElement) {
                # Detach handlers to prevent zombie events from the failed element interfering with the next one.
                if ($playerState.MediaEndedHandler) { try { $playerState.CurrentMediaElement.remove_MediaEnded($playerState.MediaEndedHandler) } catch {} }
                if ($playerState.MediaOpenedHandler) { try { $playerState.CurrentMediaElement.remove_MediaOpened($playerState.MediaOpenedHandler) } catch {} }
                if ($playerState.MediaFailedHandler) { try { $playerState.CurrentMediaElement.remove_MediaFailed($playerState.MediaFailedHandler) } catch {} }

                $playerState.CurrentMediaElement.Stop()
                $playerState.CurrentMediaElement.Close()
                # Explicitly null out the reference to the failed element.
                $playerState.CurrentMediaElement = $null
            }
            if ($playerState.ContentPresenter) { $playerState.ContentPresenter.Content = $null }
            if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
        })

        # 4. Asynchronously schedule the next media to be loaded.
        $recoveryScriptBlock = { Start-NextMedia -PlayerKey $PlayerKey }
        $null = $SyncHash.Window.Dispatcher.InvokeAsync($recoveryScriptBlock.GetNewClosure())
    }

    function Handle-MediaEnded {
        param([string]$PlayerKey)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if ($pState.IsFailed) { return } 

        $pState.PlaybackStopwatch.Stop()
        # If media "ends" almost instantly and it's not an image, it's a silent failure.
        if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500 -and -not $pState.IsImage) {
            Handle-MediaFailure -PlayerKey $PlayerKey -Reason "Playback ended instantly (bad codec)."
            return
        }
        # Normal completion, start the next media.
        Start-NextMedia -PlayerKey $PlayerKey
    }

    function Handle-MediaOpened {
        param([string]$PlayerKey, $EventArgs)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $pState) { return }
        $pState.IsFailed = $false
        $pState.PlaybackStopwatch.Restart()
        if ($pState.MediaHostGrid) { $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent }
        # Store the expected duration for the watchdog timer.
        if ($EventArgs.NaturalDuration.HasTimeSpan) { $pState.ExpectedDuration = $EventArgs.NaturalDuration.TimeSpan } else { $pState.ExpectedDuration = $null }
        # Clear the assignment time; this signifies that the media has successfully opened.
        $pState.SourceAssignmentTime = $null
        
        if (-not $pState.IsImage -and -not $EventArgs.NaturalDuration.HasTimeSpan) {
            Handle-MediaFailure -PlayerKey $PlayerKey -Reason "Invalid duration or codec."
        }
    }

    function Start-NextMedia {
        param(
            [string]$PlayerKey
        )
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        # Defensive check: If the player state doesn't exist for any reason, log it and exit the function.
        if (-not $playerState) {
            Write-Warning "Attempted to start media for a non-existent player key: '$PlayerKey'. Aborting."
            return
        }

        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }

        # Ensure the player state object has the required properties. This is a robust way to handle initialization.
        if (-not $playerState.PSObject.Properties['CurrentSource']) {
            $playerState | Add-Member -MemberType NoteProperty -Name 'CurrentSource' -Value $null
            $playerState | Add-Member -MemberType NoteProperty -Name 'IsImage' -Value $false
            $playerState | Add-Member -MemberType NoteProperty -Name 'PlaybackStopwatch' -Value (New-Object System.Diagnostics.Stopwatch)
            $playerState | Add-Member -MemberType NoteProperty -Name 'IsFailed' -Value $false
            $playerState | Add-Member -MemberType NoteProperty -Name 'ExpectedDuration' -Value $null
            $playerState | Add-Member -MemberType NoteProperty -Name 'SourceAssignmentTime' -Value $null
            $playerState | Add-Member -MemberType NoteProperty -Name 'ImageCycleStartTime' -Value $null
        }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false
        $playerState.ExpectedDuration = $null # Reset for the new media
        $playerState.ImageCycleStartTime = $null # Reset for the new media
        $playerState.SourceAssignmentTime = [datetime]::UtcNow # Mark the time we tried to start this media

        if ($SyncHash.RbSelection -eq "Filename" -and $playerState.OverlayTextBlock) {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        $playerState.IsImage = $ImageExtensions -contains $extension

        try {
            if ($ImageExtensions -contains $extension) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit(); $bitmap.Freeze()
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $image
                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $PlayerKey }
                $timer.Add_Tick({ $t = $args[0]; $key = $t.Tag; $t.Stop(); Start-NextMedia -PlayerKey $key })
                $playerState.MediaTimer = $timer; $playerState.ImageCycleStartTime = [datetime]::UtcNow; $timer.Start()
                # Manually clear the assignment time for images, as they don't fire a MediaOpened event.
                # This prevents the watchdog from incorrectly flagging them as "failed to open".
                $playerState.SourceAssignmentTime = $null
            } else { # Video
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource; Tag = $PlayerKey
                }

                # Retrieve the handlers that were set up in the visualization's setup function.
                $mediaEndedHandler = $playerState.MediaEndedHandler
                $mediaOpenedHandler = $playerState.MediaOpenedHandler
                $mediaFailedHandler = $playerState.MediaFailedHandler

                # Attach the handlers. They should already be closures from the setup function.
                $mediaElement.Add_MediaEnded($mediaEndedHandler)
                $mediaElement.Add_MediaOpened($mediaOpenedHandler)
                $mediaElement.Add_MediaFailed($mediaFailedHandler)
                
                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.CurrentMediaElement = $mediaElement
                $mediaElement.Play()
                if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
            }
        } catch {
            Handle-MediaFailure -PlayerKey $PlayerKey -Reason $_.Exception.Message
        }
    }

    function Cleanup-Visualization {
        param($SyncHash)
        if ($SyncHash.PlayerStates) {
            foreach ($key in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$key]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.CurrentMediaElement) {
                    # Detach handlers to prevent them from firing during cleanup
                    if ($playerState.MediaEndedHandler) { try { $playerState.CurrentMediaElement.remove_MediaEnded($playerState.MediaEndedHandler) } catch {} }
                    if ($playerState.MediaOpenedHandler) { try { $playerState.CurrentMediaElement.remove_MediaOpened($playerState.MediaOpenedHandler) } catch {} }
                    if ($playerState.MediaFailedHandler) { try { $playerState.CurrentMediaElement.remove_MediaFailed($playerState.MediaFailedHandler) } catch {} }
                    $playerState.CurrentMediaElement.Close()
                }
            }
            $SyncHash.PlayerStates.Clear()
        }
    }

    # --- Geometry Generation Functions ---
    function New-SphereMesh {
        param([double]$radius = 1.0, [int]$slices = 64, [int]$stacks = 32)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        for ($stack = 0; $stack -le $stacks; $stack++) {
            $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks; $y = $radius * [Math]::Sin($phi); $r = $radius * [Math]::Cos($phi)
            for ($slice = 0; $slice -le $slices; $slice++) {
                $theta = $slice * 2 * [Math]::PI / $slices; $x = $r * [Math]::Cos($theta); $z = $r * [Math]::Sin($theta)
                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
                $mesh.TextureCoordinates.Add([System.Windows.Point]::new($slice / $slices, $stack / $stacks))
            }
        }
        for ($stack = 0; $stack -lt $stacks; $stack++) {
            for ($slice = 0; $slice -lt $slices; $slice++) {
                $i0 = $stack * ($slices + 1) + $slice; $i1 = ($stack + 1) * ($slices + 1) + $slice
                $mesh.TriangleIndices.Add($i0); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i0 + 1)
                $mesh.TriangleIndices.Add($i0 + 1); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i1 + 1)
            }
        }
        return $mesh
    }

    function New-ConeMesh {
        param([double]$radius = 1.5, [double]$height = 3.0, [int]$slices = 64)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $apexY = $height; $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $apexY, 0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0))
        $baseY = 0
        for ($i = 0; $i -le $slices; $i++) {
            $theta = $i * 2 * [Math]::PI / $slices; $x = $radius * [Math]::Cos($theta); $z = $radius * [Math]::Sin($theta)
            $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $baseY, $z)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $slices, 1))
        }
        $baseCenterIndex = $mesh.Positions.Count; $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $baseY, 0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0.5))
        for ($i = 0; $i -lt $slices; $i++) { $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add($i + 2); $mesh.TriangleIndices.Add($i + 1) }
        for ($i = 0; $i -lt $slices; $i++) { $mesh.TriangleIndices.Add($baseCenterIndex); $mesh.TriangleIndices.Add($i + 1); $mesh.TriangleIndices.Add($i + 2) }
        return $mesh
    }

    function New-WagonWheelSliceModel {
        param([double]$radius=1.5, [double]$height=0.5, [double]$startAngleDeg=0, [double]$sliceAngleDeg=45, [int]$segments=8)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $h2 = $height / 2.0
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0,-$h2,0)) | Out-Null; $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0,$h2,0)) | Out-Null
        $startAngleRad = $startAngleDeg * [Math]::PI / 180.0; $x0 = $radius * [Math]::Cos($startAngleRad); $z0 = $radius * [Math]::Sin($startAngleRad)
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x0,-$h2,$z0)) | Out-Null; $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x0,$h2,$z0)) | Out-Null
        $outerFaceStartIdx = $mesh.Positions.Count
        for ($i = 0; $i -le $segments; $i++) {
            $currentAngleRad = ($startAngleDeg + ($sliceAngleDeg * $i / $segments)) * [Math]::PI / 180.0
            $x = $radius * [Math]::Cos($currentAngleRad); $z = $radius * [Math]::Sin($currentAngleRad)
            $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x,-$h2,$z)) | Out-Null; $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x,$h2,$z)) | Out-Null
        }
        for ($i = 0; $i -lt $outerFaceStartIdx; $i++) { $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,0)) | Out-Null }
        for ($i = 0; $i -le $segments; $i++) { $u = 1 - ($i / $segments); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u,1)) | Out-Null; $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u,0)) | Out-Null }
        $mesh.TriangleIndices.Add(1) | Out-Null; $mesh.TriangleIndices.Add($outerFaceStartIdx+1) | Out-Null; $mesh.TriangleIndices.Add(3) | Out-Null
        $mesh.TriangleIndices.Add(0) | Out-Null; $mesh.TriangleIndices.Add(2) | Out-Null; $mesh.TriangleIndices.Add($outerFaceStartIdx) | Out-Null
        $mesh.TriangleIndices.Add(0) | Out-Null; $mesh.TriangleIndices.Add(3) | Out-Null; $mesh.TriangleIndices.Add(1) | Out-Null; $mesh.TriangleIndices.Add(0) | Out-Null; $mesh.TriangleIndices.Add(2) | Out-Null; $mesh.TriangleIndices.Add(3) | Out-Null
        $endIdx = $mesh.Positions.Count - 2; $mesh.TriangleIndices.Add(0) | Out-Null; $mesh.TriangleIndices.Add(1) | Out-Null; $mesh.TriangleIndices.Add($endIdx+1) | Out-Null; $mesh.TriangleIndices.Add(0) | Out-Null; $mesh.TriangleIndices.Add($endIdx+1) | Out-Null; $mesh.TriangleIndices.Add($endIdx) | Out-Null
        
        $otherFacesMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $otherFacesMesh.Positions = $mesh.Positions; $otherFacesMesh.TextureCoordinates = $mesh.TextureCoordinates
        for ($i = 0; $i -lt 24; $i++) { $otherFacesMesh.TriangleIndices.Add($mesh.TriangleIndices[$i]) | Out-Null }

        $outerFaceMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $outerFaceMesh.Positions = $mesh.Positions; $outerFaceMesh.TextureCoordinates = $mesh.TextureCoordinates
        for ($i = 0; $i -lt $segments; $i++) {
            $idx0 = $outerFaceStartIdx + ($i * 2); $idx1 = $idx0 + 1; $idx2 = $idx0 + 2; $idx3 = $idx0 + 3
            $outerFaceMesh.TriangleIndices.Add($idx0) | Out-Null; $outerFaceMesh.TriangleIndices.Add($idx1) | Out-Null; $outerFaceMesh.TriangleIndices.Add($idx3) | Out-Null
            $outerFaceMesh.TriangleIndices.Add($idx0) | Out-Null; $outerFaceMesh.TriangleIndices.Add($idx3) | Out-Null; $outerFaceMesh.TriangleIndices.Add($idx2) | Out-Null
        }

        $otherFacesModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($otherFacesMesh, (New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::DarkSlateGray)))
        $outerFaceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($outerFaceMesh, $null)
        $modelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup; $modelGroup.Children.Add($otherFacesModel); $modelGroup.Children.Add($outerFaceModel)
        return @{ OuterFaceModel = $outerFaceModel; FullSliceModel = $modelGroup }
    }

    function New-SpiralingPanelMesh {
        param(
            # Parameters for Funnel and FunnelSingle
            [double]$startRadius = 0,
            [double]$endRadius = 0,
            [double]$height = 0,
            [double]$arcAngle = 0, # The angular width of the panel in degrees
            [double]$twistAngle = 0, # The total twist from top to bottom in degrees
            [int]$stacks = 50,
            [int]$slices = 2,
            # Parameters for ConcentricFunnel
            [System.Windows.Media.Media3D.Point3D]$p1, [System.Windows.Media.Media3D.Point3D]$p2, [System.Windows.Media.Media3D.Point3D]$p3, [System.Windows.Media.Media3D.Point3D]$p4
        )

        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $arcAngleRad = $arcAngle * [Math]::PI / 180.0
        $twistAngleRad = $twistAngle * [Math]::PI / 180.0

        for ($j = 0; $j -le $stacks; $j++) {
            $v = $j / $stacks # Vertical progress (0 to 1)

            # Interpolate radius, Y position, and twist for the current stack
            $v_eased = $v * $v # Use an ease-in quadratic function for a proper funnel curve
            $currentRadius = $startRadius - $v * ($startRadius - $endRadius)
            $currentY = ($height / 2) - $v_eased * $height
            $currentTwist = $v * $twistAngleRad

            for ($i = 0; $i -le $slices; $i++) {
                $u = $i / $slices # Horizontal progress (0 to 1)
                $theta = $currentTwist + ($u * $arcAngleRad) - ($arcAngleRad / 2.0)
                $x = $currentRadius * [Math]::Cos($theta)
                $z = $currentRadius * [Math]::Sin($theta)

                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $currentY, $z)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u, $v))
            }
        }

        if ($p1) { # If points are provided, use them (for Concentric Funnel)
            $mesh.Positions.Clear(); $mesh.TextureCoordinates.Clear()
            $mesh.Positions.Add($p1); $mesh.Positions.Add($p2); $mesh.Positions.Add($p3); $mesh.Positions.Add($p4)
            $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,1)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,1))
            $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(1); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(3)
        }
        else { # Otherwise, create triangles for the spiraling panel
            # Create triangle indices
            for ($j = 0; $j -lt $stacks; $j++) {
                for ($i = 0; $i -lt $slices; $i++) {
                    $row1 = $j * ($slices + 1); $row2 = ($j + 1) * ($slices + 1)
                    $mesh.TriangleIndices.Add($row1 + $i); $mesh.TriangleIndices.Add($row2 + $i); $mesh.TriangleIndices.Add($row1 + $i + 1)
                    $mesh.TriangleIndices.Add($row1 + $i + 1); $mesh.TriangleIndices.Add($row2 + $i); $mesh.TriangleIndices.Add($row2 + $i + 1)
                }
            }
        }
        $mesh.Freeze(); return $mesh
    }

    function New-CurvedPanelMesh {
        param(
            [double]$width = 2.0,
            [double]$height = 1.0,
            [double]$curveDepth = 0.5,
            [int]$widthSegments = 20,
            [int]$heightSegments = 2
        )

        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

        for ($j = 0; $j -le $heightSegments; $j++) {
            $y = $height / 2 - ($j / $heightSegments) * $height
            for ($i = 0; $i -le $widthSegments; $i++) {
                $x = -$width / 2 + ($i / $widthSegments) * $width
                $normalizedX = $x / ($width / 2) # Normalize x from -1 to 1
                $z = $curveDepth * ($normalizedX * $normalizedX)
                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
                $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $widthSegments, $j / $heightSegments))
            }
        }

        for ($j = 0; $j -lt $heightSegments; $j++) {
            for ($i = 0; $i -lt $widthSegments; $i++) {
                $row1 = $j * ($widthSegments + 1); $row2 = ($j + 1) * ($widthSegments + 1)
                $mesh.TriangleIndices.Add($row1 + $i); $mesh.TriangleIndices.Add($row1 + $i + 1); $mesh.TriangleIndices.Add($row2 + $i + 1)
                $mesh.TriangleIndices.Add($row1 + $i); $mesh.TriangleIndices.Add($row2 + $i + 1); $mesh.TriangleIndices.Add($row2 + $i)
            }
        }
        return $mesh
    }

    function New-PathRibbonMesh {
        param([array]$PathData, [int]$segments = 400, [double]$width = 0.5, [double]$verticalOffset = 0)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $worldUp = [System.Windows.Media.Media3D.Vector3D]::new(0, 1, 0)
        for ($i = 0; $i -le $segments; $i++) {
            $pointData = $PathData[$i]
            $p1 = $pointData.Point
            $up = $pointData.Up
            $normal = $pointData.Normal

            $offsetVector = $up * $verticalOffset
            $mesh.Positions.Add(($p1 + ($normal * $width / 2) + $offsetVector)); $mesh.Positions.Add(($p1 - ($normal * $width / 2) + $offsetVector))
        }
        for ($i = 0; $i -lt $segments; $i++) {
            $i0 = $i * 2; $i1 = $i0 + 1; $i2 = $i0 + 2; $i3 = $i0 + 3
            $mesh.TriangleIndices.Add($i0); $mesh.TriangleIndices.Add($i2); $mesh.TriangleIndices.Add($i1)
            $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i2); $mesh.TriangleIndices.Add($i3)
        }
        $mesh.Freeze(); return $mesh
    }

    function New-CoasterTrackModelGroup {
        param([array]$PathData, [int]$segments = 800)
        $trackModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
        $tieWidth = 0.8; $tieThickness = 0.05
        $tiesMesh = New-PathRibbonMesh -PathData $PathData -segments $segments -width $tieWidth -verticalOffset (-$tieThickness / 2)
        $tiesMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::SaddleBrown)
        $tiesModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $tiesMesh; Material = $tiesMaterial; BackMaterial = $tiesMaterial }
        $trackModelGroup.Children.Add($tiesModel)
        $railMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::DarkSlateGray)
        return $trackModelGroup
    }

    function New-PieSliceModel {
        param([System.Windows.Point]$center, [double]$radius, [double]$startAngleDeg, [double]$sliceAngleDeg, [double]$thickness = 0.1)
        $endAngleDeg = $startAngleDeg + $sliceAngleDeg; $startAngleRad = $startAngleDeg * [Math]::PI / 180.0; $endAngleRad = $endAngleDeg * [Math]::PI / 180.0; $halfThick = $thickness / 2.0
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $arcPoints = New-Object 'System.Collections.Generic.List[System.Windows.Point]'; $arcPoints.Add($center) | Out-Null
    # Determine the number of segments based on the slice angle to ensure a smooth curve.
    $numArcSegments = [double][Math]::Max(1, [Math]::Ceiling($sliceAngleDeg / 2.0))
        for ($i = 0; $i -le $numArcSegments; $i++) {
            $angle = $startAngleRad + ($i / $numArcSegments) * ($endAngleRad - $startAngleRad)
        [double]$pointX = $center.X + $radius * [Math]::Cos($angle); [double]$pointY = $center.Y + $radius * [Math]::Sin($angle)
            $arcPoints.Add((New-Object System.Windows.Point($pointX, $pointY))) | Out-Null
        }
        $frontBaseIndex = $mesh.Positions.Count
        foreach ($p in $arcPoints) { 
            $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, $halfThick)) | Out-Null
            $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5)) | Out-Null
        }
        $backBaseIndex = $mesh.Positions.Count
        foreach ($p in $arcPoints) { 
            $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, -$halfThick)) | Out-Null
            $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5)) | Out-Null
        }
        for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) {
            $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $i + 1); $mesh.TriangleIndices.Add($frontBaseIndex + $i)
        }
        for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) {
            $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $i); $mesh.TriangleIndices.Add($backBaseIndex + $i + 1)
        }
        for ($i = 1; $i -lt $arcPoints.Count; $i++) {
            $p1_front = $frontBaseIndex + $i; $p2_front = $frontBaseIndex + $i + 1; $p1_back = $backBaseIndex + $i;  $p2_back = $backBaseIndex + $i + 1
            $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p1_back); $mesh.TriangleIndices.Add($p2_back)
            $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p2_back); $mesh.TriangleIndices.Add($p2_front)
        }
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1)
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1); $mesh.TriangleIndices.Add($frontBaseIndex + 1)
        $lastIdx = $arcPoints.Count -1
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx)
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex)
        return (New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $mesh })
    }

    # --- Visualization-Specific Setup Functions ---


    function New-FacetedSphereModels {
        param([double]$radius = 1.5, [int]$slices = 8, [int]$stacks = 4)
        $facets = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.GeometryModel3D]

        $allVertices = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.Point3D]
        for ($stack = 0; $stack -le $stacks; $stack++) {
            $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks
            $y = $radius * [Math]::Sin($phi); $r = $radius * [Math]::Cos($phi)
            for ($slice = 0; $slice -le $slices; $slice++) {
                $theta = $slice * 2 * [Math]::PI / $slices
                $x = $r * [Math]::Cos($theta); $z = $r * [Math]::Sin($theta)
                $allVertices.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
            }
        }

        for ($stack = 0; $stack -lt $stacks; $stack++) {
            for ($slice = 0; $slice -lt $slices; $slice++) {
                $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
                $i0 = $stack * ($slices + 1) + $slice; $i1 = ($stack + 1) * ($slices + 1) + $slice
                $i2 = $i0 + 1; $i3 = $i1 + 1
                $p0 = $allVertices[$i0]; $p1 = $allVertices[$i1]; $p2 = $allVertices[$i2]; $p3 = $allVertices[$i3]

                $mesh.Positions.Add($p0); $mesh.Positions.Add($p1); $mesh.Positions.Add($p2)
                $mesh.Positions.Add($p2); $mesh.Positions.Add($p1); $mesh.Positions.Add($p3)

                $uv0 = [System.Windows.Point]::new(0,0); $uv1 = [System.Windows.Point]::new(0,1)
                $uv2 = [System.Windows.Point]::new(1,0); $uv3 = [System.Windows.Point]::new(1,1)
                $mesh.TextureCoordinates.Add($uv0); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv2)
                $mesh.TextureCoordinates.Add($uv2); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv3)

                $facetModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $null)
                $facets.Add($facetModel)
            }
        }
        return $facets
    }

    function Setup-FacetedSphereMulti {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Faceted Sphere (Multi-Media)"
        $Viewport.Camera.Position = "0,0,8"
        $sphereContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($sphereContainer) | Out-Null

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        # 1. Create a small, fixed number of shared players and materials.
        $numberOfSharedPlayers = 8
        $sharedMaterials = [System.Collections.ArrayList]::new()

        for ($i = 0; $i -lt $numberOfSharedPlayers; $i++) {
            $playerKey = "SharedFacetPlayer$i"
            
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            
            $playerState = [hashtable]@{
                ContentPresenter     = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
                PlaybackStopwatch    = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
                MediaEndedHandler    = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                MediaOpenedHandler   = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                MediaFailedHandler   = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $sharedMaterials.Add($material) | Out-Null
            
            Start-NextMedia -PlayerKey $playerKey
        }

        # 2. Get the facet models.
        $facetModels = New-FacetedSphereModels -radius 2.5 -slices 8 -stacks 4

        # 3. Assign the shared materials to the facets.
        for ($i = 0; $i -lt $facetModels.Count; $i++) {
            $facetModel = $facetModels[$i]
            $facetModel.Material = $sharedMaterials | Get-Random
            $sphereContainer.Content.Children.Add($facetModel) | Out-Null
        }
        
        # 4. Set up the main rotation animations (this part is unchanged).
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_Faceted", $sphereContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_Faceted", $sphereContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Faceted"); $axisAngleY = $Window.FindName("AxisAngleY_Faceted")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-FloatingStars {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Stars"
        $Viewport.Camera.Position = "0,0,15"

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()

        $sphereRadius = 0.5; $coneHeight = 0.75; $coneRadius = 0.2
        $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 64 -stacks 32
        $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 64

        $visuals = @{
            "Middle" = @{ "Mesh" = $sphereMesh; "Transform" = (New-Object System.Windows.Media.Media3D.TranslateTransform3D) }
            "Top"    = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)) }
            "Bottom" = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Right"  = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Left"   = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Front"  = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Back"   = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
        }
        $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1))); $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
        $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), -90)))); $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius, 0, 0)))
        $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), 90)))); $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius, 0, 0)))
        $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 90)))); $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, $sphereRadius)))
        $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), -90)))); $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$sphereRadius)))

        # Loop 6 times to create 6 stars, each with its own single media player.
        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Star$i" # Simple player key for the whole star

            # 1. Create ONE media host and player state for the entire star.
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{
                ContentPresenter     = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
                PlaybackStopwatch    = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
                MediaEndedHandler    = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                MediaOpenedHandler   = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                MediaFailedHandler   = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            }
            $SyncHash.PlayerStates[$playerKey] = $playerState

            # 2. Create ONE shared material for the entire star.
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $sharedMaterial = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $sharedMaterial.Color = [System.Windows.Media.Colors]::White }

            # 3. Create the parent container for the star's 3D models.
            $starContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
            $Viewport.Children.Add($starContainer) | Out-Null

            # 4. Loop through the 7 parts, create their models, and assign the SHARED material.
            foreach ($partKey in $visuals.Keys) {
                $part = $visuals[$partKey]
                
                $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D
                $geometryModel.Geometry = $part.Mesh
                $geometryModel.Transform = $part.Transform
                $geometryModel.Material = $sharedMaterial # Use the same material for all parts
                
                $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $geometryModel }
                $starContainer.Children.Add($modelVisual) | Out-Null
            }

            # 5. Set up the animation for the star container.
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D))
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform)
            $starContainer.Transform = $transformGroup
            
            $starObject = [pscustomobject]@{
                Visual = $starContainer
                Translate = $translateTransform
                Rotate = $rotateTransform
                Velocity = {
                    $randomVector = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
                    if ($randomVector.Length -gt 0) { $randomVector.Normalize() }
                    return $randomVector * (Get-Random -Minimum 1.5 -Maximum 3.0)
                }.Invoke()
                RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0))
                CurrentRotation = New-Object System.Windows.Media.Media3D.Quaternion(0,0,0,1)
            }
            $SyncHash.FloatingObjects.Add($starObject) | Out-Null

            # 6. Start the single media player for this star.
            Start-NextMedia -PlayerKey $playerKey
        }
    }

    # --- Scrolling-Specific Media Handlers ---
    function Setup-RotatingCube {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Cube"
        $Viewport.Camera.Position = "0,0,5"
        $Viewport.Camera.FieldOfView = "70"
        $cubeXAML = @"
        <ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
            <ModelVisual3D.Transform>
                <Transform3DGroup>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D>
                </Transform3DGroup>
            </ModelVisual3D.Transform>
        </ModelVisual3D>
"@
        $cubeContainer = [Windows.Markup.XamlReader]::Parse($cubeXAML)
        $Viewport.Children.Add($cubeContainer)

        $cubeModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
        $cubeContainer.Content = $cubeModelGroup

        $faceNames = @("Front", "Back", "Right", "Left", "Top", "Bottom")
        $meshes = @{
            Front  = "-1,-1,1  1,-1,1  1,1,1  -1,1,1"
            Back   = "-1,-1,-1  -1,1,-1  1,1,-1  1,-1,-1"
            Right  = "1,-1,1  1,-1,-1  1,1,-1  1,1,1"
            Left   = "-1,-1,-1  -1,-1,1  -1,1,1  -1,1,-1"
            Top    = "-1,1,1  1,1,1  1,1,-1  -1,1,-1"
            Bottom = "-1,-1,-1  1,-1,-1  1,-1,1  -1,-1,1"
        }
        $texCoords = @{
            Front  = "0,1 1,1 1,0 0,0"
            Back   = "1,1 1,0 0,0 0,1"
            Right  = "0,1 1,1 1,0 0,0"
            Left   = "0,1 1,1 1,0 0,0"
            Top    = "0,1 1,1 1,0 0,0"
            Bottom = "0,1 1,1 1,0 0,0"
        }

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        foreach ($faceName in $faceNames) {
            $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
            $mesh.Positions = $meshes[$faceName]
            $mesh.TextureCoordinates = $texCoords[$faceName]
            $mesh.TriangleIndices = "0,1,2 0,2,3"

            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            # Explicitly create a mutable hashtable to ensure properties can be added later.
            $playerState = [hashtable]@{
                ContentPresenter  = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            # Define and store the handlers on the state object
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $faceName }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $faceName -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $faceName -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$faceName] = $playerState

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $material)
            $faceModel.Transform = New-Object System.Windows.Media.Media3D.ScaleTransform3D(0.7, 0.7, 0.7)
            $cubeModelGroup.Children.Add($faceModel)

            Start-NextMedia -PlayerKey $faceName
        }

        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        
        $Window.RegisterName("AxisAngleX_Cube", $cubeContainer.Transform.Children[0].Rotation)
        $Window.RegisterName("AxisAngleY_Cube", $cubeContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Cube"); $axisAngleY = $Window.FindName("AxisAngleY_Cube")

        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Sphere {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Sphere"
        $Viewport.Camera.Position = "0,0,6"
        $sphereMesh = New-SphereMesh -radius 1.6 -slices 128 -stacks 64
        
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
        $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null

        $playerState = [hashtable]@{
            ContentPresenter  = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
        }
        # Define and store the handlers on the state object
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey "Sphere" }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey "Sphere" -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates["Sphere"] = $playerState

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }
        if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material)
        
        $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }
        $Viewport.Children.Add($modelVisual)

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup

        # Create the AxisAngleRotation3D objects first.
        $axisAngleX = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "1,0,0" }
        $axisAngleY = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "0,1,0" }

        # Then create the RotateTransform3D and assign the rotation objects to them.
        $rotateX = New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleX)
        $rotateY = New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleY)

        $transformGroup.Children.Add($rotateX); $transformGroup.Children.Add($rotateY)
        $modelVisual.Transform = $transformGroup

        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }

        Start-NextMedia -PlayerKey "Sphere"
    }

    function Setup-PulsingStar {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Pulsing Star"
        $Viewport.Camera.Position = "0,0,12"
        $starContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><Transform3DGroup><ScaleTransform3D x:Name="PulseScale" ScaleX="1" ScaleY="1" ScaleZ="1" /><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="StarRotation" Axis="1,1,0.5" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($starContainer)

        $sphereRadius = 1.2; $coneHeight = 2.4; $coneRadius = 0.8
        $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64
        $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        function New-MediaHost {
            $grid = New-Object System.Windows.Controls.Grid
            $cp = New-Object System.Windows.Controls.ContentPresenter; $grid.Children.Add($cp) | Out-Null
            $tb = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $grid.Children.Add($tb) | Out-Null
            return @{ Grid = $grid; ContentPresenter = $cp; OverlayTextBlock = $tb }
        }

        # Central Sphere
        $sphereHost = New-MediaHost
        
        # Create the object first, then set its properties. This avoids the "property is read-only" error.
        $sphereVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $sphereVisual.Geometry = $sphereMesh
        $sphereVisual.Visual = $sphereHost.Grid

        $sphereMaterial = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($sphereMaterial, $true)
        $sphereVisual.Material = $sphereMaterial
        [void]$starContainer.Children.Add($sphereVisual)
        
        $playerState = [hashtable]@{
            ContentPresenter  = $sphereHost.ContentPresenter; OverlayTextBlock = $sphereHost.OverlayTextBlock; MediaHostGrid = $sphereHost.Grid;
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
        }
        # Define and store the handlers on the state object
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey "Sphere" }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey "Sphere" -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates["Sphere"] = $playerState

        # Cones
        $conePositions = @(
            @{ Name="Top";    Transform=(New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)) },
            @{ Name="Bottom"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Right";  Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
            @{ Name="Left";   Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Front";  Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
            @{ Name="Back";   Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }
        )
        $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1))); $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
        $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1',-90))))); $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius,0,0)))
        $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1',90))))); $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius,0,0)))
        $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0',90))))); $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0,0,$sphereRadius)))
        $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0',-90))))); $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0,0,-$sphereRadius)))

        foreach ($coneInfo in $conePositions) {
            $coneHost = New-MediaHost
            
            # Create the object first, then set its properties.
            $coneVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
            $coneVisual.Geometry = $coneMesh
            $coneVisual.Visual = $coneHost.Grid
            $coneVisual.Transform = $coneInfo.Transform
            
            $coneMaterial = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($coneMaterial, $true); $coneVisual.Material = $coneMaterial
            [void]$starContainer.Children.Add($coneVisual)
            
            $playerState = [hashtable]@{
                ContentPresenter  = $coneHost.ContentPresenter; OverlayTextBlock = $coneHost.OverlayTextBlock; MediaHostGrid = $coneHost.Grid;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            $playerKey = "Cone$($coneInfo.Name)"
            # Define and store the handlers on the state object
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState
        }

        $starAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }
        $pulseAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0.85, 1.15, [TimeSpan]::FromSeconds(2)) -Property @{ AutoReverse=$true; RepeatBehavior="Forever" }
        
        $Window.RegisterName("StarRotation_Pulsing", $starContainer.Transform.Children[1].Rotation)
        $Window.RegisterName("PulseScale_Pulsing", $starContainer.Transform.Children[0])
        $starRotation = $Window.FindName("StarRotation_Pulsing"); $pulseScale = $Window.FindName("PulseScale_Pulsing")

        $starRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $starAnim)
        $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $pulseAnim)
        $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $pulseAnim)
        $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $pulseAnim)
        $SyncHash.Animations = @{ Star = $starAnim; Pulse = $pulseAnim }; $SyncHash.Rotations = @{ Star = $starRotation }; $SyncHash.Transforms = @{ Pulse = $pulseScale }
        
        $SyncHash.PlayerStates.Keys | ForEach-Object { Start-NextMedia -PlayerKey $_ }
    }

    function Setup-FacetedSphereSingle {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Faceted Sphere (Single Media)"
        $Viewport.Camera.Position = "0,0,8"

        $sphereContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($sphereContainer) | Out-Null

        # Create a new, empty mesh to hold the combined geometry of all facets.
        $sphereMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

        # Get all the individual facet models.
        $facetModels = New-FacetedSphereModels -radius 2.5 -slices 10 -stacks 5
        foreach ($facetModel in $facetModels) {
            $facetGeom = $facetModel.Geometry
            $facetGeom.Positions | ForEach-Object { $sphereMesh.Positions.Add($_) }
            $facetGeom.TextureCoordinates | ForEach-Object { $sphereMesh.TextureCoordinates.Add($_) }
        }
        
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
        $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null

        $playerState = [hashtable]@{
            ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
        }
        # Define and store the handlers on the state object
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey "Sphere" }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey "Sphere" -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates["Sphere"] = $playerState

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material)
        $sphereContainer.Content.Children.Add($model) | Out-Null

        Start-NextMedia -PlayerKey "Sphere"

        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_FacetedSingle", $sphereContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_FacetedSingle", $sphereContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_FacetedSingle"); $axisAngleY = $Window.FindName("AxisAngleY_FacetedSingle")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-FloatingSpheres {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Spheres"
        $Viewport.Camera.Position = "0,0,15"

        $sphereMesh = New-SphereMesh -radius 1.5 -slices 128 -stacks 64
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()

        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Sphere$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{
                ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            # Define and store the handlers on the state object
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material)
            $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }
            $Viewport.Children.Add($modelVisual) | Out-Null

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D))
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform)
            $modelVisual.Transform = $transformGroup

            $starObject = [pscustomobject]@{
                Visual = $modelVisual
                Translate = $translateTransform
                Rotate = $rotateTransform
                Velocity = {
                    $randomVector = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
                    if ($randomVector.Length -gt 0) { $randomVector.Normalize() }
                    return $randomVector * (Get-Random -Minimum 1.5 -Maximum 3.0)
                }.Invoke()
                RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0))
                CurrentRotation = New-Object System.Windows.Media.Media3D.Quaternion(0,0,0,1)
            }
            $SyncHash.FloatingObjects.Add($starObject) | Out-Null

            Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-MediaFlowFunnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Flow Funnel"
        $Viewport.Camera.Position = "0,2,18"
        $Viewport.Camera.LookDirection = "0,-0.1,-1"

        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0"/></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0"/></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer) | Out-Null
        $funnelModelGroup = $funnelContainer.Content

        $numberOfGroups = 6
        $sharedMaterials = @()
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $numberOfGroups; $i++) {
            $playerKey = "Group$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush }

            $SyncHash.PlayerStates[$playerKey] = @{
                ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            $playerState = $SyncHash.PlayerStates[$playerKey]
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            Start-NextMedia -PlayerKey $playerKey
        }

        $SyncHash.PanelModels = @{}
        $numberOfRings = 5; $panelsPerRing = 12; $startRadius = 7.0; $endRadius = 1.0; $totalHeight = 8.0
        $radiusStep = ($startRadius - $endRadius) / $numberOfRings; $angleStep = (2 * [Math]::PI) / $panelsPerRing
        
        for ($r = 0; $r -lt $numberOfRings; $r++) {
            $outerR = $startRadius - ($r * $radiusStep); $innerR = $startRadius - (($r + 1) * $radiusStep)
            $progressOuter = $r / $numberOfRings; $progressOuter_eased = $progressOuter * $progressOuter
            $progressInner = ($r + 1) / $numberOfRings; $progressInner_eased = $progressInner * $progressInner
            $yOuter = $totalHeight / 2 - ($progressOuter_eased * $totalHeight); $yInner = $totalHeight / 2 - ($progressInner_eased * $totalHeight)
            for ($p = 0; $p -lt $panelsPerRing; $p++) {
                $theta1 = $p * $angleStep; $theta2 = ($p + 1) * $angleStep
                $p1 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta1), $yOuter, $outerR * [Math]::Sin($theta1))
                $p2 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta2), $yOuter, $outerR * [Math]::Sin($theta2))
                $p3 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta2), $yInner, $innerR * [Math]::Sin($theta2))
                $p4 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta1), $yInner, $innerR * [Math]::Sin($theta1))
                $panelMesh = New-SpiralingPanelMesh -p1 $p1 -p2 $p2 -p3 $p3 -p4 $p4
                
                $materialToUse = $sharedMaterials | Get-Random
                $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $materialToUse; BackMaterial = $materialToUse.Clone() }
                $funnelModelGroup.Children.Add($geometryModel) | Out-Null
                $SyncHash.PanelModels[($r * $panelsPerRing) + $p] = $geometryModel
            }
        }

        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(45)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleY = $funnelContainer.Transform.Children[0].Rotation
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ Y = $animY }; $SyncHash.Rotations = @{ Y = $axisAngleY }
        $SyncHash.SharedMaterials = $sharedMaterials; $SyncHash.PanelsPerRing = $panelsPerRing
    }

    function Setup-CurvedVortex {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Curved Vortex"
        $Viewport.Camera.Position = "0,5,15"
        $Viewport.Camera.LookDirection = "0,-0.3,-1"

        $vortexContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $Viewport.Children.Add($vortexContainer) | Out-Null

        $panelCount = 16
        $curvedPanelMesh = New-CurvedPanelMesh -width 2.5 -height 1.5 -curveDepth 0.5
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $panelCount; $i++) {
            $playerKey = "Panel$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush; Color = [System.Windows.Media.Colors]::White }
            
            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $curvedPanelMesh; Material = $material; BackMaterial = $material.Clone() }
            $panelContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $panelModel }

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
            $transformGroup.Children.Add($rotateTransform) | Out-Null
            $transformGroup.Children.Add($translateTransform) | Out-Null
            $panelContainer.Transform = $transformGroup
            $vortexContainer.Children.Add($panelContainer) | Out-Null

            $SyncHash.PlayerStates[$playerKey] = @{
                ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
                TranslateTransform = $translateTransform; RotateTransform = $rotateTransform;
                CurrentAngle = (720.0 / $panelCount) * $i;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            $playerState = $SyncHash.PlayerStates[$playerKey]
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-RotatingStar {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Star"
        $Viewport.Camera.Position = "0,0,12"

        $starContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($starContainer) | Out-Null

        $totalObjectHeight = 12 * 0.75
        $sphereRadius = ($totalObjectHeight / 5.0) * 0.5
        $coneHeight = ($sphereRadius * 1.5) * 2.0
        $coneRadius = ($sphereRadius * 0.4) * 2.0

        $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64
        $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        $targets = @("Top", "Middle", "Bottom", "Left", "Right", "Front", "Back")
        foreach ($target in $targets) {
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
            $SyncHash.PlayerStates[$target] = $playerState
            # Define and store the handlers on the state object
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $target }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $target -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $target -Reason $eventArgs.ErrorException.Message }.GetNewClosure()


            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $mesh = if ($target -eq "Middle") { $sphereMesh } else { $coneMesh }
            $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $material)

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            switch ($target) {
                "Top"    { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))) }
                "Bottom" { 
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1)))
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
                }
                "Right"  {
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', -90)))))
                }
                "Left"   {
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', 90)))))
                }
                "Front"  {
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', -90)))))
                }
                "Back"   {
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
                    $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', 90)))))
                }
            }
            $model.Transform = $transformGroup

            $starContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model })) | Out-Null
            Start-NextMedia -PlayerKey $target
        }

        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        
        $Window.RegisterName("AxisAngleX_RotStar", $starContainer.Transform.Children[0].Rotation)
        $Window.RegisterName("AxisAngleY_RotStar", $starContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_RotStar"); $axisAngleY = $Window.FindName("AxisAngleY_RotStar")

        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-ButterflyEffect {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Butterfly Effect"
        $Viewport.Camera.Position = "0,0,15"

        $planeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $planeMesh.Positions = "-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0"
        $planeMesh.TriangleIndices = "0,1,2 0,2,3"
        $planeMesh.TextureCoordinates = "0,1 1,1 1,0 0,0"
        $planeMesh.Freeze()

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()
        for ($i = 1; $i -le 6; $i++) {
            $butterflyContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
            
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D))
            $flutterTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "1,0,0" }))

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $transformGroup.Children.Add($flutterTransform) | Out-Null
            $transformGroup.Children.Add($rotateTransform) | Out-Null
            $transformGroup.Children.Add($translateTransform) | Out-Null
            $butterflyContainer.Transform = $transformGroup

            foreach ($face in @("Front", "Back")) {
                $playerKey = "${i}_${face}"

                $mediaHostGrid = New-Object System.Windows.Controls.Grid
                $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
                $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
                $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null

                $playerState = [hashtable]@{
                    ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
                }
                # Define and store the handlers on the state object
                $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
                $SyncHash.PlayerStates[$playerKey] = $playerState

                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
                $material = New-Object $materialType -Property @{ Brush = $visualBrush }
                if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

                $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($planeMesh, $material)
                if ($face -eq "Back") {
                    $faceModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D("0,1,0", 180)))
                }
                $butterflyContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $faceModel })) | Out-Null
                Start-NextMedia -PlayerKey $playerKey
            }

            $Viewport.Children.Add($butterflyContainer) | Out-Null

            $startX = (Get-Random -Minimum -8 -Maximum 8)
            $startY = (Get-Random -Minimum -4 -Maximum 4)
            $translateTransform.OffsetX = $startX
            $translateTransform.OffsetY = $startY

            $butterflyObject = [pscustomobject]@{
                Visual = $butterflyContainer
                Translate = $translateTransform
                Rotate = $rotateTransform
                FlutterTransform = $flutterTransform
                Velocity = New-Object System.Windows.Media.Media3D.Vector3D(((Get-Random -Minimum 0.5 -Maximum 1.5) * (Get-Random @(1, -1))), ((Get-Random -Minimum 0.5 -Maximum 1.5) * (Get-Random @(1, -1))), 0)
                RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20 -Maximum 20), (Get-Random -Minimum -20 -Maximum 20), 0)
                CurrentRotation = [System.Windows.Media.Media3D.Quaternion]::new(0,0,0,1)
            }
            $SyncHash.FloatingObjects.Add($butterflyObject) | Out-Null
        }
    }

    function Setup-WagonWheel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Wagon Wheel"
        $Viewport.Camera.Position = "0,0,8"
        $container = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($container)

        $camera = $Viewport.Camera
        $visibleHeight = 2.0 * $camera.Position.Z * [Math]::Tan(($camera.FieldOfView * ([Math]::PI / 180.0)) / 2.0)
        $dynamicRadius = ($visibleHeight * 0.50) / 2.0
        
        $numberOfSlices = 8; $sliceAngle = 360.0 / $numberOfSlices
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $numberOfSlices; $i++) {
            $sliceParts = New-WagonWheelSliceModel -radius $dynamicRadius -startAngleDeg ($i * $sliceAngle) -sliceAngleDeg $sliceAngle
            
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $mediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; Margin='10,0,10,0'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $visualBrush = New-Object System.Windows.Media.VisualBrush($mediaHostGrid)
            $material = New-Object $materialType($visualBrush)
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $sliceParts.OuterFaceModel.Material = $material
            $container.Content.Children.Add($sliceParts.FullSliceModel) | Out-Null

            $playerState = [hashtable]@{
                ContentPresenter  = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid;
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
            }
            $playerKey = "Slice$i"
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState
            Start-NextMedia -PlayerKey $playerKey
        }

        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        
        $Window.RegisterName("AxisAngleX_Wagon", $container.Transform.Children[0].Rotation)
        $Window.RegisterName("AxisAngleY_Wagon", $container.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Wagon"); $axisAngleY = $Window.FindName("AxisAngleY_Wagon")

        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Carousel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "3D Media Carousel"
        $Viewport.Camera.Position = "0,0,12"

        $carouselContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="CarouselRotation" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($carouselContainer) | Out-Null

        $panelCount = 8; $angleIncrement = 360 / $panelCount; $panelWidth = 3.0; $panelHeight = 5.0; $carouselRadius = 4.0
        $panelMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $panelMesh.Positions = "$(-$panelWidth/2),$(-$panelHeight/2),0  $($panelWidth/2),$(-$panelHeight/2),0  $($panelWidth/2),$($panelHeight/2),0  $(-$panelWidth/2),$($panelHeight/2),0"
        $panelMesh.TriangleIndices = "0,1,2 0,2,3"; $panelMesh.TextureCoordinates = "0,1 1,1 1,0 0,0"

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $panelCount; $i++) {
            $panelAngle = $i * $angleIncrement
            $panelContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D

            foreach ($face in @("Front", "Back")) {
                $playerKey = "Panel${i}_${face}"
                $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
                $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
                $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
                $SyncHash.PlayerStates[$playerKey] = $playerState

                # Create the Viewport2DVisual3D object first, then set its properties.
                # This is more robust than using the -Property hashtable for this specific WPF object.
                $panelViewport = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
                $panelViewport.Geometry = $panelMesh.Clone() # A MeshGeometry3D can only have one parent, so we must clone it.
                $panelViewport.Visual = $mediaHostGrid

                if ($face -eq "Back") { $panelViewport.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180))) }
                $material = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($material, $true); $panelViewport.Material = $material
                $panelContainer.Children.Add($panelViewport) | Out-Null
                Start-NextMedia -PlayerKey $playerKey
            }

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$carouselRadius))) | Out-Null
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', $panelAngle))))) | Out-Null
            
            # Add the vertical translation for the undulating motion
            $verticalTranslate = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($verticalTranslate) | Out-Null

            $panelContainer.Transform = $transformGroup; $carouselContainer.Children.Add($panelContainer) | Out-Null

            # Vertical oscillation animation
            $verticalAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
                From = -0.5; To = 0.5; Duration = [TimeSpan]::FromSeconds(4); AutoReverse = $true; RepeatBehavior = "Forever"; BeginTime = [TimeSpan]::FromSeconds($i * 0.5)
            }
            $verticalTranslate.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $verticalAnim)
        }

        $carouselAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(60)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("CarouselRotation_Main", $carouselContainer.Transform.Rotation); $carouselRotation = $Window.FindName("CarouselRotation_Main")
        $carouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $carouselAnim)
        $SyncHash.Animations = @{ Carousel = $carouselAnim }; $SyncHash.Rotations = @{ Carousel = $carouselRotation }
    }

    function Setup-ConcentricFunnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Concentric Funnel"
        $Viewport.Camera.Position = "0,2,18"
        $Viewport.Camera.LookDirection = "0,-0.1,-1"

        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0"/></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0"/></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer) | Out-Null
        $funnelModelGroup = $funnelContainer.Content

        $numberOfGroups = 6
        $sharedMaterials = @()
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $numberOfGroups; $i++) {
            $playerKey = "Group$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush }

            $SyncHash.PlayerStates[$playerKey] = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
            $playerState = $SyncHash.PlayerStates[$playerKey]
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            Start-NextMedia -PlayerKey $playerKey
        }

        $numberOfRings = 5; $panelsPerRing = 12; $startRadius = 7.0; $endRadius = 1.0; $totalHeight = 8.0
        $radiusStep = ($startRadius - $endRadius) / $numberOfRings; $angleStep = (2 * [Math]::PI) / $panelsPerRing
        
        for ($r = 0; $r -lt $numberOfRings; $r++) {
            $outerR = $startRadius - ($r * $radiusStep); $innerR = $startRadius - (($r + 1) * $radiusStep)
            $progressOuter = $r / $numberOfRings; $progressOuter_eased = $progressOuter * $progressOuter
            $progressInner = ($r + 1) / $numberOfRings; $progressInner_eased = $progressInner * $progressInner
            $yOuter = $totalHeight / 2 - ($progressOuter_eased * $totalHeight); $yInner = $totalHeight / 2 - ($progressInner_eased * $totalHeight)
            for ($p = 0; $p -lt $panelsPerRing; $p++) {
                $theta1 = $p * $angleStep; $theta2 = ($p + 1) * $angleStep
                $p1 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta1), $yOuter, $outerR * [Math]::Sin($theta1))
                $p2 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta2), $yOuter, $outerR * [Math]::Sin($theta2))
                $p3 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta2), $yInner, $innerR * [Math]::Sin($theta2))
                $p4 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta1), $yInner, $innerR * [Math]::Sin($theta1))
                $panelMesh = New-SpiralingPanelMesh -p1 $p1 -p2 $p2 -p3 $p3 -p4 $p4
                $materialToUse = $sharedMaterials | Get-Random
                $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $materialToUse; BackMaterial = $materialToUse.Clone() }
                $funnelModelGroup.Children.Add($geometryModel) | Out-Null
            }
        }
    }

    function Setup-FloatingCubes {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Cubes"
        $Viewport.Camera.Position = "0,0,15"

        $cubeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $cubeMesh.Positions = "-0.5,-0.5,0.5 0.5,-0.5,0.5 0.5,0.5,0.5 -0.5,0.5,0.5 -0.5,-0.5,-0.5 0.5,-0.5,-0.5 0.5,0.5,-0.5 -0.5,0.5,-0.5"
        $cubeMesh.TriangleIndices = "0,1,2 0,2,3 4,7,6 4,6,5 0,4,5 0,5,1 1,5,6 1,6,2 2,6,7 2,7,3 3,7,4 3,4,0"
        $cubeMesh.TextureCoordinates = "0,1 1,1 1,0 0,0 1,1 0,1 0,0 1,0 0,1 1,1 1,0 0,0 0,1 1,1 1,0 0,0"

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        # This style uses WPF animations, not a manual animation loop.
        $SyncHash.Animations = [hashtable]::Synchronized(@{})

        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Cube$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($cubeMesh, $material)
            $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }
            $Viewport.Children.Add($modelVisual)

            # --- Animation Setup (Adapted from original FloatingCubes script) ---
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            
            # Add a scale transform to make the cubes larger.
            $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D(2, 2, 2)
            $transformGroup.Children.Add($scaleTransform)

            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D))
            $rotateTransform.Rotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform)
            $modelVisual.Transform = $transformGroup

            $rotAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds((Get-Random -Minimum 15 -Maximum 45))) -Property @{ RepeatBehavior = 'Forever' }
            $rotateTransform.Rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $rotAnim)

            $posAnimX = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }
            $posAnimY = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }
            $posAnimZ = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }

            $durationSeconds = (Get-Random -Minimum 20 -Maximum 60)
            $posAnimX.Duration = [TimeSpan]::FromSeconds($durationSeconds); $posAnimY.Duration = [TimeSpan]::FromSeconds($durationSeconds); $posAnimZ.Duration = [TimeSpan]::FromSeconds($durationSeconds)

            $xRadius = (Get-Random -Minimum 3 -Maximum 8); $yRadius = (Get-Random -Minimum 2 -Maximum 6); $zRadius = (Get-Random -Minimum 1 -Maximum 4)
            $timeOffset = ($durationSeconds / 4) * (Get-Random -Minimum 0 -Maximum 3)

            for ($k = 0; $k -le 4; $k++) {
                $time = [System.Windows.Media.Animation.KeyTime]::FromTimeSpan([TimeSpan]::FromSeconds(($k * $durationSeconds / 4 + $timeOffset) % $durationSeconds))
                $angle = $k * [Math]::PI / 2
                $x = $xRadius * [Math]::Cos($angle); $y = $yRadius * [Math]::Sin($angle); $z = $zRadius * [Math]::Cos($angle * 2)
                $posAnimX.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($x, $time))) | Out-Null
                $posAnimY.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($y, $time))) | Out-Null
                $posAnimZ.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($z, $time))) | Out-Null
            }

            $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $posAnimX)
            $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $posAnimY)
            $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $posAnimZ)

            $SyncHash.Animations[$modelVisual.GetHashCode()] = [pscustomobject]@{
                Rotation = $rotAnim; PositionX = $posAnimX; PositionY = $posAnimY; PositionZ = $posAnimZ
                TranslateTransform = $translateTransform; RotateTransform = $rotateTransform
            }

            Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-Funnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Funnel"
        $Viewport.Camera.Position = "0,0,15"
        $Viewport.Camera.FieldOfView = "70"

        $funnelContainerXAML = @"
        <ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
            <ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content>
            <ModelVisual3D.Transform>
                <Transform3DGroup>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0" /></RotateTransform3D.Rotation></RotateTransform3D>
                    <TranslateTransform3D OffsetY="-2.0" />
                </Transform3DGroup>
            </ModelVisual3D.Transform>
        </ModelVisual3D>
"@
        $funnelContainer = [Windows.Markup.XamlReader]::Parse($funnelContainerXAML)
        $Viewport.Children.Add($funnelContainer)
        $modelGroup = $funnelContainer.Content

        $panelCount = 8
        $panelMesh = New-SpiralingPanelMesh -startRadius 5.0 -endRadius 1.0 -height 10.0 -arcAngle (360.0 / $panelCount) -twistAngle 90
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $panelCount; $i++) {
            $playerKey = "Panel$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $material; BackMaterial = $material.Clone() }
            $angle = $i * (360.0 / $panelCount)
            $panelModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), $angle)))
            $modelGroup.Children.Add($panelModel) | Out-Null

            Start-NextMedia -PlayerKey $playerKey
        }

        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleY = $funnelContainer.Transform.Children[0].Rotation
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ Y = $animY }; $SyncHash.Rotations = @{ Y = $axisAngleY }
    }

    function Setup-FunnelSingle {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Funnel (Single)"
        $Viewport.Camera.Position = "0,0,15"
        $Viewport.Camera.FieldOfView = "70"

        $funnelContainerXAML = @"
        <ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
            <ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content>
            <ModelVisual3D.Transform>
                <Transform3DGroup>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                    <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0" /></RotateTransform3D.Rotation></RotateTransform3D>
                    <TranslateTransform3D OffsetY="-2.0" />
                </Transform3DGroup>
            </ModelVisual3D.Transform>
        </ModelVisual3D>
"@
        $funnelContainer = [Windows.Markup.XamlReader]::Parse($funnelContainerXAML)
        $Viewport.Children.Add($funnelContainer)
        $modelGroup = $funnelContainer.Content

        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
        $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null

        $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey "Funnel" }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey "Funnel" -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey "Funnel" -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates["Funnel"] = $playerState

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

        $panelCount = 8
        $panelMesh = New-SpiralingPanelMesh -startRadius 5.0 -endRadius 1.0 -height 10.0 -arcAngle (360.0 / $panelCount) -twistAngle 90
        for ($i = 0; $i -lt $panelCount; $i++) {
            $angle = $i * (360.0 / $panelCount)
            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $material; BackMaterial = $material.Clone() }
            $panelModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), $angle)))
            $modelGroup.Children.Add($panelModel) | Out-Null
        }
        Start-NextMedia -PlayerKey "Funnel"

        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleY = $funnelContainer.Transform.Children[0].Rotation
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ Y = $animY }
        $SyncHash.Rotations = @{ Y = $axisAngleY }
    }

    # --- Scrolling-Specific Media Handlers ---
    function Handle-ScrollingMediaFailure {
        param([string]$PlayerKey, [string]$Reason)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $playerState -or $playerState.IsFailed) { return }
        $playerState.IsFailed = $true

        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Scrolling media failed for player '$PlayerKey' in style '$($SyncHash.VisualizationStyle)' (File: '$fileName'). Reason: $Reason. Attempting to replace."

        # Add the bad file to the blacklist.
        if ($playerState.CurrentSource -and -not $SyncHash.BadMediaFiles.Contains($playerState.CurrentSource.LocalPath)) {
            $SyncHash.BadMediaFiles.Add($playerState.CurrentSource.LocalPath) | Out-Null
        }

        $SyncHash.Window.Dispatcher.Invoke([action]{
            if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
            $playerState.ContentPresenter.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
        })

        $recoveryScriptBlock = { Start-NextScrollingMedia -PlayerKey $PlayerKey }
        $null = $SyncHash.Window.Dispatcher.InvokeAsync($recoveryScriptBlock.GetNewClosure())
    }

    function Handle-ScrollingMediaEnded {
        param([string]$PlayerKey)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if ($pState.IsFailed) { return }

        # This is the key difference: For scrolling, an image timer completing is a SUCCESS, not a failure.
        # We only check for instant video failures.
        if (-not $pState.IsImage) {
            $pState.PlaybackStopwatch.Stop()
            if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500) {
                Handle-ScrollingMediaFailure -PlayerKey $PlayerKey -Reason "Playback ended instantly (bad codec)."
                return
            }
        }
        Start-NextScrollingMedia -PlayerKey $PlayerKey
    }

    function Handle-ScrollingMediaOpened {
        param([string]$PlayerKey, $EventArgs)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $pState) { return }
        $pState.IsFailed = $false
        $pState.PlaybackStopwatch.Restart()
        $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent

        if (-not $pState.IsImage) {
            if (-not $EventArgs.NaturalDuration.HasTimeSpan) {
                Handle-ScrollingMediaFailure -PlayerKey $PlayerKey -Reason "Invalid duration or codec."
                return
            }
            # Resize video grid based on aspect ratio
            if ($EventArgs.NaturalVideoHeight -gt 0) {
                $aspectRatio = $EventArgs.NaturalVideoWidth / $EventArgs.NaturalVideoHeight
                if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal") {
                    $pState.MediaHostGrid.Width = $pState.MediaHostGrid.Height * $aspectRatio
                } else { # ScrollingVertical
                    $pState.MediaHostGrid.Height = $pState.MediaHostGrid.Width / $aspectRatio
                }
            }
        }
    }

    function Start-NextScrollingMedia {
        param([string]$PlayerKey)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false
        if ($SyncHash.RbSelection -eq "Filename") { $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath) }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        $playerState.IsImage = $ImageExtensions -contains $extension

        try {
            if ($playerState.IsImage) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit(); $bitmap.Freeze()
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                
                # Resize image grid based on aspect ratio
                if ($bitmap.PixelHeight -gt 0) {
                    $aspectRatio = $bitmap.PixelWidth / $bitmap.PixelHeight
                    if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal") {
                        $playerState.MediaHostGrid.Width = $playerState.MediaHostGrid.Height * $aspectRatio
                    } else { # ScrollingVertical
                        $playerState.MediaHostGrid.Height = $playerState.MediaHostGrid.Width / $aspectRatio
                    }
                }
                $playerState.ContentPresenter.Content = $image
                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $PlayerKey }
                $timer.Add_Tick({ $t = $args[0]; $key = $t.Tag; $t.Stop(); Handle-ScrollingMediaEnded -PlayerKey $key })
                $playerState.MediaTimer = $timer; $timer.Start()
            } else { # Video
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource; Tag = $PlayerKey
                }
                # Retrieve and attach handlers from the player state
                $mediaElement.Add_MediaEnded($playerState.MediaEndedHandler)
                $mediaElement.Add_MediaOpened($playerState.MediaOpenedHandler)
                $mediaElement.Add_MediaFailed($playerState.MediaFailedHandler)

                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.CurrentMediaElement = $mediaElement
                $mediaElement.Play()
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            }
        } catch {
            Handle-ScrollingMediaFailure -PlayerKey $PlayerKey -Reason $_.Exception.Message
        }
    }

    function Setup-ScrollingHorizontal {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Scrolling Horizontal"
        $Viewport.Visibility = 'Collapsed'; $scrollingCanvas = $Window.FindName("scrollingCanvas"); $scrollingCanvas.Visibility = 'Visible'
        $SyncHash.scrollingPanel = $Window.FindName("scrollingPanel"); $SyncHash.scrollingPanel.Orientation = 'Horizontal'

        $numberOfPlayers = 8; $playerHeight = $Window.Height
        
        # This scriptblock will be executed AFTER the window is loaded and all initial media has had a chance to resize.
        $SyncHash.startAnimationScriptBlock = {
            param($sh)
            $win = $sh.Window
            $panel = $sh.scrollingPanel
            $totalWidth = 0
            foreach ($child in $panel.Children) { $totalWidth += $child.ActualWidth + $child.Margin.Left + $child.Margin.Right }
            if ($totalWidth -eq 0) { return }

            $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, -$totalWidth, [TimeSpan]::FromSeconds(90)) -Property @{ RepeatBehavior = "Forever" }
            $storyboard = New-Object System.Windows.Media.Animation.Storyboard; $storyboard.Children.Add($anim)
            $transformGroup = New-Object System.Windows.Media.TransformGroup; $sh.translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0)
            $transformGroup.Children.Add($sh.translateTransform) | Out-Null; $panel.RenderTransform = $transformGroup

            [System.Windows.Media.Animation.Storyboard]::SetTarget($anim, $panel)
            [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($anim, (New-Object System.Windows.PropertyPath("(UIElement.RenderTransform).(TransformGroup.Children)[0].(TranslateTransform.X)")))
            $storyboard.Begin($panel, $true)
            $sh.Storyboard = $storyboard; $sh.Animations = @{ Scroll = $anim }; $sh.Transforms = @{ Scroll = $sh.translateTransform }
        }

        for ($i = 0; $i -lt $numberOfPlayers; $i++) {
            $playerKey = "Scroller$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Height = $playerHeight; Width = $playerHeight; Margin = '5' }
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $SyncHash.scrollingPanel.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch }
            # Define and store handlers
            $playerState.MediaEndedHandler = { Handle-ScrollingMediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-ScrollingMediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-ScrollingMediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState
            Start-NextScrollingMedia -PlayerKey $playerKey
        }
        # Use the Dispatcher to invoke the animation start after the initial layout pass is complete.
        $Window.Add_Loaded({ $Window.Dispatcher.InvokeAsync({ & $SyncHash.startAnimationScriptBlock $SyncHash }, "Loaded") | Out-Null })
    }

    function Setup-ScrollingVertical {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Scrolling Vertical"
        $Viewport.Visibility = 'Collapsed'; $scrollingCanvas = $Window.FindName("scrollingCanvas"); $scrollingCanvas.Visibility = 'Visible'
        $SyncHash.scrollingPanel = $Window.FindName("scrollingPanel"); $SyncHash.scrollingPanel.Orientation = 'Vertical'

        $SyncHash.startAnimationScriptBlock = {
            param($sh)
            $win = $sh.Window
            $panel = $sh.scrollingPanel
            $totalHeight = 0
            foreach ($child in $panel.Children) { $totalHeight += $child.ActualHeight + $child.Margin.Top + $child.Margin.Bottom }
            if ($totalHeight -eq 0) { return }

            $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, -$totalHeight, [TimeSpan]::FromSeconds(90)) -Property @{ RepeatBehavior = "Forever" }
            $storyboard = New-Object System.Windows.Media.Animation.Storyboard; $storyboard.Children.Add($anim)
            $transformGroup = New-Object System.Windows.Media.TransformGroup; $sh.translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0)
            $transformGroup.Children.Add($sh.translateTransform) | Out-Null; $panel.RenderTransform = $transformGroup

            [System.Windows.Media.Animation.Storyboard]::SetTarget($anim, $panel)
            [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($anim, (New-Object System.Windows.PropertyPath("(UIElement.RenderTransform).(TransformGroup.Children)[0].(TranslateTransform.Y)")))
            $storyboard.Begin($panel, $true)
            $sh.Storyboard = $storyboard; $sh.Animations = @{ Scroll = $anim }; $sh.Transforms = @{ Scroll = $sh.translateTransform }
        }

        $numberOfPlayers = 6; $playerWidth = $Window.Width

        for ($i = 0; $i -lt $numberOfPlayers; $i++) {
            $playerKey = "Scroller$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Width = $playerWidth; Height = ($playerWidth * 0.5625); Margin = '5' } # Default to 16:9 aspect ratio
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $SyncHash.scrollingPanel.Children.Add($mediaHostGrid) | Out-Null

            $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch }
            # Define and store handlers
            $playerState.MediaEndedHandler = { Handle-ScrollingMediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-ScrollingMediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-ScrollingMediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            $SyncHash.PlayerStates[$playerKey] = $playerState
            Start-NextScrollingMedia -PlayerKey $playerKey
        }    
        $Window.Add_Loaded({ $Window.Dispatcher.InvokeAsync({ & $SyncHash.startAnimationScriptBlock $SyncHash }, "Loaded") | Out-Null })
    }

    function Setup-Aquarium {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Aquarium"
        $Viewport.Camera.Position = "0,0,10"

        # Create and store the two fish mesh geometries
        $rightFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="0.6,0,0 0.5,0.14,0 0.2,0.196,0 -0.1,0.175,0 -0.4,0.105,0 -0.6,0.14,0 -0.667,0.12,0 -0.6,0,0 -0.667,-0.12,0 -0.6,-0.14,0 -0.4,-0.105,0 -0.1,-0.175,0 0.2,-0.196,0 0.5,-0.14,0 0.667,0.05,0 0.667,-0.05,0" TriangleIndices="0,14,1 0,1,13 0,13,15 1,2,12 1,12,13 2,3,11 2,11,12 3,4,10 3,10,11 4,5,7 4,7,10 5,6,7 7,8,9 7,9,10" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
        $leftFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="-0.6,0,0 -0.5,0.14,0 -0.2,0.196,0 0.1,0.175,0 0.4,0.105,0 0.6,0.14,0 0.667,0.12,0 0.6,0,0 0.667,-0.12,0 0.6,-0.14,0 0.4,-0.105,0 0.1,-0.175,0 -0.2,-0.196,0 -0.5,-0.14,0 -0.667,0.05,0 -0.667,-0.05,0" TriangleIndices="0,1,14 0,13,1 0,15,13 1,12,2 1,13,12 2,11,3 2,12,11 3,10,4 3,11,10 4,7,5 4,10,7 5,7,6 7,9,8 7,10,9" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
        $SyncHash.RightFacingFish = [Windows.Markup.XamlReader]::Parse($rightFishXaml)
        $SyncHash.LeftFacingFish = [Windows.Markup.XamlReader]::Parse($leftFishXaml)

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()

        for ($i = 0; $i -lt 6; $i++) {
            $fishContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $fishContainer.Transform = $translateTransform
            $Viewport.Children.Add($fishContainer) | Out-Null

            foreach ($face in @("Front", "Back")) {
                $playerKey = "Fish${i}_${face}"
                $mediaHostGrid = New-Object System.Windows.Controls.Grid
                $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
                $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
                $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null

                $playerState = [hashtable]@{
                    ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
                }
                $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
                $SyncHash.PlayerStates[$playerKey] = $playerState

                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
                $material = New-Object $materialType -Property @{ Brush = $visualBrush }
                if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

                $fishModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Material = $material }
                if ($face -eq "Back") {
                    $fishModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180)))
                }
                $fishContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $fishModel })) | Out-Null

                Start-NextMedia -PlayerKey $playerKey
            }

            $velocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -0.75 -Maximum 0.75), (Get-Random -Minimum -0.25 -Maximum 0.25))
            if ($velocity.X -eq 0) { $velocity.X = 0.5 }

            $fishObject = [pscustomobject]@{
                Visual = $fishContainer
                Translate = $translateTransform
                Velocity = $velocity
                RightGeometry = $SyncHash.RightFacingFish.Clone()
                LeftGeometry = $SyncHash.LeftFacingFish.Clone()
            }
            [void]$SyncHash.FloatingObjects.Add($fishObject)
        }
    }

    function Setup-RollerCoaster {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Roller Coaster"
        $Viewport.Camera.Position = "0,5,25"
        $Viewport.Camera.LookDirection = "0,-0.2,-1"
        $Viewport.Camera.FieldOfView = "60"

        # Define the parametric path for the roller coaster
        $script:pathFunc = {
            param([double]$t) # t goes from 0 to 1
            $t_rad = $t * 2 * [Math]::PI
            $x_base = 12 * [Math]::Cos($t_rad); $z_base = 5 * [Math]::Sin($t_rad)
            $y_base = 2.5 * [Math]::Sin(3 * $t_rad) - 1.5 * [Math]::Cos(5 * $t_rad) + [Math]::Sin($t_rad)
            $loop_radius = 5.0; $loop_center_t = 0.5; $loop_width = 0.08
            $loop_influence = [Math]::Exp(-[Math]::Pow($t - $loop_center_t, 2) / (2 * [Math]::Pow($loop_width, 2)))
            $y = $y_base + $loop_radius * [Math]::Sin(($t - $loop_center_t) / $loop_width * [Math]::PI) * $loop_influence
            $x = $x_base + $loop_radius * ([Math]::Cos(($t - $loop_center_t) / $loop_width * [Math]::PI) + 1) * $loop_influence
            return [System.Windows.Media.Media3D.Point3D]::new($x, $y, $z_base)
        }
        $SyncHash.PathFunc = $script:pathFunc

        # 1. Pre-calculate all path data
        $segments = 800
        $SyncHash.CoasterSegments = $segments
        $pathDataArray = for ($i = 0; $i -le $segments; $i++) {
            $t = $i / $segments
            $p1 = & $script:pathFunc $t
            $p2 = & $script:pathFunc ($t + 0.001)
            $tangent = $p2 - $p1; $tangent.Normalize()
            $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '0,1,0')
            if ($normal.LengthSquared -lt 1e-6) { $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '1,0,0') }
            $normal.Normalize()
            $up = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($normal, $tangent)
            [pscustomobject]@{ Point = $p1; Up = $up; Normal = $normal }
        }
        $SyncHash.PathData = @{}; for($i=0; $i -lt $pathDataArray.Count; $i++){ $SyncHash.PathData[$i] = @{ Up = $pathDataArray[$i].Up } }

        # 2. Build the track model
        $trackModelGroup = New-CoasterTrackModelGroup -PathData $pathDataArray -segments $segments
        $trackContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $trackModelGroup }
        $Viewport.Children.Add($trackContainer)

        # 3. Create the sphere "cars"
        $numberOfCars = 12; $sphereRadius = 1.34
        $SyncHash.SphereRadius = $sphereRadius
        $sphereMesh = New-SphereMesh -radius $sphereRadius
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.CarStates = [hashtable]::Synchronized(@{})

        for ($i = 0; $i -lt $numberOfCars; $i++) {
            $playerKey = "Car$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null

            $SyncHash.PlayerStates[$playerKey] = [hashtable]@{
                ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
            }
            $playerState = $SyncHash.PlayerStates[$playerKey]
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()

            $viewportVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
            $viewportVisual.Geometry = $sphereMesh; $viewportVisual.Visual = $mediaHostGrid
            $material = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($material, $true)
            $viewportVisual.Material = $material
            
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $viewportVisual.Transform = $translateTransform
            
            $Viewport.Children.Add($viewportVisual)
            $SyncHash.CarStates[$i] = @{ TranslateTransform = $translateTransform; Progress = $i / $numberOfCars }
            
            Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-SphereVortex {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Sphere Vortex"
        $Viewport.Camera.Position = "0,0,15"
        $Viewport.Camera.LookDirection = "0,0,-1"
        $Viewport.Camera.FieldOfView = "70"

        $vortexContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $Viewport.Children.Add($vortexContainer) | Out-Null

        # Add a master rotation for the whole vortex that the "Random Axis" button can control
        $masterRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), 0)
        $vortexContainer.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D($masterRotation)
        $masterAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(45)) -Property @{ RepeatBehavior = "Forever" }
        $SyncHash.Animations = @{ Master = $masterAnim }; $SyncHash.Rotations = @{ Master = $masterRotation }

        $sphereCount = 32; $numberOfGroups = 6
        $sphereMesh = New-SphereMesh -radius 1.0
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        # Create the 6 unique media hosts and materials
        $sharedMaterials = @()
        for ($i = 0; $i -lt $numberOfGroups; $i++) {
            $playerKey = "Group$i"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
            $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush; Color = [System.Windows.Media.Colors]::White }

            $SyncHash.PlayerStates[$playerKey] = @{
                ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
            }
            $playerState = $SyncHash.PlayerStates[$playerKey]
            $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
            $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
            $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
            Start-NextMedia -PlayerKey $playerKey
        }
        $SyncHash.SharedMaterials = $sharedMaterials

        # Create the 32 sphere models
        $SyncHash.SphereStates = [hashtable]::Synchronized(@{})
        for ($i = 0; $i -lt $sphereCount; $i++) {
            $materialToUse = $sharedMaterials[$i % $numberOfGroups]
            $sphereMaterial = $materialToUse.Clone()

            $sphereGeometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $sphereMesh; Material = $sphereMaterial }
            $sphereContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $sphereGeometryModel }

            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
            $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform); $transformGroup.Children.Add($scaleTransform)
            $sphereContainer.Transform = $transformGroup
            
            $vortexContainer.Children.Add($sphereContainer) | Out-Null
            
            $SyncHash.SphereStates[$i] = @{
                SphereModel = $sphereGeometryModel; TranslateTransform = $translateTransform
                RotateTransform = $rotateTransform; ScaleTransform = $scaleTransform
                CurrentAngle = (720.0 / $sphereCount) * $i
            }
        }

        # Media Flow Animation Timer
        $flowTimer = New-Object System.Windows.Threading.DispatcherTimer
        $flowTimer.Interval = [TimeSpan]::FromSeconds(2.0)
        $flowTimer.Add_Tick({
            if ($SyncHash.Paused) { return }
            $lastMaterial = $SyncHash.SharedMaterials[-1]
            for ($i = $SyncHash.SharedMaterials.Count - 1; $i -gt 0; $i--) { $SyncHash.SharedMaterials[$i] = $SyncHash.SharedMaterials[$i-1] }
            $SyncHash.SharedMaterials[0] = $lastMaterial
            foreach ($i in 0..($SyncHash.SphereStates.Count-1)) { $SyncHash.SphereStates[$i].SphereModel.Material = $SyncHash.SharedMaterials[$i % $SyncHash.SharedMaterials.Count] }
        })
        $SyncHash.FlowTimer = $flowTimer; $flowTimer.Start()

        # Start the master rotation animation
        $masterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $masterAnim)
    }

    function Setup-Pie3D {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "3D Pie"
        $Viewport.Camera.Position = "0,0,8"

        $pieContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $pieModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
        $pieContainer.Content = $pieModelGroup
        $Viewport.Children.Add($pieContainer) | Out-Null

        $numberOfSlices = 8; $sliceAngle = 360.0 / $numberOfSlices; $pieRadius = 2.5; $pieCenter = New-Object System.Windows.Point(0, 0)
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $numberOfSlices; $i++) {
            $startAngle = $i * $sliceAngle
            $sliceModel = New-PieSliceModel -center $pieCenter -radius $pieRadius -startAngleDeg $startAngle -sliceAngleDeg $sliceAngle

            foreach ($face in @('Front', 'Back')) {
                $playerKey = "Slice${i}_${face}"
                $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Background = [System.Windows.Media.Brushes]::Black }
                $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; Margin='10,0,10,0'; TextAlignment='Center'; IsHitTestVisible=$false }
                $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null
                $SyncHash.PlayerStates[$playerKey] = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
                $playerState = $SyncHash.PlayerStates[$playerKey]
                $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()

                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid; Stretch = 'Fill' }
                $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                if ($face -eq 'Front') { $sliceModel.Material = $material } else { $sliceModel.BackMaterial = $material }
                Start-NextMedia -PlayerKey $playerKey
            }
            $pieModelGroup.Children.Add($sliceModel) | Out-Null
        }

        $axisAngleX = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 0)
        $axisAngleY = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), 0)
        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleX))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleY)))
        $pieContainer.Transform = $transformGroup
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(40)) -Property @{ RepeatBehavior = 'Forever' }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(60)) -Property @{ RepeatBehavior = 'Forever' }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Pinwheel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Pinwheel Kaleidoscope"
        $Viewport.Camera.Position = "0,0,8"

        # Initialize animation and rotation tracking hashtables
        $SyncHash.Animations = [hashtable]::Synchronized(@{})
        $SyncHash.Rotations = [hashtable]::Synchronized(@{})
        $SyncHash.Transforms = [hashtable]::Synchronized(@{}) # Though not used here, good practice to init

        $pinwheelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="PinwheelRotation" Axis="0,0,1" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($pinwheelContainer) | Out-Null

        $bladeCount = 8; $angleIncrement = 360 / $bladeCount; $bladeWidth = 1.5; $bladeHeight = 4.0; $bladeTiltAngle = 45; $pinwheelRadius = 1.5

        $bladeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2), (-$bladeHeight/2), 0)) )
        $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2), (-$bladeHeight/2), 0)) )
        $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2),  ($bladeHeight/2), 0)) )
        $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2),  ($bladeHeight/2), 0)) )
        $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(1); $bladeMesh.TriangleIndices.Add(2)
        $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(2); $bladeMesh.TriangleIndices.Add(3)
        $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,1)) ); $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,1)) )
        $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,0)) ); $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,0)) )

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $bladeCount; $i++) {
            $bladeAngle = $i * $angleIncrement
            $bladeContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D

            # Create front and back faces
            foreach ($face in @("Front", "Back")) {
                $playerKey = "Blade${i}_${face}"
                $mediaHostGrid = New-Object System.Windows.Controls.Grid
                $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
                $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null

                $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
                $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $playerKey }.GetNewClosure()
                $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $playerKey -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
                $SyncHash.PlayerStates[$playerKey] = $playerState

                # Create a NEW Viewport2DVisual3D for each face. This is the fix.
                $bladeViewport = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
                $bladeViewport.Geometry = $bladeMesh.Clone()
                $bladeViewport.Visual = $mediaHostGrid
                if ($face -eq "Back") { $bladeViewport.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180))) }
                
                $material = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($material, $true)
                if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                $bladeViewport.Material = $material

                $bladeContainer.Children.Add($bladeViewport) | Out-Null
                Start-NextMedia -PlayerKey $playerKey # This will trigger the media loading process
            }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            $bladeRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = '0,1,0'; Angle = 0 }
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($bladeRotation)))
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', $bladeTiltAngle)))))
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $pinwheelRadius, 0)))
            $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', $bladeAngle)))))
            $bladeContainer.Transform = $transformGroup
            $pinwheelContainer.Children.Add($bladeContainer) | Out-Null

            $bladeAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(20); RepeatBehavior = "Forever" }
            $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $bladeAnim)
            
            # Store animation and rotation objects for pause/speed control
            $SyncHash.Animations["Blade$i"] = $bladeAnim
            $SyncHash.Rotations["Blade$i"] = $bladeRotation
        }

        # Main Pinwheel Animation
        # Use a unique name to avoid conflicts if this visual is run multiple times
        $pinwheelRotationName = "PinwheelRotation_$(Get-Random)"
        $Window.RegisterName($pinwheelRotationName, $pinwheelContainer.Transform.Rotation)
        $pinwheelRotation = $Window.FindName($pinwheelRotationName)
        $pinwheelAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(60); RepeatBehavior = "Forever" }
        $pinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $pinwheelAnim)

        $SyncHash.Animations["Pinwheel"] = $pinwheelAnim
        $SyncHash.Rotations["Pinwheel"] = $pinwheelRotation
    }

    function New-ReindeerMediaObject {
        param(
            [string]$Name,
            [double]$ScaleX,
            [double]$ScaleY,
            [string]$ClipData,
            [System.Windows.Media.Media3D.MeshGeometry3D]$PlaneMesh,
            [System.Windows.Controls.Canvas]$VisualHost,
            [hashtable]$SyncHash
        )
        
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $mediaHostGrid.Width = 500; $mediaHostGrid.Height = 300
        if ($SyncHash.UseTransparentEffect) {
            $mediaHostGrid.Opacity = 0.7
        }

        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        
        if ($ClipData) {
            $geometry = [System.Windows.Media.Geometry]::Parse($ClipData).Clone()
            $bounds = $geometry.Bounds
            $scaleX_factor = if ($bounds.Width -gt 0) { $mediaHostGrid.Width / $bounds.Width } else { 1 }
            $scaleY_factor = if ($bounds.Height -gt 0) { $mediaHostGrid.Height / $bounds.Height } else { 1 }
            $scale = [Math]::Min($scaleX_factor, $scaleY_factor) * 0.95
            $translateX = ($mediaHostGrid.Width - ($bounds.Width * $scale)) / 2 - ($bounds.X * $scale)
            $translateY = ($mediaHostGrid.Height - ($bounds.Height * $scale)) / 2 - ($bounds.Y * $scale)

            $transformGroup = New-Object System.Windows.Media.TransformGroup
            $transformGroup.Children.Add((New-Object System.Windows.Media.ScaleTransform($scale, $scale)))
            $transformGroup.Children.Add((New-Object System.Windows.Media.TranslateTransform($translateX, $translateY)))
            $geometry.Transform = $transformGroup
            $mediaHostGrid.Clip = $geometry

            $borderPath = New-Object System.Windows.Shapes.Path
            $borderPath.Data = $geometry
            $borderPath.Stroke = [System.Windows.Media.Brushes]::LightGray
            $borderPath.StrokeThickness = 2
            $borderPath.Fill = [System.Windows.Media.Brushes]::Transparent
            $mediaHostGrid.Children.Add($borderPath) | Out-Null
        }

        if ($Name -eq "Deer1") {
            $nose = New-Object System.Windows.Shapes.Ellipse -Property @{
                Width = 24; Height = 24; Fill = [System.Windows.Media.Brushes]::Red
                HorizontalAlignment = 'Left'; VerticalAlignment = 'Top'
                Margin = [System.Windows.Thickness]::new(91.5, 81.5, 0, 0)
            }
            $mediaHostGrid.Children.Add($nose) | Out-Null
            $blinkAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(1.0, 0.1, [TimeSpan]::FromSeconds(1.5)) -Property @{ AutoReverse = $true; RepeatBehavior = "Forever" }
            $nose.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $blinkAnimation)
        }

        $VisualHost.Children.Add($mediaHostGrid) | Out-Null

        $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; MediaHostGrid = $mediaHostGrid; IsFailed = $false; IsImage = $false; CurrentSource = $null }
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $Name }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $Name -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $Name -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates[$Name] = $playerState

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }
        if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $PlaneMesh; Material = $material }
        $model.Transform = New-Object System.Windows.Media.Media3D.ScaleTransform3D($ScaleX, $ScaleY, 1)
        $backMaterial = $material.Clone(); $backTransform = New-Object System.Windows.Media.ScaleTransform -Property @{ ScaleX = -1 }; $backMaterial.Brush.RelativeTransform = $backTransform
        $model.BackMaterial = $backMaterial

        $objectContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, 0)
        $objectContainer.Transform = $translateTransform
        $SyncHash.PlayerStates[$Name].TranslateTransform = $translateTransform

        return $objectContainer
    }

    function Setup-ReindeerSleigh {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Reindeer Sleigh"
        $Viewport.Camera.Position = "0,0,20"

        if ($SyncHash.NightSky) {
            $backgroundCanvas = $Window.FindName("backgroundCanvas")
            if ($SyncHash.UseTransparentEffect) {
                $backgroundCanvas.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(128, 0, 0, 0))
            } else {
                $backgroundCanvas.Background = [System.Windows.Media.Brushes]::Black
            }
        }

        if ($SyncHash.TwinklingStars) {
            $starCanvas = $Window.FindName("backgroundCanvas")
            $rand = [Random]::new()
            for ($i = 0; $i -lt 200; $i++) {
                $star = New-Object System.Windows.Shapes.Ellipse
                $size = $rand.NextDouble() * 3 + 1
                $star.Width = $size; $star.Height = $size
                $star.Fill = [System.Windows.Media.Brushes]::White
                [System.Windows.Controls.Canvas]::SetLeft($star, $rand.NextDouble() * $Window.Width)
                [System.Windows.Controls.Canvas]::SetTop($star, $rand.NextDouble() * $Window.Height)
                
                $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0.1, 1.0, [TimeSpan]::FromSeconds($rand.NextDouble() * 2 + 0.5)) -Property @{ AutoReverse = $true; RepeatBehavior = "Forever" }
                $star.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $anim)
                $starCanvas.Children.Add($star) | Out-Null
            }
        }

        $aspectRatio = $Window.Width / $Window.Height; $fovRadians = 60 * ([Math]::PI / 180)
        $visibleHeight = 2 * 20 * [Math]::Tan($fovRadians / 2); $visibleWidth = $visibleHeight * $aspectRatio
        $SyncHash.RightLimit = ($visibleWidth / 2); $SyncHash.LeftLimit = -($visibleWidth / 2) - 12
        $SyncHash.LeadPositionX = $SyncHash.RightLimit
        $SyncHash.HistoryX = [System.Collections.Generic.List[double]]::new()
        $SyncHash.HistoryY = [System.Collections.Generic.List[double]]::new()
        $SyncHash.HistoryDist = [System.Collections.Generic.List[double]]::new()

        $planeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D -Property @{
            Positions = "-0.5,-0.5,0 0.5,-0.5,0 0.5,0.5,0 -0.5,0.5,0"; TriangleIndices = "0,1,2 0,2,3"; TextureCoordinates = "0,1 1,1 1,0 0,0"
        }
        $planeMesh.Freeze()

        $sleighPath = "M27,24 c0.6,0,1-0.4,1-1 c0-3.8,0.9-7.5,2.6-10.9 l1.3-2.6 c0.2-0.3,0.1-0.7,0-1 C31.7,8.2,31.3,8,31,8 L28,11 C27.6,10.8 27.1,8.6 26.4,7.2 C26.4,6 25.4,2 24,1.0 C21.6,5.6 20.8,7.2 21.2,8.8 C21.4,9.8 22.2,10.6 23.2,11.2 C22.2,11.6 21.2,12.6 21,14 C20.8,15.6 21.6,16.6 22.8,17.2 L21,16.4 c-0.7,1-1.8,1.5-2.9,1.5 c-1.3,0-2.4-0.7-3.1-1.7 c-1.2-2-3.4-3.3-5.7-3.3 H6 c-0.4,0-0.7,0.2-0.9,0.5 c-0.2,0.3-0.2,0.7,0,1 l5,9 c0.2,0.3,0.5,0.5,0.9,0.5 h1.5 l-1.2,3 l-5.8,0 c-1,0-2-0.3-2.7-1 C2.3,25.5,2,24.9,2,24.3 s0.3-1.3,0.8-1.7 c0.8-0.7,2.1-0.7,2.8,0 c0.4,0.4,1,0.4,1.4,0 c0.4-0.4,0.4-1,0-1.4 c-1.5-1.4-4-1.4-5.6,0 c-0.9,0.8-1.4,2-1.4,3.2 s0.5,2.3,1.4,3.2 c1.1,1,2.5,1.5,3.9,1.5 c0.1,0,0.1,0,0.2,0 l6.4,0 l13,0 l5,0 c0.6,0,1-0.4,1-1 s-0.4-1-1-1 l-4.3,0 l-1.2-3 H27 z"
        $reindeerPath = "M505.474,436.173c0,0-64.29-87.208-86.987-116.34c-10.781-13.882-21.414-24.368-32.196-32.048 c-0.055-0.037-0.101-0.064-0.157-0.101c0.804-0.139,1.468-0.148,2.299-0.305c26.953-5.28,33.562-41.906,31.466-57.006 c-2.086-15.101-8.538-20.528-18.313-3.286c-7.763,13.624-34.641,43.272-70.234,43.124l-58.041-0.184 c-3.775-0.037-7.42-0.147-10.928-0.37c-21.267-1.439-37.882-7.42-54.497-24.147c-8.916-8.999-12.627-24.718-15.507-39.034 c6.849,0.028,15.23,0.055,19.836,0.082c7.699,0,11.575-1.92,16.421-6.71c4.865-4.809,8.64-11.52,14.104-20.242 c6.019-9.674,0.036-18.276-6.72-18.313c-6.756,0-30.424-0.111-30.424-0.111c-5.787,0-11.592,1.92-14.51,3.84 c-2.427,1.597-7.31,5.584-12.572,10.08l-91.62-0.332c-12.444-0.074-22.855,1.957-31.272,5.427 c-13.542,5.575-22.117,14.99-26.206,25.993c-12.304,32.943,15.572,79.972,70.105,80.157l30.599,0.111 c9.609,31.568,25.789,65.72,25.789,65.72l-80.756,94.27c-7.108,5.917-9.831,19.236-1.532,27.562 c8.27,8.326,18.017,8.086,33.441-3.729l96.642-76.824l180.897,0.48l7.606,5.501l87.042,63.911 c11.852,8.935,21.598,7.145,27.894,0.083C514.252,455.464,513.44,445.394,505.474,436.173z M37.974,166.148c0.037-10.495-8.419-19.042-18.95-19.07C8.565,147.049,0.056,155.515,0,165.991 c-0.027,10.494,8.418,19.032,18.922,19.069C29.408,185.089,37.937,176.633,37.974,166.148z M154.277,142.093c-7.191,8.603-14.197,14.805-20.095,19.08h37.817c7.651-8.806,15.34-19.624,22.412-32.74 c11.261-20.897,20.934-47.62,26.361-81.19c1.145-7.052-3.655-13.689-10.67-14.842c-7.089-1.145-13.734,3.618-14.842,10.707 c-1.736,10.531-3.877,20.233-6.342,29.204l-17.722-18.424c-4.994-5.132-13.154-5.27-18.313-0.333 c-5.141,4.948-5.299,13.145-0.333,18.276l26.602,27.618c0.12,0.111,0.231,0.184,0.314,0.296 c-3.268,7.79-6.812,14.806-10.421,21.082l-16.541-11.926c-5.797-4.162-13.883-2.88-18.055,2.926 c-4.172,5.787-2.88,13.882,2.917,18.046L154.277,142.093z"

        $sleigh = New-ReindeerMediaObject -Name "Sleigh" -ScaleX 5 -ScaleY 3 -ClipData $sleighPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer1 = New-ReindeerMediaObject -Name "Deer1" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer2 = New-ReindeerMediaObject -Name "Deer2" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer3 = New-ReindeerMediaObject -Name "Deer3" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer4 = New-ReindeerMediaObject -Name "Deer4" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash

        $Viewport.Children.Add($sleigh) | Out-Null
        $Viewport.Children.Add($deer1) | Out-Null
        $Viewport.Children.Add($deer2) | Out-Null
        $Viewport.Children.Add($deer3) | Out-Null
        $Viewport.Children.Add($deer4) | Out-Null

        $SyncHash.PlayerStates.Keys | ForEach-Object { Start-NextMedia -PlayerKey $_ }
    }

    function Animate-ReindeerSleigh {
        param($SyncHash)
        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency

        $flySpeed = 4.0 * $SyncHash.SpeedMultiplier
        $bobSpeed = 2.0; $bobAmount = 1.5
        $distStep = $flySpeed * $elapsed
        $SyncHash.TotalDistance += $distStep

        $SyncHash.LeadPositionX -= ($flySpeed * $elapsed)
        
        if ($SyncHash.LeadPositionX -lt $SyncHash.LeftLimit) {
            $lapDistanceX = $SyncHash.RightLimit - $SyncHash.LeftLimit
            $SyncHash.LeadPositionX += $lapDistanceX
            for ($i = 0; $i -lt $SyncHash.HistoryX.Count; $i++) {
                $SyncHash.HistoryX[$i] += $lapDistanceX
            }
        }

        $SyncHash.LeadPositionY = [Math]::Sin($totalTime * $bobSpeed) * $bobAmount

        $SyncHash.HistoryX.Add($SyncHash.LeadPositionX)
        $SyncHash.HistoryY.Add($SyncHash.LeadPositionY)
        $SyncHash.HistoryDist.Add($SyncHash.TotalDistance)

        while ($SyncHash.HistoryDist.Count -gt 0 -and ($SyncHash.TotalDistance - $SyncHash.HistoryDist[0] -gt 25.0)) {
            $SyncHash.HistoryX.RemoveAt(0)
            $SyncHash.HistoryY.RemoveAt(0)
            $SyncHash.HistoryDist.RemoveAt(0)
        }

        $lags = @(0, 4, 8, 12, 17.7)
        $objects = @("Deer1", "Deer2", "Deer3", "Deer4", "Sleigh")
        
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $playerState = $SyncHash.PlayerStates[$objects[$i]]
            if (-not $playerState) { continue }

            $targetDist = $SyncHash.TotalDistance - $lags[$i]
            
            $idx = -1
            for ($j = $SyncHash.HistoryDist.Count - 1; $j -ge 0; $j--) {
                if ($SyncHash.HistoryDist[$j] -le $targetDist) {
                    $idx = $j
                    break
                }
            }

            if ($idx -ge 0) {
                $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[$idx]
                $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[$idx]
            } elseif ($SyncHash.HistoryX.Count -gt 0) {
                $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[0] + ($SyncHash.HistoryDist[0] - $targetDist)
                $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[0]
            }
        }
    }

    function Animate-ButterflyEffect {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "ButterflyEffect") { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency

        foreach ($plane in $SyncHash.FloatingObjects) {
            if (-not $plane.PSObject.Properties['FlutterTransform']) { continue } # Extra safety
            $flutterSpeed = 8.0; $flutterAmount = 15.0
            $plane.FlutterTransform.Rotation.Angle = [Math]::Sin($totalTime * $flutterSpeed) * $flutterAmount
        }
    }

    function Animate-Aquarium {
        param($SyncHash, $Viewport)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "Aquarium") { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        if (-not $SyncHash.StartTime) { $SyncHash.StartTime = $currentTime } # Safety for first frame
        $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency

        $camera = $Viewport.Camera
        # Use pre-calculated boundaries from the Loaded event to ensure they are accurate
        $xBoundary = $SyncHash.xBoundary; $yBoundary = $SyncHash.yBoundary; $zBoundary = 3

        # Define the margins based on half the fish's model size
        $xMargin = 0.7; $yMargin = 0.3

        $oscillation = [Math]::Sin($totalTime * 18.0) * 0.12

        foreach ($fish in $SyncHash.FloatingObjects) {
            $velX = $fish.Velocity.X; $velY = $fish.Velocity.Y; $velZ = $fish.Velocity.Z
            $nextX = $fish.Translate.OffsetX + ($velX * $elapsed); $nextY = $fish.Translate.OffsetY + ($velY * $elapsed); $nextZ = $fish.Translate.OffsetZ + ($velZ * $elapsed)

            if (($nextX + $xMargin) -gt $xBoundary -and $velX -gt 0) { $velX *= -1; $nextX = $xBoundary - $xMargin }
            elseif (($nextX - $xMargin) -lt -$xBoundary -and $velX -lt 0) { $velX *= -1; $nextX = -$xBoundary + $xMargin }
            if (($nextY + $yMargin) -gt $yBoundary -and $velY -gt 0) { $velY *= -1; $nextY = $yBoundary - $yMargin }
            elseif (($nextY - $yMargin) -lt -$yBoundary -and $velY -lt 0) { $velY *= -1; $nextY = -$yBoundary + $yMargin }
            if (($nextZ -gt $zBoundary -and $velZ -gt 0) -or ($nextZ -lt -$zBoundary -and $velZ -lt 0)) { $velZ *= -1 }
            $fish.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)

            $fish.Translate.OffsetX = $nextX; $fish.Translate.OffsetY = $nextY; $fish.Translate.OffsetZ = $nextZ

            $frontModel = $fish.Visual.Children[0].Content; $backModel = $fish.Visual.Children[1].Content
            $currentGeometry = if ($fish.Velocity.X -gt 0) { $fish.RightGeometry } else { $fish.LeftGeometry }
            if ($frontModel.Geometry -ne $currentGeometry) { $frontModel.Geometry = $currentGeometry; $backModel.Geometry = $currentGeometry }

            $positions = $currentGeometry.Positions
            $zWiggle = if ($fish.Velocity.X -gt 0) { $oscillation } else { -$oscillation }
            $positions[6] = [System.Windows.Media.Media3D.Point3D]::new($positions[6].X, $positions[6].Y, $zWiggle)
            $positions[8] = [System.Windows.Media.Media3D.Point3D]::new($positions[8].X, $positions[8].Y, $zWiggle)
            $positions[5] = [System.Windows.Media.Media3D.Point3D]::new($positions[5].X, $positions[5].Y, $zWiggle * 0.6)
            $positions[9] = [System.Windows.Media.Media3D.Point3D]::new($positions[9].X, $positions[9].Y, $zWiggle * 0.6)
        }
    }

    function Animate-RollerCoaster {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "RollerCoaster") { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        $flowSpeed = 0.10 * $SyncHash.SpeedMultiplier

        foreach ($i in $SyncHash.CarStates.Keys) {
            $carState = $SyncHash.CarStates[$i]
            $carState.Progress = ($carState.Progress + ($flowSpeed * $elapsed)) % 1.0 # Update and loop progress
            $positionOnPath = & $SyncHash.PathFunc $carState.Progress
            
            $pathIndex = [int]($carState.Progress * ($SyncHash.CoasterSegments - 1))
            $upVector = $SyncHash.PathData[$pathIndex].Up
            $finalPosition = $positionOnPath + ($upVector * ($SyncHash.SphereRadius + 0.05)) # Add rail thickness to offset
            
            $carState.TranslateTransform.OffsetX = $finalPosition.X; $carState.TranslateTransform.OffsetY = $finalPosition.Y; $carState.TranslateTransform.OffsetZ = $finalPosition.Z
        }
    }

    function Animate-SphereVortex {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "SphereVortex") { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier
        $maxScale = 1.2; $minScale = 0.1; $totalAngle = 720

        foreach ($i in $SyncHash.SphereStates.Keys) {
            $panelState = $SyncHash.SphereStates[$i]
            $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle
            $angle = $panelState.CurrentAngle; $angleRad = $angle * ([Math]::PI / 180.0)
            $progress = $angle / $totalAngle
            $currentScale = $maxScale - ($progress * ($maxScale - $minScale))

            $startRadius = 3.5; $endRadius = 0.5; $startY = 5.0; $endY = -20.0
            $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius))
            $currentY = $startY - ($progress * ($startY - $endY))
            $currentX = $currentRadius * [Math]::Cos($angleRad); $currentZ = $currentRadius * [Math]::Sin($angleRad)
            $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ
            $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, $currentY - 1.5, 0)
            $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)

            $panelState.ScaleTransform.ScaleX = $currentScale; $panelState.ScaleTransform.ScaleY = $currentScale; $panelState.ScaleTransform.ScaleZ = $currentScale
            $forward = $position - $lookAtTarget; if ($forward.Length -gt 0) { $forward.Normalize() }
            $up = '0,1,0'
            $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward); if ($right.Length -gt 0) { $right.Normalize() }
            $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)

            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity
            $matrix.M11 = $right.X;   $matrix.M12 = $right.Y;   $matrix.M13 = $right.Z
            $matrix.M21 = $newUp.X;   $matrix.M22 = $newUp.Y;   $matrix.M23 = $newUp.Z
            $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z

            # Correctly convert the Matrix3D to a Quaternion
            $trace = $matrix.M11 + $matrix.M22 + $matrix.M33
            if ($trace -gt 0) {
                $s = 0.5 / [Math]::Sqrt($trace + 1.0); $qw = 0.25 / $s; $qx = ($matrix.M32 - $matrix.M23) * $s; $qy = ($matrix.M13 - $matrix.M31) * $s; $qz = ($matrix.M21 - $matrix.M12) * $s
            } elseif (($matrix.M11 -gt $matrix.M22) -and ($matrix.M11 -gt $matrix.M33)) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M11 - $matrix.M22 - $matrix.M33); $qw = ($matrix.M32 - $matrix.M23) / $s; $qx = 0.25 * $s; $qy = ($matrix.M12 + $matrix.M21) / $s; $qz = ($matrix.M13 + $matrix.M31) / $s
            } elseif ($matrix.M22 -gt $matrix.M33) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M22 - $matrix.M11 - $matrix.M33); $qw = ($matrix.M13 - $matrix.M31) / $s; $qx = ($matrix.M12 + $matrix.M21) / $s; $qy = 0.25 * $s; $qz = ($matrix.M23 + $matrix.M32) / $s
            } else {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M33 - $matrix.M11 - $matrix.M22); $qw = ($matrix.M21 - $matrix.M12) / $s; $qx = ($matrix.M13 + $matrix.M31) / $s; $qy = ($matrix.M23 + $matrix.M32) / $s; $qz = 0.25 * $s
            }
            $quaternion = New-Object System.Windows.Media.Media3D.Quaternion($qx, $qy, $qz, $qw)
            $panelState.RotateTransform.Rotation = (New-Object System.Windows.Media.Media3D.QuaternionRotation3D($quaternion))
        }
    }

    function Animate-CurvedVortex {
        param($SyncHash)
        if ($SyncHash.Paused -or ($SyncHash.VisualizationStyle -ne "CurvedVortex")) { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier
        $startRadius = 6.0; $endRadius = 1.0; $startY = 4.0; $endY = -6.0; $totalAngle = 720

        foreach ($key in $SyncHash.PlayerStates.Keys) {
            $panelState = $SyncHash.PlayerStates[$key]
            $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle
            $angle = $panelState.CurrentAngle; $angleRad = $angle * ([Math]::PI / 180.0)
            $progress = $angle / $totalAngle
            
            $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius))
            $currentY = $startY - ($progress * ($startY - $endY))
            $currentX = $currentRadius * [Math]::Cos($angleRad); $currentZ = $currentRadius * [Math]::Sin($angleRad)

            $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ

            $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, $currentY - 1.5, 0)
            $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)
            $forward = $lookAtTarget - $position; if ($forward.Length -gt 0) { $forward.Normalize() }
            $up = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0)
            $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward); if ($right.Length -gt 0) { $right.Normalize() }
            $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)

            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity
            $matrix.M11 = $right.X;   $matrix.M12 = $right.Y;   $matrix.M13 = $right.Z
            $matrix.M21 = $newUp.X;   $matrix.M22 = $newUp.Y;   $matrix.M23 = $newUp.Z
            $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z

            $trace = $matrix.M11 + $matrix.M22 + $matrix.M33
            if ($trace -gt 0) {
                $s = 0.5 / [Math]::Sqrt($trace + 1.0); $qw = 0.25 / $s; $qx = ($matrix.M32 - $matrix.M23) * $s; $qy = ($matrix.M13 - $matrix.M31) * $s; $qz = ($matrix.M21 - $matrix.M12) * $s
            } elseif (($matrix.M11 -gt $matrix.M22) -and ($matrix.M11 -gt $matrix.M33)) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M11 - $matrix.M22 - $matrix.M33); $qw = ($matrix.M32 - $matrix.M23) / $s; $qx = 0.25 * $s; $qy = ($matrix.M12 + $matrix.M21) / $s; $qz = ($matrix.M13 + $matrix.M31) / $s
            } elseif ($matrix.M22 -gt $matrix.M33) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M22 - $matrix.M11 - $matrix.M33); $qw = ($matrix.M13 - $matrix.M31) / $s; $qx = ($matrix.M12 + $matrix.M21) / $s; $qy = 0.25 * $s; $qz = ($matrix.M23 + $matrix.M32) / $s
            } else {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M33 - $matrix.M11 - $matrix.M22); $qw = ($matrix.M21 - $matrix.M12) / $s; $qx = ($matrix.M13 + $matrix.M31) / $s; $qy = ($matrix.M23 + $matrix.M32) / $s; $qz = 0.25 * $s
            }
            $panelState.RotateTransform.Rotation = (New-Object System.Windows.Media.Media3D.QuaternionRotation3D([System.Windows.Media.Media3D.Quaternion]::new($qx, $qy, $qz, $qw)))
        }
    }

    function Animate-FloatingObjects {
        param($SyncHash, $Viewport)
        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        $xBoundary = 9; $yBoundary = 5; $zBoundary = 3

        foreach ($obj in $SyncHash.FloatingObjects) {
            $newX = $obj.Translate.OffsetX + ($obj.Velocity.X * $elapsed * $SyncHash.SpeedMultiplier)
            $newY = $obj.Translate.OffsetY + ($obj.Velocity.Y * $elapsed * $SyncHash.SpeedMultiplier)
            $newZ = $obj.Translate.OffsetZ + ($obj.Velocity.Z * $elapsed * $SyncHash.SpeedMultiplier)

            # Extra safety check for properties before using them
            if (-not $obj.PSObject.Properties['CurrentRotation']) { continue }

            $velX = $obj.Velocity.X; $velY = $obj.Velocity.Y; $velZ = $obj.Velocity.Z

            if (($newX -gt $xBoundary -and $velX -gt 0) -or ($newX -lt -$xBoundary -and $velX -lt 0)) { $velX *= -1 }
            if (($newY -gt $yBoundary -and $velY -gt 0) -or ($newY -lt -$yBoundary -and $velY -lt 0)) { $velY *= -1 }
            if (($newZ -gt $zBoundary -and $velZ -gt 0) -or ($newZ -lt -$zBoundary -and $velZ -lt 0)) { $velZ *= -1 }
            $obj.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)

            $obj.Translate.OffsetX = $newX; $obj.Translate.OffsetY = $newY; $obj.Translate.OffsetZ = $newZ

            $obj.CurrentRotation *= New-Object System.Windows.Media.Media3D.Quaternion((New-Object System.Windows.Media.Media3D.Vector3D(1,0,0)), ($obj.RotationVelocity.X * $elapsed * $SyncHash.SpeedMultiplier))
            $obj.CurrentRotation *= New-Object System.Windows.Media.Media3D.Quaternion((New-Object System.Windows.Media.Media3D.Vector3D(0,1,0)), ($obj.RotationVelocity.Y * $elapsed * $SyncHash.SpeedMultiplier))
            $obj.Rotate.Rotation = New-Object System.Windows.Media.Media3D.QuaternionRotation3D($obj.CurrentRotation)
        }
    }

    # --- Main Window Setup ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Unified MediaElement Visualizer"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Canvas x:Name="backgroundCanvas" />
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,10" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>
        </Viewport3D> 
        <Canvas x:Name="VisualHost" Opacity="0"/>
        <Canvas x:Name="scrollingCanvas" Visibility="Collapsed" ClipToBounds="True">
            <ScrollViewer x:Name="scrollingViewer" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden">
                <StackPanel x:Name="scrollingPanel" VerticalAlignment="Top" HorizontalAlignment="Center">
                    <!-- The RenderTransform will be set in code -->
                </StackPanel>
            </ScrollViewer>
        </Canvas>
        <Canvas x:Name="bubbleCanvas" IsHitTestVisible="False" />
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="pauseButton" Content="Pause" Padding="10,5" Margin="2"/>
            <Button Name="randomAxisButton" Content="Random Axis" Padding="10,5" Margin="2"/>
            <Button Name="slowDownButton" Content="&#x2190;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="speedUpButton" Content="&#x2192;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="redoButton" Content="Redo" Padding="10,5" Margin="2"/>
            <Button Name="hideControlsButton" Content="Hide Controls" Padding="10,5" Margin="2"/>
            <Button Name="closeButton" Content="X" Padding="10,5" Margin="2" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $SyncHash.Window = $window
    $viewport = $window.FindName("mainViewport")

    if ($SyncHash.VisualizationStyle -eq "Aquarium" -and $SyncHash.AddWater) {
        $waterColor = [System.Windows.Media.ColorConverter]::ConvertFromString("#80ADD8E6")
        $window.Background = [System.Windows.Media.SolidColorBrush]::new($waterColor)
    }
    $visualHost = $window.FindName("VisualHost")

    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width; $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left; $window.Top = $primaryScreen.WorkingArea.Top

    # --- Call the selected visualization's setup function ---
    switch ($SyncHash.VisualizationStyle) {
        "RotatingCube" { Setup-RotatingCube -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FloatingCubes" { Setup-FloatingCubes -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Sphere" { Setup-Sphere -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FloatingStars" { Setup-FloatingStars -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "PulsingStar" { Setup-PulsingStar -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "WagonWheel" { Setup-WagonWheel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FunnelMulti" { Setup-Funnel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "ButterflyEffect" { Setup-ButterflyEffect -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "ConcentricFunnel" { Setup-ConcentricFunnel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "RotatingStar" { Setup-RotatingStar -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Carousel" { Setup-Carousel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FacetedSphereMulti" { Setup-FacetedSphereMulti -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Aquarium" { Setup-Aquarium -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "MediaFlowFunnel" { Setup-MediaFlowFunnel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FacetedSphereSingle" { Setup-FacetedSphereSingle -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FloatingSpheres" { Setup-FloatingSpheres -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Pie3D" { Setup-Pie3D -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Pinwheel" { Setup-Pinwheel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "SphereVortex" { Setup-SphereVortex -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "CurvedVortex" { Setup-CurvedVortex -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "FunnelSingle" { Setup-FunnelSingle -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "RollerCoaster" { Setup-RollerCoaster -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "Funnel" { Setup-Funnel -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "ScrollingHorizontal" { Setup-ScrollingHorizontal -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "ScrollingVertical" { Setup-ScrollingVertical -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        "ReindeerSleigh" { Setup-ReindeerSleigh -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
        default { Setup-RotatingCube -Window $window -Viewport $viewport -VisualHost $visualHost -SyncHash $SyncHash }
    }

    # --- Common UI Event Handlers ---
    $SyncHash.closeButton = $window.FindName("closeButton"); $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton = $window.FindName("redoButton"); $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton"); $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        $SyncHash.ControlsHidden = -not $SyncHash.ControlsHidden
        $controlsPanel.Visibility = if ($SyncHash.ControlsHidden) { 'Collapsed' } else { 'Visible' }
    })
    
    $SyncHash.pauseButton = $window.FindName("pauseButton"); $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
            if ($SyncHash.VisualizationStyle -eq "FloatingCubes") {
                foreach ($animSet in $SyncHash.Animations.Values) {
                    $rotation = $animSet.RotateTransform.Rotation
                    $translation = $animSet.TranslateTransform
                    $currentAngle = $rotation.Angle; $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $rotation.Angle = $currentAngle
                    $currentX = $translation.OffsetX; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $null); $translation.OffsetX = $currentX
                    $currentY = $translation.OffsetY; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $null); $translation.OffsetY = $currentY
                    $currentZ = $translation.OffsetZ; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $null); $translation.OffsetZ = $currentZ
                }
            } elseif ($SyncHash.Storyboard) {
                $SyncHash.Storyboard.Pause()
            } else {
                if ($SyncHash.Rotations) {
                    foreach ($key in $SyncHash.Rotations.Keys) {
                        $rotation = $SyncHash.Rotations[$key]
                        $currentAngle = $rotation.Angle; $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $rotation.Angle = $currentAngle
                    }
                }
                if ($SyncHash.Transforms) {
                    $currentScale = $SyncHash.Transforms.Pulse.ScaleX
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $null); $SyncHash.Transforms.Pulse.ScaleX = $currentScale
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $null); $SyncHash.Transforms.Pulse.ScaleY = $currentScale
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $null); $SyncHash.Transforms.Pulse.ScaleZ = $currentScale
                }
            }
        } else {
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
            $SyncHash.pauseButton.Content = "Pause"
            if ($SyncHash.VisualizationStyle -eq "FloatingCubes") {
                foreach ($animSet in $SyncHash.Animations.Values) {
                    $animSet.Rotation.From = $animSet.RotateTransform.Rotation.Angle
                    $animSet.RotateTransform.Rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animSet.Rotation)
                    $animSet.TranslateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $animSet.PositionX)
                    $animSet.TranslateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $animSet.PositionY)
                    $animSet.TranslateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $animSet.PositionZ)
                }
            } elseif ($SyncHash.Storyboard) {
                $SyncHash.Storyboard.Resume()
            } else {
                if ($SyncHash.Rotations) {
                    foreach ($key in $SyncHash.Rotations.Keys) {
                        $SyncHash.Animations[$key].From = $SyncHash.Rotations[$key].Angle
                        $SyncHash.Rotations[$key].BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.Animations[$key])
                    }
                }
                if ($SyncHash.Transforms) {
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $SyncHash.Animations.Pulse)
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $SyncHash.Animations.Pulse)
                    $SyncHash.Transforms.Pulse.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $SyncHash.Animations.Pulse)
                }
            }
        }
    })

    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton"); $SyncHash.randomAxisButton.Add_Click({
        if ($SyncHash.VisualizationStyle -eq "FloatingCubes") {
            foreach ($animSet in $SyncHash.Animations.Values) {
                $animSet.RotateTransform.Rotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
            }
        } elseif ($SyncHash.Rotations) {
            $newAxis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
            foreach ($key in $SyncHash.Rotations.Keys) { $SyncHash.Rotations[$key].Axis = $newAxis }
        }
    })

    $changeSpeed = {
        param($multiplier)
        $wasPaused = $SyncHash.Paused
        if (-not $wasPaused) { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        
        if ($SyncHash.VisualizationStyle -eq "FloatingCubes") {
            foreach ($animSet in $SyncHash.Animations.Values) {
                $animSet.Rotation.Duration = [TimeSpan]::FromSeconds($animSet.Rotation.Duration.TimeSpan.TotalSeconds * $multiplier)
                $animSet.PositionX.Duration = [TimeSpan]::FromSeconds($animSet.PositionX.Duration.TimeSpan.TotalSeconds * $multiplier)
                $animSet.PositionY.Duration = [TimeSpan]::FromSeconds($animSet.PositionY.Duration.TimeSpan.TotalSeconds * $multiplier)
                $animSet.PositionZ.Duration = [TimeSpan]::FromSeconds($animSet.PositionZ.Duration.TimeSpan.TotalSeconds * $multiplier)
            }
        }
        elseif ($SyncHash.Animations) {
            foreach ($key in $SyncHash.Animations.Keys) {
                $baseDuration = switch($key) { "X" {20} "Y" {15} "Star" {30} "Pulse" {2} default {20} }
                $SyncHash.Animations[$key].Duration = [TimeSpan]::FromSeconds($baseDuration / $SyncHash.SpeedMultiplier)
            }
        }
        
        $SyncHash.SpeedMultiplier /= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        if ($SyncHash.Animations) {
            foreach ($key in $SyncHash.Animations.Keys) {
                $baseDuration = switch($key) { "X" {20} "Y" {15} "Star" {30} "Pulse" {2} default {20} }
                $SyncHash.Animations[$key].Duration = [TimeSpan]::FromSeconds($baseDuration / $SyncHash.SpeedMultiplier)
            }
        }
        
        if (-not $wasPaused) { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
    }
    $SyncHash.slowDownButton = $window.FindName("slowDownButton"); $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton = $window.FindName("speedUpButton"); $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })

    $window.Add_KeyDown({
        param($s, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "A" { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $window.Add_Closed({
        if ($animationLoop) {
            [System.Windows.Media.CompositionTarget]::remove_Rendering($animationLoop)
            $animationLoop = $null
        }

        # Stop the watchdog timer
        if ($SyncHash.WatchdogTimer) {
            $SyncHash.WatchdogTimer.Stop()
        }

        # Stop the SphereVortex-specific timer if it exists
        if ($SyncHash.FlowTimer) {
            $SyncHash.FlowTimer.Stop()
        }

        if ($animationLoop) { [System.Windows.Media.CompositionTarget]::remove_Rendering($animationLoop) }
        Cleanup-Visualization -SyncHash $SyncHash
        if ($SyncHash.VisualizationStyle -eq "FloatingCubes" -and $SyncHash.Animations) {
            foreach ($animSet in $SyncHash.Animations.Values) {
                if ($animSet) {
                    $animSet.Rotation.BeginTime = $null; $animSet.PositionX.BeginTime = $null; $animSet.PositionY.BeginTime = $null; $animSet.PositionZ.BeginTime = $null
                }
            }
        }
    })

    $window.Add_Loaded({
        # --- Configure UI Controls based on Visualization Style ---
        # By default, all controls are visible. This logic hides controls that are not applicable to the current style.
        switch ($SyncHash.VisualizationStyle) {
            "Aquarium"            { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ButterflyEffect"     { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "Carousel"            { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ConcentricFunnel"    { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "CurvedVortex"        { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingCubes"       { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingSpheres"     { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingStars"       { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "MediaFlowFunnel"     { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "RollerCoaster"       { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "SphereVortex"        { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ReindeerSleigh"      { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ScrollingHorizontal" { $SyncHash.randomAxisButton.Visibility = 'Collapsed'; $SyncHash.slowDownButton.Visibility = 'Collapsed'; $SyncHash.speedUpButton.Visibility = 'Collapsed' }
            "ScrollingVertical"   { $SyncHash.randomAxisButton.Visibility = 'Collapsed'; $SyncHash.slowDownButton.Visibility = 'Collapsed'; $SyncHash.speedUpButton.Visibility = 'Collapsed' }
        }

        if ($SyncHash.RbSelection -ne "Hidden") {
            # --- Apply Text Overlay Settings ---
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            foreach ($key in $SyncHash.PlayerStates.Keys) {
                $textBlock = $SyncHash.PlayerStates[$key].OverlayTextBlock
                if ($textBlock) {
                    $textBlock.Foreground = $brush.Clone(); $textBlock.FontFamily = $fontFamily; $textBlock.FontSize = $SyncHash.FontSize
                    $textBlock.FontWeight = $fontWeight; $textBlock.FontStyle = $fontStyle
                    if ($SyncHash.RbSelection -eq "Custom") { $textBlock.Text = $SyncHash.CustomText }
                }
            }
        }

        if ($SyncHash.VisualizationStyle -eq "Aquarium") {
            $camera = $viewport.Camera
            $distance = $camera.Position.Z
            $fovRadians = 45.0 * ([Math]::PI / 180.0)
            $viewHeight3D = 2.0 * $distance * [Math]::Tan($fovRadians / 2.0)
            $aspectRatio = if ($viewport.ActualHeight -gt 0) { $viewport.ActualWidth / $viewport.ActualHeight } else { 1.6 }
            $SyncHash.xBoundary = ((($viewHeight3D * $aspectRatio) / 2.0) / 2.0) * 1.2
            $SyncHash.yBoundary = (($viewHeight3D / 2.0) / 2.0) * 1.2
        }

        if ($SyncHash.VisualizationStyle -eq "MediaFlowFunnel") {
            $flowTimer = New-Object System.Windows.Threading.DispatcherTimer
            $flowTimer.Interval = [TimeSpan]::FromSeconds(2)
            $SyncHash.FlowTimer = $flowTimer
            $flowTimer.Add_Tick({
                if ($SyncHash.Paused) { return }
                $panelGroupsByRing = $SyncHash.PanelModels.GetEnumerator() | Group-Object { [math]::Floor($_.Name / $SyncHash.PanelsPerRing) }
                $lastMaterial = $SyncHash.SharedMaterials[-1]
                for ($i = $SyncHash.SharedMaterials.Count - 1; $i -gt 0; $i--) { $SyncHash.SharedMaterials[$i] = $SyncHash.SharedMaterials[$i-1] }
                $SyncHash.SharedMaterials[0] = $lastMaterial

                for ($i = 0; $i -lt $panelGroupsByRing.Count; $i++) {
                    $materialForThisRing = $SyncHash.SharedMaterials[$i % $SyncHash.SharedMaterials.Count]
                    foreach ($panelModel in $panelGroupsByRing[$i].Group) { $panelModel.Value.Material = $materialForThisRing; $panelModel.Value.BackMaterial = $materialForThisRing.Clone() }
                }
            })
            $flowTimer.Start()
        }
    })

    # --- Animation Loop Setup for specific styles ---
    $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $SyncHash.StartTime = $SyncHash.LastFrameTime
    if (($SyncHash.VisualizationStyle -eq "FloatingSpheres") -or ($SyncHash.VisualizationStyle -eq "FloatingStars")) {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ Animate-FloatingObjects -SyncHash $SyncHash -Viewport $viewport })
    }
    elseif ($SyncHash.VisualizationStyle -eq "ButterflyEffect") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ 
            Animate-FloatingObjects -SyncHash $SyncHash -Viewport $viewport
            Animate-ButterflyEffect -SyncHash $SyncHash
        })
    }
    elseif ($SyncHash.VisualizationStyle -eq "RollerCoaster") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ 
            Animate-RollerCoaster -SyncHash $SyncHash
        })
    }
    elseif ($SyncHash.VisualizationStyle -eq "SphereVortex") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ Animate-SphereVortex -SyncHash $SyncHash })
    }
    elseif ($SyncHash.VisualizationStyle -eq "CurvedVortex") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ Animate-CurvedVortex -SyncHash $SyncHash })
    }
    elseif ($SyncHash.VisualizationStyle -eq "Aquarium") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ Animate-Aquarium -SyncHash $SyncHash -Viewport $viewport })
        if ($SyncHash.AddBubbles) {
        $bubbleCanvas = $window.FindName('bubbleCanvas')
        $rand = [Random]::new()
        $bubbles = New-Object System.Collections.Generic.List[object]
        $maxCount = 220

        $newBubbleBrushFunc = {
            param([int]$alpha)
            $c1 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Min($alpha,255), 255, 255, 255)
            $c2 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Max($alpha-120,30), 173, 216, 230)
            $brush = New-Object Windows.Media.RadialGradientBrush
            $brush.RadiusX = 0.6; $brush.RadiusY = 0.6
            $brush.GradientOrigin = [Windows.Point]::new(0.35,0.35)
            $brush.Center = [Windows.Point]::new(0.5,0.5)
            $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c1,0.0))
            $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c2,1.0))
            return $brush
        }

        $newBubbleFunc = {
            $w = $bubbleCanvas.ActualWidth
            $h = $bubbleCanvas.ActualHeight
            if ($w -le 0 -or $h -le 0) { return }

            $size   = $rand.Next(5, 25)
            $speed  = ($rand.NextDouble() * 1.4 + 0.6)
            $drift  = (($rand.NextDouble()*2.0) - 1.0) * 0.35
            $alpha  = $rand.Next(120, 230)

            $startX = $rand.NextDouble() * [Math]::Max($w - $size, 1)
            $startY = $h - ($rand.NextDouble() * ([Math]::Max($h*0.15, 80)))

            $ellipse = New-Object Windows.Shapes.Ellipse
            $ellipse.Width  = $size; $ellipse.Height = $size
            $ellipse.Fill   = & $newBubbleBrushFunc -alpha $alpha
            $ellipse.Stroke = [System.Windows.Media.Brushes]::White
            $ellipse.StrokeThickness = [Math]::Max($size * 0.02, 0.6)

            [Windows.Controls.Canvas]::SetLeft($ellipse, $startX); [Windows.Controls.Canvas]::SetTop($ellipse,  $startY)
            $bubbleCanvas.Children.Add($ellipse) | Out-Null

            $bubble = [pscustomobject]@{ Shape = $ellipse; Vy = $speed; Vx = $drift; Spin = ($rand.NextDouble() * 0.04) - 0.02; T = $rand.NextDouble() * [Math]::PI }
            $bubbles.Add($bubble) | Out-Null
        }

        $spawnTimer = New-Object Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(220) }
        $spawnTimer.Add_Tick({ if ($bubbles.Count -lt $maxCount) { 1..($rand.Next(1,4)) | ForEach-Object { & $newBubbleFunc } } })
        $SyncHash.BubbleSpawnTimer = $spawnTimer

        $animTimer = New-Object Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(16) }
        $animTimer.Add_Tick({
            if ($bubbleCanvas -eq $null -or $SyncHash.Paused) { return }
            $h = $bubbleCanvas.ActualHeight; $w = $bubbleCanvas.ActualWidth
            for ($i = $bubbles.Count - 1; $i -ge 0; $i--) {
                $b = $bubbles[$i]; $s = [double]$b.Shape.Width
                $x = [Windows.Controls.Canvas]::GetLeft($b.Shape); $y = [Windows.Controls.Canvas]::GetTop($b.Shape)
                $b.T += $b.Spin; $x += $b.Vx + ([Math]::Sin($b.T) * 0.15); $y -= $b.Vy
                if ($x -lt -10) { $x = -10; $b.Vx = [Math]::Abs($b.Vx) }
                elseif ($x + $s -gt $w + 10) { $x = $w + 10 - $s; $b.Vx = -[Math]::Abs($b.Vx) }
                [Windows.Controls.Canvas]::SetLeft($b.Shape, $x); [Windows.Controls.Canvas]::SetTop($b.Shape,  $y)
                if ($y + $s -lt 0) { $bubbleCanvas.Children.Remove($b.Shape); $bubbles.RemoveAt($i) }
            }
        })
        $SyncHash.BubbleAnimTimer = $animTimer
        $spawnTimer.Start(); $animTimer.Start()
    }
    }
    elseif ($SyncHash.VisualizationStyle -eq "ReindeerSleigh") {
        $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering({ Animate-ReindeerSleigh -SyncHash $SyncHash })
    }


    # --- Watchdog Timer for Stuck Media ---
    # This timer proactively checks for players that have not progressed as expected,
    # bypassing the unreliable MediaFailed event for certain visualizations.
    $watchdogTimer = New-Object System.Windows.Threading.DispatcherTimer
    $watchdogTimer.Interval = [TimeSpan]::FromSeconds(3)
    $watchdogTimer.Add_Tick({
        if ($SyncHash.Paused -or ($SyncHash.VisualizationStyle -ne 'FacetedSphereMulti' -and $SyncHash.VisualizationStyle -ne 'FloatingStars')) { return }

        $now = [datetime]::UtcNow
        foreach ($playerKey in $SyncHash.PlayerStates.Keys) {
            if (-not ($playerKey.StartsWith("Facet") -or $playerKey.StartsWith("Star"))) { continue }

            $playerState = $SyncHash.PlayerStates[$playerKey]
            if ($playerState.IsFailed) { continue }

            # Check 0: Media was assigned but never opened.
            if ($playerState.SourceAssignmentTime) {
                if (($now - $playerState.SourceAssignmentTime).TotalSeconds -gt 5) { # 5-second timeout to open
                    Handle-MediaFailure -PlayerKey $playerKey -Reason "Watchdog: Media failed to open."
                    continue # Move to the next player
                }
            }

            # Check 1: Video player is stuck (playing past its duration)
            if ($playerState.ExpectedDuration -and $playerState.PlaybackStopwatch.IsRunning) {
                $gracePeriod = [TimeSpan]::FromSeconds(3)
                if ($playerState.PlaybackStopwatch.Elapsed -gt ($playerState.ExpectedDuration + $gracePeriod)) {
                    Handle-MediaFailure -PlayerKey $playerKey -Reason "Watchdog: Video playback is stuck."
                    continue # Move to the next player
                }
            }
        }
    })
    $watchdogTimer.Start()
    $SyncHash.WatchdogTimer = $watchdogTimer # Store for cleanup
    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) { break }
}
