<#
.SYNOPSIS
    A master 3D media visualizer combining multiple display styles and rendering engines.
.DESCRIPTION
    This script is the ultimate all-in-one visualizer. It launches a comprehensive GUI to select
    image and video files, and then allows the user to choose from a variety of 3D visualization styles.

    It intelligently detects if FFmpeg is installed.
    - If FFmpeg is found, the user can choose between the high-compatibility FFmpeg engine or the legacy MediaElement engine.
    - If FFmpeg is not found, it defaults to the MediaElement engine.

    This script centralizes all common functionality, such as the UI, media handling, and error
    reporting, into a single, manageable file, providing maximum flexibility.

.EXAMPLE
    PS C:\> .\Show-MasterVisualizer.ps1

    Launches the media and style selection GUI. After selecting files, an engine (if available),
    and a display style, clicking "Play" will launch the chosen 3D visualization.
.NOTES
    Name:           Show-MasterVisualizer.ps1
    Version:        1.0.0, 12/31/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. FFmpeg/ffprobe are optional but recommended for best video support.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- Script Metadata (for Script Launcher) ---
$ExternalButtonName = "Master 3D Visualizer"
$ScriptDescription = "A single script combining all 3D visualizers with a choice of FFmpeg or MediaElement engine."
$RequiredExecutables = @() # Optional, so we don't list them as required for the launcher

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

    # --- Dependency Check for FFmpeg (Non-blocking) ---
    $ffmpegAvailable = $true
    $ffmpegExecutables = @("ffmpeg.exe", "ffprobe.exe")
    foreach ($exe in $ffmpegExecutables) {
        if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) {
            $ffmpegAvailable = $false
            break
        }
    }

    #region --- File and Style Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Master Visualizer - Media & Style Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 870)
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

    # --- Engine Selection ---
    $EngineGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Rendering Engine"; Location = '10, 305'; Size = '760, 55'; Anchor = 'Top, Left, Right' }
    $FfmpegRadioButton = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "FFmpeg (Recommended for best video support)"; Location = '15, 20'; AutoSize = $true; Checked = $true }
    $MediaElementRadioButton = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "MediaElement (Legacy, limited video support)"; Location = '350, 20'; AutoSize = $true }
    $FfmpegNotFoundLabel = New-Object System.Windows.Forms.Label -Property @{ 
        Text = "FFmpeg not found. Defaulting to MediaElement engine. Install FFmpeg for best results."; 
        Location = (New-Object System.Drawing.Point(15, 25)); 
        AutoSize = $true; 
        Font = (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)); 
        ForeColor = [System.Drawing.Color]::Gray 
    }
    $EngineGroupBox.Controls.AddRange(@($FfmpegRadioButton, $MediaElementRadioButton, $FfmpegNotFoundLabel))
    $SelectForm.Controls.Add($EngineGroupBox)

    # --- Display Style Selection ---
    $StyleGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Display Style"; Location = '10, 370'; Size = '760, 195'; Anchor = 'Top, Left, Right' }
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
    $TextGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 575'; Size = '760, 250'; Anchor = 'Top, Left, Right' }
    $SelectForm.Controls.Add($TextGroupBox)

    $TextOptionsGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Options"; Location = '10, 20'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $TextOptionsGroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $TextGroupBox.Controls.Add($TextOptionsGroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '140, 20'; Size = '455, 210'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center' }
    $TextGroupBox.Controls.Add($TextBox)

    $FilenamePreviewLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Filename.mp4"; Location = '140, 20'; Size = '455, 210'; Visible = $false; TextAlign = 'MiddleCenter' }
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
    if ($ffmpegAvailable) {
        $FfmpegNotFoundLabel.Visible = $false
        $FfmpegRadioButton.Visible = $true
        $MediaElementRadioButton.Visible = $true
    } else {
        $FfmpegNotFoundLabel.Visible = $true
        $FfmpegRadioButton.Visible = $false
        $MediaElementRadioButton.Visible = $false
    }

    $formState = @{ TextColor = [System.Drawing.Color]::Black; FontFamily = "Arial" }
    
    $textOverlayEvent = {
        $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked; $isCustomText = $RadioButton3.Checked; $isFilename = $RadioButton2.Checked
        $TextBox.Visible = $isCustomText; $FontControlsGroupBox.Visible = $isTextVisible; $FilenamePreviewLabel.Visible = $isFilename
    }
    $RadioButton1.Add_Click($textOverlayEvent); $RadioButton2.Add_Click($textOverlayEvent); $RadioButton3.Add_Click($textOverlayEvent)

    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.TextColor = $colorDialog.Color; $ColorExample.BackColor = $formState.TextColor
            $TextBox.ForeColor = $formState.TextColor; $FilenamePreviewLabel.ForeColor = $formState.TextColor
        }
    })

    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont; $FilenamePreviewLabel.Font = $newFont
        } catch {
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style); $FilenamePreviewLabel.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }

    $FontButton.Add_Click({
        $fontDialog = New-Object System.Windows.Forms.FontDialog
        try { $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value) } catch { $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12) }
        if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.FontFamily = $fontDialog.Font.Name; $FontButton.Text = $formState.FontFamily
            $NumericUpDown.Value = [decimal]$fontDialog.Font.Size; $BoldCheckbox.Checked = $fontDialog.Font.Bold; $ItalicCheckbox.Checked = $fontDialog.Font.Italic
            & $updateTextBoxFont
        }
    })
    
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont); $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont); $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)
    & $updateTextBoxFont

    $SelectAllCheckbox.Add_CheckedChanged({
        $isChecked = $SelectAllCheckbox.Checked
        foreach ($row in $DataGridView.Rows) { $row.Cells["Select"].Value = $isChecked }
        $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    })

    $BrowseButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $initialPath = 'C:\Users\jdalb\Videos\Test'
        $FolderBrowser.SelectedPath = $initialPath
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SelectedPath = $FolderBrowser.SelectedPath; $FolderPathTextBox.Text = $SelectedPath; $DataGridView.Rows.Clear()
            $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
            $VideoExtensions = "*.mp4", "*.m4v", "*.wmv", "*.avi", "*.mpg", "*.mpeg"
            $AllowedExtensions = $ImageExtensions + $VideoExtensions
            $gciParams = @{ File = $true; Include = $AllowedExtensions }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true } else { $gciParams.Path = Join-Path $SelectedPath "*" }
            Get-ChildItem @gciParams | ForEach-Object { $DataGridView.Rows.Add($false, $_.Name, $_.FullName) | Out-Null }
            $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $filePath = $DataGridView.Rows[$e.RowIndex].Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) { try { Start-Process $filePath } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", "OK", "Error") } }
        }
    })

    $PlayButton.Add_Click({
        $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
        if ($selectedFiles.Count -gt 0) {
            $formState.SelectedFiles = [System.Collections.ArrayList]::new($selectedFiles)
            $formState.UseTransparentEffect = $TransparentCheckbox.Checked
            $formState.AddBubbles = $AddBubblesCheckbox.Checked; $formState.AddWater = $AddWaterCheckbox.Checked
            $formState.NightSky = $NightSkyCheckbox.Checked; $formState.TwinklingStars = $TwinkleCheckbox.Checked
            $selectedStyleRB = $StyleGroupBox.Controls | Where-Object { $_.Checked }; $formState.VisualizationStyle = if ($selectedStyleRB) { $selectedStyleRB.Tag } else { "RotatingCube" }
            if ($ffmpegAvailable) { $formState.UseFfmpeg = $FfmpegRadioButton.Checked } else { $formState.UseFfmpeg = $false }
            switch ($true) {
                { $RadioButton1.Checked } { $formState.RbSelection = "Hidden"; break }
                { $RadioButton2.Checked } { $formState.RbSelection = "Filename"; break }
                { $RadioButton3.Checked } { $formState.RbSelection = "Custom"; break }
            }
            $formState.CustomText = $TextBox.Text; $formState.FontSize = $NumericUpDown.Value
            $formState.IsBold = $BoldCheckbox.Checked; $formState.IsItalic = $ItalicCheckbox.Checked
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
        UseFfmpeg            = $formState.UseFfmpeg
        AddBubbles           = $formState.AddBubbles; AddWater = $formState.AddWater
        NightSky             = $formState.NightSky; TwinklingStars = $formState.TwinklingStars
        VisualizationStyle   = $formState.VisualizationStyle
        CurrentIndex         = -1
        BadMediaFiles        = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        PlayerStates         = [hashtable]::Synchronized(@{})
        Paused               = $false; ControlsHidden = $false; RedoClicked = $false
        LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
        SpeedMultiplier      = 1.0
        RbSelection          = $formState.RbSelection; CustomText = $formState.CustomText
        TextColor            = $formState.TextColor; FontSize = $formState.FontSize
        FontFamily           = $formState.FontFamily; IsBold = $formState.IsBold; IsItalic = $formState.IsItalic
    })

    function Get-VideoDimensions {
        param([string]$FilePath)
        try {
            $ffprobePath = (Get-Command ffprobe.exe).Source
            $arguments = "-v error -select_streams v:0 -show_entries stream=width,height -of csv=s=x:p=0 `"$FilePath`""
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo; $startInfo.FileName = $ffprobePath; $startInfo.Arguments = $arguments
            $startInfo.UseShellExecute = $false; $startInfo.RedirectStandardOutput = $true; $startInfo.CreateNoWindow = $true
            $process = [System.Diagnostics.Process]::Start($startInfo); $output = $process.StandardOutput.ReadToEnd(); $process.WaitForExit()
            if ($process.ExitCode -eq 0 -and $output -match "(\d+)x(\d+)") { return @{ Width = [int]$matches[1]; Height = [int]$matches[2] } }
        } catch { Write-Warning "ffprobe failed: $($_.Exception.Message)" }
        return $null
    }

    $globalIndexLock = New-Object object
    function Get-NextMediaIndex {
        [System.Threading.Monitor]::Enter($globalIndexLock)
        try {
            $fileCount = $SyncHash.SelectedFiles.Count
            if ($fileCount -eq 0) { return -1 }
            if ($SyncHash.BadMediaFiles.Count -ge $fileCount) { Write-Warning "All available media files have failed."; return -1 }
            $startIndex = ($SyncHash.CurrentIndex + 1) % $fileCount; $currentIndex = $startIndex
            do {
                if (-not $SyncHash.BadMediaFiles.Contains($SyncHash.SelectedFiles[$currentIndex])) { $SyncHash.CurrentIndex = $currentIndex; return $currentIndex }
                $currentIndex = ($currentIndex + 1) % $fileCount
            } while ($currentIndex -ne $startIndex)
            return -1
        } finally { [System.Threading.Monitor]::Exit($globalIndexLock) }
    }

    function Handle-MediaFailure {
        param([string]$PlayerKey, [string]$Reason)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $playerState -or $playerState.IsFailed) { return }
        $playerState.IsFailed = $true

        $fileName = if ($playerState.CurrentFilePath) { [System.IO.Path]::GetFileName($playerState.CurrentFilePath) } elseif ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for player '$PlayerKey' (File: '$fileName'). Reason: $Reason. Attempting to replace."

        $filePath = if ($playerState.CurrentFilePath) { $playerState.CurrentFilePath } else { $playerState.CurrentSource.LocalPath }
        if ($filePath -and -not $SyncHash.BadMediaFiles.Contains($filePath)) { $SyncHash.BadMediaFiles.Add($filePath) | Out-Null }

        $SyncHash.Window.Dispatcher.Invoke([action]{
            if (-not $SyncHash.UseFfmpeg) {
                if ($playerState.PlaybackStopwatch) { $playerState.PlaybackStopwatch.Stop() }
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.CurrentMediaElement) {
                    if ($playerState.MediaEndedHandler) { try { $playerState.CurrentMediaElement.remove_MediaEnded($playerState.MediaEndedHandler) } catch {} }
                    if ($playerState.MediaOpenedHandler) { try { $playerState.CurrentMediaElement.remove_MediaOpened($playerState.MediaOpenedHandler) } catch {} }
                    if ($playerState.MediaFailedHandler) { try { $playerState.CurrentMediaElement.remove_MediaFailed($playerState.MediaFailedHandler) } catch {} }
                    $playerState.CurrentMediaElement.Stop(); $playerState.CurrentMediaElement.Close(); $playerState.CurrentMediaElement = $null
                }
            }
            if ($playerState.ContentPresenter) { $playerState.ContentPresenter.Content = $null }
            if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
        })

        $recoveryScriptBlock = { Start-NextMedia -PlayerKey $PlayerKey }; $null = $SyncHash.Window.Dispatcher.InvokeAsync($recoveryScriptBlock.GetNewClosure())
    }

    function Handle-MediaEnded_ME {
        param([string]$PlayerKey)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if ($pState.IsFailed) { return }
        $pState.PlaybackStopwatch.Stop()
        if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500 -and -not $pState.IsImage) { Handle-MediaFailure -PlayerKey $PlayerKey -Reason "Playback ended instantly (bad codec)."; return }
        Start-NextMedia -PlayerKey $PlayerKey
    }

    function Handle-MediaOpened_ME {
        param([string]$PlayerKey, $EventArgs)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $pState) { return }
        $pState.IsFailed = $false; $pState.PlaybackStopwatch.Restart()
        if ($pState.MediaHostGrid) { $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent }
        if ($EventArgs.NaturalDuration.HasTimeSpan) { $pState.ExpectedDuration = $EventArgs.NaturalDuration.TimeSpan } else { $pState.ExpectedDuration = $null }
        $pState.SourceAssignmentTime = $null
        if (-not $pState.IsImage -and -not $EventArgs.NaturalDuration.HasTimeSpan) { Handle-MediaFailure -PlayerKey $PlayerKey -Reason "Invalid duration or codec." }
    }

    function Start-NextMedia {
        param([string]$PlayerKey)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $playerState) { Write-Warning "Attempted to start media for a non-existent player key: '$PlayerKey'."; return }

    # Robustly add required properties for MediaElement if they don't exist
    if (-not $SyncHash.UseFfmpeg -and -not $playerState.PSObject.Properties['PlaybackStopwatch']) {
        $playerState | Add-Member -MemberType NoteProperty -Name 'PlaybackStopwatch' -Value (New-Object System.Diagnostics.Stopwatch)
        $playerState | Add-Member -MemberType NoteProperty -Name 'CurrentSource' -Value $null
        $playerState | Add-Member -MemberType NoteProperty -Name 'ExpectedDuration' -Value $null
        $playerState | Add-Member -MemberType NoteProperty -Name 'SourceAssignmentTime' -Value $null
        $playerState | Add-Member -MemberType NoteProperty -Name 'ImageCycleStartTime' -Value $null
    }

        # --- Cleanup Previous Player State ---
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($SyncHash.UseFfmpeg) {
            if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
            if ($playerState.FrameReader) { $playerState.FrameReader.Dispose() }
            $playerState.FfmpegProcess = $null; $playerState.FrameReader = $null; $playerState.WriteableBitmap = $null
        } else {
            if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
        }

        # --- Get Next Media ---
        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }
        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.IsFailed = $false
        if ($SyncHash.RbSelection -eq "Filename" -and $playerState.OverlayTextBlock) { $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath) }
        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $playerState.IsImage = $ImageExtensions -contains [System.IO.Path]::GetExtension($filePath).ToLower()

        try {
            if ($SyncHash.UseFfmpeg) {
                # --- FFmpeg Engine Logic ---
                $playerState.CurrentFilePath = $filePath
                if ($playerState.IsImage) {
                    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage; $bitmap.BeginInit(); $bitmap.UriSource = [Uri]$filePath; $bitmap.EndInit(); $bitmap.Freeze()
                    $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                    $playerState.ContentPresenter.Content = $image
                    if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent }
                    if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal" -or $SyncHash.VisualizationStyle -eq "ScrollingVertical") {
                        if ($bitmap.PixelHeight -gt 0) {
                            $aspectRatio = $bitmap.PixelWidth / $bitmap.PixelHeight
                            if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal") { $playerState.MediaHostGrid.Width = $playerState.MediaHostGrid.Height * $aspectRatio }
                            else { $playerState.MediaHostGrid.Height = $playerState.MediaHostGrid.Width / $aspectRatio }
                        }
                    }
                    $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $PlayerKey }
                    $timer.Add_Tick({ $t = $args[0]; $key = $t.Tag; $t.Stop(); Start-NextMedia -PlayerKey $key })
                    $playerState.MediaTimer = $timer; $timer.Start()
                } else { # Video
                    $dimensions = Get-VideoDimensions -FilePath $filePath
                    if (-not $dimensions) { throw "Could not get video dimensions from ffprobe." }
                    $w = $dimensions.Width; $h = $dimensions.Height
                    if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal" -or $SyncHash.VisualizationStyle -eq "ScrollingVertical") {
                        if ($h -gt 0) {
                            $aspectRatio = $w / $h
                            if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal") { $playerState.MediaHostGrid.Width = $playerState.MediaHostGrid.Height * $aspectRatio }
                            else { $playerState.MediaHostGrid.Height = $playerState.MediaHostGrid.Width / $aspectRatio }
                        }
                    }
                    $wb = New-Object System.Windows.Media.Imaging.WriteableBitmap($w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Bgra32, $null)
                    $image = New-Object System.Windows.Controls.Image -Property @{ Source = $wb; Stretch = 'Fill' }
                    $playerState.ContentPresenter.Content = $image
                    if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent }
                    $playerState.WriteableBitmap = $wb; $playerState.FrameWidth = $w; $playerState.FrameHeight = $h
                    $playerState.FrameStride = $w * 4; $playerState.FrameBuffer = New-Object byte[] ($playerState.FrameStride * $h)

                    # Add looping for visualizations that depend on continuous media playback for their effect.
                    $loopArg = ""
                    if ($SyncHash.VisualizationStyle -in @("MediaFlowFunnel", "SphereVortex")) {
                        $loopArg = "-stream_loop -1"
                    }
                    $ffmpegPath = (Get-Command ffmpeg.exe).Source
                    $arguments = "$loopArg -i `"$filePath`" -f rawvideo -pix_fmt bgra -v quiet -"
                    $startInfo = New-Object System.Diagnostics.ProcessStartInfo; $startInfo.FileName = $ffmpegPath; $startInfo.Arguments = $arguments
                    $startInfo.UseShellExecute = $false; $startInfo.RedirectStandardOutput = $true; $startInfo.CreateNoWindow = $true
                    $process = New-Object System.Diagnostics.Process; $process.StartInfo = $startInfo; $process.EnableRaisingEvents = $true; $process.Start() | Out-Null
                    $playerState.FfmpegProcess = $process; $playerState.FrameReader = New-Object System.IO.BinaryReader($process.StandardOutput.BaseStream)
                }
            } else {
                # --- MediaElement Engine Logic ---
                $playerState.CurrentSource = [Uri]$filePath
                $playerState.SourceAssignmentTime = [datetime]::UtcNow
                if ($playerState.IsImage) {
                    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage; $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit(); $bitmap.Freeze()
                    $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                    $playerState.ContentPresenter.Content = $image
                    if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal" -or $SyncHash.VisualizationStyle -eq "ScrollingVertical") {
                        if ($bitmap.PixelHeight -gt 0) {
                            $aspectRatio = $bitmap.PixelWidth / $bitmap.PixelHeight
                            if ($SyncHash.VisualizationStyle -eq "ScrollingHorizontal") { $playerState.MediaHostGrid.Width = $playerState.MediaHostGrid.Height * $aspectRatio }
                            else { $playerState.MediaHostGrid.Height = $playerState.MediaHostGrid.Width / $aspectRatio }
                        }
                    }
                    $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $PlayerKey }
                    $timer.Add_Tick({ $t = $args[0]; $key = $t.Tag; $t.Stop(); Handle-MediaEnded_ME -PlayerKey $key })
                    $playerState.MediaTimer = $timer; $playerState.ImageCycleStartTime = [datetime]::UtcNow; $timer.Start()
                    $playerState.SourceAssignmentTime = $null
                } else { # Video
                    $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{ LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource; Tag = $PlayerKey }
                    $mediaElement.Add_MediaEnded($playerState.MediaEndedHandler); $mediaElement.Add_MediaOpened($playerState.MediaOpenedHandler); $mediaElement.Add_MediaFailed($playerState.MediaFailedHandler)
                    $playerState.ContentPresenter.Content = $mediaElement; $playerState.CurrentMediaElement = $mediaElement; $mediaElement.Play()
                    if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
                }
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
                if ($SyncHash.UseFfmpeg) {
                    if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
                    if ($playerState.FrameReader) { $playerState.FrameReader.Dispose() }
                } else {
                    if ($playerState.CurrentMediaElement) {
                        if ($playerState.MediaEndedHandler) { try { $playerState.CurrentMediaElement.remove_MediaEnded($playerState.MediaEndedHandler) } catch {} }
                        if ($playerState.MediaOpenedHandler) { try { $playerState.CurrentMediaElement.remove_MediaOpened($playerState.MediaOpenedHandler) } catch {} }
                        if ($playerState.MediaFailedHandler) { try { $playerState.CurrentMediaElement.remove_MediaFailed($playerState.MediaFailedHandler) } catch {} }
                        $playerState.CurrentMediaElement.Close()
                    }
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
        param([double]$startRadius=0, [double]$endRadius=0, [double]$height=0, [double]$arcAngle=0, [double]$twistAngle=0, [int]$stacks=50, [int]$slices=2, [System.Windows.Media.Media3D.Point3D]$p1, [System.Windows.Media.Media3D.Point3D]$p2, [System.Windows.Media.Media3D.Point3D]$p3, [System.Windows.Media.Media3D.Point3D]$p4)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $arcAngleRad = $arcAngle * [Math]::PI / 180.0; $twistAngleRad = $twistAngle * [Math]::PI / 180.0
        for ($j = 0; $j -le $stacks; $j++) {
            $v = $j / $stacks; $v_eased = $v * $v; $currentRadius = $startRadius - $v * ($startRadius - $endRadius); $currentY = ($height / 2) - $v_eased * $height; $currentTwist = $v * $twistAngleRad
            for ($i = 0; $i -le $slices; $i++) {
                $u = $i / $slices; $theta = $currentTwist + ($u * $arcAngleRad) - ($arcAngleRad / 2.0); $x = $currentRadius * [Math]::Cos($theta); $z = $currentRadius * [Math]::Sin($theta)
                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $currentY, $z)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u, $v))
            }
        }
        if ($p1) {
            $mesh.Positions.Clear(); $mesh.TextureCoordinates.Clear()
            $mesh.Positions.Add($p1); $mesh.Positions.Add($p2); $mesh.Positions.Add($p3); $mesh.Positions.Add($p4)
            $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,1)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,1))
            $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(1); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(3)
        } else {
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
        param([double]$width=2.0, [double]$height=1.0, [double]$curveDepth=0.5, [int]$widthSegments=20, [int]$heightSegments=2)
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        for ($j = 0; $j -le $heightSegments; $j++) {
            $y = $height / 2 - ($j / $heightSegments) * $height
            for ($i = 0; $i -le $widthSegments; $i++) {
                $x = -$width / 2 + ($i / $widthSegments) * $width; $normalizedX = $x / ($width / 2); $z = $curveDepth * ($normalizedX * $normalizedX)
                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $widthSegments, $j / $heightSegments))
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
        for ($i = 0; $i -le $segments; $i++) {
            $pointData = $PathData[$i]; $p1 = $pointData.Point; $up = $pointData.Up; $normal = $pointData.Normal; $offsetVector = $up * $verticalOffset
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
        $tiesMesh = New-PathRibbonMesh -PathData $PathData -segments $segments -width 0.8 -verticalOffset (-0.05 / 2)
        $tiesMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::SaddleBrown)
        $tiesModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $tiesMesh; Material = $tiesMaterial; BackMaterial = $tiesMaterial }
        $trackModelGroup.Children.Add($tiesModel)
        return $trackModelGroup
    }

    function New-PieSliceModel {
        param([System.Windows.Point]$center, [double]$radius, [double]$startAngleDeg, [double]$sliceAngleDeg, [double]$thickness = 0.1)
        $endAngleDeg = $startAngleDeg + $sliceAngleDeg; $startAngleRad = $startAngleDeg * [Math]::PI / 180.0; $endAngleRad = $endAngleDeg * [Math]::PI / 180.0; $halfThick = $thickness / 2.0
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $arcPoints = New-Object 'System.Collections.Generic.List[System.Windows.Point]'; $arcPoints.Add($center) | Out-Null
        $numArcSegments = [double][Math]::Max(1, [Math]::Ceiling($sliceAngleDeg / 2.0))
        for ($i = 0; $i -le $numArcSegments; $i++) {
            $angle = $startAngleRad + ($i / $numArcSegments) * ($endAngleRad - $startAngleRad); [double]$pointX = $center.X + $radius * [Math]::Cos($angle); [double]$pointY = $center.Y + $radius * [Math]::Sin($angle)
            $arcPoints.Add((New-Object System.Windows.Point($pointX, $pointY))) | Out-Null
        }
        $frontBaseIndex = $mesh.Positions.Count
        foreach ($p in $arcPoints) { $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, $halfThick)) | Out-Null; $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5)) | Out-Null }
        $backBaseIndex = $mesh.Positions.Count
        foreach ($p in $arcPoints) { $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, -$halfThick)) | Out-Null; $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5)) | Out-Null }
        for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) { $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $i + 1); $mesh.TriangleIndices.Add($frontBaseIndex + $i) }
        for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) { $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $i); $mesh.TriangleIndices.Add($backBaseIndex + $i + 1) }
        for ($i = 1; $i -lt $arcPoints.Count; $i++) {
            $p1_front = $frontBaseIndex + $i; $p2_front = $frontBaseIndex + $i + 1; $p1_back = $backBaseIndex + $i;  $p2_back = $backBaseIndex + $i + 1
            $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p1_back); $mesh.TriangleIndices.Add($p2_back); $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p2_back); $mesh.TriangleIndices.Add($p2_front)
        }
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1); $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1); $mesh.TriangleIndices.Add($frontBaseIndex + 1)
        $lastIdx = $arcPoints.Count -1; $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex)
        return (New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $mesh })
    }

    # --- Visualization-Specific Setup Functions ---
    function New-FacetedSphereModels {
        param([double]$radius = 1.5, [int]$slices = 8, [int]$stacks = 4)
        $facets = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.GeometryModel3D]
        $allVertices = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.Point3D]
        for ($stack = 0; $stack -le $stacks; $stack++) {
            $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks; $y = $radius * [Math]::Sin($phi); $r = $radius * [Math]::Cos($phi)
            for ($slice = 0; $slice -le $slices; $slice++) {
                $theta = $slice * 2 * [Math]::PI / $slices; $x = $r * [Math]::Cos($theta); $z = $r * [Math]::Sin($theta)
                $allVertices.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
            }
        }
        for ($stack = 0; $stack -lt $stacks; $stack++) {
            for ($slice = 0; $slice -lt $slices; $slice++) {
                $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
                $i0 = $stack * ($slices + 1) + $slice; $i1 = ($stack + 1) * ($slices + 1) + $slice; $i2 = $i0 + 1; $i3 = $i1 + 1
                $p0 = $allVertices[$i0]; $p1 = $allVertices[$i1]; $p2 = $allVertices[$i2]; $p3 = $allVertices[$i3]
                $mesh.Positions.Add($p0); $mesh.Positions.Add($p1); $mesh.Positions.Add($p2); $mesh.Positions.Add($p2); $mesh.Positions.Add($p1); $mesh.Positions.Add($p3)
                $uv0 = [System.Windows.Point]::new(0,0); $uv1 = [System.Windows.Point]::new(0,1); $uv2 = [System.Windows.Point]::new(1,0); $uv3 = [System.Windows.Point]::new(1,1)
                $mesh.TextureCoordinates.Add($uv0); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv2); $mesh.TextureCoordinates.Add($uv2); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv3)
                $facets.Add((New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $null)))
            }
        }
        return $facets
    }

    function Setup-FacetedSphereMulti {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Faceted Sphere (Multi-Media)"; $Viewport.Camera.Position = "0,0,8"
        $sphereContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($sphereContainer) | Out-Null
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $numberOfSharedPlayers = 8; $sharedMaterials = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt $numberOfSharedPlayers; $i++) {
            $playerKey = "SharedFacetPlayer$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $sharedMaterials.Add($material) | Out-Null; Start-NextMedia -PlayerKey $playerKey
        }
        $facetModels = New-FacetedSphereModels -radius 2.5 -slices 8 -stacks 4
        for ($i = 0; $i -lt $facetModels.Count; $i++) { $facetModels[$i].Material = $sharedMaterials | Get-Random; $sphereContainer.Content.Children.Add($facetModels[$i]) | Out-Null }
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_Faceted", $sphereContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_Faceted", $sphereContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Faceted"); $axisAngleY = $Window.FindName("AxisAngleY_Faceted")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-FloatingStars {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Stars"; $Viewport.Camera.Position = "0,0,15"
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()
        $sphereRadius = 0.5; $coneHeight = 0.75; $coneRadius = 0.2; $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 64 -stacks 32; $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 64
        $visuals = @{ "Middle"=@{"Mesh"=$sphereMesh; "Transform"=(New-Object System.Windows.Media.Media3D.TranslateTransform3D)}; "Top"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))}; "Bottom"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.Transform3DGroup)}; "Right"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.Transform3DGroup)}; "Left"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.Transform3DGroup)}; "Front"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.Transform3DGroup)}; "Back"=@{"Mesh"=$coneMesh; "Transform"=(New-Object System.Windows.Media.Media3D.Transform3DGroup)} }
        $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1))); $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
        $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), -90)))); $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius, 0, 0)))
        $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), 90)))); $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius, 0, 0)))
        $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 90)))); $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, $sphereRadius)))
        $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), -90)))); $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$sphereRadius)))
        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Star$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $sharedMaterial = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $sharedMaterial.Color = [System.Windows.Media.Colors]::White }
            $starContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $Viewport.Children.Add($starContainer) | Out-Null
            foreach ($partKey in $visuals.Keys) {
                $part = $visuals[$partKey]; $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D; $geometryModel.Geometry = $part.Mesh; $geometryModel.Transform = $part.Transform; $geometryModel.Material = $sharedMaterial
                $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $geometryModel }; $starContainer.Children.Add($modelVisual) | Out-Null
            }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D)); $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform); $starContainer.Transform = $transformGroup
            $starObject = [pscustomobject]@{ Visual = $starContainer; Translate = $translateTransform; Rotate = $rotateTransform; Velocity = { $v = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0)); if ($v.Length -gt 0) { $v.Normalize() }; return $v * (Get-Random -Minimum 1.5 -Maximum 3.0) }.Invoke(); RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0)); CurrentRotation = New-Object System.Windows.Media.Media3D.Quaternion(0,0,0,1) }
            $SyncHash.FloatingObjects.Add($starObject) | Out-Null; Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-RotatingCube {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Cube"; $Viewport.Camera.Position = "0,0,5"; $Viewport.Camera.FieldOfView = "70"
        $cubeContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($cubeContainer); $cubeModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup; $cubeContainer.Content = $cubeModelGroup
        $faceNames = @("Front", "Back", "Right", "Left", "Top", "Bottom")
        $meshes = @{ Front="-1,-1,1 1,-1,1 1,1,1 -1,1,1"; Back="-1,-1,-1 -1,1,-1 1,1,-1 1,-1,-1"; Right="1,-1,1 1,-1,-1 1,1,-1 1,1,1"; Left="-1,-1,-1 -1,-1,1 -1,1,1 -1,1,-1"; Top="-1,1,1 1,1,1 1,1,-1 -1,1,-1"; Bottom="-1,-1,-1 1,-1,-1 1,-1,1 -1,-1,1" }
        $texCoords = @{ Front="0,1 1,1 1,0 0,0"; Back="1,1 1,0 0,0 0,1"; Right="0,1 1,1 1,0 0,0"; Left="0,1 1,1 1,0 0,0"; Top="0,1 1,1 1,0 0,0"; Bottom="0,1 1,1 1,0 0,0" }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        foreach ($faceName in $faceNames) {
            $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $mesh.Positions = $meshes[$faceName]; $mesh.TextureCoordinates = $texCoords[$faceName]; $mesh.TriangleIndices = "0,1,2 0,2,3"
            $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $faceName }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $faceName -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $faceName -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$faceName] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $material); $faceModel.Transform = New-Object System.Windows.Media.Media3D.ScaleTransform3D(0.7, 0.7, 0.7); $cubeModelGroup.Children.Add($faceModel)
            Start-NextMedia -PlayerKey $faceName
        }
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_Cube", $cubeContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_Cube", $cubeContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Cube"); $axisAngleY = $Window.FindName("AxisAngleY_Cube")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Sphere {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Sphere"; $Viewport.Camera.Position = "0,0,6"; $sphereMesh = New-SphereMesh -radius 1.6 -slices 128 -stacks 64
        $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null
        if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
        else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey "Sphere" }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey "Sphere" -Reason $e.ErrorException.Message }.GetNewClosure() } }
        $SyncHash.PlayerStates["Sphere"] = $playerState
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material); $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }; $Viewport.Children.Add($modelVisual)
        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $axisAngleX = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "1,0,0" }; $axisAngleY = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "0,1,0" }
        $rotateX = New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleX); $rotateY = New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleY)
        $transformGroup.Children.Add($rotateX); $transformGroup.Children.Add($rotateY); $modelVisual.Transform = $transformGroup
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }; Start-NextMedia -PlayerKey "Sphere"
    }

    function Setup-PulsingStar {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Pulsing Star"; $Viewport.Camera.Position = "0,0,12"
        $starContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><Transform3DGroup><ScaleTransform3D x:Name="PulseScale" ScaleX="1" ScaleY="1" ScaleZ="1" /><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="StarRotation" Axis="1,1,0.5" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($starContainer)
        $sphereRadius = 1.2; $coneHeight = 2.4; $coneRadius = 0.8; $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64; $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        function New-MediaHost { $grid = New-Object System.Windows.Controls.Grid; $cp = New-Object System.Windows.Controls.ContentPresenter; $grid.Children.Add($cp) | Out-Null; $tb = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $grid.Children.Add($tb) | Out-Null; return @{ Grid = $grid; ContentPresenter = $cp; OverlayTextBlock = $tb } }
        $sphereHost = New-MediaHost; $VisualHost.Children.Add($sphereHost.Grid) | Out-Null; $sphereVisualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $sphereHost.Grid }; $sphereMaterial = New-Object $materialType -Property @{ Brush = $sphereVisualBrush }; if ($SyncHash.UseTransparentEffect) { $sphereMaterial.Color = [System.Windows.Media.Colors]::White }
        $sphereModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $sphereMaterial); $sphereModelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $sphereModel }; $starContainer.Children.Add($sphereModelVisual) | Out-Null
        if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $sphereHost.ContentPresenter; OverlayTextBlock = $sphereHost.OverlayTextBlock; MediaHostGrid = $sphereHost.Grid } }
        else { $playerState = @{ ContentPresenter = $sphereHost.ContentPresenter; OverlayTextBlock = $sphereHost.OverlayTextBlock; MediaHostGrid = $sphereHost.Grid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey "Sphere" }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey "Sphere" -Reason $e.ErrorException.Message }.GetNewClosure() } }
        $SyncHash.PlayerStates["Sphere"] = $playerState
        $conePositions = @( @{ Name="Top"; Transform=(New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)) }, @{ Name="Bottom"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Right"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Left"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Front"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }, @{ Name="Back"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) } )
        $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1))); $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
        $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1',-90))))); $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius,0,0)))
        $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1',90))))); $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius,0,0)))
        $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0',90))))); $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0,0,$sphereRadius)))
        $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0',-90))))); $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0,0,-$sphereRadius)))
        foreach ($coneInfo in $conePositions) {
            $coneHost = New-MediaHost; $VisualHost.Children.Add($coneHost.Grid) | Out-Null; $coneVisualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $coneHost.Grid }; $coneMaterial = New-Object $materialType -Property @{ Brush = $coneVisualBrush }; if ($SyncHash.UseTransparentEffect) { $coneMaterial.Color = [System.Windows.Media.Colors]::White }
            $coneModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($coneMesh, $coneMaterial); $coneModel.Transform = $coneInfo.Transform; $coneModelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $coneModel }; $starContainer.Children.Add($coneModelVisual) | Out-Null
            $playerKey = "Cone$($coneInfo.Name)"
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $coneHost.ContentPresenter; OverlayTextBlock = $coneHost.OverlayTextBlock; MediaHostGrid = $coneHost.Grid } }
            else { $playerState = @{ ContentPresenter = $coneHost.ContentPresenter; OverlayTextBlock = $coneHost.OverlayTextBlock; MediaHostGrid = $coneHost.Grid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
        }
        $starAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }; $pulseAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0.85, 1.15, [TimeSpan]::FromSeconds(2)) -Property @{ AutoReverse=$true; RepeatBehavior="Forever" }
        $Window.RegisterName("StarRotation_Pulsing", $starContainer.Transform.Children[1].Rotation); $Window.RegisterName("PulseScale_Pulsing", $starContainer.Transform.Children[0])
        $starRotation = $Window.FindName("StarRotation_Pulsing"); $pulseScale = $Window.FindName("PulseScale_Pulsing")
        $starRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $starAnim); $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $pulseAnim); $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $pulseAnim); $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $pulseAnim)
        $SyncHash.Animations = @{ Star = $starAnim; Pulse = $pulseAnim }; $SyncHash.Rotations = @{ Star = $starRotation }; $SyncHash.Transforms = @{ Pulse = $pulseScale }; $SyncHash.PlayerStates.Keys | ForEach-Object { Start-NextMedia -PlayerKey $_ }
    }

    function Setup-FacetedSphereSingle {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Faceted Sphere (Single Media)"; $Viewport.Camera.Position = "0,0,8"
        $sphereContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($sphereContainer) | Out-Null; $sphereMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $facetModels = New-FacetedSphereModels -radius 2.5 -slices 10 -stacks 5
        foreach ($facetModel in $facetModels) { $facetGeom = $facetModel.Geometry; $facetGeom.Positions | ForEach-Object { $sphereMesh.Positions.Add($_) }; $facetGeom.TextureCoordinates | ForEach-Object { $sphereMesh.TextureCoordinates.Add($_) } }
        $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null
        if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
        else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey "Sphere" }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey "Sphere" -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey "Sphere" -Reason $e.ErrorException.Message }.GetNewClosure() } }
        $SyncHash.PlayerStates["Sphere"] = $playerState
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material); $sphereContainer.Content.Children.Add($model) | Out-Null; Start-NextMedia -PlayerKey "Sphere"
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_FacetedSingle", $sphereContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_FacetedSingle", $sphereContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_FacetedSingle"); $axisAngleY = $Window.FindName("AxisAngleY_FacetedSingle")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-FloatingSpheres {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Spheres"; $Viewport.Camera.Position = "0,0,15"
        $sphereMesh = New-SphereMesh -radius 1.5 -slices 128 -stacks 64; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Sphere$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material); $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }; $Viewport.Children.Add($modelVisual) | Out-Null
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D)); $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform); $modelVisual.Transform = $transformGroup
            $starObject = [pscustomobject]@{ Visual = $modelVisual; Translate = $translateTransform; Rotate = $rotateTransform; Velocity = { $v = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0)); if ($v.Length -gt 0) { $v.Normalize() }; return $v * (Get-Random -Minimum 1.5 -Maximum 3.0) }.Invoke(); RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0)); CurrentRotation = New-Object System.Windows.Media.Media3D.Quaternion(0,0,0,1) }
            $SyncHash.FloatingObjects.Add($starObject) | Out-Null; Start-NextMedia -PlayerKey $playerKey
        }
    }

       function Setup-MediaFlowFunnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Flow Funnel"; $Viewport.Camera.Position = "0,2,18"; $Viewport.Camera.LookDirection = "0,-0.1,-1"
        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0"/></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0"/></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer) | Out-Null
        $funnelModelGroup = $funnelContainer.Content
        $numberOfGroups = 6
        $sharedMaterials = [System.Collections.ArrayList]::new()
        $sharedBackMaterials = [System.Collections.ArrayList]::new()
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
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            $sharedMaterials.Add($material) | Out-Null
            $sharedBackMaterials.Add($material.Clone()) | Out-Null

            if ($SyncHash.UseFfmpeg) {
                $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid }
            } else {
                $playerState = @{
                    ContentPresenter     = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid
                    PlaybackStopwatch    = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false
                    MediaEndedHandler    = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure()
                    MediaOpenedHandler   = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure()
                    MediaFailedHandler   = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure()
                }
            }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            Start-NextMedia -PlayerKey $playerKey
        }

        $SyncHash.PanelModels = @{}
        $numberOfRings = 5; $panelsPerRing = 12; $startRadius = 7.0; $endRadius = 1.0; $totalHeight = 8.0
        $radiusStep = ($startRadius - $endRadius) / $numberOfRings
        $angleStep = (2 * [Math]::PI) / $panelsPerRing
        
        for ($r = 0; $r -lt $numberOfRings; $r++) {
            $outerR = $startRadius - ($r * $radiusStep)
            $innerR = $startRadius - (($r + 1) * $radiusStep)
            $progressOuter = $r / $numberOfRings
            $progressOuter_eased = $progressOuter * $progressOuter
            $progressInner = ($r + 1) / $numberOfRings
            $progressInner_eased = $progressInner * $progressInner
            $yOuter = $totalHeight / 2 - ($progressOuter_eased * $totalHeight)
            $yInner = $totalHeight / 2 - ($progressInner_eased * $totalHeight)
            
            for ($p = 0; $p -lt $panelsPerRing; $p++) {
                $theta1 = $p * $angleStep
                $theta2 = ($p + 1) * $angleStep
                $p1 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta1), $yOuter, $outerR * [Math]::Sin($theta1))
                $p2 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta2), $yOuter, $outerR * [Math]::Sin($theta2))
                $p3 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta2), $yInner, $innerR * [Math]::Sin($theta2))
                $p4 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta1), $yInner, $innerR * [Math]::Sin($theta1))
                $panelMesh = New-SpiralingPanelMesh -p1 $p1 -p2 $p2 -p3 $p3 -p4 $p4
                
                $randIndex = Get-Random -Maximum $sharedMaterials.Count
                $materialToUse = $sharedMaterials[$randIndex]
                $backMaterialToUse = $sharedBackMaterials[$randIndex]
                
                $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{

                    Geometry = $panelMesh
                    Material = $materialToUse
                    BackMaterial = $backMaterialToUse
                }
                $funnelModelGroup.Children.Add($geometryModel) | Out-Null
                $SyncHash.PanelModels[($r * $panelsPerRing) + $p] = $geometryModel
            }

        }

        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(45)) -Property @{ RepeatBehavior="Forever" }
        $axisAngleY = $funnelContainer.Transform.Children[0].Rotation
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        
        $SyncHash.Animations = @{ Y = $animY }
        $SyncHash.Rotations = @{ Y = $axisAngleY }
        $SyncHash.SharedMaterials = $sharedMaterials
        $SyncHash.SharedBackMaterials = $sharedBackMaterials
        $SyncHash.PanelsPerRing = $panelsPerRing
    }

    function Setup-CurvedVortex {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Curved Vortex"; $Viewport.Camera.Position = "0,5,15"; $Viewport.Camera.LookDirection = "0,-0.3,-1"
        $vortexContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $Viewport.Children.Add($vortexContainer) | Out-Null
        $panelCount = 16; $curvedPanelMesh = New-CurvedPanelMesh -width 2.5 -height 1.5 -curveDepth 0.5; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $panelCount; $i++) {
            $playerKey = "Panel$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null; $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush; Color = [System.Windows.Media.Colors]::White }
            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $curvedPanelMesh; Material = $material; BackMaterial = $material.Clone() }; $panelContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $panelModel }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
            $transformGroup.Children.Add($rotateTransform) | Out-Null; $transformGroup.Children.Add($translateTransform) | Out-Null; $panelContainer.Transform = $transformGroup; $vortexContainer.Children.Add($panelContainer) | Out-Null
            $playerStateParams = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; TranslateTransform = $translateTransform; RotateTransform = $rotateTransform; CurrentAngle = (720.0 / $panelCount) * $i }
            if (-not $SyncHash.UseFfmpeg) { $playerStateParams.MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); $playerStateParams.MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); $playerStateParams.MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() }
            $SyncHash.PlayerStates[$playerKey] = $playerStateParams; Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-RotatingStar {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Rotating Star"; $Viewport.Camera.Position = "0,0,12"
        $starContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($starContainer) | Out-Null; $totalObjectHeight = 12 * 0.75; $sphereRadius = ($totalObjectHeight / 5.0) * 0.5; $coneHeight = ($sphereRadius * 1.5) * 2.0; $coneRadius = ($sphereRadius * 0.4) * 2.0
        $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64; $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $targets = @("Top", "Middle", "Bottom", "Left", "Right", "Front", "Back")
        foreach ($target in $targets) {
            $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $target }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $target -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $target -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$target] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $mesh = if ($target -eq "Middle") { $sphereMesh } else { $coneMesh }; $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $material); $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
            switch ($target) { "Top" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))) }; "Bottom" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0))) }; "Right" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', -90))))) }; "Left" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', 90))))) }; "Front" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', -90))))) }; "Back" { $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', 90))))) } }
            $model.Transform = $transformGroup; $starContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model })) | Out-Null; Start-NextMedia -PlayerKey $target
        }
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_RotStar", $starContainer.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_RotStar", $starContainer.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_RotStar"); $axisAngleY = $Window.FindName("AxisAngleY_RotStar")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-ButterflyEffect {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Butterfly Effect"; $Viewport.Camera.Position = "0,0,15"
        $planeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $planeMesh.Positions = "-0.667,-0.5,0 0.667,-0.5,0 0.667,0.5,0 -0.667,0.5,0"; $planeMesh.TriangleIndices = "0,1,2 0,2,3"; $planeMesh.TextureCoordinates = "0,1 1,1 1,0 0,0"; $planeMesh.Freeze()
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()
        for ($i = 1; $i -le 6; $i++) {
            $butterflyContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D)); $flutterTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = "1,0,0" }))
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $transformGroup.Children.Add($flutterTransform) | Out-Null; $transformGroup.Children.Add($rotateTransform) | Out-Null; $transformGroup.Children.Add($translateTransform) | Out-Null; $butterflyContainer.Transform = $transformGroup
            foreach ($face in @("Front", "Back")) {
                $playerKey = "${i}_${face}"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null
                if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
                else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
                $SyncHash.PlayerStates[$playerKey] = $playerState
                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($planeMesh, $material); if ($face -eq "Back") { $faceModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D("0,1,0", 180))) }
                $butterflyContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $faceModel })) | Out-Null; Start-NextMedia -PlayerKey $playerKey
            }
            $Viewport.Children.Add($butterflyContainer) | Out-Null; $startX = (Get-Random -Minimum -8 -Maximum 8); $startY = (Get-Random -Minimum -4 -Maximum 4); $translateTransform.OffsetX = $startX; $translateTransform.OffsetY = $startY
            $butterflyObject = [pscustomobject]@{ Visual = $butterflyContainer; Translate = $translateTransform; Rotate = $rotateTransform; FlutterTransform = $flutterTransform; Velocity = New-Object System.Windows.Media.Media3D.Vector3D(((Get-Random -Minimum 0.5 -Maximum 1.5) * (Get-Random @(1, -1))), ((Get-Random -Minimum 0.5 -Maximum 1.5) * (Get-Random @(1, -1))), 0); RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20 -Maximum 20), (Get-Random -Minimum -20 -Maximum 20), 0); CurrentRotation = [System.Windows.Media.Media3D.Quaternion]::new(0,0,0,1) }
            $SyncHash.FloatingObjects.Add($butterflyObject) | Out-Null
        }
    }

    function Setup-WagonWheel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Wagon Wheel"; $Viewport.Camera.Position = "0,0,8"
        $container = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($container); $camera = $Viewport.Camera; $visibleHeight = 2.0 * $camera.Position.Z * [Math]::Tan(($camera.FieldOfView * ([Math]::PI / 180.0)) / 2.0); $dynamicRadius = ($visibleHeight * 0.50) / 2.0
        $numberOfSlices = 8; $sliceAngle = 360.0 / $numberOfSlices; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $numberOfSlices; $i++) {
            $sliceParts = New-WagonWheelSliceModel -radius $dynamicRadius -startAngleDeg ($i * $sliceAngle) -sliceAngleDeg $sliceAngle
            $mediaHostGrid = New-Object System.Windows.Controls.Grid; $mediaHostGrid.Background = [System.Windows.Media.Brushes]::Black; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; Margin='10,0,10,0'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGrid.Children.Add($contentPresenter) | Out-Null; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null; $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            $visualBrush = New-Object System.Windows.Media.VisualBrush($mediaHostGrid); $material = New-Object $materialType($visualBrush); if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $sliceParts.OuterFaceModel.Material = $material; $container.Content.Children.Add($sliceParts.FullSliceModel) | Out-Null
            $playerKey = "Slice$i"
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; CurrentSource = $null; IsFailed = $false; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState; Start-NextMedia -PlayerKey $playerKey
        }
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)) -Property @{ RepeatBehavior="Forever" }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)) -Property @{ RepeatBehavior="Forever" }
        $Window.RegisterName("AxisAngleX_Wagon", $container.Transform.Children[0].Rotation); $Window.RegisterName("AxisAngleY_Wagon", $container.Transform.Children[1].Rotation)
        $axisAngleX = $Window.FindName("AxisAngleX_Wagon"); $axisAngleY = $Window.FindName("AxisAngleY_Wagon")
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Carousel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "3D Media Carousel"; $Viewport.Camera.Position = "0,0,12"
        $carouselContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="CarouselRotation" Axis="0,1,0" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($carouselContainer) | Out-Null; $panelCount = 8; $angleIncrement = 360 / $panelCount; $panelWidth = 3.0; $panelHeight = 5.0; $carouselRadius = 4.0
        $panelMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $panelMesh.Positions = "$(-$panelWidth/2),$(-$panelHeight/2),0 $($panelWidth/2),$(-$panelHeight/2),0 $($panelWidth/2),$($panelHeight/2),0 $(-$panelWidth/2),$($panelHeight/2),0"; $panelMesh.TriangleIndices = "0,1,2 0,2,3"; $panelMesh.TextureCoordinates = "0,1 1,1 1,0 0,0"
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $panelCount; $i++) {
            $panelAngle = $i * $angleIncrement; $panelContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
            foreach ($face in @("Front", "Back")) {
                $playerKey = "Panel${i}_${face}"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
                else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
                $SyncHash.PlayerStates[$playerKey] = $playerState
                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($panelMesh.Clone(), $material); if ($face -eq "Back") { $faceModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180))) }
                $panelContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $faceModel })) | Out-Null; Start-NextMedia -PlayerKey $playerKey
            }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$carouselRadius))) | Out-Null; $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', $panelAngle))))) | Out-Null
            $verticalTranslate = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $transformGroup.Children.Add($verticalTranslate) | Out-Null; $panelContainer.Transform = $transformGroup; $carouselContainer.Children.Add($panelContainer) | Out-Null
            $verticalAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = -0.5; To = 0.5; Duration = [TimeSpan]::FromSeconds(4); AutoReverse = $true; RepeatBehavior = "Forever"; BeginTime = [TimeSpan]::FromSeconds($i * 0.5) }; $verticalTranslate.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $verticalAnim)
        }
        $carouselAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(60)) -Property @{ RepeatBehavior="Forever" }; $Window.RegisterName("CarouselRotation_Main", $carouselContainer.Transform.Rotation); $carouselRotation = $Window.FindName("CarouselRotation_Main")
        $carouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $carouselAnim); $SyncHash.Animations = @{ Carousel = $carouselAnim }; $SyncHash.Rotations = @{ Carousel = $carouselRotation }
    }

    function Setup-ConcentricFunnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Concentric Funnel"; $Viewport.Camera.Position = "0,2,18"; $Viewport.Camera.LookDirection = "0,-0.1,-1"
        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0"/></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0"/></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer) | Out-Null; $funnelModelGroup = $funnelContainer.Content; $numberOfGroups = 6; $sharedMaterials = @(); $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $numberOfGroups; $i++) {
            $playerKey = "Group$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null; $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState; Start-NextMedia -PlayerKey $playerKey
        }
        $numberOfRings = 5; $panelsPerRing = 12; $startRadius = 7.0; $endRadius = 1.0; $totalHeight = 8.0; $radiusStep = ($startRadius - $endRadius) / $numberOfRings; $angleStep = (2 * [Math]::PI) / $panelsPerRing
        for ($r = 0; $r -lt $numberOfRings; $r++) {
            $outerR = $startRadius - ($r * $radiusStep); $innerR = $startRadius - (($r + 1) * $radiusStep); $progressOuter = $r / $numberOfRings; $progressOuter_eased = $progressOuter * $progressOuter; $progressInner = ($r + 1) / $numberOfRings; $progressInner_eased = $progressInner * $progressInner
            $yOuter = $totalHeight / 2 - ($progressOuter_eased * $totalHeight); $yInner = $totalHeight / 2 - ($progressInner_eased * $totalHeight)
            for ($p = 0; $p -lt $panelsPerRing; $p++) {
                $theta1 = $p * $angleStep; $theta2 = ($p + 1) * $angleStep; $p1 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta1), $yOuter, $outerR * [Math]::Sin($theta1)); $p2 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta2), $yOuter, $outerR * [Math]::Sin($theta2)); $p3 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta2), $yInner, $innerR * [Math]::Sin($theta2)); $p4 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta1), $yInner, $innerR * [Math]::Sin($theta1))
                $panelMesh = New-SpiralingPanelMesh -p1 $p1 -p2 $p2 -p3 $p3 -p4 $p4; $materialToUse = $sharedMaterials | Get-Random; $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $materialToUse; BackMaterial = $materialToUse.Clone() }
                $funnelModelGroup.Children.Add($geometryModel) | Out-Null
            }
        }
    }

    function Setup-FloatingCubes {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Floating Cubes"; $Viewport.Camera.Position = "0,0,15"
        $cubeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $cubeMesh.Positions = "-0.5,-0.5,0.5 0.5,-0.5,0.5 0.5,0.5,0.5 -0.5,0.5,0.5 -0.5,-0.5,-0.5 0.5,-0.5,-0.5 0.5,0.5,-0.5 -0.5,0.5,-0.5"; $cubeMesh.TriangleIndices = "0,1,2 0,2,3 4,7,6 4,6,5 0,4,5 0,5,1 1,5,6 1,6,2 2,6,7 2,7,3 3,7,4 3,4,0"; $cubeMesh.TextureCoordinates = "0,1 1,1 1,0 0,0 1,1 0,1 0,0 1,0 0,1 1,1 1,0 0,0 0,1 1,1 1,0 0,0"
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }; $SyncHash.Animations = [hashtable]::Synchronized(@{})
        for ($i = 0; $i -lt 6; $i++) {
            $playerKey = "Cube$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $model = New-Object System.Windows.Media.Media3D.GeometryModel3D($cubeMesh, $material); $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }; $Viewport.Children.Add($modelVisual)
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D(2, 2, 2); $transformGroup.Children.Add($scaleTransform)
            $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D)); $rotateTransform.Rotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0)); $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform); $modelVisual.Transform = $transformGroup
            $rotAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds((Get-Random -Minimum 15 -Maximum 45))) -Property @{ RepeatBehavior = 'Forever' }; $rotateTransform.Rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $rotAnim)
            $posAnimX = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }; $posAnimY = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }; $posAnimZ = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames -Property @{ RepeatBehavior = 'Forever' }
            $durationSeconds = (Get-Random -Minimum 20 -Maximum 60); $posAnimX.Duration = [TimeSpan]::FromSeconds($durationSeconds); $posAnimY.Duration = [TimeSpan]::FromSeconds($durationSeconds); $posAnimZ.Duration = [TimeSpan]::FromSeconds($durationSeconds)
            $xRadius = (Get-Random -Minimum 3 -Maximum 8); $yRadius = (Get-Random -Minimum 2 -Maximum 6); $zRadius = (Get-Random -Minimum 1 -Maximum 4); $timeOffset = ($durationSeconds / 4) * (Get-Random -Minimum 0 -Maximum 3)
            for ($k = 0; $k -le 4; $k++) {
                $time = [System.Windows.Media.Animation.KeyTime]::FromTimeSpan([TimeSpan]::FromSeconds(($k * $durationSeconds / 4 + $timeOffset) % $durationSeconds)); $angle = $k * [Math]::PI / 2; $x = $xRadius * [Math]::Cos($angle); $y = $yRadius * [Math]::Sin($angle); $z = $zRadius * [Math]::Cos($angle * 2)
                $posAnimX.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($x, $time))) | Out-Null; $posAnimY.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($y, $time))) | Out-Null; $posAnimZ.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($z, $time))) | Out-Null
            }
            $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $posAnimX); $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $posAnimY); $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $posAnimZ)
            $SyncHash.Animations[$modelVisual.GetHashCode()] = [pscustomobject]@{ Rotation = $rotAnim; PositionX = $posAnimX; PositionY = $posAnimY; PositionZ = $posAnimZ; TranslateTransform = $translateTransform; RotateTransform = $rotateTransform }; Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-Funnel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Funnel"; $Viewport.Camera.Position = "0,0,15"; $Viewport.Camera.FieldOfView = "70"
        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0" /></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0" /></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer); $modelGroup = $funnelContainer.Content; $panelCount = 8; $panelMesh = New-SpiralingPanelMesh -startRadius 5.0 -endRadius 1.0 -height 10.0 -arcAngle (360.0 / $panelCount) -twistAngle 90
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $panelCount; $i++) {
            $playerKey = "Panel$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $material; BackMaterial = $material.Clone() }; $angle = $i * (360.0 / $panelCount)
            $panelModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), $angle))); $modelGroup.Children.Add($panelModel) | Out-Null; Start-NextMedia -PlayerKey $playerKey
        }
        $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }; $axisAngleY = $funnelContainer.Transform.Children[0].Rotation; $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ Y = $animY }; $SyncHash.Rotations = @{ Y = $axisAngleY }
    }

    function Setup-FunnelSingle {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Media Funnel (Single)"; $Viewport.Camera.Position = "0,0,15"; $Viewport.Camera.FieldOfView = "70"
        $funnelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Content><Model3DGroup/></ModelVisual3D.Content><ModelVisual3D.Transform><Transform3DGroup><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Angle="40" Axis="1,0,0" /></RotateTransform3D.Rotation></RotateTransform3D><TranslateTransform3D OffsetY="-2.0" /></Transform3DGroup></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($funnelContainer); $modelGroup = $funnelContainer.Content; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null
        if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
        else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey "Funnel" }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey "Funnel" -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey "Funnel" -Reason $e.ErrorException.Message }.GetNewClosure() } }
        $SyncHash.PlayerStates["Funnel"] = $playerState
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
        $panelCount = 8; $panelMesh = New-SpiralingPanelMesh -startRadius 5.0 -endRadius 1.0 -height 10.0 -arcAngle (360.0 / $panelCount) -twistAngle 90
        for ($i = 0; $i -lt $panelCount; $i++) {
            $angle = $i * (360.0 / $panelCount); $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $panelMesh; Material = $material; BackMaterial = $material.Clone() }; $panelModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), $angle))); $modelGroup.Children.Add($panelModel) | Out-Null
        }
        Start-NextMedia -PlayerKey "Funnel"; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) -Property @{ RepeatBehavior="Forever" }; $axisAngleY = $funnelContainer.Transform.Children[0].Rotation; $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
        $SyncHash.Animations = @{ Y = $animY }; $SyncHash.Rotations = @{ Y = $axisAngleY }
    }

    function Setup-ScrollingHorizontal {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Scrolling Horizontal"; $Viewport.Visibility = 'Collapsed'; $scrollingCanvas = $Window.FindName("scrollingCanvas"); $scrollingCanvas.Visibility = 'Visible'; $SyncHash.scrollingPanel = $Window.FindName("scrollingPanel"); $SyncHash.scrollingPanel.Orientation = 'Horizontal'
        $numberOfPlayers = 8; $playerHeight = $Window.Height
        $SyncHash.startAnimationScriptBlock = {
            param($sh)
            $panel = $sh.scrollingPanel; $totalWidth = 0; foreach ($child in $panel.Children) { $totalWidth += $child.ActualWidth + $child.Margin.Left + $child.Margin.Right }; if ($totalWidth -eq 0) { return }
            $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, -$totalWidth, [TimeSpan]::FromSeconds(90)) -Property @{ RepeatBehavior = "Forever" }; $storyboard = New-Object System.Windows.Media.Animation.Storyboard; $storyboard.Children.Add($anim)
            $transformGroup = New-Object System.Windows.Media.TransformGroup; $sh.translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0); $transformGroup.Children.Add($sh.translateTransform) | Out-Null; $panel.RenderTransform = $transformGroup
            [System.Windows.Media.Animation.Storyboard]::SetTarget($anim, $panel); [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($anim, (New-Object System.Windows.PropertyPath("(UIElement.RenderTransform).(TransformGroup.Children)[0].(TranslateTransform.X)"))); $storyboard.Begin($panel, $true)
            $sh.Storyboard = $storyboard; $sh.Animations = @{ Scroll = $anim }; $sh.Transforms = @{ Scroll = $sh.translateTransform }
        }
        for ($i = 0; $i -lt $numberOfPlayers; $i++) {
            $playerKey = "Scroller$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Height = $playerHeight; Width = $playerHeight; Margin = '5' }; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $SyncHash.scrollingPanel.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState; Start-NextMedia -PlayerKey $playerKey
        }
        $Window.Add_Loaded({ $Window.Dispatcher.InvokeAsync({ & $SyncHash.startAnimationScriptBlock $SyncHash }, "Loaded") | Out-Null })
    }

    function Setup-ScrollingVertical {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Scrolling Vertical"; $Viewport.Visibility = 'Collapsed'; $scrollingCanvas = $Window.FindName("scrollingCanvas"); $scrollingCanvas.Visibility = 'Visible'; $SyncHash.scrollingPanel = $Window.FindName("scrollingPanel"); $SyncHash.scrollingPanel.Orientation = 'Vertical'
        $SyncHash.startAnimationScriptBlock = {
            param($sh)
            $panel = $sh.scrollingPanel; $totalHeight = 0; foreach ($child in $panel.Children) { $totalHeight += $child.ActualHeight + $child.Margin.Top + $child.Margin.Bottom }; if ($totalHeight -eq 0) { return }
            $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, -$totalHeight, [TimeSpan]::FromSeconds(90)) -Property @{ RepeatBehavior = "Forever" }; $storyboard = New-Object System.Windows.Media.Animation.Storyboard; $storyboard.Children.Add($anim)
            $transformGroup = New-Object System.Windows.Media.TransformGroup; $sh.translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0); $transformGroup.Children.Add($sh.translateTransform) | Out-Null; $panel.RenderTransform = $transformGroup
            [System.Windows.Media.Animation.Storyboard]::SetTarget($anim, $panel); [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($anim, (New-Object System.Windows.PropertyPath("(UIElement.RenderTransform).(TransformGroup.Children)[0].(TranslateTransform.Y)"))); $storyboard.Begin($panel, $true)
            $sh.Storyboard = $storyboard; $sh.Animations = @{ Scroll = $anim }; $sh.Transforms = @{ Scroll = $sh.translateTransform }
        }
        $numberOfPlayers = 6; $playerWidth = $Window.Width
        for ($i = 0; $i -lt $numberOfPlayers; $i++) {
            $playerKey = "Scroller$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Width = $playerWidth; Height = ($playerWidth * 0.5625); Margin = '5' }; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $SyncHash.scrollingPanel.Children.Add($mediaHostGrid) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState; Start-NextMedia -PlayerKey $playerKey
        }    
        $Window.Add_Loaded({ $Window.Dispatcher.InvokeAsync({ & $SyncHash.startAnimationScriptBlock $SyncHash }, "Loaded") | Out-Null })
    }

    function Setup-Aquarium {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Aquarium"; $Viewport.Camera.Position = "0,0,10"
        $rightFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="0.6,0,0 0.5,0.14,0 0.2,0.196,0 -0.1,0.175,0 -0.4,0.105,0 -0.6,0.14,0 -0.667,0.12,0 -0.6,0,0 -0.667,-0.12,0 -0.6,-0.14,0 -0.4,-0.105,0 -0.1,-0.175,0 0.2,-0.196,0 0.5,-0.14,0 0.667,0.05,0 0.667,-0.05,0" TriangleIndices="0,14,1 0,1,13 0,13,15 1,2,12 1,12,13 2,3,11 2,11,12 3,4,10 3,10,11 4,5,7 4,7,10 5,6,7 7,8,9 7,9,10" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
        $leftFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="-0.6,0,0 -0.5,0.14,0 -0.2,0.196,0 0.1,0.175,0 0.4,0.105,0 0.6,0.14,0 0.667,0.12,0 0.6,0,0 0.667,-0.12,0 0.6,-0.14,0 0.4,-0.105,0 0.1,-0.175,0 -0.2,-0.196,0 -0.5,-0.14,0 -0.667,0.05,0 -0.667,-0.05,0" TriangleIndices="0,1,14 0,13,1 0,15,13 1,12,2 1,13,12 2,11,3 2,12,11 3,10,4 3,11,10 4,7,5 4,10,7 5,7,6 7,9,8 7,10,9" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
        $SyncHash.RightFacingFish = [Windows.Markup.XamlReader]::Parse($rightFishXaml); $SyncHash.LeftFacingFish = [Windows.Markup.XamlReader]::Parse($leftFishXaml)
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }; $SyncHash.FloatingObjects = [System.Collections.ArrayList]::new()
        for ($i = 0; $i -lt 6; $i++) {
            $fishContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $fishContainer.Transform = $translateTransform; $Viewport.Children.Add($fishContainer) | Out-Null
            foreach ($face in @("Front", "Back")) {
                $playerKey = "Fish${i}_${face}"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null
                if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
                else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
                $SyncHash.PlayerStates[$playerKey] = $playerState
                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                $fishModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Material = $material }; if ($face -eq "Back") { $fishModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180))) }
                $fishContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $fishModel })) | Out-Null; Start-NextMedia -PlayerKey $playerKey
            }
            $velocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -0.75 -Maximum 0.75), (Get-Random -Minimum -0.25 -Maximum 0.25)); if ($velocity.X -eq 0) { $velocity.X = 0.5 }
            $fishObject = [pscustomobject]@{ Visual = $fishContainer; Translate = $translateTransform; Velocity = $velocity; RightGeometry = $SyncHash.RightFacingFish.Clone(); LeftGeometry = $SyncHash.LeftFacingFish.Clone() }; [void]$SyncHash.FloatingObjects.Add($fishObject)
        }
    }

    function Setup-RollerCoaster {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Roller Coaster"; $Viewport.Camera.Position = "0,5,25"; $Viewport.Camera.LookDirection = "0,-0.2,-1"; $Viewport.Camera.FieldOfView = "60"
        $script:pathFunc = { param([double]$t) $t_rad = $t * 2 * [Math]::PI; $x_base = 12 * [Math]::Cos($t_rad); $z_base = 5 * [Math]::Sin($t_rad); $y_base = 2.5 * [Math]::Sin(3 * $t_rad) - 1.5 * [Math]::Cos(5 * $t_rad) + [Math]::Sin($t_rad); $loop_radius = 5.0; $loop_center_t = 0.5; $loop_width = 0.08; $loop_influence = [Math]::Exp(-[Math]::Pow($t - $loop_center_t, 2) / (2 * [Math]::Pow($loop_width, 2))); $y = $y_base + $loop_radius * [Math]::Sin(($t - $loop_center_t) / $loop_width * [Math]::PI) * $loop_influence; $x = $x_base + $loop_radius * ([Math]::Cos(($t - $loop_center_t) / $loop_width * [Math]::PI) + 1) * $loop_influence; return [System.Windows.Media.Media3D.Point3D]::new($x, $y, $z_base) }
        $SyncHash.PathFunc = $script:pathFunc; $segments = 800; $SyncHash.CoasterSegments = $segments
        $pathDataArray = for ($i = 0; $i -le $segments; $i++) { $t = $i / $segments; $p1 = & $script:pathFunc $t; $p2 = & $script:pathFunc ($t + 0.001); $tangent = $p2 - $p1; $tangent.Normalize(); $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '0,1,0'); if ($normal.LengthSquared -lt 1e-6) { $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '1,0,0') }; $normal.Normalize(); $up = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($normal, $tangent); [pscustomobject]@{ Point = $p1; Up = $up; Normal = $normal } }
        $SyncHash.PathData = @{}; for($i=0; $i -lt $pathDataArray.Count; $i++){ $SyncHash.PathData[$i] = @{ Up = $pathDataArray[$i].Up } }
        $trackModelGroup = New-CoasterTrackModelGroup -PathData $pathDataArray -segments $segments; $trackContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $trackModelGroup }; $Viewport.Children.Add($trackContainer)
        $numberOfCars = 12; $sphereRadius = 1.34; $SyncHash.SphereRadius = $sphereRadius; $sphereMesh = New-SphereMesh -radius $sphereRadius; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }; $SyncHash.CarStates = [hashtable]::Synchronized(@{})
        for ($i = 0; $i -lt $numberOfCars; $i++) {
            $playerKey = "Car$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState
            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
            $carModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($sphereMesh, $material); $carVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $carModel }; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $carVisual.Transform = $translateTransform
            $Viewport.Children.Add($carVisual) | Out-Null; $SyncHash.CarStates[$i] = @{ TranslateTransform = $translateTransform; Progress = $i / $numberOfCars }; Start-NextMedia -PlayerKey $playerKey
        }
    }

    function Setup-SphereVortex {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Sphere Vortex"; $Viewport.Camera.Position = "0,0,15"; $Viewport.Camera.LookDirection = "0,0,-1"; $Viewport.Camera.FieldOfView = "70"
        $vortexContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $Viewport.Children.Add($vortexContainer) | Out-Null
        $masterRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), 0); $vortexContainer.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D($masterRotation); $masterAnim = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(45)) -Property @{ RepeatBehavior = "Forever" }; $SyncHash.Animations = @{ Master = $masterAnim }; $SyncHash.Rotations = @{ Master = $masterRotation }
        $sphereCount = 32; $numberOfGroups = 6; $sphereMesh = New-SphereMesh -radius 1.0; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $sharedMaterials = @()
        for ($i = 0; $i -lt $numberOfGroups; $i++) {
            $playerKey = "Group$i"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
            $VisualHost.Children.Add($mediaHostGrid) | Out-Null; $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush; Color = [System.Windows.Media.Colors]::White }
            if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
            else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
            $SyncHash.PlayerStates[$playerKey] = $playerState; Start-NextMedia -PlayerKey $playerKey
        }
        $SyncHash.SharedMaterials = $sharedMaterials; $SyncHash.SphereStates = [hashtable]::Synchronized(@{})
        for ($i = 0; $i -lt $sphereCount; $i++) {
            $materialToUse = $sharedMaterials[$i % $numberOfGroups]; $sphereMaterial = $materialToUse.Clone(); $sphereGeometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $sphereMesh; Material = $sphereMaterial }; $sphereContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $sphereGeometryModel }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D; $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D; $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D
            $transformGroup.Children.Add($rotateTransform); $transformGroup.Children.Add($translateTransform); $transformGroup.Children.Add($scaleTransform); $sphereContainer.Transform = $transformGroup; $vortexContainer.Children.Add($sphereContainer) | Out-Null
            $SyncHash.SphereStates[$i] = @{ SphereModel = $sphereGeometryModel; TranslateTransform = $translateTransform; RotateTransform = $rotateTransform; ScaleTransform = $scaleTransform; CurrentAngle = (720.0 / $sphereCount) * $i }
        }
        $flowTimer = New-Object System.Windows.Threading.DispatcherTimer; $flowTimer.Interval = [TimeSpan]::FromSeconds(2.0)
        $flowTimer.Add_Tick({ if ($SyncHash.Paused) { return }; $lastMaterial = $SyncHash.SharedMaterials[-1]; for ($i = $SyncHash.SharedMaterials.Count - 1; $i -gt 0; $i--) { $SyncHash.SharedMaterials[$i] = $SyncHash.SharedMaterials[$i-1] }; $SyncHash.SharedMaterials[0] = $lastMaterial; foreach ($i in 0..($SyncHash.SphereStates.Count-1)) { $SyncHash.SphereStates[$i].SphereModel.Material = $SyncHash.SharedMaterials[$i % $SyncHash.SharedMaterials.Count] } })
        $SyncHash.FlowTimer = $flowTimer; $flowTimer.Start(); $masterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $masterAnim)
    }

    function Setup-Pie3D {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "3D Pie"; $Viewport.Camera.Position = "0,0,8"
        $pieContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D; $pieModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup; $pieContainer.Content = $pieModelGroup; $Viewport.Children.Add($pieContainer) | Out-Null
        $numberOfSlices = 8; $sliceAngle = 360.0 / $numberOfSlices; $pieRadius = 2.5; $pieCenter = New-Object System.Windows.Point(0, 0); $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $numberOfSlices; $i++) {
            $startAngle = $i * $sliceAngle; $sliceModel = New-PieSliceModel -center $pieCenter -radius $pieRadius -startAngleDeg $startAngle -sliceAngleDeg $sliceAngle
            foreach ($face in @('Front', 'Back')) {
                $playerKey = "Slice${i}_${face}"; $mediaHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Background = [System.Windows.Media.Brushes]::Black }; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; Margin='10,0,10,0'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null
                if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
                else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
                $SyncHash.PlayerStates[$playerKey] = $playerState
                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid; Stretch = 'Fill' }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                if ($face -eq 'Front') { $sliceModel.Material = $material } else { $sliceModel.BackMaterial = $material }; Start-NextMedia -PlayerKey $playerKey
            }
            $pieModelGroup.Children.Add($sliceModel) | Out-Null
        }
        $axisAngleX = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 0); $axisAngleY = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), 0); $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleX))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($axisAngleY))); $pieContainer.Transform = $transformGroup
        $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(40)) -Property @{ RepeatBehavior = 'Forever' }; $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(60)) -Property @{ RepeatBehavior = 'Forever' }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY); $SyncHash.Animations = @{ X = $animX; Y = $animY }; $SyncHash.Rotations = @{ X = $axisAngleX; Y = $axisAngleY }
    }

    function Setup-Pinwheel {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Pinwheel Kaleidoscope"; $Viewport.Camera.Position = "0,0,8"; $SyncHash.Animations = [hashtable]::Synchronized(@{}); $SyncHash.Rotations = [hashtable]::Synchronized(@{}); $SyncHash.Transforms = [hashtable]::Synchronized(@{})
        $pinwheelContainer = [Windows.Markup.XamlReader]::Parse('<ModelVisual3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><ModelVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="PinwheelRotation" Axis="0,0,1" Angle="0"/></RotateTransform3D.Rotation></RotateTransform3D></ModelVisual3D.Transform></ModelVisual3D>')
        $Viewport.Children.Add($pinwheelContainer) | Out-Null; $bladeCount = 8; $angleIncrement = 360 / $bladeCount; $bladeWidth = 1.5; $bladeHeight = 4.0; $bladeTiltAngle = 45; $pinwheelRadius = 1.5
        $bladeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2), (-$bladeHeight/2), 0)) ); $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2), (-$bladeHeight/2), 0)) ); $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2),  ($bladeHeight/2), 0)) ); $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2),  ($bladeHeight/2), 0)) )
        $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(1); $bladeMesh.TriangleIndices.Add(2); $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(2); $bladeMesh.TriangleIndices.Add(3)
        $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,1)) ); $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,1)) ); $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,0)) ); $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,0)) )
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        for ($i = 0; $i -lt $bladeCount; $i++) {
            $bladeAngle = $i * $angleIncrement; $bladeContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
            foreach ($face in @("Front", "Back")) {
                $playerKey = "Blade${i}_${face}"; $mediaHostGrid = New-Object System.Windows.Controls.Grid; $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
                $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }; $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
                $VisualHost.Children.Add($mediaHostGrid) | Out-Null
                if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid } }
                else { $playerState = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $playerKey }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $playerKey -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $playerKey -Reason $e.ErrorException.Message }.GetNewClosure() } }
                $SyncHash.PlayerStates[$playerKey] = $playerState
                $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
                $faceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($bladeMesh.Clone(), $material); if ($face -eq "Back") { $faceModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180))) }
                $bladeContainer.Children.Add((New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $faceModel })) | Out-Null; Start-NextMedia -PlayerKey $playerKey
            }
            $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup; $bladeRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{ Axis = '0,1,0'; Angle = 0 }; $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($bladeRotation))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', $bladeTiltAngle))))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $pinwheelRadius, 0))); $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', $bladeAngle))))); $bladeContainer.Transform = $transformGroup; $pinwheelContainer.Children.Add($bladeContainer) | Out-Null
            $bladeAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(20); RepeatBehavior = "Forever" }; $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $bladeAnim)
            $SyncHash.Animations["Blade$i"] = $bladeAnim; $SyncHash.Rotations["Blade$i"] = $bladeRotation
        }
        $pinwheelRotationName = "PinwheelRotation_$(Get-Random)"; $Window.RegisterName($pinwheelRotationName, $pinwheelContainer.Transform.Rotation); $pinwheelRotation = $Window.FindName($pinwheelRotationName)
        $pinwheelAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(60); RepeatBehavior = "Forever" }; $pinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $pinwheelAnim)
        $SyncHash.Animations["Pinwheel"] = $pinwheelAnim; $SyncHash.Rotations["Pinwheel"] = $pinwheelRotation
    }

    function New-ReindeerMediaObject {
        param([string]$Name, [double]$ScaleX, [double]$ScaleY, [string]$ClipData, [System.Windows.Media.Media3D.MeshGeometry3D]$PlaneMesh, [System.Windows.Controls.Canvas]$VisualHost, [hashtable]$SyncHash)
        $mediaHostGrid = New-Object System.Windows.Controls.Grid; $mediaHostGrid.Width = 500; $mediaHostGrid.Height = 300; if ($SyncHash.UseTransparentEffect) { $mediaHostGrid.Opacity = 0.7 }
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter; $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        if ($ClipData) {
            $geometry = [System.Windows.Media.Geometry]::Parse($ClipData).Clone(); $bounds = $geometry.Bounds; $scaleX_factor = if ($bounds.Width -gt 0) { $mediaHostGrid.Width / $bounds.Width } else { 1 }; $scaleY_factor = if ($bounds.Height -gt 0) { $mediaHostGrid.Height / $bounds.Height } else { 1 }; $scale = [Math]::Min($scaleX_factor, $scaleY_factor) * 0.95
            $translateX = ($mediaHostGrid.Width - ($bounds.Width * $scale)) / 2 - ($bounds.X * $scale); $translateY = ($mediaHostGrid.Height - ($bounds.Height * $scale)) / 2 - ($bounds.Y * $scale)
            $transformGroup = New-Object System.Windows.Media.TransformGroup; $transformGroup.Children.Add((New-Object System.Windows.Media.ScaleTransform($scale, $scale))); $transformGroup.Children.Add((New-Object System.Windows.Media.TranslateTransform($translateX, $translateY))); $geometry.Transform = $transformGroup; $mediaHostGrid.Clip = $geometry
            $borderPath = New-Object System.Windows.Shapes.Path; $borderPath.Data = $geometry; $borderPath.Stroke = [System.Windows.Media.Brushes]::LightGray; $borderPath.StrokeThickness = 2; $borderPath.Fill = [System.Windows.Media.Brushes]::Transparent; $mediaHostGrid.Children.Add($borderPath) | Out-Null
        }
        if ($Name -eq "Deer1") {
            $nose = New-Object System.Windows.Shapes.Ellipse -Property @{ Width = 24; Height = 24; Fill = [System.Windows.Media.Brushes]::Red; HorizontalAlignment = 'Left'; VerticalAlignment = 'Top'; Margin = [System.Windows.Thickness]::new(91.5, 81.5, 0, 0) }; $mediaHostGrid.Children.Add($nose) | Out-Null
            $blinkAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(1.0, 0.1, [TimeSpan]::FromSeconds(1.5)) -Property @{ AutoReverse = $true; RepeatBehavior = "Forever" }; $nose.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $blinkAnimation)
        }
        $VisualHost.Children.Add($mediaHostGrid) | Out-Null
        if ($SyncHash.UseFfmpeg) { $playerState = @{ ContentPresenter = $contentPresenter; MediaHostGrid = $mediaHostGrid } }
        else { $playerState = @{ ContentPresenter = $contentPresenter; MediaHostGrid = $mediaHostGrid; MediaEndedHandler = { Handle-MediaEnded_ME -PlayerKey $Name }.GetNewClosure(); MediaOpenedHandler = { Handle-MediaOpened_ME -PlayerKey $Name -EventArgs $args[0] }.GetNewClosure(); MediaFailedHandler = { param($s, $e) Handle-MediaFailure -PlayerKey $Name -Reason $e.ErrorException.Message }.GetNewClosure() } }
        $SyncHash.PlayerStates[$Name] = $playerState
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }; $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }; $material = New-Object $materialType -Property @{ Brush = $visualBrush }; if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }
        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $PlaneMesh; Material = $material }; $model.Transform = New-Object System.Windows.Media.Media3D.ScaleTransform3D($ScaleX, $ScaleY, 1); $backMaterial = $material.Clone(); $backTransform = New-Object System.Windows.Media.ScaleTransform -Property @{ ScaleX = -1 }; $backMaterial.Brush.RelativeTransform = $backTransform; $model.BackMaterial = $backMaterial
        $objectContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }; $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, 0); $objectContainer.Transform = $translateTransform; $SyncHash.PlayerStates[$Name].TranslateTransform = $translateTransform
        return $objectContainer
    }

    function Setup-ReindeerSleigh {
        param($Window, $Viewport, $VisualHost, $SyncHash)
        $Window.Title = "Reindeer Sleigh"; $Viewport.Camera.Position = "0,0,20"
        if ($SyncHash.NightSky) { $backgroundCanvas = $Window.FindName("backgroundCanvas"); if ($SyncHash.UseTransparentEffect) { $backgroundCanvas.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(128, 0, 0, 0)) } else { $backgroundCanvas.Background = [System.Windows.Media.Brushes]::Black } }
        if ($SyncHash.TwinklingStars) {
            $starCanvas = $Window.FindName("backgroundCanvas"); $rand = [Random]::new()
            for ($i = 0; $i -lt 200; $i++) {
                $star = New-Object System.Windows.Shapes.Ellipse; $size = $rand.NextDouble() * 3 + 1; $star.Width = $size; $star.Height = $size; $star.Fill = [System.Windows.Media.Brushes]::White
                [System.Windows.Controls.Canvas]::SetLeft($star, $rand.NextDouble() * $Window.Width); [System.Windows.Controls.Canvas]::SetTop($star, $rand.NextDouble() * $Window.Height)
                $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0.1, 1.0, [TimeSpan]::FromSeconds($rand.NextDouble() * 2 + 0.5)) -Property @{ AutoReverse = $true; RepeatBehavior = "Forever" }; $star.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $anim); $starCanvas.Children.Add($star) | Out-Null
            }
        }
        $aspectRatio = $Window.Width / $Window.Height; $fovRadians = 60 * ([Math]::PI / 180); $visibleHeight = 2 * 20 * [Math]::Tan($fovRadians / 2); $visibleWidth = $visibleHeight * $aspectRatio
        $SyncHash.RightLimit = ($visibleWidth / 2); $SyncHash.LeftLimit = -($visibleWidth / 2) - 12; $SyncHash.LeadPositionX = $SyncHash.RightLimit
        $SyncHash.HistoryX = [System.Collections.Generic.List[double]]::new(); $SyncHash.HistoryY = [System.Collections.Generic.List[double]]::new(); $SyncHash.HistoryDist = [System.Collections.Generic.List[double]]::new()
        $planeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D -Property @{ Positions = "-0.5,-0.5,0 0.5,-0.5,0 0.5,0.5,0 -0.5,0.5,0"; TriangleIndices = "0,1,2 0,2,3"; TextureCoordinates = "0,1 1,1 1,0 0,0" }; $planeMesh.Freeze()
        $sleighPath = "M27,24 c0.6,0,1-0.4,1-1 c0-3.8,0.9-7.5,2.6-10.9 l1.3-2.6 c0.2-0.3,0.1-0.7,0-1 C31.7,8.2,31.3,8,31,8 L28,11 C27.6,10.8 27.1,8.6 26.4,7.2 C26.4,6 25.4,2 24,1.0 C21.6,5.6 20.8,7.2 21.2,8.8 C21.4,9.8 22.2,10.6 23.2,11.2 C22.2,11.6 21.2,12.6 21,14 C20.8,15.6 21.6,16.6 22.8,17.2 L21,16.4 c-0.7,1-1.8,1.5-2.9,1.5 c-1.3,0-2.4-0.7-3.1-1.7 c-1.2-2-3.4-3.3-5.7-3.3 H6 c-0.4,0-0.7,0.2-0.9,0.5 c-0.2,0.3-0.2,0.7,0,1 l5,9 c0.2,0.3,0.5,0.5,0.9,0.5 h1.5 l-1.2,3 l-5.8,0 c-1,0-2-0.3-2.7-1 C2.3,25.5,2,24.9,2,24.3 s0.3-1.3,0.8-1.7 c0.8-0.7,2.1-0.7,2.8,0 c0.4,0.4,1,0.4,1.4,0 c0.4-0.4,0.4-1,0-1.4 c-1.5-1.4-4-1.4-5.6,0 c-0.9,0.8-1.4,2-1.4,3.2 s0.5,2.3,1.4,3.2 c1.1,1,2.5,1.5,3.9,1.5 c0.1,0,0.1,0,0.2,0 l6.4,0 l13,0 l5,0 c0.6,0,1-0.4,1-1 s-0.4-1-1-1 l-4.3,0 l-1.2-3 H27 z"
        $reindeerPath = "M505.474,436.173c0,0-64.29-87.208-86.987-116.34c-10.781-13.882-21.414-24.368-32.196-32.048 c-0.055-0.037-0.101-0.064-0.157-0.101c0.804-0.139,1.468-0.148,2.299-0.305c26.953-5.28,33.562-41.906,31.466-57.006 c-2.086-15.101-8.538-20.528-18.313-3.286c-7.763,13.624-34.641,43.272-70.234,43.124l-58.041-0.184 c-3.775-0.037-7.42-0.147-10.928-0.37c-21.267-1.439-37.882-7.42-54.497-24.147c-8.916-8.999-12.627-24.718-15.507-39.034 c6.849,0.028,15.23,0.055,19.836,0.082c7.699,0,11.575-1.92,16.421-6.71c4.865-4.809,8.64-11.52,14.104-20.242 c6.019-9.674,0.036-18.276-6.72-18.313c-6.756,0-30.424-0.111-30.424-0.111c-5.787,0-11.592,1.92-14.51,3.84 c-2.427,1.597-7.31,5.584-12.572,10.08l-91.62-0.332c-12.444-0.074-22.855,1.957-31.272,5.427 c-13.542,5.575-22.117,14.99-26.206,25.993c-12.304,32.943,15.572,79.972,70.105,80.157l30.599,0.111 c9.609,31.568,25.789,65.72,25.789,65.72l-80.756,94.27c-7.108,5.917-9.831,19.236-1.532,27.562 c8.27,8.326,18.017,8.086,33.441-3.729l96.642-76.824l180.897,0.48l7.606,5.501l87.042,63.911 c11.852,8.935,21.598,7.145,27.894,0.083C514.252,455.464,513.44,445.394,505.474,436.173z M37.974,166.148c0.037-10.495-8.419-19.042-18.95-19.07C8.565,147.049,0.056,155.515,0,165.991 c-0.027,10.494,8.418,19.032,18.922,19.069C29.408,185.089,37.937,176.633,37.974,166.148z M154.277,142.093c-7.191,8.603-14.197,14.805-20.095,19.08h37.817c7.651-8.806,15.34-19.624,22.412-32.74 c11.261-20.897,20.934-47.62,26.361-81.19c1.145-7.052-3.655-13.689-10.67-14.842c-7.089-1.145-13.734,3.618-14.842,10.707 c-1.736,10.531-3.877,20.233-6.342,29.204l-17.722-18.424c-4.994-5.132-13.154-5.27-18.313-0.333 c-5.141,4.948-5.299,13.145-0.333,18.276l26.602,27.618c0.12,0.111,0.231,0.184,0.314,0.296 c-3.268,7.79-6.812,14.806-10.421,21.082l-16.541-11.926c-5.797-4.162-13.883-2.88-18.055,2.926 c-4.172,5.787-2.88,13.882,2.917,18.046L154.277,142.093z"
        $sleigh = New-ReindeerMediaObject -Name "Sleigh" -ScaleX 5 -ScaleY 3 -ClipData $sleighPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer1 = New-ReindeerMediaObject -Name "Deer1" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer2 = New-ReindeerMediaObject -Name "Deer2" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer3 = New-ReindeerMediaObject -Name "Deer3" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $deer4 = New-ReindeerMediaObject -Name "Deer4" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPath -PlaneMesh $planeMesh -VisualHost $VisualHost -SyncHash $SyncHash
        $Viewport.Children.Add($sleigh) | Out-Null; $Viewport.Children.Add($deer1) | Out-Null; $Viewport.Children.Add($deer2) | Out-Null; $Viewport.Children.Add($deer3) | Out-Null; $Viewport.Children.Add($deer4) | Out-Null
        $SyncHash.PlayerStates.Keys | ForEach-Object { Start-NextMedia -PlayerKey $_ }
    }

    function Animate-ReindeerSleigh {
        param($SyncHash)
        if ($SyncHash.Paused) { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime; $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency
        $flySpeed = 4.0 * $SyncHash.SpeedMultiplier; $bobSpeed = 2.0; $bobAmount = 1.5; $distStep = $flySpeed * $elapsed; $SyncHash.TotalDistance += $distStep; $SyncHash.LeadPositionX -= ($flySpeed * $elapsed)
        if ($SyncHash.LeadPositionX -lt $SyncHash.LeftLimit) { $lapDistanceX = $SyncHash.RightLimit - $SyncHash.LeftLimit; $SyncHash.LeadPositionX += $lapDistanceX; for ($i = 0; $i -lt $SyncHash.HistoryX.Count; $i++) { $SyncHash.HistoryX[$i] += $lapDistanceX } }
        $SyncHash.LeadPositionY = [Math]::Sin($totalTime * $bobSpeed) * $bobAmount; $SyncHash.HistoryX.Add($SyncHash.LeadPositionX); $SyncHash.HistoryY.Add($SyncHash.LeadPositionY); $SyncHash.HistoryDist.Add($SyncHash.TotalDistance)
        while ($SyncHash.HistoryDist.Count -gt 0 -and ($SyncHash.TotalDistance - $SyncHash.HistoryDist[0] -gt 25.0)) { $SyncHash.HistoryX.RemoveAt(0); $SyncHash.HistoryY.RemoveAt(0); $SyncHash.HistoryDist.RemoveAt(0) }
        $lags = @(0, 4, 8, 12, 17.7); $objects = @("Deer1", "Deer2", "Deer3", "Deer4", "Sleigh")
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $playerState = $SyncHash.PlayerStates[$objects[$i]]; if (-not $playerState) { continue }; $targetDist = $SyncHash.TotalDistance - $lags[$i]; $idx = -1
            for ($j = $SyncHash.HistoryDist.Count - 1; $j -ge 0; $j--) { if ($SyncHash.HistoryDist[$j] -le $targetDist) { $idx = $j; break } }
            if ($idx -ge 0) { $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[$idx]; $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[$idx] }
            elseif ($SyncHash.HistoryX.Count -gt 0) { $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[0] + ($SyncHash.HistoryDist[0] - $targetDist); $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[0] }
        }
    }

    function Animate-ButterflyEffect {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "ButterflyEffect") { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency
        foreach ($plane in $SyncHash.FloatingObjects) { if (-not $plane.PSObject.Properties['FlutterTransform']) { continue }; $plane.FlutterTransform.Rotation.Angle = [Math]::Sin($totalTime * 8.0) * 15.0 }
    }

    function Animate-Aquarium {
        param($SyncHash, $Viewport)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "Aquarium") { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime
        if (-not $SyncHash.StartTime) { $SyncHash.StartTime = $currentTime }; $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency
        $xBoundary = $SyncHash.xBoundary; $yBoundary = $SyncHash.yBoundary; $zBoundary = 3; $xMargin = 0.7; $yMargin = 0.3; $oscillation = [Math]::Sin($totalTime * 18.0) * 0.12
        foreach ($fish in $SyncHash.FloatingObjects) {
            $velX = $fish.Velocity.X; $velY = $fish.Velocity.Y; $velZ = $fish.Velocity.Z; $nextX = $fish.Translate.OffsetX + ($velX * $elapsed); $nextY = $fish.Translate.OffsetY + ($velY * $elapsed); $nextZ = $fish.Translate.OffsetZ + ($velZ * $elapsed)
            if (($nextX + $xMargin) -gt $xBoundary -and $velX -gt 0) { $velX *= -1; $nextX = $xBoundary - $xMargin } elseif (($nextX - $xMargin) -lt -$xBoundary -and $velX -lt 0) { $velX *= -1; $nextX = -$xBoundary + $xMargin }
            if (($nextY + $yMargin) -gt $yBoundary -and $velY -gt 0) { $velY *= -1; $nextY = $yBoundary - $yMargin } elseif (($nextY - $yMargin) -lt -$yBoundary -and $velY -lt 0) { $velY *= -1; $nextY = -$yBoundary + $yMargin }
            if (($nextZ -gt $zBoundary -and $velZ -gt 0) -or ($nextZ -lt -$zBoundary -and $velZ -lt 0)) { $velZ *= -1 }; $fish.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)
            $fish.Translate.OffsetX = $nextX; $fish.Translate.OffsetY = $nextY; $fish.Translate.OffsetZ = $nextZ
            $frontModel = $fish.Visual.Children[0].Content; $backModel = $fish.Visual.Children[1].Content; $currentGeometry = if ($fish.Velocity.X -gt 0) { $fish.RightGeometry } else { $fish.LeftGeometry }
            if ($frontModel.Geometry -ne $currentGeometry) { $frontModel.Geometry = $currentGeometry; $backModel.Geometry = $currentGeometry }
            $positions = $currentGeometry.Positions; $zWiggle = if ($fish.Velocity.X -gt 0) { $oscillation } else { -$oscillation }
            $positions[6] = [System.Windows.Media.Media3D.Point3D]::new($positions[6].X, $positions[6].Y, $zWiggle); $positions[8] = [System.Windows.Media.Media3D.Point3D]::new($positions[8].X, $positions[8].Y, $zWiggle)
            $positions[5] = [System.Windows.Media.Media3D.Point3D]::new($positions[5].X, $positions[5].Y, $zWiggle * 0.6); $positions[9] = [System.Windows.Media.Media3D.Point3D]::new($positions[9].X, $positions[9].Y, $zWiggle * 0.6)
        }
    }

    function Animate-RollerCoaster {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "RollerCoaster") { return }; $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime; $flowSpeed = 0.05 * $SyncHash.SpeedMultiplier
        foreach ($i in $SyncHash.CarStates.Keys) {
            $carState = $SyncHash.CarStates[$i]; $carState.Progress = ($carState.Progress + ($flowSpeed * $elapsed)) % 1.0; $positionOnPath = & $SyncHash.PathFunc $carState.Progress
            $pathIndex = [int]($carState.Progress * ($SyncHash.CoasterSegments - 1)); $upVector = $SyncHash.PathData[$pathIndex].Up; $finalPosition = $positionOnPath + ($upVector * ($SyncHash.SphereRadius + 0.05))
            $carState.TranslateTransform.OffsetX = $finalPosition.X; $carState.TranslateTransform.OffsetY = $finalPosition.Y; $carState.TranslateTransform.OffsetZ = $finalPosition.Z
        }
    }

    function Animate-SphereVortex {
        param($SyncHash)
        if ($SyncHash.Paused -or $SyncHash.VisualizationStyle -ne "SphereVortex") { return }; $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime
        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier; $maxScale = 1.2; $minScale = 0.1; $totalAngle = 720
        foreach ($i in $SyncHash.SphereStates.Keys) {
            $panelState = $SyncHash.SphereStates[$i]; $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle; $angle = $panelState.CurrentAngle; $angleRad = $angle * ([Math]::PI / 180.0); $progress = $angle / $totalAngle; $currentScale = $maxScale - ($progress * ($maxScale - $minScale))
            $startRadius = 3.5; $endRadius = 0.5; $startY = 5.0; $endY = -20.0; $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius)); $currentY = $startY - ($progress * ($startY - $endY)); $currentX = $currentRadius * [Math]::Cos($angleRad); $currentZ = $currentRadius * [Math]::Sin($angleRad)
            $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ; $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, $currentY - 1.5, 0); $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)
            $panelState.ScaleTransform.ScaleX = $currentScale; $panelState.ScaleTransform.ScaleY = $currentScale; $panelState.ScaleTransform.ScaleZ = $currentScale; $forward = $position - $lookAtTarget; if ($forward.Length -gt 0) { $forward.Normalize() }; $up = '0,1,0'; $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward); if ($right.Length -gt 0) { $right.Normalize() }; $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)
            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity; $matrix.M11 = $right.X; $matrix.M12 = $right.Y; $matrix.M13 = $right.Z; $matrix.M21 = $newUp.X; $matrix.M22 = $newUp.Y; $matrix.M23 = $newUp.Z; $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z
            $trace = $matrix.M11 + $matrix.M22 + $matrix.M33
            if ($trace -gt 0) { $s = 0.5 / [Math]::Sqrt($trace + 1.0); $qw = 0.25 / $s; $qx = ($matrix.M32 - $matrix.M23) * $s; $qy = ($matrix.M13 - $matrix.M31) * $s; $qz = ($matrix.M21 - $matrix.M12) * $s }
            elseif (($matrix.M11 -gt $matrix.M22) -and ($matrix.M11 -gt $matrix.M33)) { $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M11 - $matrix.M22 - $matrix.M33); $qw = ($matrix.M32 - $matrix.M23) / $s; $qx = 0.25 * $s; $qy = ($matrix.M12 + $matrix.M21) / $s; $qz = ($matrix.M13 + $matrix.M31) / $s }
            elseif ($matrix.M22 -gt $matrix.M33) { $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M22 - $matrix.M11 - $matrix.M33); $qw = ($matrix.M13 - $matrix.M31) / $s; $qx = ($matrix.M12 + $matrix.M21) / $s; $qy = 0.25 * $s; $qz = ($matrix.M23 + $matrix.M32) / $s }
            else { $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M33 - $matrix.M11 - $matrix.M22); $qw = ($matrix.M21 - $matrix.M12) / $s; $qx = ($matrix.M13 + $matrix.M31) / $s; $qy = ($matrix.M23 + $matrix.M32) / $s; $qz = 0.25 * $s }
            $panelState.RotateTransform.Rotation = (New-Object System.Windows.Media.Media3D.QuaternionRotation3D([System.Windows.Media.Media3D.Quaternion]::new($qx, $qy, $qz, $qw)))
        }
    }

    function Animate-CurvedVortex {
        param($SyncHash)
        if ($SyncHash.Paused -or ($SyncHash.VisualizationStyle -ne "CurvedVortex")) { return }; $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime
        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier; $startRadius = 6.0; $endRadius = 1.0; $startY = 4.0; $endY = -6.0; $totalAngle = 720
        foreach ($key in $SyncHash.PlayerStates.Keys) {
            $panelState = $SyncHash.PlayerStates[$key]; $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle; $angle = $panelState.CurrentAngle; $angleRad = $angle * ([Math]::PI / 180.0); $progress = $angle / $totalAngle
            $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius)); $currentY = $startY - ($progress * ($startY - $endY)); $currentX = $currentRadius * [Math]::Cos($angleRad); $currentZ = $currentRadius * [Math]::Sin($angleRad)
            $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ; $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, $currentY - 1.5, 0); $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)
            $forward = $lookAtTarget - $position; if ($forward.Length -gt 0) { $forward.Normalize() }; $up = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0); $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward); if ($right.Length -gt 0) { $right.Normalize() }; $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)
            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity; $matrix.M11 = $right.X; $matrix.M12 = $right.Y; $matrix.M13 = $right.Z; $matrix.M21 = $newUp.X; $matrix.M22 = $newUp.Y; $matrix.M23 = $newUp.Z; $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z
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
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency; $SyncHash.LastFrameTime = $currentTime
        $xBoundary = 9; $yBoundary = 5; $zBoundary = 3
        foreach ($obj in $SyncHash.FloatingObjects) {
            $newX = $obj.Translate.OffsetX + ($obj.Velocity.X * $elapsed * $SyncHash.SpeedMultiplier); $newY = $obj.Translate.OffsetY + ($obj.Velocity.Y * $elapsed * $SyncHash.SpeedMultiplier); $newZ = $obj.Translate.OffsetZ + ($obj.Velocity.Z * $elapsed * $SyncHash.SpeedMultiplier)
            if (-not $obj.PSObject.Properties['CurrentRotation']) { continue }
            $velX = $obj.Velocity.X; $velY = $obj.Velocity.Y; $velZ = $obj.Velocity.Z
            if (($newX -gt $xBoundary -and $velX -gt 0) -or ($newX -lt -$xBoundary -and $velX -lt 0)) { $velX *= -1 }
            if (($newY -gt $yBoundary -and $velY -gt 0) -or ($newY -lt -$yBoundary -and $velY -lt 0)) { $velY *= -1 }
            if (($newZ -gt $zBoundary -and $velZ -gt 0) -or ($newZ -lt -$zBoundary -and $velZ -lt 0)) { $velZ *= -1 }
            $obj.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ); $obj.Translate.OffsetX = $newX; $obj.Translate.OffsetY = $newY; $obj.Translate.OffsetZ = $newZ
            $obj.CurrentRotation *= New-Object System.Windows.Media.Media3D.Quaternion((New-Object System.Windows.Media.Media3D.Vector3D(1,0,0)), ($obj.RotationVelocity.X * $elapsed * $SyncHash.SpeedMultiplier))
            $obj.CurrentRotation *= New-Object System.Windows.Media.Media3D.Quaternion((New-Object System.Windows.Media.Media3D.Vector3D(0,1,0)), ($obj.RotationVelocity.Y * $elapsed * $SyncHash.SpeedMultiplier))
            $obj.Rotate.Rotation = New-Object System.Windows.Media.Media3D.QuaternionRotation3D($obj.CurrentRotation)
        }
    }

    function Animate-FfmpegFrames {
        param($SyncHash)
        foreach ($key in $SyncHash.PlayerStates.Keys) {
            $playerState = $SyncHash.PlayerStates[$key]
            if ($playerState.IsImage -or -not $playerState.FfmpegProcess) { continue }
            try {
                $totalBytesRead = 0; $frameSize = $playerState.FrameBuffer.Length
                while ($totalBytesRead -lt $frameSize) {
                    $bytesRead = $playerState.FrameReader.Read($playerState.FrameBuffer, $totalBytesRead, $frameSize - $totalBytesRead)
                    if ($bytesRead -eq 0) { break }
                    $totalBytesRead += $bytesRead
                }
                if ($totalBytesRead -eq $frameSize) {
                    $dirtyRect = New-Object System.Windows.Int32Rect(0, 0, $playerState.FrameWidth, $playerState.FrameHeight)
                    $playerState.WriteableBitmap.WritePixels($dirtyRect, $playerState.FrameBuffer, $playerState.FrameStride, 0)
                } elseif ($playerState.FfmpegProcess.HasExited) {
                    $playerState.FrameReader.Dispose(); $playerState.FfmpegProcess.Dispose(); $playerState.FfmpegProcess = $null
                    Start-NextMedia -PlayerKey $key
                }
            } catch { Handle-MediaFailure -PlayerKey $key -Reason "Error reading frame from FFmpeg stream: $($_.Exception.Message)" }
        }
    }

    # --- Main Window Setup ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Master Visualizer"
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
                <StackPanel x:Name="scrollingPanel" VerticalAlignment="Top" HorizontalAlignment="Center" />
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
    $reader = New-Object System.Xml.XmlNodeReader $xaml; $window = [Windows.Markup.XamlReader]::Load($reader); $SyncHash.Window = $window
    $viewport = $window.FindName("mainViewport"); $visualHost = $window.FindName("VisualHost")
    if ($SyncHash.VisualizationStyle -eq "Aquarium" -and $SyncHash.AddWater) { $window.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString("#80ADD8E6")) }
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen; $window.Width = $primaryScreen.WorkingArea.Width; $window.Height = $primaryScreen.WorkingArea.Height; $window.Left = $primaryScreen.WorkingArea.Left; $window.Top = $primaryScreen.WorkingArea.Top

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
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton"); $SyncHash.hideControlsButton.Add_Click({ $controlsPanel = $window.FindName("controlsPanel"); $SyncHash.ControlsHidden = -not $SyncHash.ControlsHidden; $controlsPanel.Visibility = if ($SyncHash.ControlsHidden) { 'Collapsed' } else { 'Visible' } })
    $SyncHash.pauseButton = $window.FindName("pauseButton"); $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        $SyncHash.pauseButton.Content = if ($SyncHash.Paused) { "Resume" } else { "Pause" }
        if ($SyncHash.Paused) {
            if ($SyncHash.Storyboard) { $SyncHash.Storyboard.Pause($SyncHash.scrollingPanel) }
            elseif ($SyncHash.Animations) {
                foreach ($key in $SyncHash.Animations.Keys) {
                    if ($SyncHash.Rotations[$key]) { $rotation = $SyncHash.Rotations[$key]; $currentAngle = $rotation.Angle; $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $rotation.Angle = $currentAngle }
                    elseif ($SyncHash.Transforms[$key]) { $transform = $SyncHash.Transforms[$key]; $currentValue = $transform.ScaleX; $transform.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $null); $transform.ScaleX = $currentValue; $transform.ScaleY = $currentValue; $transform.ScaleZ = $currentValue }
                }
            }
        } else {
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
            if ($SyncHash.Storyboard) { $SyncHash.Storyboard.Resume($SyncHash.scrollingPanel) }
            elseif ($SyncHash.Animations) {
                foreach ($key in $SyncHash.Animations.Keys) {
                    if ($SyncHash.Rotations[$key]) { $SyncHash.Animations[$key].From = $SyncHash.Rotations[$key].Angle; $SyncHash.Rotations[$key].BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.Animations[$key]) }
                    elseif ($SyncHash.Transforms[$key]) { $transform = $SyncHash.Transforms[$key]; $transform.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $SyncHash.Animations[$key]); $transform.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $SyncHash.Animations[$key]); $transform.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $SyncHash.Animations[$key]) }
                }
            }
        }
    })
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton"); $SyncHash.randomAxisButton.Add_Click({ if ($SyncHash.Rotations) { $newAxis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0)); foreach ($key in $SyncHash.Rotations.Keys) { $SyncHash.Rotations[$key].Axis = $newAxis } } })
    $changeSpeed = { param($multiplier) $SyncHash.SpeedMultiplier /= $multiplier; if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }; if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }; if ($SyncHash.Animations) { foreach ($key in $SyncHash.Animations.Keys) { $baseDuration = switch($key) { "X" {20} "Y" {15} "Star" {30} "Pulse" {2} "Carousel" {60} "Scroll" {90} default {20} }; $SyncHash.Animations[$key].Duration = [TimeSpan]::FromSeconds($baseDuration / $SyncHash.SpeedMultiplier) } } }
    $SyncHash.slowDownButton = $window.FindName("slowDownButton"); $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton = $window.FindName("speedUpButton"); $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })
    $window.Add_KeyDown({ param($s, $e) switch ($e.Key) { "Escape" { $window.Close() } "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } "A" { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } } })
    $window.Add_Closed({ if ($animationLoop) { [System.Windows.Media.CompositionTarget]::remove_Rendering($animationLoop); $animationLoop = $null }; if ($SyncHash.WatchdogTimer) { $SyncHash.WatchdogTimer.Stop() }; if ($SyncHash.FlowTimer) { $SyncHash.FlowTimer.Stop() }; Cleanup-Visualization -SyncHash $SyncHash })
    $window.Add_Loaded({
        switch ($SyncHash.VisualizationStyle) {
            "Aquarium" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ButterflyEffect" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "Carousel" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ConcentricFunnel" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "CurvedVortex" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingCubes" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingSpheres" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "FloatingStars" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "MediaFlowFunnel" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "RollerCoaster" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "SphereVortex" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ReindeerSleigh" { $SyncHash.randomAxisButton.Visibility = 'Collapsed' }
            "ScrollingHorizontal" { $SyncHash.randomAxisButton.Visibility = 'Collapsed'; $SyncHash.slowDownButton.Visibility = 'Collapsed'; $SyncHash.speedUpButton.Visibility = 'Collapsed' }
            "ScrollingVertical" { $SyncHash.randomAxisButton.Visibility = 'Collapsed'; $SyncHash.slowDownButton.Visibility = 'Collapsed'; $SyncHash.speedUpButton.Visibility = 'Collapsed' }
        }
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B); $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor); $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily); $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }; $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }
            foreach ($key in $SyncHash.PlayerStates.Keys) { $textBlock = $SyncHash.PlayerStates[$key].OverlayTextBlock; if ($textBlock) { $textBlock.Foreground = $brush.Clone(); $textBlock.FontFamily = $fontFamily; $textBlock.FontSize = $SyncHash.FontSize; $textBlock.FontWeight = $fontWeight; $textBlock.FontStyle = $fontStyle; if ($SyncHash.RbSelection -eq "Custom") { $textBlock.Text = $SyncHash.CustomText } } }
        }
        if ($SyncHash.VisualizationStyle -eq "Aquarium") { $camera = $viewport.Camera; $distance = $camera.Position.Z; $fovRadians = 45.0 * ([Math]::PI / 180.0); $viewHeight3D = 2.0 * $distance * [Math]::Tan($fovRadians / 2.0); $aspectRatio = if ($viewport.ActualHeight -gt 0) { $viewport.ActualWidth / $viewport.ActualHeight } else { 1.6 }; $SyncHash.xBoundary = ((($viewHeight3D * $aspectRatio) / 2.0) / 2.0) * 1.2; $SyncHash.yBoundary = (($viewHeight3D / 2.0) / 2.0) * 1.2 }
        if ($SyncHash.VisualizationStyle -eq "MediaFlowFunnel") {
            $flowTimer = New-Object System.Windows.Threading.DispatcherTimer
            $flowTimer.Interval = [TimeSpan]::FromSeconds(2)

            $SyncHash.FlowTimer = $flowTimer
            $flowTimer.Add_Tick({
                if ($SyncHash.Paused) { return }
                
                $panelGroupsByRing = $SyncHash.PanelModels.GetEnumerator() | Group-Object { [math]::Floor($_.Name / $SyncHash.PanelsPerRing) } | Sort-Object { [int]$_.Name }
                
                # Shuffle both front and back materials in sync
                $lastFrontMaterial = $SyncHash.SharedMaterials[-1]
                $lastBackMaterial = $SyncHash.SharedBackMaterials[-1]
                for ($i = $SyncHash.SharedMaterials.Count - 1; $i -gt 0; $i--) {
                    $SyncHash.SharedMaterials[$i] = $SyncHash.SharedMaterials[$i-1]
                    $SyncHash.SharedBackMaterials[$i] = $SyncHash.SharedBackMaterials[$i-1]
                }
                $SyncHash.SharedMaterials[0] = $lastFrontMaterial
                $SyncHash.SharedBackMaterials[0] = $lastBackMaterial

                for ($i = 0; $i -lt $panelGroupsByRing.Count; $i++) {
                    $frontMaterialForThisRing = $SyncHash.SharedMaterials[$i % $SyncHash.SharedMaterials.Count]
                    $backMaterialForThisRing = $SyncHash.SharedBackMaterials[$i % $SyncHash.SharedBackMaterials.Count]
                    foreach ($panelModel in $panelGroupsByRing[$i].Group) {
                        # Assign materials directly, without cloning in the timer

                        $panelModel.Value.Material = $frontMaterialForThisRing;
                        $panelModel.Value.BackMaterial = $backMaterialForThisRing
                    }
                }
            })
            $flowTimer.Start()
        }
    })

    # --- Animation Loop Setup ---
    $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp(); $SyncHash.StartTime = $SyncHash.LastFrameTime
    $animationScriptBlock = {
        if ($SyncHash.UseFfmpeg) { Animate-FfmpegFrames -SyncHash $SyncHash }
        switch ($SyncHash.VisualizationStyle) {
            "FloatingSpheres" { Animate-FloatingObjects -SyncHash $SyncHash -Viewport $viewport }
            "FloatingStars"   { Animate-FloatingObjects -SyncHash $SyncHash -Viewport $viewport }
            "ButterflyEffect" { Animate-FloatingObjects -SyncHash $SyncHash -Viewport $viewport; Animate-ButterflyEffect -SyncHash $SyncHash }
            "RollerCoaster"   { Animate-RollerCoaster -SyncHash $SyncHash }
            "SphereVortex"    { Animate-SphereVortex -SyncHash $SyncHash }
            "CurvedVortex"    { Animate-CurvedVortex -SyncHash $SyncHash }
            "ReindeerSleigh"  { Animate-ReindeerSleigh -SyncHash $SyncHash }
            "Aquarium"        { Animate-Aquarium -SyncHash $SyncHash -Viewport $viewport }
        }
    }
    $animationLoop = [System.Windows.Media.CompositionTarget]::add_Rendering($animationScriptBlock.GetNewClosure())

    if ($SyncHash.VisualizationStyle -eq "Aquarium" -and $SyncHash.AddBubbles) {
        $bubbleCanvas = $window.FindName('bubbleCanvas'); $rand = [Random]::new(); $bubbles = New-Object System.Collections.Generic.List[object]; $maxCount = 220
        $newBubbleBrushFunc = { param([int]$alpha) $c1 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Min($alpha,255), 255, 255, 255); $c2 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Max($alpha-120,30), 173, 216, 230); $brush = New-Object Windows.Media.RadialGradientBrush; $brush.RadiusX = 0.6; $brush.RadiusY = 0.6; $brush.GradientOrigin = [Windows.Point]::new(0.35,0.35); $brush.Center = [Windows.Point]::new(0.5,0.5); $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c1,0.0)); $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c2,1.0)); return $brush }
        $newBubbleFunc = { $w = $bubbleCanvas.ActualWidth; $h = $bubbleCanvas.ActualHeight; if ($w -le 0 -or $h -le 0) { return }; $size = $rand.Next(5, 25); $speed = ($rand.NextDouble() * 1.4 + 0.6); $drift = (($rand.NextDouble()*2.0) - 1.0) * 0.35; $alpha = $rand.Next(120, 230); $startX = $rand.NextDouble() * [Math]::Max($w - $size, 1); $startY = $h - ($rand.NextDouble() * ([Math]::Max($h*0.15, 80))); $ellipse = New-Object Windows.Shapes.Ellipse; $ellipse.Width = $size; $ellipse.Height = $size; $ellipse.Fill = & $newBubbleBrushFunc -alpha $alpha; $ellipse.Stroke = [System.Windows.Media.Brushes]::White; $ellipse.StrokeThickness = [Math]::Max($size * 0.02, 0.6); [Windows.Controls.Canvas]::SetLeft($ellipse, $startX); [Windows.Controls.Canvas]::SetTop($ellipse, $startY); $bubbleCanvas.Children.Add($ellipse) | Out-Null; $bubble = [pscustomobject]@{ Shape = $ellipse; Vy = $speed; Vx = $drift; Spin = ($rand.NextDouble() * 0.04) - 0.02; T = $rand.NextDouble() * [Math]::PI }; $bubbles.Add($bubble) | Out-Null }
        $spawnTimer = New-Object Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(220) }; $spawnTimer.Add_Tick({ if ($bubbles.Count -lt $maxCount) { 1..($rand.Next(1,4)) | ForEach-Object { & $newBubbleFunc } } }); $SyncHash.BubbleSpawnTimer = $spawnTimer
        $animTimer = New-Object Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(16) }; $animTimer.Add_Tick({ if ($bubbleCanvas -eq $null -or $SyncHash.Paused) { return }; $h = $bubbleCanvas.ActualHeight; $w = $bubbleCanvas.ActualWidth; for ($i = $bubbles.Count - 1; $i -ge 0; $i--) { $b = $bubbles[$i]; $s = [double]$b.Shape.Width; $x = [Windows.Controls.Canvas]::GetLeft($b.Shape); $y = [Windows.Controls.Canvas]::GetTop($b.Shape); $b.T += $b.Spin; $x += $b.Vx + ([Math]::Sin($b.T) * 0.15); $y -= $b.Vy; if ($x -lt -10) { $x = -10; $b.Vx = [Math]::Abs($b.Vx) } elseif ($x + $s -gt $w + 10) { $x = $w + 10 - $s; $b.Vx = -[Math]::Abs($b.Vx) }; [Windows.Controls.Canvas]::SetLeft($b.Shape, $x); [Windows.Controls.Canvas]::SetTop($b.Shape, $y); if ($y + $s -lt 0) { $bubbleCanvas.Children.Remove($b.Shape); $bubbles.RemoveAt($i) } } }); $SyncHash.BubbleAnimTimer = $animTimer; $spawnTimer.Start(); $animTimer.Start()
    }

    if (-not $SyncHash.UseFfmpeg) {
        $watchdogTimer = New-Object System.Windows.Threading.DispatcherTimer; $watchdogTimer.Interval = [TimeSpan]::FromSeconds(3)
        $watchdogTimer.Add_Tick({
            if ($SyncHash.Paused) { return }
            $now = [datetime]::UtcNow
            foreach ($playerKey in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$playerKey]; if ($playerState.IsFailed) { continue }
                if ($playerState.SourceAssignmentTime -and ($now - $playerState.SourceAssignmentTime).TotalSeconds -gt 5) { Handle-MediaFailure -PlayerKey $playerKey -Reason "Watchdog: Media failed to open."; continue }
                if ($playerState.ExpectedDuration -and $playerState.PlaybackStopwatch.IsRunning -and $playerState.PlaybackStopwatch.Elapsed -gt ($playerState.ExpectedDuration + [TimeSpan]::FromSeconds(3))) { Handle-MediaFailure -PlayerKey $playerKey -Reason "Watchdog: Video playback is stuck."; continue }
            }
        }); $watchdogTimer.Start(); $SyncHash.WatchdogTimer = $watchdogTimer
    }
    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
