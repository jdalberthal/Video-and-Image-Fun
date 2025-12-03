<#
.SYNOPSIS
    Displays media on a series of vertical panels arranged in a rotating 3D carousel with an undulating motion.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto multiple,
    vertical 3D panels. These panels are arranged in a circular carousel that rotates around a
    central axis.

    Each panel also moves up and down in a smooth, oscillating wave pattern, creating a dynamic
    and mesmerizing visual effect.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system (e.g., MP4, WMV, AVI).

.EXAMPLE
    PS C:\> .\Show-ImagesVideosCarouselMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D carousel window.

.NOTES
    Name:           Show-ImagesVideosCarouselMediaElement.ps1
    Version:        1.0.0, 11/23/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "3D Media Carousel`n(MediaElement)"
$ScriptDescription = "Displays media on vertical panels in a rotating carousel with an undulating wave motion. Uses the built-in Windows MediaElement."
$RequiredExecutables = @() # No external executables needed

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "3D Carousel (MediaElement) - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25'
    }
    $SelectForm.Controls.Add($BrowseButton)

    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '120, 10'; Size = '450, 25'; ReadOnly = $true
    }
    $SelectForm.Controls.Add($FolderPathTextBox)

    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false
    }
    $SelectForm.Controls.Add($RecursiveCheckBox)

    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false
    }
    $SelectForm.Controls.Add($TransparentCheckbox)

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false
    }
    $SelectForm.Controls.Add($SelectAllCheckbox)

    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'
        AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65
    }
    $SelectForm.Controls.Add($DataGridView)

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{
        Name = "Select"; HeaderText = ""; Width = 30
    }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null

    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null

    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Play Selected"; Location = '600, 40'; Size = '170, 30'
    }
    $SelectForm.Controls.Add($PlayButton)

    # --- Text Overlay Controls ---
    $GroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 440'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $GroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $SelectForm.Controls.Add($GroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '140, 440'; Size = '455, 180'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center'
    }
    $SelectForm.Controls.Add($TextBox)

    $CurrentColor = New-Object System.Windows.Forms.Label -Property @{ Text = "Text Color:"; Location = '600, 477'; AutoSize = $true; Visible = $false }
    $ColorExample = New-Object System.Windows.Forms.Label -Property @{ Text = "     "; Location = '660, 477'; AutoSize = $true; BackColor = [System.Drawing.Color]::Black; Visible = $false }
    $SelectColorButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change"; Location = '685, 470'; Size = '80, 30'; Visible = $false }
    $SizeLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Font Size:"; AutoSize = $true; Location = '600, 522'; Visible = $false }
    $NumericUpDown = New-Object System.Windows.Forms.NumericUpDown -Property @{ Location = '660, 520'; Size = '50, 20'; Visible = $false; Minimum = 8; Maximum = 72; Value = 24 }
    $FontButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change Font"; Location = '600, 570'; Size = '170, 25'; Visible = $false }
    $ItalicCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Italic"; Location = '600, 620'; Size = '75, 20'; Checked = $false; Visible = $false }
    $BoldCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Bold"; Location = '680, 620'; Size = '75, 20'; Checked = $true; Visible = $false }

    $SelectForm.Controls.AddRange(@(
            $CurrentColor, $ColorExample, $SelectColorButton, $SizeLabel,
            $NumericUpDown, $FontButton, $ItalicCheckbox, $BoldCheckbox
        ))

    $textOverlayEvent = {
        $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked
        $isCustomText = $RadioButton3.Checked
        $TextBox.Visible = $isCustomText
        @($CurrentColor, $ColorExample, $SelectColorButton, $SizeLabel, $NumericUpDown, $FontButton, $ItalicCheckbox, $BoldCheckbox).ForEach({ $_.Visible = $isTextVisible })
    }
    $RadioButton1.Add_Click($textOverlayEvent)
    $RadioButton2.Add_Click($textOverlayEvent)
    $RadioButton3.Add_Click($textOverlayEvent)

    $formState = @{ TextColor = [System.Drawing.Color]::Black; FontFamily = "Arial" }
    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
            $colorDialog = New-Object System.Windows.Forms.ColorDialog
            if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.TextColor = $colorDialog.Color
                $ColorExample.BackColor = $formState.TextColor
                $TextBox.ForeColor = $formState.TextColor # Sync the textbox color
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            # Initialize the dialog with the current settings from the form
            try {
                $currentStyle = [System.Drawing.FontStyle]::Regular
                if ($BoldCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Bold }
                if ($ItalicCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Italic }
                $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $currentStyle)
            } catch {
                $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12)
            }

            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                # When the dialog closes, update all related controls on the main form
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
                $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
                $BoldCheckbox.Checked = $fontDialog.Font.Bold
                $ItalicCheckbox.Checked = $fontDialog.Font.Italic
                # Trigger the central font update function
                & $updateTextBoxFont
            }
        })

    # Central script block to update the TextBox font based on the state of all controls
    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
        } catch {
            # Fallback to a default font if the selected one is invalid
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont)
    $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont)
    $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)
    
    # Call the update function once at the start to synchronize the initial state
    & $updateTextBoxFont

    $SelectAllCheckbox.Add_CheckedChanged({
            $isChecked = $SelectAllCheckbox.Checked
            foreach ($row in $DataGridView.Rows) { $row.Cells["Select"].Value = $isChecked }
            $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        })

    $BrowseButton.Add_Click({
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $FolderBrowser.Description = "Select the folder to scan."
            if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $SelectedPath = $FolderBrowser.SelectedPath
                $FolderPathTextBox.Text = $SelectedPath
                $DataGridView.Rows.Clear()

                $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
                $VideoExtensions = "*.mp4", "*.m4v", "*.wmv", "*.avi", "*.mpg", "*.mpeg"
                $AllowedExtensions = $ImageExtensions + $VideoExtensions
            
                $gciParams = @{ File = $true; Include = $AllowedExtensions }
                if ($RecursiveCheckBox.Checked) {
                    $gciParams.Path = $SelectedPath
                    $gciParams.Recurse = $true
                } else {
                    $gciParams.Path = Join-Path $SelectedPath "*"
                }
                $files = Get-ChildItem @gciParams
                foreach ($file in $files) {
                    $DataGridView.Rows.Add($false, $file.Name, $file.FullName)
                }

                foreach ($row in $DataGridView.Rows) {
                    if ($row.IsNewRow) { continue }
                    $row.HeaderCell.Value = "Play"
                }
            }
        })
    
    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $row = $DataGridView.Rows[$e.RowIndex]
            $filePath = $row.Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                # Using Start-Process to open with default player. ffplay might not be available.
                try { Start-Process $filePath } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", "OK", "Error") }
            } else { [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", "OK", "Error") }
        }
    })

    $PlayButton.Add_Click({
            $formState.SelectedFiles = @(
                foreach ($Row in $DataGridView.Rows) {
                    if ($Row.Cells["Select"].Value) { $Row.Cells["FilePath"].Value }
                }
            )
            if ($formState.SelectedFiles.Count -gt 0) {
                $formState.UseTransparentEffect = $TransparentCheckbox.Checked
                if ($RadioButton1.Checked) { $formState.RbSelection = "Hidden" }
                if ($RadioButton2.Checked) { $formState.RbSelection = "Filename" }
                if ($RadioButton3.Checked) { $formState.RbSelection = "Custom" }
                $formState.CustomText = $TextBox.Text
                $formState.FontSize = $NumericUpDown.Value
                $formState.IsBold = $BoldCheckbox.Checked
                $formState.IsItalic = $ItalicCheckbox.Checked
                $SelectForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", "OK", "Warning")
            }
        })

    $null = $SelectForm.ShowDialog()
    $SelectForm.Dispose()

    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break 
    }

    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
            SelectedFiles        = [System.Collections.ArrayList]::new($formState.SelectedFiles)
            UseTransparentEffect = $formState.UseTransparentEffect
            CurrentIndex         = -1
            PlayerStates         = [hashtable]::Synchronized(@{})
            Paused               = $false
            ControlsHidden       = $false
            RedoClicked          = $false
            # Text Overlay Settings
            RbSelection          = $formState.RbSelection
            CustomText           = $formState.CustomText
            TextColor            = $formState.TextColor
            FontSize             = $formState.FontSize
            FontFamily           = $formState.FontFamily
            IsBold               = $formState.IsBold
            IsItalic             = $formState.IsItalic
        })
    $SyncHash.SpeedMultiplier = 1.0 # Initialize speed multiplier

    # --- XAML Definition ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="3D Media Carousel"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,12" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
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

            <!-- This is the container for all the carousel panels -->
            <ModelVisual3D x:Name="CarouselContainer">
                <ModelVisual3D.Transform>
                    <RotateTransform3D>
                        <RotateTransform3D.Rotation>
                            <AxisAngleRotation3D x:Name="CarouselRotation" Axis="0,1,0" Angle="0"/>
                        </RotateTransform3D.Rotation>
                    </RotateTransform3D>
                </ModelVisual3D.Transform>
            </ModelVisual3D>

        </Viewport3D>
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

    # --- Set Window to Full Screen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    # --- Find UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $SyncHash.CarouselRotation = $window.FindName("CarouselRotation")

    # --- Centralized, Thread-Safe Indexing ---
    $globalIndexLock = New-Object object
    function Get-NextMediaIndex {
        [System.Threading.Monitor]::Enter($globalIndexLock)
        try {
            if ($SyncHash.SelectedFiles.Count -eq 0) { return -1 }
            $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $SyncHash.SelectedFiles.Count
            return $SyncHash.CurrentIndex
        } finally {
            [System.Threading.Monitor]::Exit($globalIndexLock)
        }
    }

    # --- Media Handling Functions ---
    function HandleMediaFailure {
        param([int]$PanelIndex, [string]$Reason)

        $playerState = $SyncHash.PlayerStates[$PanelIndex]
        if ($playerState.IsFailed) { return } # Prevent re-entry

        $playerState.IsFailed = $true

        # 1. Log directly to the console.
        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for Panel $PanelIndex (File: '$fileName'). Reason: $Reason. Attempting to replace."

        # 2. Remove the bad file from the list.
        if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
            $SyncHash.SelectedFiles.Remove($playerState.CurrentSource.LocalPath) | Out-Null
        }

        # 3. Dispatch ONLY the pure UI updates (blacking out the panel).
        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState.ContentPresenter.Content = $null
            $playerState.ContentPresenterBack.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Black
        })

        # 4. Immediately call for the next media file, OUTSIDE the dispatcher. This is the critical step.
        # Use a short-lived timer to avoid potential re-entrancy issues within the event handler.
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(50); Tag = $PanelIndex }
        $recoveryTimer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaOnPanel -PanelIndex $idx })
        $playerState.RecoveryTimer = $recoveryTimer
        $recoveryTimer.Start()
    }

    function Start-NextMediaOnPanel {
        param([int]$PanelIndex)
        
        $playerState = $SyncHash.PlayerStates[$PanelIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
        if ($playerState.CurrentMediaElementBack) { $playerState.CurrentMediaElementBack.Close() }
        $playerState.CurrentMediaElement = $null
        $playerState.CurrentMediaElementBack = $null

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false # Reset failure flag for the new attempt

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
            $playerState.OverlayTextBlockBack.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        if ($ImageExtensions -contains $extension) {
            try {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit()
                
                $imageFront = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $imageBack = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $imageFront
                $playerState.ContentPresenterBack.Content = $imageBack

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(10)
                $timer.Tag = $PanelIndex
                $timer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaOnPanel -PanelIndex $idx })
                $playerState.MediaTimer = $timer
                $timer.Start()
            } catch {
                HandleMediaFailure -PanelIndex $PanelIndex -Reason "Failed to load image file."
            }
        }
        else { # It's a video, use MediaElement
            try {
                # --- Front Face Player ---
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource
                }
                $mediaElement.Tag = $PanelIndex
                $mediaElement.Add_MediaEnded({ 
                    $pIndex = $args[0].Tag
                    $pState = $SyncHash.PlayerStates[$pIndex]
                    if ($pState.IsFailed) { return }
                    $pState.PlaybackStopwatch.Stop()

                    # This is the critical check for silent failures.
                    if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500) {
                        HandleMediaFailure -PanelIndex $pIndex -Reason "Playback ended instantly (bad codec)."
                        return
                    }

                    Start-NextMediaOnPanel -PanelIndex $pIndex
                })
                $mediaElement.Add_MediaOpened({
                    $pIndex = $args[0].Tag
                    $pState = $SyncHash.PlayerStates[$pIndex]
                    $pState.PlaybackStopwatch.Restart()
                    if (-not $args[0].NaturalDuration.HasTimeSpan) {
                        HandleMediaFailure -PanelIndex $pIndex -Reason "Invalid duration or codec."
                    }
                })
                $mediaElement.Add_MediaFailed({ HandleMediaFailure -PanelIndex $args[0].Tag -Reason $args[1].ErrorException.Message })
                
                # --- Back Face Player (for sync) ---
                $mediaElementBack = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource
                }
                # The back face just loops on its own; its events do not trigger the next file.
                $mediaElementBack.Add_MediaEnded({ $thisElement = $args[0]; $thisElement.Position = [TimeSpan]::Zero; $thisElement.Play() })
                
                # Assign to presenters and start playback
                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.ContentPresenterBack.Content = $mediaElementBack
                $playerState.CurrentMediaElement = $mediaElement
                $playerState.CurrentMediaElementBack = $mediaElementBack
                
                $mediaElement.Play()
                $mediaElementBack.Play()
                # Blackout the panel until MediaOpened confirms the media is valid.
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
                $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Black

            } catch {
                HandleMediaFailure -PanelIndex $PanelIndex -Reason $_.Exception.Message
            }
        }
    }

    # --- Carousel Panel Generation ---
    $carouselContainer = $window.FindName("CarouselContainer")
    $panelCount = 8
    $angleIncrement = 360 / $panelCount
    $panelWidth = 3.0
    $panelHeight = 5.0
    $carouselRadius = 4.0

    $panelMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $panelMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$panelWidth/2), (-$panelHeight/2), 0)) )
    $panelMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($panelWidth/2), (-$panelHeight/2), 0)) )
    $panelMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($panelWidth/2),  ($panelHeight/2), 0)) )
    $panelMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$panelWidth/2),  ($panelHeight/2), 0)) )
    $panelMesh.TriangleIndices.Add(0); $panelMesh.TriangleIndices.Add(1); $panelMesh.TriangleIndices.Add(2)
    $panelMesh.TriangleIndices.Add(0); $panelMesh.TriangleIndices.Add(2); $panelMesh.TriangleIndices.Add(3)
    $panelMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,1)) )
    $panelMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,1)) )
    $panelMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,0)) )
    $panelMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,0)) )

    for ($i = 0; $i -lt $panelCount; $i++) {
        $panelAngle = $i * $angleIncrement

        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        [void]$mediaHostGrid.Children.Add($contentPresenter)

        $mediaHostGridBack = New-Object System.Windows.Controls.Grid
        $contentPresenterBack = New-Object System.Windows.Controls.ContentPresenter
        [void]$mediaHostGridBack.Children.Add($contentPresenterBack)

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'
            TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        [void]$mediaHostGrid.Children.Add($overlayTextBlock)

        $overlayTextBlockBack = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'
            TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        [void]$mediaHostGridBack.Children.Add($overlayTextBlockBack)

        $panelViewport = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $panelViewport.Geometry = $panelMesh
        $panelViewport.Visual = $mediaHostGrid

        $panelViewportBack = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $panelViewportBack.Geometry = $panelMesh
        $panelViewportBack.Visual = $mediaHostGridBack
        $panelViewportBack.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180)))

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $panelMaterial = New-Object $materialType
        [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($panelMaterial, $true)
        $panelViewport.Material = $panelMaterial

        $panelMaterialBack = New-Object $materialType
        [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($panelMaterialBack, $true)
        $panelViewportBack.Material = $panelMaterialBack

        $panelContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        [void]$panelContainer.Children.Add($panelViewport)
        [void]$panelContainer.Children.Add($panelViewportBack)

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup

        # 1. Move the panel out from the center along the Z-axis
        [void]$transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$carouselRadius)))

        # 2. Rotate the panel into its position in the carousel circle (around the Y-axis)
        $placementRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{
            Axis = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0); Angle = $panelAngle
        }
        [void]$transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($placementRotation)))

        # 3. Add the vertical translation for the undulating motion
        $verticalTranslate = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        [void]$transformGroup.Children.Add($verticalTranslate)

        $panelContainer.Transform = $transformGroup
        [void]$carouselContainer.Children.Add($panelContainer)

        # --- Animations ---
        # Vertical oscillation animation
        $verticalAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
            From = -0.5; To = 0.5; Duration = [TimeSpan]::FromSeconds(4); 
            AutoReverse = $true; RepeatBehavior = "Forever";
            BeginTime = [TimeSpan]::FromSeconds($i * 0.5) # Stagger the start time for the wave effect
        }
        $verticalTranslate.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $verticalAnim)
        
        $SyncHash.PlayerStates[$i] = @{
            ContentPresenter = $contentPresenter
            ContentPresenterBack = $contentPresenterBack
            OverlayTextBlock = $overlayTextBlock
            OverlayTextBlockBack = $overlayTextBlockBack
            MediaHostGrid = $mediaHostGrid
            MediaHostGridBack = $mediaHostGridBack
            VerticalAnimation = $verticalAnim # Store for speed changes
            MediaTimer = $null
            RecoveryTimer = $null
            CurrentMediaElement = $null # This property is required by Start-NextMediaOnPanel
            CurrentMediaElementBack = $null # This property is required by Start-NextMediaOnPanel
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
            IsFailed = $false
        }
    }

    # --- Main Carousel Animation ---
    $carouselAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
        From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(60); RepeatBehavior = "Forever"
    }
    $SyncHash.CarouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $carouselAnim)
    $SyncHash.CarouselAnim = $carouselAnim

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        if ($SyncHash.Paused) {
            $SyncHash.CarouselAnim.From = $SyncHash.CarouselRotation.Angle
            $SyncHash.CarouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.CarouselAnim)
            
            # No individual panel rotation to resume, but we could resume the vertical animation if we paused it.
            # For now, we'll let the vertical animation continue during pause for a nice effect.

            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.Paused = $false
        } else {
            $SyncHash.CarouselRotation.Angle = $SyncHash.CarouselRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
            $SyncHash.CarouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)

            # We'll let the vertical oscillation continue during pause.

            $SyncHash.pauseButton.Content = "Resume"
            $SyncHash.Paused = $true
        }
    })

    $SyncHash.randomAxisButton.Add_Click({
        $SyncHash.CarouselRotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
    })

    $changeSpeed = {
        param($multiplier)

        $wasPaused = $SyncHash.Paused
        
        if (-not $wasPaused) {
            $SyncHash.CarouselRotation.Angle = $SyncHash.CarouselRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
            $SyncHash.CarouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        }

        $SyncHash.SpeedMultiplier *= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        
        # Update durations based on base values
        $SyncHash.CarouselAnim.Duration = [TimeSpan]::FromSeconds(60.0 / $SyncHash.SpeedMultiplier)
        0..($panelCount-1) | ForEach-Object { $SyncHash.PlayerStates[$_].VerticalAnimation.Duration = [TimeSpan]::FromSeconds(4.0 / $SyncHash.SpeedMultiplier) }

        if (-not $wasPaused) {
            $SyncHash.CarouselAnim.From = $SyncHash.CarouselRotation.Angle
            $SyncHash.CarouselRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.CarouselAnim)
        }
    }

    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 0.5 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 2.0 })

    $SyncHash.redoButton.Add_Click({
        $SyncHash.RedoClicked = $true
        $window.Close()
    })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        if ($SyncHash.ControlsHidden) {
            $controlsPanel.Visibility = 'Visible'
            $SyncHash.ControlsHidden = $false
        } else {
            $controlsPanel.Visibility = 'Collapsed'
            $SyncHash.ControlsHidden = $true
        }
    })

    # --- Window Events ---
    $window.Add_KeyDown({
        param($sender, $e)
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
        if ($SyncHash.PlayerStates) {
            foreach ($i in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$i]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
                if ($playerState.CurrentMediaElementBack) { $playerState.CurrentMediaElementBack.Close() }
            }
        }
    })

    # --- Start the show ---
    $window.Add_Loaded({
        # Apply Text Overlay Settings
        $textToSet = $null
        if ($SyncHash.RbSelection -eq "Custom") {
            $textToSet = $SyncHash.CustomText
        }

        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            for ($i = 0; $i -lt $panelCount; $i++) {
                $tbFront = $SyncHash.PlayerStates[$i].OverlayTextBlock
                $tbBack = $SyncHash.PlayerStates[$i].OverlayTextBlockBack
                foreach ($tb in @($tbFront, $tbBack)) {
                    $tb.Foreground = $brush.Clone()
                    $tb.FontFamily = $fontFamily
                    $tb.FontSize = $SyncHash.FontSize
                    $tb.FontWeight = $fontWeight
                    $tb.FontStyle = $fontStyle
                    if ($textToSet) { $tb.Text = $textToSet }
                }
            }
        }

        # Start media on each panel
        for ($i = 0; $i -lt $panelCount; $i++) {
            Start-NextMediaOnPanel -PanelIndex $i
        }
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}