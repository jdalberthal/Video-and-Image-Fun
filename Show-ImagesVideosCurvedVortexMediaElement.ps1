<#
.SYNOPSIS
    Displays media on curved panels that spiral down a stationary 3D vortex.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto multiple
    3D panels that are geometrically curved. These panels are arranged in a descending spiral,
    creating the visual effect of a vortex or black hole.

    The funnel structure is stationary, but each curved panel individually travels down the spiral path,
    disappearing at the bottom and reappearing at the top for a continuous flow effect.
    This is achieved using per-frame animation calculations for precise control over the complex motion.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system (e.g., MP4, WMV, AVI).

.EXAMPLE
    PS C:\> .\Show-ImagesVideosCurvedVortexMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D vortex window.

.NOTES
    Name:           Show-ImagesVideosCurvedVortexMediaElement.ps1
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
$ExternalButtonName = "3D Curved Vortex`n(MediaElement)"
$ScriptDescription = "Displays media on curved panels that spiral down a stationary 3D vortex. Uses the built-in Windows MediaElement."

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "3D Curved Vortex (MediaElement) - Media Selector"
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
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            try {
                $currentStyle = [System.Drawing.FontStyle]::Regular
                if ($BoldCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Bold }
                if ($ItalicCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Italic }
                $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $currentStyle)
            } catch {
                $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12)
            }

            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
                $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
                $BoldCheckbox.Checked = $fontDialog.Font.Bold
                $ItalicCheckbox.Checked = $fontDialog.Font.Italic
                & $updateTextBoxFont
            }
        })

    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
        } catch {
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }
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
            SelectedFiles        = $formState.SelectedFiles
            UseTransparentEffect = $formState.UseTransparentEffect
            CurrentIndex         = -1
            PlayerStates         = [hashtable]::Synchronized(@{})
            Paused               = $false
            ControlsHidden       = $false
            RedoClicked          = $false
            LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
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
        Title="3D Curved Media Vortex"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,5,15" LookDirection="0,-0.3,-1" UpDirection="0,1,0" FieldOfView="70"/>
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

            <!-- This is the container for all the vortex panels -->
            <ModelVisual3D x:Name="VortexContainer" />

        </Viewport3D>
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="pauseButton" Content="Pause" Padding="10,5" Margin="2"/>
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
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")

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
        param($PanelIndex, [string]$Reason = "Unknown Error")
        
        $handleFailureScriptBlock = {
            param($pIndex, $reason)
            $pState = $SyncHash.PlayerStates[$pIndex]
            $fUri = $pState.CurrentSource
            $fName = if ($fUri) { [System.IO.Path]::GetFileName($fUri.LocalPath) } else { "an unknown file" }
            Write-Warning "Media failed for Panel $pIndex (File: '$fName'). Reason: $reason. Removing from list and continuing."
        }
        & $handleFailureScriptBlock -pIndex $PanelIndex -reason $Reason

        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState = $SyncHash.PlayerStates[$PanelIndex]
            if ($playerState.IsFailed) { return }
            $playerState.IsFailed = $true

            $playerState.ContentPresenter.Content = $null
            $playerState.ContentPresenterBack.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Black

            if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
                $SyncHash.SelectedFiles = @($SyncHash.SelectedFiles | Where-Object { $_ -ne $playerState.CurrentSource.LocalPath })
            }

            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{
                Interval = [TimeSpan]::FromMilliseconds(100); Tag = $PanelIndex
            }
            $recoveryTimer.Add_Tick({
                $timer = $args[0]; $pIndex = $timer.Tag; $timer.Stop()
                Start-NextMediaOnPanel -PanelIndex $pIndex
            })
            $playerState.RecoveryTimer = $recoveryTimer
            $recoveryTimer.Start()
        })
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

        $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
        $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Transparent

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
                $playerState.IsFailed = $false

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
                    if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500) {
                        HandleMediaFailure -PanelIndex $pIndex -Reason "Playback ended instantly (bad codec)."
                    } else {
                        $SyncHash.Window.Dispatcher.Invoke([action]{
                            Start-NextMediaOnPanel -PanelIndex $pIndex
                        })
                    }
                })
                $mediaElement.Add_MediaOpened({
                    $pIndex = $args[0].Tag
                    $pState = $SyncHash.PlayerStates[$pIndex]
                    $pState.IsFailed = $false
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
            } catch {
                HandleMediaFailure -PanelIndex $PanelIndex -Reason $_.Exception.Message
            }
        }
    }

    # --- Curved Panel Generation Function ---
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
                
                # Parabolic curve for the Z coordinate
                $normalizedX = $x / ($width / 2) # Normalize x from -1 to 1
                $z = $curveDepth * ($normalizedX * $normalizedX)

                $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
                $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $widthSegments, $j / $heightSegments))
            }
        }

        for ($j = 0; $j -lt $heightSegments; $j++) {
            for ($i = 0; $i -lt $widthSegments; $i++) {
                $row1 = $j * ($widthSegments + 1)
                $row2 = ($j + 1) * ($widthSegments + 1)

                $mesh.TriangleIndices.Add($row1 + $i)
                $mesh.TriangleIndices.Add($row1 + $i + 1)
                $mesh.TriangleIndices.Add($row2 + $i + 1)

                $mesh.TriangleIndices.Add($row1 + $i)
                $mesh.TriangleIndices.Add($row2 + $i + 1)
                $mesh.TriangleIndices.Add($row2 + $i)
            }
        }
        return $mesh
    }

    # --- Vortex Panel Generation ---
    $vortexContainer = $window.FindName("VortexContainer")
    $panelCount = 16

    $curvedPanelMesh = New-CurvedPanelMesh -width 2.5 -height 1.5 -curveDepth 0.5

    for ($i = 0; $i -lt $panelCount; $i++) {
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
        $panelViewport.Geometry = $curvedPanelMesh
        $panelViewport.Visual = $mediaHostGrid

        $panelViewportBack = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $panelViewportBack.Geometry = $curvedPanelMesh
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
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
        [void]$transformGroup.Children.Add($rotateTransform)
        [void]$transformGroup.Children.Add($translateTransform)
        $panelContainer.Transform = $transformGroup
        
        [void]$vortexContainer.Children.Add($panelContainer)
        
        $SyncHash.PlayerStates[$i] = @{
            ContentPresenter = $contentPresenter
            ContentPresenterBack = $contentPresenterBack
            OverlayTextBlock = $overlayTextBlock
            OverlayTextBlockBack = $overlayTextBlockBack
            MediaHostGrid = $mediaHostGrid
            MediaHostGridBack = $mediaHostGridBack
            TranslateTransform = $translateTransform
            RotateTransform = $rotateTransform
            CurrentAngle = (720.0 / $panelCount) * $i # Stagger the start angle across the full spiral
            MediaTimer = $null
            RecoveryTimer = $null
            CurrentMediaElement = $null
            CurrentMediaElementBack = $null
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
            IsFailed = $false
        }
    }

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
        } else {
            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        }
    })

    $SyncHash.slowDownButton.Add_Click({ $SyncHash.SpeedMultiplier *= 0.5 })
    $SyncHash.speedUpButton.Add_Click({ $SyncHash.SpeedMultiplier *= 2.0 })

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

    # --- Per-Frame Animation Handler ---
    $animationHandler = {
        param($sender, $e)

        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        # Vortex parameters
        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier # degrees per second
        $startRadius = 6.0
        $endRadius = 1.0
        $startY = 4.0
        $endY = -6.0
        $totalAngle = 360 * 2 # Two full rotations to go from top to bottom

        foreach ($i in 0..($panelCount-1)) {
            $panelState = $SyncHash.PlayerStates[$i]
            
            # Update and wrap the angle
            $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle
            $angle = $panelState.CurrentAngle
            
            # Convert to radians for trig functions
            $angleRad = $angle * ([Math]::PI / 180.0)

            # Interpolate radius and Y position based on the angle
            $progress = $angle / $totalAngle
            $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius))
            $currentY = $startY - ($progress * ($startY - $endY))

            # Calculate X and Z position on the circle
            $currentX = $currentRadius * [Math]::Cos($angleRad)
            $currentZ = $currentRadius * [Math]::Sin($angleRad)

            # Apply the translation
            $panelState.TranslateTransform.OffsetX = $currentX
            $panelState.TranslateTransform.OffsetY = $currentY
            $panelState.TranslateTransform.OffsetZ = $currentZ

            # Calculate rotation to make the panel face the center (Y-axis)
            # and also tilt it to match the funnel's slope
            $lookAtY = $currentY - 1.5
            [System.Windows.Media.Media3D.Point3D]$lookAtTarget = New-Object System.Windows.Media.Media3D.Point3D(0, $lookAtY, 0)
            [System.Windows.Media.Media3D.Point3D]$position = New-Object System.Windows.Media.Media3D.Point3D($currentX, $currentY, $currentZ)
            $forward = [System.Windows.Media.Media3D.Point3D]::Subtract($lookAtTarget, $position)
            $forward.Normalize()
            $up = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0)
            $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward)
            $right.Normalize()
            $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)

            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity
            $matrix.M11 = $right.X;   $matrix.M12 = $right.Y;   $matrix.M13 = $right.Z
            $matrix.M21 = $newUp.X;   $matrix.M22 = $newUp.Y;   $matrix.M23 = $newUp.Z
            $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z

            # Create a Quaternion from the rotation matrix
            $qw = 0; $qx = 0; $qy = 0; $qz = 0
            $trace = $matrix.M11 + $matrix.M22 + $matrix.M33

            if ($trace -gt 0) {
                $s = 0.5 / [Math]::Sqrt($trace + 1.0)
                $qw = 0.25 / $s; $qx = ($matrix.M32 - $matrix.M23) * $s
                $qy = ($matrix.M13 - $matrix.M31) * $s; $qz = ($matrix.M21 - $matrix.M12) * $s
            } elseif (($matrix.M11 -gt $matrix.M22) -and ($matrix.M11 -gt $matrix.M33)) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M11 - $matrix.M22 - $matrix.M33)
                $qw = ($matrix.M32 - $matrix.M23) / $s; $qx = 0.25 * $s
                $qy = ($matrix.M12 + $matrix.M21) / $s; $qz = ($matrix.M13 + $matrix.M31) / $s
            } elseif ($matrix.M22 -gt $matrix.M33) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M22 - $matrix.M11 - $matrix.M33)
                $qw = ($matrix.M13 - $matrix.M31) / $s; $qx = ($matrix.M12 + $matrix.M21) / $s
                $qy = 0.25 * $s; $qz = ($matrix.M23 + $matrix.M32) / $s
            } else {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M33 - $matrix.M11 - $matrix.M22)
                $qw = ($matrix.M21 - $matrix.M12) / $s; $qx = ($matrix.M13 + $matrix.M31) / $s
                $qy = ($matrix.M23 + $matrix.M32) / $s; $qz = 0.25 * $s
            }
            $quaternion = New-Object System.Windows.Media.Media3D.Quaternion($qx, $qy, $qz, $qw)
            $panelState.RotateTransform.Rotation = (New-Object System.Windows.Media.Media3D.QuaternionRotation3D($quaternion))
        }
    }

    # --- Window Events ---
    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $window.Add_Closed({
        [System.Windows.Media.CompositionTarget]::remove_Rendering($animationHandler)
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

        # Start the animation loop
        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        [System.Windows.Media.CompositionTarget]::add_Rendering($animationHandler)
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}