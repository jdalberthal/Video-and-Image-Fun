<#
.SYNOPSIS
    Displays media on spheres that spiral down a rotating 3D vortex or funnel.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto multiple
    3D spheres. These spheres are arranged in a descending spiral, creating the visual effect of
    a vortex or a black hole pulling in planets.

    The entire structure rotates, and each sphere individually travels down the spiral path,
    disappearing at the bottom and reappearing at the top for a continuous flow effect.
    This is achieved using per-frame animation calculations for precise control over the complex motion.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system (e.g., MP4, WMV, AVI).

.EXAMPLE
    PS C:\> .\Show-ImagesVideosSphereVortexMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D vortex window.

.NOTES
    Name:           Show-ImagesVideosSphereVortexMediaElement.ps1
    Version:        1.0.0, 11/26/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "3D Vortex (MediaElement) - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{ Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'; AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65 }
    
    $SelectForm.Controls.AddRange(@($BrowseButton, $FolderPathTextBox, $RecursiveCheckBox, $TransparentCheckbox, $SelectAllCheckbox, $DataGridView))

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true }
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Play Selected"; Location = '600, 40'; Size = '170, 30' }
    $SelectForm.Controls.Add($PlayButton)

    # --- Text Overlay Controls ---
    $GroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 440'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $GroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $SelectForm.Controls.Add($GroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '140, 440'; Size = '455, 180'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center' }
    $CurrentColor = New-Object System.Windows.Forms.Label -Property @{ Text = "Text Color:"; Location = '600, 477'; AutoSize = $true; Visible = $false }
    $ColorExample = New-Object System.Windows.Forms.Label -Property @{ Text = "     "; Location = '660, 477'; AutoSize = $true; BackColor = [System.Drawing.Color]::Black; Visible = $false }
    $SelectColorButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change"; Location = '685, 470'; Size = '80, 30'; Visible = $false }
    $SizeLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Font Size:"; AutoSize = $true; Location = '600, 522'; Visible = $false }
    $NumericUpDown = New-Object System.Windows.Forms.NumericUpDown -Property @{ Location = '660, 520'; Size = '50, 20'; Visible = $false; Minimum = 8; Maximum = 72; Value = 24 }
    $FontButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Change Font"; Location = '600, 570'; Size = '170, 25'; Visible = $false }
    $ItalicCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Italic"; Location = '600, 620'; Size = '75, 20'; Checked = $false; Visible = $false }
    $BoldCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Bold"; Location = '680, 620'; Size = '75, 20'; Checked = $true; Visible = $false }

    $SelectForm.Controls.AddRange(@($TextBox, $CurrentColor, $ColorExample, $SelectColorButton, $SizeLabel, $NumericUpDown, $FontButton, $ItalicCheckbox, $BoldCheckbox))

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
    
    $OrientationGroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Vortex Orientation"; Location = '300, 40'; Size = '200, 50' }
    $VortexVerticalRadioButton = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Vertical"; Location = '15, 20'; Checked = $true; AutoSize = $true }
    $VortexHorizontalRadioButton = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Horizontal"; Location = '100, 20'; AutoSize = $true }
    $OrientationGroupBox.Controls.AddRange(@($VortexVerticalRadioButton, $VortexHorizontalRadioButton))
    $SelectForm.Controls.Add($OrientationGroupBox)

    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
            $colorDialog = New-Object System.Windows.Forms.ColorDialog
            if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.TextColor = $colorDialog.Color
                $ColorExample.BackColor = $formState.TextColor
                & $updateTextBoxFont
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            $fontDialog.ShowColor = $true
            try {
                $currentStyle = [System.Drawing.FontStyle]::Regular
                if ($BoldCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Bold }
                if ($ItalicCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Italic }
                $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $currentStyle)
                $fontDialog.Color = $formState.TextColor
            } catch {
                $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12)
            }

            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
                $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
                $BoldCheckbox.Checked = $fontDialog.Font.Bold
                $ItalicCheckbox.Checked = $fontDialog.Font.Italic
                $formState.TextColor = $fontDialog.Color
                $ColorExample.BackColor = $formState.TextColor
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
            $TextBox.ForeColor = $formState.TextColor
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
                if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true } 
                else { $gciParams.Path = Join-Path $SelectedPath "*" }
                
                Get-ChildItem @gciParams | ForEach-Object { $DataGridView.Rows.Add($false, $_.Name, $_.FullName) }
                $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
            }
        })
    
    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $filePath = $DataGridView.Rows[$e.RowIndex].Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                try { Start-Process $filePath } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", "OK", "Error") }
            } else { [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", "OK", "Error") }
        }
    })

    $PlayButton.Add_Click({
            $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
            if ($selectedFiles.Count -gt 0) {                
                $formState.UseTransparentEffect = $TransparentCheckbox.Checked
                $formState.SelectedFiles = [System.Collections.ArrayList]::new($selectedFiles)
                if ($RadioButton1.Checked) { $formState.RbSelection = "Hidden" }
                if ($RadioButton2.Checked) { $formState.RbSelection = "Filename" }
                if ($RadioButton3.Checked) { $formState.RbSelection = "Custom" }
                $formState.CustomText = $TextBox.Text
                $formState.VortexOrientation = if ($VortexVerticalRadioButton.Checked) { "Vertical" } else { "Horizontal" }
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
            SphereStates         = [hashtable]::Synchronized(@{}) # For animation
            MediaGroupStates     = [hashtable]::Synchronized(@{}) # For media players
            Paused               = $false
            ControlsHidden       = $false
            RedoClicked          = $false
            LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
            VortexOrientation    = $formState.VortexOrientation
            RbSelection          = $formState.RbSelection
            CustomText           = $formState.CustomText
            TextColor            = $formState.TextColor
            FontSize             = $formState.FontSize
            FontFamily           = $formState.FontFamily
            IsBold               = $formState.IsBold
            IsItalic             = $formState.IsItalic
        })
    $SyncHash.SpeedMultiplier = 1.0

    # --- XAML Definition ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="3D Media Vortex"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,15" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="70"/>
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

            <ModelVisual3D x:Name="VortexContainer">
                <ModelVisual3D.Transform>
                    <RotateTransform3D>
                        <RotateTransform3D.Rotation>
                            <AxisAngleRotation3D x:Name="VortexMasterRotation" Axis="0,1,0" Angle="0"/>
                        </RotateTransform3D.Rotation>
                    </RotateTransform3D>
                </ModelVisual3D.Transform>
            </ModelVisual3D>
        </Viewport3D>
        <Canvas x:Name="VisualHost" Opacity="0"/>
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

    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width; $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left; $window.Top = $primaryScreen.WorkingArea.Top

    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $SyncHash.VortexMasterRotation = $window.FindName("VortexMasterRotation")

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

    # --- Media Handling Functions (Adopted from MediaFlowFunnel script for proven stability) ---
    function Handle-MediaFailure {
        param([int]$GroupIndex, [string]$Reason)

        $playerState = $SyncHash.MediaGroupStates[$GroupIndex]
        if ($playerState.IsFailed) { return } # Prevent re-entry

        $playerState.IsFailed = $true

        # 1. Log directly to the console.
        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for Group $GroupIndex (File: '$fileName'). Reason: $Reason. Attempting to replace."

        # 2. Remove the bad file from the list.
        if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
            $SyncHash.SelectedFiles.Remove($playerState.CurrentSource.LocalPath) | Out-Null
        }

        # 3. Dispatch ONLY the pure UI updates (blacking out the panel).
        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState.ContentPresenter.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
        })

        # 4. Immediately call for the next media file.
        Start-NextMediaForGroup -GroupIndex $GroupIndex
    }

    function Start-NextMediaForGroup {
        param([int]$GroupIndex)

        $playerState = $SyncHash.MediaGroupStates[$GroupIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
        $playerState.CurrentMediaElement = $null
        $playerState.CurrentSource = $null

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { 
            $playerState.ContentPresenter.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            return
        }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false # Reset failure flag for the new attempt

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

        try {
            if ($ImageExtensions -contains $extension) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit(); $bitmap.Freeze()
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                $playerState.ContentPresenter.Content = $image

                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $GroupIndex }
                $timer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaForGroup -GroupIndex $idx })
                $playerState.MediaTimer = $timer
                $timer.Start()
            }
            else { # Video
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource
                }
                $mediaElement.Tag = $GroupIndex

                $mediaElement.Add_MediaEnded({ 
                    $gIndex = $args[0].Tag
                    $pState = $SyncHash.MediaGroupStates[$gIndex]
                    if ($pState.IsFailed) { return }
                    $pState.PlaybackStopwatch.Stop()

                    if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500) {
                        Handle-MediaFailure -GroupIndex $gIndex -Reason "Playback ended instantly (bad codec or invalid file)"
                        return
                    }
                    Start-NextMediaForGroup -GroupIndex $gIndex
                })

                $mediaElement.Add_MediaOpened({
                    $gIndex = $args[0].Tag
                    $pState = $SyncHash.MediaGroupStates[$gIndex]
                    if (-not $args[0].NaturalDuration.HasTimeSpan) {
                        Handle-MediaFailure -GroupIndex $gIndex -Reason "Invalid duration or codec"
                    } else {
                        $pState.PlaybackStopwatch.Restart()
                        $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                    }
                })

                $mediaElement.Add_MediaFailed({ Handle-MediaFailure -GroupIndex $args[0].Tag -Reason $args[1].ErrorException.Message })
                
                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.CurrentMediaElement = $mediaElement
                $mediaElement.Play()
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            }
        } catch {
            # Fallback catch
        }
    }

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

    # --- Vortex Sphere Generation ---
    $vortexContainer = $window.FindName("VortexContainer")
    $sphereCount = 32
    $numberOfGroups = 6 # Number of unique media players
    $sphereMesh = New-SphereMesh -radius 1.0

    # Create the 6 unique media hosts and materials
    $sharedMaterials = @()
    for ($i = 0; $i -lt $numberOfGroups; $i++) {
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush; Color = [System.Windows.Media.Colors]::White }

        $SyncHash.MediaGroupStates[$i] = @{
            MediaHostGrid = $mediaHostGrid; ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock
            MediaTimer = $null; RecoveryTimer = $null; CurrentMediaElement = $null; CurrentSource = $null
            IsFailed = $false; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
        }
    }
    # The Grids must be part of the live visual tree for the VisualBrushes to work.
    $visualHostCanvas = $window.FindName("VisualHost")
    $SyncHash.MediaGroupStates.GetEnumerator() | ForEach-Object { $visualHostCanvas.Children.Add($_.Value.MediaHostGrid) | Out-Null }

    # Create the 32 sphere models
    for ($i = 0; $i -lt $sphereCount; $i++) {
        $materialToUse = $sharedMaterials[$i % $numberOfGroups]
        $sphereMaterial = $materialToUse.Clone()

        $sphereGeometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $sphereMesh; Material = $sphereMaterial }
        $sphereContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $sphereGeometryModel }

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
        $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D
        [void]$transformGroup.Children.Add($rotateTransform); [void]$transformGroup.Children.Add($translateTransform); [void]$transformGroup.Children.Add($scaleTransform)
        $sphereContainer.Transform = $transformGroup
        
        [void]$vortexContainer.Children.Add($sphereContainer)
        
        $SyncHash.SphereStates[$i] = @{
            SphereModel = $sphereGeometryModel
            TranslateTransform = $translateTransform
            RotateTransform = $rotateTransform
            ScaleTransform = $scaleTransform
            CurrentAngle = (720.0 / $sphereCount) * $i
        }
    }

    $vortexAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(45); RepeatBehavior = "Forever"; IsCumulative = $true }
    $SyncHash.VortexAnim = $vortexAnim

    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
            if ($SyncHash.flowTimer) { $SyncHash.flowTimer.Stop() }
        } else {
            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.VortexAnim.From = $SyncHash.VortexMasterRotation.Angle
            $SyncHash.VortexMasterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.VortexAnim)
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
            if ($SyncHash.flowTimer) { $SyncHash.flowTimer.Start() }
        }
    })

    $changeSpeed = {
        param($multiplier)
        $SyncHash.SpeedMultiplier *= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        $SyncHash.VortexAnim.Duration = [TimeSpan]::FromSeconds(45.0 / $SyncHash.SpeedMultiplier)
        if ($SyncHash.flowTimer) {
            $newInterval = 2.0 / $SyncHash.SpeedMultiplier
            if ($newInterval -lt 0.1) { $newInterval = 0.1 }
            $SyncHash.flowTimer.Interval = [TimeSpan]::FromSeconds($newInterval)
        }
    }
    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 0.5 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 2.0 })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        $SyncHash.ControlsHidden = -not $SyncHash.ControlsHidden
        $controlsPanel.Visibility = if ($SyncHash.ControlsHidden) { 'Collapsed' } else { 'Visible' }
    })

    $animationHandler = {
        param($sender, $e)
        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier
        $maxScale = 1.2; $minScale = 0.1; $totalAngle = 720

        foreach ($i in 0..($sphereCount-1)) {
            $panelState = $SyncHash.SphereStates[$i]
            $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle
            $angle = $panelState.CurrentAngle; $angleRad = $angle * ([Math]::PI / 180.0)
            $progress = $angle / $totalAngle
            $currentScale = $maxScale - ($progress * ($maxScale - $minScale))

            if ($SyncHash.VortexOrientation -eq "Vertical") {
                $startRadius = 3.5; $endRadius = 0.5; $startY = 5.0; $endY = -20.0
                $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius))
                $currentY = $startY - ($progress * ($startY - $endY))
                $currentX = $currentRadius * [Math]::Cos($angleRad); $currentZ = $currentRadius * [Math]::Sin($angleRad)
                $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ
                $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, $currentY - 1.5, 0)
                $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)
            } else {
                $startRadiusH = 5.0; $endRadiusH = 0.5; $startZ = 5.0; $endZ = -20.0
                $currentRadius = $startRadiusH - ($progress * ($startRadiusH - $endRadiusH))
                $currentZ = $startZ - ($progress * ($startZ - $endZ))
                $currentX = $currentRadius * [Math]::Cos($angleRad); $currentY = $currentRadius * [Math]::Sin($angleRad)
                $panelState.TranslateTransform.OffsetX = $currentX; $panelState.TranslateTransform.OffsetY = $currentY; $panelState.TranslateTransform.OffsetZ = $currentZ
                $lookAtTarget = [System.Windows.Media.Media3D.Point3D]::new(0, 0, $currentZ - 1.5)
                $position = [System.Windows.Media.Media3D.Point3D]::new($currentX, $currentY, $currentZ)
            }

            $panelState.ScaleTransform.ScaleX = $currentScale; $panelState.ScaleTransform.ScaleY = $currentScale; $panelState.ScaleTransform.ScaleZ = $currentScale
            $forward = $position - $lookAtTarget; if ($forward.Length -gt 0) { $forward.Normalize() }
            $up = if ($SyncHash.VortexOrientation -ne "Vertical" -and [Math]::Abs([System.Windows.Media.Media3D.Vector3D]::DotProduct($forward, '0,1,0')) -gt 0.99) { '1,0,0' } else { '0,1,0' }
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

    $window.Add_KeyDown({
        param($s, $e)
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
        if ($SyncHash.flowTimer) { $SyncHash.flowTimer.Stop() }
        if ($SyncHash.MediaGroupStates) {
            foreach ($i in $SyncHash.MediaGroupStates.Keys) {
                $playerState = $SyncHash.MediaGroupStates[$i]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
            }
        }
    })

    $window.Add_Loaded({
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            for ($i = 0; $i -lt $numberOfGroups; $i++) {
                $tbFront = $SyncHash.MediaGroupStates[$i].OverlayTextBlock
                $tbFront.Foreground = $brush.Clone(); $tbFront.FontFamily = $fontFamily; $tbFront.FontSize = $SyncHash.FontSize
                $tbFront.FontWeight = $fontWeight; $tbFront.FontStyle = $fontStyle
                if ($SyncHash.RbSelection -eq "Custom") { $tbFront.Text = $SyncHash.CustomText }
            }
        }
        
        # --- Media Flow Animation Timer ---
        $flowTimer = New-Object System.Windows.Threading.DispatcherTimer
        $flowTimer.Interval = [TimeSpan]::FromSeconds(2.0 / $SyncHash.SpeedMultiplier)
        $flowTimer.Add_Tick({
            $lastMaterial = $sharedMaterials[-1]
            for ($i = $sharedMaterials.Count - 1; $i -gt 0; $i--) { $sharedMaterials[$i] = $sharedMaterials[$i-1] }
            $sharedMaterials[0] = $lastMaterial

            foreach ($i in 0..($sphereCount-1)) {
                $materialForThisSphere = $sharedMaterials[$i % $sharedMaterials.Count]
                $SyncHash.SphereStates[$i].SphereModel.Material = $materialForThisSphere
            }
        })
        $SyncHash.flowTimer = $flowTimer
        $flowTimer.Start()

        # Start media on the 6 unique players
        for ($i = 0; $i -lt $numberOfGroups; $i++) { Start-NextMediaForGroup -GroupIndex $i }

        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $SyncHash.VortexMasterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.VortexAnim)
        [System.Windows.Media.CompositionTarget]::add_Rendering($animationHandler)
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}
