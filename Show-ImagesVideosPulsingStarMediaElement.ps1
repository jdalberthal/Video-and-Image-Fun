<#
.SYNOPSIS
    Displays media on a 3D object resembling a pulsing star, made of a central sphere and multiple cones.
.DESCRIPTION
    This script launches a GUI to select media files, then renders them onto a complex 3D object.
    The object consists of a central sphere and six cones pointing outwards, resembling a star.
    Each of the seven components (one sphere, six cones) has its own independent media player.

    The cones are animated to pulse in and out from the center, creating a dynamic "breathing" 
    effect, while the entire star object rotates.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system (e.g., MP4, WMV, AVI).
.EXAMPLE
    PS C:\> .\Show-ImagesVideosPulsingStarMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D pulsing star window.
.NOTES
    Name:           Show-ImagesVideosPulsingStarMediaElement.ps1
    Version:        1.0.0, 11/19/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Geometry Generation Functions ---
function New-SphereMesh {
    param([double]$radius = 1.5, [int]$slices = 64, [int]$stacks = 32)
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

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Pulsing Star - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{ Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'; AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65 }
    
    $SelectForm.Controls.AddRange(@($BrowseButton, $FolderPathTextBox, $RecursiveCheckBox, $SelectAllCheckbox, $TransparentCheckbox, $DataGridView))

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 200; ReadOnly = $true }
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 330; ReadOnly = $true }
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
    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.TextColor = $colorDialog.Color
            $ColorExample.BackColor = $formState.TextColor
            & $updateTextBoxFont # Update the textbox color immediately
        }
    })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            $fontDialog.ShowColor = $true # Allow color selection in the font dialog
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

            Get-ChildItem @gciParams | ForEach-Object {
                $DataGridView.Rows.Add($false, $_.Name, $_.FullName)
            }
            $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGridView.Rows.Count) {
            $filePath = $DataGridView.Rows[$e.RowIndex].Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                try { Start-Process $filePath } catch { [System.Windows.Forms.MessageBox]::Show("Could not open file: $($_.Exception.Message)", "Error", "OK", "Error") }
            }
        }
    })

    $PlayButton.Add_Click({
        $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
        if ($selectedFiles.Count -gt 0) {
            # Convert the fixed-size array to a resizable ArrayList here.
            # This is critical for the Handle-MediaFailure function to be able to remove items.
            $formState.SelectedFiles = [System.Collections.ArrayList]::new($selectedFiles)
            
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
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
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
        SelectedFiles = $formState.SelectedFiles
        CurrentIndex = -1
        PlayerStates = [hashtable]::Synchronized(@{})
        Paused = $false
        ControlsHidden = $false
        RedoClicked = $false
        UseTransparentEffect = $formState.UseTransparentEffect
        RbSelection = $formState.RbSelection
        CustomText = $formState.CustomText
        TextColor = $formState.TextColor
        FontSize = $formState.FontSize
        FontFamily = $formState.FontFamily
        IsBold = $formState.IsBold
        IsItalic = $formState.IsItalic
    })
    $SyncHash.SpeedMultiplier = 1.0

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="StarWindow"
        Title="Pulsing Star"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,12" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>

            <ModelVisual3D x:Name="StarContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <ScaleTransform3D x:Name="PulseScale" ScaleX="1" ScaleY="1" ScaleZ="1" />
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="StarRotation" Axis="1,1,0.5" Angle="0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                    </Transform3DGroup>
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

    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width; $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left; $window.Top = $primaryScreen.WorkingArea.Top

    $starContainer = $window.FindName("StarContainer")
    $SyncHash.Window = $window

    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $SyncHash.StarRotation = $window.FindName("StarRotation")

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
    function Handle-MediaFailure {
        param([string]$ObjectName, [string]$Reason)

        $playerState = $SyncHash.PlayerStates[$ObjectName]
        if ($playerState.IsFailed) { return } # Prevent re-entry

        $playerState.IsFailed = $true

        # 1. Log directly to the console.
        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for object '$ObjectName' (File: '$fileName'). Reason: $Reason. Attempting to replace."

        # 2. Remove the bad file from the list.
        if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
            $SyncHash.SelectedFiles.Remove($playerState.CurrentSource.OriginalString) | Out-Null
        }

        # 3. Dispatch ONLY the pure UI updates (blacking out the panel).
        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState.ContentPresenter.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
        })

        # 4. Immediately call for the next media file, OUTSIDE the dispatcher. This is the critical step.
        # Use a short-lived timer to avoid potential re-entrancy issues within the event handler.
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(50); Tag = $ObjectName }
        $recoveryTimer.Add_Tick({ $t = $args[0]; $objName = $t.Tag; $t.Stop(); Start-NextMediaForObject -ObjectName $objName })
        $playerState.RecoveryTimer = $recoveryTimer
        $recoveryTimer.Start()
    }


    function Start-NextMediaForObject {
        param([string]$ObjectName) # e.g., "Sphere", "Cone1", etc.

        $playerState = $SyncHash.PlayerStates[$ObjectName]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false # Reset failure flag

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

        try {
            if ($ImageExtensions -contains $extension) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit()
                $bitmap.Freeze()

                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                $playerState.ContentPresenter.Content = $image

                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $ObjectName }
                $timer.Add_Tick({ $t = $args[0]; $objName = $t.Tag; $t.Stop(); Start-NextMediaForObject -ObjectName $objName })
                $playerState.MediaTimer = $timer
                $timer.Start()
            }
            else { # Video
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource
                }
                $mediaElement.Tag = $ObjectName

                $mediaElement.Add_MediaEnded({
                    $objName = $args[0].Tag
                    $pState = $SyncHash.PlayerStates[$objName]
                    if ($pState.IsFailed) { return }
                    $pState.PlaybackStopwatch.Stop()
                    if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1500) {
                        Handle-MediaFailure -ObjectName $objName -Reason "Playback ended instantly (bad codec)."
                        return
                    }
                    Start-NextMediaForObject -ObjectName $objName
                })
                $mediaElement.Add_MediaOpened({
                    $objName = $args[0].Tag
                    $pState = $SyncHash.PlayerStates[$objName]
                    if (-not $args[0].NaturalDuration.HasTimeSpan) {
                        Handle-MediaFailure -ObjectName $objName -Reason "Invalid duration or codec."
                    } else {
                        $pState.PlaybackStopwatch.Restart()
                        $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                    }
                })
                $mediaElement.Add_MediaFailed({ Handle-MediaFailure -ObjectName $args[0].Tag -Reason $args[1].ErrorException.Message })

                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.CurrentMediaElement = $mediaElement
                $mediaElement.Play()
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            }
        } catch {
            Handle-MediaFailure -ObjectName $ObjectName -Reason $_.Exception.Message
        }
    }

    # --- Create Star Components ---
    # Reduce the base size of the star to prevent it from pulsing off-screen.
    # The original values were 1.5, 3.0, and 1.0 respectively.
    $sphereRadius = 1.2; $coneHeight = 2.4; $coneRadius = 0.8
    $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64
    $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

    function New-MediaHost {
        $grid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $grid.Children.Add($contentPresenter) | Out-Null
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $grid.Children.Add($overlayTextBlock) | Out-Null
        return @{ Grid = $grid; ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; PlaybackStopwatch = (New-Object System.Diagnostics.Stopwatch) }
    }

    # 1. Central Sphere
    $sphereHosts = New-MediaHost
    $sphereVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $sphereMesh; Visual = $sphereHosts.Grid }
    $sphereMaterial = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($sphereMaterial, $true)
    $sphereHosts.Grid.Name = "Sphere_Grid"
    $sphereVisual.Material = $sphereMaterial
    $starContainer.Children.Add($sphereVisual)
    $SyncHash.PlayerStates["Sphere"] = @{ 
        ContentPresenter = $sphereHosts.ContentPresenter; OverlayTextBlock = $sphereHosts.OverlayTextBlock; MediaHostGrid = $sphereHosts.Grid; PlaybackStopwatch = $sphereHosts.PlaybackStopwatch
    }

    # 2. Cones
    $conePositions = @(
        @{ Name="ConeTop";    Transform=(New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)) },
        @{ Name="ConeBottom"; Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
        @{ Name="ConeRight";  Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
        @{ Name="ConeLeft";   Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
        @{ Name="ConeFront";  Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) },
        @{ Name="ConeBack";   Transform=(New-Object System.Windows.Media.Media3D.Transform3DGroup) }
    )
    # ConeBottom
    $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1)))
    $conePositions[1].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
    # ConeRight
    $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', -90)))))
    $conePositions[2].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius, 0, 0)))
    # ConeLeft
    $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,0,1', 90)))))
    $conePositions[3].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius, 0, 0)))
    # ConeFront
    $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', 90)))))
    $conePositions[4].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, $sphereRadius)))
    # ConeBack
    $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('1,0,0', -90)))))
    $conePositions[5].Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$sphereRadius)))

    foreach ($coneInfo in $conePositions) {
        $coneHosts = New-MediaHost
        $coneVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh; Visual = $coneHosts.Grid }
        $coneMaterial = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($coneMaterial, $true) 
        $coneVisual.Material = $coneMaterial
        
        # The transform is now just for placement; the pulsing is handled by the parent container.
        $coneVisual.Transform = $coneInfo.Transform

        $starContainer.Children.Add($coneVisual)
        $SyncHash.PlayerStates[$coneInfo.Name] = @{ 
            ContentPresenter = $coneHosts.ContentPresenter; OverlayTextBlock = $coneHosts.OverlayTextBlock;
            MediaHostGrid = $coneHosts.Grid; PlaybackStopwatch = $coneHosts.PlaybackStopwatch
        }
    }

    # --- Animations ---
    $starAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{ From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(30); RepeatBehavior = "Forever" }
    $SyncHash.StarRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $starAnim)
    $SyncHash.StarAnim = $starAnim

    # Replicate the pulsing scale animation from the FFmpeg version
    $pulseAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
        From = 0.85; To = 1.15; Duration = [TimeSpan]::FromSeconds(2); AutoReverse = $true; RepeatBehavior = "Forever"
    }
    $pulseScale = $window.FindName("PulseScale")
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $pulseAnim)
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $pulseAnim)
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $pulseAnim)
    $SyncHash.PulseAnim = $pulseAnim

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        if ($SyncHash.Paused) {
            $SyncHash.StarAnim.From = $SyncHash.StarRotation.Angle
            $SyncHash.StarRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.StarAnim)
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $SyncHash.PulseAnim)
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $SyncHash.PulseAnim)
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $SyncHash.PulseAnim)
            $SyncHash.pauseButton.Content = "Pause"; $SyncHash.Paused = $false
        } else {
            $SyncHash.StarRotation.Angle = $SyncHash.StarRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
            $SyncHash.StarRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $currentScale = $pulseScale.ScaleX
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $null); $pulseScale.ScaleX = $currentScale
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $null); $pulseScale.ScaleY = $currentScale
            $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $null); $pulseScale.ScaleZ = $currentScale
            $SyncHash.pauseButton.Content = "Resume"; $SyncHash.Paused = $true
        }
    })

    $SyncHash.randomAxisButton.Add_Click({
        $SyncHash.StarRotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
    })

    $changeSpeed = {
        param($multiplier)
        $wasPaused = $SyncHash.Paused
        if (-not $wasPaused) { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }

        $SyncHash.SpeedMultiplier *= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        
        $SyncHash.StarAnim.Duration = [TimeSpan]::FromSeconds(30.0 / $SyncHash.SpeedMultiplier)
        $SyncHash.PulseAnim.Duration = [TimeSpan]::FromSeconds(2.0 / $SyncHash.SpeedMultiplier)

        if (-not $wasPaused) { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
    }

    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 0.5 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 2.0 })

    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        $controlsPanel.Visibility = if ($controlsPanel.Visibility -eq 'Visible') { 'Collapsed' } else { 'Visible' }
        $SyncHash.ControlsHidden = -not $SyncHash.ControlsHidden
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
            foreach ($key in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$key]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                if ($playerState.CurrentMediaElement) {
                    $playerState.CurrentMediaElement.Stop()
                    $playerState.CurrentMediaElement.Close()
                }
            }
        }
    })

    # --- Start the show ---
    $window.Add_Loaded({
        # Apply Text Overlay Settings
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            foreach ($key in $SyncHash.PlayerStates.Keys) {
                $textBlock = $SyncHash.PlayerStates[$key].OverlayTextBlock
                $textBlock.Foreground = $brush.Clone()
                $textBlock.FontFamily = $fontFamily
                $textBlock.FontSize = $SyncHash.FontSize
                $textBlock.FontWeight = $fontWeight
                $textBlock.FontStyle = $fontStyle
                if ($SyncHash.RbSelection -eq "Custom") { $textBlock.Text = $SyncHash.CustomText }
            }
        }

        # Start media on each object
        foreach ($key in $SyncHash.PlayerStates.Keys) {
            Start-NextMediaForObject -ObjectName $key
        }
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}
