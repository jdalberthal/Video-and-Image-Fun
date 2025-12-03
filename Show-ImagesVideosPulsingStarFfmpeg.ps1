<#
.SYNOPSIS
    Displays media on a rotating 3D "star" shape with a radiating, pulsing effect.
.DESCRIPTION
    This script creates a 3D object composed of a central sphere and four cones. The media is
    animated in a "pulsing" or "breathing" effect: an image or video appears on the central sphere,
    and after a delay, it radiates outwards, appearing on the four cones simultaneously. As this
    happens, a new media item appears on the central sphere, and the cycle repeats.

    The entire object rotates, and the animation is managed by five separate, coordinated media players.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosPulsingStarFfmpeg.ps1
.NOTES
    Name:           Show-ImagesVideosPulsingStarFfmpeg.ps1
    Version:        1.0.0, 11/21/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Dependencies:   FFmpeg (ffmpeg.exe, ffplay.exe, ffprobe.exe)
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Dependency Check ---
$RequiredExecutables = @("ffmpeg.exe", "ffplay.exe", "ffprobe.exe")
$dependenciesMissing = $false
$dependencyStatus = foreach ($exe in $RequiredExecutables) {
    $isFound = (Get-Command $exe -ErrorAction SilentlyContinue) -or (Test-Path (Join-Path $PSScriptRoot $exe))
    if (-not $isFound) { $dependenciesMissing = $true }
    [pscustomobject]@{
        Name   = $exe
        Status = if ($isFound) { 'Found' } else { 'NOT FOUND' }
    }
}
if ($dependenciesMissing) {
    $messageLines = @(
        "One or more required executables were not found in your system's PATH or the script's directory."
        "Please install FFmpeg (including ffplay and ffprobe) and ensure they are accessible."
        ""
        "Required executable status:"
    ) + ($dependencyStatus | ForEach-Object { " - $($_.Status): $($_.Name)" })
    [System.Windows.Forms.MessageBox]::Show(($messageLines -join "`n"), "Dependency Error", "OK", "Error")
    exit
}

# --- Geometry Generation Functions ---
function New-SphereMesh {
    param([double]$radius=1.5, [int]$slices=64, [int]$stacks=32)
    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    for ($stack = 0; $stack -le $stacks; $stack++) {
        $phi = [Math]::PI/2 - $stack * [Math]::PI / $stacks; $y = $radius * [Math]::Sin($phi); $r = $radius * [Math]::Cos($phi)
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
    param([double]$radius=1.5, [double]$height=3.0, [int]$slices=64)
    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $height, 0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0))
    for ($i = 0; $i -le $slices; $i++) {
        $theta = $i * 2 * [Math]::PI / $slices; $x = $radius * [Math]::Cos($theta); $z = $radius * [Math]::Sin($theta)
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, 0, $z)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $slices, 1))
    }
    $baseCenterIndex = $mesh.Positions.Count
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, 0, 0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0.5))
    for ($i = 0; $i -lt $slices; $i++) { $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add($i + 2); $mesh.TriangleIndices.Add($i + 1) }
    for ($i = 0; $i -lt $slices; $i++) { $mesh.TriangleIndices.Add($baseCenterIndex); $mesh.TriangleIndices.Add($i + 1); $mesh.TriangleIndices.Add($i + 2) }
    return $mesh
}

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form -Property @{ Text = "Pulsing Star (FFmpeg) - Media Selector"; Size = '800, 680'; StartPosition = "CenterScreen" }
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
        } catch { $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12) }

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
        } catch { $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style) }
    }
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont)
    $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont)
    $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)
    & $updateTextBoxFont

    $SelectAllCheckbox.Add_CheckedChanged({ $isChecked = $SelectAllCheckbox.Checked; foreach ($row in $DataGridView.Rows) { $row.Cells["Select"].Value = $isChecked }; $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) })
    $BrowseButton.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SelectedPath = $FolderBrowser.SelectedPath; $FolderPathTextBox.Text = $SelectedPath; $DataGridView.Rows.Clear()
            $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
            $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.mov", "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2", "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v", "*.f4p", "*.f4a", "*.f4b"
            $gciParams = @{ File = $true; Include = ($ImageExtensions + $VideoExtensions) }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true } else { $gciParams.Path = Join-Path $SelectedPath "*" }
            Get-ChildItem @gciParams | ForEach-Object { $DataGridView.Rows.Add($false, $_.Name, $_.FullName) }
            $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
        }
    })
    $DataGridView.Add_RowHeaderMouseClick({ param($s, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGridView.Rows.Count) {
            $filePath = $DataGridView.Rows[$e.RowIndex].Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) { Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -i `"$filePath`"" -NoNewWindow }
        }
    })
    $PlayButton.Add_Click({
        $formState.SelectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
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
    $null = $SelectForm.ShowDialog(); $SelectForm.Dispose()
    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0) { Write-Host "Exiting."; break }

    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
        SelectedFiles = $formState.SelectedFiles
        UseTransparentEffect = $formState.UseTransparentEffect
        CurrentSphereIndex = -1
        CurrentConeIndex = -1
        PlayerStates = [hashtable]::Synchronized(@{})
        Paused = $false
        RedoClicked = $false
        # Text Overlay Settings
        RbSelection = $formState.RbSelection
        CustomText = $formState.CustomText
        TextColor = $formState.TextColor
        FontSize = $formState.FontSize
        FontFamily = $formState.FontFamily
        IsBold = $formState.IsBold
        IsItalic = $formState.IsItalic
    })

    # --- XAML Definition ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="MainWindow"
        Title="Pulsing Star" WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,12" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>
            <ModelVisual3D x:Name="ObjectContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <ScaleTransform3D x:Name="PulseScale" ScaleX="1" ScaleY="1" ScaleZ="1" />
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                    </Transform3DGroup>
                </ModelVisual3D.Transform>
            </ModelVisual3D>
        </Viewport3D>
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="pauseButton" Content="Pause" Padding="10,5" Margin="2"/>
            <Button Name="slowDownButton" Content="&#x2190;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="speedUpButton" Content="&#x2192;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="rotSlowDownButton" Content="Rot-" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="rotSpeedUpButton" Content="Rot+" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="redoButton" Content="Redo" Padding="10,5" Margin="2"/>
            <Button Name="closeButton" Content="X" Padding="10,5" Margin="2" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $reader = New-Object System.Xml.XmlNodeReader $xaml; $window = [Windows.Markup.XamlReader]::Load($reader)
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width; $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left; $window.Top = $primaryScreen.WorkingArea.Top
    $objectContainer = $window.FindName("ObjectContainer"); $SyncHash.Window = $window
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.rotSlowDownButton = $window.FindName("rotSlowDownButton")
    $SyncHash.rotSpeedUpButton = $window.FindName("rotSpeedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.closeButton = $window.FindName("closeButton")

    # --- Media Handling Functions ---

    # --- Animation and Pulse Logic ---
    $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(25)); $animX.RepeatBehavior = 'Forever'
    $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(35)); $animY.RepeatBehavior = 'Forever'
    $axisAngleX = $window.FindName("AxisAngleX"); $axisAngleY = $window.FindName("AxisAngleY")
    $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
    $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
    
    # Add the physical pulsing animation
    $pulseAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
        From = 0.765; To = 1.265; Duration = [TimeSpan]::FromSeconds(1.5) # Pulse from smaller to larger
        AutoReverse = $true; RepeatBehavior = 'Forever'
    }
    $pulseScale = $window.FindName("PulseScale")
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $pulseAnim)
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $pulseAnim)
    $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $pulseAnim)
    $SyncHash.pulseAnim = $pulseAnim # Store for speed control
    $SyncHash.animX = $animX; $SyncHash.animY = $animY
    $SyncHash.AxisAngleX = $axisAngleX; $SyncHash.AxisAngleY = $axisAngleY; $SyncHash.pulseScale = $pulseScale

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })
    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pulseTimer.Stop()
            # Correctly pause animations by detaching them
            $currentAngleX = $SyncHash.AxisAngleX.Angle; $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $SyncHash.AxisAngleX.Angle = $currentAngleX
            $currentAngleY = $SyncHash.AxisAngleY.Angle; $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $SyncHash.AxisAngleY.Angle = $currentAngleY
            $currentScale = $SyncHash.pulseScale.ScaleX; $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $null); $SyncHash.pulseScale.ScaleX = $currentScale
            $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $null); $SyncHash.pulseScale.ScaleY = $currentScale
            $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $null); $SyncHash.pulseScale.ScaleZ = $currentScale
            $SyncHash.pauseButton.Content = "Resume"
        } else {
            # Correctly resume animations by re-applying them from the current state
            $SyncHash.animX.From = $SyncHash.AxisAngleX.Angle; $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animX)
            $SyncHash.animY.From = $SyncHash.AxisAngleY.Angle; $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animY)           
            # For the pulse animation, we don't set .From, we just restart it to maintain the cycle.
            $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $SyncHash.pulseAnim)
            $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleYProperty, $SyncHash.pulseAnim)
            $SyncHash.pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleZProperty, $SyncHash.pulseAnim)
            $SyncHash.pulseTimer.Start()
            $SyncHash.pauseButton.Content = "Pause"
        }
    })
    $window.Add_KeyDown({ param($s, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Ctrl+Left" { $SyncHash.rotSlowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Ctrl+Right" { $SyncHash.rotSpeedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $changePulseSpeed = {
        param($multiplier)
        $newDuration = [TimeSpan]::FromSeconds(($SyncHash.pulseAnim.Duration.TimeSpan.TotalSeconds * $multiplier))
        if ($newDuration.TotalSeconds -lt 0.1) { $newDuration = [TimeSpan]::FromSeconds(0.1) }
        $SyncHash.pulseAnim.Duration = $newDuration

        if (-not $SyncHash.Paused) {
            # Briefly pause and resume all animations to apply the new speed
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Pause
            Start-Sleep -Milliseconds 50
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Resume
        }
    }
    $SyncHash.slowDownButton.Add_Click({ & $changePulseSpeed 2.0 })
    $SyncHash.speedUpButton.Add_Click({ & $changePulseSpeed 0.5 })

    $changeRotationSpeed = {
        param($multiplier)
        $newDurationX = [TimeSpan]::FromSeconds(($SyncHash.animX.Duration.TimeSpan.TotalSeconds * $multiplier))
        $newDurationY = [TimeSpan]::FromSeconds(($SyncHash.animY.Duration.TimeSpan.TotalSeconds * $multiplier))
        if ($newDurationX.TotalSeconds -lt 0.5) { $newDurationX = [TimeSpan]::FromSeconds(0.5) }
        if ($newDurationY.TotalSeconds -lt 0.5) { $newDurationY = [TimeSpan]::FromSeconds(0.5) }
        $SyncHash.animX.Duration = $newDurationX
        $SyncHash.animY.Duration = $newDurationY

        if (-not $SyncHash.Paused) {
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Pause
            Start-Sleep -Milliseconds 50
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Resume
        }
    }
    $SyncHash.rotSlowDownButton.Add_Click({ & $changeRotationSpeed 2.0 })
    $SyncHash.rotSpeedUpButton.Add_Click({ & $changeRotationSpeed 0.5 })

    $window.Add_Closed({
        if ($SyncHash.pulseTimer) { $SyncHash.pulseTimer.Stop() }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $pulseScale.BeginAnimation([System.Windows.Media.Media3D.ScaleTransform3D]::ScaleXProperty, $null)
        foreach ($target in $SyncHash.PlayerStates.Keys) {
            $playerState = $SyncHash.PlayerStates[$target]
            if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
            if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
        }
    })
    $window.Add_Loaded({
        # --- Media Handling Functions ---
        $script:Start_MediaForTarget = {
            param([string]$Target, [string]$FilePath)
            $playerState = $SyncHash.PlayerStates[$Target]
            if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
            if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }

            # Update text overlay if set to "Filename"
            $fileName = [System.IO.Path]::GetFileName($FilePath)
            if ($SyncHash.RbSelection -eq "Filename") {
                $playerState.OverlayTextBlock.Text = $fileName
            }
            
            $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
            if ($ImageExtensions -contains [System.IO.Path]::GetExtension($FilePath).ToLower()) {
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$FilePath); Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $image
            } else {
                $width=1280; $height=720; $frameSize=$width*$height*3
                $bitmap = New-Object System.Windows.Media.Imaging.WriteableBitmap($width, $height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $imageControl = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = "Fill" }
                $playerState.ContentPresenter.Content = $imageControl
                $args = "-hide_banner -loglevel error -stream_loop -1 -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"
                $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{ FileName="ffmpeg.exe"; Arguments=$args; RedirectStandardOutput=$true; UseShellExecute=$false; CreateNoWindow=$true }
                $proc = [System.Diagnostics.Process]::Start($psi); $playerState.FfmpegProcess = $proc
                $stream = $proc.StandardOutput.BaseStream; $bytes = New-Object byte[] $frameSize; $rect = [System.Windows.Int32Rect]::new(0,0,$width,$height); $stride=$width*3
                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval=[TimeSpan]::FromMilliseconds(33); Tag=$Target }
                $tickScriptBlock = {
                    $t = $args[0]; try {
                        $totalRead=0; while($totalRead -lt $frameSize){ $read=$stream.Read($bytes, $totalRead, $frameSize-$totalRead); if($read -le 0){$t.Stop();return}; $totalRead+=$read }
                        if($totalRead -eq $frameSize){ $bitmap.Lock(); $bitmap.WritePixels($rect, $bytes, $stride, 0); $bitmap.Unlock() }
                    } catch { $t.Stop() }
                }; $timer.Add_Tick($tickScriptBlock.GetNewClosure()); $playerState.MediaTimer = $timer; $timer.Start()
            }
        }

        # --- Create Star Geometry and Media Players ---
        # Correctly calculate the visible height at the origin based on camera distance and FOV.
        $camera = $window.FindName('mainViewport').Camera
        $distance = $camera.Position.Z
        $fovRadians = $camera.FieldOfView * ([Math]::PI / 180.0)
        $visibleHeight = 2.0 * $distance * [Math]::Tan($fovRadians / 2.0)

        $totalObjectHeight = $visibleHeight * 0.4 # Make the object 40% of the visible height to make it smaller.
        $sphereRadius = $totalObjectHeight / 5.0; $coneHeight = $sphereRadius * 1.5; $coneRadius = $sphereRadius * 0.4
        $sphereMesh = New-SphereMesh -radius $sphereRadius; $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        
        $targets = "Sphere", "TopCone", "BottomCone", "LeftCone", "RightCone"
        foreach ($target in $targets) {
            # Create the 2D content host
            $grid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
                HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
            }
            $grid.Children.Add($contentPresenter) | Out-Null
            $grid.Children.Add($overlayTextBlock) | Out-Null

            # Apply Text Overlay Settings
            if ($SyncHash.RbSelection -ne "Hidden") {
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $overlayTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $overlayTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
                $overlayTextBlock.FontSize = $SyncHash.FontSize
                if ($SyncHash.IsBold) { $overlayTextBlock.FontWeight = 'Bold' }
                if ($SyncHash.IsItalic) { $overlayTextBlock.FontStyle = 'Italic' }
                if ($SyncHash.RbSelection -eq "Custom") { $overlayTextBlock.Text = $SyncHash.CustomText }
            }
            
            # Create the Viewport2DVisual3D, which is the robust way to host 2D content in 3D
            $viewportVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
            $geometry = if ($target -eq "Sphere") { $sphereMesh } else { $coneMesh }
            $viewportVisual.Geometry = $geometry
            $viewportVisual.Visual = $grid

            # Create a special material that links to the hosted visual
            $material = New-Object $materialType
            [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($material, $true)
            $viewportVisual.Material = $material
            
            $transform = switch ($target) {
                "TopCone"    { New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0) }
                "BottomCone" { $tg=New-Object System.Windows.Media.Media3D.Transform3DGroup; $tg.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1,-1,1))); $tg.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0,-$sphereRadius,0))); $tg }
                "LeftCone"   { $tg=New-Object System.Windows.Media.Media3D.Transform3DGroup; $tg.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,0,1),90))))); $tg.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius,0,0))); $tg }
                "RightCone"  { $tg=New-Object System.Windows.Media.Media3D.Transform3DGroup; $tg.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,0,1),-90))))); $tg.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius,0,0))); $tg }
            }
            if ($transform) { $viewportVisual.Transform = $transform }
            
            # Add the final Viewport2DVisual3D to the scene
            $objectContainer.Children.Add($viewportVisual)
            $SyncHash.PlayerStates[$target] = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock }
        }

        $pulseTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(5) }
        $pulseTimer.Add_Tick({
            $SyncHash.CurrentConeIndex = $SyncHash.CurrentSphereIndex
            $coneFilePath = $SyncHash.SelectedFiles[$SyncHash.CurrentConeIndex]
            foreach ($coneTarget in @("TopCone", "BottomCone", "LeftCone", "RightCone")) { & $script:Start_MediaForTarget -Target $coneTarget -FilePath $coneFilePath }
            $SyncHash.CurrentSphereIndex = ($SyncHash.CurrentSphereIndex + 1) % $SyncHash.SelectedFiles.Count
            $sphereFilePath = $SyncHash.SelectedFiles[$SyncHash.CurrentSphereIndex]
            & $script:Start_MediaForTarget -Target "Sphere" -FilePath $sphereFilePath
        })
        $SyncHash.pulseTimer = $pulseTimer

        # Initial load
        # Load the first media item onto the sphere
        $SyncHash.CurrentSphereIndex = 0
        $sphereFilePath = $SyncHash.SelectedFiles[$SyncHash.CurrentSphereIndex]
        & $script:Start_MediaForTarget -Target "Sphere" -FilePath $sphereFilePath
        
        # Immediately load the second media item onto the cones to start fully populated
        $coneIndex = if ($SyncHash.SelectedFiles.Count -gt 1) { 1 } else { 0 }
        $SyncHash.CurrentConeIndex = $coneIndex
        $coneFilePath = $SyncHash.SelectedFiles[$coneIndex]
        foreach ($coneTarget in @("TopCone", "BottomCone", "LeftCone", "RightCone")) {
            & $script:Start_MediaForTarget -Target $coneTarget -FilePath $coneFilePath
        }
        $pulseTimer.Start()
    })

    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
