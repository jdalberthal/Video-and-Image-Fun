<#
.SYNOPSIS
    Displays a playlist of media on the faces of a rotating 3D pie using FFmpeg.
.DESCRIPTION
    This script creates a WPF window and renders a 3D pie shape. It prompts the
    user to select image and video files, then displays them on both the front
    and back faces of each of the 8 slices.

    The 3D pie tumbles and rotates in space, revealing all 16 media players.
    The rotation can be paused, and the rotation axis can be randomized.

    This version uses FFmpeg for video playback, providing support for a wide range of video formats.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosPie3DFfmpeg.ps1

    Launches the file selection GUI. After selecting at least one file and clicking "Play", the
    script will launch the 3D rotating pie window.
.NOTES
    Name:           Show-ImagesVideosPie3DFfmpeg.ps1
    Version:        1.0.0, 11/04/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffprobe.exe, ffplay.exe): https://www.ffmpeg.org/download.html
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- 3D Pie Slice Generation Function ---
function New-PieSliceModel {
    param(
        [System.Windows.Point]$center,
        [double]$radius,
        [double]$startAngleDeg,
        [double]$sliceAngleDeg,
        [double]$thickness = 0.1
    )
    
    $endAngleDeg = $startAngleDeg + $sliceAngleDeg
    $startAngleRad = $startAngleDeg * [Math]::PI / 180.0
    $endAngleRad = $endAngleDeg * [Math]::PI / 180.0
    $halfThick = $thickness / 2.0

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

    # Define the points for the arc
    $arcPoints = New-Object 'System.Collections.Generic.List[System.Windows.Point]'
    $arcPoints.Add($center)
    $numArcSegments = [double][Math]::Max(1, [int]($sliceAngleDeg / 5)) # Cast to double for floating point division
    for ($i = 0; $i -le $numArcSegments; $i++) {
        $angle = $startAngleRad + ($i / $numArcSegments) * ($endAngleRad - $startAngleRad)
        [double]$pointX = $center.X + $radius * [Math]::Cos($angle)
        [double]$pointY = $center.Y + $radius * [Math]::Sin($angle)
        $arcPoints.Add((New-Object System.Windows.Point($pointX, $pointY)))
    }

    # --- Generate Vertices and Texture Coordinates ---
    # Front Face
    $frontBaseIndex = $mesh.Positions.Count
    foreach ($p in $arcPoints) { 
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, $halfThick))
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5))
    }
    # Back Face
    $backBaseIndex = $mesh.Positions.Count
    foreach ($p in $arcPoints) { 
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($p.X, $p.Y, -$halfThick)) 
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(($p.X / (2*$radius)) + 0.5, -($p.Y / (2*$radius)) + 0.5))
    }

    # --- Generate Triangle Indices ---
    # Front Face Triangles
    for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) {
        $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $i + 1); $mesh.TriangleIndices.Add($frontBaseIndex + $i)
    }
    # Back Face Triangles
    for ($i = 1; $i -lt ($arcPoints.Count - 1); $i++) {
        $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $i); $mesh.TriangleIndices.Add($backBaseIndex + $i + 1)
    }
    # Outer Edge Triangles
    for ($i = 1; $i -lt $arcPoints.Count; $i++) {
        $p1_front = $frontBaseIndex + $i; $p2_front = $frontBaseIndex + $i + 1
        $p1_back = $backBaseIndex + $i;  $p2_back = $backBaseIndex + $i + 1
        $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p1_back); $mesh.TriangleIndices.Add($p2_back)
        $mesh.TriangleIndices.Add($p1_front); $mesh.TriangleIndices.Add($p2_back); $mesh.TriangleIndices.Add($p2_front)
    }
    # Side Edge 1 (Start)
    $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1)
    $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + 1); $mesh.TriangleIndices.Add($frontBaseIndex + 1)
    # Side Edge 2 (End)
    $lastIdx = $arcPoints.Count -1
    $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($frontBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx)
    $mesh.TriangleIndices.Add($frontBaseIndex); $mesh.TriangleIndices.Add($backBaseIndex + $lastIdx); $mesh.TriangleIndices.Add($backBaseIndex)

    $model = New-Object System.Windows.Media.Media3D.GeometryModel3D
    $model.Geometry = $mesh
    return $model
}

# --- Main Application Loop ---
while ($true) {

    # -------------------- File Selection Form --------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "3D Pie Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }

    # --- Dependency Check ---
    $RequiredExecutables = @("ffmpeg.exe", "ffprobe.exe", "ffplay.exe")
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
        return
    }
    $SelectForm.Controls.Add($BrowseButton)

    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $SelectForm.Controls.Add($FolderPathTextBox)

    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $SelectForm.Controls.Add($RecursiveCheckBox)

    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $SelectForm.Controls.Add($TransparentCheckbox)

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $SelectForm.Controls.Add($SelectAllCheckbox)

    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'
        AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersVisible = $true; RowHeadersWidth = 65
    }
    $SelectForm.Controls.Add($DataGridView)

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    [void]$DataGridView.Columns.Add($CheckBoxColumn)
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true }
    [void]$DataGridView.Columns.Add($FileNameColumn)
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 380; ReadOnly = $true }
    [void]$DataGridView.Columns.Add($FilePathColumn)

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Play Selected"; Location = '600, 40'; Size = '170, 30' }
    $SelectForm.Controls.Add($PlayButton)

    # --- Text Overlay Controls ---
    $GroupBox     = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 440'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $GroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $SelectForm.Controls.Add($GroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '140, 440'; Size = '455, 180'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center'
    }
    $SelectForm.Controls.Add($TextBox)

    $CurrentColor      = New-Object System.Windows.Forms.Label    -Property @{ Text = "Text Color:"; Location = '600, 477'; AutoSize = $true; Visible = $false }
    $ColorExample      = New-Object System.Windows.Forms.Label    -Property @{ Text = "     "; Location = '660, 477'; AutoSize = $true; BackColor = [System.Drawing.Color]::Black; Visible = $false }
    $SelectColorButton = New-Object System.Windows.Forms.Button   -Property @{ Text = "Change"; Location = '685, 470'; Size = '80, 30'; Visible = $false }
    $SizeLabel         = New-Object System.Windows.Forms.Label    -Property @{ Text = "Font Size:"; AutoSize = $true; Location = '600, 522'; Visible = $false }
    $NumericUpDown     = New-Object System.Windows.Forms.NumericUpDown -Property @{ Location = '660, 520'; Size = '50, 20'; Visible = $false; Minimum = 8; Maximum = 72; Value = 24 }
    $FontButton        = New-Object System.Windows.Forms.Button   -Property @{ Text = "Change Font"; Location = '600, 570'; Size = '170, 25'; Visible = $false }
    $ItalicCheckbox    = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Italic"; Location = '600, 620'; Size = '75, 20'; Checked = $false; Visible = $false }
    $BoldCheckbox      = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Bold"; Location = '680, 620'; Size = '75, 20'; Checked = $true;  Visible = $false }

    $SelectForm.Controls.AddRange(@(
        $CurrentColor, $ColorExample, $SelectColorButton, $SizeLabel,
        $NumericUpDown, $FontButton, $ItalicCheckbox, $BoldCheckbox
    ))

    # --- Script-scoped form state with safe defaults ---
    $script:formState = @{
      TextColor            = [System.Drawing.Color]::Black
      FontFamily           = "Arial"
      FontSize             = 24
      IsBold               = $true
      IsItalic             = $false
      RbSelection          = "Hidden"
      CustomText           = ""
      SelectedFiles        = @()
      UseTransparentEffect = $false
    }

    # Toggle visibility of text controls
    $textOverlayEvent = {
        $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked
        $isCustomText  = $RadioButton3.Checked

        $TextBox.Visible           = $isCustomText
        $CurrentColor.Visible      = $isTextVisible
        $ColorExample.Visible      = $isTextVisible
        $SelectColorButton.Visible = $isTextVisible
        $SizeLabel.Visible         = $isTextVisible
        $NumericUpDown.Visible     = $isTextVisible
        $FontButton.Visible        = $isTextVisible
        $ItalicCheckbox.Visible    = $isTextVisible
        $BoldCheckbox.Visible      = $isTextVisible
    }
    $RadioButton1.Add_Click($textOverlayEvent)
    $RadioButton2.Add_Click($textOverlayEvent)
    $RadioButton3.Add_Click($textOverlayEvent)

    # --- Event Handlers for Text Customization ---
    $ColorExample.BackColor = $script:formState.TextColor
    $SelectColorButton.Add_Click({
            $colorDialog = New-Object System.Windows.Forms.ColorDialog
            if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $script:formState.TextColor = $colorDialog.Color
                $ColorExample.BackColor = $script:formState.TextColor
                $TextBox.ForeColor = $script:formState.TextColor
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            $currentFont = New-Object System.Drawing.Font($script:formState.FontFamily, [float]$NumericUpDown.Value)
            $fontDialog.Font = $currentFont

            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $selectedFont = $fontDialog.Font
                $script:formState.FontFamily = $selectedFont.Name
                $FontButton.Text = $script:formState.FontFamily
                $NumericUpDown.Value = [decimal]$selectedFont.Size
                $BoldCheckbox.Checked = $selectedFont.Bold
                $ItalicCheckbox.Checked = $selectedFont.Italic
                & $updateTextBoxFont
            }
        })

    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($script:formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
        } catch {
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont)
    $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont)
    $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)

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
            $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.mov", "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2", "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v", "*.f4p", "*.f4a", "*.f4b"
            $AllowedExtensions = $ImageExtensions + $VideoExtensions

            $gciParams = @{
                File = $true
                Include = $AllowedExtensions
            }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true }
            else { $gciParams.Path = Join-Path $SelectedPath "*" }
            $files = Get-ChildItem @gciParams
            foreach ($file in $files) {
                $rowIdx = $DataGridView.Rows.Add($false, $file.Name, $file.FullName)
                $DataGridView.Rows[$rowIdx].HeaderCell.Value = "Play"
            }
        }
    })

    $PlayButton.Add_Click({
        $script:formState.SelectedFiles = @(
            foreach ($Row in $DataGridView.Rows) {
                if ($Row.Cells["Select"].Value) { $Row.Cells["FilePath"].Value }
            }
        )

        if ($script:formState.SelectedFiles.Count -gt 0) {
            $script:formState.UseTransparentEffect = $TransparentCheckbox.Checked
            if ($RadioButton1.Checked) { $script:formState.RbSelection = "Hidden" }
            if ($RadioButton2.Checked) { $script:formState.RbSelection = "Filename" }
            if ($RadioButton3.Checked) { $script:formState.RbSelection = "Custom" }
            $script:formState.CustomText = $TextBox.Text

            try {
                $script:formState.FontSize = [double]$NumericUpDown.Value
                if ($script:formState.FontSize -le 0) { $script:formState.FontSize = 24 }
            } catch { $script:formState.FontSize = 24 }
            $script:formState.IsBold   = $BoldCheckbox.Checked
            $script:formState.IsItalic = $ItalicCheckbox.Checked
            $SelectForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        $rowIndex = $_.RowIndex
        if ($rowIndex -lt 0) { return }
        $row = $DataGridView.Rows[$rowIndex]
        $videoPath = ($row.Cells["FilePath"].Value)
        if ([System.IO.File]::Exists($videoPath)) {
            Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -loop 0 -i `"$videoPath`"" -NoNewWindow
        } else {
            [System.Windows.Forms.MessageBox]::Show("File not found: $videoPath", "Error", "OK", "Error")
        }
    })

    $null = $SelectForm.ShowDialog()
    $SelectForm.Dispose()
    if ($script:formState.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break
    }

    # -------------------- Main WPF Window --------------------
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="3D Pie Viewer"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        WindowState="Maximized">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,8" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="#505050"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>
            <ModelVisual3D x:Name="PieContainer" />
        </Viewport3D>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="randomAxisButton" Content="Random Axis" Padding="10,5" Margin="2"/>
            <Button Name="slowDownButton" Content="&#x2190;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="speedUpButton" Content="&#x2192;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="pauseButton" Content="Pause" Padding="10,5" Margin="2"/>
            <Button Name="redoButton" Content="Redo" Padding="10,5" Margin="2"/>
            <Button Name="closeButton" Content="X" Padding="10,5" Margin="2" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $fontSz  = 24
    try { if ($script:formState.FontSize -and [double]$script:formState.FontSize -gt 0) { $fontSz = [double]$script:formState.FontSize } } catch { }

    $SyncLock = New-Object object
    $SyncHash = [pscustomobject]@{
        SelectedFiles = $script:formState.SelectedFiles
        RedoClicked = $false
        Paused = $false
        UseTransparentEffect = $script:formState.UseTransparentEffect
        RbSelection = $script:formState.RbSelection; CustomText = $script:formState.CustomText; TextColor = $script:formState.TextColor
        FontSize = $fontSz; FontFamily = $script:formState.FontFamily; IsBold = $script:formState.IsBold; IsItalic = $script:formState.IsItalic
        PlayerState = @{}; ImageExtensions = @(".bmp",".jpeg",".jpg",".png",".tif",".tiff",".gif",".wmp",".ico"); GlobalCounter = -1; ImageHoldSeconds = 10
        AnimationX = $null; AnimationY = $null; AxisAngleX = $null; AxisAngleY = $null
        Window = $null; SliceControls = (New-Object 'System.Collections.Generic.List[psobject]')
        GetNextIndex = $null; ApplyOverlayFor = $null; AssignNext = $null; HandleMediaFailure = $null; StartFfmpegStream = $null
        redoButton = $null; closeButton = $null ; pauseButton = $null; randomAxisButton = $null; slowDownButton = $null; speedUpButton = $null
        pauseOrResumeAnimation = $null; changeSpeed = $null
    }

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $SyncHash.Window = $window

    $pieContainer = $window.FindName("PieContainer")
    $pieModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
    $pieContainer.Content = $pieModelGroup

    $numberOfSlices = 8 # 8 front + 8 back = 16 media elements
    $sliceAngle = 360.0 / $numberOfSlices
    $pieRadius = 2.5 
    $pieCenter = New-Object System.Windows.Point(0, 0)

    for ($i = 0; $i -lt $numberOfSlices; $i++) {
        $startAngle = $i * $sliceAngle

        # Create a slice model (the 3D mesh)
        $sliceModel = New-PieSliceModel -center $pieCenter -radius $pieRadius -startAngleDeg $startAngle -sliceAngleDeg $sliceAngle

        # Create materials and media elements for BOTH front and back
        foreach ($face in @('Front', 'Back')) {
            $sliceIndex = if ($face -eq 'Front') { $i } else { $i + $numberOfSlices }

            $imageControl = New-Object System.Windows.Controls.Image -Property @{ Stretch = 'Fill' }

            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
                HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; Margin = '10,0,10,0'
                TextAlignment = 'Center'; IsHitTestVisible = $false
            }
            $contentGrid = New-Object System.Windows.Controls.Grid -Property @{ Background = [System.Windows.Media.Brushes]::Black }

            [void]$contentGrid.Children.Add($imageControl)
            [void]$contentGrid.Children.Add($overlayTextBlock)

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $contentGrid; Stretch = 'UniformToFill' }

            $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
            $material = New-Object $materialType
            $material.Brush = $visualBrush

            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White; $material.Brush.Opacity = 0.7 }

            if ($face -eq 'Front') { $sliceModel.Material = $material }
            else { $sliceModel.BackMaterial = $material }

            $sliceControlSet = [pscustomobject]@{
                Index = $sliceIndex
                ImageControl = $imageControl
                Overlay = $overlayTextBlock # This is part of the OriginalGrid
                OriginalGrid = $contentGrid
            }
            # Pre-initialize the state for every slice to prevent errors on first load.
            $SyncHash.PlayerState[$sliceIndex] = @{
                CurrentPath = $null; IsImage = $false; ImageTimer = $null; IsFailed = $false; RecoveryTimer = $null
                FfmpegProcess = $null; FfmpegTimer = $null; WriteableBmp = $null
                VisualBrush = $visualBrush # Store the brush for later manipulation
                OriginalGrid = $contentGrid # Store the original grid to restore it after an error
            }

            $SyncHash.SliceControls.Add($sliceControlSet)
        }
        $pieModelGroup.Children.Add($sliceModel)
    }

    $GetNextIndex = {
        $count = $SyncHash.SelectedFiles.Count
        if ($count -eq 0) { return $null }
        [System.Threading.Monitor]::Enter($SyncLock)
        try { $SyncHash.GlobalCounter = ($SyncHash.GlobalCounter + 1) % $count }
        finally { [System.Threading.Monitor]::Exit($SyncLock) }
        return $SyncHash.GlobalCounter
    }
    $SyncHash.GetNextIndex = $GetNextIndex

    $ApplyOverlayFor = {
        param($sliceIndex, [Uri]$uriOrNull)
        $overlay = $SyncHash.SliceControls[$sliceIndex].Overlay
        switch ($SyncHash.RbSelection) {
            "Hidden"   { $overlay.Visibility = 'Collapsed' }
            "Filename" { if ($uriOrNull) { $overlay.Text = [System.IO.Path]::GetFileName($uriOrNull.LocalPath) }; $overlay.Visibility = 'Visible' }
            "Custom"   { $overlay.Text = $SyncHash.CustomText; $overlay.Visibility = 'Visible' }
        }
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $overlay.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $overlay.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $overlay.FontSize = $SyncHash.FontSize
            if ($SyncHash.IsBold)  { $overlay.FontWeight = 'Bold' } else { $overlay.FontWeight = 'Normal' }
            if ($SyncHash.IsItalic){ $overlay.FontStyle  = 'Italic' } else { $overlay.FontStyle = 'Normal' }
        }
    }
    $SyncHash.ApplyOverlayFor = $ApplyOverlayFor

    $AssignNext = {
        param($sliceIndex)
        if ($sliceIndex -lt 0) { return }

        $st = $SyncHash.PlayerState[$sliceIndex]

        $controls = $SyncHash.SliceControls[$sliceIndex]
        $imageControl = $controls.ImageControl

        if ($st) {
            if ($st.ImageTimer)    { $st.ImageTimer.Stop();    $st.ImageTimer    = $null }
            if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop(); $st.RecoveryTimer = $null }
            if ($st.FfmpegProcess -and -not $st.FfmpegProcess.HasExited) { try { $st.FfmpegProcess.Kill() } catch {} }
            if ($st.FfmpegTimer)   { $st.FfmpegTimer.Stop(); $st.FfmpegTimer = $null }
        }

        $tries = 0; $maxTries = [Math]::Max(1, $SyncHash.SelectedFiles.Count)
        do {
            $next = & $SyncHash.GetNextIndex
            if ($next -eq $null) { return }
            $filePath = $SyncHash.SelectedFiles[$next]
            $uri = [Uri]$filePath
            $currentPath = if ($st) { $st.CurrentPath } else { $null }
            $sameAsCurrent = ($currentPath -eq $uri.LocalPath)
            $tries++
        } while ($sameAsCurrent -and $tries -lt $maxTries)

        $ext = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
        $st.CurrentPath = $uri.LocalPath
        $st.IsImage = ($SyncHash.ImageExtensions -contains $ext)
        $st.IsFailed = $false # Reset failure state

        # Restore the original grid visual in case this slice was previously in an error state.
        $st.VisualBrush.Visual = $controls.OriginalGrid

        if ($st.IsImage) {
            try {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $uri; $bitmap.EndInit()
                $imageControl.Source = $bitmap

                & $SyncHash.ApplyOverlayFor $sliceIndex $uri

                $t = New-Object System.Windows.Threading.DispatcherTimer
                $t.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds)
                $t.Tag = $sliceIndex
                $t.Add_Tick({ $timer = $args[0]; $idx = $timer.Tag; $timer.Stop(); & $SyncHash.AssignNext $idx })
                $st.ImageTimer = $t; $t.Start()
            } catch {
                & $SyncHash.HandleMediaFailure -SliceIndex $sliceIndex -Reason "Failed to load image."
            }
        } else {
            # Set overlay text just before starting the stream
            & $SyncHash.ApplyOverlayFor $sliceIndex $uri
            & $SyncHash.StartFfmpegStream -SliceIndex $sliceIndex -FilePath $uri.LocalPath
        }
    }
    $SyncHash.AssignNext = $AssignNext

    $StartFfmpegStream = {
        param($SliceIndex, $FilePath)
        $controls = $SyncHash.SliceControls[$SliceIndex]; $st = $SyncHash.PlayerState[$SliceIndex]
        $imageControl = $controls.ImageControl

        try {
            $psi_probe = New-Object System.Diagnostics.ProcessStartInfo -Property @{ FileName = "ffprobe.exe"; Arguments = "-v error -select_streams v:0 -show_entries stream=width,height -of csv=s=x:p=0 `"$FilePath`""; RedirectStandardOutput = $true; UseShellExecute = $false; CreateNoWindow = $true }
            $probe_proc = [System.Diagnostics.Process]::Start($psi_probe); $ffprobeOutput = $probe_proc.StandardOutput.ReadToEnd(); $probe_proc.WaitForExit()
            $width, $height = $ffprobeOutput.Trim() -split 'x'
            if (($probe_proc.ExitCode -ne 0) -or -not ($width -and $height -and [int]$width -gt 0 -and [int]$height -gt 0)) { throw "ffprobe failed or returned invalid dimensions." }

            # Apply overlay text only after ffprobe confirms the file is likely valid.
            & $SyncHash.ApplyOverlayFor $SliceIndex ([Uri]$FilePath)

            $st.WriteableBmp = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
            $imageControl.Source = $st.WriteableBmp

            $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{ FileName = "ffmpeg"; Arguments = "-hide_banner -loglevel error -i `"$FilePath`" -vf scale=${width}:${height} -f rawvideo -pix_fmt bgr24 -"; RedirectStandardOutput = $true; RedirectStandardError = $true; UseShellExecute = $false; CreateNoWindow = $true }
            $ffmpeg = [System.Diagnostics.Process]::Start($psi); $st.FfmpegProcess = $ffmpeg
            $stream = $ffmpeg.StandardOutput.BaseStream; $bytesPerFrame = [int]$width * [int]$height * 3; $buffer = New-Object byte[] $bytesPerFrame
            $rect = [System.Windows.Int32Rect]::new(0, 0, [int]$width, [int]$height); $stride = [int]$width * 3

            $frameTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(33); Tag = $SliceIndex }
            $frameTimer.Add_Tick({ $timer = $args[0]; $sIndex = $timer.Tag; $pState = $SyncHash.PlayerState[$sIndex]; $totalRead = 0; while ($totalRead -lt $bytesPerFrame) { $bytesRead = $stream.Read($buffer, $totalRead, $bytesPerFrame - $totalRead); if ($bytesRead -le 0) { $timer.Stop(); & $SyncHash.AssignNext $sIndex; return }; $totalRead += $bytesRead }; if ($totalRead -eq $bytesPerFrame) { $pState.WriteableBmp.Lock(); $pState.WriteableBmp.WritePixels($rect, $buffer, $stride, 0); $pState.WriteableBmp.Unlock() } }.GetNewClosure())
            $st.FfmpegTimer = $frameTimer; $frameTimer.Start()
        } catch { & $SyncHash.HandleMediaFailure -SliceIndex $SliceIndex -Reason $_.Exception.Message }
    }
    $SyncHash.StartFfmpegStream = $StartFfmpegStream

    $HandleMediaFailure = {
        param($SliceIndex, [string]$Reason)
        if ($SliceIndex -lt 0) { return }
        $controls = $SyncHash.SliceControls[$SliceIndex]; $st = $SyncHash.PlayerState[$SliceIndex]
        if ($st -and $st.IsFailed) { return }; if (-not $st) { return }
        $st.IsFailed = $true

        # Create a new, dedicated Grid for the error message to ensure it renders correctly.
        $errorGrid = New-Object System.Windows.Controls.Grid
        $errorGrid.Background = [System.Windows.Media.Brushes]::Black
        
        $errorTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            Foreground = [System.Windows.Media.Brushes]::Red
            HorizontalAlignment = 'Center'
            VerticalAlignment = 'Center'
            TextAlignment = 'Center'
            TextWrapping = 'Wrap'
        }
        $fileName = if ($st.CurrentPath) { [System.IO.Path]::GetFileName($st.CurrentPath) } else { "media" }
        $fileIndex = $SyncHash.SelectedFiles.IndexOf($st.CurrentPath)
        $fileNumberText = if ($fileIndex -ge 0) { " (File #$($fileIndex + 1))" } else { "" }
        $errorTextBlock.Text = "ERROR playing:`n$fileName$fileNumberText`n`n$Reason";

        # Create a fixed-size container for the text to prevent it from being stretched too large.
        $containerBorder = New-Object System.Windows.Controls.Border -Property @{
            Width = 256; Height = 256; Child = $errorTextBlock
        }
        $errorGrid.Children.Add($containerBorder)

        $st.VisualBrush.Visual = $errorGrid # Set the visual of the brush to our new error grid

        if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer; $recoveryTimer.Interval = [TimeSpan]::FromSeconds(10) # Show error for 10 seconds
        $recoveryTimer.Tag = $SliceIndex
        $recoveryTimer.Add_Tick({
            $timer = $args[0]; $sIndex = $timer.Tag; $timer.Stop()
            # Reset the failed state for this slice so AssignNext can proceed correctly.
            $SyncHash.PlayerState[$sIndex].IsFailed = $false
            # Restore the original grid so the next media can be displayed
            $SyncHash.PlayerState[$sIndex].VisualBrush.Visual = $SyncHash.SliceControls[$sIndex].OriginalGrid
            & $SyncHash.AssignNext $sIndex
        })
        $st.RecoveryTimer = $recoveryTimer; $recoveryTimer.Start() # Start the timer
    }
    $SyncHash.HandleMediaFailure = $HandleMediaFailure

    # Initial population of all slices
    for ($i = 0; $i -lt ($numberOfSlices * 2); $i++) { & $SyncHash.AssignNext $i }

    $window.Add_Loaded({
        # --- UI Controls and 3D Animation ---
        $mainGrid = $window.FindName("MainGrid")
        $SyncHash.AxisAngleX = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 0)
        $SyncHash.AxisAngleY = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), 0)
        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($SyncHash.AxisAngleX)))
        $transformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D($SyncHash.AxisAngleY)))
        $pieContainer.Transform = $transformGroup

        $SyncHash.AnimationX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(40)) -Property @{ RepeatBehavior = 'Forever' }
        $SyncHash.AnimationY = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(60)) -Property @{ RepeatBehavior = 'Forever' }
        $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.AnimationX)
        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.AnimationY)

        $SyncHash.pauseOrResumeAnimation = {
            if ($SyncHash.Paused) {
                $SyncHash.AnimationX.From = $SyncHash.AxisAngleX.Angle; $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.AnimationX)
                $SyncHash.AnimationY.From = $SyncHash.AxisAngleY.Angle; $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.AnimationY)
                $SyncHash.pauseButton.Content = "Pause"; $SyncHash.Paused = $false
            } else {
                $currentAngleX = $SyncHash.AxisAngleX.Angle; $currentAngleY = $SyncHash.AxisAngleY.Angle
                $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $SyncHash.AxisAngleX.Angle = $currentAngleX
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $SyncHash.AxisAngleY.Angle = $currentAngleY
                $SyncHash.pauseButton.Content = "Resume"; $SyncHash.Paused = $true
            }
        }

        $SyncHash.changeSpeed = {
            param([double]$multiplier)
            $newDurationX = [TimeSpan]::FromSeconds(($SyncHash.AnimationX.Duration.TimeSpan.TotalSeconds * $multiplier))
            $newDurationY = [TimeSpan]::FromSeconds(($SyncHash.AnimationY.Duration.TimeSpan.TotalSeconds * $multiplier))
            if ($newDurationX.TotalSeconds -lt 0.5) { $newDurationX = [TimeSpan]::FromSeconds(0.5) }
            if ($newDurationY.TotalSeconds -lt 0.5) { $newDurationY = [TimeSpan]::FromSeconds(0.5) }
            
            $SyncHash.AnimationX.Duration = $newDurationX
            $SyncHash.AnimationY.Duration = $newDurationY

            # If currently playing, pause and resume to apply the new speed
            if (-not $SyncHash.Paused) {
                & $SyncHash.pauseOrResumeAnimation # Pause
                # Use Dispatcher to invoke the resume on the next UI cycle
                $window.Dispatcher.InvokeAsync([action]{ & $SyncHash.pauseOrResumeAnimation }) | Out-Null
            }
        }

        $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
        $SyncHash.pauseButton = $window.FindName("pauseButton")
        $SyncHash.redoButton = $window.FindName("redoButton")
        $SyncHash.closeButton = $window.FindName("closeButton")
        $SyncHash.slowDownButton = $window.FindName("slowDownButton")
        $SyncHash.speedUpButton = $window.FindName("speedUpButton")

        $SyncHash.randomAxisButton.Add_Click({
            $SyncHash.AxisAngleX.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
            $SyncHash.AxisAngleY.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
        })
        $SyncHash.pauseButton.Add_Click($SyncHash.pauseOrResumeAnimation)
        $SyncHash.closeButton.Add_Click({ $window.Close() })
        $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $SyncHash.Window.Close() })
        $SyncHash.slowDownButton.Add_Click({ param($s, $e) $SyncHash.changeSpeed.Invoke(2.0) }) # Slower = longer duration
        $SyncHash.speedUpButton.Add_Click({  param($s, $e) $SyncHash.changeSpeed.Invoke(0.5) })  # Faster = shorter duration

        $mainGrid.Add_MouseDown($SyncHash.pauseOrResumeAnimation)
    })

    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "R"      { if ($SyncHash.redoButton) { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } }
            "P"      { if ($SyncHash.pauseButton) { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } }
            "A"      { if ($SyncHash.randomAxisButton) { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } }
            "Left"   { if ($SyncHash.slowDownButton) { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } }
            "Right"  { if ($SyncHash.speedUpButton) { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) } }
        }
    })

    $window.Add_Closed({
        if ($SyncHash.AnimationX) { $SyncHash.AnimationX.BeginTime = $null }
        if ($SyncHash.AnimationY) { $SyncHash.AnimationY.BeginTime = $null }
        foreach ($slice in $SyncHash.SliceControls) {
            $st = $SyncHash.PlayerState[$slice.Index]; if (-not $st) { continue }
            if ($st) {
                if ($st.ImageTimer)    { $st.ImageTimer.Stop() }
                if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
                if ($st.FfmpegTimer)   { $st.FfmpegTimer.Stop() }
                if ($st.FfmpegProcess -and -not $st.FfmpegProcess.HasExited) { try { $st.FfmpegProcess.Kill() } catch {} }
            }
        }
    })

    # Initialize overlays
    for ($i=0; $i -lt ($numberOfSlices * 2); $i++) { & $SyncHash.ApplyOverlayFor $i $null }

    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
