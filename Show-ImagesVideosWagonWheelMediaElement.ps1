<#
.SYNOPSIS
    Displays a playlist of media on the outer faces of a rotating 3D wagon wheel using WPF's MediaElement.
.DESCRIPTION
    This script creates a WPF window and renders a 3D wagon wheel composed of multiple
    "slice" or "spoke" geometries. It prompts the user to select image and video files,
    which are then displayed on the outer curved face of each slice. Each slice plays
    media from the playlist independently.

    This version uses the built-in Windows MediaElement for video playback. As a result, video
    format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI).

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosWagonWheelMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D wagon wheel window.
.NOTES
    Name:           Show-ImagesVideosWagonWheelMediaElement.ps1
    Version:        1.0.0, 11/04/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Wagon Wheel (MediaElement)"
$ScriptDescription = "Displays media on the outer curved faces of rotating 3D wagon wheel slices. Each slice plays media independently. Uses the built-in Windows MediaElement."

# --- Dependency Check ---
$RequiredExecutables = @() # No external executables needed

# --- Pie Slice Generation Function ---
function New-PieSliceModel {
    param(
        [double]$radius = 1.5,
        [double]$height = 0.5,
        [double]$startAngleDeg = 0,
        [double]$sliceAngleDeg = 45,
        [int]$segments = 8 # Segments for the curved outer face
    )

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $h2 = $height / 2.0

    # --- Vertices ---
    # 0,1: Center bottom, top
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, -$h2, 0))
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $h2, 0))

    # 2,3: Start angle edge (bottom, top)
    $startAngleRad = $startAngleDeg * [Math]::PI / 180.0
    $x0 = $radius * [Math]::Cos($startAngleRad)
    $z0 = $radius * [Math]::Sin($startAngleRad)
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x0, -$h2, $z0))
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x0, $h2, $z0))

    # Add vertices for the curved outer face
    $outerFaceStartIdx = $mesh.Positions.Count
    for ($i = 0; $i -le $segments; $i++) {
        $currentAngleRad = ($startAngleDeg + ($sliceAngleDeg * $i / $segments)) * [Math]::PI / 180.0
        $x = $radius * [Math]::Cos($currentAngleRad)
        $z = $radius * [Math]::Sin($currentAngleRad)
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, -$h2, $z)) # Bottom vertex
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $h2, $z))  # Top vertex
    }

    # --- Texture Coordinates ---
    # Add dummy UVs for non-textured faces
    for ($i = 0; $i -lt $outerFaceStartIdx; $i++) {
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0, 0))
    }
    # Add UVs for the curved outer face
    for ($i = 0; $i -le $segments; $i++) {
        # Invert the U-coordinate to flip the texture horizontally. This corrects the mirrored text.
        $u = 1 - ($i / $segments)
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u, 1)) # Bottom UV
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u, 0)) # Top UV
    }

    # --- Triangle Indices ---
    # Top face
    $mesh.TriangleIndices.Add(1); $mesh.TriangleIndices.Add($outerFaceStartIdx + 1); $mesh.TriangleIndices.Add(3)
    # Bottom face
    $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add($outerFaceStartIdx)
    # Side face 1 (start angle)
    $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(3); $mesh.TriangleIndices.Add(1)
    $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(3)
    # Side face 2 (end angle)
    $endIdx = $mesh.Positions.Count - 2
    $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(1); $mesh.TriangleIndices.Add($endIdx + 1)
    $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add($endIdx + 1); $mesh.TriangleIndices.Add($endIdx)

    # Curved outer face
    for ($i = 0; $i -lt $segments; $i++) {
        $idx0 = $outerFaceStartIdx + ($i * 2)       # Current bottom
        $idx1 = $outerFaceStartIdx + ($i * 2) + 1   # Current top
        $idx2 = $outerFaceStartIdx + ($i * 2) + 2   # Next bottom
        $idx3 = $outerFaceStartIdx + ($i * 2) + 3   # Next top

        $mesh.TriangleIndices.Add($idx0); $mesh.TriangleIndices.Add($idx1); $mesh.TriangleIndices.Add($idx3)
        $mesh.TriangleIndices.Add($idx0); $mesh.TriangleIndices.Add($idx3); $mesh.TriangleIndices.Add($idx2)
    }

    # --- Create Models ---
    $sliceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D
    $sliceModel.Geometry = $mesh

    # The main material will be for the outer face.
    # We need a separate material for the other faces.
    $otherFacesMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::DarkSlateGray)

    # Create a separate model for the non-textured faces
    $otherFacesMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $otherFacesMesh.Positions = $mesh.Positions
    $otherFacesMesh.TextureCoordinates = $mesh.TextureCoordinates
    # Add only the indices for top, bottom, and side faces
    for ($i = 0; $i -lt 24; $i++) {
        $otherFacesMesh.TriangleIndices.Add($mesh.TriangleIndices[$i])
    }
    $otherFacesModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($otherFacesMesh, $otherFacesMaterial)

    # Create a model just for the textured outer face
    $outerFaceMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $outerFaceMesh.Positions = $mesh.Positions
    $outerFaceMesh.TextureCoordinates = $mesh.TextureCoordinates
    # Add only the indices for the outer face
    for ($i = 24; $i -lt $mesh.TriangleIndices.Count; $i++) {
        $outerFaceMesh.TriangleIndices.Add($mesh.TriangleIndices[$i])
    }
    $outerFaceModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($outerFaceMesh, $null) # Material will be set later

    # Group them together
    $modelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
    $modelGroup.Children.Add($otherFacesModel)
    $modelGroup.Children.Add($outerFaceModel)

    return @{
        OuterFaceModel = $outerFaceModel
        FullSliceModel = $modelGroup
    }
}

# --- Main Application Loop ---
while ($true) {

    # -------------------- File Selection Form --------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Pie Slice Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
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
        AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65
    }
    $SelectForm.Controls.Add($DataGridView)

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    [void]$DataGridView.Columns.Add($CheckBoxColumn)
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true }
    [void]$DataGridView.Columns.Add($FileNameColumn)
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true }
    [void]$DataGridView.Columns.Add($FilePathColumn)

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Play Selected"; Location = '600, 40'; Size = '170, 30' }
    $SelectForm.Controls.Add($PlayButton)

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $isPreviewPaused = $false
            $row = $DataGridView.Rows[$e.RowIndex]
            $videoPath = $row.Cells["FilePath"].Value

            [xml]$previewXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preview - Click to Pause/Resume" Height="450" Width="800"
        WindowStartupLocation="CenterScreen" WindowStyle="ToolWindow" Background="Black">
    <Grid x:Name="TheGrid">
        <MediaElement x:Name="MediaPlayer" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
    </Grid>
</Window>
"@
            $previewReader = (New-Object System.Xml.XmlNodeReader $previewXaml)
            $previewWindow = [Windows.Markup.XamlReader]::Load($previewReader)
            $previewGrid = $previewWindow.FindName("TheGrid")
            $previewPlayer = $previewWindow.FindName("MediaPlayer")
            $previewPlayer.Source = [Uri]$videoPath
            $previewGrid.Add_MouseDown({
                if ($isPreviewPaused) { $previewPlayer.Play(); $isPreviewPaused = $false } else { $previewPlayer.Pause(); $isPreviewPaused = $true }
            })
            $previewPlayer.Add_MediaEnded({ $previewPlayer.Position = [TimeSpan]::Zero; $previewPlayer.Play() })
            $previewPlayer.Play(); $previewWindow.ShowDialog() | Out-Null
        }
    })

    $DataGridView.Add_CellPainting({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -lt 0) {
            $e.PaintBackground($e.ClipBounds, $true)
            $fmt = New-Object System.Drawing.StringFormat
            $fmt.Alignment = [System.Drawing.StringAlignment]::Center
            $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
            $rectF = New-Object System.Drawing.RectangleF($e.CellBounds.X, $e.CellBounds.Y, $e.CellBounds.Width, $e.CellBounds.Height)
            $e.Graphics.DrawString($e.FormattedValue.ToString(), $e.CellStyle.Font, [System.Drawing.Brushes]::Black, $rectF, $fmt)
            $e.Handled = $true
        }
    })

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
      UseTransparentEffect = $false
      SelectedFiles        = @()
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

            $ImageExtensionsScan = "*.bmp","*.jpeg","*.jpg","*.png","*.tif","*.tiff","*.gif","*.wmp","*.ico"
            $VideoExtensionsScan = "*.webm","*.mkv","*.flv","*.vob","*.ogv","*.ogg","*.mov","*.avi","*.qt","*.wmv","*.yuv","*.rm","*.asf","*.amv","*.mp4","*.m4p","*.m4v","*.mpg","*.mp2","*.mpeg","*.mpe","*.mpv","*.svi","*.3gp","*.3g2","*.mxf","*.roq","*.nsv","*.f4v","*.f4p","*.f4a","*.f4b"
            $AllowedExtensions = $ImageExtensionsScan + $VideoExtensionsScan

            $gciParams = @{ File = $true; Include = $AllowedExtensions }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true }
            else { $gciParams.Path = Join-Path $SelectedPath "*" }
            $files = Get-ChildItem @gciParams
            foreach ($file in $files) { [void]$DataGridView.Rows.Add($false, $file.Name, $file.FullName) }
            foreach ($row in $DataGridView.Rows) { if (-not $row.IsNewRow) { $row.HeaderCell.Value = "Play" } }
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
        Title="Rotating Pie"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
  <Grid x:Name="MainGrid">
    <Viewport3D x:Name="mainViewport">
      <Viewport3D.Camera>
        <PerspectiveCamera Position="0,0,8" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
      </Viewport3D.Camera>
      <ModelVisual3D x:Name="PieContainer">
        <ModelVisual3D.Content>
          <Model3DGroup>
            <AmbientLight Color="Gray"/>
            <DirectionalLight Color="White" Direction="-1,-1,-2"/>
            <DirectionalLight Color="White" Direction="1,1,2"/>
          </Model3DGroup>
        </ModelVisual3D.Content>
        <ModelVisual3D.Transform>
          <Transform3DGroup>
            <RotateTransform3D>
              <RotateTransform3D.Rotation>
                <AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/>
              </RotateTransform3D.Rotation>
            </RotateTransform3D>
            <RotateTransform3D>
              <RotateTransform3D.Rotation>
                <AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/>
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

    $fontSz  = 24
    try { if ($script:formState.FontSize -and [double]$script:formState.FontSize -gt 0) { $fontSz = [double]$script:formState.FontSize } } catch { }

    $SyncLock = New-Object object
    $SyncHash = [pscustomobject]@{
        SelectedFiles = $script:formState.SelectedFiles; UseTransparentEffect = $script:formState.UseTransparentEffect
        Paused = $false; ControlsHidden = $false; RedoClicked = $false
        RbSelection = $script:formState.RbSelection; CustomText = $script:formState.CustomText; TextColor = $script:formState.TextColor;
        FontSize = $fontSz; FontFamily = $script:formState.FontFamily; IsBold = $script:formState.IsBold; IsItalic = $script:formState.IsItalic;
        PlayerState = @{}; ImageExtensions = @(".bmp",".jpeg",".jpg",".png",".tif",".tiff",".gif",".wmp",".ico")
        GlobalCounter = -1; ImageHoldSeconds = 10
        Window = $null; ImageControls = (New-Object 'System.Collections.Generic.List[System.Windows.Controls.MediaElement]')
        OverlayTextBlocks = (New-Object 'System.Collections.Generic.List[System.Windows.Controls.TextBlock]'); SliceOuterFaceModels = $null
        GetNextIndex = $null; ApplyOverlayFor = $null; AssignNext = $null; HandleMediaFailure = $null
        animX = $null; animY = $null; AxisAngleX = $null; AxisAngleY = $null
        pauseButton = $null; randomAxisButton = $null; slowDownButton = $null; speedUpButton = $null; redoButton = $null; hideControlsButton = $null; closeButton = $null
    }

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $SyncHash.Window = $window

    $pieContainer = $window.FindName("PieContainer")

    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top
    
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
    $numberOfSlices = 8
    $sliceAngle = 360.0 / $numberOfSlices
    $sliceOuterFaceModels = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.GeometryModel3D]

    # --- Dynamic Radius Calculation ---
    # Calculate the radius so the wheel's diameter fits within a percentage of the screen's working height.
    $camera = $window.FindName("mainViewport").Camera
    $cameraDistance = $camera.Position.Z
    $cameraFovRadians = $camera.FieldOfView * ([Math]::PI / 180.0)
    $visibleHeightAtOrigin = 2.0 * $cameraDistance * [Math]::Tan($cameraFovRadians / 2.0)
    $desiredHeightPercentage = 0.50 # Use 50% of the viewport height
    $dynamicRadius = ($visibleHeightAtOrigin * $desiredHeightPercentage) / 2.0

    for ($i = 0; $i -lt $numberOfSlices; $i++) {
        $startAngle = $i * $sliceAngle
        $sliceParts = New-PieSliceModel -radius $dynamicRadius -startAngleDeg $startAngle -sliceAngleDeg $sliceAngle
        $sliceOuterFaceModels.Add($sliceParts.OuterFaceModel)
        
        # Use a MediaElement for both images and videos
        $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
            Stretch = 'Fill'; LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Visibility = 'Collapsed'
        }
        $SyncHash.ImageControls.Add($mediaElement) # Reusing this list name for consistency

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; Margin = '10,0,10,0'
            TextAlignment = 'Center'; IsHitTestVisible = $false;
        }
        $SyncHash.OverlayTextBlocks.Add($overlayTextBlock)

        $grid = New-Object System.Windows.Controls.Grid
        $grid.Background = [System.Windows.Media.Brushes]::Black
        [void]$grid.Children.Add($mediaElement)
        [void]$grid.Children.Add($overlayTextBlock)

        $visualBrush = New-Object System.Windows.Media.VisualBrush
        $visualBrush.Visual = $grid

        $material = New-Object $materialType
        $material.Brush = $visualBrush
        if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

        $sliceParts.OuterFaceModel.Material = $material
        $pieContainer.Content.Children.Add($sliceParts.FullSliceModel)
    }
    $SyncHash.SliceOuterFaceModels = $sliceOuterFaceModels

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
        $overlay = $SyncHash.OverlayTextBlocks[$sliceIndex]
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
            if ($SyncHash.IsBold)  { $overlay.FontWeight = 'Bold' }
            if ($SyncHash.IsItalic){ $overlay.FontStyle  = 'Italic' }
        }
    }
    $SyncHash.ApplyOverlayFor = $ApplyOverlayFor

    $AssignNext = {
        param($sliceIndex)
        if ($sliceIndex -lt 0) { return }
        $mediaElement = $SyncHash.ImageControls[$sliceIndex]

        $st = $SyncHash.PlayerState[$mediaElement.GetHashCode()]
        if ($st) {
            if ($st.ImageTimer)    { $st.ImageTimer.Stop();    $st.ImageTimer    = $null }
            if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop(); $st.RecoveryTimer = $null }
        }
        $mediaElement.Stop()

        # Detach old event handlers to prevent memory leaks and duplicate triggers
        if ($st.MediaOpenedHandler) { $mediaElement.remove_MediaOpened($st.MediaOpenedHandler) }
        if ($st.MediaEndedHandler)  { $mediaElement.remove_MediaEnded($st.MediaEndedHandler) }
        if ($st.MediaFailedHandler) { $mediaElement.remove_MediaFailed($st.MediaFailedHandler) }

        # --- Event Handlers for this specific MediaElement instance ---
        $mediaOpenedHandler = {
            param($sender, $e)
            $me = $sender; $sIndex = $me.Tag; $pState = $SyncHash.PlayerState[$me.GetHashCode()]
            if (-not $pState) { return }
            $pState.IsFailed = $false; $pState.PlaybackStopwatch.Restart(); $me.Visibility = 'Visible'
            if ($pState.IsImage) {
                $me.Pause(); if ($pState.ImageTimer) { $pState.ImageTimer.Stop() }
                $t = New-Object System.Windows.Threading.DispatcherTimer; $t.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds); $t.Tag = $sIndex
                $t.Add_Tick({ $timer = $args[0]; $idx = $timer.Tag; $timer.Stop(); & $SyncHash.AssignNext $idx }); $pState.ImageTimer = $t; $t.Start()
            } elseif (-not $me.NaturalDuration.HasTimeSpan) { & $SyncHash.HandleMediaFailure -SliceIndex $sIndex -Reason "No duration found (silent failure)." }
        }
        $mediaEndedHandler = {
            $me = $args[0]; $sIndex = $me.Tag; if ($sIndex -lt 0) { return }
            $pState = $SyncHash.PlayerState[$me.GetHashCode()]; if ($pState -and -not $pState.IsImage -and $pState.PlaybackStopwatch) { $pState.PlaybackStopwatch.Stop(); if ($pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 2000) { & $SyncHash.HandleMediaFailure -SliceIndex $sIndex -Reason "Playback failed instantly."; return } }
            & $SyncHash.AssignNext $sIndex
        }
        $mediaFailedHandler = { param($sender, $e) & $SyncHash.HandleMediaFailure -SliceIndex $sender.Tag -Reason $e.ErrorException.Message }


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

        & $SyncHash.ApplyOverlayFor $sliceIndex $uri

        $ext = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
        if (-not $st) {
            $st = @{ CurrentPath = $null; IsImage = $false; ImageTimer = $null; IsFailed = $false; RecoveryTimer = $null; PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch }
            $SyncHash.PlayerState[$mediaElement.GetHashCode()] = $st
        }
        $st.CurrentPath = $uri.LocalPath
        $st.IsImage = ($SyncHash.ImageExtensions -contains $ext)
        $st.MediaOpenedHandler = $mediaOpenedHandler.GetNewClosure()
        $st.MediaEndedHandler = $mediaEndedHandler.GetNewClosure()
        $st.MediaFailedHandler = $mediaFailedHandler.GetNewClosure()
        $mediaElement.Tag = $sliceIndex
        $mediaElement.Add_MediaOpened($st.MediaOpenedHandler)
        $mediaElement.Add_MediaEnded($st.MediaEndedHandler)
        $mediaElement.Add_MediaFailed($st.MediaFailedHandler)

        $mediaElement.Visibility = 'Collapsed'
        $mediaElement.Source = $uri
        $mediaElement.Play()
    }
    $SyncHash.AssignNext = $AssignNext

    $HandleMediaFailure = {
        param($SliceIndex, [string]$Reason)
        if ($SliceIndex -lt 0) { return }
        $mediaElement = $SyncHash.ImageControls[$SliceIndex]; $st = $SyncHash.PlayerState[$mediaElement.GetHashCode()]
        if ($st -and $st.IsFailed) { return }; if (-not $st) { return }
        $st.IsFailed = $true
        
        $mediaElement.Visibility = 'Collapsed'
        $overlay = $SyncHash.OverlayTextBlocks[$SliceIndex]
        $fileName = if ($st.CurrentPath) { [System.IO.Path]::GetFileName($st.CurrentPath) } else { "media" }
        $overlay.Text = "ERROR playing:`n$fileName`n`n$Reason";
        $overlay.Foreground = [System.Windows.Media.Brushes]::Red; $overlay.Visibility = 'Visible'

        if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer; $recoveryTimer.Interval = [TimeSpan]::FromSeconds(5)
        $recoveryTimer.Tag = $SliceIndex
        $recoveryTimer.Add_Tick({
            $timer = $args[0]; $sIndex = $timer.Tag; $timer.Stop();
            $s2 = $SyncHash.PlayerState[$SyncHash.ImageControls[$sIndex].GetHashCode()]; if ($s2) { $s2.IsFailed = $false }
            & $SyncHash.AssignNext $sIndex
        });
        $st.RecoveryTimer = $recoveryTimer; $recoveryTimer.Start()
    }
    $SyncHash.HandleMediaFailure = $HandleMediaFailure

    for ($i = 0; $i -lt $SyncHash.ImageControls.Count; $i++) { & $SyncHash.AssignNext $i }

    $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20)); $animX.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleX = $window.FindName("AxisAngleX"); $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)
    $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15)); $animY.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleY = $window.FindName("AxisAngleY"); $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
    $SyncHash.animX = $animX; $SyncHash.animY = $animY; $SyncHash.AxisAngleX = $axisAngleX; $SyncHash.AxisAngleY = $axisAngleY

    $SyncHash.pauseButton = $window.FindName("pauseButton"); $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton"); $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton"); $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        if ($SyncHash.Paused) {
            $SyncHash.animX.From = $SyncHash.AxisAngleX.Angle; $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animX)
            $SyncHash.animY.From = $SyncHash.AxisAngleY.Angle; $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animY)
            $SyncHash.pauseButton.Content = "Pause"; $SyncHash.Paused = $false
        } else {
            $currentAngleX = $SyncHash.AxisAngleX.Angle; $currentAngleY = $SyncHash.AxisAngleY.Angle
            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $SyncHash.AxisAngleX.Angle = $currentAngleX; $SyncHash.AxisAngleY.Angle = $currentAngleY
            $SyncHash.pauseButton.Content = "Resume"; $SyncHash.Paused = $true
        }
    })

    $SyncHash.randomAxisButton.Add_Click({
        $SyncHash.AxisAngleX.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
        $SyncHash.AxisAngleY.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
    })

    $changeSpeed = {
        param($multiplier)
        $newDurationX = [TimeSpan]::FromSeconds(($SyncHash.animX.Duration.TimeSpan.TotalSeconds * $multiplier))
        $newDurationY = [TimeSpan]::FromSeconds(($SyncHash.animY.Duration.TimeSpan.TotalSeconds * $multiplier))
        if ($newDurationX.TotalSeconds -lt 0.5) { $newDurationX = [TimeSpan]::FromSeconds(0.5) }
        if ($newDurationY.TotalSeconds -lt 0.5) { $newDurationY = [TimeSpan]::FromSeconds(0.5) }
        $SyncHash.animX.Duration = $newDurationX; $SyncHash.animY.Duration = $newDurationY
        if (-not $SyncHash.Paused) {
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            Start-Sleep -Milliseconds 50
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        }
    }
    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })

    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $SyncHash.Window.Close() })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        if ($SyncHash.ControlsHidden) { $controlsPanel.Visibility = 'Visible'; $SyncHash.ControlsHidden = $false }
        else { $controlsPanel.Visibility = 'Collapsed'; $SyncHash.ControlsHidden = $true }
    })

    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P"      { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "A"      { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R"      { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H"      { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left"   { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right"  { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })
    $mainGrid = $window.FindName("MainGrid")
    $mainGrid.Add_MouseDown({ $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) })

    $window.Add_Closed({
        foreach ($imageControl in $SyncHash.ImageControls) {
            $st = $SyncHash.PlayerState[$imageControl.GetHashCode()];
            if ($st) {
                if ($st.ImageTimer) { $st.ImageTimer.Stop() }
                if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
                if ($st.MediaOpenedHandler) { $imageControl.remove_MediaOpened($st.MediaOpenedHandler) }
                if ($st.MediaEndedHandler)  { $imageControl.remove_MediaEnded($st.MediaEndedHandler) }
                if ($st.MediaFailedHandler) { $imageControl.remove_MediaFailed($st.MediaFailedHandler) }
            }
            $imageControl.Source = $null
        }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
    })

    for ($i=0; $i -lt $SyncHash.OverlayTextBlocks.Count; $i++) { & $SyncHash.ApplyOverlayFor $i $null }

    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
