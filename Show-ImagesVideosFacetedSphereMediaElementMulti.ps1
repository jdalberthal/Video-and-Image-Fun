<#
.SYNOPSIS
    Displays a playlist of media on an interactive, rotating, faceted 3D sphere.
.DESCRIPTION
    This script creates a WPF window and renders a 3D sphere. The sphere's geometry is
    programmatically generated with flat facets. It then prompts the user to select an image
    or video file, which is displayed on each facet of the sphere.

    The sphere has a continuous rotation animation and includes UI controls to pause, change speed, and randomize the rotation.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosFacetedSphereMediaElementMulti.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D faceted sphere window.
.NOTES
    Name:           Show-ImagesVideosFacetedSphereMediaElementMulti.ps1
    Version:        1.0.0, 10/25/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Rotating Faceted Sphere"
$ScriptDescription  = "Displays selected images/videos on a rotating 3D sphere; each facet advances independently in queue order."

# --- Sphere Generation Function ---
function New-SphereMesh {
    param(
        [double]$radius = 1.5,
        [int]$slices = 8, # Longitude
        [int]$stacks = 3  # Latitude
    )
    $facets = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.GeometryModel3D]

    $allVertices = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.Point3D]
    for ($stack = 0; $stack -le $stacks; $stack++) {
        $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks
        $y = $radius * [Math]::Sin($phi)
        $r = $radius * [Math]::Cos($phi)
        for ($slice = 0; $slice -le $slices; $slice++) {
            $theta = $slice * 2 * [Math]::PI / $slices
            $x = $r * [Math]::Cos($theta)
            $z = $r * [Math]::Sin($theta)
            $allVertices.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
        }
    }

    for ($stack = 0; $stack -lt $stacks; $stack++) {
        for ($slice = 0; $slice -lt $slices; $slice++) {
            $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
            $i0 = $stack * ($slices + 1) + $slice
            $i1 = ($stack + 1) * ($slices + 1) + $slice
            $i2 = $i0 + 1
            $i3 = $i1 + 1
            $p0 = $allVertices[$i0]; $p1 = $allVertices[$i1]; $p2 = $allVertices[$i2]; $p3 = $allVertices[$i3]

            # Two triangles per facet
            $mesh.Positions.Add($p0); $mesh.Positions.Add($p1); $mesh.Positions.Add($p2)
            $mesh.Positions.Add($p2); $mesh.Positions.Add($p1); $mesh.Positions.Add($p3)

            # UVs
            $uv0 = [System.Windows.Point]::new(0,0)
            $uv1 = [System.Windows.Point]::new(0,1)
            $uv2 = [System.Windows.Point]::new(1,0)
            $uv3 = [System.Windows.Point]::new(1,1)
            $mesh.TextureCoordinates.Add($uv0); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv2)
            $mesh.TextureCoordinates.Add($uv2); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv3)

            $facetModel = New-Object System.Windows.Media.Media3D.GeometryModel3D($mesh, $null)
            $facets.Add($facetModel)
        }
    }
    return $facets
}

# --- Main Application Loop ---
while ($true) {

    # -------------------- File Selection Form --------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Faceted Sphere - Media Selector"
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

    # ---- Preview XAML for row header click ----
    [xml]$VideoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preview Video - Click to Pause/Resume" Height="450" Width="800"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        SizeToContent="Manual" WindowState="Normal" WindowStyle="ToolWindow" Background="Black">
  <Grid x:Name="TheGrid">
    <MediaElement x:Name="MediaPlayer" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
  </Grid>
</Window>
"@

    $DataGridView.Add_RowHeaderMouseClick({
        $rowIndex = $_.RowIndex
        $row = $DataGridView.Rows[$rowIndex]
        $videoPath = ($row.Cells["FilePath"].Value)

        $VideoReader = (New-Object System.Xml.XmlNodeReader $VideoXaml)
        $VideoWindow = [Windows.Markup.XamlReader]::Load($VideoReader)
        $TheGrid = $VideoWindow.FindName("TheGrid")
        $MediaPlayer = $VideoWindow.FindName("MediaPlayer")
        $MediaPlayer.Source = [Uri]$videoPath

        $PreviewPaused = $false
        $TheGrid.Add_MouseDown({
            if ($PreviewPaused) { $MediaPlayer.Play(); $PreviewPaused = $false } else { $MediaPlayer.Pause(); $PreviewPaused = $true }
        })
        $MediaPlayer.Add_MediaEnded({ $MediaPlayer.Position = [TimeSpan]::Zero; $MediaPlayer.Play() })
        $MediaPlayer.Play()
        $VideoWindow.ShowDialog() | Out-Null
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
            # Initialize the dialog with the current font settings
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
                # The updateTextBoxFont call is now implicitly triggered by the ValueChanged/CheckedChanged events
                & $updateTextBoxFont
            }
        })

    # Centralized function to update the TextBox font based on current selections
    $updateTextBoxFont = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }

        try {
            $newFont = New-Object System.Drawing.Font($script:formState.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
        } catch {
            # Fallback to a default font if the selected one fails for any reason
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

            # Wildcards for filesystem scan
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
            # TextColor and FontFamily are already updated by their respective dialogs

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

    # Exit if no files were selected or the form was closed
    if ($script:formState.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break
    }

    # -------------------- Main WPF Window --------------------
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rotating Sphere"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
  <Grid x:Name="MainGrid">
    <Viewport3D x:Name="mainViewport">
      <Viewport3D.Camera>
        <PerspectiveCamera Position="0,0,8" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
      </Viewport3D.Camera>
      <ModelVisual3D x:Name="SphereContainer">
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

    # --- Safe font size for SyncHash ---
    $fontSz  = 24
    try {
        if ($script:formState.FontSize -and [double]$script:formState.FontSize -gt 0) {
            $fontSz = [double]$script:formState.FontSize
        }
    } catch { $fontSz = 24 }

    # Separate lock for atomic increments
    $SyncLock = New-Object object

    # --- Shared state (pscustomobject so dot-notation works) ---
    $SyncHash = [pscustomobject]@{
        # Inputs
        SelectedFiles        = $script:formState.SelectedFiles
        UseTransparentEffect = $script:formState.UseTransparentEffect

        # UI state
        Paused               = $false
        ControlsHidden       = $false
        RedoClicked          = $false

        # Text overlay
        RbSelection          = $script:formState.RbSelection
        CustomText           = $script:formState.CustomText
        TextColor            = $script:formState.TextColor
        FontSize             = $fontSz
        FontFamily           = $script:formState.FontFamily
        IsBold               = $script:formState.IsBold
        IsItalic             = $script:formState.IsItalic

        # Per-facet
        PlayerState          = @{}  # facet MediaElement hashcode -> state
        ImageExtensions      = @(".bmp",".jpeg",".jpg",".png",".tif",".tiff",".gif",".wmp",".ico")

        # Queue state
        GlobalCounter        = -1           # first increment -> 0
        ImageHoldSeconds     = 10

        # Handles & collections
        Window               = $null
        MediaPlayers         = (New-Object 'System.Collections.Generic.List[System.Windows.Controls.MediaElement]')
        OverlayTextBlocks    = (New-Object 'System.Collections.Generic.List[System.Windows.Controls.TextBlock]')

        # Delegates / functions (filled later)
        GetNextIndex         = $null
        ApplyOverlayFor      = $null
        AssignNext           = $null
        HandleMediaFailure   = $null
        MediaEndedHandler    = $null

        # Anim & transforms
        animX                = $null
        animY                = $null
        AxisAngleX           = $null
        AxisAngleY           = $null

        # Control refs
        pauseButton          = $null
        randomAxisButton     = $null
        slowDownButton       = $null
        speedUpButton        = $null
        redoButton           = $null
        hideControlsButton   = $null
        closeButton          = $null
    }

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $SyncHash.Window = $window

    $sphereContainer = $window.FindName("SphereContainer")

    # --- Geometry ---
    $cameraDistance = 8.0
    $cameraFovRadians = (60.0 * [Math]::PI / 180.0)
    $visibleHeightAtOrigin = 2.0 * $cameraDistance * [Math]::Tan($cameraFovRadians / 2.0)
    $dynamicRadius = ($visibleHeightAtOrigin * 0.50) / 2.0
    $facetModels = New-SphereMesh -radius $dynamicRadius

    # --- Fullscreen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    # --- Visuals per facet (static) ---
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

    foreach ($facetModel in $facetModels) {
        $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
            LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; ScrubbingEnabled = $true
        }
        $SyncHash.MediaPlayers.Add($mediaElement)

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; Margin = '10,0,10,0'
            TextAlignment = 'Center'; IsHitTestVisible = $false
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

        $facetModel.Material = $material
        $sphereContainer.Content.Children.Add($facetModel)
    }

    # --- Helper: next index (thread-safe via Monitor lock on $SyncLock) ---
    $GetNextIndex = {
        $count = $SyncHash.SelectedFiles.Count
        if ($count -eq 0) { return $null }

        [System.Threading.Monitor]::Enter($SyncLock)
        try {
            $SyncHash.GlobalCounter = $SyncHash.GlobalCounter + 1
            $idx = $SyncHash.GlobalCounter % $count
        }
        finally {
            [System.Threading.Monitor]::Exit($SyncLock)
        }
        return $idx
    }
    $SyncHash.GetNextIndex = $GetNextIndex

    # --- Helper: apply overlay text for a player index ---
    $ApplyOverlayFor = {
        param($facetIndex, [Uri]$uriOrNull)
        $overlay = $SyncHash.OverlayTextBlocks[$facetIndex]
        switch ($SyncHash.RbSelection) {
            "Hidden"   { $overlay.Visibility = 'Collapsed' }
            "Filename" {
                if ($uriOrNull) { $overlay.Text = [System.IO.Path]::GetFileName($uriOrNull.LocalPath) }
                $overlay.Visibility = 'Visible'
            }
            "Custom"   { $overlay.Text = $SyncHash.CustomText; $overlay.Visibility = 'Visible' }
        }
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $overlay.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $overlay.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)

            $safeFont = 24
            try {
                if ($SyncHash.FontSize -and [double]$SyncHash.FontSize -gt 0) { $safeFont = [double]$SyncHash.FontSize }
            } catch { $safeFont = 24 }
            $overlay.FontSize = $safeFont

            if ($SyncHash.IsBold)  { $overlay.FontWeight = 'Bold' }
            if ($SyncHash.IsItalic){ $overlay.FontStyle  = 'Italic' }
        }
    }
    $SyncHash.ApplyOverlayFor = $ApplyOverlayFor

    # --- Assign next media to a specific player (ONLY this facet updates) ---
    $AssignNext = {
        param([System.Windows.Controls.MediaElement]$player)
        if (-not $player) { return }
        $facetIndex = $SyncHash.MediaPlayers.IndexOf($player)
        if ($facetIndex -lt 0) { return }

        # Stop any existing facet-local timers
        $st = $SyncHash.PlayerState[$player.GetHashCode()]
        if ($st) {
            if ($st.ImageTimer)    { $st.ImageTimer.Stop();    $st.ImageTimer    = $null }
            if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop(); $st.RecoveryTimer = $null }
        }

        # Choose next file; avoid immediately reusing the same file on this facet (unless there's only one)
        $tries = 0
        $maxTries = [Math]::Max(1, $SyncHash.SelectedFiles.Count)
        do {
            $next = & $SyncHash.GetNextIndex
            if ($next -eq $null) { return }
            $filePath = $SyncHash.SelectedFiles[$next]
            $uri = [Uri]$filePath

            $sameAsCurrent = $false
            if ($player.Source) {
                $sameAsCurrent = ($player.Source.IsFile -and ($player.Source.LocalPath -eq $uri.LocalPath))
            }
            $tries++
        } while ($sameAsCurrent -and $tries -lt $maxTries)

        $player.Stop()
        $player.Source = $uri
        & $SyncHash.ApplyOverlayFor $facetIndex $uri

        $ext = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
        if (-not $st) {
            $st = @{
                IsImage        = $false
                ImageTimer     = $null
                IsFailed       = $false
                RecoveryTimer  = $null
            }
            $SyncHash.PlayerState[$player.GetHashCode()] = $st
        }
        $st.IsImage = ($SyncHash.ImageExtensions -contains $ext)

        if ($st.IsImage) {
            # MediaOpened pauses and starts per-facet timer
        } else {
            $player.Play()
        }
    }
    $SyncHash.AssignNext = $AssignNext

    # --- Failure handler: advance this facet to next item ---
    $HandleMediaFailure = {
        param($FailedPlayer, [string]$Reason)
        if (-not $FailedPlayer) { return }
        $st = $SyncHash.PlayerState[$FailedPlayer.GetHashCode()]
        if ($st -and $st.IsFailed) { return }
        if (-not $st) { return }

        $st.IsFailed = $true
        $facetIndex = $SyncHash.MediaPlayers.IndexOf($FailedPlayer)
        if ($facetIndex -ge 0) {
            $fileName = if ($FailedPlayer.Source) { [System.IO.Path]::GetFileName($FailedPlayer.Source.LocalPath) } else { "unknown media" }
            $ov = $SyncHash.OverlayTextBlocks[$facetIndex]
            $ov.Text = "ERROR`n$fileName`n$Reason"
            $ov.Visibility = 'Visible'
            $FailedPlayer.Visibility = 'Collapsed'
        }

        if ($st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
        $recoveryTimer.Interval = [TimeSpan]::FromSeconds(2)
        $recoveryTimer.Tag = $FailedPlayer
        $recoveryTimer.Add_Tick({
            $timer = $args[0]; $pl = $timer.Tag; $timer.Stop()
            $s2 = $SyncHash.PlayerState[$pl.GetHashCode()]
            $s2.IsFailed = $false
            $pl.Visibility = 'Visible'
            & $SyncHash.AssignNext $pl
        })
        $st.RecoveryTimer = $recoveryTimer
        $recoveryTimer.Start()
    }
    $SyncHash.HandleMediaFailure = $HandleMediaFailure

    # --- Media event handlers ---
    $MediaEndedHandler = {
        param($Sender, $EventArgs)
        $finishedPlayer = $Sender
        if (-not $finishedPlayer) { return }
        $st = $SyncHash.PlayerState[$finishedPlayer.GetHashCode()]
        if ($st -and $st.ImageTimer) { $st.ImageTimer.Stop(); $st.ImageTimer = $null }
        if ($st -and $st.IsFailed) { return }
        # Advance ONLY this facet
        & $SyncHash.AssignNext $finishedPlayer
    }
    $SyncHash.MediaEndedHandler = $MediaEndedHandler

    $MediaOpenedHandler = {
        param($Sender, $EventArgs)
        $player = $Sender
        if (-not $player) { return }
        $st = $SyncHash.PlayerState[$player.GetHashCode()]
        if (-not $st) {
            $st = @{
                IsImage        = $false
                ImageTimer     = $null
                IsFailed       = $false
                RecoveryTimer  = $null
            }
            $SyncHash.PlayerState[$player.GetHashCode()] = $st
        }
        $st.IsFailed = $false

        $facetIndex = $SyncHash.MediaPlayers.IndexOf($player)
        if ($facetIndex -ge 0) {
            $player.Visibility = 'Visible'
            # Overlay text already applied in AssignNext
        }

        if ($st.IsImage) {
            # Hold current image for N seconds, per facet
            $player.Pause()
            if ($st.ImageTimer) { $st.ImageTimer.Stop() }
            $t = New-Object System.Windows.Threading.DispatcherTimer
            $t.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds)
            $t.Tag = $player
            $t.Add_Tick({
                $timer = $args[0]; $p = $timer.Tag; $timer.Stop()
                & $SyncHash.MediaEndedHandler -Sender $p -EventArgs $null
            })
            $st.ImageTimer = $t
            $t.Start()
        } else {
            if (-not $player.NaturalDuration.HasTimeSpan) {
                $st.IsFailed = $true
                & $SyncHash.HandleMediaFailure -FailedPlayer $player -Reason "Invalid duration or codec."
            }
        }
    }

    # --- Hook players & seed initial items (each facet gets next item in queue) ---
    for ($i = 0; $i -lt $SyncHash.MediaPlayers.Count; $i++) { } # no-op if already filled; we fill below

    foreach ($facetModel in $facetModels) { } # geometry already added above

    for ($i = 0; $i -lt $SyncHash.MediaPlayers.Count; $i++) {
        $player = $SyncHash.MediaPlayers[$i]
        $player.Add_MediaEnded($MediaEndedHandler)
        $player.Add_MediaOpened($MediaOpenedHandler)
        $player.Add_MediaFailed({ param($s,$e) & $SyncHash.HandleMediaFailure -FailedPlayer $s -Reason $e.ErrorException.Message })

        $SyncHash.PlayerState[$player.GetHashCode()] = @{
            IsImage           = $false
            ImageTimer        = $null
            IsFailed          = $false
            RecoveryTimer     = $null
        }

        & $SyncHash.AssignNext $player   # initial assignment for THIS facet
    }

    # --- Animate the Sphere ---
    $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20))
    $animX.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleX = $window.FindName("AxisAngleX")
    $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

    $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15))
    $animY.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleY = $window.FindName("AxisAngleY")
    $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

    $SyncHash.animX = $animX; $SyncHash.animY = $animY
    $SyncHash.AxisAngleX = $axisAngleX; $SyncHash.AxisAngleY = $axisAngleY

    $window.Add_Loaded({
        foreach ($player in $SyncHash.MediaPlayers) { $player.Play() }  # images will pause in MediaOpened
    })

    # --- UI Controls ---
    $SyncHash.pauseButton        = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton   = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton     = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton      = $window.FindName("speedUpButton")
    $SyncHash.redoButton         = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton        = $window.FindName("closeButton")

    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        if ($SyncHash.Paused) {
            $SyncHash.animX.From = $SyncHash.AxisAngleX.Angle
            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animX)
            $SyncHash.animY.From = $SyncHash.AxisAngleY.Angle
            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animY)
            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.Paused = $false
        } else {
            $currentAngleX = $SyncHash.AxisAngleX.Angle
            $currentAngleY = $SyncHash.AxisAngleY.Angle
            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $SyncHash.AxisAngleX.Angle = $currentAngleX
            $SyncHash.AxisAngleY.Angle = $currentAngleY
            $SyncHash.pauseButton.Content = "Resume"
            $SyncHash.Paused = $true
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
        $SyncHash.animX.Duration = $newDurationX
        $SyncHash.animY.Duration = $newDurationY
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

    # --- Keys & Click ---
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
        foreach ($player in $SyncHash.MediaPlayers) {
            $st = $SyncHash.PlayerState[$player.GetHashCode()]
            if ($st -and $st.ImageTimer)    { $st.ImageTimer.Stop() }
            if ($st -and $st.RecoveryTimer) { $st.RecoveryTimer.Stop() }
            $player.Stop()
            $player.Source = $null
            $player.Close()
        }
        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
    })

    # --- Apply initial overlay style (text/color/font) ---
    for ($i=0; $i -lt $SyncHash.OverlayTextBlocks.Count; $i++) {
        & $SyncHash.ApplyOverlayFor $i $null
    }

    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
