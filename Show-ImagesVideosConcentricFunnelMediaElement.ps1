<#
.SYNOPSIS
    Displays media on a static 3D funnel mesh built from concentric rings of trapezoid panels.
.DESCRIPTION
    This script launches a GUI to select media files and then renders them onto a custom 3D mesh
    that resembles a spiraling funnel or vortex. The mesh is constructed from multiple concentric
    rings of discrete trapezoid panels. Each ring is smaller and more steeply angled than the last,
    creating a stepped, inward-curving funnel.

    The entire structure is static. It uses a small number of shared media players (e.g., 6)
    distributed across the many panels to provide visual variety without the extreme performance
    cost of one player per panel.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system (e.g., MP4, WMV, AVI).
.EXAMPLE
    PS C:\> .\Show-ImagesVideosConcentricFunnelMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D window with the static concentric funnel.
.NOTES
    Name:           Show-ImagesVideosConcentricFunnelMediaElement.ps1
    Version:        1.1.0, 11/19/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Concentric Funnel Generation Function ---
function New-ConcentricFunnelModelGroup {
    param(
        [int]$numberOfRings = 5,
        [int]$panelsPerRing = 12,
        [double]$startRadius = 7.0,
        [double]$endRadius = 1.0,
        [double]$totalHeight = 8.0,
        [array]$SharedMaterials
    )

    $modelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup
    $panelIndex = 0 # Keep track of the absolute panel number

    # Helper to create a single trapezoid panel mesh
    $createPanelMesh = {
        param($p1, $p2, $p3, $p4) # Four corner points
        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
        $mesh.Positions.Add($p1); $mesh.Positions.Add($p2); $mesh.Positions.Add($p3); $mesh.Positions.Add($p4)
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,0)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,0))
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(1,1)); $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0,1))
        $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(1); $mesh.TriangleIndices.Add(2)
        $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add(2); $mesh.TriangleIndices.Add(3)
        $mesh.Freeze()
        return $mesh
    }

    $radiusStep = ($startRadius - $endRadius) / $numberOfRings
    $angleStep = (2 * [Math]::PI) / $panelsPerRing

    for ($r = 0; $r -lt $numberOfRings; $r++) {
        $outerR = $startRadius - ($r * $radiusStep)
        $innerR = $startRadius - (($r + 1) * $radiusStep)

        # Use an easing function for the Y position to create the curve
        $progressOuter = $r / $numberOfRings; $progressOuter_eased = $progressOuter * $progressOuter
        $progressInner = ($r + 1) / $numberOfRings; $progressInner_eased = $progressInner * $progressInner

        $yOuter = $totalHeight / 2 - ($progressOuter_eased * $totalHeight)
        $yInner = $totalHeight / 2 - ($progressInner_eased * $totalHeight)

        for ($p = 0; $p -lt $panelsPerRing; $p++) {
            $theta1 = $p * $angleStep
            $theta2 = ($p + 1) * $angleStep

            # Define the 4 corners of the trapezoid panel
            $p1 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta1), $yOuter, $outerR * [Math]::Sin($theta1))
            $p2 = [System.Windows.Media.Media3D.Point3D]::new($outerR * [Math]::Cos($theta2), $yOuter, $outerR * [Math]::Sin($theta2))
            $p3 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta2), $yInner, $innerR * [Math]::Sin($theta2))
            $p4 = [System.Windows.Media.Media3D.Point3D]::new($innerR * [Math]::Cos($theta1), $yInner, $innerR * [Math]::Sin($theta1))

            $panelMesh = & $createPanelMesh $p1 $p2 $p3 $p4

            # Determine which of the shared materials to use for this panel
            $materialToUse = $SharedMaterials | Get-Random

            $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{
                Geometry = $panelMesh
                Material = $materialToUse
                BackMaterial = $materialToUse.Clone()
            }
            $modelGroup.Children.Add($geometryModel)
            $panelIndex++
        }
    }
    return $modelGroup
}

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Concentric Funnel - Media Selector"
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
        PlayerStates = [hashtable]::Synchronized(@{}) # For the 6 media groups
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

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="MainWindow"
        Title="Concentric Funnel"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,2,18" LookDirection="0,-0.1,-1" UpDirection="0,1,0" FieldOfView="70"/>
            </Viewport3D.Camera>

            <ModelVisual3D x:Name="ObjectContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="#555555"/>
                        <DirectionalLight Color="#FFFFFF" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="#FFFFFF" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <!-- TILT: This static rotation tilts the funnel forward. -->
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D Angle="40" Axis="1,0,0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <!-- MOVE: This static translation moves the funnel down. -->
                        <TranslateTransform3D OffsetY="-2.0" />
                    </Transform3DGroup>
                </ModelVisual3D.Transform>
            </ModelVisual3D>
        </Viewport3D>
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
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
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    $objectContainer = $window.FindName("ObjectContainer")
    $SyncHash.Window = $window

    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")

    # --- Centralized, Thread-Safe Indexing ---
    $globalIndexLock = New-Object object
    function Get-NextMediaIndex {
        [System.Threading.Monitor]::Enter($globalIndexLock)
        try {
            if ($SyncHash.SelectedFiles.Count -eq 0) { return -1 }
            $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1)
            if ($SyncHash.CurrentIndex -ge $SyncHash.SelectedFiles.Count) { $SyncHash.CurrentIndex = 0 }
            return $SyncHash.CurrentIndex
        } finally {
            [System.Threading.Monitor]::Exit($globalIndexLock)
        }
    }

    # --- Media Handling Functions ---
    function Start-NextMediaForGroup {
        param([int]$GroupIndex)

        $playerState = $SyncHash.PlayerStates[$GroupIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        if ($ImageExtensions -contains $extension) {
            $image = New-Object System.Windows.Controls.Image -Property @{
                Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$filePath)
                Stretch = 'Fill'
            }
            $playerState.ContentPresenter.Content = $image

            $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $GroupIndex }
            $timer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaForGroup -GroupIndex $idx })
            $playerState.MediaTimer = $timer
            $timer.Start()
        }
        else { # Video
            $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                LoadedBehavior = 'Play'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = [Uri]$filePath
            }
            $mediaElement.Add_MediaEnded({ 
                $thisElement = $args[0]
                $thisElement.Position = [TimeSpan]::Zero
                $thisElement.Play()
            })
            $playerState.ContentPresenter.Content = $mediaElement
            $playerState.CurrentMediaElement = $mediaElement
        }
    }

    # --- Create the Concentric Funnel ---
    $numberOfGroups = 6
    $sharedMaterials = @()

    # Create 6 unique media hosts and materials
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
        $sharedMaterials += New-Object $materialType -Property @{ Brush = $visualBrush }

        $SyncHash.PlayerStates[$i] = @{
            ContentPresenter = $contentPresenter
            OverlayTextBlock = $overlayTextBlock
        }
    }

    # Now, build the funnel, passing in the array of shared materials
    $funnelModelGroup = New-ConcentricFunnelModelGroup -SharedMaterials $sharedMaterials
    $objectContainer.Content.Children.Add($funnelModelGroup)

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        $controlsPanel.Visibility = if ($controlsPanel.Visibility -eq 'Visible') { 'Collapsed' } else { 'Visible' }
    })

    $SyncHash.redoButton.Add_Click({
        $SyncHash.RedoClicked = $true
        $SyncHash.Window.Close()
    })

    $window.Add_Closed({
        if ($SyncHash.PlayerStates) {
            foreach ($i in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$i]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }
            }
        }
    })

    $window.Add_Loaded({
        # Apply Text Overlay Settings
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            foreach ($i in $SyncHash.PlayerStates.Keys) {
                $textBlock = $SyncHash.PlayerStates[$i].OverlayTextBlock
                $textBlock.Foreground = $brush.Clone()
                $textBlock.FontFamily = $fontFamily
                $textBlock.FontSize = $SyncHash.FontSize
                $textBlock.FontWeight = $fontWeight
                $textBlock.FontStyle = $fontStyle
                if ($SyncHash.RbSelection -eq "Custom") { $textBlock.Text = $SyncHash.CustomText }
            }
        }

        # Start a media cycle for each of the 6 groups
        for ($i = 0; $i -lt $numberOfGroups; $i++) { Start-NextMediaForGroup -GroupIndex $i }
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}