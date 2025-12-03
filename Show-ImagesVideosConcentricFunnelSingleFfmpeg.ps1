<#
.SYNOPSIS
    Displays media on a static 3D funnel mesh built from concentric rings of trapezoid panels.
.DESCRIPTION
    This script launches a GUI to select media files and then renders them onto a custom 3D mesh
    that resembles a spiraling funnel or vortex. The mesh is constructed from multiple concentric
    rings of discrete trapezoid panels. Each ring is smaller and more steeply angled than the last,
    creating a stepped, inward-curving funnel.

    The entire structure is static. Each individual panel has its own media player, creating a
    complex and interesting mosaic of the selected images and videos.

    It uses FFmpeg to decode video frames in real-time, allowing for broad video format support.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosConcentricFunnelFfmpeg.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D window with the static concentric funnel.
.NOTES
    Name:           Show-ImagesVideosConcentricFunnelFfmpeg.ps1
    Version:        1.0.0, 11/19/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffplay.exe, ffprobe.exe): https://www.ffmpeg.org/download.html
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

# --- Concentric Funnel Generation Function ---
function New-ConcentricFunnelModelGroup {
    param(
        [int]$numberOfRings = 5,
        [int]$panelsPerRing = 12,
        [double]$startRadius = 7.0,
        [double]$endRadius = 1.0,
        [double]$totalHeight = 8.0,
        [System.Windows.Media.Media3D.Material]$SharedMaterial
    )

    $modelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup

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

            $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{
                Geometry = $panelMesh
                Material = $SharedMaterial
                BackMaterial = $SharedMaterial.Clone()
            }
            $modelGroup.Children.Add($geometryModel)
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
            # Initialize the dialog with the current settings from the form
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
                # When the dialog closes, update all related controls on the main form
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
                $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
                $BoldCheckbox.Checked = $fontDialog.Font.Bold
                $ItalicCheckbox.Checked = $fontDialog.Font.Italic
                $formState.TextColor = $fontDialog.Color
                $ColorExample.BackColor = $formState.TextColor

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
            $TextBox.ForeColor = $formState.TextColor
        } catch {
            # Fallback to a default font if the selected one is invalid
            $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style)
        }
    }
    $NumericUpDown.Add_ValueChanged($updateTextBoxFont)
    $ItalicCheckbox.Add_CheckedChanged($updateTextBoxFont)
    $BoldCheckbox.Add_CheckedChanged($updateTextBoxFont)
    & $updateTextBoxFont # Call once to set initial state

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
            $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.mov", "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2", "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v", "*.f4p", "*.f4a", "*.f4b"
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
                Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -i `"$filePath`"" -NoNewWindow
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
        MediaTimer = $null
        FfmpegProcess = $null
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

    $SyncHash.ContentPresenter = $null # Will be set after creation
    $SyncHash.OverlayTextBlock = $null # Will be set after creation

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
    function Start-NextMedia {
        # Clean up previous media resources
        if ($SyncHash.MediaTimer) { $SyncHash.MediaTimer.Stop(); $SyncHash.MediaTimer = $null }
        if ($SyncHash.FfmpegProcess -and -not $SyncHash.FfmpegProcess.HasExited) {
            try { $SyncHash.FfmpegProcess.Kill() } catch {}
            $SyncHash.FfmpegProcess.Dispose()
            $SyncHash.FfmpegProcess = $null
        }
        
        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        if ($SyncHash.RbSelection -eq "Filename") {
            $SyncHash.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        if ($ImageExtensions -contains $extension) {
            $image = New-Object System.Windows.Controls.Image -Property @{
                Source  = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$filePath)
                Stretch = 'Fill'
            }
            $SyncHash.ContentPresenter.Content = $image

            $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10) }
            $timer.Add_Tick({ $args[0].Stop(); Start-NextMedia })
            $SyncHash.MediaTimer = $timer
            $timer.Start()
        }
        else { # Video
            $width = 1920; $height = 1080
            $frameSize = $width * $height * 3
            $bitmap = [System.Windows.Media.Imaging.WriteableBitmap]::new($width, $height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
            $imageControl = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = "Fill" }
            $SyncHash.ContentPresenter.Content = $imageControl

            $args = "-hide_banner -loglevel error -stream_loop -1 -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"
            $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                FileName = "ffmpeg.exe"; Arguments = $args; RedirectStandardOutput = $true
                UseShellExecute = $false; CreateNoWindow = $true
            }
            $proc = [System.Diagnostics.Process]::Start($psi)
            $SyncHash.FfmpegProcess = $proc

            $stream = $proc.StandardOutput.BaseStream
            $bytes = New-Object byte[] $frameSize
            $rect = [System.Windows.Int32Rect]::new(0, 0, $width, $height)
            $stride = $width * 3

            $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(33) }
            $tickScriptBlock = {
                try {
                    $totalRead = 0
                    while ($totalRead -lt $frameSize) {
                        $read = $stream.Read($bytes, $totalRead, $frameSize - $totalRead)
                        if ($read -le 0) { $args[0].Stop(); Start-NextMedia; return }
                        $totalRead += $read
                    }
                    if ($totalRead -eq $frameSize) {
                        $bitmap.Lock(); $bitmap.WritePixels($rect, $bytes, $stride, 0); $bitmap.Unlock()
                    }
                } catch { $t.Stop() }
            }
            $timer.Add_Tick($tickScriptBlock.GetNewClosure())
            $SyncHash.MediaTimer = $timer
            $timer.Start()
        }
    }

    # --- Create the Concentric Funnel ---
    # Create ONE media host that will be shared by all panels
    $mediaHostGrid = New-Object System.Windows.Controls.Grid
    $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
    $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
    $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
        HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
    }
    $mediaHostGrid.Children.Add($overlayTextBlock) | Out-Null
    $SyncHash.ContentPresenter = $contentPresenter
    $SyncHash.OverlayTextBlock = $overlayTextBlock

    # Create ONE material from that ONE media host
    $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
    $sharedMaterial = New-Object $materialType -Property @{ Brush = $visualBrush }

    # Now, build the funnel, passing in the single shared material
    $funnelModelGroup = New-ConcentricFunnelModelGroup -SharedMaterial $sharedMaterial
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
        # Clean up the single media player resources
        if ($SyncHash.MediaTimer) { $SyncHash.MediaTimer.Stop() }
        if ($SyncHash.FfmpegProcess -and -not $SyncHash.FfmpegProcess.HasExited) {
            try { $SyncHash.FfmpegProcess.Kill(); $SyncHash.FfmpegProcess.Dispose() } catch {}
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

            $textBlock = $SyncHash.OverlayTextBlock
            $textBlock.Foreground = $brush
            $textBlock.FontFamily = $fontFamily
            $textBlock.FontSize = $SyncHash.FontSize
            $textBlock.FontWeight = $fontWeight
            $textBlock.FontStyle = $fontStyle
            if ($SyncHash.RbSelection -eq "Custom") { $textBlock.Text = $SyncHash.CustomText }
        }

        # Start the single media player
        Start-NextMedia
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}