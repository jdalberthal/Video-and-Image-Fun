<#
.SYNOPSIS
    Displays media on 6 spheres that travel along a static 3D roller coaster track.
.DESCRIPTION
    This script first creates a complex, static 3D mesh that resembles a looping, twisting roller
    coaster track. It then creates 6 separate 3D spheres, each with its own media player.

    These 6 spheres are animated on a per-frame basis to travel along the path of the static
    track, like cars on a roller coaster. The user can control the speed and pause the animation.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosCoasterCarsFfmpeg.ps1
.NOTES
    Name:           Show-ImagesVideosCoasterCarsFfmpeg.ps1
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
function New-PathRibbonMesh {
    param([array]$PathData, [int]$segments = 400, [double]$width = 0.5, [double]$verticalOffset = 0)
    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D; $epsilon = 0.001
    $worldUp = [System.Windows.Media.Media3D.Vector3D]::new(0, 1, 0)
    for ($i = 0; $i -le $segments; $i++) {
        $pointData = $PathData[$i]
        $p1 = $pointData.Point
        $up = $pointData.Up
        $normal = $pointData.Normal

        $offsetVector = $up * $verticalOffset
        $mesh.Positions.Add(($p1 + ($normal * $width / 2) + $offsetVector)); $mesh.Positions.Add(($p1 - ($normal * $width / 2) + $offsetVector))
    }
    for ($i = 0; $i -lt $segments; $i++) {
        $i0 = $i * 2; $i1 = $i0 + 1; $i2 = $i0 + 2; $i3 = $i0 + 3
        $mesh.TriangleIndices.Add($i0); $mesh.TriangleIndices.Add($i2); $mesh.TriangleIndices.Add($i1)
        $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i2); $mesh.TriangleIndices.Add($i3)
    }
    $mesh.Freeze(); return $mesh
}

function New-CoasterTrackModelGroup {
    param([array]$PathData, [int]$segments = 800)

    $trackModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup

    # 1. Create the Ties (a wide, flat ribbon)
    $tieWidth = 0.8; $tieThickness = 0.05
    $tiesMesh = New-PathRibbonMesh -PathData $PathData -segments $segments -width $tieWidth -verticalOffset (-$tieThickness / 2)
    $tiesMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::SaddleBrown)
    $tiesModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{
        Geometry = $tiesMesh; Material = $tiesMaterial; BackMaterial = $tiesMaterial
    }
    $trackModelGroup.Children.Add($tiesModel)

    # 2. Create the Rails (two narrower, thicker ribbons on top of the ties)
    $railWidth = 0.1; $railThickness = 0.15
    $railOffsetFromCenter = ($tieWidth / 2) - ($railWidth / 2) + 0.05 # Adjust to sit nicely

    # Create new PathData arrays for the left and right rails, offset from the center
    $leftRailPathData = @()
    $rightRailPathData = @()
    for ($i = 0; $i -lt $PathData.Count; $i++) {
        $centerPointData = $PathData[$i]
        $offsetVector = $centerPointData.Normal * $railOffsetFromCenter
        
        $leftRailPathData += [pscustomobject]@{
            Point = $centerPointData.Point - $offsetVector
            Up = $centerPointData.Up
            Normal = $centerPointData.Normal
        }
        $rightRailPathData += [pscustomobject]@{
            Point = $centerPointData.Point + $offsetVector
            Up = $centerPointData.Up
            Normal = $centerPointData.Normal
        }
    }

    # Use the pre-calculated path data to build the rail meshes
    $leftRailMesh = New-PathRibbonMesh -PathData $leftRailPathData -segments $segments -width $railWidth -verticalOffset ($railThickness / 2)
    $rightRailMesh = New-PathRibbonMesh -PathData $rightRailPathData -segments $segments -width $railWidth -verticalOffset ($railThickness / 2)
    $railMaterial = New-Object System.Windows.Media.Media3D.DiffuseMaterial([System.Windows.Media.Brushes]::DarkSlateGray)
    
    $leftRailModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $leftRailMesh; Material = $railMaterial; BackMaterial = $railMaterial }
    $rightRailModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $rightRailMesh; Material = $railMaterial; BackMaterial = $railMaterial }

    $trackModelGroup.Children.Add($leftRailModel)
    $trackModelGroup.Children.Add($rightRailModel)

    return $trackModelGroup
}

function New-SphereMesh {
    param([double]$radius=1.0, [int]$slices=32, [int]$stacks=16)
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

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form (Standard UI) ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form -Property @{ Text = "Coaster Cars (FFmpeg) - Media Selector"; Size = '800, 680'; StartPosition = "CenterScreen" }
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
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null; $DataGridView.Columns.Add($FileNameColumn) | Out-Null; $DataGridView.Columns.Add($FilePathColumn) | Out-Null
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
        $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
        if ($selectedFiles.Count -gt 0) {
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
        } else { [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", "OK", "Warning") }
    })
    $null = $SelectForm.ShowDialog(); $SelectForm.Dispose()
    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0) { Write-Host "Exiting."; break }

    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
        SelectedFiles = $formState.SelectedFiles
        UseTransparentEffect = $formState.UseTransparentEffect
        PlayerStates = [hashtable]::Synchronized(@{})
        CarStates = [hashtable]::Synchronized(@{})
        PathData = @{} # To store path orientation data
        Paused = $false
        RedoClicked = $false
        SpeedMultiplier = 1.0
        MediaOffset = 0
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
        Title="Roller Coaster Cars" WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,5,25" LookDirection="0,-0.2,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>
            <ModelVisual3D x:Name="SceneContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>
        </Viewport3D>
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="pauseButton" Content="Pause" Padding="10,5" Margin="2"/>
            <Button Name="slowDownButton" Content="&#x2190;" Padding="10,5" Margin="2" FontWeight="Bold"/>
            <Button Name="speedUpButton" Content="&#x2192;" Padding="10,5" Margin="2" FontWeight="Bold"/>
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
    $sceneContainer = $window.FindName("SceneContainer"); $SyncHash.Window = $window
    $SyncHash.pauseButton = $window.FindName("pauseButton"); $SyncHash.slowDownButton = $window.FindName("slowDownButton"); $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton"); $SyncHash.closeButton = $window.FindName("closeButton")

    # --- Media Handling Functions ---
    function Handle-MediaFailure {
        param([int]$CarIndex, [string]$Reason)

        $playerState = $SyncHash.PlayerStates[$CarIndex]
        if ($playerState.IsFailed) { return }

        $maxReasonLength = 250
        if ($Reason.Length -gt $maxReasonLength) {
            $Reason = $Reason.Substring(0, $maxReasonLength) + "..."
        }

        $playerState.IsFailed = $true

        $fileName = if ($playerState.CurrentSource) { [System.IO.Path]::GetFileName($playerState.CurrentSource) } else { "an unknown file" }
        Write-Warning "Media failed for car $CarIndex (File: '$fileName'). Reason: $Reason. Attempting to replace."

        if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
            $SyncHash.SelectedFiles.Remove($playerState.CurrentSource) | Out-Null
        }

        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState.ContentPresenter.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
        })

        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromMilliseconds(50); Tag = $CarIndex }
        $recoveryTimer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaForCar -CarIndex $idx })
        $playerState.RecoveryTimer = $recoveryTimer
        $recoveryTimer.Start()
    }

    function Start-NextMediaForCar {
        param([int]$CarIndex)
        $playerState = $SyncHash.PlayerStates[$CarIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
        
        $mediaIndex = ($CarIndex + $SyncHash.MediaOffset) % $SyncHash.SelectedFiles.Count
        $filePath = $SyncHash.SelectedFiles[$mediaIndex]
        $playerState.CurrentSource = $filePath
        $playerState.IsFailed = $false

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        try {
            $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
            if ($ImageExtensions -contains [System.IO.Path]::GetExtension($filePath).ToLower()) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = [Uri]$filePath; $bitmap.EndInit(); $bitmap.Freeze()
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
                $playerState.ContentPresenter.Content = $image
            } else {
                # Hardcode dimensions to 1280x720 to avoid ffprobe race conditions, matching the working PulsingStar script.
                $width=1280; $height=720
                if (-not [System.IO.File]::Exists($filePath)) { throw "File not found: $filePath" }
                
                $bitmap = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $imageControl = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = "Fill" }
                $playerState.ContentPresenter.Content = $imageControl
                $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent

                $args = "-hide_banner -loglevel error -stream_loop -1 -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=$($width):$($height) -"
                $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{ FileName="ffmpeg.exe"; Arguments=$args; RedirectStandardOutput=$true; UseShellExecute=$false; CreateNoWindow=$true }
                $proc = [System.Diagnostics.Process]::Start($psi); $playerState.FfmpegProcess = $proc
                
                # Check if the process exited immediately, which indicates a fatal error (e.g., file not found, bad format).
                Start-Sleep -Milliseconds 100 # Give the process a moment to start or fail.
                if ($proc.HasExited) { throw "ffmpeg process exited immediately. Invalid file or codec." }

                $stream = $proc.StandardOutput.BaseStream; $bytes = New-Object byte[] ($width*$height*3); $rect = [System.Windows.Int32Rect]::new(0,0,$width,$height); $stride=$width*3
                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval=[TimeSpan]::FromMilliseconds(33); Tag=$CarIndex }
                $tickScriptBlock = {
                    $t = $args[0]; $idx = $t.Tag; try {
                        $totalRead=0; while($totalRead -lt $bytes.Length){ $read=$stream.Read($bytes, $totalRead, $bytes.Length-$totalRead); if($read -le 0){$t.Stop(); Handle-MediaFailure -CarIndex $idx -Reason "ffmpeg stream ended unexpectedly."; return}; $totalRead+=$read }
                        if($totalRead -eq $bytes.Length){ $bitmap.Lock(); $bitmap.WritePixels($rect, $bytes, $stride, 0); $bitmap.Unlock() }
                    } catch { $t.Stop(); Handle-MediaFailure -CarIndex $idx -Reason "Error reading ffmpeg stream." }
                }; $timer.Add_Tick($tickScriptBlock.GetNewClosure()); $playerState.MediaTimer = $timer; $timer.Start()
            }
        } catch {
            Handle-MediaFailure -CarIndex $CarIndex -Reason $_.Exception.Message
        }
    }

    # --- Create Scene Geometry ---
    $window.Add_Loaded({
        # Define the parametric path for the roller coaster
        $script:pathFunc = {
            param([double]$t) # t goes from 0 to 1
            $t_rad = $t * 2 * [Math]::PI

            # Base path: An oval with rolling hills.
            $x_base = 12 * [Math]::Cos($t_rad)
            $z_base = 5 * [Math]::Sin($t_rad)
            $y_base = 2.5 * [Math]::Sin(3 * $t_rad) - 1.5 * [Math]::Cos(5 * $t_rad) + [Math]::Sin($t_rad)

            # --- Add a smooth vertical loop ---
            $loop_radius = 5.0
            $loop_center_t = 0.5 # Center the loop at the far end of the track (x = -12)
            $loop_width = 0.08  # Controls how wide the loop section is

            # Use a Gaussian (bell curve) function for a very smooth influence
            $loop_influence = [Math]::Exp(-[Math]::Pow($t - $loop_center_t, 2) / (2 * [Math]::Pow($loop_width, 2)))

            # Add the loop's circular motion to the Y and Z coordinates, scaled by the influence
            $y = $y_base + $loop_radius * [Math]::Sin(($t - $loop_center_t) / $loop_width * [Math]::PI) * $loop_influence
            $x = $x_base + $loop_radius * ([Math]::Cos(($t - $loop_center_t) / $loop_width * [Math]::PI) + 1) * $loop_influence
            $z = $z_base

            return [System.Windows.Media.Media3D.Point3D]::new($x, $y, $z)
        }
        
        # 1. Pre-calculate all path data to avoid deadlocks
        $segments = 800
        $pathDataArray = for ($i = 0; $i -le $segments; $i++) {
            $t = $i / $segments
            $p1 = & $script:pathFunc $t
            $p2 = & $script:pathFunc ($t + 0.001)
            $tangent = $p2 - $p1; $tangent.Normalize()
            $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '0,1,0')
            if ($normal.LengthSquared -lt 1e-6) { $normal = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($tangent, '1,0,0') }
            $normal.Normalize()
            $up = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($normal, $tangent)
            [pscustomobject]@{ Point = $p1; Up = $up; Normal = $normal }
        }
        # Store the calculated Up vectors for the animation handler
        for($i=0; $i -lt $pathDataArray.Count; $i++){
            $SyncHash.PathData[$i] = @{ Up = $pathDataArray[$i].Up }
        }

        # 2. Build the track model using the pre-calculated data
        $trackModelGroup = New-CoasterTrackModelGroup -PathData $pathDataArray -segments $segments
        $sceneContainer.Content.Children.Add($trackModelGroup)

        # 3. Create the sphere "cars"
        $numberOfCars = 12
        $sphereRadius = 1.34 # Doubled from 0.67
        $SyncHash.SphereRadius = $sphereRadius # Store in SyncHash for global access
        $sphereMesh = New-SphereMesh -radius $sphereRadius
        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $numberOfCars; $i++) {
            $grid = New-Object System.Windows.Controls.Grid
            $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false }
            $grid.Children.Add($contentPresenter) | Out-Null
            $grid.Children.Add($overlayTextBlock) | Out-Null

            if ($SyncHash.RbSelection -ne "Hidden") {
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
                $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
                $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }
                $overlayTextBlock.Foreground = $brush.Clone(); $overlayTextBlock.FontFamily = $fontFamily
                $overlayTextBlock.FontSize = $SyncHash.FontSize; $overlayTextBlock.FontWeight = $fontWeight
                $overlayTextBlock.FontStyle = $fontStyle
                if ($SyncHash.RbSelection -eq "Custom") { $overlayTextBlock.Text = $SyncHash.CustomText }
            }

            $SyncHash.PlayerStates[$i] = @{ ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock; MediaHostGrid = $grid }
            
            $viewportVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
            $viewportVisual.Geometry = $sphereMesh; $viewportVisual.Visual = $grid
            $material = New-Object $materialType; [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($material, $true)
            $viewportVisual.Material = $material
            
            $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
            $viewportVisual.Transform = $translateTransform
            
            $sceneContainer.Children.Add($viewportVisual)
            $SyncHash.CarStates[$i] = @{ TranslateTransform = $translateTransform; Progress = $i / $numberOfCars }
            
            Start-NextMediaForCar -CarIndex $i
        }
        
        # Start the animation loop
        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        [System.Windows.Media.CompositionTarget]::add_Rendering($animationHandler)
    })

    # --- Per-Frame Animation Handler ---
    $animationHandler = {
        param($sender, $e)
        if ($SyncHash.Paused) { return }
        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        $flowSpeed = 0.05 * $SyncHash.SpeedMultiplier

        foreach ($i in $SyncHash.CarStates.Keys) {
            $carState = $SyncHash.CarStates[$i]
            $carState.Progress = ($carState.Progress + ($flowSpeed * $elapsed)) % 1.0 # Update and loop progress
            $positionOnPath = & $script:pathFunc $carState.Progress
            
            # Get the 'Up' vector for the track at this point to offset the sphere
            $pathIndex = [int]($carState.Progress * ($SyncHash.PathData.Count - 1))
            $upVector = $SyncHash.PathData[$pathIndex].Up
            $finalPosition = $positionOnPath + ($upVector * ($SyncHash.SphereRadius + 0.05)) # Add rail thickness to offset
            
            $carState.TranslateTransform.OffsetX = $finalPosition.X
            $carState.TranslateTransform.OffsetY = $finalPosition.Y
            $carState.TranslateTransform.OffsetZ = $finalPosition.Z
        }
    }

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })
    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) { $SyncHash.pauseButton.Content = "Resume" } 
        else { $SyncHash.pauseButton.Content = "Pause"; $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp() }
    })
    $SyncHash.slowDownButton.Add_Click({ $SyncHash.SpeedMultiplier *= 0.5 })
    $SyncHash.speedUpButton.Add_Click({ $SyncHash.SpeedMultiplier *= 2.0 })
    
    $window.Add_KeyDown({ param($s, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })
    $window.Add_Closed({
        [System.Windows.Media.CompositionTarget]::remove_Rendering($animationHandler)
        foreach ($i in $SyncHash.PlayerStates.Keys) {
            $playerState = $SyncHash.PlayerStates[$i]
            if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
        }
    })

    $null = $window.ShowDialog()
    if (-not $SyncHash.RedoClicked) { break }
}
