<#
.SYNOPSIS
    Displays selected images and videos on the faces of six floating, rotating 3D cubes using FFmpeg.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto the faces
    of six independently moving and rotating 3D cubes in a WPF window. This creates a "zero-gravity"
    or "space" visual effect.

    This version uses FFmpeg for video decoding, providing support for a wide range of video formats
    without relying on system-installed codecs.

    The 3D view is interactive, with controls to pause the animation, change the rotation speed,
    and hide the UI for an unobstructed view. It also supports text overlays on the cube faces.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosFloatingCubesFfmpeg.ps1

    Launches the file selection GUI. After selecting at least one file and clicking "Play", the
    script will launch the 3D cube window with six floating cubes.

.NOTES
    Name:           Show-ImagesVideosFloatingCubesFfmpeg.ps1
    Version:        1.0.0, 10/26/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffprobe.exe, ffplay.exe): https://www.ffmpeg.org/download.html
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Floating Cubes `n(FFmpeg)"
$ScriptDescription = "Displays media on six independently floating and rotating 3D cubes. Uses FFmpeg for broad video format support."
$RequiredExecutables = @("ffmpeg.exe", "ffprobe.exe", "ffplay.exe")

# --- Dependency Check ---
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

# --- Main Application Loop ---
while ($true) {
    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
        WindowReady = $false # Flag for the loading form
        SelectedFiles = @()
        GlobalCounter = -1
        PlayerState = [hashtable]::Synchronized(@{})
        ImageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")
        ImageHoldSeconds = 10
        AllAnimations = [hashtable]::Synchronized(@{})
        Paused = $false
        ControlsHidden = $false
        RedoClicked = $false
        UseTransparentEffect = $false
        RbSelection = "Hidden"
        CustomText = ""
        TextColor = [System.Drawing.Color]::Black
        FontSize = 24
        FontFamily = "Arial"
        IsBold = $true
        IsItalic = $false
    })

    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Floating Cubes (FFmpeg) - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
    $SelectForm.Controls.Add($BrowseButton)

    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $SelectForm.Controls.Add($FolderPathTextBox)

    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $SelectForm.Controls.Add($RecursiveCheckBox)

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $SelectForm.Controls.Add($SelectAllCheckbox)

    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $SelectForm.Controls.Add($TransparentCheckbox)

    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'
        AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65
    }
    $SelectForm.Controls.Add($DataGridView)

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null

    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 200; ReadOnly = $true }
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null

    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 330; ReadOnly = $true }
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

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '140, 440'; Size = '455, 180'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"
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

    # Event handler for radio buttons to show/hide text controls
    $textOverlayEvent = {
        $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked
        $isCustomText = $RadioButton3.Checked

        $TextBox.Visible = $isCustomText
        $CurrentColor.Visible = $isTextVisible
        $ColorExample.Visible = $isTextVisible
        $SelectColorButton.Visible = $isTextVisible
        $SizeLabel.Visible = $isTextVisible
        $NumericUpDown.Visible = $isTextVisible
        $FontButton.Visible = $isTextVisible
        $ItalicCheckbox.Visible = $isTextVisible
        $BoldCheckbox.Visible = $isTextVisible
    }
    $RadioButton1.Add_Click($textOverlayEvent)
    $RadioButton2.Add_Click($textOverlayEvent)
    $RadioButton3.Add_Click($textOverlayEvent)

    # --- Event Handlers for Text Customization ---
    $ColorExample.BackColor = $SyncHash.TextColor
    $SelectColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SyncHash.TextColor = $colorDialog.Color
            $ColorExample.BackColor = $SyncHash.TextColor
            $TextBox.ForeColor = $SyncHash.TextColor
        }
    })

    $FontButton.Add_Click({
        $fontDialog = New-Object System.Windows.Forms.FontDialog
        try {
            $currentStyle = [System.Drawing.FontStyle]::Regular
            if ($BoldCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Bold }
            if ($ItalicCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Italic }
            $fontDialog.Font = New-Object System.Drawing.Font($SyncHash.FontFamily, [float]$NumericUpDown.Value, $currentStyle)
        } catch { $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12) }

        if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $SyncHash.FontFamily = $fontDialog.Font.Name
            $FontButton.Text = $SyncHash.FontFamily
            $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
            $BoldCheckbox.Checked = $fontDialog.Font.Bold
            $ItalicCheckbox.Checked = $fontDialog.Font.Italic
        }
    })

    $updateFontStyle = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
        try {
            $newFont = New-Object System.Drawing.Font($SyncHash.FontFamily, [float]$NumericUpDown.Value, $style)
            $TextBox.Font = $newFont
        } catch { $TextBox.Font = New-Object System.Drawing.Font("Arial", 12, $style) }
    }
    $ItalicCheckbox.Add_CheckedChanged($updateFontStyle)
    $BoldCheckbox.Add_CheckedChanged($updateFontStyle)
    $NumericUpDown.Add_ValueChanged($updateFontStyle)
    & $updateFontStyle

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
            
            $gciParams = @{ File = $true; Include = $AllowedExtensions }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true }
            else { $gciParams.Path = Join-Path $SelectedPath "*" }
            $files = Get-ChildItem @gciParams
            foreach ($file in $files) { $DataGridView.Rows.Add($false, $file.Name, $file.FullName) }
            foreach ($row in $DataGridView.Rows) { if (-not $row.IsNewRow) { $row.HeaderCell.Value = "Play" } }
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0) {
            $row = $DataGridView.Rows[$e.RowIndex]
            $filePath = $row.Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -i `"$filePath`"" -NoNewWindow
            } else { [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", "OK", "Error") }
        }
    })

    $PlayButton.Add_Click({
        $SyncHash.SelectedFiles = @(
            foreach ($Row in $DataGridView.Rows) {
                if ($Row.Cells["Select"].Value) { $Row.Cells["FilePath"].Value }
            }
        )
        if ($SyncHash.SelectedFiles.Count -gt 0) {
            $SyncHash.UseTransparentEffect = $TransparentCheckbox.Checked
            if ($RadioButton1.Checked) { $SyncHash.RbSelection = "Hidden" }
            if ($RadioButton2.Checked) { $SyncHash.RbSelection = "Filename" }
            if ($RadioButton3.Checked) { $SyncHash.RbSelection = "Custom" }
            $SyncHash.CustomText = $TextBox.Text
            $SyncHash.FontSize = $NumericUpDown.Value
            $SyncHash.IsBold = $BoldCheckbox.Checked
            $SyncHash.IsItalic = $ItalicCheckbox.Checked
            $SelectForm.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })

    # --- Loading Form in a Separate Runspace ---
    $loadingRunspace = $null
    $loadingPS = $null
    $loadingJob = $null

    $SelectForm.Add_FormClosing({
        param($sender, $e)

        # Only start the loading form if files were actually selected.
        if ($SyncHash.SelectedFiles.Count -gt 0) {
            $loadingScriptBlock = {
                param($SyncHash)
                Add-Type -AssemblyName System.Windows.Forms, System.Drawing
                [System.Windows.Forms.Application]::EnableVisualStyles()

                $loadingForm = New-Object System.Windows.Forms.Form -Property @{ Text = "Loading..."; Size = '300,120'; StartPosition = "CenterScreen"; FormBorderStyle = "FixedDialog"; ControlBox = $false }
                $loadingLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Loading media, please wait..."; Location = '20,20'; AutoSize = $true }
                $progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{ Style = "Marquee"; Location = '20,50'; Size = '250,20' }
                $loadingForm.Controls.AddRange(@($loadingLabel, $progressBar))
                $loadingForm.Show()

                while (-not $SyncHash.WindowReady) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 50
                }
                $loadingForm.Close(); $loadingForm.Dispose()
            }

            $loadingRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $loadingRunspace.ApartmentState = "STA"; $loadingRunspace.Open()
            $loadingPS = [PowerShell]::Create().AddScript($loadingScriptBlock).AddArgument($SyncHash)
            $loadingPS.Runspace = $loadingRunspace
            $loadingJob = $loadingPS.BeginInvoke()
        }
    })

    $null = $SelectForm.ShowDialog()
    $SelectForm.Dispose()

    if ($SyncHash.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break
    }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Floating Cubes (FFmpeg)"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,15" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="#404040"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>
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

    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    $mainViewport = $window.FindName("mainViewport")
    $SyncHash.Window = $window

    # --- Helper Functions ---

    function New-CubeVisual3D {
        param(
            [int]$CubeIndex,
            [hashtable]$SyncHash
        )

        $cubeVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $cubeModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup

        $faceNames = @("Front", "Back", "Right", "Left", "Top", "Bottom")
        $positions = @(
            "-1,-1,1  1,-1,1  1,1,1  -1,1,1",   # Front
            "-1,-1,-1  -1,1,-1  1,1,-1  1,-1,-1",   # Back
            "1,-1,1  1,-1,-1  1,1,-1  1,1,1",     # Right
            "-1,-1,-1  -1,-1,1  -1,1,1  -1,1,-1",   # Left
            "-1,1,1  1,1,1  1,1,-1  -1,1,-1",     # Top
            "-1,-1,-1  1,-1,-1  1,-1,1  -1,-1,1"    # Bottom
        )
        $normals = @(
            "0,0,1 0,0,1 0,0,1 0,0,1", "0,0,-1 0,0,-1 0,0,-1 0,0,-1", "1,0,0 1,0,0 1,0,0 1,0,0",
            "-1,0,0 -1,0,0 -1,0,0 -1,0,0", "0,1,0 0,1,0 0,1,0 0,1,0", "0,-1,0 0,-1,0 0,-1,0 0,-1,0"
        )
        $textureCoords = @(
            "0,1 1,1 1,0 0,0", "1,1 1,0 0,0 0,1", "0,1 1,1 1,0 0,0",
            "0,1 1,1 1,0 0,0", "0,1 1,1 1,0 0,0", "0,1 1,1 1,0 0,0"
        )

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

        for ($i = 0; $i -lt $faceNames.Length; $i++) {
            $faceName = $faceNames[$i]
            $facetIndex = ($CubeIndex * 6) + $i

            $imageControl = New-Object System.Windows.Controls.Image -Property @{ Stretch = 'Fill' }
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
                HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap';
                TextAlignment = 'Center'; IsHitTestVisible = $false; Margin = '10,0,10,0'
            }
            $grid = New-Object System.Windows.Controls.Grid -Property @{ Background = [System.Windows.Media.Brushes]::Black }
            [void]$grid.Children.Add($imageControl)
            [void]$grid.Children.Add($overlayTextBlock)

            $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $grid }
            $material = New-Object $materialType -Property @{ Brush = $visualBrush }
            if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

            $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D -Property @{
                Positions = $positions[$i]; TriangleIndices = "0,1,2 0,2,3"; Normals = $normals[$i]; TextureCoordinates = $textureCoords[$i]
            }
            $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $mesh; Material = $material }
            [void]$cubeModelGroup.Children.Add($geometryModel)

            $SyncHash.PlayerState[$facetIndex] = @{
                ImageControl  = $imageControl
                Overlay       = $overlayTextBlock
                IsImage       = $false
                ImageTimer    = $null
                IsFailed      = $false
                RecoveryTimer = $null
                FfmpegProcess = $null
                FrameReader   = $null
                WriteableBmp  = $null
                CurrentPath   = $null
            }
        }

        $cubeVisual.Content = $cubeModelGroup
        return $cubeVisual
    }

    $globalIndexLock = New-Object object
    function Get-NextMediaIndex {
        param([hashtable]$SyncHash)
        [System.Threading.Monitor]::Enter($globalIndexLock)
        try {
            $SyncHash.GlobalCounter = ($SyncHash.GlobalCounter + 1) % $SyncHash.SelectedFiles.Count
            return $SyncHash.GlobalCounter
        } finally {
            [System.Threading.Monitor]::Exit($globalIndexLock)
        }
    }

    function Assign-NextMediaToFace {
        param([int]$FacetIndex, [hashtable]$SyncHash)

        $playerState = $SyncHash.PlayerState[$FacetIndex]
        if (-not $playerState) { return }

        # Stop any existing timers or processes for this face
        if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }

        $nextIndex = Get-NextMediaIndex -SyncHash $SyncHash
        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $uri = [Uri]$filePath

        # Apply overlay text
        $overlay = $playerState.Overlay
        switch ($SyncHash.RbSelection) {
            "Hidden"   { $overlay.Visibility = 'Collapsed' }
            "Filename" { $overlay.Text = [System.IO.Path]::GetFileName($uri.LocalPath); $overlay.Visibility = 'Visible' }
            "Custom"   { $overlay.Text = $SyncHash.CustomText; $overlay.Visibility = 'Visible' }
        }

        $playerState.CurrentPath = $uri.LocalPath
        $ext = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
        $playerState.IsImage = ($SyncHash.ImageExtensions -contains $ext)

        if ($playerState.IsImage) {
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit(); $bitmap.UriSource = $uri; $bitmap.EndInit()
            $playerState.ImageControl.Source = $bitmap

            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds)
            $timer.Tag = $FacetIndex
            $timer.Add_Tick({
                $t = $args[0]; $fIndex = $t.Tag; $t.Stop()
                Assign-NextMediaToFace -FacetIndex $fIndex -SyncHash $SyncHash
            })
            $playerState.ImageTimer = $timer
            $timer.Start()
        } else {
            # Video playback with FFmpeg
            try {
                $psi_probe = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                    FileName = "ffprobe.exe"; Arguments = "-v error -select_streams v:0 -show_entries stream=width,height -of csv=s=x:p=0 `"$($uri.LocalPath)`"";
                    RedirectStandardOutput = $true; UseShellExecute = $false; CreateNoWindow = $true
                }
                $probe_proc = [System.Diagnostics.Process]::Start($psi_probe)
                $ffprobeOutput = $probe_proc.StandardOutput.ReadToEnd()
                $probe_proc.WaitForExit()

                $width, $height = $ffprobeOutput -split 'x'
                if (($probe_proc.ExitCode -ne 0) -or -not ($width -and $height -and [int]$width -gt 0 -and [int]$height -gt 0)) { throw "Invalid dimensions or corrupt file." }

                $playerState.WriteableBmp = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $playerState.ImageControl.Source = $playerState.WriteableBmp

                $loopArg = if ($SyncHash.SelectedFiles.Count -le 36) { "-stream_loop -1" } else { "" }

                $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                    FileName = "ffmpeg"; Arguments = "-hide_banner -loglevel error $loopArg -i `"$($uri.LocalPath)`" -vf scale=${width}:${height} -f rawvideo -pix_fmt bgr24 -";
                    RedirectStandardOutput = $true; RedirectStandardError = $true; UseShellExecute = $false; CreateNoWindow = $true
                }
                $ffmpeg = [System.Diagnostics.Process]::Start($psi)

                $playerState.FfmpegProcess = $ffmpeg
                $stream = $ffmpeg.StandardOutput.BaseStream
                $bytesPerFrame = [int]$width * [int]$height * 3
                $buffer = New-Object byte[] $bytesPerFrame
                $rect = [System.Windows.Int32Rect]::new(0, 0, [int]$width, [int]$height)
                $stride = [int]$width * 3

                $frameTimer = New-Object System.Windows.Threading.DispatcherTimer
                $frameTimer.Interval = [TimeSpan]::FromMilliseconds(33) # ~30fps
                $frameTimer.Tag = $FacetIndex

                $tickScriptBlock = {
                    $timer = $args[0]; $fIndex = $timer.Tag
                    $pState = $SyncHash.PlayerState[$fIndex]
                    $totalRead = 0
                    while ($totalRead -lt $bytesPerFrame) {
                        $bytesRead = $stream.Read($buffer, $totalRead, $bytesPerFrame - $totalRead)
                        if ($bytesRead -le 0) { # End of stream or error
                            $timer.Stop()
                            Assign-NextMediaToFace -FacetIndex $fIndex -SyncHash $SyncHash
                            return
                        }
                        $totalRead += $bytesRead
                    }
                    if ($totalRead -eq $bytesPerFrame) {
                        $pState.WriteableBmp.Lock()
                        $pState.WriteableBmp.WritePixels($rect, $buffer, $stride, 0)
                        $pState.WriteableBmp.Unlock()
                    }
                }
                $frameTimer.Add_Tick($tickScriptBlock.GetNewClosure())
                $playerState.ImageTimer = $frameTimer # Reuse ImageTimer property for cleanup
                $frameTimer.Start()
            } catch {
                Handle-MediaFailure -FacetIndex $FacetIndex -Reason $_.Exception.Message -SyncHash $SyncHash
            }
        }
    }

    function Handle-MediaFailure {
        param([int]$FacetIndex, [string]$Reason, [hashtable]$SyncHash)

        $playerState = $SyncHash.PlayerState[$FacetIndex]
        if (-not $playerState -or $playerState.IsFailed) { return }

        $playerState.IsFailed = $true
        $fileName = if ($playerState.CurrentPath) { [System.IO.Path]::GetFileName($playerState.CurrentPath) } else { "media" }
        $playerState.Overlay.Text = "ERROR`n$fileName`n$Reason"
        $playerState.Overlay.Visibility = 'Visible'
        $playerState.ImageControl.Source = $null # Clear the image

        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
        $recoveryTimer.Interval = [TimeSpan]::FromSeconds(5)
        $recoveryTimer.Tag = $FacetIndex
        $recoveryTimer.Add_Tick({
            $timer = $args[0]; $fIndex = $timer.Tag; $timer.Stop()
            $state = $SyncHash.PlayerState[$fIndex]
            $state.IsFailed = $false
            Assign-NextMediaToFace -FacetIndex $fIndex -SyncHash $SyncHash
        })
        $playerState.RecoveryTimer = $recoveryTimer
        $recoveryTimer.Start()
    }

    # --- Create and Animate 6 Cubes ---
    for ($i = 0; $i -lt 6; $i++) {
        $cubeVisual = New-CubeVisual3D -CubeIndex $i -SyncHash $SyncHash

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
        $rotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D
        $rotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
        $rotateTransform.Rotation = $rotation
        $transformGroup.Children.Add($rotateTransform)

        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        $transformGroup.Children.Add($translateTransform)
        $cubeVisual.Transform = $transformGroup

        $rotAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
            From = 0; To = 360; Duration = [TimeSpan]::FromSeconds((Get-Random -Minimum 15 -Maximum 45))
            RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
        }
        $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $rotAnim)

        $posAnimX = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames
        $posAnimY = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames
        $posAnimZ = New-Object System.Windows.Media.Animation.DoubleAnimationUsingKeyFrames

        $durationSeconds = (Get-Random -Minimum 20 -Maximum 60)
        $posAnimX.Duration = [TimeSpan]::FromSeconds($durationSeconds)
        $posAnimY.Duration = [TimeSpan]::FromSeconds($durationSeconds)
        $posAnimZ.Duration = [TimeSpan]::FromSeconds($durationSeconds)
        $posAnimX.RepeatBehavior = 'Forever'; $posAnimY.RepeatBehavior = 'Forever'; $posAnimZ.RepeatBehavior = 'Forever'

        $xRadius = (Get-Random -Minimum 3 -Maximum 8)
        $yRadius = (Get-Random -Minimum 2 -Maximum 6)
        $zRadius = (Get-Random -Minimum 1 -Maximum 4)
        $timeOffset = ($durationSeconds / 4) * (Get-Random -Minimum 0 -Maximum 3)

        for ($k = 0; $k -le 4; $k++) {
            $time = [System.Windows.Media.Animation.KeyTime]::FromTimeSpan([TimeSpan]::FromSeconds(($k * $durationSeconds / 4 + $timeOffset) % $durationSeconds))
            $angle = $k * [Math]::PI / 2
            $x = $xRadius * [Math]::Cos($angle)
            $y = $yRadius * [Math]::Sin($angle)
            $z = $zRadius * [Math]::Cos($angle * 2)
            [void]$posAnimX.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($x, $time)))
            [void]$posAnimY.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($y, $time)))
            [void]$posAnimZ.KeyFrames.Add((New-Object System.Windows.Media.Animation.SplineDoubleKeyFrame($z, $time)))
        }

        $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $posAnimX)
        $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $posAnimY)
        $translateTransform.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $posAnimZ)
        
        $SyncHash.AllAnimations[$cubeVisual.GetHashCode()] = [pscustomobject]@{
            RotationAnimation = $rotAnim; PositionAnimationX = $posAnimX; PositionAnimationY = $posAnimY; PositionAnimationZ = $posAnimZ
        }
        [void]$mainViewport.Children.Add($cubeVisual)
    }

    # --- Initial Media Loading ---
    for ($i = 0; $i -lt ($SyncHash.PlayerState.Keys.Count); $i++) {
        Assign-NextMediaToFace -FacetIndex $i -SyncHash $SyncHash
    }

    # --- Apply Text Overlay Settings on Load ---
    foreach ($playerState in $SyncHash.PlayerState.Values) {
        $overlay = $playerState.Overlay
        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $overlay.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $overlay.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $overlay.FontSize = $SyncHash.FontSize
            if ($SyncHash.IsBold) { $overlay.FontWeight = 'Bold' }
            if ($SyncHash.IsItalic) { $overlay.FontStyle = 'Italic' }
        }
    }

    # --- UI Controls and Event Handlers ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")

    $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        if ($SyncHash.ControlsHidden) { $controlsPanel.Visibility = 'Visible'; $SyncHash.ControlsHidden = $false }
        else { $controlsPanel.Visibility = 'Collapsed'; $SyncHash.ControlsHidden = $true }
    })

    $pauseOrResumeAnimations = {
        if ($SyncHash.Paused) {
            # --- RESUME ---
            foreach ($visual in $mainViewport.Children) {
                if (-not ($visual -is [System.Windows.Media.Media3D.ModelVisual3D] -and $visual.Transform)) { continue }
                $animSet = $SyncHash.AllAnimations[$visual.GetHashCode()]; if (-not $animSet) { continue }
                $rotation = $visual.Transform.Children[0].Rotation
                $translation = $visual.Transform.Children[1]
                $animSet.RotationAnimation.From = $rotation.Angle
                $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animSet.RotationAnimation)
                $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $animSet.PositionAnimationX)
                $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $animSet.PositionAnimationY)
                $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $animSet.PositionAnimationZ)
            }
            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.Paused = $false
        } else {
            # --- PAUSE ---
            foreach ($visual in $mainViewport.Children) {
                if (-not ($visual -is [System.Windows.Media.Media3D.ModelVisual3D] -and $visual.Transform -and $visual.Transform.Children)) { continue }
                $rotation = $visual.Transform.Children[0].Rotation
                $translation = $visual.Transform.Children[1]
                $currentAngle = $rotation.Angle; $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null); $rotation.Angle = $currentAngle
                $currentX = $translation.OffsetX; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetXProperty, $null); $translation.OffsetX = $currentX
                $currentY = $translation.OffsetY; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetYProperty, $null); $translation.OffsetY = $currentY
                $currentZ = $translation.OffsetZ; $translation.BeginAnimation([System.Windows.Media.Media3D.TranslateTransform3D]::OffsetZProperty, $null); $translation.OffsetZ = $currentZ
            }
            $SyncHash.pauseButton.Content = "Resume"
            $SyncHash.Paused = $true
        }
    }
    $SyncHash.pauseButton.Add_Click($pauseOrResumeAnimations)

    $changeSpeed = {
        param($multiplier)
        foreach ($animSet in $SyncHash.AllAnimations.Values) {
            $animSet.RotationAnimation.Duration = [TimeSpan]::FromSeconds($animSet.RotationAnimation.Duration.TimeSpan.TotalSeconds * $multiplier)
            $animSet.PositionAnimationX.Duration = [TimeSpan]::FromSeconds($animSet.PositionAnimationX.Duration.TimeSpan.TotalSeconds * $multiplier)
            $animSet.PositionAnimationY.Duration = [TimeSpan]::FromSeconds($animSet.PositionAnimationY.Duration.TimeSpan.TotalSeconds * $multiplier)
            $animSet.PositionAnimationZ.Duration = [TimeSpan]::FromSeconds($animSet.PositionAnimationZ.Duration.TimeSpan.TotalSeconds * $multiplier)
        }
        if (-not $SyncHash.Paused) {
            & $pauseOrResumeAnimations # Pause
            $window.Dispatcher.InvokeAsync([action]{ & $pauseOrResumeAnimations }, "Background") | Out-Null
        }
    }
    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 }) # Slower = longer duration
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })  # Faster = shorter duration

    # --- Window-level Events ---
    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P"      { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R"      { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H"      { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left"   { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right"  { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $mainGrid = $window.FindName("MainGrid")
    $mainGrid.Add_MouseDown({ & $pauseOrResumeAnimations })

    $window.Add_Closed({
        foreach ($animSet in $SyncHash.AllAnimations.Values) {
            if ($animSet) {
                $animSet.RotationAnimation.BeginTime = $null; $animSet.PositionAnimationX.BeginTime = $null
                $animSet.PositionAnimationY.BeginTime = $null; $animSet.PositionAnimationZ.BeginTime = $null
            }
        }
        foreach ($playerState in $SyncHash.PlayerState.Values) {
            if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) { try { $playerState.FfmpegProcess.Kill() } catch {} }
        }
    })

    $window.Add_ContentRendered({
        # This event fires after the window is rendered. It's the perfect time
        # to signal that the main window is ready and the loading form can close.
        $SyncHash.WindowReady = $true
    })

    # --- Show the Window ---
    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }

    # Clean up the loading runspace if it was created
    if ($loadingJob) { $loadingPS.EndInvoke($loadingJob) }
    if ($loadingPS) { $loadingPS.Dispose() }
    if ($loadingRunspace) { $loadingRunspace.Dispose() }
    $loadingRunspace = $null
    $loadingPS = $null
}
