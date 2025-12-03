<#
.SYNOPSIS
    Displays media on spheres that spiral down a rotating 3D vortex or funnel.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto multiple
    3D spheres. These spheres are arranged in a descending spiral, creating the visual effect of
    a vortex or a black hole pulling in planets.

    The entire structure rotates, and each panel individually travels down the spiral path,
    disappearing at the bottom and reappearing at the top for a continuous flow effect.
    This is achieved using per-frame animation calculations for precise control over the complex motion.

    This version uses FFmpeg for video decoding, providing support for a wide range of video formats.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosSphereVortexFfmpeg.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D vortex window.

.NOTES
    Name:           Show-ImagesVideosSphereVortexFfmpeg.ps1
    Version:        1.0.0, 11/18/2025
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
$ExternalButtonName = "3D Sphere Vortex`n(FFmpeg)"
$ScriptDescription = "Displays media on spheres that spiral down a rotating 3D vortex. Uses FFmpeg for broad video format support."
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
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "3D Vortex (FFmpeg) - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 680)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25'
    }
    $SelectForm.Controls.Add($BrowseButton)

    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '120, 10'; Size = '450, 25'; ReadOnly = $true
    }
    $SelectForm.Controls.Add($FolderPathTextBox)

    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false
    }
    $SelectForm.Controls.Add($RecursiveCheckBox)

    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false
    }
    $SelectForm.Controls.Add($TransparentCheckbox)

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false
    }
    $SelectForm.Controls.Add($SelectAllCheckbox)

    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{
        Location = '10, 95'; Size = '760, 330'; Anchor = 'Top, Bottom, Left, Right'
        AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65
    }
    $SelectForm.Controls.Add($DataGridView)

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{
        Name = "Select"; HeaderText = ""; Width = 30
    }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null

    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null

    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Play Selected"; Location = '600, 40'; Size = '170, 30'
    }
    $SelectForm.Controls.Add($PlayButton)

    # --- Text Overlay Controls ---
    $GroupBox = New-Object System.Windows.Forms.GroupBox -Property @{ Text = "Text Overlay"; Location = '10, 440'; Size = '125, 130' }
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Hide Text Overlay"; Location = '10, 30'; Width = 114; Checked = $true }
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Filename"; Location = '10, 60' }
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton -Property @{ Text = "Custom Text"; Location = '10, 90' }
    $GroupBox.Controls.AddRange(@($RadioButton1, $RadioButton2, $RadioButton3))
    $SelectForm.Controls.Add($GroupBox)

    $TextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = '140, 440'; Size = '455, 180'; Multiline = $true; Visible = $false; ScrollBars = "Vertical"; Font = "Arial, 12"; TextAlign = 'Center'
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
    
    # --- Vortex Orientation Controls ---
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
            & $updateTextBoxFont # Call the update function to apply the color
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            try {
                $currentStyle = [System.Drawing.FontStyle]::Regular
                if ($BoldCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Bold }
                if ($ItalicCheckbox.Checked) { $currentStyle = $currentStyle -bor [System.Drawing.FontStyle]::Italic }
                $fontDialog.Font = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $currentStyle)
            } catch {
                $fontDialog.Font = New-Object System.Drawing.Font("Arial", 12)
            }

            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
                $NumericUpDown.Value = [decimal]$fontDialog.Font.Size
                $BoldCheckbox.Checked = $fontDialog.Font.Bold
                $ItalicCheckbox.Checked = $fontDialog.Font.Italic
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
            $TextBox.ForeColor = $formState.TextColor # Apply the selected color
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
                $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.mov", "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2", "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v", "*.f4p", "*.f4a", "*.f4b"
                $AllowedExtensions = $ImageExtensions + $VideoExtensions
            
                $gciParams = @{ File = $true; Include = $AllowedExtensions }
                if ($RecursiveCheckBox.Checked) {
                    $gciParams.Path = $SelectedPath
                    $gciParams.Recurse = $true
                } else {
                    $gciParams.Path = Join-Path $SelectedPath "*"
                }
                $files = Get-ChildItem @gciParams
                foreach ($file in $files) {
                    $DataGridView.Rows.Add($false, $file.Name, $file.FullName)
                }

                foreach ($row in $DataGridView.Rows) {
                    if ($row.IsNewRow) { continue }
                    $row.HeaderCell.Value = "Play"
                }
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
            $formState.SelectedFiles = @(
                foreach ($Row in $DataGridView.Rows) {
                    if ($Row.Cells["Select"].Value) { $Row.Cells["FilePath"].Value }
                }
            )
            if ($formState.SelectedFiles.Count -gt 0) {
                $formState.UseTransparentEffect = $TransparentCheckbox.Checked
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
            SelectedFiles        = $formState.SelectedFiles
            UseTransparentEffect = $formState.UseTransparentEffect
            CurrentIndex         = -1
            PlayerStates         = [hashtable]::Synchronized(@{})
            Paused               = $false
            ControlsHidden       = $false
            RedoClicked          = $false
            LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
            VortexOrientation    = $formState.VortexOrientation
            # Text Overlay Settings
            RbSelection          = $formState.RbSelection
            CustomText           = $formState.CustomText
            TextColor            = $formState.TextColor
            FontSize             = $formState.FontSize
            FontFamily           = $formState.FontFamily
            IsBold               = $formState.IsBold
            IsItalic             = $formState.IsItalic
        })
    $SyncHash.SpeedMultiplier = 1.0 # Initialize speed multiplier

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

            <!-- This is the container for all the vortex panels -->
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

    # --- Set Window to Full Screen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    # --- Find UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $SyncHash.VortexMasterRotation = $window.FindName("VortexMasterRotation")

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
    function HandleMediaFailure {
        param([int]$SphereIndex, [string]$Reason = "Unknown Error")
        
        $handleFailureScriptBlock = {
            param($pIndex, $reason)
            $pState = $SyncHash.PlayerStates[$pIndex]
            $fUri = $pState.CurrentSource
            $fName = if ($fUri) { [System.IO.Path]::GetFileName($fUri.LocalPath) } else { "an unknown file" }
            Write-Warning "Media failed for Sphere $pIndex (File: '$fName'). Reason: $reason."
        }
        & $handleFailureScriptBlock -pIndex $SphereIndex -reason $Reason

        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState = $SyncHash.PlayerStates[$SphereIndex]
            if ($playerState.IsFailed) { return }
            $playerState.IsFailed = $true

            # For spheres, we clear the content by setting the VisualBrush's visual to null and making the background black.
            $playerState.VisualBrush.Visual = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black

            if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
                $SyncHash.SelectedFiles = @($SyncHash.SelectedFiles | Where-Object { $_ -ne $playerState.CurrentSource.LocalPath })
            }

            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{
                Interval = [TimeSpan]::FromMilliseconds(100); Tag = $SphereIndex
            }
            $recoveryTimer.Add_Tick({
                $timer = $args[0]; $pIndex = $timer.Tag; $timer.Stop()
                Start-NextMediaOnSphere -SphereIndex $pIndex
            })
            $playerState.RecoveryTimer = $recoveryTimer
            $recoveryTimer.Start()
        })
    }

    function Start-NextMediaOnSphere {
        param([int]$SphereIndex)
        
        $playerState = $SyncHash.PlayerStates[$SphereIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) {
            try { $playerState.FfmpegProcess.Kill() } catch {}
        }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath

        $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent # Reset background on new media load

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        if ($ImageExtensions -contains $extension) {
            try {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit()
                
                $imageControl = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $imageControl
                $playerState.IsFailed = $false

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(10)
                $timer.Tag = $SphereIndex
                $timer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaOnSphere -SphereIndex $idx })
                $playerState.MediaTimer = $timer
                $timer.Start()
            } catch {
                HandleMediaFailure -SphereIndex $SphereIndex -Reason "Failed to load image file."
            }
        }
        else { # It's a video, use FFmpeg
            try {
                $psi_probe = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                    FileName = "ffprobe.exe"; Arguments = "-v error -select_streams v:0 -show_entries stream=width,height -of csv=s=x:p=0 `"$filePath`"";
                    RedirectStandardOutput = $true; UseShellExecute = $false; CreateNoWindow = $true
                }
                $probe_proc = [System.Diagnostics.Process]::Start($psi_probe)
                $ffprobeOutput = $probe_proc.StandardOutput.ReadToEnd()
                $probe_proc.WaitForExit()

                $width, $height = $ffprobeOutput -split 'x'
                if (($probe_proc.ExitCode -ne 0) -or -not ($width -and $height -and [int]$width -gt 0 -and [int]$height -gt 0)) { throw "Invalid dimensions or corrupt file." }

                $writeableBmpFront = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $imageControlFront = New-Object System.Windows.Controls.Image -Property @{ Source = $writeableBmpFront; Stretch = 'Fill' }
                
                $playerState.ContentPresenter.Content = $imageControlFront
                $playerState.IsFailed = $false

                $loopArg = if ($SyncHash.SelectedFiles.Count -le $sphereCount) { "-stream_loop -1" } else { "" }
                $args = "-hide_banner -loglevel error $loopArg -i `"$filePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"
                $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                    FileName = "ffmpeg.exe"; Arguments = $args; RedirectStandardOutput = $true
                    UseShellExecute = $false; CreateNoWindow = $true
                }
                $proc = [System.Diagnostics.Process]::Start($psi)
                $playerState.FfmpegProcess = $proc

                $stream = $proc.StandardOutput.BaseStream
                $bytesPerFrame = [int]$width * [int]$height * 3
                $buffer = New-Object byte[] $bytesPerFrame
                $rect = [System.Windows.Int32Rect]::new(0, 0, [int]$width, [int]$height)
                $stride = [int]$width * 3

                $frameTimer = New-Object System.Windows.Threading.DispatcherTimer
                $frameTimer.Interval = [TimeSpan]::FromMilliseconds(33) # ~30fps
                $frameTimer.Tag = $SphereIndex

                $tickScriptBlock = {
                    $timer = $args[0]; $pIndex = $timer.Tag
                    try {
                        $totalRead = 0
                        while ($totalRead -lt $bytesPerFrame) {
                            $read = $stream.Read($buffer, $totalRead, $bytesPerFrame - $totalRead)
                            if ($read -le 0) { # End of stream
                                $timer.Stop(); Start-NextMediaOnSphere -SphereIndex $pIndex
                                return
                            }
                            $totalRead += $read
                        }
                        if ($totalRead -eq $bytesPerFrame) {
                            $writeableBmpFront.Lock(); $writeableBmpFront.WritePixels($rect, $buffer, $stride, 0); $writeableBmpFront.Unlock()
                        }
                    } catch { $timer.Stop() }
                }
                $frameTimer.Add_Tick($tickScriptBlock.GetNewClosure())
                $playerState.MediaTimer = $frameTimer
                $frameTimer.Start()
            } catch {
                HandleMediaFailure -SphereIndex $SphereIndex -Reason $_.Exception.Message
            }
        }
    }

    # --- Sphere Generation Function ---
    function New-SphereMesh {
        param(
            [double]$radius = 1.0,
            [int]$slices = 64, # Longitude
            [int]$stacks = 32  # Latitude
        )

        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

        for ($stack = 0; $stack -le $stacks; $stack++) {
            $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks
            $y = $radius * [Math]::Sin($phi)
            $r = $radius * [Math]::Cos($phi)

            for ($slice = 0; $slice -le $slices; $slice++) {
                $theta = $slice * 2 * [Math]::PI / $slices
                $x = $r * [Math]::Cos($theta)
                $z = $r * [Math]::Sin($theta)

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

    $sphereMesh = New-SphereMesh -radius 1.0

    for ($i = 0; $i -lt $sphereCount; $i++) {
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        [void]$mediaHostGrid.Children.Add($contentPresenter)

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'
            TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        [void]$mediaHostGrid.Children.Add($overlayTextBlock)
        
        # Add a second textblock for the back side
        $overlayTextBlockBack = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        [void]$mediaHostGrid.Children.Add($overlayTextBlockBack)

        # Create a VisualBrush from the 2D content
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $sphereMaterial = New-Object $materialType -Property @{ Brush = $visualBrush }

        # Create the GeometryModel3D which combines the shape (sphereMesh) and the surface (sphereMaterial)
        $sphereGeometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{
            Geometry = $sphereMesh
            Material = $sphereMaterial # Front material
            BackMaterial = $sphereMaterial # Back material
        }

        $sphereContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $sphereContainer.Content = $sphereGeometryModel

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
        $scaleTransform = New-Object System.Windows.Media.Media3D.ScaleTransform3D
        [void]$transformGroup.Children.Add($rotateTransform)
        [void]$transformGroup.Children.Add($translateTransform)
        [void]$transformGroup.Children.Add($scaleTransform)
        $sphereContainer.Transform = $transformGroup
        
        [void]$vortexContainer.Children.Add($sphereContainer)
        
        $SyncHash.PlayerStates[$i] = @{
            ContentPresenter = $contentPresenter
            VisualBrush      = $visualBrush # Store the brush to modify it later
            OverlayTextBlock = $overlayTextBlock
            OverlayTextBlockBack = $overlayTextBlockBack
            MediaHostGrid    = $mediaHostGrid
            TranslateTransform = $translateTransform
            RotateTransform  = $rotateTransform
            ScaleTransform   = $scaleTransform
            # Stagger the start angle across the *entire* spiral path (720 degrees)
            CurrentAngle = (720.0 / $sphereCount) * $i 
            MediaTimer = $null
            RecoveryTimer = $null
            FfmpegProcess = $null
            IsFailed = $false
        }
    }

    # --- Master Vortex Animation ---
    $vortexAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
        From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(45); RepeatBehavior = "Forever"
    }
    $SyncHash.VortexAnim = $vortexAnim
    # We will start this animation in the Window.Loaded event

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
        } else {
            $SyncHash.pauseButton.Content = "Pause"
            # Also resume the master rotation
            $SyncHash.VortexAnim.From = $SyncHash.VortexMasterRotation.Angle
            $SyncHash.VortexMasterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.VortexAnim)
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        }
    })

    $changeSpeed = {
        param($multiplier)
        $SyncHash.SpeedMultiplier *= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        $SyncHash.VortexAnim.Duration = [TimeSpan]::FromSeconds(45.0 / $SyncHash.SpeedMultiplier)
    }
    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 0.5 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 2.0 })

    $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        if ($SyncHash.ControlsHidden) {
            $controlsPanel.Visibility = 'Visible'
            $SyncHash.ControlsHidden = $false
        } else {
            $controlsPanel.Visibility = 'Collapsed'
            $SyncHash.ControlsHidden = $true
        }
    })

    # --- Per-Frame Animation Handler ---
    $animationHandler = {
        param($sender, $e)

        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        # Vortex parameters
        $rotationSpeed = 30.0 * $SyncHash.SpeedMultiplier # degrees per second
        $startRadius = 3.5  # Make the funnel narrower at the top
        $endRadius = 0.5    # Make the funnel narrower at the bottom
        $startY = 5.0       # Start the funnel higher up
        $endY = -20.0       # Make the funnel much deeper to reach the bottom of the screen
        $maxScale = 1.2     # Reduce the max size of the spheres to prevent overlap
        $minScale = 0.1     # Make the spheres even smaller at the end for more perspective
        $totalAngle = 360 * 2 # Two full rotations to go from top to bottom

        foreach ($i in 0..($sphereCount-1)) {
            $panelState = $SyncHash.PlayerStates[$i]
            
            # Update and wrap the angle
            $panelState.CurrentAngle = ($panelState.CurrentAngle + ($rotationSpeed * $elapsed)) % $totalAngle
            $angle = $panelState.CurrentAngle
            
            # Convert to radians for trig functions
            $angleRad = $angle * ([Math]::PI / 180.0)

            if ($SyncHash.VortexOrientation -eq "Vertical") {
                # --- VERTICAL VORTEX LOGIC ---
                $progress = $angle / $totalAngle
                $currentRadius = $startRadius - ($progress * ($startRadius - $endRadius))
                $currentY = $startY - ($progress * ($startY - $endY))
                $currentScale = $maxScale - ($progress * ($maxScale - $minScale))

                $currentX = $currentRadius * [Math]::Cos($angleRad)
                $currentZ = $currentRadius * [Math]::Sin($angleRad)

                $panelState.TranslateTransform.OffsetX = $currentX
                $panelState.TranslateTransform.OffsetY = $currentY
                $panelState.TranslateTransform.OffsetZ = $currentZ

                $lookAtY = $currentY - 1.5
                [System.Windows.Media.Media3D.Point3D]$lookAtTarget = New-Object System.Windows.Media.Media3D.Point3D(0, $lookAtY, 0)
                [System.Windows.Media.Media3D.Point3D]$position = New-Object System.Windows.Media.Media3D.Point3D($currentX, $currentY, $currentZ)
            } else {
                # --- HORIZONTAL VORTEX LOGIC ---
                $startRadiusH = 5.0; $endRadiusH = 0.5
                $startZ = 5.0; $endZ = -20.0

                $progress = $angle / $totalAngle
                $currentRadius = $startRadiusH - ($progress * ($startRadiusH - $endRadiusH))
                $currentZ = $startZ - ($progress * ($startZ - $endZ))
                $currentScale = $maxScale - ($progress * ($maxScale - $minScale))

                $currentX = $currentRadius * [Math]::Cos($angleRad)
                $currentY = $currentRadius * [Math]::Sin($angleRad)

                $panelState.TranslateTransform.OffsetX = $currentX
                $panelState.TranslateTransform.OffsetY = $currentY
                $panelState.TranslateTransform.OffsetZ = $currentZ

                $lookAtZ = $currentZ - 1.5
                [System.Windows.Media.Media3D.Point3D]$lookAtTarget = New-Object System.Windows.Media.Media3D.Point3D(0, 0, $lookAtZ)
                [System.Windows.Media.Media3D.Point3D]$position = New-Object System.Windows.Media.Media3D.Point3D($currentX, $currentY, $currentZ)
            }

            # --- COMMON LOGIC FOR SCALING AND ROTATION ---
            $panelState.ScaleTransform.ScaleX = $currentScale
            $panelState.ScaleTransform.ScaleY = $currentScale
            $panelState.ScaleTransform.ScaleZ = $currentScale

            $forward = [System.Windows.Media.Media3D.Point3D]::Subtract($position, $lookAtTarget)
            if ($forward.Length -gt 0) { $forward.Normalize() }

            $up = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0)
            if ($SyncHash.VortexOrientation -ne "Vertical") {
                # For horizontal, if the sphere is directly above/below, the standard 'up' is parallel to 'forward'.
                # We use a different 'up' vector in that case to prevent issues.
                if ([Math]::Abs([System.Windows.Media.Media3D.Vector3D]::DotProduct($forward, $up)) -gt 0.99) {
                    $up = New-Object System.Windows.Media.Media3D.Vector3D(1, 0, 0)
                }
            }

            $right = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($up, $forward)
            if ($right.Length -gt 0) { $right.Normalize() }
            $newUp = [System.Windows.Media.Media3D.Vector3D]::CrossProduct($forward, $right)

            $matrix = [System.Windows.Media.Media3D.Matrix3D]::Identity
            $matrix.M11 = $right.X;   $matrix.M12 = $right.Y;   $matrix.M13 = $right.Z
            $matrix.M21 = $newUp.X;   $matrix.M22 = $newUp.Y;   $matrix.M23 = $newUp.Z
            $matrix.M31 = $forward.X; $matrix.M32 = $forward.Y; $matrix.M33 = $forward.Z

            $qw = 0; $qx = 0; $qy = 0; $qz = 0
            $trace = $matrix.M11 + $matrix.M22 + $matrix.M33

            if ($trace -gt 0) {
                $s = 0.5 / [Math]::Sqrt($trace + 1.0)
                $qw = 0.25 / $s; $qx = ($matrix.M32 - $matrix.M23) * $s
                $qy = ($matrix.M13 - $matrix.M31) * $s; $qz = ($matrix.M21 - $matrix.M12) * $s
            } elseif (($matrix.M11 -gt $matrix.M22) -and ($matrix.M11 -gt $matrix.M33)) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M11 - $matrix.M22 - $matrix.M33)
                $qw = ($matrix.M32 - $matrix.M23) / $s; $qx = 0.25 * $s
                $qy = ($matrix.M12 + $matrix.M21) / $s; $qz = ($matrix.M13 + $matrix.M31) / $s
            } elseif ($matrix.M22 -gt $matrix.M33) {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M22 - $matrix.M11 - $matrix.M33)
                $qw = ($matrix.M13 - $matrix.M31) / $s; $qx = ($matrix.M12 + $matrix.M21) / $s
                $qy = 0.25 * $s; $qz = ($matrix.M23 + $matrix.M32) / $s
            } else {
                $s = 2.0 * [Math]::Sqrt(1.0 + $matrix.M33 - $matrix.M11 - $matrix.M22)
                $qw = ($matrix.M21 - $matrix.M12) / $s; $qx = ($matrix.M13 + $matrix.M31) / $s
                $qy = ($matrix.M23 + $matrix.M32) / $s; $qz = 0.25 * $s
            }
            $quaternion = New-Object System.Windows.Media.Media3D.Quaternion($qx, $qy, $qz, $qw)
            $panelState.RotateTransform.Rotation = (New-Object System.Windows.Media.Media3D.QuaternionRotation3D($quaternion))
        }
    }

    # --- Window Events ---
    $window.Add_KeyDown({
        param($sender, $e)
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
        if ($SyncHash.PlayerStates) {
            foreach ($i in $SyncHash.PlayerStates.Keys) {
                $playerState = $SyncHash.PlayerStates[$i]
                if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) {
                    try { $playerState.FfmpegProcess.Kill() } catch {}
                }
            }
        }
    })

    # --- Start the show ---
    $window.Add_Loaded({
        # Apply Text Overlay Settings
        $textToSet = $null
        if ($SyncHash.RbSelection -eq "Custom") {
            $textToSet = $SyncHash.CustomText
        }

        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $fontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $fontWeight = if ($SyncHash.IsBold) { 'Bold' } else { 'Normal' }
            $fontStyle = if ($SyncHash.IsItalic) { 'Italic' } else { 'Normal' }

            for ($i = 0; $i -lt $sphereCount; $i++) {
                $tbFront = $SyncHash.PlayerStates[$i].OverlayTextBlock
                $tbBack = $SyncHash.PlayerStates[$i].OverlayTextBlockBack
                foreach ($tb in @($tbFront, $tbBack)) {
                    $tb.Foreground = $brush.Clone()
                    $tb.FontFamily = $fontFamily
                    $tb.FontSize = $SyncHash.FontSize
                    $tb.FontWeight = $fontWeight
                    $tb.FontStyle = $fontStyle
                    if ($textToSet) { $tb.Text = $textToSet }
                }
            }
        }

        # Start media on each sphere
        for ($i = 0; $i -lt $sphereCount; $i++) {
            Start-NextMediaOnSphere -SphereIndex $i
        }

        # Start the animation loop
        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $SyncHash.VortexMasterRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.VortexAnim)
        [System.Windows.Media.CompositionTarget]::add_Rendering($animationHandler)
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}