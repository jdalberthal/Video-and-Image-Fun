<#
.SYNOPSIS
    Displays media on a series of rotating 3D planes arranged in a pinwheel or kaleidoscope pattern using FFmpeg.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto multiple,
    independently rotating 3D planes. These planes are arranged in a circular, pinwheel-like
    formation that also rotates as a whole, creating a dynamic kaleidoscope effect.

    This version uses FFmpeg for video decoding, providing support for a wide range of video formats
    without relying on system-installed codecs.

    The 3D view is interactive, with controls to pause, change rotation speed, and randomize
    the rotation axes for different visual effects.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosPinwheelFfmpeg.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the 3D pinwheel window.

.NOTES
    Name:           Show-ImagesVideosPinwheelFfmpeg.ps1
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
$ExternalButtonName = "Pinwheel Kaleidoscope`n(FFmpeg)"
$ScriptDescription = "Displays media on multiple, independently rotating 3D planes arranged in a pinwheel pattern. Uses FFmpeg for broad video format support."
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
    $SelectForm.Text = "Pinwheel (FFmpeg) - Media Selector"
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
            }
        })

    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
            }
        })

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
        Title="Pinwheel Kaleidoscope"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,8" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
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

            <!-- This is the container for all the pinwheel blades -->
            <ModelVisual3D x:Name="PinwheelContainer">
                <ModelVisual3D.Transform>
                    <RotateTransform3D>
                        <RotateTransform3D.Rotation>
                            <AxisAngleRotation3D x:Name="PinwheelRotation" Axis="0,0,1" Angle="0"/>
                        </RotateTransform3D.Rotation>
                    </RotateTransform3D>
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
    $SyncHash.Window = $window

    # --- Set Window to Full Screen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    # --- Find UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    # Find the rotation object AFTER the window is loaded and store it for reliable access.
    $SyncHash.PinwheelRotation = $window.FindName("PinwheelRotation")

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
        param($BladeIndex, [string]$Reason = "Unknown Error")
        
        $handleFailureScriptBlock = {
            param($bIndex, $reason)
            $pState = $SyncHash.PlayerStates[$bIndex]
            $fUri = $pState.CurrentSource
            $fName = if ($fUri) { [System.IO.Path]::GetFileName($fUri.LocalPath) } else { "an unknown file" }
            Write-Warning "Media failed for Blade $bIndex (File: '$fName'). Reason: $reason. Removing from list and continuing."
        }
        & $handleFailureScriptBlock -bIndex $BladeIndex -reason $Reason

        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState = $SyncHash.PlayerStates[$BladeIndex]
            if ($playerState.IsFailed) { return }
            $playerState.IsFailed = $true

            $playerState.ContentPresenter.Content = $null
            $playerState.ContentPresenterBack.Content = $null
            $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black
            $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Black

            if ($playerState.CurrentSource -and $SyncHash.SelectedFiles.Count -gt 1) {
                $SyncHash.SelectedFiles = @($SyncHash.SelectedFiles | Where-Object { $_ -ne $playerState.CurrentSource.LocalPath })
            }

            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{
                Interval = [TimeSpan]::FromMilliseconds(100); Tag = $BladeIndex
            }
            $recoveryTimer.Add_Tick({
                $timer = $args[0]; $bIndex = $timer.Tag; $timer.Stop()
                Start-NextMediaOnBlade -BladeIndex $bIndex
            })
            $playerState.RecoveryTimer = $recoveryTimer
            $recoveryTimer.Start()
        })
    }

    function Start-NextMediaOnBlade {
        param([int]$BladeIndex)
        
        $playerState = $SyncHash.PlayerStates[$BladeIndex]
        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) {
            try { $playerState.FfmpegProcess.Kill() } catch {}
        }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath

        $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent
        $playerState.MediaHostGridBack.Background = [System.Windows.Media.Brushes]::Transparent

        if ($SyncHash.RbSelection -eq "Filename") {
            $playerState.OverlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
            $playerState.OverlayTextBlockBack.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        if ($ImageExtensions -contains $extension) {
            try {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit()
                
                # Create two separate image controls for front and back.
                $imageFront = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $imageBack = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $imageFront
                $playerState.ContentPresenterBack.Content = $imageBack
                $playerState.IsFailed = $false

                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromSeconds(10)
                $timer.Tag = $BladeIndex
                $timer.Add_Tick({ $t = $args[0]; $idx = $t.Tag; $t.Stop(); Start-NextMediaOnBlade -BladeIndex $idx })
                $playerState.MediaTimer = $timer
                $timer.Start()
            } catch {
                HandleMediaFailure -BladeIndex $BladeIndex -Reason "Failed to load image file."
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

                # Create two separate WriteableBitmaps and Image controls for front and back.
                $writeableBmpFront = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $writeableBmpBack = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $imageControlFront = New-Object System.Windows.Controls.Image -Property @{ Source = $writeableBmpFront; Stretch = 'Fill' }
                $imageControlBack = New-Object System.Windows.Controls.Image -Property @{ Source = $writeableBmpBack; Stretch = 'Fill' }
                
                $playerState.ContentPresenter.Content = $imageControlFront
                $playerState.ContentPresenterBack.Content = $imageControlBack
                $playerState.IsFailed = $false

                $loopArg = if ($SyncHash.SelectedFiles.Count -le $bladeCount) { "-stream_loop -1" } else { "" }
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
                $frameTimer.Tag = $BladeIndex

                $tickScriptBlock = {
                    $timer = $args[0]; $bIndex = $timer.Tag
                    try {
                        $totalRead = 0
                        while ($totalRead -lt $bytesPerFrame) {
                            $read = $stream.Read($buffer, $totalRead, $bytesPerFrame - $totalRead)
                            if ($read -le 0) { # End of stream
                                $timer.Stop()
                                Start-NextMediaOnBlade -BladeIndex $bIndex
                                return
                            }
                            $totalRead += $read
                        }
                        if ($totalRead -eq $bytesPerFrame) {
                            # Update BOTH bitmaps with the same frame data.
                            $writeableBmpFront.Lock(); $writeableBmpFront.WritePixels($rect, $buffer, $stride, 0); $writeableBmpFront.Unlock()
                            $writeableBmpBack.Lock(); $writeableBmpBack.WritePixels($rect, $buffer, $stride, 0); $writeableBmpBack.Unlock()
                        }
                    } catch { $timer.Stop() }
                }
                $frameTimer.Add_Tick($tickScriptBlock.GetNewClosure())
                $playerState.MediaTimer = $frameTimer
                $frameTimer.Start()
            } catch {
                HandleMediaFailure -BladeIndex $BladeIndex -Reason $_.Exception.Message
            }
        }
    }

    # --- Pinwheel Blade Generation ---
    $pinwheelContainer = $window.FindName("PinwheelContainer")
    $bladeCount = 8
    $angleIncrement = 360 / $bladeCount
    $bladeWidth = 1.5
    $bladeHeight = 4.0
    $bladeTiltAngle = 45 # Degrees
    $pinwheelRadius = 1.5

    $bladeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2), (-$bladeHeight/2), 0)) )
    $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2), (-$bladeHeight/2), 0)) )
    $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D( ($bladeWidth/2),  ($bladeHeight/2), 0)) )
    $bladeMesh.Positions.Add( (New-Object System.Windows.Media.Media3D.Point3D((-$bladeWidth/2),  ($bladeHeight/2), 0)) )
    $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(1); $bladeMesh.TriangleIndices.Add(2)
    $bladeMesh.TriangleIndices.Add(0); $bladeMesh.TriangleIndices.Add(2); $bladeMesh.TriangleIndices.Add(3)
    $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,1)) )
    $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,1)) )
    $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(1,0)) )
    $bladeMesh.TextureCoordinates.Add( (New-Object System.Windows.Point(0,0)) )

    for ($i = 0; $i -lt $bladeCount; $i++) {
        $bladeAngle = $i * $angleIncrement

        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter)

        $mediaHostGridBack = New-Object System.Windows.Controls.Grid
        $contentPresenterBack = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGridBack.Children.Add($contentPresenterBack)

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'
            TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $mediaHostGrid.Children.Add($overlayTextBlock)

        $overlayTextBlockBack = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'
            TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $mediaHostGridBack.Children.Add($overlayTextBlockBack)

        $bladeViewport = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $bladeViewport.Geometry = $bladeMesh
        $bladeViewport.Visual = $mediaHostGrid

        $bladeViewportBack = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
        $bladeViewportBack.Geometry = $bladeMesh
        $bladeViewportBack.Visual = $mediaHostGridBack
        $bladeViewportBack.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D('0,1,0', 180)))

        $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
        $bladeMaterial = New-Object $materialType
        [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($bladeMaterial, $true)
        $bladeViewport.Material = $bladeMaterial

        $bladeMaterialBack = New-Object $materialType
        [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($bladeMaterialBack, $true)
        $bladeViewportBack.Material = $bladeMaterialBack

        $bladeContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D
        $bladeContainer.Children.Add($bladeViewport)
        $bladeContainer.Children.Add($bladeViewportBack)

        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup

        $bladeRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{
            Axis = New-Object System.Windows.Media.Media3D.Vector3D(0, 1, 0); Angle = 0
        }
        $transformGroup.Children.Add( (New-Object System.Windows.Media.Media3D.RotateTransform3D($bladeRotation)) )

        $tiltRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{
            Axis = New-Object System.Windows.Media.Media3D.Vector3D(1, 0, 0); Angle = $bladeTiltAngle
        }
        $transformGroup.Children.Add( (New-Object System.Windows.Media.Media3D.RotateTransform3D($tiltRotation)) )

        $transformGroup.Children.Add( (New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $pinwheelRadius, 0)) )

        $placementRotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D -Property @{
            Axis = New-Object System.Windows.Media.Media3D.Vector3D(0, 0, 1); Angle = $bladeAngle
        }
        $transformGroup.Children.Add( (New-Object System.Windows.Media.Media3D.RotateTransform3D($placementRotation)) )

        $bladeContainer.Transform = $transformGroup
        $pinwheelContainer.Children.Add($bladeContainer)

        $bladeAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
            From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(20); RepeatBehavior = "Forever"
        }
        
        $SyncHash.PlayerStates[$i] = @{
            ContentPresenter = $contentPresenter
            ContentPresenterBack = $contentPresenterBack
            OverlayTextBlock = $overlayTextBlock
            OverlayTextBlockBack = $overlayTextBlockBack
            MediaHostGrid = $mediaHostGrid
            MediaHostGridBack = $mediaHostGridBack
            BladeRotation = $bladeRotation
            BladeAnimation = $bladeAnim
            MediaTimer = $null
            RecoveryTimer = $null
            FfmpegProcess = $null
            IsFailed = $false
        }

        $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $bladeAnim)
    }

    # --- Main Pinwheel Animation ---
    $pinwheelAnim = New-Object System.Windows.Media.Animation.DoubleAnimation -Property @{
        From = 0; To = 360; Duration = [TimeSpan]::FromSeconds(60); RepeatBehavior = "Forever"
    }
    $SyncHash.PinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $pinwheelAnim)
    $SyncHash.PinwheelAnim = $pinwheelAnim

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
        if ($SyncHash.Paused) {
            $SyncHash.PinwheelAnim.From = $SyncHash.PinwheelRotation.Angle
            $SyncHash.PinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.PinwheelAnim)
            
            foreach($i in 0..($bladeCount-1)) {
                $bladeRotation = $SyncHash.PlayerStates[$i].BladeRotation
                $bladeAnim = $SyncHash.PlayerStates[$i].BladeAnimation
                $bladeAnim.From = $bladeRotation.Angle
                $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $bladeAnim)
            }

            $SyncHash.pauseButton.Content = "Pause"
            $SyncHash.Paused = $false
        } else {
            $SyncHash.PinwheelRotation.Angle = $SyncHash.PinwheelRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
            $SyncHash.PinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)

            foreach($i in 0..($bladeCount-1)) {
                $bladeRotation = $SyncHash.PlayerStates[$i].BladeRotation
                $bladeRotation.Angle = $bladeRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            }

            $SyncHash.pauseButton.Content = "Resume"
            $SyncHash.Paused = $true
        }
    })

    $SyncHash.randomAxisButton.Add_Click({
        $SyncHash.PinwheelRotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
        foreach ($i in 0..($bladeCount-1)) {
            $SyncHash.PlayerStates[$i].BladeRotation.Axis = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
        }
    })

    $changeSpeed = {
        param($multiplier)

        $wasPaused = $SyncHash.Paused
        
        if (-not $wasPaused) {
            $SyncHash.PinwheelRotation.Angle = $SyncHash.PinwheelRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
            $SyncHash.PinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            foreach($i in 0..($bladeCount-1)) {
                $bladeRotation = $SyncHash.PlayerStates[$i].BladeRotation
                $bladeRotation.Angle = $bladeRotation.GetValue([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            }
        }

        $SyncHash.SpeedMultiplier *= $multiplier
        if ($SyncHash.SpeedMultiplier -gt 16.0) { $SyncHash.SpeedMultiplier = 16.0 }
        if ($SyncHash.SpeedMultiplier -lt 0.0625) { $SyncHash.SpeedMultiplier = 0.0625 }
        $SyncHash.PinwheelAnim.Duration = [TimeSpan]::FromSeconds(60.0 / $SyncHash.SpeedMultiplier)
        0..($bladeCount-1) | ForEach-Object { $SyncHash.PlayerStates[$_].BladeAnimation.Duration = [TimeSpan]::FromSeconds(20.0 / $SyncHash.SpeedMultiplier) }

        if (-not $wasPaused) {
            $SyncHash.PinwheelAnim.From = $SyncHash.PinwheelRotation.Angle
            $SyncHash.PinwheelRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.PinwheelAnim)
            foreach($i in 0..($bladeCount-1)) {
                $bladeRotation = $SyncHash.PlayerStates[$i].BladeRotation
                $bladeAnim = $SyncHash.PlayerStates[$i].BladeAnimation
                $bladeAnim.From = $bladeRotation.Angle
                $bladeRotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $bladeAnim)
            }
        }
    }

    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 0.5 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 2.0 })

    $SyncHash.redoButton.Add_Click({
        $SyncHash.RedoClicked = $true
        $window.Close()
    })

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

            for ($i = 0; $i -lt $bladeCount; $i++) {
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

        # Start media on each blade
        for ($i = 0; $i -lt $bladeCount; $i++) {
            Start-NextMediaOnBlade -BladeIndex $i
        }
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) {
        break
    }
}
