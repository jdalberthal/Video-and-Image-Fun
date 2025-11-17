<#
.SYNOPSIS
    Displays selected images and videos on the surfaces of six floating, bouncing 3D stars using FFmpeg.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto the surfaces
    of six independently moving 3D stars in a WPF window. The stars move in straight lines and
    bounce off the edges of the screen, creating a "zero-gravity" or "screen saver" visual effect.

    A "star" is a composite 3D object made of a central sphere and six cones. All parts of a single
    star display the same media content.

    This version uses FFmpeg for video decoding, providing support for a wide range of video formats
    without relying on system-installed codecs.

    The 3D view is interactive, with controls to pause the animation and hide the UI for an
    unobstructed view.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosFloatingStarsFfmpeg.ps1

    Launches the file selection GUI. After selecting at least one file and clicking "Play", the
    script will launch the 3D window with six floating, bouncing stars.

.NOTES
    Name:           Show-ImagesVideosFloatingStarsFfmpeg.ps1
    Version:        1.0.0, 11/15/2025
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
$ExternalButtonName = "Floating Stars `n(FFmpeg)"
$ScriptDescription = "Displays media on six independently floating and bouncing 3D stars. Uses FFmpeg for broad video format support."
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
        # For bouncing animation
        StarObjects = @()
        LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
    })

    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Floating Stars (FFmpeg) - Media Selector"
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

    $HelpLabel = New-Object System.Windows.Forms.Label -Property @{
        Text      = "F1 - Help"
        AutoSize  = $true
        Location  = '700, 10'
    }
    $SelectForm.Controls.Add($HelpLabel)

    [xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help" Height="340" Width="450" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <RichTextBox x:Name="MyRichTextBox" Grid.Row="0" Margin="5" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
            <FlowDocument>
                <FlowDocument.Resources>
                    <Style TargetType="{x:Type Paragraph}">
                        <Setter Property="Margin" Value="0"/>
                    </Style>
                </FlowDocument.Resources>
                <Paragraph>
                    <Run Text="Hopefully the selection dialog is self-explanatory. :-)"/><LineBreak/>
                    <Run Text=" "/>
                </Paragraph>
                <Paragraph>
                    <Run Text="Commands for after media is playing:"/><LineBreak/>
                </Paragraph>    
                <Paragraph TextAlignment="Left" FontFamily="Consolas">
                    <Bold><Run Text="Button            : Key : Action" TextDecorations="Underline"/><LineBreak/></Bold>
                    <Run Text="X                 : Esc : Exit"/><LineBreak/>
                    <Run Text="Pause/Resume      :  P  : Pause/Resume Animation"/><LineBreak/>
                    <Run Text="Redo              :  R  : Reselect Media"/><LineBreak/>
                    <Run Text="Hide Controls     :  H  : Hide/Show Controls"/><LineBreak/><LineBreak/>
                    <Run Text="*Click a star to Pause/Resume*"/><LineBreak/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0" IsDefault="True"/>
    </Grid>
</Window>
"@

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

    $SelectForm.KeyPreview = $true
    $SelectForm.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq "F1") {
            $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
            $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
            $OkButton = $PopupWindow.FindName("OKButton")
            $OkButton.Add_Click({ $PopupWindow.Close() })
            $PopupWindow.ShowDialog() | Out-Null
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
        Title="Floating Stars (FFmpeg)"
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
                $i0 = $stack * ($slices + 1) + $slice
                $i1 = ($stack + 1) * ($slices + 1) + $slice

                $mesh.TriangleIndices.Add($i0); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i0 + 1)
                $mesh.TriangleIndices.Add($i0 + 1); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i1 + 1)
            }
        }
        return $mesh
    }

    function New-ConeMesh {
        param(
            [double]$radius = 0.4,
            [double]$height = 1.5,
            [int]$slices = 64
        )

        $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

        # Apex vertex
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $height, 0))
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0))

        # Base vertices
        for ($i = 0; $i -le $slices; $i++) {
            $theta = $i * 2 * [Math]::PI / $slices
            $x = $radius * [Math]::Cos($theta)
            $z = $radius * [Math]::Sin($theta)
            $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, 0, $z))
            $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $slices, 1))
        }

        # Base center vertex for the bottom cap
        $baseCenterIndex = $mesh.Positions.Count
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, 0, 0))
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0.5))

        # Add triangle indices for the cone sides
        for ($i = 0; $i -lt $slices; $i++) {
            $mesh.TriangleIndices.Add(0); $mesh.TriangleIndices.Add($i + 2); $mesh.TriangleIndices.Add($i + 1)
        }

        # Add triangle indices for the base cap
        for ($i = 0; $i -lt $slices; $i++) {
            $mesh.TriangleIndices.Add($baseCenterIndex); $mesh.TriangleIndices.Add($i + 1); $mesh.TriangleIndices.Add($i + 2)
        }

        return $mesh
    }

    function New-StarVisual3D {
        param(
            [int]$StarIndex,
            [hashtable]$SyncHash
        )

        $starContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D

        $sphereRadius = 0.5
        $coneHeight = 1.5
        $coneRadius = 0.4

        $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 64 -stacks 32
        $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 64

        $visuals = @{
            "Middle" = @{ "Mesh" = $sphereMesh; "Transform" = (New-Object System.Windows.Media.Media3D.TranslateTransform3D) }
            "Top"    = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)) }
            "Bottom" = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Right"  = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Left"   = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Front"  = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
            "Back"   = @{ "Mesh" = $coneMesh;   "Transform" = (New-Object System.Windows.Media.Media3D.Transform3DGroup) }
        }

        # Configure transforms for cones
        $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1)))
        $visuals.Bottom.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))

        $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), -90))))
        $visuals.Right.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D($sphereRadius, 0, 0)))

        $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), 90))))
        $visuals.Left.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(-$sphereRadius, 0, 0)))

        $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 90))))
        $visuals.Front.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, $sphereRadius)))

        $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D([System.Windows.Media.Media3D.AxisAngleRotation3D]::new([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), -90))))
        $visuals.Back.Transform.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, -$sphereRadius)))

        # Create the 2D content (a Grid with an Image and TextBlock) for this star.
        $imageControl = New-Object System.Windows.Controls.Image -Property @{ Stretch = 'Fill' }
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap';
            TextAlignment = 'Center'; IsHitTestVisible = $false; Margin = '10,0,10,0'
        }
        $visualHostGrid = New-Object System.Windows.Controls.Grid -Property @{ Background = [System.Windows.Media.Brushes]::Black }
        
        [void]$visualHostGrid.Children.Add($imageControl)
        [void]$visualHostGrid.Children.Add($overlayTextBlock)

        # Create a single VisualBrush from the 2D content. This brush can be shared.
        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $visualHostGrid }
        # Create a visible ModelVisual3D for each part of the star.
        # Each part gets its own material, but all materials will share the same VisualBrush.
        foreach ($key in $visuals.Keys) {
            $part = $visuals[$key]
            $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D
            $geometryModel.Geometry = $part.Mesh
            $geometryModel.Transform = $part.Transform

            $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
            $geometryModel.Material = New-Object $materialType -Property @{ Brush = $visualBrush }

            $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D
            $modelVisual.Content = $geometryModel
            $starContainer.Children.Add($modelVisual)
        }

        $SyncHash.PlayerState[$StarIndex] = @{
            ImageControl  = $imageControl
            Overlay       = $overlayTextBlock
            IsImage       = $false
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
            ImageTimer    = $null
            IsFailed      = $false
            RecoveryTimer = $null
            FfmpegProcess = $null
            FrameReader   = $null
            WriteableBmp  = $null
            CurrentPath   = $null
        }

        return $starContainer
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

    function Assign-NextMediaToStar {
        param([int]$StarIndex, [hashtable]$SyncHash)

        $playerState = $SyncHash.PlayerState[$StarIndex]
        if (-not $playerState) { return }

        # Stop any existing timers or processes for this star
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
            $playerState.PlaybackStopwatch.Restart() # Start timer for consistency with video path
            $timer.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds)
            $timer.Tag = $StarIndex
            $timer.Add_Tick({
                $t = $args[0]; $sIndex = $t.Tag; $t.Stop()
                Assign-NextMediaToStar -StarIndex $sIndex -SyncHash $SyncHash
            })
            $playerState.ImageTimer = $timer
            $timer.Start()
            $playerState.IsFailed = $false # Mark as successful
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
                if (($probe_proc.ExitCode -ne 0) -or -not ($width -and $height -and [int]$width -gt 0 -and [int]$height -gt 0)) { throw "ffprobe could not get valid dimensions." }

                $playerState.WriteableBmp = New-Object System.Windows.Media.Imaging.WriteableBitmap([int]$width, [int]$height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
                $playerState.ImageControl.Source = $playerState.WriteableBmp

                $loopArg = if ($SyncHash.SelectedFiles.Count -le 6) { "-stream_loop -1" } else { "" }

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
                $frameTimer.Tag = $StarIndex

                $playerState.PlaybackStopwatch.Restart() # Start timing the playback
                $playerState.IsFailed = $false # Tentatively mark as successful

                $tickScriptBlock = {
                    $timer = $args[0]; $sIndex = $timer.Tag
                    $pState = $SyncHash.PlayerState[$sIndex]
                    $totalRead = 0
                    while ($totalRead -lt $bytesPerFrame) {
                        $bytesRead = $stream.Read($buffer, $totalRead, $bytesPerFrame - $totalRead)
                        
                        # Check for instant failure
                        $pState.PlaybackStopwatch.Stop()
                        if ($bytesRead -le 0 -and $pState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 1000) {
                            Handle-MediaFailure -StarIndex $sIndex -Reason "Playback stream ended instantly." -SyncHash $SyncHash
                            return
                        }

                        if ($bytesRead -le 0) { # End of stream or error
                            $timer.Stop()
                            Assign-NextMediaToStar -StarIndex $sIndex -SyncHash $SyncHash
                            return
                        }
                        $totalRead += $bytesRead
                    }
                    if ($totalRead -eq $bytesPerFrame) {
                        $pState.WriteableBmp.Lock()
                        $pState.WriteableBmp.WritePixels($rect, $buffer, $stride, 0)
                        $pState.WriteableBmp.Unlock()
                    }
                    $pState.PlaybackStopwatch.Restart() # Restart for next frame interval check
                }
                $frameTimer.Add_Tick($tickScriptBlock.GetNewClosure())
                $playerState.ImageTimer = $frameTimer # Reuse ImageTimer property for cleanup
                $frameTimer.Start()
            } catch {
                Handle-MediaFailure -StarIndex $StarIndex -Reason $_.Exception.Message -SyncHash $SyncHash
            }
        }
    }

    function Handle-MediaFailure {
        param([int]$StarIndex, [string]$Reason, [hashtable]$SyncHash)
        $playerState = $SyncHash.PlayerState[$StarIndex]
        if (-not $playerState -or $playerState.IsFailed) { return }
        $playerState.IsFailed = $true

        $failedPath = $playerState.CurrentPath
        $fileName = if ($failedPath) { [System.IO.Path]::GetFileName($failedPath) } else { "an unknown file" }
        Write-Warning "Media failed for Star $StarIndex (File: '$fileName'). Reason: $Reason. Removing from list and continuing."

        # Remove the bad file from the list in a thread-safe way
        if ($failedPath) {
            $SyncHash.SelectedFiles = @($SyncHash.SelectedFiles | Where-Object { $_ -ne $failedPath })
        }

        # Immediately try to load the next file.
        # This needs to be on the main thread, but Handle-MediaFailure is called from there.
        Assign-NextMediaToStar -StarIndex $StarIndex -SyncHash $SyncHash
    }

    # --- Create and Prepare 6 Stars ---
    for ($i = 0; $i -lt 6; $i++) {
        $starVisual = New-StarVisual3D -StarIndex $i -SyncHash $SyncHash

        # Set up transforms for movement and rotation
        $transformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
        $rotateTransform = New-Object System.Windows.Media.Media3D.RotateTransform3D
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D
        $transformGroup.Children.Add($rotateTransform)
        $transformGroup.Children.Add($translateTransform)
        $starVisual.Transform = $transformGroup

        # Store object references for the animation loop
        $starObject = [pscustomobject]@{
            Visual = $starVisual
            Translate = $translateTransform
            Rotate = $rotateTransform
            # Create a random direction vector, normalize it to get a pure direction, then multiply by a random speed.
            # This ensures a more uniform "outward burst" from the center.
            Velocity = {
                $randomVector = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -1.0 -Maximum 1.0))
                if ($randomVector.Length -gt 0) { $randomVector.Normalize() } # Avoid division by zero if the vector is (0,0,0)
                return $randomVector * (Get-Random -Minimum 1.5 -Maximum 3.0)
            }.Invoke()
            RotationVelocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0), (Get-Random -Minimum -20.0 -Maximum 20.0))
            CurrentRotation = New-Object System.Windows.Media.Media3D.Quaternion(0,0,0,1)
        }
        $SyncHash.StarObjects += $starObject

        [void]$mainViewport.Children.Add($starVisual)
    }

    # --- Initial Media Loading ---
    for ($i = 0; $i -lt ($SyncHash.PlayerState.Keys.Count); $i++) {
        Assign-NextMediaToStar -StarIndex $i -SyncHash $SyncHash
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
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
        } else {
            $SyncHash.pauseButton.Content = "Pause"
            # Reset frame timer to avoid a large jump after unpausing
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        }
    }
    $SyncHash.pauseButton.Add_Click($pauseOrResumeAnimations)

    # --- Bouncing Animation Logic ---
    $animationHandler = {
        param($sender, $e)

        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime

        # Define boundaries (approximate based on camera FOV and distance)
        $xBoundary = 9
        $yBoundary = 5
        $zBoundary = 3

        foreach ($star in $SyncHash.StarObjects) {
            # Update Position
            $newX = $star.Translate.OffsetX + ($star.Velocity.X * $elapsed)
            $newY = $star.Translate.OffsetY + ($star.Velocity.Y * $elapsed)
            $newZ = $star.Translate.OffsetZ + ($star.Velocity.Z * $elapsed)

            # Since Vector3D is a struct, its properties are read-only. We must create a new vector to change it.
            $velX = $star.Velocity.X
            $velY = $star.Velocity.Y
            $velZ = $star.Velocity.Z

            # Check for bounces
            if (($newX -gt $xBoundary -and $velX -gt 0) -or ($newX -lt -$xBoundary -and $velX -lt 0)) {
                $velX *= -1
            }
            if (($newY -gt $yBoundary -and $velY -gt 0) -or ($newY -lt -$yBoundary -and $velY -lt 0)) {
                $velY *= -1
            }
            if (($newZ -gt $zBoundary -and $velZ -gt 0) -or ($newZ -lt -$zBoundary -and $velZ -lt 0)) {
                $velZ *= -1
            }
            $star.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)

            $star.Translate.OffsetX = $newX
            $star.Translate.OffsetY = $newY
            $star.Translate.OffsetZ = $newZ

            # Update Rotation
            $axisX = New-Object System.Windows.Media.Media3D.Vector3D(1,0,0)
            $axisY = New-Object System.Windows.Media.Media3D.Vector3D(0,1,0)
            $axisZ = New-Object System.Windows.Media.Media3D.Vector3D(0,0,1)

            $deltaRotationX = New-Object System.Windows.Media.Media3D.Quaternion($axisX, ($star.RotationVelocity.X * $elapsed))
            $deltaRotationY = New-Object System.Windows.Media.Media3D.Quaternion($axisY, ($star.RotationVelocity.Y * $elapsed))
            $deltaRotationZ = New-Object System.Windows.Media.Media3D.Quaternion($axisZ, ($star.RotationVelocity.Z * $elapsed))
            
            $star.CurrentRotation = $star.CurrentRotation * $deltaRotationX * $deltaRotationY * $deltaRotationZ
            $star.Rotate.Rotation = New-Object System.Windows.Media.Media3D.QuaternionRotation3D($star.CurrentRotation)
        }
    }

    # --- Window-level Events ---
    $window.Add_KeyDown({
        param($sender, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P"      { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R"      { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H"      { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "F1"
            {
                $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
                $OkButton = $PopupWindow.FindName("OKButton")
                $OkButton.Add_Click({
                        $PopupWindow.Close()
                    })
                $PopupWindow.ShowDialog() | Out-Null
            }
        }
    })

    $mainGrid = $window.FindName("MainGrid")
    $mainGrid.Add_MouseDown({
        param($sender, $e)
        $viewport = $window.FindName('mainViewport')
        $mousePosition = $e.GetPosition($viewport)
        $hitVisual = $null

        $hitTestCallback = [System.Windows.Media.HitTestResultCallback]{
            param($result)
            if ($result -is [System.Windows.Media.Media3D.RayMeshGeometry3DHitTestResult]) {
                $script:hitVisual = $result.VisualHit
                return [System.Windows.Media.HitTestResultBehavior]::Stop
            }
            return [System.Windows.Media.HitTestResultBehavior]::Continue
        }

        $hitTestParams = [System.Windows.Media.PointHitTestParameters]::new($mousePosition)
        [System.Windows.Media.VisualTreeHelper]::HitTest($viewport, $null, $hitTestCallback, $hitTestParams)

        if ($hitVisual) { # If we hit any part of a star
            & $pauseOrResumeAnimations
        }
    })

    $window.Add_Closed({
        # Stop the animation handler
        [System.Windows.Media.CompositionTarget]::remove_Rendering($animationHandler)

        # Clean up media resources
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

    $window.Add_Loaded({
        # Start the animation loop once the window is loaded
        # Reset the timer here to prevent a large initial time jump from app startup.
        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        [System.Windows.Media.CompositionTarget]::add_Rendering($animationHandler)
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