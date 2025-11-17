<#
.SYNOPSIS
    Displays selected images and videos on a rotating 3D sphere using FFmpeg.
.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them one at a time
    on a rotating 3D sphere in a WPF window.

    It uses FFmpeg to decode video frames in real-time and stream them to a WriteableBitmap,
    which is then applied as a texture to the sphere. This allows for broad video format support.

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays.
.EXAMPLE
    PS C:\> .\Show-ImageVideoSphereFfmpeg.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D sphere window.
.NOTES
    Name:           Show-ImageVideoSphereFfmpeg.ps1
    Version:        1.0.0, 10/25/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffplay.exe): https://www.ffmpeg.org/download.html
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Rotating Image/Video Sphere `n Uses Ffmpeg"
$ScriptDescription = "Loops through and displays selected images or videos on a rotating 3D sphere. Uses FFmpeg for video decoding, providing broad format support."

# --- Dependency Check ---
$RequiredExecutables = @("ffmpeg.exe", "ffplay.exe")
$dependencyStatus = @()
$allDependenciesMet = $true

foreach ($exe in $RequiredExecutables) {
    $localPath = Join-Path $PSScriptRoot $exe
    if ((Get-Command $exe -ErrorAction SilentlyContinue) -or (Test-Path -Path $localPath)) {
        $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'Found' }
    } else {
        $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'NOT FOUND' }
        $allDependenciesMet = $false
    }
}

if (-not $allDependenciesMet) {
    $messageLines = @(
        "One or more required executables were not found in your system's PATH. Please install them and try again.",
        "",
        "Required executable status:"
    )
    foreach ($status in $dependencyStatus) { $messageLines += " - $($status.Status): $($status.Name)" }
    $message = $messageLines -join "`n"
    [System.Windows.Forms.MessageBox]::Show($message, "Dependency Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return # Exit the script
}

# --- Sphere Generation Function ---
function New-SphereMesh {
    param(
        [double]$radius = 1.5,
        [int]$slices = 64, # Longitude
        [int]$stacks = 32  # Latitude
    )

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

    # Add vertices and texture coordinates
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

    # Add triangle indices
    for ($stack = 0; $stack -lt $stacks; $stack++) {
        for ($slice = 0; $slice -lt $slices; $slice++) {
            $i0 = $stack * ($slices + 1) + $slice; $i1 = ($stack + 1) * ($slices + 1) + $slice
            $mesh.TriangleIndices.Add($i0); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i0 + 1)
            $mesh.TriangleIndices.Add($i0 + 1); $mesh.TriangleIndices.Add($i1); $mesh.TriangleIndices.Add($i1 + 1)
        }
    }
    return $mesh
}

# --- Cone Generation Function ---
function New-ConeMesh {
    param(
        [double]$radius = 1.5,
        [double]$height = 3.0,
        [int]$slices = 64
    )

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

    # Apex vertex
    $apexY = $height # Apex is now at the top
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $apexY, 0))
    $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0)) # Apex texture coord

    # Base vertices
    $baseY = 0 # Base is at the origin (Y=0)
    for ($i = 0; $i -le $slices; $i++) {
        $theta = $i * 2 * [Math]::PI / $slices
        $x = $radius * [Math]::Cos($theta)
        $z = $radius * [Math]::Sin($theta)
        $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $baseY, $z))
        $mesh.TextureCoordinates.Add([System.Windows.Point]::new($i / $slices, 1))
    }

    # Base center vertex for the bottom cap
    $baseCenterIndex = $mesh.Positions.Count
    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new(0, $baseY, 0))
    $mesh.TextureCoordinates.Add([System.Windows.Point]::new(0.5, 0.5))

    # Add triangle indices for the cone sides
    for ($i = 0; $i -lt $slices; $i++) {
        $mesh.TriangleIndices.Add(0)               # Apex
        $mesh.TriangleIndices.Add($i + 2)
        $mesh.TriangleIndices.Add($i + 1)
    }

    # Add triangle indices for the base cap
    for ($i = 0; $i -lt $slices; $i++) {
        $mesh.TriangleIndices.Add($baseCenterIndex)
        $mesh.TriangleIndices.Add($i + 1)
        $mesh.TriangleIndices.Add($i + 2)
    }

    return $mesh
}

# --- Main Application Loop ---
while ($true) {
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Sphere - Media Selector"
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

    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false
    }
    $SelectForm.Controls.Add($SelectAllCheckbox)

    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false
    }
    $SelectForm.Controls.Add($TransparentCheckbox)

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
        Name = "FileName"; HeaderText = "File Name"; Width = 200; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null

    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
        Name = "FilePath"; HeaderText = "File Path"; Width = 330; ReadOnly = $true
    }
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Play Selected"; Location = '600, 40'; Size = '170, 30'
    }
    $SelectForm.Controls.Add($PlayButton)

    # --- Text Overlay Controls (copied from cube script) ---
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
    # --- Script-scoped form state with safe defaults ---
    $formState = @{
      TextColor            = [System.Drawing.Color]::Black
      FontFamily           = "Arial"
      FontSize             = 24
      IsBold               = $true
      IsItalic             = $false
    }

    # Initialize the textbox font to match the default state
    $initialFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$formState.FontSize, [System.Drawing.FontStyle]::Bold)
    $TextBox.Font = $initialFont

    # Initialize the textbox font to match the default state
    $initialFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$formState.FontSize, [System.Drawing.FontStyle]::Bold)
    $TextBox.Font = $initialFont

    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $formState.TextColor = $colorDialog.Color
            $ColorExample.BackColor = $formState.TextColor
            $TextBox.ForeColor = $formState.TextColor
        }
    })

    $FontButton.Add_Click({
        $fontDialog = New-Object System.Windows.Forms.FontDialog
        $currentFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value)
        $fontDialog.Font = $currentFont

        if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $selectedFont = $fontDialog.Font
            $formState.FontFamily = $selectedFont.Name
            $FontButton.Text = $formState.FontFamily
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
            $newFont = New-Object System.Drawing.Font($formState.FontFamily, [float]$NumericUpDown.Value, $style)
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
        foreach ($row in $DataGridView.Rows) {
            $row.Cells["Select"].Value = $isChecked
        }
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
                File    = $true
                Include = $AllowedExtensions
            }
            if ($RecursiveCheckBox.Checked) {
                $gciParams.Path = $SelectedPath
                $gciParams.Recurse = $true
            } else {
                $gciParams.Path = Join-Path $SelectedPath "*"
            }
            $files = Get-ChildItem @gciParams
            foreach ($file in $files) {
                $DataGridView.Rows.Add($false, [System.IO.Path]::GetFileName($file), $file)
            }

            foreach ($row in $DataGridView.Rows) {
                if ($row.IsNewRow) { continue }
                $row.HeaderCell.Value = "Play"
            }
        }
    })

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.RowIndex -lt $DataGridView.Rows.Count) {
            $row = $DataGridView.Rows[$e.RowIndex]
            if ($row.IsNewRow) { return }
            $filePath = $row.Cells["FilePath"].Value
            if ([System.IO.File]::Exists($filePath)) {
                Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -i `"$filePath`"" -NoNewWindow
            }
        }
    })


    $PlayButton.Add_Click({
        $formState.SelectedFiles = @(
            foreach ($Row in $DataGridView.Rows) {
                if ($Row.Cells["Select"].Value) {
                    $Row.Cells["FilePath"].Value
                }
            }
        )
        if ($formState.SelectedFiles.Count -gt 0) {
            $formState.UseTransparentEffect = $TransparentCheckbox.Checked
            
            # Capture text settings
            if ($RadioButton1.Checked) { $formState.RbSelection = "Hidden" }
            if ($RadioButton2.Checked) { $formState.RbSelection = "Filename" }
            if ($RadioButton3.Checked) { $formState.RbSelection = "Custom" }
            $formState.CustomText = $TextBox.Text

            try {
                $formState.FontSize = [double]$NumericUpDown.Value
                if ($formState.FontSize -le 0) { $formState.FontSize = 24 }
            } catch { $formState.FontSize = 24 }
            $formState.IsBold   = $BoldCheckbox.Checked
            $formState.IsItalic = $ItalicCheckbox.Checked

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
                    <Run Text="Pause/Resume      :  P  : Pause/Resume Spinning"/><LineBreak/>
                    <Run Text="Redo              :  R  : Reselect Media"/><LineBreak/>
                    <Run Text="Random Axis       :  A  : Change Rotation Axis"/><LineBreak/>
                    <Run Text="Hide Controls     :  H  : Hide/Show Controls"/><LineBreak/>
                    <Run Text="Left Arrow        :  &#x2190;  : Slow Down Spinning"/><LineBreak/>
                    <Run Text="Right Arrow       :  &#x2192;  : Speed Up Spinning"/><LineBreak/><LineBreak/>
                    <Run Text="*Click star to Pause/Resume*"/><LineBreak/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0" IsDefault="True"/>
    </Grid>
</Window>
"@

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

    # Exit if no files were selected or the form was closed
    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0) {
        Write-Host "No files were selected or form was closed. Exiting."
        break # Exit the main while loop
    }

    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
        SelectedFiles = $formState.SelectedFiles
        CurrentIndex = -1 # Global index, will be incremented before first use.
        MediaTimers = @{}; FfmpegProcesses = @{} # Hashtables to store resources by target
        Paused = $false
        ControlsHidden = $false
        RedoClicked = $false
        UseTransparentEffect = $formState.UseTransparentEffect
        # Text Overlay Settings
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
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="SphereWindow"
        Title="Rotating Sphere"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,8" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>

            <!-- This ModelVisual3D will contain our lights and the sphere model -->
            <ModelVisual3D x:Name="SphereContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <!-- Lights -->
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" />
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

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # --- Set Window to Full Screen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top


    # --- Find the container for our sphere ---
    $sphereContainer = $window.FindName("SphereContainer")
    $SyncHash.Window = $window # Store window for event handlers

    # --- Find UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")
    $window.Title = "Rotating Cone"

    # --- Create the Cone using Viewport2DVisual3D for robust media display ---
    # Calculate a dynamic size to make the cone's height a percentage of the viewport height.
    $cameraDistance = 8.0
    $cameraFovDegrees = 60.0
    $cameraFovRadians = $cameraFovDegrees * ([Math]::PI / 180.0)
    $visibleHeightAtOrigin = 2.0 * $cameraDistance * [Math]::Tan($cameraFovRadians / 2.0)

    # --- Define the geometry for all three objects ---
    # Define the total height of the composite object to be 90% of the visible viewport height.
    $totalObjectHeight = $visibleHeightAtOrigin * 0.75 # Reduced from 0.90 to make it smaller

    # The total height is (coneHeight + sphereRadius) * 2. The cone height is 1.5 * sphereRadius.
    # So, totalHeight = (1.5 * sphereRadius + sphereRadius) * 2 = 2.5 * sphereRadius * 2 = 5 * sphereRadius.
    $sphereRadius = ($totalObjectHeight / 5.0) * 0.5 # Shrink sphere by 50%
    $coneHeight = ($sphereRadius * 1.5) * 2.0 # Double the cone height relative to the sphere
    $coneRadius = ($sphereRadius * 0.4) * 2.0 # Double the cone radius relative to the sphere

    # 1. Central Sphere
    $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64

    # 2. Cone Meshes (we will use the same mesh and transform it)
    $coneMesh = New-ConeMesh -radius $coneRadius -height $coneHeight -slices 128

    # --- Create Visuals and Materials for all three objects ---
    $topConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }
    $middleSphereVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $sphereMesh }
    $bottomConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }
    $leftConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }
    $rightConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }
    $frontConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }
    $backConeVisual = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D -Property @{ Geometry = $coneMesh }

    # --- Position the cones using Transforms (This is the key fix for the rotation issue) ---
    # Top Cone: Translate it up by the sphere's radius
    $topTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)
    $topConeVisual.Transform = $topTransform

    # Bottom Cone: Invert it on Y-axis, then translate it down
    $bottomTransformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
    $bottomTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.ScaleTransform3D(1, -1, 1)))
    $bottomTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, -$sphereRadius, 0)))
    $bottomConeVisual.Transform = $bottomTransformGroup

    # Right Cone: Translate it up, then rotate it 90 degrees around Z-axis to point right
    $rightTransformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
    $rightTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
    $rightTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), -90)))))
    $rightConeVisual.Transform = $rightTransformGroup

    # Left Cone: Translate it up, then rotate it -90 degrees around Z-axis to point left
    $leftTransformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
    $leftTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
    $leftTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,0,1), 90)))))
    $leftConeVisual.Transform = $leftTransformGroup

    # Front Cone: Translate it up, then rotate it -90 degrees around X-axis to point forward
    $frontTransformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
    $frontTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
    $frontTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), -90)))))
    $frontConeVisual.Transform = $frontTransformGroup

    # Back Cone: Translate it up, then rotate it 90 degrees around X-axis to point backward
    $backTransformGroup = New-Object System.Windows.Media.Media3D.Transform3DGroup
    $backTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, $sphereRadius, 0)))
    $backTransformGroup.Children.Add((New-Object System.Windows.Media.Media3D.RotateTransform3D((New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(1,0,0), 90)))))
    $backConeVisual.Transform = $backTransformGroup

    # Create the material and link it to the visual host
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
    $sphereMaterial = New-Object $materialType
    [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($sphereMaterial, $true)
    $topConeVisual.Material = $sphereMaterial
    $middleSphereVisual.Material = $sphereMaterial.Clone()
    $bottomConeVisual.Material = $sphereMaterial.Clone()
    $leftConeVisual.Material = $sphereMaterial.Clone()
    $rightConeVisual.Material = $sphereMaterial.Clone()
    $frontConeVisual.Material = $sphereMaterial.Clone()
    $backConeVisual.Material = $sphereMaterial.Clone()

    # --- Create THREE separate 2D content hosts, one for each object ---
    function New-MediaHost {
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Background = [System.Windows.Media.Brushes]::White
        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $grid.Children.Add($contentPresenter)
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $grid.Children.Add($overlayTextBlock)
        return @{ Grid = $grid; ContentPresenter = $contentPresenter; OverlayTextBlock = $overlayTextBlock }
    }

    $topHosts = New-MediaHost
    $middleHosts = New-MediaHost
    $bottomHosts = New-MediaHost
    $leftHosts = New-MediaHost
    $rightHosts = New-MediaHost
    $frontHosts = New-MediaHost
    $backHosts = New-MediaHost

    $topConeVisual.Visual = $topHosts.Grid
    $middleSphereVisual.Visual = $middleHosts.Grid
    $bottomConeVisual.Visual = $bottomHosts.Grid
    $leftConeVisual.Visual = $leftHosts.Grid
    $rightConeVisual.Visual = $rightHosts.Grid
    $frontConeVisual.Visual = $frontHosts.Grid
    $backConeVisual.Visual = $backHosts.Grid

    # Add all three objects to the main scene so they rotate together
    $sphereContainer.Children.Add($topConeVisual)
    $sphereContainer.Children.Add($middleSphereVisual)
    $sphereContainer.Children.Add($bottomConeVisual)
    $sphereContainer.Children.Add($leftConeVisual)
    $sphereContainer.Children.Add($rightConeVisual)
    $sphereContainer.Children.Add($frontConeVisual)
    $sphereContainer.Children.Add($backConeVisual)

    # --- Media Handling Functions ---
    # Store ALL content presenters and text blocks in the synchash
    $SyncHash.TopContentPresenter = $topHosts.ContentPresenter
    $SyncHash.TopOverlayTextBlock = $topHosts.OverlayTextBlock
    $SyncHash.MiddleContentPresenter = $middleHosts.ContentPresenter
    $SyncHash.MiddleOverlayTextBlock = $middleHosts.OverlayTextBlock
    $SyncHash.BottomContentPresenter = $bottomHosts.ContentPresenter
    $SyncHash.BottomOverlayTextBlock = $bottomHosts.OverlayTextBlock
    $SyncHash.LeftContentPresenter = $leftHosts.ContentPresenter
    $SyncHash.LeftOverlayTextBlock = $leftHosts.OverlayTextBlock
    $SyncHash.RightContentPresenter = $rightHosts.ContentPresenter
    $SyncHash.RightOverlayTextBlock = $rightHosts.OverlayTextBlock
    $SyncHash.FrontContentPresenter = $frontHosts.ContentPresenter
    $SyncHash.FrontOverlayTextBlock = $frontHosts.OverlayTextBlock
    $SyncHash.BackContentPresenter = $backHosts.ContentPresenter
    $SyncHash.BackOverlayTextBlock = $backHosts.OverlayTextBlock


    function Start-NextMedia {
        param(
            [Parameter(Mandatory=$true)]
            [string]$Target # 'Top', 'Middle', 'Bottom', 'Left', 'Right', 'Front', 'Back'
        )

        # Clean up previous media resources
        if ($SyncHash.MediaTimers[$Target]) { $SyncHash.MediaTimers[$Target].Stop() }
        if ($SyncHash.FfmpegProcesses[$Target] -and -not $SyncHash.FfmpegProcesses[$Target].HasExited) {
            $SyncHash.FfmpegProcesses[$Target].Kill()
            $SyncHash.FfmpegProcesses[$Target].Dispose()
            $SyncHash.FfmpegProcesses[$Target] = $null
        }

        # Get next media file using the global index, looping if necessary
        $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $SyncHash.SelectedFiles.Count
        $filePath = $SyncHash.SelectedFiles[$SyncHash.CurrentIndex]

        # Update text overlay if set to "Filename"
        if ($SyncHash.RbSelection -eq "Filename") {
            $SyncHash."${Target}OverlayTextBlock".Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

        if ($ImageExtensions -contains $extension) {
            Start-ImageOnSphere -FilePath $filePath -Target $Target
        } else {
            Start-VideoOnSphere -FilePath $filePath -Target $Target
        }
    }

    function Start-ImageOnSphere {
        param([string]$FilePath, [string]$Target)
        
        $image = New-Object System.Windows.Controls.Image
        $image.Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$FilePath)
        $image.Stretch = "Fill"
        $SyncHash."${Target}ContentPresenter".Content = $image

        # Set a timer to show the next media item after 10 seconds
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds(10)
        $tickScriptBlock = {
            $timer.Stop()
            Start-NextMedia -Target $Target
        }
        $timer.Add_Tick($tickScriptBlock.GetNewClosure())
        $timer.Start()
        $SyncHash.MediaTimers[$Target] = $timer
    }

    function Start-VideoOnSphere {
        param([string]$FilePath, [string]$Target)

        $width = 1280
        $height = 720
        $frameSize = $width * $height * 3 # Bgr24 is 3 bytes per pixel

        $bitmap = [System.Windows.Media.Imaging.WriteableBitmap]::new($width, $height, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null)
        $imageControl = New-Object System.Windows.Controls.Image
        $imageControl.Source = $bitmap
        $imageControl.Stretch = "Fill"
        $SyncHash."${Target}ContentPresenter".Content = $imageControl

        # Start ffmpeg to stream frames
        $args = "-hide_banner -loglevel error -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"
        $psi = New-Object System.Diagnostics.ProcessStartInfo -Property @{
            FileName = "ffmpeg.exe"; Arguments = $args; RedirectStandardOutput = $true
            UseShellExecute = $false; CreateNoWindow = $true
        }
        $proc = [System.Diagnostics.Process]::Start($psi)
        $SyncHash.FfmpegProcesses[$Target] = $proc

        $stream = $proc.StandardOutput.BaseStream
        $bytes = New-Object byte[] $frameSize
        $rect = [System.Windows.Int32Rect]::new(0, 0, $width, $height)
        $stride = $width * 3

        # Set a timer to read frames from ffmpeg
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(33) # ~30fps

        $tickScriptBlock = {
            try {
                $totalRead = 0
                while ($totalRead -lt $frameSize) {
                    $read = $stream.Read($bytes, $totalRead, $frameSize - $totalRead)
                    if ($read -le 0) { # End of stream
                        Start-NextMedia -Target $Target
                        return # Exit this tick
                    }
                    $totalRead += $read
                }

                if ($totalRead -eq $frameSize) {
                    $bitmap.Lock()
                    $bitmap.WritePixels($rect, $bytes, $stride, 0)
                    $bitmap.Unlock()
                }
            } catch {
                # An error likely means the process was killed or stream closed.
                # Stop the timer to prevent further errors.
                $timer.Stop()
            }
        }
        $timer.Add_Tick($tickScriptBlock.GetNewClosure())
        $timer.Start()
        $SyncHash.MediaTimers[$Target] = $timer
    }
    # --- Animate the Cone ---
    $animX = New-Object System.Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(20))
    $animX.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleX = $window.FindName("AxisAngleX")
    $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

    $animY = New-Object System.Windows.Media.Animation.DoubleAnimation(360, 0, [TimeSpan]::FromSeconds(15))
    $animY.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
    $axisAngleY = $window.FindName("AxisAngleY")
    $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)
    $SyncHash.animX = $animX
    $SyncHash.animY = $animY
    $SyncHash.AxisAngleX = $axisAngleX
    $SyncHash.AxisAngleY = $axisAngleY

    # --- UI Event Handlers ---
    $SyncHash.closeButton.Add_Click({ $window.Close() })

    # --- Handle Window Events ---
    $window.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq 'Escape') {
            $window.Close()
        }
        switch ($e.Key) {
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "A" { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
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
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Pause
            Start-Sleep -Milliseconds 50 # Give it a moment to process
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Resume
        }
    }

    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })

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

    $SyncHash.redoButton.Add_Click({
        $SyncHash.RedoClicked = $true
        $SyncHash.Window.Close()
    })

    $mainGrid = $window.FindName("MainGrid")
    $mainGrid.Add_MouseDown({
        param($sender, $e)
        $viewport = $window.FindName('mainViewport')
        $mousePosition = $e.GetPosition($viewport)
        $SyncHash.hitVisual = $null # Use hitVisual for Viewport2DVisual3D

        # Define the callback for the hit test
        $hitTestCallback = [System.Windows.Media.HitTestResultCallback]{
            param($result)
            if ($result -is [System.Windows.Media.Media3D.RayMeshGeometry3DHitTestResult]) {
                $SyncHash.hitVisual = $result.VisualHit
                return [System.Windows.Media.HitTestResultBehavior]::Stop
            }
            return [System.Windows.Media.HitTestResultBehavior]::Continue
        }

        # Perform the hit test
        $hitTestParams = [System.Windows.Media.PointHitTestParameters]::new($mousePosition)
        [System.Windows.Media.VisualTreeHelper]::HitTest($viewport, $null, $hitTestCallback, $hitTestParams)

        # If the hit visual is our sphere, trigger the pause button
        if ($SyncHash.hitVisual -is [System.Windows.Media.Media3D.Viewport2DVisual3D]) {
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
        }
    })

    $window.Add_Closed({
        # Stop animations to prevent resource leaks
        foreach ($timer in $SyncHash.MediaTimers.Values) { $timer.Stop() }
        foreach ($proc in $SyncHash.FfmpegProcesses.Values) {
            if ($proc -and -not $proc.HasExited) {
                try {
                    $proc.Kill()
                    $proc.Dispose()
                }
                catch { }
            }
        }

        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
    })

    # --- Start the show ---
    $window.Add_Loaded({
        # Start the media cycle once the window is fully loaded
        Start-NextMedia -Target "Top"
        Start-NextMedia -Target "Middle"
        Start-NextMedia -Target "Bottom"
        Start-NextMedia -Target "Left"
        Start-NextMedia -Target "Right"
        Start-NextMedia -Target "Front"
        Start-NextMedia -Target "Back"
    })

    # --- Apply Text Overlay Settings on Load ---
    foreach ($target in @("Top", "Middle", "Bottom", "Left", "Right", "Front", "Back")) {
        $textBlock = $SyncHash."${target}OverlayTextBlock"
        switch ($SyncHash.RbSelection) {
            "Hidden" {
                $textBlock.Visibility = 'Collapsed'
            }
            "Filename" {
                # Text is set in Start-NextMedia
            }
            "Custom" {
                $textBlock.Text = $SyncHash.CustomText
            }
        }

        if ($SyncHash.RbSelection -ne "Hidden") {
            $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
            $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
            $textBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
            $textBlock.FontSize = $SyncHash.FontSize
            if ($SyncHash.IsBold) { $textBlock.FontWeight = 'Bold' }
            if ($SyncHash.IsItalic) { $textBlock.FontStyle = 'Italic' }
        }
    }

    $null = $window.ShowDialog()

    # After window closes, check if we need to loop
    if (-not $SyncHash.RedoClicked) {
        break # Exit the main while loop
    }
}