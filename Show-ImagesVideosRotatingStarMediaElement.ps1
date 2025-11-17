<#
.SYNOPSIS
    Displays selected images and videos on a rotating 3D star-shaped object using MediaElement.
.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them one at a time
    on a rotating 3D star-shaped object in a WPF window.

    This version uses the built-in Windows MediaElement for video playback. As a result, video
    format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI).

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosRotatingStarMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D star window.
.NOTES
    Name:           Show-ImagesVideosRotatingStarMediaElement.ps1
    Version:        1.0.0, 11/05/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing


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
    $SelectForm.Text = "Star - Media Selector"
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
                $DataGridView.Rows.Add($false, [System.IO.Path]::GetFileName($file.FullName), $file.FullName)
            }

            foreach ($row in $DataGridView.Rows) {
                if ($row.IsNewRow) { continue }
                $row.HeaderCell.Value = "Play"
            }
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

    $PlayButton.Add_Click({
        $formState.SelectedFiles = @(
            $selectedRows = $DataGridView.Rows | Where-Object { $_.Cells["Select"].Value -eq $true }
            foreach ($row in $selectedRows) {
                # The value is the full path string we added
                $row.Cells["FilePath"].Value
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
        PlayerState = [hashtable]::Synchronized(@{})
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

            <!-- All 3D content must be inside a single root ModelVisual3D -->
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <!-- Lights -->
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="White" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <!-- This container will hold the rotating star object -->
                <ModelVisual3D.Children>
                    <ModelVisual3D x:Name="sceneContainer">
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
                        <!-- The star's components will be added here via code -->
                    </ModelVisual3D>
                </ModelVisual3D.Children>
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
    $sceneContainer = $window.FindName("sceneContainer")
    $SyncHash.Window = $window # Store window for event handlers

    # --- Find UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton") # Corrected from SphereWindow
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
    # Define the total height of the composite object to be a percentage of the visible viewport height.
    $totalObjectHeight = $visibleHeightAtOrigin * 0.75 # Reduced from 0.90 to make it smaller

    # The total height is (coneHeight + sphereRadius) * 2. The cone height is 1.5 * sphereRadius.
    # So, totalHeight = (1.5 * sphereRadius + sphereRadius) * 2 = 2.5 * sphereRadius * 2 = 5 * sphereRadius.
    $sphereRadius = ($totalObjectHeight / 5.0) * 0.5 # Shrink sphere by 50%
    $coneHeight = ($sphereRadius * 1.5) * 2.0 # Double the cone height relative to the sphere
    $coneRadius = ($sphereRadius * 0.4) * 2.0 # Double the cone radius relative to the sphere

    # 1. Central Sphere
    $sphereMesh = New-SphereMesh -radius $sphereRadius -slices 128 -stacks 64

    # 2. Cone Meshes (we use the same mesh and transform it for each point of the star)
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

    # --- Create separate 2D content hosts, one for each object ---
    function New-MediaHost { param([string]$TargetName)
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Background = [System.Windows.Media.Brushes]::White
        
        $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
            Name = "mediaElement_$TargetName"; Stretch = "Fill"; LoadedBehavior = "Manual"; UnloadedBehavior = "Stop"
        }
        $grid.Children.Add($mediaElement)

        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $grid.Children.Add($overlayTextBlock)

        $errorTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            Name = "errorTextBlock_$TargetName"; Foreground = [System.Windows.Media.Brushes]::Red; VerticalAlignment = 'Center'; HorizontalAlignment = 'Center'; Visibility = 'Collapsed'; TextWrapping = 'Wrap'
        }
        $grid.Children.Add($errorTextBlock)

        return @{ Grid = $grid; MediaElement = $mediaElement; OverlayTextBlock = $overlayTextBlock; ErrorTextBlock = $errorTextBlock }
    }

    $targets = @("Top", "Middle", "Bottom", "Left", "Right", "Front", "Back")
    foreach ($target in $targets) {
        $hosts = New-MediaHost -TargetName $target
        $SyncHash."${target}MediaElement" = $hosts.MediaElement
        $SyncHash."${target}OverlayTextBlock" = $hosts.OverlayTextBlock
        $SyncHash."${target}ErrorTextBlock" = $hosts.ErrorTextBlock

        # Assign the grid to the visual property of the corresponding 3D object
        $visual3D = Get-Variable -Name "${target}ConeVisual" -ValueOnly -ErrorAction SilentlyContinue
        if (-not $visual3D) { $visual3D = Get-Variable -Name "${target}SphereVisual" -ValueOnly } # For the middle sphere
        $visual3D.Visual = $hosts.Grid

        # Initialize player state
        $SyncHash.PlayerState[$hosts.MediaElement.Name] = [hashtable]::Synchronized(@{
            Target = $target
            IsImage = $false
            ImageTimer = $null
            RecoveryTimer = $null
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
            IsFailed = $false
        })
    }

    $topHosts = New-MediaHost -TargetName "Top"
    $middleHosts = New-MediaHost -TargetName "Middle"
    $bottomHosts = New-MediaHost -TargetName "Bottom"
    $leftHosts = New-MediaHost -TargetName "Left"
    $rightHosts = New-MediaHost -TargetName "Right"
    $frontHosts = New-MediaHost -TargetName "Front"
    $backHosts = New-MediaHost -TargetName "Back"

    $topConeVisual.Visual = $topHosts.Grid
    $middleSphereVisual.Visual = $middleHosts.Grid
    $bottomConeVisual.Visual = $bottomHosts.Grid
    $leftConeVisual.Visual = $leftHosts.Grid
    $rightConeVisual.Visual = $rightHosts.Grid
    $frontConeVisual.Visual = $frontHosts.Grid
    $backConeVisual.Visual = $backHosts.Grid

    # Add all objects to the main scene so they rotate together
    $sceneContainer.Children.Add($topConeVisual)
    $sceneContainer.Children.Add($middleSphereVisual)
    $sceneContainer.Children.Add($bottomConeVisual)
    $sceneContainer.Children.Add($leftConeVisual)
    $sceneContainer.Children.Add($rightConeVisual)
    $sceneContainer.Children.Add($frontConeVisual)
    $sceneContainer.Children.Add($backConeVisual)

    # Store all media elements and text blocks in the synchash
    $SyncHash.TopMediaElement = $topHosts.MediaElement
    $SyncHash.TopOverlayTextBlock = $topHosts.OverlayTextBlock
    $SyncHash.TopErrorTextBlock = $topHosts.ErrorTextBlock
    $SyncHash.MiddleMediaElement = $middleHosts.MediaElement
    $SyncHash.MiddleOverlayTextBlock = $middleHosts.OverlayTextBlock
    $SyncHash.MiddleErrorTextBlock = $middleHosts.ErrorTextBlock
    $SyncHash.BottomMediaElement = $bottomHosts.MediaElement
    $SyncHash.BottomOverlayTextBlock = $bottomHosts.OverlayTextBlock
    $SyncHash.BottomErrorTextBlock = $bottomHosts.ErrorTextBlock
    $SyncHash.LeftMediaElement = $leftHosts.MediaElement
    $SyncHash.LeftOverlayTextBlock = $leftHosts.OverlayTextBlock
    $SyncHash.LeftErrorTextBlock = $leftHosts.ErrorTextBlock
    $SyncHash.RightMediaElement = $rightHosts.MediaElement
    $SyncHash.RightOverlayTextBlock = $rightHosts.OverlayTextBlock
    $SyncHash.RightErrorTextBlock = $rightHosts.ErrorTextBlock
    $SyncHash.FrontMediaElement = $frontHosts.MediaElement
    $SyncHash.FrontOverlayTextBlock = $frontHosts.OverlayTextBlock
    $SyncHash.FrontErrorTextBlock = $frontHosts.ErrorTextBlock
    $SyncHash.BackMediaElement = $backHosts.MediaElement
    $SyncHash.BackOverlayTextBlock = $backHosts.OverlayTextBlock
    $SyncHash.BackErrorTextBlock = $backHosts.ErrorTextBlock

    # --- Media Handling Logic (Adapted from Cube Script) ---
    $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"

    function Start-NextMedia {
        param(
            [Parameter(Mandatory=$true)]
            [string]$Target # 'Top', 'Middle', 'Bottom', 'Left', 'Right', 'Front', 'Back'
        )
        $mediaElement = $SyncHash."${Target}MediaElement"
        $playerState = $SyncHash.PlayerState[$mediaElement.Name]

        $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $SyncHash.SelectedFiles.Count
        $fileObject = $SyncHash.SelectedFiles[$SyncHash.CurrentIndex]
        $uri = [System.Uri]$fileObject # This now correctly uses the string path

        # Update text overlay if set to "Filename"
        if ($SyncHash.RbSelection -eq "Filename") {
            $SyncHash."${Target}OverlayTextBlock".Text = [System.IO.Path]::GetFileName($fileObject)
        }

        $extension = [System.IO.Path]::GetExtension($fileObject).ToLower()
        $playerState.IsImage = ($ImageExtensions -contains $extension)
        $mediaElement.Source = $uri
        $mediaElement.Play()
    }

    $HandleMediaFailure = {
        param($ErrorElement, [string]$Reason = "Unknown Error")
        $SyncHash.Window.Dispatcher.Invoke([action]{
            $playerState = $SyncHash.PlayerState[$ErrorElement.Name]
            if ($playerState.IsFailed) { return }
            $playerState.IsFailed = $true

            $fileName = if ($ErrorElement.Source) { [System.IO.Path]::GetFileName($ErrorElement.Source.LocalPath) } else { "an unknown media file" }
            $errorText = "Error: $($fileName)`n$Reason"
            $SyncHash."$($playerState.Target)ErrorTextBlock".Text = $errorText
            $SyncHash."$($playerState.Target)ErrorTextBlock".Visibility = "Visible"
            $ErrorElement.Visibility = "Collapsed"
            $ErrorElement.Stop()

            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
            $recoveryTimer.Interval = [TimeSpan]::FromSeconds(10)
            $recoveryTimer.Tag = $ErrorElement
            $recoveryTick = {
                $timer = $args[0]; $failedElement = $timer.Tag; $timer.Stop()
                $SyncHash.PlayerState[$failedElement.Name].IsFailed = $false
                & $MediaEndedHandler -Sender $failedElement -e $null -IsRecovery
            }
            $recoveryTimer.Add_Tick($recoveryTick)
            $playerState.RecoveryTimer = $recoveryTimer
            $recoveryTimer.Start()
        })
    }

    $MediaFailedHandler = {
        param($Sender, $EventArgs)
        $reason = if ($EventArgs.ErrorException) { $EventArgs.ErrorException.Message } else { "MediaFailed event fired." }
        & $HandleMediaFailure -ErrorElement $Sender -Reason $reason
    }

    $MediaOpenedHandler = {
        param($Sender, $EventArgs)
        $playerState = $SyncHash.PlayerState[$Sender.Name]
        $playerState.IsFailed = $false
        $playerState.PlaybackStopwatch.Restart()

        $SyncHash.Window.Dispatcher.Invoke([action]{
            $Sender.Visibility = "Visible"
            $SyncHash."$($playerState.Target)ErrorTextBlock".Visibility = "Collapsed"
        })

        if ($playerState.IsImage) {
            $Sender.Pause()
            if ($SyncHash.SelectedFiles.Count -gt $targets.Count) {
                if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
                $playerState.ImageTimer = New-Object System.Windows.Threading.DispatcherTimer
                $playerState.ImageTimer.Interval = [TimeSpan]::FromSeconds(10)
                $playerState.ImageTimer.Tag = $Sender
                $tickScriptBlock = {
                    $timer = $args[0]; $mediaElement = $timer.Tag; $timer.Stop()
                    & $MediaEndedHandler -Sender $mediaElement -e $null
                }
                $playerState.ImageTimer.Add_Tick($tickScriptBlock)
                $playerState.ImageTimer.Start()
            }
        } elseif (-not $Sender.NaturalDuration.HasTimeSpan) {
            & $HandleMediaFailure -ErrorElement $Sender -Reason "No duration found (bad codec/file)."
        } else {
            if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
        }
    }

    $MediaEndedHandler = {
        param($Sender, $e, [switch]$IsRecovery)
        if (-not $Sender -or -not $Sender.Name) { return }
        $playerState = $SyncHash.PlayerState[$Sender.Name]
        if ($playerState.IsFailed) { return }

        if (-not $IsRecovery) {
            $playerState.PlaybackStopwatch.Stop()
            if (($playerState.PlaybackStopwatch.Elapsed.TotalMilliseconds -lt 2000) -and (-not $playerState.IsImage)) {
                & $HandleMediaFailure -ErrorElement $Sender -Reason "Playback ended instantly."
                return
            }
        }

        if ($SyncHash.SelectedFiles.Count -le $targets.Count) {
            $Sender.Position = [TimeSpan]::FromSeconds(0)
            $Sender.Play()
        } else {
            Start-NextMedia -Target $playerState.Target
        }
    }

    foreach ($target in $targets) {
        $mediaElement = $SyncHash."${target}MediaElement"
        $mediaElement.Add_MediaFailed($MediaFailedHandler)
        $mediaElement.Add_MediaOpened($MediaOpenedHandler)
        $mediaElement.Add_MediaEnded($MediaEndedHandler)
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
        foreach ($playerName in $SyncHash.PlayerState.Keys) {
            $playerState = $SyncHash.PlayerState[$playerName]
            if ($playerState) {
                if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            }
        }
        foreach ($target in $targets) {
            $mediaElement = $SyncHash."${target}MediaElement"
            $mediaElement.Stop()
            $mediaElement.Source = $null
            $mediaElement.Close()
        }

        $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
    })

    # --- Start the show ---
    $window.Add_Loaded({
        # Start the media cycle once the window is fully loaded
        foreach ($target in $targets) { Start-NextMedia -Target $target }
    })

    # --- Apply Text Overlay Settings on Load ---
    foreach ($target in $targets) {
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