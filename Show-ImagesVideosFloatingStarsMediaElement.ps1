<#
.SYNOPSIS
    Displays selected images and videos on the surfaces of six floating, bouncing 3D stars using MediaElement.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto the surfaces
    of six independently moving 3D stars in a WPF window. The stars move in straight lines and
    bounce off the edges of the screen, creating a "zero-gravity" or "screen saver" visual effect.

    A "star" is a composite 3D object made of a central sphere and six cones. All parts of a single
    star display the same media content.

    This version uses the built-in Windows MediaElement for video playback. As a result, video
    format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI).

    The 3D view is interactive, with controls to pause the animation and hide the UI for an
    unobstructed view.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosFloatingStarsMediaElement.ps1

    Launches the file selection GUI. After selecting at least one file and clicking "Play", the
    script will launch the 3D window with six floating, bouncing stars.

.NOTES
    Name:           Show-ImagesVideosFloatingStarsMediaElement.ps1
    Version:        1.0.0, 11/15/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Floating Stars `n(MediaElement)"
$ScriptDescription = "Displays media on six independently floating and bouncing 3D stars. Uses the built-in Windows MediaElement."
$RequiredExecutables = @() # No external executables needed

# --- Main Application Loop ---
while ($true) {
    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
        WindowReady = $false # Flag for the loading form
        SelectedFiles = @()
        GlobalCounter = -1
        PlayerState = [hashtable]::Synchronized(@{})
        MediaReadyCounter = 0 # Counter for initial media loads
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
    $SelectForm.Text = "Floating Stars (MediaElement) - Media Selector"
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

    [xml]$VideoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preview Video - Click to Pause/Resume" Height="450" Width="800"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize" SizeToContent="Manual"
        WindowState="Normal" WindowStyle="ToolWindow" Background="Black">
    <Grid x:Name="TheGrid">
        <MediaElement x:Name="MediaPlayer" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
    </Grid>
</Window>
"@

    $DataGridView.Add_RowHeaderMouseClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $row = $DataGridView.Rows[$e.RowIndex]
        $filePath = $row.Cells["FilePath"].Value

        $isPreviewPaused = $false
        $VideoReader = (New-Object System.Xml.XmlNodeReader $VideoXaml)
        $VideoWindow = [Windows.Markup.XamlReader]::Load($VideoReader)

        $TheGrid = $VideoWindow.FindName("TheGrid")
        $MediaPlayer = $VideoWindow.FindName("MediaPlayer")
        $MediaPlayer.Source = [Uri]$filePath

        $TheGrid.Add_MouseDown({
            if($isPreviewPaused) { $MediaPlayer.Play(); $isPreviewPaused = $false }
            else { $MediaPlayer.Pause(); $isPreviewPaused = $true }
        })

        $MediaPlayer.Add_MediaEnded({ $MediaPlayer.Position = [TimeSpan]::Zero; $MediaPlayer.Play() })
        $MediaPlayer.Play()
        $VideoWindow.ShowDialog() | Out-Null
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
        Title="Floating Stars (MediaElement)"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D x:Name="mainViewport" Visibility="Collapsed">
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

    # Creates an isolated 2D host for media content.
    function New-MediaHost {
        param([string]$Name)
        $grid = New-Object System.Windows.Controls.Grid
        
        $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
            Name = "mediaElement_$Name"; Stretch = "Fill"; LoadedBehavior = "Manual"; UnloadedBehavior = "Stop"; Visibility = 'Collapsed'
        }
        $imageElement = New-Object System.Windows.Controls.Image -Property @{
            Name = "imageElement_$Name"; Stretch = "Fill"; Visibility = 'Collapsed'
        }
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false
        }
        $errorBorder = New-Object System.Windows.Controls.Border -Property @{
            Name = "errorBorder_$Name"; Background = [System.Windows.Media.Brushes]::Black; Visibility = 'Collapsed'
        }
        $errorTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            Name = "errorTextBlock_$Name"; Foreground = [System.Windows.Media.Brushes]::Red; VerticalAlignment = 'Center'; HorizontalAlignment = 'Center'; TextWrapping = 'Wrap'; Margin = '10'
        }
        $errorBorder.Child = $errorTextBlock

        $grid.Children.Add($mediaElement)
        $grid.Children.Add($imageElement)
        $grid.Children.Add($overlayTextBlock)
        $grid.Children.Add($errorBorder)

        return @{ Grid = $grid; MediaElement = $mediaElement; ImageElement = $imageElement; OverlayTextBlock = $overlayTextBlock; ErrorBorder = $errorBorder; ErrorTextBlock = $errorTextBlock }
    }

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

    # This function creates a complete, isolated 2D visual host (Grid, MediaElement, etc.)
    # and returns a hashtable containing the brush and the controls. This prevents a single failure from
    # invalidating the material for the entire star.
    function New-IsolatedVisualBrush {
        $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
            LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; ScrubbingEnabled = $true
        }
        $imageElement = New-Object System.Windows.Controls.Image -Property @{
            Stretch = 'Fill'; Visibility = 'Collapsed'
        }
        $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
            HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'; TextWrapping = 'Wrap';
            TextAlignment = 'Center'; IsHitTestVisible = $false; Margin = '10,0,10,0'
        }
        # The Grid's background is a solid black border that is always present.
        # This is the key to preventing transparency on error.
        $backgroundBorder = New-Object System.Windows.Controls.Border -Property @{ Background = [System.Windows.Media.Brushes]::Black }
        
        $visualHostGrid = New-Object System.Windows.Controls.Grid
        [void]$visualHostGrid.Children.Add($backgroundBorder)
        [void]$visualHostGrid.Children.Add($imageElement)
        [void]$visualHostGrid.Children.Add($mediaElement)
        [void]$visualHostGrid.Children.Add($overlayTextBlock)

        return @{
            Brush = (New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $visualHostGrid })
            Player = $mediaElement; ImageElement = $imageElement; Overlay = $overlayTextBlock; Background = $backgroundBorder
        }
    }

    function New-StarVisual3D {
        param(
            [int]$StarIndex,
            [hashtable]$SyncHash
        )
        
        $starModelGroup = New-Object System.Windows.Media.Media3D.Model3DGroup

        $starContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D

        $sphereRadius = 0.5
        $coneHeight = 0.75
        $coneRadius = 0.2

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

        # Create a visible ModelVisual3D for each part of the star.
        # Each part gets its own material AND its own VisualBrush to ensure isolation.
        $playerComponents = @()
        foreach ($key in $visuals.Keys) {
            $part = $visuals[$key]
            $geometryModel = New-Object System.Windows.Media.Media3D.GeometryModel3D
            $geometryModel.Geometry = $part.Mesh
            $geometryModel.Transform = $part.Transform

            $isolatedBrushData = New-IsolatedVisualBrush
            $playerComponents += $isolatedBrushData # Store for linking events

            if ($SyncHash.UseTransparentEffect) {
                # For EmissiveMaterial, the Color must be White for the Brush to show correctly.
                $geometryModel.Material = New-Object System.Windows.Media.Media3D.EmissiveMaterial -Property @{ Brush = $isolatedBrushData.Brush; Color = [System.Windows.Media.Colors]::White }
            } else {
                $geometryModel.Material = New-Object System.Windows.Media.Media3D.DiffuseMaterial -Property @{ Brush = $isolatedBrushData.Brush }
            }

            $modelVisual = New-Object System.Windows.Media.Media3D.ModelVisual3D
            $modelVisual.Content = $geometryModel
            $starContainer.Children.Add($modelVisual)
        }

        $SyncHash.PlayerState[$StarIndex] = @{
            PlayerComponents = $playerComponents # Store the list of 7 component sets
            IsImage = $false
            ImageTimer = $null
            IsFailed = $false
            PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
            RecoveryTimer = $null
            InitialLoadComplete = $false
        }
        return $starContainer
    }

    function Check-LoadingComplete {
        param([hashtable]$SyncHash)
    
        [System.Threading.Monitor]::Enter($SyncHash) # Lock for thread-safe increment and check
        try {
            $SyncHash.MediaReadyCounter++
            if ($SyncHash.MediaReadyCounter -ge 6) { # Hardcoded to 6 stars
                # All media loaded. Use dispatcher to update UI and signal loading form.
                $SyncHash.Window.Dispatcher.Invoke([action]{
                    $mainViewport = $SyncHash.Window.FindName("mainViewport")
                    $mainViewport.Visibility = 'Visible'
                    $SyncHash.WindowReady = $true # Signal loading form to close
                })
            }
        }
        finally { [System.Threading.Monitor]::Exit($SyncHash) }
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

        # Determine media type BEFORE interacting with the MediaElement to avoid race conditions.
        $nextIndex = Get-NextMediaIndex -SyncHash $SyncHash
        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $uri = [Uri]$filePath
        $ext = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
        $playerState.IsImage = ($SyncHash.ImageExtensions -contains $ext)

        # All UI interactions must be on the UI thread
        $SyncHash.Window.Dispatcher.Invoke([action]{
            # Apply the same source and text to all 7 components of the star
            foreach ($component in $playerState.PlayerComponents) {
                $component.Player.Stop()
                $component.Player.Source = $uri
                # Always call Play(). The MediaOpened event is responsible for pausing if it's an image.
                $component.Player.Play()
                $overlay = $component.Overlay
                switch ($SyncHash.RbSelection) {
                    "Hidden"   { $overlay.Visibility = 'Collapsed' }
                    "Filename" { $overlay.Text = [System.IO.Path]::GetFileName($uri.LocalPath); $overlay.Visibility = 'Visible' }
                    "Custom"   { $overlay.Text = $SyncHash.CustomText; $overlay.Visibility = 'Visible' }
                }
            }
        })
    }

    function Handle-MediaFailure {
        param([int]$StarIndex, [string]$Reason, [hashtable]$SyncHash)
        $playerState = $SyncHash.PlayerState[$StarIndex]
        if (-not $playerState -or $playerState.IsFailed) { return }
        $playerState.IsFailed = $true

        $failedUri = $playerState.PlayerComponents[0].Player.Source
        $fileName = if ($failedUri) { [System.IO.Path]::GetFileName($failedUri.LocalPath) } else { "an unknown file" }
        Write-Warning "Media failed for Star $StarIndex (File: '$fileName'). Reason: $Reason. Removing from list and continuing."

        # Remove the bad file from the list in a thread-safe way
        if ($failedUri) {
            $SyncHash.SelectedFiles = @($SyncHash.SelectedFiles | Where-Object { $_ -ne $failedUri.LocalPath })
        }

        $SyncHash.Window.Dispatcher.Invoke([action]{
            if (-not $playerState.InitialLoadComplete) {
                $playerState.InitialLoadComplete = $true
                Check-LoadingComplete -SyncHash $SyncHash
            }
        })

        # Immediately try to load the next file.
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
    for ($i = 0; $i -lt 6; $i++) {
        $playerState = $SyncHash.PlayerState[$i]

        # We only need to attach event handlers to the FIRST media player of each star,
        # as they all share the same source and will behave identically.
        if (-not $playerState.PlayerComponents) {
            Write-Warning "Could not find player components for Star Index $i. Skipping event handler attachment."
            continue
        }
        $mainPlayer = $playerState.PlayerComponents[0].Player
        $mainPlayer.Tag = $i # Store star index in the player's tag

        $mainPlayer.Add_MediaEnded({
                param($Sender, $EventArgs)
                $sIndex = $Sender.Tag
                $playerState = $SyncHash.PlayerState[$sIndex]
                if ($playerState -and $playerState.IsFailed) { return }

                $playerState.PlaybackStopwatch.Stop()
                $elapsedMilliseconds = $playerState.PlaybackStopwatch.Elapsed.TotalMilliseconds

                # If media "ends" in under 1 second and it's not an image, it's a failure.
                if (($elapsedMilliseconds -lt 1000) -and (-not $playerState.IsImage)) {
                    Handle-MediaFailure -StarIndex $sIndex -Reason "Playback failed or ended instantly." -SyncHash $SyncHash
                    return
                }

                Assign-NextMediaToStar -StarIndex $sIndex -SyncHash $SyncHash
            })

        $mainPlayer.Add_MediaOpened({
                param($Sender, $EventArgs)
                $sIndex = $Sender.Tag
                $pState = $SyncHash.PlayerState[$sIndex]
                if (-not $pState) { return }

                $SyncHash.Window.Dispatcher.Invoke([action]{
                    $pState.IsFailed = $false
                    $pState.PlaybackStopwatch.Restart() # Start timing the playback.

                    # On success, reset the UI state for ALL components of the star
                    foreach ($component in $pState.PlayerComponents) {
                        $component.Player.Opacity = 1 # Ensure player is visible
                        if ($SyncHash.RbSelection -ne "Hidden") {
                            $component.Overlay.Visibility = 'Visible'
                        }
                    }

                    if (-not $pState.InitialLoadComplete) {
                        $pState.InitialLoadComplete = $true
                        Check-LoadingComplete -SyncHash $SyncHash
                    }

                    if ($pState.IsImage) {
                        # Pause ALL media elements for this star, not just the sender.
                        foreach ($component in $pState.PlayerComponents) {
                            $component.Player.Pause()
                        }
                        if ($pState.ImageTimer) { $pState.ImageTimer.Stop() }
                        $timer = New-Object System.Windows.Threading.DispatcherTimer
                        $timer.Interval = [TimeSpan]::FromSeconds($SyncHash.ImageHoldSeconds)
                        $timer.Tag = $sIndex
                        $timer.Add_Tick({ $tickTimer = $args[0]; $tickIndex = $tickTimer.Tag; $tickTimer.Stop(); Assign-NextMediaToStar -StarIndex $tickIndex -SyncHash $SyncHash })
                        $pState.ImageTimer = $timer
                        $timer.Start()
                    }
                    elseif (-not $Sender.NaturalDuration.HasTimeSpan) {
                        # This is a silent failure for a non-image file (e.g., bad video codec).
                        Handle-MediaFailure -StarIndex $sIndex -Reason "Invalid duration or codec." -SyncHash $SyncHash
                    }
                })
            })

        $mainPlayer.Add_MediaFailed({
                param($s, $e)
                $sIndex = $s.Tag
                Handle-MediaFailure -StarIndex $sIndex -Reason $e.ErrorException.Message -SyncHash $SyncHash
            })

        Assign-NextMediaToStar -StarIndex $i -SyncHash $SyncHash

    }

    # --- Apply Text Overlay Settings on Load ---
    foreach ($playerState in $SyncHash.PlayerState.Values) {
        if (-not $playerState.PlayerComponents) { continue }

        foreach ($component in $playerState.PlayerComponents) {
            if ($SyncHash.RbSelection -ne "Hidden") {
                $overlay = $component.Overlay
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $overlay.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $overlay.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
                $overlay.FontSize = $SyncHash.FontSize
                if ($SyncHash.IsBold) { $overlay.FontWeight = 'Bold' }
                if ($SyncHash.IsItalic) { $overlay.FontStyle = 'Italic' }
            }
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

        $xBoundary = 9; $yBoundary = 5; $zBoundary = 3

        foreach ($star in $SyncHash.StarObjects) {
            $newX = $star.Translate.OffsetX + ($star.Velocity.X * $elapsed)
            $newY = $star.Translate.OffsetY + ($star.Velocity.Y * $elapsed)
            $newZ = $star.Translate.OffsetZ + ($star.Velocity.Z * $elapsed)

            $velX = $star.Velocity.X; $velY = $star.Velocity.Y; $velZ = $star.Velocity.Z

            if (($newX -gt $xBoundary -and $velX -gt 0) -or ($newX -lt -$xBoundary -and $velX -lt 0)) { $velX *= -1 }
            if (($newY -gt $yBoundary -and $velY -gt 0) -or ($newY -lt -$yBoundary -and $velY -lt 0)) { $velY *= -1 }
            if (($newZ -gt $zBoundary -and $velZ -gt 0) -or ($newZ -lt -$zBoundary -and $velZ -lt 0)) { $velZ *= -1 }
            $star.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)

            $star.Translate.OffsetX = $newX; $star.Translate.OffsetY = $newY; $star.Translate.OffsetZ = $newZ

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

        foreach ($playerState in $SyncHash.PlayerState.Values) {
            if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
            if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
            if ($playerState.PlayerComponents) {
                foreach($component in $playerState.PlayerComponents) {
                    $component.Player.Stop()
                    $component.Player.Source = $null
                    $component.Player.Close()
                }
            }
        }
    })

    $window.Add_Loaded({
        # Correctly calculate the 3D boundaries based on the camera's FOV and distance.
        $viewport = $SyncHash.Window.FindName("mainViewport")
        $camera = $viewport.Camera
        if ($camera -is [System.Windows.Media.Media3D.PerspectiveCamera]) {
            $distance = $camera.Position.Z
            $fovRadians = $camera.FieldOfView * ([Math]::PI / 180.0)
            $viewHeight3D = 2.0 * $distance * [Math]::Tan($fovRadians / 2.0)
            $aspectRatio = if ($viewport.ActualHeight -gt 0) { $viewport.ActualWidth / $viewport.ActualHeight } else { 1.0 }
            $viewWidth3D = $viewHeight3D * $aspectRatio
            $SyncHash.xBoundary = $viewWidth3D / 2.0
            $SyncHash.yBoundary = $viewHeight3D / 2.0
        } else {
            # Fallback for an unexpected camera type
            $SyncHash.xBoundary = 9.0
            $SyncHash.yBoundary = 5.0
        }

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
