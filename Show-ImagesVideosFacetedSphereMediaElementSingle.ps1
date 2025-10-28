<#
.SYNOPSIS
    Displays a playlist of media on an interactive, rotating, faceted 3D sphere.
.DESCRIPTION
    This script creates a WPF window and renders a 3D sphere. The sphere's geometry is
    programmatically generated with flat facets. It then prompts the user to select an image
    or video file, which is displayed on each facet of the sphere.

    The sphere has a continuous rotation animation and includes UI controls to pause, change speed, and randomize the rotation.

    This version uses the built-in Windows MediaElement for video playback, so video format
    support is limited to codecs installed on the local system.
.EXAMPLE
    PS C:\> .\Show-ImagesVideosFacetedSphereMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D faceted sphere window.
.NOTES
    Name:           Show-ImagesVideosFacetedSphereMediaElement.ps1
    Version:        1.0.0, 10/25/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Rotating Faceted Sphere"
$ScriptDescription = "Loops through and displays selected images or videos on each facet of a rotating 3D sphere. Uses the built-in Windows MediaElement."

# --- Dependency Check ---
$RequiredExecutables = @() # No external executables needed
if ($RequiredExecutables.Count -gt 0)
{
    # This block is kept for structural consistency with other scripts.
    # It will not run unless $RequiredExecutables is populated.
    $dependencyStatus = @()
    $allDependenciesMet = $true
    foreach ($exe in $RequiredExecutables)
    {
        if (-not (Get-Command $exe -ErrorAction SilentlyContinue))
        {
            $allDependenciesMet = $false
        }
    }
}

# --- Sphere Generation Function ---
function New-SphereMesh
{
    param(
        [double]$radius = 1.5,
        [int]$slices = 10, # Longitude
        [int]$stacks = 5  # Latitude
    )

    $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D

    # Step 1: Generate all unique vertices and store them in a temporary list, just like the working smooth sphere.
    $allVertices = New-Object System.Collections.Generic.List[System.Windows.Media.Media3D.Point3D]
    for ($stack = 0; $stack -le $stacks; $stack++)
    {
        $phi = [Math]::PI / 2 - $stack * [Math]::PI / $stacks
        $y = $radius * [Math]::Sin($phi)
        $r = $radius * [Math]::Cos($phi)

        for ($slice = 0; $slice -le $slices; $slice++)
        {
            $theta = $slice * 2 * [Math]::PI / $slices
            $x = $r * [Math]::Cos($theta)
            $z = $r * [Math]::Sin($theta)
            $allVertices.Add([System.Windows.Media.Media3D.Point3D]::new($x, $y, $z))
        }
    }

    # Step 2: Use the proven Triangle-Index logic from the working script to add vertices to the final mesh.
    # This guarantees the correct winding order, creating a solid object.
    for ($stack = 0; $stack -lt $stacks; $stack++)
    {
        for ($slice = 0; $slice -lt $slices; $slice++)
        {
            $i0 = $stack * ($slices + 1) + $slice
            $i1 = ($stack + 1) * ($slices + 1) + $slice
            $i2 = $i0 + 1
            $i3 = $i1 + 1

            # Get the vertices for the four corners of the facet
            $p0 = $allVertices[$i0]; $p1 = $allVertices[$i1]; $p2 = $allVertices[$i2]; $p3 = $allVertices[$i3]

            # Add the two triangles that form the facet. This order is proven to work.
            $mesh.Positions.Add($p0); $mesh.Positions.Add($p1); $mesh.Positions.Add($p2)
            $mesh.Positions.Add($p2); $mesh.Positions.Add($p1); $mesh.Positions.Add($p3)

            # Add TextureCoordinates for each vertex. This maps the whole media to each facet.
            $uv0 = [System.Windows.Point]::new(0, 0); $uv1 = [System.Windows.Point]::new(0, 1); $uv2 = [System.Windows.Point]::new(1, 0); $uv3 = [System.Windows.Point]::new(1, 1)
            $mesh.TextureCoordinates.Add($uv0); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv2)
            $mesh.TextureCoordinates.Add($uv2); $mesh.TextureCoordinates.Add($uv1); $mesh.TextureCoordinates.Add($uv3)
        }
    }
    return $mesh
}

# --- Main Application Loop ---
while ($true)
{
    # --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Faceted Sphere - Media Selector"
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

    [xml]$VideoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Preview Video - Click to Pause/Resume" Height="450" Width="800"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        SizeToContent="Manual"
        WindowState="Normal"
        WindowStyle="ToolWindow"
        Background="Black">
    <Grid x:Name="TheGrid">
        <!-- MediaElement for video playback -->
        <MediaElement x:Name="MediaPlayer" 
            LoadedBehavior="Manual" 
            UnloadedBehavior="Stop" />
                      
    </Grid>
</Window>
"@

    $DataGridView.Add_RowHeaderMouseClick({
            # Access the row index from the event arguments
            $rowIndex = $_.RowIndex

            # Get the row object
            $row = $DataGridView.Rows[$rowIndex]

            # Now, you can access the cell value by its column header (name)
            # Replace "YourColumnHeader" with the actual header of the column you want
            $videoPath = ($row.Cells["FilePath"].Value)

            $SyncHash.PreviewPaused = $False

            $VideoReader = (New-Object System.Xml.XmlNodeReader $VideoXaml)
            $VideoWindow = [Windows.Markup.XamlReader]::Load($VideoReader)

            # Find controls by their x:Name
            $TheGrid = $VideoWindow.FindName("TheGrid")
            $MediaPlayer = $VideoWindow.FindName("MediaPlayer")

            $NewUri = New-Object System.Uri($videoPath, [System.UriKind]::Absolute)

            $MediaPlayer.Source = $NewUri

            # Define the MouseDown event handler
            $TheGrid.Add_MouseDown({
                    if($SyncHash.PreviewPaused)
                    {
                        $MediaPlayer.Play()
                        $SyncHash.PreviewPaused = $False
                    }
                    else
                    {
                        $MediaPlayer.Pause()
                        $SyncHash.PreviewPaused = $True
                    }
                })

            $MediaPlayer.Add_MediaEnded({
                    $MediaPlayer.Position = [TimeSpan]::FromSeconds(0)
                    $MediaPlayer.Play()
                })

            # Display the window
            $MediaPlayer.Play()
            $VideoWindow.ShowDialog() | Out-Null
        })

    $DataGridView.Add_CellPainting({
        param($sender, $e)
        # Check if it's a row header cell and not the top-left header cell
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -lt 0) {
            $e.PaintBackground($e.ClipBounds, $true) # Paint the background
            
            # Define text format for centering
            $format = New-Object System.Drawing.StringFormat
            $format.Alignment = [System.Drawing.StringAlignment]::Center
            $format.LineAlignment = [System.Drawing.StringAlignment]::Center
            
            # Draw the text from the HeaderCell.Value
            # We must explicitly use a RectangleF to avoid overload resolution errors with DrawString.
            $rectF = New-Object System.Drawing.RectangleF($e.CellBounds.X, $e.CellBounds.Y, $e.CellBounds.Width, $e.CellBounds.Height)
            $e.Graphics.DrawString($e.FormattedValue.ToString(), $e.CellStyle.Font, [System.Drawing.Brushes]::Black, $rectF, $format)
            $e.Handled = $true # Mark the event as handled
        }
    })

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
    $formState = @{ TextColor = [System.Drawing.Color]::Black }
    $ColorExample.BackColor = $formState.TextColor
    $SelectColorButton.Add_Click({
            $colorDialog = New-Object System.Windows.Forms.ColorDialog
            if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $formState.TextColor = $colorDialog.Color
                $ColorExample.BackColor = $formState.TextColor
            }
        })

    $formState.FontFamily = "Arial"
    $FontButton.Add_Click({
            $fontDialog = New-Object System.Windows.Forms.FontDialog
            if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $formState.FontFamily = $fontDialog.Font.Name
                $FontButton.Text = $formState.FontFamily
            }
        })

    $updateFontStyle = {
        $style = [System.Drawing.FontStyle]::Regular
        if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
        if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
    }
    $ItalicCheckbox.Add_CheckedChanged($updateFontStyle)
    $BoldCheckbox.Add_CheckedChanged($updateFontStyle)

    $SelectAllCheckbox.Add_CheckedChanged({
            $isChecked = $SelectAllCheckbox.Checked
            foreach ($row in $DataGridView.Rows)
            {
                $row.Cells["Select"].Value = $isChecked
            }
            $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        })

    $BrowseButton.Add_Click({
            $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $FolderBrowser.Description = "Select the folder to scan."
            if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
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
                if ($RecursiveCheckBox.Checked)
                {
                    $gciParams.Path = $SelectedPath
                    $gciParams.Recurse = $true
                }
                else
                {
                    $gciParams.Path = Join-Path $SelectedPath "*"
                }
                $files = Get-ChildItem @gciParams
                foreach ($file in $files)
                {
                    $DataGridView.Rows.Add($false, $file.Name, $file.FullName)
                }

                foreach ($row in $DataGridView.Rows)
                {
                    if ($row.IsNewRow) { continue }
                    $row.HeaderCell.Value = "Play"
                }
            }
        })

    $PlayButton.Add_Click({
            $formState.SelectedFiles = @(
                foreach ($Row in $DataGridView.Rows)
                {
                    if ($Row.Cells["Select"].Value)
                    {
                        $Row.Cells["FilePath"].Value
                    }
                }
            )
            if ($formState.SelectedFiles.Count -gt 0)
            {
                $formState.UseTransparentEffect = $TransparentCheckbox.Checked
            
                # Capture text settings
                if ($RadioButton1.Checked) { $formState.RbSelection = "Hidden" }
                if ($RadioButton2.Checked) { $formState.RbSelection = "Filename" }
                if ($RadioButton3.Checked) { $formState.RbSelection = "Custom" }
                $formState.CustomText = $TextBox.Text
                $formState.FontSize = $NumericUpDown.Value
                $formState.IsBold = $BoldCheckbox.Checked
                $formState.IsItalic = $ItalicCheckbox.Checked
                $SelectForm.Close()
            }
            else
            {
                [System.Windows.Forms.MessageBox]::Show("No files selected.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        })

    $null = $SelectForm.ShowDialog()
    $SelectForm.Dispose()

    # Exit if no files were selected or the form was closed
    if (-not $formState.ContainsKey("SelectedFiles") -or $formState.SelectedFiles.Count -eq 0)
    {
        Write-Host "No files were selected or form was closed. Exiting."
        break # Exit the main while loop
    }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
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
                                <AxisAngleRotation3D x:Name="AxisAngleX" Axis="1,0,0" Angle="0"/>
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0"/>
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

    # --- Synchronized Hashtable for state management ---
    $SyncHash = [hashtable]::Synchronized(@{
            SelectedFiles        = $formState.SelectedFiles
            UseTransparentEffect = $formState.UseTransparentEffect
            CurrentIndex         = -1 # Will be incremented to 0 on first run
            MediaTimer           = $null
            CurrentMediaElement  = $null
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

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # --- Find the container for our sphere ---
    $sphereContainer = $window.FindName("SphereContainer")

    # --- Create the Sphere Geometry ---
    # Calculate a dynamic radius to make the sphere's height a percentage of the viewport height.
    $cameraDistance = 8.0
    $cameraFovDegrees = 60.0
    $cameraFovRadians = $cameraFovDegrees * ([Math]::PI / 180.0)
    $visibleHeightAtOrigin = 2.0 * $cameraDistance * [Math]::Tan($cameraFovRadians / 2.0)

    $desiredHeightPercentage = 0.50 # 50% of the viewport height
    $dynamicRadius = ($visibleHeightAtOrigin * $desiredHeightPercentage) / 2.0
    $sphereMesh = New-SphereMesh -radius $dynamicRadius

    # --- Set Window to Full Screen ---
    $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $window.Width = $primaryScreen.WorkingArea.Width
    $window.Height = $primaryScreen.WorkingArea.Height
    $window.Left = $primaryScreen.WorkingArea.Left
    $window.Top = $primaryScreen.WorkingArea.Top

    $SyncHash.Window = $window # Store window for event handlers

    # --- Create the 3D visual element for the sphere ---
    $sphereViewport = New-Object System.Windows.Media.Media3D.Viewport2DVisual3D
    $sphereViewport.Geometry = $sphereMesh

    # --- Create the material and link it to the visual host ---
    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
    $sphereMaterial = New-Object $materialType
    [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($sphereMaterial, $true)
    $sphereViewport.Material = $sphereMaterial

    # --- Create the 2D content host (Grid and ContentPresenter) that will be displayed on the sphere ---
    $mediaHostGrid = New-Object System.Windows.Controls.Grid
    $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
    $mediaHostGrid.Children.Add($contentPresenter)

    # Create and add the text overlay block
    $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{
        HorizontalAlignment = 'Center'; VerticalAlignment = 'Center'
        TextWrapping = 'Wrap'; TextAlignment = 'Center'; IsHitTestVisible = $false;
    }
    $mediaHostGrid.Children.Add($overlayTextBlock)

    $sphereViewport.Visual = $mediaHostGrid
    $SyncHash.ContentPresenter = $contentPresenter

    # Add the complete sphere visual to the main scene
    $sphereContainer.Children.Add($sphereViewport)

    # --- Media Handling Functions ---
    function Start-NextMedia
    {
        # Clean up previous media resources
        if ($SyncHash.MediaTimer) { $SyncHash.MediaTimer.Stop() }
        if ($SyncHash.CurrentMediaElement)
        {
            $SyncHash.CurrentMediaElement.Stop()
            $SyncHash.CurrentMediaElement.Source = $null
            $SyncHash.CurrentMediaElement = $null
        }

        # Get next media file, looping if necessary
        $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $SyncHash.SelectedFiles.Count
        $filePath = $SyncHash.SelectedFiles[$SyncHash.CurrentIndex]

        # Update text overlay if set to "Filename"
        if ($SyncHash.RbSelection -eq "Filename")
        {
            $overlayTextBlock.Text = [System.IO.Path]::GetFileName($filePath)
        }

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

        if ($ImageExtensions -contains $extension)
        {
            $image = New-Object System.Windows.Controls.Image
            $image.Source = [System.Windows.Media.Imaging.BitmapImage]::new([Uri]$filePath)
            $image.Stretch = "Fill"
            $SyncHash.ContentPresenter.Content = $image

            # Set a timer to show the next media item after 10 seconds
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds(10)
            $timer.Add_Tick({
                    $args[0].Stop() # The timer object is passed as the first argument
                    Start-NextMedia
                })
            $timer.Start()
            $SyncHash.MediaTimer = $timer
        }
        else
        {
            $mediaElement = New-Object System.Windows.Controls.MediaElement
            $mediaElement.LoadedBehavior = 'Manual' # Set to Manual for programmatic control
            $mediaElement.UnloadedBehavior = 'Stop'
            $mediaElement.Stretch = 'Fill'
            $mediaElement.Source = [Uri]$filePath
            $mediaElement.Add_MediaEnded({ Start-NextMedia })
            $mediaElement.Add_MediaFailed({ Start-NextMedia })
            $mediaElement.Play() # Explicitly start playback
            $SyncHash.ContentPresenter.Content = $mediaElement
            $SyncHash.CurrentMediaElement = $mediaElement
        }
    }

    # --- Animate the Sphere ---
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

    # --- Find and Wire Up UI Controls ---
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.closeButton = $window.FindName("closeButton")

    $SyncHash.closeButton.Add_Click({ $window.Close() })

    $SyncHash.pauseButton.Add_Click({
            if ($SyncHash.Paused)
            {
                $SyncHash.animX.From = $SyncHash.AxisAngleX.Angle
                $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animX)
                $SyncHash.animY.From = $SyncHash.AxisAngleY.Angle
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $SyncHash.animY)
                $SyncHash.pauseButton.Content = "Pause"
                $SyncHash.Paused = $false
            }
            else
            {
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

        if (-not $SyncHash.Paused)
        {
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Pause
            Start-Sleep -Milliseconds 50 # Give it a moment to process
            $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) # Resume
        }
    }

    $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })

    $SyncHash.redoButton.Add_Click({
            $SyncHash.RedoClicked = $true
            $SyncHash.Window.Close()
        })

    $SyncHash.hideControlsButton.Add_Click({
            $controlsPanel = $window.FindName("controlsPanel")
            if ($SyncHash.ControlsHidden)
            {
                $controlsPanel.Visibility = 'Visible'
                $SyncHash.ControlsHidden = $false
            }
            else
            {
                $controlsPanel.Visibility = 'Collapsed'
                $SyncHash.ControlsHidden = $true
            }
        })

    # --- Handle Window Events ---
    $window.Add_KeyDown({
            param($sender, $e)
            switch ($e.Key)
            {
                "Escape" { $window.Close() }
                "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                "A" { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                "Left" { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                "Right" { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            }
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
                if ($result -is [System.Windows.Media.Media3D.RayMeshGeometry3DHitTestResult])
                {
                    $SyncHash.hitVisual = $result.VisualHit
                    return [System.Windows.Media.HitTestResultBehavior]::Stop
                }
                return [System.Windows.Media.HitTestResultBehavior]::Continue
            }

            # Perform the hit test
            $hitTestParams = [System.Windows.Media.PointHitTestParameters]::new($mousePosition)
            [System.Windows.Media.VisualTreeHelper]::HitTest($viewport, $null, $hitTestCallback, $hitTestParams)

            # If the hit visual is our sphere, trigger the pause button
            if ($SyncHash.hitVisual -is [System.Windows.Media.Media3D.Viewport2DVisual3D])
            {
                $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        })

    $window.Add_Closed({
            # Stop animations and timers to prevent resource leaks
            if ($SyncHash.MediaTimer) { $SyncHash.MediaTimer.Stop() }
            if ($SyncHash.CurrentMediaElement)
            {
                $SyncHash.CurrentMediaElement.Stop()
                $SyncHash.CurrentMediaElement.Source = $null
            }
            $axisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
            $axisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null)
        })

    # --- Start the show ---
    $window.Add_Loaded({
            # Start the media cycle once the window is fully loaded
            Start-NextMedia
        })

    # --- Apply Text Overlay Settings on Load ---
    switch ($SyncHash.RbSelection)
    {
        "Hidden"
        {
            $overlayTextBlock.Visibility = 'Collapsed'
        }
        "Filename"
        {
            # Text is set in Start-NextMedia
        }
        "Custom"
        {
            $overlayTextBlock.Text = $SyncHash.CustomText
        }
    }

    if ($SyncHash.RbSelection -ne "Hidden")
    {
        $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
        $overlayTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
        $overlayTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.FontFamily)
        $overlayTextBlock.FontSize = $SyncHash.FontSize
        if ($SyncHash.IsBold) { $overlayTextBlock.FontWeight = 'Bold' }
        if ($SyncHash.IsItalic) { $overlayTextBlock.FontStyle = 'Italic' }
    }

    $null = $window.ShowDialog()

    # After window closes, check if we need to loop
    if (-not $SyncHash.RedoClicked)
    {
        break # Exit the main while loop
    }
}