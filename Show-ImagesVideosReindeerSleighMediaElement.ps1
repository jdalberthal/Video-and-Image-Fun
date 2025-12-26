<#
.SYNOPSIS
    Displays a sleigh pulled by 4 reindeer flying across the screen, with media playing on each object.

.DESCRIPTION
    This script launches a GUI to select image and video files. It then renders a 3D scene
    featuring a sleigh and four reindeer. Each of these 5 objects displays media from the
    selected playlist. The team flies across the screen from right to left, undulating
    up and down, and loops continuously.

    This version uses the built-in Windows MediaElement for video playback.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosReindeerSleighMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the festive 3D window.

.NOTES
    Name:           Show-ImagesVideosReindeerSleighMediaElement.ps1
    Version:        1.0.0, 12/25/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access.
#>

Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# --- Script Metadata ---
$ExternalButtonName = "Reindeer & Sleigh"
$ScriptDescription = "A festive display of 4 reindeer pulling a sleigh, all playing media."
$RequiredExecutables = @()

# --- Main Application Loop ---
while ($true) {
    #region --- File Selection Form ---
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $SelectForm = New-Object System.Windows.Forms.Form
    $SelectForm.Text = "Reindeer & Sleigh - Media Selector"
    $SelectForm.Size = New-Object System.Drawing.Size(800, 600)
    $SelectForm.StartPosition = "CenterScreen"

    $BrowseButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Browse Folder"; Location = '10, 10'; Size = '100, 25' }
    $FolderPathTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = '120, 10'; Size = '450, 25'; ReadOnly = $true }
    $RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Include Subfolders"; AutoSize = $true; Location = '10, 40'; Checked = $false }
    $TransparentCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Make Semi-Transparent"; AutoSize = $true; Location = '150, 40'; Checked = $false }
    $NightSkyCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Dark Night Sky"; AutoSize = $true; Location = '320, 40'; Checked = $true }
    $TwinkleCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Twinkling Stars"; AutoSize = $true; Location = '450, 40'; Checked = $true }
    
    $SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Select All"; AutoSize = $true; Location = '10, 70'; Checked = $false }
    $DataGridView = New-Object System.Windows.Forms.DataGridView -Property @{ Location = '10, 95'; Size = '760, 400'; Anchor = 'Top, Bottom, Left, Right'; AutoGenerateColumns = $false; AllowUserToAddRows = $false; RowHeadersWidth = 65 }
    
    $SelectForm.Controls.AddRange(@($BrowseButton, $FolderPathTextBox, $RecursiveCheckBox, $TransparentCheckbox, $NightSkyCheckbox, $TwinkleCheckbox, $SelectAllCheckbox, $DataGridView))

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = "Select"; HeaderText = ""; Width = 30 }
    $FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FileName"; HeaderText = "File Name"; Width = 250; ReadOnly = $true }
    $FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = "FilePath"; HeaderText = "File Path"; Width = 450; ReadOnly = $true }
    $DataGridView.Columns.Add($CheckBoxColumn) | Out-Null
    $DataGridView.Columns.Add($FileNameColumn) | Out-Null
    $DataGridView.Columns.Add($FilePathColumn) | Out-Null

    $PlayButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Play Selected"; Location = '600, 40'; Size = '170, 30' }
    $SelectForm.Controls.Add($PlayButton)

    # --- Form Logic ---
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
            $VideoExtensions = "*.mp4", "*.m4v", "*.wmv", "*.avi", "*.mpg", "*.mpeg"
            $AllowedExtensions = $ImageExtensions + $VideoExtensions
            
            $gciParams = @{ File = $true; Include = $AllowedExtensions }
            if ($RecursiveCheckBox.Checked) { $gciParams.Path = $SelectedPath; $gciParams.Recurse = $true } 
            else { $gciParams.Path = Join-Path $SelectedPath "*" }

            Get-ChildItem @gciParams | ForEach-Object { $DataGridView.Rows.Add($false, $_.Name, $_.FullName) | Out-Null }
            $DataGridView.Rows | ForEach-Object { if (-not $_.IsNewRow) { $_.HeaderCell.Value = "Play" } }
        }
    })

    $formState = @{}
    $PlayButton.Add_Click({
        $selectedFiles = @($DataGridView.Rows | Where-Object { $_.Cells["Select"].Value } | ForEach-Object { $_.Cells["FilePath"].Value })
        if ($selectedFiles.Count -gt 0) {
            $formState.SelectedFiles = [System.Collections.ArrayList]::new($selectedFiles)
            $formState.UseTransparentEffect = $TransparentCheckbox.Checked
            $formState.NightSky = $NightSkyCheckbox.Checked
            $formState.TwinklingStars = $TwinkleCheckbox.Checked
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
    #endregion

    # --- Central State ---
    $SyncHash = [hashtable]::Synchronized(@{
        SelectedFiles        = $formState.SelectedFiles
        UseTransparentEffect = $formState.UseTransparentEffect
        TwinklingStars       = $formState.TwinklingStars
        CurrentIndex         = -1
        PlayerStates         = [hashtable]::Synchronized(@{})
        Paused               = $false; ControlsHidden = $false; RedoClicked = $false
        LastFrameTime        = [System.Diagnostics.Stopwatch]::GetTimestamp()
        SpeedMultiplier      = 1.0
        LeadPositionX        = 30.0
        LeadPositionY        = 0.0
        TotalDistance        = 0.0
        HistoryX             = [System.Collections.Generic.List[double]]::new()
        HistoryY             = [System.Collections.Generic.List[double]]::new()
        HistoryDist          = [System.Collections.Generic.List[double]]::new()
    })

    # --- Helper Functions ---
    function Get-NextMediaIndex {
        $fileCount = $SyncHash.SelectedFiles.Count
        if ($fileCount -eq 0) { return -1 }
        $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $fileCount
        return $SyncHash.CurrentIndex
    }

    function Handle-MediaEnded {
        param([string]$PlayerKey)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if ($pState.IsFailed) { return } 
        Start-NextMedia -PlayerKey $PlayerKey
    }

    function Handle-MediaOpened {
        param([string]$PlayerKey, $EventArgs)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $pState) { return }
        $pState.IsFailed = $false
        if ($pState.MediaHostGrid) { $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Transparent }
    }

    function Handle-MediaFailure {
        param([string]$PlayerKey, [string]$Reason)
        $pState = $SyncHash.PlayerStates[$PlayerKey]
        $pState.IsFailed = $true
        $SyncHash.Window.Dispatcher.Invoke([action]{
            if ($pState.CurrentMediaElement) { $pState.CurrentMediaElement.Close() }
            if ($pState.MediaHostGrid) { $pState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
        })
        # Try next media after short delay
        $recoveryScriptBlock = { Start-NextMedia -PlayerKey $PlayerKey }
        $null = $SyncHash.Window.Dispatcher.InvokeAsync($recoveryScriptBlock.GetNewClosure())
    }

    function Start-NextMedia {
        param([string]$PlayerKey)
        $playerState = $SyncHash.PlayerStates[$PlayerKey]
        if (-not $playerState) { return }

        if ($playerState.MediaTimer) { $playerState.MediaTimer.Stop() }
        if ($playerState.CurrentMediaElement) { $playerState.CurrentMediaElement.Close() }

        $nextIndex = Get-NextMediaIndex
        if ($nextIndex -lt 0) { return }

        $filePath = $SyncHash.SelectedFiles[$nextIndex]
        $playerState.CurrentSource = [Uri]$filePath
        $playerState.IsFailed = $false

        $ImageExtensions = ".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico"
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        $playerState.IsImage = $ImageExtensions -contains $extension

        try {
            if ($playerState.IsImage) {
                $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                $bitmap.BeginInit(); $bitmap.UriSource = $playerState.CurrentSource; $bitmap.EndInit(); $bitmap.Freeze()
                $image = New-Object System.Windows.Controls.Image -Property @{ Source = $bitmap; Stretch = 'Fill' }
                $playerState.ContentPresenter.Content = $image
                $timer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(10); Tag = $PlayerKey }
                $timer.Add_Tick({ $t = $args[0]; $key = $t.Tag; $t.Stop(); Handle-MediaEnded -PlayerKey $key })
                $playerState.MediaTimer = $timer; $timer.Start()
            } else { # Video
                $mediaElement = New-Object System.Windows.Controls.MediaElement -Property @{
                    LoadedBehavior = 'Manual'; UnloadedBehavior = 'Stop'; Stretch = 'Fill'; Source = $playerState.CurrentSource; Tag = $PlayerKey
                }
                $mediaElement.Add_MediaEnded($playerState.MediaEndedHandler)
                $mediaElement.Add_MediaOpened($playerState.MediaOpenedHandler)
                $mediaElement.Add_MediaFailed($playerState.MediaFailedHandler)
                
                $playerState.ContentPresenter.Content = $mediaElement
                $playerState.CurrentMediaElement = $mediaElement
                $mediaElement.Play()
                if ($playerState.MediaHostGrid) { $playerState.MediaHostGrid.Background = [System.Windows.Media.Brushes]::Black }
            }
        } catch {
            Handle-MediaFailure -PlayerKey $PlayerKey -Reason $_.Exception.Message
        }
    }

    # --- 3D Setup ---
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Reindeer &amp; Sleigh"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid x:Name="MainGrid">
        <Canvas x:Name="StarCanvas" />
        <Viewport3D x:Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0,0,20" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="60"/>
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
        </Viewport3D> 
        <Canvas x:Name="VisualHost" Opacity="0"/>
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
    $viewport = $window.FindName("mainViewport")
    $starCanvas = $window.FindName("StarCanvas")
    $visualHost = $window.FindName("VisualHost")

    if ($formState.NightSky) {
        if ($SyncHash.UseTransparentEffect) {
            $window.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(128, 0, 0, 0)) # 50% transparent black
        } else {
            $window.Background = [System.Windows.Media.Brushes]::Black
        }
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $window.Width = $workArea.Width; $window.Height = $workArea.Height
    $window.Left = $workArea.Left; $window.Top = $workArea.Top

    if ($SyncHash.TwinklingStars) {
        $rand = [Random]::new()
        for ($i = 0; $i -lt 200; $i++) {
            $star = New-Object System.Windows.Shapes.Ellipse
            $size = $rand.NextDouble() * 3 + 1
            $star.Width = $size; $star.Height = $size
            $star.Fill = [System.Windows.Media.Brushes]::White
            [System.Windows.Controls.Canvas]::SetLeft($star, $rand.NextDouble() * $window.Width)
            [System.Windows.Controls.Canvas]::SetTop($star, $rand.NextDouble() * $window.Height)
            
            $anim = New-Object System.Windows.Media.Animation.DoubleAnimation(0.1, 1.0, [TimeSpan]::FromSeconds($rand.NextDouble() * 2 + 0.5))
            $anim.AutoReverse = $true
            $anim.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
            $star.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $anim)
            $starCanvas.Children.Add($star) | Out-Null
        }
    }

    # Calculate visible boundaries to determine start/end points
    $aspectRatio = $window.Width / $window.Height
    $fovRadians = 60 * ([Math]::PI / 180)
    $visibleHeight = 2 * 20 * [Math]::Tan($fovRadians / 2)
    $visibleWidth = $visibleHeight * $aspectRatio
    $SyncHash.RightLimit = ($visibleWidth / 2)
    $SyncHash.LeftLimit = -($visibleWidth / 2) - 12
    $SyncHash.LeadPositionX = $SyncHash.RightLimit

    # --- Create Models ---
    # Simple plane mesh for screens
    $planeMesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
    $planeMesh.Positions = "-0.5,-0.5,0 0.5,-0.5,0 0.5,0.5,0 -0.5,0.5,0"
    $planeMesh.TriangleIndices = "0,1,2 0,2,3"
    $planeMesh.TextureCoordinates = "0,1 1,1 1,0 0,0"
    $planeMesh.Freeze()

    $materialType = if ($SyncHash.UseTransparentEffect) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

    # Helper to create a media object
    function New-MediaObject {
        param($Name, $ScaleX, $ScaleY, $ClipData)
        
        $mediaHostGrid = New-Object System.Windows.Controls.Grid
        $mediaHostGrid.Width = 500; $mediaHostGrid.Height = 300
        if ($SyncHash.UseTransparentEffect) {
            $mediaHostGrid.Opacity = 0.7 # Make the media itself semi-transparent
        }

        $contentPresenter = New-Object System.Windows.Controls.ContentPresenter
        $mediaHostGrid.Children.Add($contentPresenter) | Out-Null
        if ($ClipData) {
            # Parse the geometry and clone it to make it modifiable (not frozen)
            $geometry = [System.Windows.Media.Geometry]::Parse($ClipData).Clone()
            
            $bounds = $geometry.Bounds
            
            # Calculate scale factors to fit the geometry within the grid, maintaining aspect ratio.
            $scaleX_factor = if ($bounds.Width -gt 0) { $mediaHostGrid.Width / $bounds.Width } else { 1 }
            $scaleY_factor = if ($bounds.Height -gt 0) { $mediaHostGrid.Height / $bounds.Height } else { 1 }
            $scale = [Math]::Min($scaleX_factor, $scaleY_factor) * 0.95 # Use 95% to add a small margin

            # Center the scaled geometry within the grid.
            $scaledWidth = $bounds.Width * $scale
            $scaledHeight = $bounds.Height * $scale
            $translateX = ($mediaHostGrid.Width - $scaledWidth) / 2 - ($bounds.X * $scale)
            $translateY = ($mediaHostGrid.Height - $scaledHeight) / 2 - ($bounds.Y * $scale)

            $transformGroup = New-Object System.Windows.Media.TransformGroup
            $transformGroup.Children.Add((New-Object System.Windows.Media.ScaleTransform($scale, $scale)))
            $transformGroup.Children.Add((New-Object System.Windows.Media.TranslateTransform($translateX, $translateY)))
            $geometry.Transform = $transformGroup
            
            $mediaHostGrid.Clip = $geometry

            # Create a Path element to draw the border on top of the media
            $borderPath = New-Object System.Windows.Shapes.Path
            $borderPath.Data = $geometry # Use the same transformed geometry
            $borderPath.Stroke = [System.Windows.Media.Brushes]::LightGray
            $borderPath.StrokeThickness = 2
            $borderPath.Fill = [System.Windows.Media.Brushes]::Transparent # Ensure the inside of the path is not filled
            $mediaHostGrid.Children.Add($borderPath) | Out-Null
        }

    if ($Name -eq "Deer1") {
        $nose = New-Object System.Windows.Shapes.Ellipse -Property @{
            Width = 24
            Height = 24
            Fill = [System.Windows.Media.Brushes]::Red
            HorizontalAlignment = 'Left'
            VerticalAlignment = 'Top'
            Margin = [System.Windows.Thickness]::new(91.5, 81.5, 0, 0) # Position over the nose of the clipped shape
        }
        $mediaHostGrid.Children.Add($nose) | Out-Null

        # Create the blinking animation for the nose
        $blinkAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation(1.0, 0.1, [TimeSpan]::FromSeconds(1.5))
        $blinkAnimation.AutoReverse = $true
        $blinkAnimation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
        $nose.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $blinkAnimation)
    }

        $visualHost.Children.Add($mediaHostGrid) | Out-Null

        $playerState = [hashtable]@{ ContentPresenter = $contentPresenter; MediaHostGrid = $mediaHostGrid; IsFailed = $false; IsImage = $false; CurrentSource = $null }
        $playerState.MediaEndedHandler = { Handle-MediaEnded -PlayerKey $Name }.GetNewClosure()
        $playerState.MediaOpenedHandler = { Handle-MediaOpened -PlayerKey $Name -EventArgs $args[0] }.GetNewClosure()
        $playerState.MediaFailedHandler = { param($sender, $eventArgs) Handle-MediaFailure -PlayerKey $Name -Reason $eventArgs.ErrorException.Message }.GetNewClosure()
        $SyncHash.PlayerStates[$Name] = $playerState

        $visualBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGrid }
        $material = New-Object $materialType -Property @{ Brush = $visualBrush }
        if ($SyncHash.UseTransparentEffect) { $material.Color = [System.Windows.Media.Colors]::White }

        $model = New-Object System.Windows.Media.Media3D.GeometryModel3D -Property @{ Geometry = $planeMesh; Material = $material }
        $model.Transform = New-Object System.Windows.Media.Media3D.ScaleTransform3D($ScaleX, $ScaleY, 1) # Apply scale to the geometry

        $backMaterial = $material.Clone()
        # Flip back material so it's visible from behind if needed
        $backTransform = New-Object System.Windows.Media.ScaleTransform; $backTransform.ScaleX = -1
        $backMaterial.Brush.RelativeTransform = $backTransform
        $model.BackMaterial = $backMaterial

        $objectContainer = New-Object System.Windows.Media.Media3D.ModelVisual3D -Property @{ Content = $model }
        
        # Each object gets its own translate transform for the animation loop to control
        $translateTransform = New-Object System.Windows.Media.Media3D.TranslateTransform3D(0, 0, 0)
        $objectContainer.Transform = $translateTransform
        
        # Store the transform for the animation loop
        $SyncHash.PlayerStates[$Name].TranslateTransform = $translateTransform

        # Add the object's container directly to the viewport
        $viewport.Children.Add($objectContainer) | Out-Null
        
        Start-NextMedia -PlayerKey $Name
    }

    # Sleigh (Back/Right) - Modified to include Santa silhouette
    # Runners + Body with Santa bump
$sleighPath = "
M27,24
c0.6,0,1-0.4,1-1
c0-3.8,0.9-7.5,2.6-10.9
l1.3-2.6
c0.2-0.3,0.1-0.7,0-1
C31.7,8.2,31.3,8,31,8

L28,11
C27.6,10.8 27.1,8.6 26.4,7.2
C26.4,6 25.4,2 24,1.0
C21.6,5.6 20.8,7.2 21.2,8.8
C21.4,9.8 22.2,10.6 23.2,11.2
C22.2,11.6 21.2,12.6 21,14
C20.8,15.6 21.6,16.6 22.8,17.2
L21,16.4

c-0.7,1-1.8,1.5-2.9,1.5
c-1.3,0-2.4-0.7-3.1-1.7
c-1.2-2-3.4-3.3-5.7-3.3
H6
c-0.4,0-0.7,0.2-0.9,0.5
c-0.2,0.3-0.2,0.7,0,1
l5,9
c0.2,0.3,0.5,0.5,0.9,0.5
h1.5
l-1.2,3
l-5.8,0
c-1,0-2-0.3-2.7-1
C2.3,25.5,2,24.9,2,24.3
s0.3-1.3,0.8-1.7
c0.8-0.7,2.1-0.7,2.8,0
c0.4,0.4,1,0.4,1.4,0
c0.4-0.4,0.4-1,0-1.4
c-1.5-1.4-4-1.4-5.6,0
c-0.9,0.8-1.4,2-1.4,3.2
s0.5,2.3,1.4,3.2
c1.1,1,2.5,1.5,3.9,1.5
c0.1,0,0.1,0,0.2,0
l6.4,0
l13,0
l5,0
c0.6,0,1-0.4,1-1
s-0.4-1-1-1
l-4.3,0
l-1.2-3
H27
z
"


    New-MediaObject -Name "Sleigh" -ScaleX 5 -ScaleY 3 -ClipData $sleighPath

    # Reindeer in a line (Front is Left/Negative X), with a custom clip path for their shape
    $reindeerPathImproved = "M505.474,436.173c0,0-64.29-87.208-86.987-116.34c-10.781-13.882-21.414-24.368-32.196-32.048 c-0.055-0.037-0.101-0.064-0.157-0.101c0.804-0.139,1.468-0.148,2.299-0.305c26.953-5.28,33.562-41.906,31.466-57.006 c-2.086-15.101-8.538-20.528-18.313-3.286c-7.763,13.624-34.641,43.272-70.234,43.124l-58.041-0.184 c-3.775-0.037-7.42-0.147-10.928-0.37c-21.267-1.439-37.882-7.42-54.497-24.147c-8.916-8.999-12.627-24.718-15.507-39.034 c6.849,0.028,15.23,0.055,19.836,0.082c7.699,0,11.575-1.92,16.421-6.71c4.865-4.809,8.64-11.52,14.104-20.242 c6.019-9.674,0.036-18.276-6.72-18.313c-6.756,0-30.424-0.111-30.424-0.111c-5.787,0-11.592,1.92-14.51,3.84 c-2.427,1.597-7.31,5.584-12.572,10.08l-91.62-0.332c-12.444-0.074-22.855,1.957-31.272,5.427 c-13.542,5.575-22.117,14.99-26.206,25.993c-12.304,32.943,15.572,79.972,70.105,80.157l30.599,0.111 c9.609,31.568,25.789,65.72,25.789,65.72l-80.756,94.27c-7.108,5.917-9.831,19.236-1.532,27.562 c8.27,8.326,18.017,8.086,33.441-3.729l96.642-76.824l180.897,0.48l7.606,5.501l87.042,63.911 c11.852,8.935,21.598,7.145,27.894,0.083C514.252,455.464,513.44,445.394,505.474,436.173z M37.974,166.148c0.037-10.495-8.419-19.042-18.95-19.07C8.565,147.049,0.056,155.515,0,165.991 c-0.027,10.494,8.418,19.032,18.922,19.069C29.408,185.089,37.937,176.633,37.974,166.148z M154.277,142.093c-7.191,8.603-14.197,14.805-20.095,19.08h37.817c7.651-8.806,15.34-19.624,22.412-32.74 c11.261-20.897,20.934-47.62,26.361-81.19c1.145-7.052-3.655-13.689-10.67-14.842c-7.089-1.145-13.734,3.618-14.842,10.707 c-1.736,10.531-3.877,20.233-6.342,29.204l-17.722-18.424c-4.994-5.132-13.154-5.27-18.313-0.333 c-5.141,4.948-5.299,13.145-0.333,18.276l26.602,27.618c0.12,0.111,0.231,0.184,0.314,0.296 c-3.268,7.79-6.812,14.806-10.421,21.082l-16.541-11.926c-5.797-4.162-13.883-2.88-18.055,2.926 c-4.172,5.787-2.88,13.882,2.917,18.046L154.277,142.093z"
    New-MediaObject -Name "Deer1" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPathImproved
    New-MediaObject -Name "Deer2" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPathImproved
    New-MediaObject -Name "Deer3" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPathImproved
    New-MediaObject -Name "Deer4" -ScaleX 3 -ScaleY 2 -ClipData $reindeerPathImproved

    # --- Animation Loop ---
    $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
    $SyncHash.StartTime = $SyncHash.LastFrameTime
    
    $renderHandler = [System.EventHandler]{
        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency

        # Movement Parameters
        $flySpeed = 4.0 * $SyncHash.SpeedMultiplier
        $bobSpeed = 2.0
        $bobAmount = 1.5
        $distStep = $flySpeed * $elapsed
        $SyncHash.TotalDistance += $distStep

        # Update Lead Reindeer's (Deer1) Position
        $SyncHash.LeadPositionX -= ($flySpeed * $elapsed)
        
        # Reset if off-screen left
        if ($SyncHash.LeadPositionX -lt $SyncHash.LeftLimit) {
            # Calculate the distance of one full lap to shift the history
            $lapDistanceX = $SyncHash.RightLimit - $SyncHash.LeftLimit
            
            # Move the leader's logical position by one lap, creating a seamless loop
            $SyncHash.LeadPositionX += $lapDistanceX
            
            # Shift all historical X positions by the same amount to keep the train intact
            for ($i = 0; $i -lt $SyncHash.HistoryX.Count; $i++) {
                $SyncHash.HistoryX[$i] += $lapDistanceX
            }
        }

        # Update Y Position (Undulate)
        $SyncHash.LeadPositionY = [Math]::Sin($totalTime * $bobSpeed) * $bobAmount

        # Store current lead position and cumulative distance in history
        $SyncHash.HistoryX.Add($SyncHash.LeadPositionX)
        $SyncHash.HistoryY.Add($SyncHash.LeadPositionY)
        $SyncHash.HistoryDist.Add($SyncHash.TotalDistance)

        # Prune history (keep enough for the sleigh, ~20 units)
        while ($SyncHash.HistoryDist.Count -gt 0 -and ($SyncHash.TotalDistance - $SyncHash.HistoryDist[0] -gt 25.0)) {
            $SyncHash.HistoryX.RemoveAt(0)
            $SyncHash.HistoryY.RemoveAt(0)
            $SyncHash.HistoryDist.RemoveAt(0)
        }

        # Update positions of all objects based on distance lag
        # Spacing: Deer Width=3 -> Gap=1 -> Lag=4. Sleigh Width=5 -> Gap=1.7 -> Lag=5.7
        $lags = @(0, 4, 8, 12, 17.7)
        $objects = @("Deer1", "Deer2", "Deer3", "Deer4", "Sleigh")
        
        for ($i = 0; $i -lt $objects.Count; $i++) {
            $playerState = $SyncHash.PlayerStates[$objects[$i]]
            if (-not $playerState) { continue }

            $targetDist = $SyncHash.TotalDistance - $lags[$i]
            
            # Find history point closest to target distance (searching backwards)
            $idx = -1
            for ($j = $SyncHash.HistoryDist.Count - 1; $j -ge 0; $j--) {
                if ($SyncHash.HistoryDist[$j] -le $targetDist) {
                    $idx = $j
                    break
                }
            }

            if ($idx -ge 0) {
                $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[$idx]
                $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[$idx]
            } elseif ($SyncHash.HistoryX.Count -gt 0) {
                # Extrapolate if history isn't long enough (at start)
                # Since X decreases as Dist increases, we add the missing dist to X
                $playerState.TranslateTransform.OffsetX = $SyncHash.HistoryX[0] + ($SyncHash.HistoryDist[0] - $targetDist)
                $playerState.TranslateTransform.OffsetY = $SyncHash.HistoryY[0]
            }
        }
    }
    [System.Windows.Media.CompositionTarget]::add_Rendering($renderHandler)

    # --- UI Event Handlers ---
    $SyncHash.closeButton = $window.FindName("closeButton"); $SyncHash.closeButton.Add_Click({ $window.Close() })
    $SyncHash.redoButton = $window.FindName("redoButton"); $SyncHash.redoButton.Add_Click({ $SyncHash.RedoClicked = $true; $window.Close() })
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton"); $SyncHash.hideControlsButton.Add_Click({
        $controlsPanel = $window.FindName("controlsPanel")
        $SyncHash.ControlsHidden = -not $SyncHash.ControlsHidden
        $controlsPanel.Visibility = if ($SyncHash.ControlsHidden) { 'Collapsed' } else { 'Visible' }
    })
    
    $SyncHash.pauseButton = $window.FindName("pauseButton"); $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        $SyncHash.pauseButton.Content = if ($SyncHash.Paused) { "Resume" } else { "Pause" }
        if (-not $SyncHash.Paused) { $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp() }
    })

    $changeSpeed = { param($multiplier) $SyncHash.SpeedMultiplier /= $multiplier }
    $SyncHash.slowDownButton = $window.FindName("slowDownButton"); $SyncHash.slowDownButton.Add_Click({ & $changeSpeed 2.0 })
    $SyncHash.speedUpButton = $window.FindName("speedUpButton"); $SyncHash.speedUpButton.Add_Click({ & $changeSpeed 0.5 })

    $window.Add_KeyDown({
        param($s, $e)
        switch ($e.Key) {
            "Escape" { $window.Close() }
            "P" { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "R" { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
            "H" { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
        }
    })

    $window.Add_Closed({
        [System.Windows.Media.CompositionTarget]::remove_Rendering($renderHandler)
        foreach ($key in $SyncHash.PlayerStates.Keys) {
            $ps = $SyncHash.PlayerStates[$key]
            if ($ps.MediaTimer) { $ps.MediaTimer.Stop() }
            if ($ps.CurrentMediaElement) { $ps.CurrentMediaElement.Close() }
        }
    })

    $null = $window.ShowDialog()

    if (-not $SyncHash.RedoClicked) { break }
}
