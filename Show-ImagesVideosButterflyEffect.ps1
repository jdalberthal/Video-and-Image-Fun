<#
.SYNOPSIS
    Displays multiple, independently moving 3D planes, each showing a looping playlist of images and videos selected by the user.
.DESCRIPTION
    This script first presents a dialog to select image and video files. Once files are selected, it creates a dynamic visual display featuring six 3D planes that move around the screen in a butterfly-like pattern. Each face of each plane independently plays media from the user-selected playlist.

    The script uses FFmpeg for robust video decoding, allowing a wide variety of video formats to be displayed in real-time. The 3D planes are animated procedurally to bounce off the edges of the screen. Users can interact by clicking on the background to randomize the rotation axis of all planes.
.EXAMPLE
    .\SpinningDPlane.ps1
    Launches the file selection dialog. After selection, it launches the WPF window and begins the animation.
.NOTES
    Name:           SpinningDPlane.ps1
    Version:        1.0.0, 10/23/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffplay.exe): https://www.ffmpeg.org/download.html
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName WindowsFormsIntegration, System.Xaml
[System.Windows.Forms.Application]::EnableVisualStyles()

$ExternalButtonName = "Spinning Media Planes"
$ScriptDescription = "Displays multiple, independently moving 3D planes, each showing a looping playlist of images and videos."
$RequiredExecutables = @("ffmpeg.exe", "ffplay.exe")

$SyncHash = [hashtable]::Synchronized(@{}) # For passing data between runspaces
$imageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")

#region --- Animation Functions ---

function Update-RandomMotion {
    param($SyncHash)
    # This function now only needs to update the rotation axis, as the storyboard handles the angle animation.
    for ($i = 1; $i -le 6; $i++)
    {
        $rotationTransform = $SyncHash.Window.FindName("rotation$i")
        if ($rotationTransform)
        {
            $randX = Get-Random -Minimum -1.0 -Maximum 1.0
            $randY = Get-Random -Minimum -1.0 -Maximum 1.0
            $rotationAxis = New-Object System.Windows.Media.Media3D.Vector3D($randX, $randY, 0) # Keep rotation on XY plane
            if ($rotationAxis.Length -eq 0) { $rotationAxis.Y = 1 } # Default to Y-axis
            $rotationTransform.Axis = $rotationAxis
        }
    }
}

function Stop-MediaResources {
    param(
        [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]$Timers,
        [System.Collections.Generic.List[System.Diagnostics.Process]]$Processes,
        [System.Collections.Generic.List[System.Management.Automation.PowerShell]]$Runspaces
    )
    foreach ($timer in $Timers) { $timer.Stop() }
    $Timers.Clear()

    foreach ($proc in $Processes)
    {
        if (-not $proc.HasExited) { $proc.Kill() }
    }
    $Processes.Clear()

    foreach ($ps in $Runspaces) { $ps.Stop(); $ps.Dispose() }
    $Runspaces.Clear()
}

function Get-NextPlaylistIndex {
    param($SyncHash)
    $playlistCount = $SyncHash.playlist.Count
    if ($playlistCount -eq 0) { return -1 }

    $nextIndex = -1
    [System.Threading.Monitor]::Enter($SyncHash.SyncRoot)
    try {
        $nextIndex = $SyncHash.FileCounter
        $SyncHash.FileCounter = ($SyncHash.FileCounter + 1) % $playlistCount
    }
    finally {
        [System.Threading.Monitor]::Exit($SyncHash.SyncRoot)
    }
    return $nextIndex
}

function Show-Video {
    param(
        $SyncHash,
        [string]$FilePath,
        [int]$PlaneIndex,
        [System.Windows.Controls.Image]$ImageControl,
        [scriptblock]$OnVideoEnd,
        [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]$Timers,
        [System.Collections.Generic.List[System.Diagnostics.Process]]$Processes,
        [System.Collections.Generic.List[System.Management.Automation.PowerShell]]$Runspaces
    )
    $width = 640
    $height = 480
    $frameSize = $width * $height * 3
    $bitmap = [Windows.Media.Imaging.WriteableBitmap]::new($width, $height, 96, 96, [Windows.Media.PixelFormats]::Bgr24, $null)
    $ImageControl.Source = $bitmap

    $ffmpeg_args = "-hide_banner -loglevel error -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "ffmpeg.exe"
    $psi.Arguments = $ffmpeg_args
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    try {
        $proc = [System.Diagnostics.Process]::Start($psi)
        $Processes.Add($proc)
    }
    catch { # This catch block will handle if ffmpeg fails to start
        Write-Error "Failed to start ffmpeg.exe. Ensure it is in your PATH. Error: $($_.Exception.Message)"
        return
    }

    $stream = $proc.StandardOutput.BaseStream
    $rect = [System.Windows.Int32Rect]::new(0, 0, $width, $height)
    $stride = $width * 3

    # If the process exited immediately (e.g., file not found, invalid format), handle it gracefully.
    if ($proc.HasExited) {
        Write-Warning "ffmpeg process for '$FilePath' exited unexpectedly. The file may be invalid or inaccessible."
        & $OnVideoEnd -SyncHash $SyncHash -PlaneIndex $PlaneIndex
        return
    }

    $ps = [PowerShell]::Create()
    $null = $ps.AddScript({
        param($stream, $frameSize, $frameQueue)
        while ($true) {
            $frameBytes = New-Object byte[] $frameSize
            $totalRead = 0
            while ($totalRead -lt $frameSize) {
                try {
                    $bytesRead = $stream.Read($frameBytes, $totalRead, $frameSize - $totalRead)
                    if ($bytesRead -le 0) {
                        $frameQueue.Enqueue($null)
                        return
                    }
                    $totalRead += $bytesRead
                }
                catch { return }
            }
            $frameQueue.Enqueue($frameBytes)
        }
    })

    $frameQueue = [System.Collections.Concurrent.ConcurrentQueue[byte[]]]::new()
    $null = $ps.AddParameters(@($stream, $frameSize, $frameQueue))
    $runspaceHandle = $ps.BeginInvoke()
    $Runspaces.Add($ps)

    $uiTimer = New-Object Windows.Threading.DispatcherTimer
    $uiTimer.Interval = [TimeSpan]::FromMilliseconds(33)
    $Timers.Add($uiTimer)

    $currentPlaneIndex = $PlaneIndex
    $tickScriptBlock = {
        $frame = [byte[]]$null
        if ($frameQueue.TryDequeue([ref]$frame)) {
            if ($frame) {
                $bitmap.WritePixels($rect, $frame, $stride, 0)
            } else {
                $uiTimer.Stop() # Stop this timer
                & $OnVideoEnd -SyncHash $SyncHash -PlaneIndex $currentPlaneIndex
            }
        } elseif ($runspaceHandle.IsCompleted -and $frameQueue.IsEmpty) { # Fallback for when the stream ends
            $uiTimer.Stop()
            & $OnVideoEnd -SyncHash $SyncHash -PlaneIndex $currentPlaneIndex
        }
    }
    $uiTimer.Add_Tick($tickScriptBlock.GetNewClosure())
    $uiTimer.Start()
}

function Show-Image {
    param(
        $SyncHash,
        [string]$FilePath,
        [int]$PlaneIndex,
        [System.Windows.Controls.Image]$ImageControl,
        [scriptblock]$OnImageEnd,
        [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]$Timers,
        [int]$imageDisplaySeconds
    )
    try {
        $bitmapImage = [Windows.Media.Imaging.BitmapImage]::new([Uri]$FilePath)
        $ImageControl.Source = $bitmapImage
    }
    catch {
        Write-Warning "Failed to load image: $($FilePath)."
        & $OnImageEnd -SyncHash $SyncHash -PlaneIndex $PlaneIndex
        return
    }

    $imageTimer = New-Object Windows.Threading.DispatcherTimer
    $imageTimer.Interval = [TimeSpan]::FromSeconds($imageDisplaySeconds)
    $Timers.Add($imageTimer)

    $currentPlaneIndex = $PlaneIndex
    # Use a scriptblock to ensure the timer stops itself and then calls the next function
    $tickScriptBlock = {
        $imageTimer.Stop()
        & $OnImageEnd -SyncHash $SyncHash -PlaneIndex $currentPlaneIndex
    }
    $imageTimer.Add_Tick($tickScriptBlock.GetNewClosure())
    $imageTimer.Start()
}

function Start-NextMediaItemForFront {
    param($SyncHash, $PlaneIndex)
    Stop-MediaResources -Timers $SyncHash."Plane${PlaneIndex}FrontTimers" -Processes $SyncHash."Plane${PlaneIndex}FrontProcesses" -Runspaces $SyncHash."Plane${PlaneIndex}FrontRunspaces"
    $playlist = $SyncHash.playlist
    if ($playlist.Count -eq 0) { Write-Warning "Playlist is empty."; return }

    $nextIndex = Get-NextPlaylistIndex -SyncHash $SyncHash
    if ($nextIndex -lt 0) { return }
    $filePath = $playlist[$nextIndex]
    $FrontImageControl = $SyncHash."Plane${PlaneIndex}FrontImageControl"

    # Update text if Filename overlay is active
    if ($SyncHash.RbSelection -eq "Filename") {
        $textBlock = $SyncHash."Plane${PlaneIndex}FrontTextBlock"
        if ($textBlock) {
            try {
                $textBlock.Text = (Split-Path -Path $filePath -Leaf)
            } catch {
                $textBlock.Text = ""
            }
        }
    }

    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

    if ($SyncHash.imageExtensions -contains $extension) {
        Show-Image -SyncHash $SyncHash -FilePath $filePath -ImageControl $FrontImageControl -PlaneIndex $PlaneIndex `
            -OnImageEnd ${function:Start-NextMediaItemForFront} -Timers $SyncHash."Plane${PlaneIndex}FrontTimers" -imageDisplaySeconds $SyncHash.imageDisplaySeconds
    } else {
        Show-Video -SyncHash $SyncHash -FilePath $filePath -ImageControl $FrontImageControl -PlaneIndex $PlaneIndex `
            -OnVideoEnd ${function:Start-NextMediaItemForFront} -Timers $SyncHash."Plane${PlaneIndex}FrontTimers" `
            -Processes $SyncHash."Plane${PlaneIndex}FrontProcesses" -Runspaces $SyncHash."Plane${PlaneIndex}FrontRunspaces"
    }
}

function Start-NextMediaItemForBack {
    param($SyncHash, $PlaneIndex)
    Stop-MediaResources -Timers $SyncHash."Plane${PlaneIndex}BackTimers" -Processes $SyncHash."Plane${PlaneIndex}BackProcesses" -Runspaces $SyncHash."Plane${PlaneIndex}BackRunspaces"
    $playlist = $SyncHash.playlist
    if ($playlist.Count -eq 0) { Write-Warning "Playlist is empty."; return }

    $nextIndex = Get-NextPlaylistIndex -SyncHash $SyncHash
    if ($nextIndex -lt 0) { return }
    $filePath = $playlist[$nextIndex]
    $BackImageControl = $SyncHash."Plane${PlaneIndex}BackImageControl"

    # Update text if Filename overlay is active
    if ($SyncHash.RbSelection -eq "Filename") {
        $textBlock = $SyncHash."Plane${PlaneIndex}BackTextBlock"
        if ($textBlock) {
            try {
                $textBlock.Text = (Split-Path -Path $filePath -Leaf)
            } catch {
                $textBlock.Text = ""
            }
        }
    }

    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

    if ($SyncHash.imageExtensions -contains $extension) {
        Show-Image -SyncHash $SyncHash -FilePath $filePath -ImageControl $BackImageControl -PlaneIndex $PlaneIndex `
            -OnImageEnd ${function:Start-NextMediaItemForBack} -Timers $SyncHash."Plane${PlaneIndex}BackTimers" -imageDisplaySeconds $SyncHash.imageDisplaySeconds
    } else {
        Show-Video -SyncHash $SyncHash -FilePath $filePath -ImageControl $BackImageControl -PlaneIndex $PlaneIndex `
            -OnVideoEnd ${function:Start-NextMediaItemForBack} -Timers $SyncHash."Plane${PlaneIndex}BackTimers" `
            -Processes $SyncHash."Plane${PlaneIndex}BackProcesses" -Runspaces $SyncHash."Plane${PlaneIndex}BackRunspaces"
    }
}

function Show-ImagesVideosButterflyEffect
{
    param(
        [hashtable]$SyncHash,
        [array]$playlist,
        [bool]$UseTransparentEffect
    )
    $SyncHash.Paused = $false
    $SyncHash.imageDisplaySeconds = 10
    $SyncHash.imageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")

    function Start-ButterflyMovement {
        param($WorkAreaWidth, $WorkAreaHeight)
        $edgeMargin = 150 # Increased margin to account for the plane's size during rotation
        $SyncHash.Planes = @()
    
        $startupTimer = New-Object Windows.Threading.DispatcherTimer
        $startupTimer.Interval = [TimeSpan]::FromMilliseconds(250)
        $SyncHash.Timers.Add($startupTimer)

        $startupTimer.Add_Tick({
                $startupTimer.Stop()
                $viewport = $SyncHash.Window.FindName("mainViewport")
    
                for ($i = 1; $i -le 6; $i++)
                {
                    $frontImageControl = $SyncHash.Window.FindName("videoImage$i")
                    $backImageControl = $SyncHash.Window.FindName("backImage$i")
                    $SyncHash."Plane${i}FrontImageControl" = $frontImageControl
                    $SyncHash."Plane${i}BackImageControl" = $backImageControl

                    $SyncHash."Plane${i}FrontTimers" = [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]::new()
                    $SyncHash."Plane${i}FrontProcesses" = [System.Collections.Generic.List[System.Diagnostics.Process]]::new()
                    $SyncHash."Plane${i}FrontRunspaces" = [System.Collections.Generic.List[System.Management.Automation.PowerShell]]::new()
                    $SyncHash."Plane${i}BackTimers" = [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]::new()
                    $SyncHash."Plane${i}BackProcesses" = [System.Collections.Generic.List[System.Diagnostics.Process]]::new()
                    $SyncHash."Plane${i}BackRunspaces" = [System.Collections.Generic.List[System.Management.Automation.PowerShell]]::new()
            
                    # --- New: Set a safe, random starting position ---
                    $startX = (Get-Random -Minimum -3.5 -Maximum 3.5)
                    $startY = (Get-Random -Minimum -2.5 -Maximum 2.5)
                    ($SyncHash.Window.FindName("translation$i")).OffsetX = $startX
                    ($SyncHash.Window.FindName("translation$i")).OffsetY = $startY

                     $plane = @{
                        Translation = $SyncHash.Window.FindName("translation$i")
                        ModelVisual = $SyncHash.Window.FindName("planeModelVisual$i")
                        VelocityX   = (Get-Random -Minimum 0.015 -Maximum 0.035) * (Get-Random @(1, -1))
                        VelocityY   = (Get-Random -Minimum 0.015 -Maximum 0.035) * (Get-Random @(1, -1))
                    }
                    if ($plane.Translation -and $plane.ModelVisual)
                    {
                        $SyncHash.Planes += $plane
                    }

                    if ($frontImageControl) { Start-NextMediaItemForFront -SyncHash $SyncHash -PlaneIndex $i }
                    if ($backImageControl) { Start-NextMediaItemForBack -SyncHash $SyncHash -PlaneIndex $i }
                }

                if ($SyncHash.Planes.Count -gt 0 -and $viewport)
                {   
                    $movementTimer = New-Object Windows.Threading.DispatcherTimer
                    $movementTimer.Interval = [TimeSpan]::FromMilliseconds(20)
                    $SyncHash.Timers.Add($movementTimer)
                    $SyncHash.movementTimer = $movementTimer # Store for pause/resume
            
                    $movementTimer.Add_Tick({
                            foreach ($plane in $SyncHash.Planes)
                            {   
                                $transformToRoot = $plane.ModelVisual.TransformToAncestor($viewport)
                                if (-not $transformToRoot) { continue }
    
                                $centerScreenPoint = $transformToRoot.Transform([System.Windows.Media.Media3D.Point3D]::new(0, 0, 0))
                                $cornerScreenPoint = $transformToRoot.Transform([System.Windows.Media.Media3D.Point3D]::new(-0.667, 0.5, 0))
                                $radius = [Math]::Sqrt([Math]::Pow($centerScreenPoint.X - $cornerScreenPoint.X, 2) + [Math]::Pow($centerScreenPoint.Y - $cornerScreenPoint.Y, 2)) + 10

                                # --- Predictive Bouncing Logic ---
                                # Predict the next position to prevent overshooting the boundary.
                                # We use a scaling factor (e.g., 200) to approximate the conversion from 3D velocity to 2D screen movement.
                                $predictedScreenX = $centerScreenPoint.X + ($plane.VelocityX * 200)
                                $predictedScreenY = $centerScreenPoint.Y + ($plane.VelocityY * 200)

                                if (($predictedScreenX + $radius) -ge $WorkAreaWidth -or ($predictedScreenX - $radius) -le 0) {
                                    $plane.VelocityX *= -1 # Reverse X velocity
                                }

                                if (($predictedScreenY + $radius) -ge $WorkAreaHeight -or ($predictedScreenY - $radius) -le 0) {
                                    $plane.VelocityY *= -1 # Reverse Y velocity
                                }

                                $plane.Translation.OffsetX += $plane.VelocityX
                                $plane.Translation.OffsetY += $plane.VelocityY
                            }
                        }.GetNewClosure())
                    $movementTimer.Start()
                }
            }.GetNewClosure())

        $startupTimer.Start()
    }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Spinning 3D Plane"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">
    <Grid Background="Transparent" IsHitTestVisible="True">
        <Grid.Triggers>
            <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                <BeginStoryboard>
                    <Storyboard x:Name="mainStoryboard">
                        <DoubleAnimation Storyboard.TargetName="rotation1" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                        <DoubleAnimation Storyboard.TargetName="rotation2" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                        <DoubleAnimation Storyboard.TargetName="rotation3" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                        <DoubleAnimation Storyboard.TargetName="rotation4" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                        <DoubleAnimation Storyboard.TargetName="rotation5" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                        <DoubleAnimation Storyboard.TargetName="rotation6" Storyboard.TargetProperty="Angle" From="0" To="360" Duration="0:0:10" RepeatBehavior="Forever" />
                    </Storyboard>
                </BeginStoryboard>
            </EventTrigger>
        </Grid.Triggers>
        <Border Name="viewportBorder" Background="Transparent">
            <Viewport3D Name="mainViewport">
            <Viewport3D.Camera>
                <PerspectiveCamera Position="0, 0, 10" LookDirection="0, 0, -1" />
            </Viewport3D.Camera>
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="#FF7F7F7F" />
                        <DirectionalLight Color="White" Direction="-1, -1, -2" />
                        <DirectionalLight Color="White" Direction="1, 1, 2" />
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual1">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation1" />
                       <RotateTransform3D>
                           <RotateTransform3D.Rotation>
                               <AxisAngleRotation3D x:Name="rotation1" Axis="0, 1, 0" Angle="0" />
                           </RotateTransform3D.Rotation>
                       </RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" />
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid Background="Transparent">
                            <Image Name="videoImage1" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront1" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform>
                            <RotateTransform3D>
                                <RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation>
                            </RotateTransform3D>
                        </Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" />
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid Background="Transparent">
                            <Image Name="backImage1" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack1" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual2">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation2" />
                       <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="rotation2" Axis="0, 1, 0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="videoImage2" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront2" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="backImage2" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack2" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual3">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation3" />
                       <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="rotation3" Axis="0, 1, 0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="videoImage3" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront3" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="backImage3" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack3" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual4">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation4" />
                       <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="rotation4" Axis="0, 1, 0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="videoImage4" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront4" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="backImage4" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack4" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual5">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation5" />
                       <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="rotation5" Axis="0, 1, 0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="videoImage5" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront5" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="backImage5" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack5" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual6">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation6" />
                       <RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D x:Name="rotation6" Axis="0, 1, 0" Angle="0" /></RotateTransform3D.Rotation></RotateTransform3D>
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="videoImage6" Stretch="Fill" />
                            <TextBlock Name="textOverlayFront6" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D>
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D Positions="-0.667,-0.5,0  0.667,-0.5,0  0.667,0.5,0  -0.667,0.5,0" TriangleIndices="0,1,2  0,2,3" TextureCoordinates="0,1 1,1 1,0 0,0" /></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent">
                            <Image Name="backImage6" Stretch="Fill" />
                            <TextBlock Name="textOverlayBack6" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            </Viewport3D>
        </Border>
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

    try
    {
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
    }
    catch
    {
        Write-Error "Failed to load XAML: $($_.Exception.Message)"
        return
    }

    if ($UseTransparentEffect) {
        # Find all Viewport2DVisual3D elements. This is simpler than naming them all.
        $allViewports = @($window.FindName("mainViewport").Children | Where-Object { $_ -is [System.Windows.Media.Media3D.ModelVisual3D] } | ForEach-Object { $_.Children } | Where-Object { $_ -is [System.Windows.Media.Media3D.Viewport2DVisual3D] })
        
        foreach ($viewport in $allViewports) {
            # Create a new EmissiveMaterial.
            $emissiveMaterial = New-Object System.Windows.Media.Media3D.EmissiveMaterial
            $emissiveMaterial.Brush = [System.Windows.Media.Brushes]::White # Keep the brush
            
            # This is the crucial part: setting the attached property in code.
            [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($emissiveMaterial, $true)
            
            # Assign the new material to the viewport.
            $viewport.Material = $emissiveMaterial
        }
    }

    $SyncHash.playlist = $playlist
    $SyncHash.FileCounter = 0

    $SyncHash.Timers = [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]::new()
    $SyncHash.Processes = [System.Collections.Generic.List[System.Diagnostics.Process]]::new()
    $SyncHash.Runspaces = [System.Collections.Generic.List[System.Management.Automation.PowerShell]]::new()

    $SyncHash.Window = $window
    $SyncHash.mainStoryboard = $window.FindName("mainStoryboard")

    # Store controls for easy access
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.randomAxisButton = $window.FindName("randomAxisButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")

    # Store animations for pause/resume
    $SyncHash.Animations = @{}
    foreach ($animation in $SyncHash.mainStoryboard.Children) {
        $targetName = [System.Windows.Media.Animation.Storyboard]::GetTargetName($animation)
        $SyncHash.Animations[$targetName] = $animation
    }
    
    # Store TextBlocks for overlay
    for ($i = 1; $i -le 6; $i++) {
        $SyncHash."Plane${i}FrontTextBlock" = $SyncHash.Window.FindName("textOverlayFront$i")
        $SyncHash."Plane${i}BackTextBlock" = $SyncHash.Window.FindName("textOverlayBack$i")
    }

    $closeButton = $SyncHash.Window.FindName('closeButton')
    Add-Type -AssemblyName System.Windows.Forms
    $PrimaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $WorkAreaWidth = $PrimaryScreen.WorkingArea.Width
    $WorkAreaHeight = $PrimaryScreen.WorkingArea.Height

    $SyncHash.Window.Width = [System.Windows.SystemParameters]::WorkArea.Width
    $SyncHash.Window.Height = [System.Windows.SystemParameters]::WorkArea.Height
    $SyncHash.Window.Left = 0
    $SyncHash.Window.Top = 0

    # --- Text Overlay Logic ---
    $applyTextStyles = {
        param($textBlock)
        if (-not $textBlock) { return }
        $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
        $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
        $textBlock.FontFamily = $SyncHash.SelectedFont
        $textBlock.FontSize = $SyncHash.SelectedFontSize
        $textBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
        $textBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
    }

    switch ($SyncHash.RbSelection) {
        "Hidden" {
            # Text is hidden by default, nothing to do.
        }
        "Filename" {
            # Apply styles to all text blocks initially
            for ($i = 1; $i -le 6; $i++) {
                & $applyTextStyles -textBlock $SyncHash."Plane${i}FrontTextBlock"
                & $applyTextStyles -textBlock $SyncHash."Plane${i}BackTextBlock"
            }
            # Set initial filenames if that option is selected
            for ($i = 1; $i -le 6; $i++) {
                $frontIndex = ($i - 1)
                $backIndex = ($i - 1) + 6
                if ($frontIndex -lt $SyncHash.playlist.Count) {
                    $filePath = $SyncHash.playlist[$frontIndex]
                    $SyncHash."Plane${i}FrontTextBlock".Text = (Split-Path -Path $filePath -Leaf)
                }
                if ($backIndex -lt $SyncHash.playlist.Count) {
                    $filePath = $SyncHash.playlist[$backIndex]
                    $SyncHash."Plane${i}BackTextBlock".Text = (Split-Path -Path $filePath -Leaf)
                }
            }
        }
        "Custom" {
            $applyTextStyles = {
                param($textBlock)
                if (-not $textBlock) { return }
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $textBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $textBlock.FontFamily = $SyncHash.SelectedFont
                $textBlock.FontSize = $SyncHash.SelectedFontSize
                $textBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
                $textBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
            }
            $customText = $SyncHash.TextBox.Text

            for ($i = 1; $i -le 6; $i++) {
                $frontTextBlock = $SyncHash."Plane${i}FrontTextBlock"
                $backTextBlock = $SyncHash."Plane${i}BackTextBlock"
                if ($frontTextBlock) { & $applyTextStyles -textBlock $frontTextBlock; $frontTextBlock.Text = $customText }
                if ($backTextBlock) { & $applyTextStyles -textBlock $backTextBlock; $backTextBlock.Text = $customText }
            }
        }
    }

    # --- Keyboard and Button Events ---
    $SyncHash.Window.Add_KeyDown({
            param($sender, $e)
            switch ($e.Key) {
                'Escape' { $SyncHash.Window.Close() }
                'P' { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'A' { $SyncHash.randomAxisButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'H' { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'R' { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'Left' { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'Right' { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'F1' {
                    $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                    $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
                    $OkButton = $PopupWindow.FindName("OKButton")
                    $OkButton.Add_Click({ $PopupWindow.Close() })
                    $PopupWindow.ShowDialog() | Out-Null
                }
            }
        })

    $SyncHash.pauseButton.Add_Click({
            if ($SyncHash.Paused) {
                # Resume: Restart animations from their last saved angle
                foreach ($targetName in $SyncHash.Animations.Keys) {
                    $rotation = $SyncHash.Window.FindName($targetName)
                    $animation = $SyncHash.Animations[$targetName]
                    $animation.From = $rotation.Angle # Start from the saved angle
                    $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animation)
                }
                if ($SyncHash.movementTimer) { $SyncHash.movementTimer.Start() }
                $SyncHash.pauseButton.Content = "Pause"
                $SyncHash.Paused = $false
            } else {
                # Pause: Stop animations and save their current angle
                foreach ($targetName in $SyncHash.Animations.Keys) {
                    $rotation = $SyncHash.Window.FindName($targetName)
                    $currentAngle = $rotation.Angle
                    $rotation.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $null) # Stop animation
                    $rotation.Angle = $currentAngle # Set the angle explicitly
                }
                if ($SyncHash.movementTimer) { $SyncHash.movementTimer.Stop() }
                $SyncHash.pauseButton.Content = "Resume"
                $SyncHash.Paused = $true
            }
        })

    $SyncHash.randomAxisButton.Add_Click({
            Update-RandomMotion -SyncHash $SyncHash
        })

    $SyncHash.hideControlsButton.Add_Click({
            $controlsPanel = $SyncHash.Window.FindName("controlsPanel")
            if ($controlsPanel.Visibility -eq 'Visible') {
                $controlsPanel.Visibility = 'Collapsed'
            } else {
                $controlsPanel.Visibility = 'Visible'
            }
        })

    $SyncHash.redoButton.Add_Click({
            $SyncHash.Window.Close()
            $SelectFolderForm.Show()
        })

    $changeSpeed = {
        param($multiplier)
        foreach ($animation in $SyncHash.mainStoryboard.Children) {
            $currentDuration = $animation.Duration.TimeSpan.TotalSeconds
            $newDuration = $currentDuration * $multiplier
            if ($newDuration -lt 0.5) { $newDuration = 0.5 } # Prevent it from going too fast
            if ($newDuration -gt 600) { $newDuration = 600 } # Prevent it from going too slow
            $animation.Duration = [System.Windows.Duration]::new([TimeSpan]::FromSeconds($newDuration))
        }
        # Restart the storyboard to apply the new durations
        $SyncHash.mainStoryboard.Begin($SyncHash.Window, $true)
        if ($SyncHash.Paused) { $SyncHash.mainStoryboard.Pause() }
    }

    $SyncHash.slowDownButton.Add_Click({
            & $changeSpeed 2.0 # Double the duration to slow down
        })

    $SyncHash.speedUpButton.Add_Click({
            & $changeSpeed 0.5 # Halve the duration to speed up
        })

    $SyncHash.Window.Add_Closed({
            param($sender, $e)
            # Stop all timers and kill all processes/runspaces associated with each plane
            for ($i = 1; $i -le 6; $i++) {
                Stop-MediaResources -Timers $SyncHash."Plane${i}FrontTimers" -Processes $SyncHash."Plane${i}FrontProcesses" -Runspaces $SyncHash."Plane${i}FrontRunspaces"
                Stop-MediaResources -Timers $SyncHash."Plane${i}BackTimers" -Processes $SyncHash."Plane${i}BackProcesses" -Runspaces $SyncHash."Plane${i}BackRunspaces"
            }

            # Stop the main movement and startup timers
            foreach ($timer in $SyncHash.Timers) { $timer.Stop() }
            $SyncHash.Timers.Clear()

            # As a final safety measure, ensure no ffmpeg processes from this script are left running.
            # This is a good practice in case a process was not tracked correctly.
            Get-Process "ffmpeg" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

            # Clear the main storyboard to stop all animations
            $SyncHash.mainStoryboard.Stop()
        })

    $closeButton.Add_Click({ $SyncHash.Window.Close() })

    $border = $SyncHash.Window.FindName('viewportBorder')
    $border.Add_MouseDown({
            param($sender, $e)
            $viewport = $SyncHash.Window.FindName('mainViewport')
            $mousePosition = $e.GetPosition($viewport)
            $SyncHash.hitModel = $null
            $hitTestCallback = [System.Windows.Media.HitTestResultCallback]{
                param($result)
                if ($result -is [System.Windows.Media.Media3D.RayMeshGeometry3DHitTestResult])
                {
                    $SyncHash.hitModel = $result.ModelHit
                    return [System.Windows.Media.HitTestResultBehavior]::Stop
                }
                return [System.Windows.Media.HitTestResultBehavior]::Continue
            }
            $hitTestParams = [System.Windows.Media.PointHitTestParameters]::new($mousePosition)
            [System.Windows.Media.VisualTreeHelper]::HitTest($viewport, $null, $hitTestCallback, $hitTestParams)

            if ($SyncHash.hitModel -is [System.Windows.Media.Media3D.GeometryModel3D])
            {
                # Trigger the pause/resume functionality, just like the 'P' key or pause button
                $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        })

    Update-RandomMotion -SyncHash $SyncHash
    Start-ButterflyMovement -WorkAreaWidth $WorkAreaWidth -WorkAreaHeight $WorkAreaHeight
    $null = $window.ShowDialog()
}

#endregion

#region --- Help Popup XAML ---
[xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help" Height="320" Width="400" WindowStartupLocation="CenterScreen" WindowStyle="ToolWindow">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <RichTextBox Grid.Row="0" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
            <FlowDocument>
                <Paragraph>
                    <Run Text="Animation Controls:" FontWeight="Bold"/><LineBreak/>
                </Paragraph>
                <Paragraph TextAlignment="Left" FontFamily="Consolas">
                    <Bold>
                        <Run Text="Key         : Action" TextDecorations="Underline"/><LineBreak/>
                    </Bold>
                    <Run Text="Esc         : Exit Application"/><LineBreak/>
                    <Run Text="P           : Pause / Resume Spinning"/><LineBreak/>
                    <Run Text="R           : Reselect Media Files (Redo)"/><LineBreak/>
                    <Run Text="A           : Change Rotation Axis (All Planes)"/><LineBreak/>
                    <Run Text="H           : Hide / Show Controls"/><LineBreak/>
                    <Run Text="&#x2190; (Left)    : Slow Down Spinning"/><LineBreak/>
                    <Run Text="&#x2192; (Right)   : Speed Up Spinning"/><LineBreak/><LineBreak/>
                    <Run Text="*Click a plane to pause/resume."/><LineBreak/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@
#endregion

#region --- UI and Main Logic ---

# --- Dependency Check ---
if ($RequiredExecutables)
{
    $dependencyStatus = @()
    $allDependenciesMet = $true

    foreach ($exe in $RequiredExecutables)
    {
        $localPath = Join-Path $PSScriptRoot $exe
        if ((Get-Command $exe -ErrorAction SilentlyContinue) -or (Test-Path -Path $localPath))
        {
            $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'Found' }
        }
        else
        {
            $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'NOT FOUND' }
            $allDependenciesMet = $false
        }
    }

    if (-not $allDependenciesMet)
    {
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
}

# Create the form
$SelectFolderForm = New-Object System.Windows.Forms.Form
$SelectFolderForm.Text = "Spinning Planes - Media Selector"
$SelectFolderForm.Size = New-Object System.Drawing.Size(800, 680)
$SelectFolderForm.StartPosition = "CenterScreen"

# Create a Button for browsing folders
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Browse Folder"
$BrowseButton.Location = New-Object System.Drawing.Point(10, 10)
$BrowseButton.Size = New-Object System.Drawing.Size(100, 25)
$SelectFolderForm.Controls.Add($BrowseButton)

# Create a TextBox to display the selected folder path
$FolderPathTextBox = New-Object System.Windows.Forms.TextBox
$FolderPathTextBox.Location = New-Object System.Drawing.Point(120, 10)
$FolderPathTextBox.Size = New-Object System.Drawing.Size(450, 25)
$FolderPathTextBox.ReadOnly = $True
$SelectFolderForm.Controls.Add($FolderPathTextBox)

# Create a CheckBox for recursive scanning
$RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox
$RecursiveCheckBox.Text = "Include Subfolders"
$RecursiveCheckBox.AutoSize = $True
$RecursiveCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$RecursiveCheckBox.Size = New-Object System.Drawing.Size(150, 20)
$RecursiveCheckBox.Checked = $False
$SelectFolderForm.Controls.Add($RecursiveCheckBox)

# Create a CheckBox for material type
$TransparentFacesCheckbox = New-Object System.Windows.Forms.CheckBox
$TransparentFacesCheckbox.Text = "Make Semi-Transparent"
$TransparentFacesCheckbox.AutoSize = $True
$TransparentFacesCheckbox.Location = New-Object System.Drawing.Point(150, 40)
$TransparentFacesCheckbox.Checked = $False
$SelectFolderForm.Controls.Add($TransparentFacesCheckbox)

# Create a DataGridView to display files
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(10, 95)
$DataGridView.Size = New-Object System.Drawing.Size(760, 330)
$DataGridView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$DataGridView.AutoGenerateColumns = $False # We'll define columns manually
$DataGridView.AllowUserToAddRows = $False
$DataGridView.RowHeadersWidth = 65
$SelectFolderForm.Controls.Add($DataGridView)

# Add a CheckBox column to the DataGridView
$CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$CheckBoxColumn.Name = "Select"
$CheckBoxColumn.HeaderText = ""
$CheckBoxColumn.Width = 30
$DataGridView.Columns.Add($CheckBoxColumn) | Out-Null

# Add a column for file names
$FileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$FileNameColumn.Name = "FileName"
$FileNameColumn.HeaderText = "File Name"
$FileNameColumn.Width = 300
$FileNameColumn.ReadOnly = $True
$DataGridView.Columns.Add($FileNameColumn) | Out-Null

# Add a column for full file paths
$FilePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$FilePathColumn.Name = "FilePath"
$FilePathColumn.HeaderText = "File Path"
$FilePathColumn.Width = 430
$FilePathColumn.ReadOnly = $True
$DataGridView.Columns.Add($FilePathColumn) | Out-Null

$SelectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$SelectAllCheckbox.Text = "Select All"
$SelectAllCheckbox.Location = New-Object System.Drawing.Point(10, 70)
$SelectAllCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$SelectAllCheckbox.Checked = $False
$SelectFolderForm.Controls.Add($SelectAllCheckbox)

# Create a Button to perform an action on selected files
$PlayButton = New-Object System.Windows.Forms.Button
$PlayButton.Text = "Play Selected Item(s)"
$PlayButton.Location = New-Object System.Drawing.Point(600, 40)
$PlayButton.Size = New-Object System.Drawing.Size(170, 30)
$SelectFolderForm.Controls.Add($PlayButton)

#region --- Text Overlay and F1 Help UI (Copied from Cube Script) ---

$MyFont = New-Object System.Drawing.Font("Arial", 12)

# Create a GroupBox to contain the radio buttons
$GroupBox = New-Object System.Windows.Forms.GroupBox
$GroupBox.Text = "Text Overlay"
$GroupBox.Location = New-Object System.Drawing.Point(10, 440)
$GroupBox.Size = New-Object System.Drawing.Size(125, 130)

# Create the three Radio Buttons
$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Text = "Hide Text Overlay"
$RadioButton1.Location = New-Object System.Drawing.Point(10, 30)
$RadioButton1.Width = 114
$RadioButton1.Checked = $True # Set one as default selected

$RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Text = "Filename"
$RadioButton2.Location = New-Object System.Drawing.Point(10, 60)

$RadioButton3 = New-Object System.Windows.Forms.RadioButton
$RadioButton3.Text = "Custom Text"
$RadioButton3.Location = New-Object System.Drawing.Point(10, 90)

# Add the Radio Buttons to the GroupBox
$GroupBox.Controls.Add($RadioButton1)
$GroupBox.Controls.Add($RadioButton2)
$GroupBox.Controls.Add($RadioButton3)

# Add the GroupBox to the form
$SelectFolderForm.Controls.Add($GroupBox)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Point(140, 440)
$TextBox.Size = New-Object System.Drawing.Size(455, 180)
$TextBox.Multiline = $True
$TextBox.Visible = $False
$TextBox.ScrollBars = "Vertical"
$TextBox.Font = $MyFont
$SelectFolderForm.Controls.Add($TextBox)
$SyncHash.TextBox = $TextBox

$SyncHash.TextColor = [PSCustomObject]@{ A = 255; R = 0; G = 0; B = 0 } # Default to Black

$CurrentColor = New-Object System.Windows.Forms.Label
$CurrentColor.Text = "Text Color:"
$CurrentColor.Location = New-Object System.Drawing.Point(600, 477)
$CurrentColor.AutoSize = $True
$CurrentColor.Visible = $False
$SelectFolderForm.Controls.Add($CurrentColor)

$ColorExample = New-Object System.Windows.Forms.Label
$ColorExample.Text = "     "
$ColorExample.Location = New-Object System.Drawing.Point(660, 477)
$ColorExample.AutoSize = $True
$ColorExample.BackColor = [System.Drawing.Color]::Black
$ColorExample.Visible = $False
$SelectFolderForm.Controls.Add($ColorExample)

$SelectColorButton = New-Object System.Windows.Forms.Button
$SelectColorButton.Text = "Change"
$SelectColorButton.Location = New-Object System.Drawing.Point(685, 470)
$SelectColorButton.Size = New-Object System.Drawing.Size(80, 30)
$SelectColorButton.Visible = $False
$SelectFolderForm.Controls.Add($SelectColorButton)

$SizeLabel = New-Object System.Windows.Forms.Label
$SizeLabel.Text = "Font Size:"
$SizeLabel.AutoSize = $True
$SizeLabel.Location = New-Object System.Drawing.Point(600, 522)
$SizeLabel.Visible = $False
$SelectFolderForm.Controls.Add($SizeLabel)

$NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$NumericUpDown.Location = New-Object System.Drawing.Point(660, 520)
$NumericUpDown.Size = New-Object System.Drawing.Size(50, 20)
$NumericUpDown.Visible = $False
$NumericUpDown.Minimum = 8
$NumericUpDown.Maximum = 72
$NumericUpDown.Value = 24 # Default font size
$SelectFolderForm.Controls.Add($NumericUpDown)
$SyncHash.NumericUpDown = $NumericUpDown
$SyncHash.SelectedFontSize = $SyncHash.NumericUpDown.Value

$FontButton = New-Object System.Windows.Forms.Button
$FontButton.Text = "Change Font"
$FontButton.Location = New-Object System.Drawing.Point(600, 570)
$FontButton.Size = New-Object System.Drawing.Size(170, 25)
$FontButton.Font = $MyFont
$FontButton.Visible = $False
$SelectFolderForm.Controls.Add($FontButton)
$SyncHash.SelectedFont = "Arial"
$SyncHash.FontButton = $FontButton

$ItalicCheckbox = New-Object System.Windows.Forms.CheckBox
$ItalicCheckbox.Text = "Italic"
$ItalicCheckbox.Location = New-Object System.Drawing.Point(600, 620)
$ItalicCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$ItalicCheckbox.Checked = $False
$ItalicCheckbox.Visible = $False
$SelectFolderForm.Controls.Add($ItalicCheckbox)
$SyncHash.ItalicCheckbox = $ItalicCheckbox

$BoldCheckbox = New-Object System.Windows.Forms.CheckBox
$BoldCheckbox.Text = "Bold"
$BoldCheckbox.Location = New-Object System.Drawing.Point(680, 620)
$BoldCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$BoldCheckbox.Checked = $True # Default to Bold
$BoldCheckbox.Visible = $False
$SelectFolderForm.Controls.Add($BoldCheckbox)
$SyncHash.BoldCheckbox = $BoldCheckbox

$HelpLabel = New-Object System.Windows.Forms.Label
$HelpLabel.Text = "F1 - Help"
$HelpLabel.AutoSize = $True
$HelpLabel.Location = New-Object System.Drawing.Point(700, 0)
$SelectFolderForm.Controls.Add($HelpLabel)

$Event = {
    $isTextVisible = $RadioButton2.Checked -or $RadioButton3.Checked
    $isCustomText = $RadioButton3.Checked

    $SyncHash.RbSelection = if ($RadioButton1.Checked) { "Hidden" } elseif ($RadioButton2.Checked) { "Filename" } else { "Custom" }

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

$RadioButton1.Add_Click($Event)
$RadioButton2.Add_Click($Event)
$RadioButton3.Add_Click($Event)

$NumericUpDown.Add_ValueChanged({ $SyncHash.SelectedFontSize = $NumericUpDown.Value })

$colorDialog = New-Object System.Windows.Forms.ColorDialog
$SelectColorButton.Add_Click({
    if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ColorExample.BackColor = $colorDialog.Color
        $TextBox.ForeColor = $colorDialog.Color
        $SyncHash.TextColor = $colorDialog.Color
    }
})

$FontButton.Add_Click({
    $fontDialog = New-Object System.Windows.Forms.FontDialog
    $fontDialog.Font = $TextBox.Font
    if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextBox.Font = $fontDialog.Font
        $FontButton.Text = $fontDialog.Font.Name
        $SyncHash.SelectedFont = $fontDialog.Font.Name
        $SyncHash.ItalicCheckbox.Checked = $fontDialog.Font.Italic
        $SyncHash.BoldCheckbox.Checked = $fontDialog.Font.Bold
        $SyncHash.NumericUpDown.Value = [Math]::Round($fontDialog.Font.Size)
    }
})

$updateFontStyle = {
    $style = [System.Drawing.FontStyle]::Regular
    if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
    if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
    $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, $style)
}
$ItalicCheckbox.Add_CheckedChanged($updateFontStyle)
$BoldCheckbox.Add_CheckedChanged($updateFontStyle)

$SelectFolderForm.KeyPreview = $True
$SelectFolderForm.Add_KeyDown({
    param($Sender, $e)
    if ($e.KeyCode -eq "F1") {
        $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
        $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
        $OkButton = $PopupWindow.FindName("OKButton")
        $OkButton.Add_Click({ $PopupWindow.Close() })
        $PopupWindow.ShowDialog() | Out-Null
    }
})

$dataGridView.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $row = $dataGridView.Rows[$e.RowIndex]
        $filePath = $row.Cells["FilePath"].Value
        if ([System.IO.File]::Exists($filePath)) {
            Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -autoexit -i `"$filePath`""
        } else {
            [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Set initial state of font controls
& $Event

#endregion

$SelectAllCheckbox.Add_CheckedChanged({
        $CheckedState = $SelectAllCheckbox.Checked
        foreach ($Row in $DataGridView.Rows)
        {
            $Row.Cells["Select"].Value = $CheckedState
        }
        $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    })
    
# Event handler for the Browse Folder button
$BrowseButton.Add_Click({
        $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
        "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
        "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
        "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
        "*.ts", "*.vob"

        $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
        $AllowedExtension = $VideoExtensions + $ImageExtensions

        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select the folder to scan."

        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $SelectedPath = $FolderBrowser.SelectedPath
            $FolderPathTextBox.Text = $SelectedPath
            $DataGridView.Rows.Clear()

            # Get files from the selected folder
            $Files = Get-ChildItem -Path "$($SelectedPath)\*" -File -Include $AllowedExtension

            if ($RecursiveCheckBox.Checked)
            {
                $Files = Get-ChildItem -Path $SelectedPath -File -Include $AllowedExtension -Recurse
            }

            foreach ($File in $Files)
            {
                $DataGridView.Rows.Add($False, $File.Name, $File.FullName) # Add a row with a checkbox (false initially), filename, and File Path
            }
        }
    })

# Event handler for the Process Selected Files button
$PlayButton.Add_Click({
        $selectedFiles = @()
        foreach ($Row in $DataGridView.Rows)
        {
            if ($Row.Cells["Select"].Value)
            {
                $selectedFiles += $Row.Cells["FilePath"].Value
            }
        }

        if($selectedFiles.Count -eq 0)
        {
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        else
        {
            $SelectFolderForm.Hide()
            $useTransparent = $TransparentFacesCheckbox.Checked
            Show-ImagesVideosButterflyEffect -SyncHash $SyncHash -playlist $selectedFiles -UseTransparentEffect $useTransparent
        }
    })

$SelectFolderForm.ShowDialog() | Out-Null
$SelectFolderForm.Dispose()

#endregion
