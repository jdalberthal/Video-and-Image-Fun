<#
.SYNOPSIS
    Displays multiple, independently swimming 3D fish, each showing a looping playlist of images and videos selected by the user.
.DESCRIPTION
    This script first presents a dialog to select image and video files. Once files are selected, it creates a dynamic visual display featuring six 3D fish that swim around the screen in an aquarium-like environment. Each side of each fish independently plays media from the user-selected playlist.

    This version uses the built-in Windows MediaElement for video playback. As a result, video format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI). For broader format support, use the FFmpeg version of this script. The 3D fish are animated procedurally to bounce off the edges of the screen. Users can interact by clicking on the background to randomize the rotation axis of all planes.
.EXAMPLE
    .\Show-ImagesVideosAquariumMediaElement.ps1
    Launches the file selection dialog. After selection, it launches the WPF window and begins the animation.
.NOTES
    Name:           Show-ImagesVideosAquariumMediaElement.ps1
    Version:        1.0.0, 11/14/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. Video playback is limited to formats
                    supported by the built-in Windows MediaElement.
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName WindowsFormsIntegration, System.Xaml
[System.Windows.Forms.Application]::EnableVisualStyles()

$ExternalButtonName = "Aquarium Animation - MediaElement"
$ScriptDescription = "Displays multiple, independently swimming 3D fish, each showing a looping playlist of images and videos using MediaElement."
$RequiredExecutables = @() # No external executables needed

$SyncHash = [hashtable]::Synchronized(@{}) # For passing data between runspaces
$imageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")

#region --- Animation Functions ---

function Stop-MediaResources
{
    param(
        [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]$Timers
    )
    foreach ($timer in $Timers) { $timer.Stop() }
    $Timers.Clear()
}

function Get-NextPlaylistIndex
{
    param($SyncHash)
    $playlistCount = $SyncHash.playlist.Count
    if ($playlistCount -eq 0) { return -1 }

    $nextIndex = -1
    [System.Threading.Monitor]::Enter($SyncHash.SyncRoot)
    try
    {
        $nextIndex = $SyncHash.FileCounter
        $SyncHash.FileCounter = ($SyncHash.FileCounter + 1) % $playlistCount
    }
    finally
    {
        [System.Threading.Monitor]::Exit($SyncHash.SyncRoot)
    }
    return $nextIndex
}

function Start-NextMediaItem
{
    param($SyncHash, $PlaneIndex, $Face) # Face is "Front" or "Back"

    $mediaElement = $SyncHash."Plane${PlaneIndex}${Face}MediaElement"
    $imageElement = $SyncHash."Plane${PlaneIndex}${Face}ImageElement"
    $textBlock = $SyncHash."Plane${PlaneIndex}${Face}TextBlock"
    $errorBorder = $SyncHash."Plane${PlaneIndex}${Face}ErrorBorder"
    $errorTextBlock = $SyncHash."Plane${PlaneIndex}${Face}ErrorTextBlock"

    # Stop any existing timers for this face
    $playerState = $SyncHash.PlayerState["${PlaneIndex}_${Face}"]
    if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
    if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }

    $playlist = $SyncHash.playlist
    if ($playlist.Count -eq 0) { Write-Warning "Playlist is empty."; return }

    $nextIndex = Get-NextPlaylistIndex -SyncHash $SyncHash
    if ($nextIndex -lt 0) { return }
    $filePath = $playlist[$nextIndex]
    $uri = [Uri]$filePath

    # Hide error messages before loading new media
    $errorBorder.Visibility = 'Collapsed'

    # Update text if Filename overlay is active
    if ($SyncHash.RbSelection -eq "Filename" -and $textBlock)
    {
        $textBlock.Text = (Split-Path -Path $filePath -Leaf)
    }

    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

    if ($SyncHash.imageExtensions -contains $extension)
    {
        $playerState.IsImage = $true
        $mediaElement.Visibility = 'Collapsed'
        $mediaElement.Source = $null # Unload video
        $imageElement.Source = [Windows.Media.Imaging.BitmapImage]::new($uri)
        $imageElement.Visibility = 'Visible'

        # Start a timer to simulate MediaEnded for the image
        $imageTimer = New-Object Windows.Threading.DispatcherTimer
        $imageTimer.Interval = [TimeSpan]::FromSeconds($SyncHash.imageDisplaySeconds)
        $playerState.ImageTimer = $imageTimer

        $tickScriptBlock = {
            $imageTimer.Stop()
            Start-NextMediaItem -SyncHash $SyncHash -PlaneIndex $PlaneIndex -Face $Face
        }
        $imageTimer.Add_Tick($tickScriptBlock.GetNewClosure())
        $imageTimer.Start()
    }
    else
    {
        $playerState.IsImage = $false
        $imageElement.Visibility = 'Collapsed'
        $imageElement.Source = $null # Unload image
        $mediaElement.Source = $uri
        $mediaElement.Visibility = 'Visible'
        $mediaElement.Play()
    }
}

function Start-FishMovement
{
    param($SyncHash)
    $SyncHash.FishObjects = [System.Collections.ArrayList]::new()
    $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()

    $startupTimer = New-Object Windows.Threading.DispatcherTimer
    $startupTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $SyncHash.Timers.Add($startupTimer)

    $startupTimer.Add_Tick({
            $startupTimer.Stop()

            # Create and store the two fish mesh geometries
            $rightFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="0.6,0,0 0.5,0.14,0 0.2,0.196,0 -0.1,0.175,0 -0.4,0.105,0 -0.6,0.14,0 -0.667,0.12,0 -0.6,0,0 -0.667,-0.12,0 -0.6,-0.14,0 -0.4,-0.105,0 -0.1,-0.175,0 0.2,-0.196,0 0.5,-0.14,0 0.667,0.05,0 0.667,-0.05,0" TriangleIndices="0,14,1 0,1,13 0,13,15 1,2,12 1,12,13 2,3,11 2,11,12 3,4,10 3,10,11 4,5,7 4,7,10 5,6,7 7,8,9 7,9,10" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
            $leftFishXaml = '<MeshGeometry3D xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Positions="-0.6,0,0 -0.5,0.14,0 -0.2,0.196,0 0.1,0.175,0 0.4,0.105,0 0.6,0.14,0 0.667,0.12,0 0.6,0,0 0.667,-0.12,0 0.6,-0.14,0 0.4,-0.105,0 0.1,-0.175,0 -0.2,-0.196,0 -0.5,-0.14,0 -0.667,0.05,0 -0.667,-0.05,0" TriangleIndices="0,1,14 0,13,1 0,15,13 1,12,2 1,13,12 2,11,3 2,12,11 3,10,4 3,11,10 4,7,5 4,10,7 5,7,6 7,9,8 7,10,9" TextureCoordinates="0.95,0.5 0.85,0.1 0.6,0 0.4,0.05 0.2,0.2 0.05,0.05 0,0.3 0.1,0.5 0,0.7 0.05,0.95 0.2,0.8 0.4,0.95 0.6,1 0.85,0.9 1,0.6 1,0.4" />'
            
            $SyncHash.RightFacingFish = [Windows.Markup.XamlReader]::Parse($rightFishXaml)
            $SyncHash.LeftFacingFish = [Windows.Markup.XamlReader]::Parse($leftFishXaml)

            for ($i = 1; $i -le 6; $i++)
            {
                $SyncHash."Plane${i}FrontMediaElement" = $SyncHash.Window.FindName("videoImage$i")
                $SyncHash."Plane${i}FrontImageElement" = $SyncHash.Window.FindName("staticImage$i")
                $SyncHash."Plane${i}BackMediaElement" = $SyncHash.Window.FindName("backVideoImage$i")
                $SyncHash."Plane${i}BackImageElement" = $SyncHash.Window.FindName("backStaticImage$i")
                
                $SyncHash.PlayerState["${i}_Front"] = @{ IsImage = $false; ImageTimer = $null; RecoveryTimer = $null }
                $SyncHash.PlayerState["${i}_Back"] = @{ IsImage = $false; ImageTimer = $null; RecoveryTimer = $null }

                $planeModel = $SyncHash.Window.FindName("planeModelVisual$i")
                $translateTransform = $SyncHash.Window.FindName("translation$i")
                $frontVp = $SyncHash.Window.FindName("frontViewport$i")
                $backVp = $SyncHash.Window.FindName("backViewport$i")

                # Create a random velocity vector
                $velocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -0.75 -Maximum 0.75), (Get-Random -Minimum -0.25 -Maximum 0.25))
                if ($velocity.X -eq 0) { $velocity.X = 0.5 } # Ensure there's always horizontal movement

                # Give each fish its own clone of the geometry so they can be animated independently
                $rightGeom = $SyncHash.RightFacingFish.Clone()
                $leftGeom = $SyncHash.LeftFacingFish.Clone()

                $fishObject = [pscustomobject]@{
                    Visual = $planeModel
                    Translate = $translateTransform
                    Velocity = $velocity
                    RightGeometry = $rightGeom
                    LeftGeometry = $leftGeom
                }
                [void]$SyncHash.FishObjects.Add($fishObject)

                Start-NextMediaItem -SyncHash $SyncHash -PlaneIndex $i -Face "Front"
                Start-NextMediaItem -SyncHash $SyncHash -PlaneIndex $i -Face "Back"
            }
        }.GetNewClosure())

    $startupTimer.Start()
}

function Show-AquariumAnimation
{
    param(
        [hashtable]$SyncHash,
        [array]$playlist
    )
    $SyncHash.Paused = $false
    $SyncHash.imageDisplaySeconds = 15
    $SyncHash.imageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")
    $SyncHash.TailWiggleSpeed = 18.0 # Controls how fast the tail wiggles
    $SyncHash.TailWiggleAmount = 0.12 # Controls how far the tail wiggles

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Name="AquariumWindow"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Aquarium"
        WindowStartupLocation="CenterScreen" 
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" >
    <Grid Name="MainGrid" Background="Transparent" IsHitTestVisible="True">
        <Border Name="viewportBorder" Background="Transparent">
            <!-- The 3D scene with the fish -->
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
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport1">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage1" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage1" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront1" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront1" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront1" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport1">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage1" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage1" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack1" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack1" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack1" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual2">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation2" />
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport2">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage2" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage2" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront2" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront2" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront2" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport2">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage2" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage2" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack2" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack2" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack2" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual3">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation3" />
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport3">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage3" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage3" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront3" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront3" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront3" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport3">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage3" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage3" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack3" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack3" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack3" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual4">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation4" />
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport4">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage4" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage4" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront4" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront4" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront4" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport4">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage4" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage4" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack4" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack4" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack4" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual5">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation5" />
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport5">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage5" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage5" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront5" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront5" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront5" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport5">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage5" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage5" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack5" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack5" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack5" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            <ModelVisual3D x:Name="planeModelVisual6">
                <ModelVisual3D.Transform>
                   <Transform3DGroup>
                       <TranslateTransform3D x:Name="translation6" />
                   </Transform3DGroup>
                </ModelVisual3D.Transform>
                <ModelVisual3D.Children>
                    <Viewport2DVisual3D x:Name="frontViewport6">
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="videoImage6" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="staticImage6" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayFront6" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderFront6" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockFront6" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                    <Viewport2DVisual3D x:Name="backViewport6">
                        <Viewport2DVisual3D.Transform><RotateTransform3D><RotateTransform3D.Rotation><AxisAngleRotation3D Axis="0,1,0" Angle="180"/></RotateTransform3D.Rotation></RotateTransform3D></Viewport2DVisual3D.Transform>
                        <Viewport2DVisual3D.Geometry><MeshGeometry3D/></Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material><DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" /></Viewport2DVisual3D.Material>                        
                        <Grid Background="Transparent" Width="800" Height="400">
                            <MediaElement Name="backVideoImage6" Stretch="Fill" LoadedBehavior="Manual" UnloadedBehavior="Stop" />
                            <Image Name="backStaticImage6" Stretch="Fill" Visibility="Collapsed" />
                            <TextBlock Name="textOverlayBack6" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="5" IsHitTestVisible="False" />
                            <Border Name="errorBorderBack6" Panel.ZIndex="1" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="200" Height="200">
                                <TextBlock Name="errorTextBlockBack6" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
            </Viewport3D>
        </Border>
        <!-- 2D Canvas for bubbles, overlaid on top of the 3D viewport -->
        <Canvas Name="bubbleCanvas" IsHitTestVisible="False" />
        <StackPanel Name="controlsPanel" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="5">
            <Button Name="tapGlassButton" Content="Tap Glass" Padding="10,5" Margin="2"/>
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

    $SyncHash.playlist = $playlist | Get-Random -Count $playlist.Count # Shuffle the playlist
    $SyncHash.FileCounter = 0
    $SyncHash.SyncRoot = New-Object System.Object

    $SyncHash.Timers = [System.Collections.Generic.List[System.Windows.Threading.DispatcherTimer]]::new()

    $SyncHash.PlayerState = [hashtable]::Synchronized(@{})
    $SyncHash.Window = $window

    # Set background based on user selection
    if ($SyncHash.AddWater) {
        $waterColor = [System.Windows.Media.ColorConverter]::ConvertFromString("#80ADD8E6") # Semi-transparent light blue
        $window.Background = [System.Windows.Media.SolidColorBrush]::new($waterColor)
    }

    # Store controls for easy access
    $SyncHash.pauseButton = $window.FindName("pauseButton")
    $SyncHash.slowDownButton = $window.FindName("slowDownButton")
    $SyncHash.speedUpButton = $window.FindName("speedUpButton")
    $SyncHash.redoButton = $window.FindName("redoButton")
    $SyncHash.hideControlsButton = $window.FindName("hideControlsButton")
    $SyncHash.tapGlassButton = $window.FindName("tapGlassButton")
    
    # Store TextBlocks and Error Borders for overlay
    for ($i = 1; $i -le 6; $i++)
    {
        $SyncHash."Plane${i}FrontTextBlock" = $SyncHash.Window.FindName("textOverlayFront$i")
        $SyncHash."Plane${i}BackTextBlock" = $SyncHash.Window.FindName("textOverlayBack$i")
        $SyncHash."Plane${i}FrontErrorBorder" = $SyncHash.Window.FindName("errorBorderFront$i")
        $SyncHash."Plane${i}BackErrorBorder" = $SyncHash.Window.FindName("errorBorderBack$i")
        $SyncHash."Plane${i}FrontErrorTextBlock" = $SyncHash.Window.FindName("errorTextBlockFront$i")
        $SyncHash."Plane${i}BackErrorTextBlock" = $SyncHash.Window.FindName("errorTextBlockBack$i")
    }

    $closeButton = $SyncHash.Window.FindName('closeButton')
    $PrimaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
    $WorkAreaWidth = $PrimaryScreen.WorkingArea.Width
    $WorkAreaHeight = $PrimaryScreen.WorkingArea.Height

    $SyncHash.Window.Width = [System.Windows.SystemParameters]::WorkArea.Width
    $SyncHash.Window.Height = [System.Windows.SystemParameters]::WorkArea.Height
    $SyncHash.Window.Left = 0
    $SyncHash.Window.Top = 0

    # --- Media Event Handlers ---
    $mediaEndedHandler = {
        param($sender, $e)
        $mediaElement = $sender
        $planeIndex = $mediaElement.Name -replace '\D'
        $face = if ($mediaElement.Name -like "back*") { "Back" } else { "Front" }
        Start-NextMediaItem -SyncHash $SyncHash -PlaneIndex $planeIndex -Face $face
    }

    $mediaFailedHandler = {
        param($sender, $e)
        $mediaElement = $sender
        $planeIndex = $mediaElement.Name -replace '\D'
        $face = if ($mediaElement.Name -like "back*") { "Back" } else { "Front" }
        
        $errorBorder = $SyncHash."Plane${planeIndex}${face}ErrorBorder"
        $errorTextBlock = $SyncHash."Plane${planeIndex}${face}ErrorTextBlock"
        $errorTextBlock.Text = "Media Failed: $($mediaElement.Source.Segments[-1])"
        $errorBorder.Visibility = 'Visible'

        # Start a recovery timer
        $playerState = $SyncHash.PlayerState["${planeIndex}_${face}"]
        if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
        $recoveryTimer = New-Object Windows.Threading.DispatcherTimer
        $recoveryTimer.Interval = [TimeSpan]::FromSeconds(5)
        $playerState.RecoveryTimer = $recoveryTimer
        
        $tickScriptBlock = {
            $recoveryTimer.Stop()
            Start-NextMediaItem -SyncHash $SyncHash -PlaneIndex $planeIndex -Face $face
        }
        $recoveryTimer.Add_Tick($tickScriptBlock.GetNewClosure())
        $recoveryTimer.Start()
    }

    # Attach handlers to all MediaElements
    for ($i = 1; $i -le 6; $i++)
    {
        $frontME = $SyncHash.Window.FindName("videoImage$i")
        $backME = $SyncHash.Window.FindName("backVideoImage$i")
        $frontME.Add_MediaEnded($mediaEndedHandler)
        $backME.Add_MediaEnded($mediaEndedHandler)
        $frontME.Add_MediaFailed($mediaFailedHandler)
        $backME.Add_MediaFailed($mediaFailedHandler)
    }

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

    switch ($SyncHash.RbSelection)
    {
        "Hidden"
        {
            # Text is hidden by default, nothing to do.
        }
        "Filename"
        {
            # Apply styles to all text blocks initially
            for ($i = 1; $i -le 6; $i++)
            {
                & $applyTextStyles -textBlock $SyncHash."Plane${i}FrontTextBlock"
                & $applyTextStyles -textBlock $SyncHash."Plane${i}BackTextBlock"
            }
        }
        "Custom"
        {
            $customText = $SyncHash.TextBox.Text
            for ($i = 1; $i -le 6; $i++)
            {
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
            switch ($e.Key)
            {
                'Escape' { $SyncHash.Window.Close() }
                'P' { $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'H' { $SyncHash.hideControlsButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'R' { $SyncHash.redoButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'T' { $SyncHash.tapGlassButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'Left' { $SyncHash.slowDownButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'Right' { $SyncHash.speedUpButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                'F1'
                {
                    $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                    $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
                    $OkButton = $PopupWindow.FindName("OKButton")
                    $OkButton.Add_Click({ $PopupWindow.Close() })
                    $PopupWindow.ShowDialog() | Out-Null
                }
            }
        })

    $SyncHash.tapGlassButton.Add_Click({
        # Give each fish a new random velocity
        foreach ($fish in $SyncHash.FishObjects) {
            $fish.Velocity = New-Object System.Windows.Media.Media3D.Vector3D((Get-Random -Minimum -1.0 -Maximum 1.0), (Get-Random -Minimum -0.75 -Maximum 0.75), (Get-Random -Minimum -0.25 -Maximum 0.25))
            if ($fish.Velocity.X -eq 0) { $fish.Velocity.X = 0.5 } # Ensure there's always horizontal movement
        }
    })
    
    $SyncHash.pauseButton.Add_Click({
        $SyncHash.Paused = -not $SyncHash.Paused
        if ($SyncHash.Paused) {
            $SyncHash.pauseButton.Content = "Resume"
        } else {
            $SyncHash.pauseButton.Content = "Pause"
            # Reset frame timer to avoid a large jump after unpausing
            $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        }
    })

    $SyncHash.hideControlsButton.Add_Click({
            $controlsPanel = $SyncHash.Window.FindName("controlsPanel")
            $controlsPanel.Visibility = if ($controlsPanel.Visibility -eq 'Visible') { 'Collapsed' } else { 'Visible' }
        })

    $SyncHash.redoButton.Add_Click({
            $SyncHash.Window.Close()
            $SelectFolderForm.Show()
        })

    $changeSpeed = {
        param($multiplier)
        foreach ($fish in $SyncHash.FishObjects) {
            $fish.Velocity *= $multiplier
        }
    }

    $SyncHash.slowDownButton.Add_Click({
            & $changeSpeed 0.5 # Slower
        })

    $SyncHash.speedUpButton.Add_Click({
            & $changeSpeed 2.0 # Faster
        })

    # --- Define the per-frame rendering logic ---
    $SyncHash.RenderHandler = {
        param($sender, $e)
        if ($SyncHash.Paused) { return }

        $currentTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $elapsed = ($currentTime - $SyncHash.LastFrameTime) / [System.Diagnostics.Stopwatch]::Frequency
        $SyncHash.LastFrameTime = $currentTime
        $totalTime = ($currentTime - $SyncHash.StartTime) / [System.Diagnostics.Stopwatch]::Frequency

        $xBoundary = $SyncHash.xBoundary; $yBoundary = $SyncHash.yBoundary; $zBoundary = 3

        # Define the margins based on half the fish's model size
        $xMargin = 0.7 
        $yMargin = 0.3

        # Calculate the tail oscillation value based on time
        $oscillation = [Math]::Sin($totalTime * $SyncHash.TailWiggleSpeed) * $SyncHash.TailWiggleAmount

        foreach ($fish in $SyncHash.FishObjects) {
            # 1. Get current velocity
            $velX = $fish.Velocity.X
            $velY = $fish.Velocity.Y
            $velZ = $fish.Velocity.Z

            # 2. Predict next position of the fish's center
            $nextX = $fish.Translate.OffsetX + ($velX * $elapsed)
            $nextY = $fish.Translate.OffsetY + ($velY * $elapsed)
            $nextZ = $fish.Translate.OffsetZ + ($velZ * $elapsed)

            # 3. Check if the *edge* of the fish will hit the boundary and rebound if so
            if (($nextX + $xMargin) -gt $xBoundary -and $velX -gt 0) {
                $velX *= -1
                $nextX = $xBoundary - $xMargin # Clamp position to the edge
            } elseif (($nextX - $xMargin) -lt -$xBoundary -and $velX -lt 0) {
                $velX *= -1
                $nextX = -$xBoundary + $xMargin # Clamp position to the edge
            }

            if (($nextY + $yMargin) -gt $yBoundary -and $velY -gt 0) {
                $velY *= -1
                $nextY = $yBoundary - $yMargin # Clamp position to the edge
            } elseif (($nextY - $yMargin) -lt -$yBoundary -and $velY -lt 0) {
                $velY *= -1
                $nextY = -$yBoundary + $yMargin # Clamp position to the edge
            }

            if (($nextZ -gt $zBoundary -and $velZ -gt 0) -or ($nextZ -lt -$zBoundary -and $velZ -lt 0)) {
                $velZ *= -1
            }
            $fish.Velocity = New-Object System.Windows.Media.Media3D.Vector3D($velX, $velY, $velZ)
            
            # 4. Apply the (potentially clamped) new position
            $fish.Translate.OffsetX = $nextX
            $fish.Translate.OffsetY = $nextY
            $fish.Translate.OffsetZ = $nextZ

            # 5. Update Directional Geometry and Animate Tail
            $frontVp = $fish.Visual.Children[0]
            $backVp = $fish.Visual.Children[1]
            $currentGeometry = $null

            if ($fish.Velocity.X -gt 0) {
                if ($frontVp.Geometry -ne $fish.RightGeometry) {
                    $frontVp.Geometry = $fish.RightGeometry
                    $backVp.Geometry = $fish.RightGeometry
                }
                $currentGeometry = $fish.RightGeometry
            } else {
                if ($frontVp.Geometry -ne $fish.LeftGeometry) {
                    $frontVp.Geometry = $fish.LeftGeometry
                    $backVp.Geometry = $fish.LeftGeometry
                }
                $currentGeometry = $fish.LeftGeometry
            }

            if ($currentGeometry) {
                $positions = $currentGeometry.Positions
                # Animate the tail vertices (6, 7, 8, 5, 9)
                # The Z value is flipped for left-facing fish to keep the wiggle direction consistent
                $zWiggle = if ($fish.Velocity.X -gt 0) { $oscillation } else { -$oscillation }

                $positions[6] = [System.Windows.Media.Media3D.Point3D]::new($positions[6].X, $positions[6].Y, $zWiggle)
                $positions[8] = [System.Windows.Media.Media3D.Point3D]::new($positions[8].X, $positions[8].Y, $zWiggle)
                $positions[5] = [System.Windows.Media.Media3D.Point3D]::new($positions[5].X, $positions[5].Y, $zWiggle * 0.6)
                $positions[9] = [System.Windows.Media.Media3D.Point3D]::new($positions[9].X, $positions[9].Y, $zWiggle * 0.6)
            }
        }
    }.GetNewClosure()

    $window.Add_Loaded({
        # Start the per-frame rendering loop
        # --- Correctly Calculate the 3D Boundaries using the camera's properties ---
        $viewport = $SyncHash.Window.FindName("mainViewport")
        $camera = $viewport.Camera

        if ($camera -is [System.Windows.Media.Media3D.PerspectiveCamera]) {
            # This is the standard formula for calculating the frustum size at a given distance.
            $distance = $camera.Position.Z # Distance from camera to the Z=0 plane
            $fovRadians = 45.0 * ([Math]::PI / 180.0) # The default FOV is 45 degrees.
            $viewHeight3D = 2.0 * $distance * [Math]::Tan($fovRadians / 2.0)
            $aspectRatio = if ($viewport.ActualHeight -gt 0) { $viewport.ActualWidth / $viewport.ActualHeight } else { 1.0 }
            $viewWidth3D = $viewHeight3D * $aspectRatio

            # Halving the boundaries as a test, as requested.
            $SyncHash.xBoundary = ($viewWidth3D / 2.0) / 2.0
            $SyncHash.yBoundary = ($viewHeight3D / 2.0) / 2.0
        } else {
            # Fallback for an unexpected camera type
            $SyncHash.xBoundary = 8.9
            $SyncHash.yBoundary = 5.0
        }
        $SyncHash.LastFrameTime = [System.Diagnostics.Stopwatch]::GetTimestamp()
        $SyncHash.StartTime = [System.Diagnostics.Stopwatch]::GetTimestamp() # For tail animation timing
        [System.Windows.Media.CompositionTarget]::add_Rendering($SyncHash.RenderHandler)
    })

    $SyncHash.Window.Add_Closed({
            param($sender, $e)
            # Stop the per-frame rendering loop
            if ($SyncHash.RenderHandler) {
                [System.Windows.Media.CompositionTarget]::remove_Rendering($SyncHash.RenderHandler)
            }

            # Stop all timers
            foreach ($timer in $SyncHash.Timers) { $timer.Stop() }
            $SyncHash.Timers.Clear()

            # Stop all MediaElements
            for ($i = 1; $i -le 6; $i++)
            {
                $frontME = $SyncHash."Plane${i}FrontMediaElement"
                $backME = $SyncHash."Plane${i}BackMediaElement"
                if ($frontME) { $frontME.Close() }
                if ($backME) { $backME.Close() }
            }
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
                    $SyncHash.hitModel = $result.VisualHit
                    return [System.Windows.Media.HitTestResultBehavior]::Stop
                }
                return [System.Windows.Media.HitTestResultBehavior]::Continue
            }
            $hitTestParams = [System.Windows.Media.PointHitTestParameters]::new($mousePosition)
            [System.Windows.Media.VisualTreeHelper]::HitTest($viewport, $null, $hitTestCallback, $hitTestParams)

            if ($SyncHash.hitModel -is [System.Windows.Media.Media3D.Viewport2DVisual3D])
            {
                # Trigger the pause/resume functionality, just like the 'P' key or pause button
                $SyncHash.pauseButton.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
            }
        })

    Start-FishMovement -SyncHash $SyncHash

    # --- Bubble Animation Logic (from Bubbles.ps1) ---
    if ($SyncHash.AddBubbles) {
        $bubbleCanvas = $window.FindName('bubbleCanvas')
        $rand = [Random]::new()
        $bubbles = New-Object System.Collections.Generic.List[object]
        $maxCount = 220

        $newBubbleBrushFunc = {
            param([int]$alpha)
            $c1 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Min($alpha,255), 255, 255, 255)
            $c2 = [System.Windows.Media.Color]::FromArgb([byte][Math]::Max($alpha-120,30), 173, 216, 230)
            $brush = New-Object Windows.Media.RadialGradientBrush
            $brush.RadiusX = 0.6; $brush.RadiusY = 0.6
            $brush.GradientOrigin = [Windows.Point]::new(0.35,0.35)
            $brush.Center = [Windows.Point]::new(0.5,0.5)
            $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c1,0.0))
            $brush.GradientStops.Add([Windows.Media.GradientStop]::new($c2,1.0))
            return $brush
        }

        $newBubbleFunc = {
            $w = [double]($bubbleCanvas.ActualWidth)
            $h = [double]($bubbleCanvas.ActualHeight)
            if ($w -le 0 -or $h -le 0) { return }

            $size   = [double]($rand.Next(8, 42))
            $speed  = [double]($rand.NextDouble() * 1.4 + 0.6)
            $drift  = [double]((($rand.NextDouble()*2.0) - 1.0) * 0.35)
            $alpha  = $rand.Next(120, 230)

            $startX = [double]($rand.NextDouble() * [Math]::Max($w - $size, 1))
            $startY = [double]($h + $rand.NextDouble() * ([Math]::Max($h*0.15, 80)))

            $ellipse = New-Object Windows.Shapes.Ellipse
            $ellipse.Width  = $size
            $ellipse.Height = $size
            $ellipse.Fill   = & $newBubbleBrushFunc -alpha $alpha
            $ellipse.Stroke = [Windows.Media.Brushes]::White
            $ellipse.StrokeThickness = [Math]::Max($size * 0.02, 0.6)

            [Windows.Controls.Canvas]::SetLeft($ellipse, $startX)
            [Windows.Controls.Canvas]::SetTop($ellipse,  $startY)

            $bubbleCanvas.Children.Add($ellipse) | Out-Null

            $bubble = [pscustomobject]@{ Shape = $ellipse; Vy = $speed; Vx = $drift; Spin = ($rand.NextDouble() * 0.04) - 0.02; T = $rand.NextDouble() * [Math]::PI }
            $bubbles.Add($bubble) | Out-Null
        }

        $spawnTimer = New-Object Windows.Threading.DispatcherTimer
        $spawnTimer.Interval = [TimeSpan]::FromMilliseconds(220)
        $spawnTimer.Add_Tick({
            if ($bubbles.Count -lt $maxCount) {
                $count = 1 + $rand.Next(0,3)
                for ($i=0; $i -lt $count; $i++) { & $newBubbleFunc }
            }
        })

        $animTimer = New-Object Windows.Threading.DispatcherTimer
        $animTimer.Interval = [TimeSpan]::FromMilliseconds(16)
        $animTimer.Add_Tick({
            if ($bubbleCanvas -eq $null) { return }
            if ($SyncHash.Paused) { return } # Pause bubbles with the fish
            $h = [double]($bubbleCanvas.ActualHeight)
            $w = [double]($bubbleCanvas.ActualWidth)

            for ($i = $bubbles.Count - 1; $i -ge 0; $i--) {
                $b = $bubbles[$i]
                $s = [double]$b.Shape.Width
                $x = [Windows.Controls.Canvas]::GetLeft($b.Shape)
                $y = [Windows.Controls.Canvas]::GetTop($b.Shape)

                $b.T += $b.Spin
                $x += $b.Vx + ([Math]::Sin($b.T) * 0.15)
                $y -= $b.Vy

                if ($x -lt -10) { $x = -10; $b.Vx = [Math]::Abs($b.Vx) }
                elseif ($x + $s -gt $w + 10) { $x = $w + 10 - $s; $b.Vx = -[Math]::Abs($b.Vx) }

                [Windows.Controls.Canvas]::SetLeft($b.Shape, $x)
                [Windows.Controls.Canvas]::SetTop($b.Shape,  $y)

                if ($y + $s -lt 0) {
                    $bubbleCanvas.Children.Remove($b.Shape)
                    $bubbles.RemoveAt($i)
                }
            }
        })

        $SyncHash.Timers.Add($spawnTimer)
        $SyncHash.Timers.Add($animTimer)

        $spawnTimer.Start()
        $animTimer.Start()
    }

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
                    <Run Text="P           : Pause / Resume Animation"/><LineBreak/>
                    <Run Text="R           : Reselect Media Files (Redo)"/><LineBreak/>
                    <Run Text="T           : Tap Glass (New Random Paths)"/><LineBreak/>
                    <Run Text="H           : Hide / Show Controls"/><LineBreak/>
                    <Run Text="&#x2190; (Left)    : Slow Down Animation"/><LineBreak/>
                    <Run Text="&#x2192; (Right)   : Speed Up Animation"/><LineBreak/><LineBreak/>
                    <Run Text="*Click a fish to pause/resume."/><LineBreak/>
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
if ($RequiredExecutables -and $RequiredExecutables.Count -gt 0)
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
$SelectFolderForm.Text = "Aquarium - Media Selector"
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

# Create a CheckBox for adding water background
$AddWaterCheckbox = New-Object System.Windows.Forms.CheckBox
$AddWaterCheckbox.Text = "Add Water"
$AddWaterCheckbox.AutoSize = $True
$AddWaterCheckbox.Location = New-Object System.Drawing.Point(150, 40)
$AddWaterCheckbox.Checked = $True
$SelectFolderForm.Controls.Add($AddWaterCheckbox)

# Create a CheckBox for adding bubbles
$BubblesCheckbox = New-Object System.Windows.Forms.CheckBox
$BubblesCheckbox.Text = "Add Bubbles"
$BubblesCheckbox.AutoSize = $True
$BubblesCheckbox.Location = New-Object System.Drawing.Point(250, 40)
$BubblesCheckbox.Checked = $True
$SelectFolderForm.Controls.Add($BubblesCheckbox)

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

$HeaderLabel = New-Object System.Windows.Forms.Label
$HeaderLabel.Text = "Play Media"
$HeaderLabel.Location = New-Object System.Drawing.Point(5, 5) # Position at top-left
$HeaderLabel.AutoSize = $True # Ensures the label resizes to fit the text
$HeaderLabel.BackColor = [System.Drawing.Color]::Transparent
$HeaderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$DataGridView.Controls.Add($HeaderLabel)

# Create a Button to perform an action on selected files
$PlayButton = New-Object System.Windows.Forms.Button
$PlayButton.Text = "Play Selected Item(s)"
$PlayButton.Location = New-Object System.Drawing.Point(600, 40)
$PlayButton.Size = New-Object System.Drawing.Size(170, 30)
$SelectFolderForm.Controls.Add($PlayButton)

#region --- Text Overlay and F1 Help UI (Copied from Cube Script) ---

$MyFont = New-Object System.Drawing.Font("Arial", 24)

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
$TextBox.TextAlign = "Center"
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
$FontButton.Font = "Arial, 24"
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

$colorDialog = New-Object System.Windows.Forms.ColorDialog
$SelectColorButton.Add_Click({
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $ColorExample.BackColor = $colorDialog.Color
            $TextBox.ForeColor = $colorDialog.Color
            $SyncHash.TextColor = $colorDialog.Color
        }
    })

$updateFontFromControls = {
    $style = [System.Drawing.FontStyle]::Regular
    if ($BoldCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Bold }
    if ($ItalicCheckbox.Checked) { $style = $style -bor [System.Drawing.FontStyle]::Italic }
    $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, [float]$NumericUpDown.Value, $style)
    $SyncHash.SelectedFontSize = $NumericUpDown.Value
}

$FontButton.Add_Click({
        $fontDialog = New-Object System.Windows.Forms.FontDialog
        $fontDialog.Font = $TextBox.Font
        if ($fontDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $TextBox.Font = $fontDialog.Font
            $FontButton.Text = $fontDialog.Font.Name
            $SyncHash.SelectedFont = $fontDialog.Font.Name
            $SyncHash.ItalicCheckbox.Checked = $fontDialog.Font.Italic
            $SyncHash.BoldCheckbox.Checked = $fontDialog.Font.Bold
            $SyncHash.NumericUpDown.Value = [Math]::Round($fontDialog.Font.Size)
        }
    })

$NumericUpDown.Add_ValueChanged($updateFontFromControls)
$ItalicCheckbox.Add_CheckedChanged($updateFontFromControls)
$BoldCheckbox.Add_CheckedChanged($updateFontFromControls)

$SelectFolderForm.KeyPreview = $True
$SelectFolderForm.Add_KeyDown({
        param($Sender, $e)
        if ($e.KeyCode -eq "F1")
        {
            $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
            $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
            $OkButton = $PopupWindow.FindName("OKButton")
            $OkButton.Add_Click({ $PopupWindow.Close() })
            $PopupWindow.ShowDialog() | Out-Null
        }
    })

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
        $VideoExtensions = "*.wmv", "*.asf", "*.mpg", "*.mpeg", "*.mp2", "*.mpe", "*.mpv", "*.avi", "*.mp4", "*.mov", "*.webm", "*.mkv"
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
            $SyncHash.AddWater = $AddWaterCheckbox.Checked
            $SyncHash.AddBubbles = $BubblesCheckbox.Checked
            Show-AquariumAnimation -SyncHash $SyncHash -playlist $selectedFiles 
        }
    })

$SelectFolderForm.ShowDialog() | Out-Null
$SelectFolderForm.Dispose()

#endregion
