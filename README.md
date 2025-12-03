# PowerShell Media & 3D Visualization Suite

This is a collection of powerful, GUI-based PowerShell scripts designed for a variety of media and file management tasks. The suite includes tools for analyzing video corruption, extracting detailed file metadata, and displaying images and videos in unique, dynamic ways—such as on a rotating 3D cube or in a continuous scroller.

The main entry point is `Show-ScriptLauncher.ps1`, which provides a user-friendly interface to discover and launch all other scripts in the collection, checking for dependencies and grouping them automatically.

## Getting Started

1.  Ensure all prerequisites are met (see below).
2.  Place all `.ps1` script files into the same directory.
3.  Run the main launcher script from a PowerShell terminal: `.\Show-ScriptLauncher.ps1`

## The Scripts

### Launcher

- **`Show-ScriptLauncher.ps1`**
  - **Description**: A dynamic GUI that scans its directory for other PowerShell scripts and creates launch buttons for them, grouping them by their dependencies.
  - **Dependencies**: PowerShell with .NET Framework access.

### Tools

- **`Get-VideoCorruptionGPUFfmpeg.ps1`**
  - **Description**: A GUI-based tool to scan, analyze, and attempt repairs on video files using FFmpeg.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`, `recover_mp4.exe`.

- **`Get-AllFilesData.ps1`**
  - **Description**: Scans selected files or folders to extract and display detailed metadata, including EXIF data for images and technical properties for other files.
  - **Dependencies**: PowerShell with .NET Framework access.

### Media Viewers (FFmpeg-Based)

- **`Show-RotatingImagesVideosCubeFfmpeg.ps1`**
  - **Description**: Displays selected images and videos on the faces of a rotating 3D cube using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosCarouselFfmpeg.ps1`**
  - **Description**: Displays media on vertical panels in a rotating 3D carousel with an undulating wave motion. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosConcentricFunnelFfmpeg.ps1`**
  - **Description**: Displays media on a static 3D funnel mesh built from concentric rings of trapezoid panels. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosButterflyEffectFfmpeg.ps1`**
  - **Description**: Creates a dynamic visual display featuring six 3D planes that move around the screen in a butterfly-like pattern, with each face independently playing media from a user-selected playlist.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-FloatingCubesFfmpeg.ps1`**
  - **Description**: Displays media on multiple, independently moving and rotating 3D cubes using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosFacetedSphereFfmpegMulti.ps1`**
  - **Description**: Displays multiple media files at once, with each facet of a rotating 3D sphere showing a different file from the playlist. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`, `ffprobe.exe`.

- **`Show-ImagesVideosFacetedSphereFfmpegMulti.ps1`**
  - **Description**: Displays multiple media files at once, with each facet of a rotating 3D sphere showing a different file from the playlist. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`, `ffprobe.exe`.

- **`Show-ImagesVideosFacetedSphereFfmpegSingle.ps1`**
  - **Description**: Displays one media file at a time from a playlist, showing the same media on all facets of a rotating 3D sphere. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosFloatingStarsFfmpeg.ps1`**
  - **Description**: Displays media on six independently floating and bouncing 3D stars using FFmpeg for broad video format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosFunnelFfmpeg.ps1`**
  - **Description**: Displays selected images and videos on the surface of a rotating 3D funnel using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosMediaFlowFunnelFfmpeg.ps1`**
  - **Description**: Displays media that "flows" down the panels of a static, concentric 3D funnel. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosFunnelFfmpeg.ps1`**
  - **Description**: Displays selected images and videos on the surface of a rotating 3D funnel using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosMediaFlowFunnelFfmpeg.ps1`**
  - **Description**: Displays media that "flows" down the panels of a static, concentric 3D funnel. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosPie3DFfmpeg.ps1`**
  - **Description**: Displays media on the front and back of 8 rotating 3D pie slices. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`, `ffprobe.exe`.

- **`Show-ImagesVideosPinwheelFfmpeg.ps1`**
  - **Description**: Displays media on multiple, independently rotating 3D planes arranged in a pinwheel pattern. Uses FFmpeg for broad video format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosRotatingStarFfmpeg.ps1`**
  - **Description**: Displays media on a rotating 3D star-shaped object. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ScrollingImagesVideosHorizontalFfmpeg.ps1`**
  - **Description**: Creates a continuous horizontal-scrolling display of selected images and videos. Uses FFmpeg for video decoding.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ScrollingImagesVideosHorizontalFfmpeg.ps1`**
  - **Description**: Creates a continuous horizontal-scrolling display of selected images and videos. Uses FFmpeg for video decoding.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ScrollingImagesVideosVerticalFfmpeg.ps1`**
  - **Description**: Creates a continuous vertical-scrolling display of selected images and videos. Uses FFmpeg for video decoding.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosSphereFfmpeg.ps1`**
  - **Description**: Displays selected images and videos on a rotating 3D sphere using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosSphereVortexFfmpeg.ps1`**
  - **Description**: Displays media on spheres that spiral down a rotating 3D vortex. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosSphereVortexFfmpeg.ps1`**
  - **Description**: Displays media on spheres that spiral down a rotating 3D vortex. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosWagonWheelFfmpeg.ps1`**
  - **Description**: Displays media on the outer curved faces of rotating 3D wagon wheel slices. Each slice plays media independently. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

- **`Show-ImagesVideosCarouselFfmpeg.ps1`**
  - **Description**: Displays media on vertical panels in a rotating 3D carousel with an undulating wave motion. Uses FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffprobe.exe`, `ffplay.exe`.

### Media Viewers (MediaElement-Based)

- **`Show-RotatingImagesVideosCubeMediaElement.ps1`**
  - **Description**: Displays selected images and videos on a rotating 3D cube using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosButterflyEffectMediaElement.ps1`**
  - **Description**: Creates a dynamic visual display featuring six 3D planes that move around the screen in a butterfly-like pattern, with each face independently playing media from a user-selected playlist.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosCarouselMediaElement.ps1`**
  - **Description**: Displays media on vertical panels in a rotating 3D carousel with an undulating wave motion. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosConcentricFunnelMediaElement.ps1`**
  - **Description**: Displays media on a static 3D funnel mesh built from concentric rings of trapezoid panels. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-FloatingCubesMediaElement.ps1`**
  - **Description**: Displays media on multiple, independently moving and rotating 3D cubes using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosFacetedSphereMediaElementMulti.ps1`**
  - **Description**: Displays multiple media files at once, with each facet of a rotating 3D sphere showing a different file from the playlist. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosFacetedSphereMediaElementSingle.ps1`**
  - **Description**: Displays one media file at a time from a playlist, showing the same media on all facets of a rotating 3D sphere. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosFloatingStarsMediaElement.ps1`**
  - **Description**: Displays media on six independently floating and bouncing 3D stars using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosFunnelMediaElement.ps1`**
  - **Description**: Displays selected images and videos on the surface of a rotating 3D funnel using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosMediaFlowFunnelMediaElement.ps1`**
  - **Description**: Displays media that "flows" down the panels of a static, concentric 3D funnel. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosPulsingStarMediaElement.ps1`**
  - **Description**: Displays media on a 3D object resembling a pulsing star, made of a central sphere and multiple cones. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosMediaFlowFunnelMediaElement.ps1`**
  - **Description**: Displays media that "flows" down the panels of a static, concentric 3D funnel. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosPie3DMediaElement.ps1`**
  - **Description**: Displays media on the front and back of 8 rotating 3D pie slices using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosPinwheelMediaElement.ps1`**
  - **Description**: Displays media on multiple, independently rotating 3D planes arranged in a pinwheel pattern using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosRotatingStarMediaElement.ps1`**
  - **Description**: Displays selected images and videos on a rotating 3D star-shaped object using MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ScrollingImagesVideosHorizontalMediaElement.ps1`**
  - **Description**: Creates a horizontal-scrolling display using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ScrollingImagesVideosHorizontalMediaElement.ps1`**
  - **Description**: Creates a horizontal-scrolling display using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ScrollingImagesVideosVerticalMediaElement.ps1`**
  - **Description**: Creates a vertical-scrolling display using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosSphereMediaElement.ps1`**
  - **Description**: Loops through and displays selected images or videos on a rotating 3D sphere. Uses the built-in Windows MediaElement, which may have more limited video format support.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosSphereVortexMediaElement.ps1`**
  - **Description**: Displays media on spheres that spiral down a rotating 3D vortex. Uses the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ImagesVideosWagonWheelMediaElement.ps1`**
  - **Description**: Displays media on the outer curved faces of rotating 3D wagon wheel slices. Each slice plays media independently. Uses the built-in Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

### Prerequisites

Before running the scripts, ensure you have the following installed and accessible in your system's PATH.

1. **PowerShell**: Version 5.1 or higher.
2. **.NET Framework**: Required for the GUI components. This is included by default in modern Windows versions.
3. **FFmpeg**: Required for the more advanced video scripts.
    - **Download**: ffmpeg.org/download.html
    - **Installation**: Download the binaries and add the `bin` folder (containing `ffmpeg.exe`, `ffprobe.exe`, and `ffplay.exe`) to your system's PATH environment variable.
4. **recover_mp4.exe**: Required for the video repair functionality.
    - **Download**: videohelp.com/software/recover-mp4-to-h264
    - **Installation**: Place `recover_mp4.exe` in the same directory as the scripts or in a folder that is in your system's PATH.

## Author

- **JD Alberthal**
  - Website: jdalberthal.com
  - GitHub: @jdalberthal
  - Email: `jd@jdalberthal.com`

## License

This project is licensed under the MIT License - see the `LICENSE.md` file for details.
