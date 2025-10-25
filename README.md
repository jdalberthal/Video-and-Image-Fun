# PowerShell Media & File Utilities

This is a collection of powerful, GUI-based PowerShell scripts designed for a variety of media and file management tasks. The suite includes tools for analyzing video corruption, extracting detailed file metadata, and displaying images and videos in unique, dynamic ways—such as on a rotating 3D cube or in a continuous scroller.

The main entry point is `Show-ScriptLauncher.ps1`, which provides a user-friendly interface to discover and launch all other scripts in the collection, checking for dependencies and grouping them automatically.

## Features

- **Dynamic Script Launcher**: Automatically discovers and creates launch buttons for all scripts in its directory.
- **Video Corruption Analysis**: Scan video files for general corruption, container mismatches, and `moov atom` placement using FFmpeg.
- **Video Repair**: Attempt to repair corrupted video files by re-muxing, re-encoding, or fixing the `moov atom`.
- **Deep File Metadata Extraction**: Pulls detailed EXIF data from images and comprehensive shell properties from any file type.
- **3D Media Cube**: Display images or videos on the faces of an interactive, rotating 3D cube.
- **Continuous Media Scroller**: Create a seamless horizontal or vertical scrolling display of images and videos.
- **Flexible Video Playback**: Scripts are available in two flavors: one using **FFmpeg** for broad format support and another using the native Windows **MediaElement** for simplicity.
- **User-Friendly GUIs**: All tools are wrapped in intuitive graphical interfaces built with Windows Forms and WPF.

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

- **`Show-RotatingImageVideoCubeFfmpeg.ps1`**
  - **Description**: Displays selected images and videos on the faces of a rotating 3D cube using FFmpeg for broad format support.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ImagesVideosButterflyEffect.ps1`**
  - **Description**: Creates a dynamic visual display featuring six 3D planes that move around the screen in a butterfly-like pattern, with each face independently playing media from a user-selected playlist.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ScrollingImagesVideosHorizontalFfmpeg.ps1`**
  - **Description**: Creates a continuous horizontal-scrolling display of selected images and videos using FFmpeg.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

- **`Show-ScrollingImagesVideosVerticalFfmpeg.ps1`**
  - **Description**: Creates a continuous vertical-scrolling display of selected images and videos using FFmpeg.
  - **Dependencies**: `ffmpeg.exe`, `ffplay.exe`.

### Media Viewers (MediaElement-Based)

- **`Show-RotatingImageVideoCubeMediaElement.ps1`**
  - **Description**: Displays selected images and videos on a rotating 3D cube using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ScrollingImagesVideosHorizontalMediaElement.ps1`**
  - **Description**: Creates a horizontal-scrolling display using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

- **`Show-ScrollingImagesVideosVerticalMediaElement.ps1`**
  - **Description**: Creates a vertical-scrolling display using the native Windows MediaElement.
  - **Dependencies**: PowerShell with .NET/WPF access.

## Getting Started

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

### Usage

1. Place all 9 `.ps1` script files in the same directory.
1. Place all 10 `.ps1` script files in the same directory.
2. Ensure all prerequisites are met.
3. Run the main launcher script from a PowerShell terminal:

    ```powershell
    .\Show-ScriptLauncher.ps1
    ```

4. The launcher GUI will appear, showing buttons for each available script. Buttons for scripts with missing dependencies will be disabled.
5. Click a button to launch the desired tool.

## Author

- **JD Alberthal**
  - Website: jdalberthal.com
  - GitHub: @jdalberthal
  - Email: `jd@jdalberthal.com`

## License

This project is licensed under the MIT License - see the `LICENSE.md` file for details.
