<#
.SYNOPSIS
    A GUI-based tool to scan, analyze, and attempt repairs on video files using FFmpeg.

.DESCRIPTION
    This script provides a comprehensive suite of tools for video file maintenance, wrapped in a
    user-friendly GUI. It leverages the power of the FFmpeg toolkit (ffmpeg, ffprobe, ffplay) and
    recover_mp4.exe to perform various checks and repairs.

    Key Features:
    - Multiple Scan Modes:
        - General Corruption: Decodes video files to detect a wide range of errors.
        - Container Mismatch: Verifies that a file's extension (e.g., .mkv) matches its internal
          container format.
        - Moov Atom Check: For MP4/MOV files, verifies the 'moov atom' is at the beginning for
          fast streaming and playback.
    - Asynchronous Operations: Scans run in the background, keeping the UI responsive.
    - Repair Attempts: Offers context-sensitive repair options based on the type of error
      detected, such as re-muxing, re-encoding, or fixing the moov atom position.
    - Detailed Analysis: Provides a deep-dive view of a single video's format, streams, and metadata.
    - Multi-Video Player: Allows for simultaneous playback of up to four videos for comparison.

.EXAMPLE
    PS C:\> .\Get-VideoCorruptionGPUFfmpeg.ps1

    Launches the main GUI. From there, select a path to scan and choose one of the available
    scan types. Results are displayed in a new interactive grid.

.NOTES
    Name:           Get-VideoCorruptionGPUFfmpeg.ps1
    Version:        1.0.0, 10/18/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffprobe.exe, ffplay.exe): https://www.ffmpeg.org/download.html
                    - recover_mp4.exe: https://www.videohelp.com/software/recover-mp4-to-h264
#>
Clear-Host
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type -AssemblyName WindowsFormsIntegration, System.Xaml
[System.Windows.Forms.Application]::EnableVisualStyles()

$ExternalButtonName = "Get Video Corruption"
$ScriptDescription = "A comprehensive tool to scan video files for various types of corruption, check container/extension mismatches, and verify the 'moov atom' position. It also provides options to attempt repairs on corrupted files."
$RequiredExecutables = @("ffmpeg.exe", "ffprobe.exe", "ffplay.exe", "recover_mp4.exe")

# --- Dependency Check ---
if ($RequiredExecutables)
{
    $dependencyStatus = @()
    $allDependenciesMet = $true

    # First, check all dependencies without writing to the console yet
    foreach ($exe in $RequiredExecutables)
    {
        # Check in the PATH and in the script's local directory ($PSScriptRoot)
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

    # If any dependency is missing, then write the status of all of them
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

# --- Function Definitions ---
# Define all functions that will be used by the background runspace here.
# IMPORTANT: These functions MUST use Write-Progress for status updates. They cannot access UI controls directly.
function RetrieveOnly
{
    param (
        $ScanPath,
        $IsRescan,
        $ErrorsToMatch,
        [switch]$Recursive
    )
    $FfmpegResults = @()

    $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
    "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
    "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
    "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
    "*.ts", "*.vob"

    Write-Progress -Activity "Scanning Files" -Status "Getting file list..."
    $gciParams = @{ File = $true; Include = $VideoExtensions }
    if ($Recursive.IsPresent)
    {
        $gciParams.Path = $ScanPath
        $gciParams.Recurse = $true
    }
    else
    {
        $gciParams.Path = (Join-Path $ScanPath "*")
    }
    $AllFiles = Get-ChildItem @gciParams

    $fileCount = $AllFiles.Count
    $processedCount = 0

    foreach($A in $AllFiles)
    {
        $processedCount++
        $percentComplete = if ($fileCount -gt 0) { ($processedCount / $fileCount) * 100 } else { 0 }
        $remaining = $fileCount - $processedCount + 1
        Write-Progress -Activity "Processing" -Status "$($A.Name)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $fileCount"

        $Folder = (Split-Path $A.FullName)
        $File = (Split-Path $A.FullName -Leaf)
                
        $ErrorResult = $false
        $ProbeResult = $null
        $Result = $null
        
        try
        {
            $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_entries format -i ($A.FullName) | ConvertFrom-Json)
            $FormatEntries = $FfprobeOutput.format
            $ProbeResult = "$($FormatEntries.format_name)"
                    
            $VideoDuration = $FormatEntries.duration
            $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
            $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")

            $SizeMB = [math]::Round($A.Length / 1MB, 2)

            $Result = "Scan not performed."   

            $FfmpegResults += [PSCustomObject]@{
                Error     = $ErrorResult
                File      = $File
                Folder    = $Folder
                Extension = $A.Extension
                SizeMB    = $SizeMB
                Duration  = $FormattedTime
                Results   = $Result
            }
        }
        catch
        {
            Write-Progress -Activity "Error" -Status "Skipping $($A.Name): $($_.Exception.Message)" -PercentComplete $percentComplete
        }
    }
 
    return @{ScanType = "RetrieveOnly"; Results = $FfmpegResults }
}

function CompareContainersAndExtensions
{
    param (
        $ScanPath,
        $IsRescan,
        $ErrorsToMatch,
        [switch]$Recursive
    )
    $FfmpegResults = @()

    $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
    "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
    "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
    "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
    "*.ts", "*.vob"

    Write-Progress -Activity "Scanning Files" -Status "Getting file list..."

    # Using a hashtable for lookups is much more efficient and readable than a long if/elseif chain.
    $containerMappings = @{
        ".wmv"   = "asf"
        ".mpg"   = "mpeg|mpegvideo"
        ".mkv"   = "matroska,webm"
        ".f4v"   = "mov,mp4,m4a,3gp,3g2,mj2"
        ".m2ts"  = "mpegts"
        ".m4v"   = "mov,mp4,m4a,3gp,3g2,mj2"
        ".mjpeg" = "jpeg_pipe"
        ".mts"   = "mpegts"
        ".ogv"   = "ogg"
        ".vob"   = "mpeg"
        ".m2v"   = "mpegvideo"
    }
    # Create a regex pattern for all other extensions that should match their container name.
    $otherExtensionsPattern = 'avi|qt|yuv|rm|amv|mp4|m4p|mpg|mp2|mpeg|mpe|mpv|svi|3gp|3g2|mxf|roq|nsv|flv|f4a|f4b|mod|wtv|hevc|ts'

    $gciParams = @{ File = $true; Include = $VideoExtensions }
    if ($Recursive.IsPresent)
    {
        $gciParams.Path = $ScanPath
        $gciParams.Recurse = $true
    }
    else
    {
        $gciParams.Path = (Join-Path $ScanPath "*")
    }
    $AllFiles = Get-ChildItem @gciParams

    $fileCount = $AllFiles.Count
    $processedCount = 0

    foreach($A in $AllFiles)
    {
        $processedCount++
        $percentComplete = if ($fileCount -gt 0) { ($processedCount / $fileCount) * 100 } else { 0 }
        $remaining = $fileCount - $processedCount + 1
        Write-Progress -Activity "Processing" -Status "$($A.Name)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $fileCount"

        $Folder = (Split-Path $A.FullName)
        $File = (Split-Path $A.FullName -Leaf)
                
        $ErrorResult = $false
        $ProbeResult = $null
        $Result = $null
        
        try
        {
            $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_entries format -i ($A.FullName) | ConvertFrom-Json)
            $FormatEntries = $FfprobeOutput.format
            $ProbeResult = "$($FormatEntries.format_name)"
                    
            $VideoDuration = $FormatEntries.duration
            $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
            $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")

            $SizeMB = [math]::Round($A.Length / 1MB, 2)

            $extension = $A.Extension
            # Check if the extension has a specific mapping.
            if ($containerMappings.ContainsKey($extension))
            {
                $expectedContainer = $containerMappings[$extension]
                if ($ProbeResult -match $expectedContainer)
                {
                    # It's a match, but we can add the extension for clarity if it's not already there.
                    if ($ProbeResult -notmatch $extension.TrimStart('.'))
                    {
                        $ProbeResult = "$($ProbeResult),$($extension.TrimStart('.'))"
                    }
                }
                else
                {
                    # It's a mismatch.
                    $ErrorResult = $true
                }
            }
            elseif ($ProbeResult -notmatch $extension.TrimStart('.'))
            {
                # For all other standard extensions, the container should match the extension name.
                $ErrorResult = $true
            }
            $Result = $ProbeResult
                    

            $FfmpegResults += [PSCustomObject]@{
                Error     = $ErrorResult
                File      = $File
                Folder    = $Folder
                Extension = $A.Extension
                SizeMB    = $SizeMB
                Duration  = $FormattedTime
                Results   = $Result -join ", "
            }
        }
        catch
        {
            Write-Progress -Activity "Error" -Status "Skipping $($A.Name): $($_.Exception.Message)" -PercentComplete $percentComplete
        }
    }
 
    return @{ScanType = "CompareContainersAndExtensions"; Results = $FfmpegResults }
}

function GetMoovAtom
{
    param (
        $ScanPath,
        $IsRescan,
        $ErrorsToMatch,
        [switch]$Recursive
    )
    $FfmpegResults = @()

    $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
    "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
    "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
    "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
    "*.ts", "*.vob"
    
    Write-Progress -Activity "Scanning Files" -Status "Getting file list..."
    $gciParams = @{ File = $true; Include = $VideoExtensions }
    if ($Recursive.IsPresent)
    {
        $gciParams.Path = $ScanPath
        $gciParams.Recurse = $true
    }
    else
    {
        $gciParams.Path = (Join-Path $ScanPath "*")
    }
    $AllFiles = Get-ChildItem @gciParams
    $fileCount = $AllFiles.Count
    $processedCount = 0

    foreach($A in $AllFiles)
    {
        $processedCount++
        $percentComplete = if ($fileCount -gt 0) { ($processedCount / $fileCount) * 100 } else { 0 }
        $remaining = $fileCount - $processedCount + 1
        Write-Progress -Activity "Processing" -Status "$($A.Name)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $fileCount"

        $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_entries format -i ($A.FullName) | ConvertFrom-Json)
        $FormatEntries = $FfprobeOutput.format
        $VideoDuration = $FormatEntries.duration
        $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
        $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")
        $SizeMB = [math]::Round($A.Length / 1MB, 2) 

        $Folder = (Split-Path $A.FullName)
        $File = (Split-Path $A.FullName -Leaf)
        
        if($A.Extension -notmatch ".mp4|.mov|.qt|.3gp|.m4v|.3g2 ")
        {
            $Result = "Not applicable for this file type."
            $ErrorResult = $false
        }
        else
        {
            # This new method is more robust. It starts ffmpeg and reads the error stream line-by-line.
            # It stops as soon as it finds the first 'moov' or 'mdat' atom.
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "ffmpeg.exe"
            $psi.Arguments = "-v trace -i `"$($A.FullName)`" -f null -"
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $psi
            $proc.Start() | Out-Null

            $firstAtom = $null
            while (-not $proc.StandardError.EndOfStream)
            {
                $line = $proc.StandardError.ReadLine()
                if ($line -match "type:'moov'")
                {
                    $firstAtom = 'moov'
                    break
                }
                if ($line -match "type:'mdat'")
                {
                    $firstAtom = 'mdat'
                    break
                }
            }
            if (-not $proc.HasExited) { $proc.Kill() }

            if ($firstAtom -eq 'moov')
            {
                $Result = "The 'moov' atom is before the 'mdat' atom (faststart)."
                $ErrorResult = $false
            }
            elseif ($firstAtom -eq 'mdat')
            {
                $Result = "The 'moov' atom is after the 'mdat' atom (not faststart)."
                $ErrorResult = $true
            }
            else
            {
                $Result = "The 'moov' atom was NOT detected during FFmpeg trace."
                $ErrorResult = $true
            }
        }

        $FfmpegResults += [PSCustomObject]@{
            Error     = $ErrorResult
            File      = $File
            Folder    = $Folder
            Extension = $A.Extension
            SizeMB    = $SizeMB
            Duration  = $FormattedTime
            Results   = $Result
        }
    }

    return @{ScanType = "GetMoovAtom"; Results = $FfmpegResults } 
}

function GetSomeVideoCorruption
{
    param (
        $ScanPath,
        $IsRescan,
        $ErrorsToMatch,
        [switch]$Recursive
    )
    $FfmpegResults = @()

    $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
    "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
    "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
    "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
    "*.ts", "*.vob"

    Write-Progress -Activity "Scanning Files" -Status "Getting file list..."
    $gciParams = @{ File = $true; Include = $VideoExtensions }
    if ($Recursive.IsPresent)
    {
        $gciParams.Path = $ScanPath
        $gciParams.Recurse = $true
    }
    else
    {
        $gciParams.Path = (Join-Path $ScanPath "*")
    }
    $AllFiles = Get-ChildItem @gciParams
    $fileCount = $AllFiles.Count
    $processedCount = 0

    foreach($A in $AllFiles)
    {
        $processedCount++
        $percentComplete = if ($fileCount -gt 0) { ($processedCount / $fileCount) * 100 } else { 0 }
        $remaining = $fileCount - $processedCount + 1
        Write-Progress -Activity "Processing" -Status "$($A.Name)" -PercentComplete $percentComplete -CurrentOperation "$processedCount of $fileCount"
        
        $Folder = (Split-Path $A.FullName)
        $File = (Split-Path $A.FullName -Leaf)

        $ErrorResult = $false
        $ProbeResult = $null
        $Result = $null

        $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_entries format -i ($A.FullName) | ConvertFrom-Json)
        $FormatEntries = $FfprobeOutput.format
        $ProbeResult = "$($FormatEntries.format_name)"
                
        $VideoDuration = $FormatEntries.duration
        $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
        $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")
        $SizeMB = [math]::Round($A.Length / 1MB, 2) 

        $ffmpegOutput = ffmpeg.exe -hwaccel auto -v error -i ($A.FullName) -f null - 2>&1 
        # Process each line of the output:
        # 1. Find lines with error keywords.
        # 2. For each line, remove the `[...]` prefix and the variable data after the last colon.
        # 3. Get the unique, cleaned-up error strings.
        $ErrorLines = $ffmpegOutput | Select-String -Pattern $ErrorsToMatch | ForEach-Object {
            $line = $_.Line -replace '\[.*?\]\s*' # Remove all occurrences of [...] prefixes

            # Special handling for DTS errors to preserve the stream number
            if ($line -match "non monotonically increasing dts to muxer in stream")
            {
                $line = $line -replace '(:).*', '$1' # Keep the stream number but remove the values after it
            }
            else
            {
                # General cleanup for other errors
                $line = $line -replace ':\s*.*$' `
                    -replace '\(.*?\)|\d+$' `
                    -replace '\s[\d\.]+\s', ' ' `
                    -replace ' with size$'
            }
            $line.Trim().TrimEnd(':. ').Trim()
        } | Select-Object -Unique
        
        if ($ErrorLines)
        {
            $ErrorResult = $true
            $Result = $ErrorLines -join [Environment]::NewLine
        }
        else
        {
            $ErrorResult = $false
            $Result = "No errors found."
        }

        $FfmpegResults += [PSCustomObject]@{
            Error     = $ErrorResult
            File      = $File
            Folder    = $Folder
            Extension = $A.Extension
            SizeMB    = $SizeMB
            Duration  = $FormattedTime
            Results   = $Result
        }
    }
          
    return @{ScanType = "GetSomeVideoCorruption"; Results = $FfmpegResults } 
}

function GetSingleVideoDetails
{
    param (
        $ErrorsToMatch
    )
    $FfmpegResults = @()
    $Path = $null 
    
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Title = "Select a File"
    $FileBrowser.InitialDirectory = [Environment]::GetFolderPath("MyComputer")
    $FileBrowser.Filter = "Video Files (*.mp4, *.mkv, *.avi, *.mov, *.wmv, *.flv, *.webm)|*.mp4;*.mkv;*.avi;*.mov;*.wmv;*.flv;*.webm|All files (*.*)|*.*"
    $FileBrowser.Multiselect = $false

    if ($FileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Path = $FileBrowser.FileName
    }
    else
    {
        return
    }
    # Since this function runs on the UI thread, write directly to the StatusTextBox.
    $StatusTextBox.AppendText("Processing File - Processing: $(Split-Path $Path -Leaf)`r`n")
        
    $Folder = (Split-Path $Path)
    $File = (Split-Path $Path -Leaf)

    $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_format -show_streams -i $Path | ConvertFrom-Json)
    $FormatEntries = $ffprobeOutput.format
    $VideoDuration = $FormatEntries.duration
    $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
    $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")

    $SizeMB = [math]::Round((Get-Item $Path).Length / 1MB, 2) 

    $ffmpegOutput = ffmpeg.exe -hwaccel auto -v error -i "$Path" -f null - 2>&1 
    # Process each line of the output:
    # 1. Find lines with error keywords.
    # 2. For each line, remove the `[...]` prefix and the variable data after the last colon.
    # 3. Get the unique, cleaned-up error strings.
    $ErrorLines = $ffmpegOutput | Select-String -Pattern $ErrorsToMatch | ForEach-Object {
        $line = $_.Line -replace '\[.*?\]\s*' # Remove all occurrences of [...] prefixes

        # Special handling for DTS errors to preserve the stream number
        if ($line -match "non monotonically increasing dts to muxer in stream")
        {
            $line = $line -replace '(:).*', '$1' # Keep the stream number but remove the values after it
        }
        else
        {
            # General cleanup for other errors
            $line = $line -replace ':\s*.*$' `
                -replace '\(.*?\)|\d+$' `
                -replace '\s[\d\.]+\s', ' ' `
                -replace ' with size$'
        }
        $line.Trim().TrimEnd(':. ').Trim()
    } | Select-Object -Unique

    if ($ErrorLines)
    {
        $ErrorResult = $true
        $Result = $ErrorLines -join [Environment]::NewLine
    }
    else
    {
        $ErrorResult = $false
        $Result = "No errors found."
    }

    $Output = New-Object System.Collections.ArrayList

    $Output.Add([PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "General"; Name = "Error Detection -->"; Value = $Result }) | Out-Null


    $isEmpty = $true
    foreach ($property in $ffprobeOutput.PSObject.Properties)
    {
        if ($property.Value -ne $null -and $property.Value -ne "")
        {
            $isEmpty = $false
            break
        }
    }

    if ($isEmpty -ne $true)
    {
        $FormatProperties = $ffprobeOutput.format | Get-Member -MemberType NoteProperty
        $videoStream = $ffprobeOutput.streams | Where-Object { $_.codec_type -eq 'video' }
        $audioStream = $ffprobeOutput.streams | Where-Object { $_.codec_type -eq 'audio' }


        if ($videoStream) { $VideoProperties = $videoStream | Get-Member -MemberType NoteProperty }
        if ($audioStream) { $AudioProperties = $audioStream | Get-Member -MemberType NoteProperty }

        $Format = foreach ($property in $FormatProperties)
        {
            $name = $property.Name
            $value = $ffprobeOutput.format.$name
            [PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "Format"; Name = $name; Value = $value }
        }
        $Output.AddRange($Format)

        $Video = @(foreach ($property in $VideoProperties)
            {
                $name = $property.Name
                $value = $videoStream.$name
                [PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "Video"; Name = $name; Value = $value }
            })
        $Output.AddRange($Video)

        $Audio = @(foreach ($property in $AudioProperties)
            {
                $name = $property.Name
                $value = $audioStream.$name
                [PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "Audio"; Name = $name; Value = $value }
            })
        $Output.AddRange($Audio)

        if ($videoStream) { $VideoDispositionProperties = $videoStream.disposition | Get-Member -MemberType NoteProperty }
        if ($audioStream) { $AudioDispositionProperties = $audioStream.disposition | Get-Member -MemberType NoteProperty }

        $VideoDispositionData = @(foreach ($property in $VideoDispositionProperties)
            {
                $name = $property.Name
                $value = $videoStream.disposition.$name
                [PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "Video Disposition"; Name = $name; Value = $value }
            })
        $Output.AddRange($VideoDispositionData)

        $AudioDispositionData = @(foreach ($property in $AudioDispositionProperties)
            {
                $name = $property.Name
                $value = $audioStream.disposition.$name
                [PSCustomObject]@{ File = $File; Folder = $Folder; Extension = [System.IO.Path]::GetExtension($Path); SizeMB = $SizeMB; Duration = $FormattedTime; Section = "Audio Disposition"; Name = $name; Value = $value }
            })
        $Output.AddRange($AudioDispositionData)
    }
    $FfmpegResults = $Output
    return @{ScanType = "GetSingleVideoDetails"; Results = $FfmpegResults } 
}

function Test-PathContains
{
    param (
        [string]$ParentPath,
        [string]$ChildPath
    )

    # Ensure paths are not null or empty before proceeding
    if ([string]::IsNullOrEmpty($ParentPath) -or [string]::IsNullOrEmpty($ChildPath))
    {
        return $false
    }

    $fullParentPath = [System.IO.Path]::GetFullPath($ParentPath).TrimEnd('\/')
    $fullChildPath = [System.IO.Path]::GetFullPath($ChildPath).TrimEnd('\/')

    # Case-insensitive comparison
    return $fullChildPath.ToLower().StartsWith($fullParentPath.ToLower())
}


function Start-RepairAttempt
{
    param (
        $SyncHash,
        $CheckedItems
    )

    $SkippedFiles = @()
    
    # This function now runs in the background, so it uses Write-Progress for UI updates.
    Write-Progress -Activity "Preparing Repair" -Status "Initializing..."

    if ($CheckedItems.Count -eq 0)
    {
        # This case should be handled by Start-AsyncRepair, but as a safeguard:
        Write-Warning "No items were passed to the repair function."
        return
    }

    if (-not (Test-Path -Path $SyncHash.RepairOutputPath))
    {
        New-Item -ItemType Directory -Path $SyncHash.RepairOutputPath -Force | Out-Null
    }

    $Fixed = 0
    $processedCount = 0
    $fileCount = $CheckedItems.Count

    foreach ($C in $CheckedItems)
    {
        $processedCount++
        $percentComplete = ($processedCount / $fileCount) * 100
        Write-Progress -Activity "Repairing" -Status "Processing: $($C.File)" -PercentComplete $percentComplete

        $FileToRepair = Join-Path $C.Folder $C.File
        $FullRepairOutputPath = Join-Path $SyncHash.RepairOutputPath $C.File
        $CurrentExtension = $C.Extension
        $CurrentResults = $C.Results
        
        # Initialize variables with safe defaults
        $PsVideoCodec, $PsAvgFrameRate, $PsAudioCodec, $PsAudioBitRate, $FeVideoBitRate = $null, $null, $null, "0k", "0k"

        try
        {
            $ProbeStreams = ffprobe.exe -v quiet -print_format json -show_streams -i $FileToRepair 2>$null | ConvertFrom-Json
            $ProbeFormat = (ffprobe.exe -v quiet -print_format json -show_entries format -i $FileToRepair 2>$null | ConvertFrom-Json)

            $VideoStream = $ProbeStreams.streams | Where-Object { $_.codec_type -eq "video" } | Select-Object -First 1
            $AudioStream = $ProbeStreams.streams | Where-Object { $_.codec_type -eq "audio" } | Select-Object -First 1
            $FormatEntries = $ProbeFormat.format

            $PsVideoCodec = $VideoStream.codec_name
            $PsAvgFrameRate = $VideoStream.avg_frame_rate # or r_frame_rate
            $PsAudioCodec = $AudioStream.codec_name

            if ($AudioStream.bit_rate -gt 0)
            {
                $PsAudioBitRate = ([Math]::Round($AudioStream.bit_rate / 1000).ToString()) + "k"
            }
            if ($FormatEntries.duration -gt 0)
            {
                $FeVideoBitRate = ([Math]::Round(($FormatEntries.size * 8) / $FormatEntries.duration / 1000).ToString()) + "k"
            }
        }
        catch
        {
            # ffprobe failed, proceed with default values. This is expected for some corrupt files.
            Write-Warning "Could not get media info for '$($C.File)'. Proceeding with repair attempt using default parameters."
        }

        Remove-Item -Path "$FullRepairOutputPath" -Force -ErrorAction SilentlyContinue

        switch ($SyncHash.ScanType)
        {
            "CompareContainersAndExtensions"
            {
                $format = ffprobe.exe -v error -select_streams v:0 -show_entries format='format_name' -of default=noprint_wrappers=1:nokey=1 "$FileToRepair" 2>$null
                $TrimmedExtension = ($CurrentExtension.TrimStart("."))
                if($CurrentResults -notmatch $TrimmedExtension)
                {
                    if ([string]::IsNullOrWhiteSpace($format))
                    {
                        $NewExtension = ".$($TrimmedExtension).unknown"
                        $reason = "Could not determine container format (likely corrupt). Renamed to '$($C.File).unknown'"
                        $SkippedFiles += "File: $($C.File)`nReason: $Reason`n`n"
                    }
                    else
                    {
                        $Fixed += 1
                        $possibleFormats = @($format -split ',' | ForEach-Object { $_.Trim() })
                        if ($possibleFormats -contains 'mp4')
                        {
                            $NewExtension = '.mp4'
                        }
                        else
                        {
                            # Fallback to the first format if mp4 is not an option
                            $NewExtension = '.' + $possibleFormats[0]
                        }
                    }
                    $NewFile = [System.IO.Path]::ChangeExtension($FullRepairOutputPath, $NewExtension)
                    try { Copy-Item -Path $FileToRepair -Destination $NewFile -ErrorAction Stop }
                    catch { Write-Error "Error during copy operation for $($C.File): $_" }
                }
                else
                {
                    $reason = "Extension '$($C.Extension)' already matches container '$($C.Results)'."
                    $SkippedFiles += "File: $($C.File)`nReason: $Reason`n`n"
                }
            }
            "GetMoovAtom"
            {
                # Only attempt a repair if the moov atom is explicitly reported as being at the end (not faststart).
                if ($CurrentResults -eq "The 'moov' atom is after the 'mdat' atom (not faststart).")
                {
                    $Fixed += 1
                    try { ffmpeg.exe -hwaccel auto -y -i "$FileToRepair" -c:v copy -c:a copy -movflags +faststart "$FullRepairOutputPath" 2>&1 | Out-Null }
                    catch { Write-Error "Error during moov atom operation for $($C.File): $_" }
                }
                else
                {
                    $SkippedFiles += "$($C.File) - $($C.Results)"
                }
            }
            { ($_ -match "GetSomeVideoCorruption" ) -or ($_ -match "GetSingleVideoDetails") }
            {
                if ( ($CurrentExtension -match "mpg") -and (($CurrentResults -match "header missing") -and ($CurrentResults -match "Error submitting packet to decoder")))
                {
                    $Fixed += 1
                    ffmpeg -i "$FileToRepair" -c:v mpeg2video -q:v 1 -b:v 2500k -c:a mp3 -b:a 192k "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "Invalid NAL unit size|missing picture in access unit with size")
                {
                    $Fixed += 1
                    $TooReplace = $FileToRepair.Substring($($FileToRepair.LastIndexOf(".")))
                    $FullRepairOutputPathOne = $FullRepairOutputPath.Replace($TooReplace, "-Option01$TooReplace")
                    $FullRepairOutputPathTwo = $FullRepairOutputPath.Replace($TooReplace, "-Option02$TooReplace")
                    Set-Location -Path $SyncHash.RepairOutputPath
                    recover_mp4.exe $FileToRepair --analyze
                    recover_mp4.exe $FileToRepair recovered.h264 recovered.aac
                    ffmpeg.exe -hwaccel auto -i recovered.h264 -c:v $SyncHash.Gpu -preset p7 -cq:v 33 -rc:v vbr -movflags +faststart temp-recovered.h264 2>&1 | Out-Null
                    ffmpeg.exe -hwaccel auto -i temp-recovered.h264 -i recovered.aac -bsf:a aac_adtstoasc -c:v copy -c:a copy "$FullRepairOutputPathOne" 2>&1 | Out-Null
                    ffmpeg.exe -hwaccel auto -i "$FullRepairOutputPathOne" -c:v $SyncHash.Gpu -preset p7 -cq:v 30 -rc:v vbr -movflags +faststart "$FullRepairOutputPathTwo" 2>&1 | Out-Null
                    Remove-Item -Path .\recovered.h264, .\recovered.aac, .\video.hdr, .\audio.hdr, .\temp-recovered.h264 -ErrorAction SilentlyContinue
                }
                elseif ($CurrentResults -match "Header missing|len -1 invalid|Error submitting packet to decoder|ac-tex damaged at|invalid new backstep|cabac decode of qscale diff failed at|error while decoding MB")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -c:v $SyncHash.Gpu -preset p7 -cq:v 33 -rc:v vbr -movflags +faststart "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif (($CurrentResults -match "Number of bands.*exceeds limit") -or ($CurrentResults -match "exceeds limit.*Number of bands"))
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -c:v copy -c:a aac -b:a $PsAudioBitRate "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "Rematrix is needed")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -c:v copy -c:a aac -b:a 192k "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "Missing AMF_END_OF_OBJECT")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -analyzeduration 10M -probesize 10M -i "$FileToRepair" -c copy "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "incomplete frame")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -b:v "$FeVideoBitRate" -qscale:v 2 -async 0 -vsync 1 "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "Error at MB|Packet mismatch|illegal ac vlc code at")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -c:v $SyncHash.Gpu -preset p7 -cq:v 33 -rc:v vbr -c:a aac "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif($CurrentResults -match "non monotonically increasing dts to muxer in stream 0")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -fflags +genpts -i "$FileToRepair" -r $PsAvgFrameRate -c:v $SyncHash.Gpu -preset p7 -cq:v 33 -rc:v vbr -c:a copy -movflags +faststart "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                elseif ($CurrentResults -match "non monotonically increasing dts to muxer in stream 1")
                {
                    $Fixed += 1
                    ffmpeg.exe -hwaccel auto -i "$FileToRepair" -fflags +genpts+igndts -c:v copy -c:a "$PsAudioCodec" -b:a "$PsAudioBitRate"  -movflags +faststart "$FullRepairOutputPath" 2>&1 | Out-Null
                }
                else { $SkippedFiles += "$($C.File) - No repair known for error: $($C.Results)" }
            }
        }
    }
    
    # Return a result object that the UI thread can process
    return @{
        FixedCount   = $Fixed
        SkippedFiles = $SkippedFiles
    }
}

# --- Global Variables & Initial Setup ---
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.ScanPath = $null
$SyncHash.ScanType = ""
$SyncHash.ScriptRoot = $PSScriptRoot
$SyncHash.RescanPath = ""
$SyncHash.IsRescan = $false
$SyncHash.RepairOutputPath = $null
$SyncHash.ErrorsToMatch = "error|missing|invalid|corrupt|exceeds limit|not allocated|Input buffer exhausted|Prediction is not allowed|non-existing PPS|POCs unavailable|damaged|failed|Packet mismatch|Invalid NAL unit size|Error splitting the input|Reserved bit set|ms_present = is reserved|skip_data_stream_element|SBR was found|decode_pce|Pulse tool not allowed|TYPE_FIL|Rematrix is needed|Failed to configure output pad|Error reinitializing filters|Task finished with error code|Terminating thread with return code|Cannot determine format of input|Nothing was written into output file|AMF_END_OF_OBJECT|channel element is not allocated|invalid band type"

$gpuInfo = Get-WmiObject Win32_VideoController | Select-Object -ExpandProperty Description

# Determine the GPU type (NVIDIA, AMD, or other)
if ($gpuInfo -like "*NVIDIA*") { $gpuType = "NVIDIA" }
elseif ($gpuInfo -like "*AMD*") { $gpuType = "AMD" }
else { $gpuType = "Other" }

# Add hardware acceleration based on GPU type
if ($gpuType -eq "NVIDIA") { $SyncHash.Gpu = "h264_nvenc" }
elseif ($gpuType -eq "AMD") { $SyncHash.Gpu = "h264_amf" }
else { $SyncHash.Gpu = "libx264" }

# --- UI Form Definitions ---
$SyncHash.NewResultsForm = 
{
    param (
        $ScanPath = $SyncHash.ScanPath,
        $SyncHash = $SyncHash,
        $StatusTextBox = $SyncHash.StatusTextBox,
        $ScriptRoot = $SyncHash.ScriptRoot,
        $DataGridView = $SyncHash.DataGridView,
        $FfmpegResults
    ) 

    $arrayList = New-Object System.Collections.ArrayList
    if ($null -ne $FfmpegResults)
    {
        $arrayList.AddRange($FfmpegResults)
    }

    $ResultsForm = New-Object System.Windows.Forms.Form
    $ResultsForm.Text = "Data Grid View Window"
    $ResultsForm.Size = New-Object System.Drawing.Size(1280, 800)
    $ResultsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $ResultsForm.BackColor = "Navy"

    # Create the Rescan form and its textbox here, where they are used.
    $HelpLabel = New-Object System.Windows.Forms.Label
    $HelpLabel.Text = "F1 - Help"
    $HelpLabel.AutoSize = $true
    $HelpLabel.Location = New-Object System.Drawing.Point(1100, 15)
    $HelpLabel.ForeColor = [System.Drawing.Color]::Yellow
    $HelpLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $ResultsForm.Controls.Add($HelpLabel)

[xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help - Results Grid" Height="550" Width="700" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <RichTextBox x:Name="MyRichTextBox" Grid.Row="0" Margin="5" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
            <FlowDocument>
                <FlowDocument.Resources>
                    <Style TargetType="{x:Type Paragraph}">
                        <Setter Property="Margin" Value="0,0,0,5"/>
                    </Style>
                </FlowDocument.Resources>
                <Paragraph><Bold><Run Text="Results Grid Controls:"/></Bold></Paragraph>
                <Paragraph><Bold>Errors Only:</Bold> Filters the grid to show only rows where an error was detected.</Paragraph>
                <Paragraph><Bold>Show All Results:</Bold> Clears any active filters and shows all rows.</Paragraph>
                <Paragraph><Bold>Show Duplicates:</Bold> (Retrieve/Two Folders scans only) Filters the grid to show only files that appear more than once (i.e., have duplicate filenames).</Paragraph>
                <Paragraph><Bold>Show Unique:</Bold> (Retrieve/Two Folders scans only) Filters the grid to show only files that appear once.</Paragraph>
                <Paragraph><Bold>Export to CSV:</Bold> Exports the currently selected (checked) and visible rows to a CSV file.</Paragraph>
                <Paragraph><Bold>Export Path:</Bold> Allows selecting a different path and filename for the CSV export.</Paragraph>
                <Paragraph><Bold>Attempt Repair on Selected:</Bold> Starts the repair process for the selected (checked) items. This is only visible on the initial scan results.</Paragraph>
                <Paragraph><Bold>Folder to output repair attempts:</Bold> Allows selecting a different folder where repaired files will be saved.</Paragraph>
                <Paragraph><Bold>Copy Repaired:</Bold> (Visible on rescan results) Copies selected repaired files back to their original folders, appending "-Copy" to the filename.</Paragraph>
                <Paragraph><Bold>Replace Original(s):</Bold> (Visible on rescan results) Overwrites the original files with the selected repaired versions. <Bold>This is a destructive action!</Bold></Paragraph>
                <Paragraph><Bold>Play Selected Videos (1-4):</Bold> Opens up to 4 selected videos in a multi-video player for simultaneous viewing.</Paragraph>
                <Paragraph><Bold>Delete Selected:</Bold> Deletes the selected files. By default, it sends them to the Recycle Bin.</Paragraph>
                <Paragraph><Bold>Permanently Delete:</Bold> If checked, the "Delete Selected" button will permanently delete the files, bypassing the Recycle Bin.</Paragraph>
                <Paragraph><Bold>Play (Row Header):</Bold> Clicking the "Play" button in the far-left row header will play that single video file using ffplay.</Paragraph>
            </FlowDocument>
        </RichTextBox>
        
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$ResultsForm.KeyPreview = $True
$ResultsForm.Add_KeyDown({
    param($Sender, $e)
    if ($e.KeyCode -eq 'F1') {
        $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
        $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
        $OkButton = $PopupWindow.FindName("OKButton")
        $OkButton.Add_Click({ $PopupWindow.Close() })
        $PopupWindow.ShowDialog() | Out-Null
    }
})

    $RescanForm = New-Object System.Windows.Forms.Form
    $RescanForm.Text = "Rescanning..."
    $RescanForm.Size = New-Object System.Drawing.Size(400, 400)
    $RescanForm.StartPosition = "CenterScreen"
    $RescanTextBox = New-Object System.Windows.Forms.RichTextBox
    $RescanTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $RescanTextBox.Size = New-Object System.Drawing.Size(360, 340)
    $RescanTextBox.Multiline = $true
    $RescanTextBox.ReadOnly = $true
    $RescanTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $RescanForm.Controls.Add($RescanTextBox)
    $SyncHash.RescanForm = $RescanForm
    $SyncHash.RescanTextBox = $RescanTextBox

    $ErrorsOnly = New-Object System.Windows.Forms.Button
    $ErrorsOnly.Text = "Errors Only"
    $ErrorsOnly.Location = New-Object System.Drawing.Point(20, 10)
    $ErrorsOnly.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($ErrorsOnly)

    $ShowDuplicates = New-Object System.Windows.Forms.Button
    $ShowDuplicates.Text = "Show Duplicates"
    $ShowDuplicates.Location = New-Object System.Drawing.Point(20, 10)
    $ShowDuplicates.Size = New-Object System.Drawing.Size(120, 40)
    $ShowDuplicates.Visible = $false
    $ResultsForm.Controls.Add($ShowDuplicates)

    $ShowDuplicates.Add_Click({

            if($ShowDuplicates.Text -match "Show Duplicates")
            {
                # 1. Get all values from the target column
                $columnValues = ($dataGridView.Rows | ForEach-Object { $_.Cells['File'].Value }).Trim()

                # 2. Identify unique values
                $uniqueValues = $columnValues | Group-Object | Where-Object { $_.Count -le 1 } | Select-Object -ExpandProperty Name

                # 3. Iterate through rows and hide based on uniqueness
                foreach ($row in $dataGridView.Rows)
                {
                    $DataGridView.CurrentCell = $null
                    $currentColumnValue = $row.Cells['File'].Value
                    if ($uniqueValues -contains $currentColumnValue)
                    {
                        $row.Visible = $false
                    }
                }
                $ShowDuplicates.Text = "Show All"
            }
            elseif($ShowDuplicates.Text -match "Show All")
            {
                $DataGridView.CurrentCell = $null
                foreach ($row in $DataGridView.Rows)
                {
                    if (!$DataGridView.SelectedRows.Contains($row))
                    {
                        $row.Visible = $true
                    }
                }
                $ShowDuplicates.Text = "Show Duplicates"
            }
            # Get number of visible rows
            $visibleRows = $dataGridView.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible)

            # Update the Label text
            $RowCount.Text = "$visibleRows"
        })

    $AllResults = New-Object System.Windows.Forms.Button
    $AllResults.Text = "Show All Results"
    $AllResults.Location = New-Object System.Drawing.Point(150, 10)
    $AllResults.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($AllResults)

    $ShowUnique = New-Object System.Windows.Forms.Button
    $ShowUnique.Text = "Show Unique"
    $ShowUnique.Location = New-Object System.Drawing.Point(150, 10)
    $ShowUnique.Size = New-Object System.Drawing.Size(120, 40)
    $ShowUnique.Visible = $false
    $ResultsForm.Controls.Add($ShowUnique)

    $ShowUnique.Add_Click({
            if($ShowUnique.Text -match "Show Unique")
            {
                # 1. Get all values from the target column
                $columnValues = $dataGridView.Rows | ForEach-Object { $_.Cells['File'].Value }

                # 2. Identify unique values
                $uniqueValues = $columnValues | Group-Object | Where-Object { $_.Count -le 1 } | Select-Object -ExpandProperty Name

                # 3. Iterate through rows and hide based on uniqueness
                foreach ($row in $dataGridView.Rows)
                {
                    $DataGridView.CurrentCell = $null
                    $currentColumnValue = $row.Cells['File'].Value
                    if ($uniqueValues -notcontains $currentColumnValue)
                    {
                        $row.Visible = $false
                    }
                }
                $ShowUnique.Text = "Show All"
            }
            elseif($ShowUnique.Text -match "Show All")
            {
                $DataGridView.CurrentCell = $null
                foreach ($row in $DataGridView.Rows)
                {
                    if (!$DataGridView.SelectedRows.Contains($row))
                    {
                        $row.Visible = $true
                    }
                }
                $ShowUnique.Text = "Show Unique"
            }
            # Get number of visible rows
            $visibleRows = $dataGridView.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible)

            # Update the Label text
            $RowCount.Text = "$visibleRows"
        })

    $ExportToCSV = New-Object System.Windows.Forms.Button
    $ExportToCSV.Text = "Export to CSV"
    $ExportToCSV.Location = New-Object System.Drawing.Point(280, 10)
    $ExportToCSV.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($ExportToCSV)

    $SelectExportPath = New-Object System.Windows.Forms.Button
    $SelectExportPath.Text = "Export Path"
    $SelectExportPath.Location = New-Object System.Drawing.Point(410, 10)
    $SelectExportPath.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($SelectExportPath)

    $ExportPath = New-Object System.Windows.Forms.TextBox
    $ExportPath.Location = New-Object System.Drawing.Point(540, 25)
    $ExportPath.Size = New-Object System.Drawing.Size(450, 40)
    $ExportPath.ReadOnly = $true

    # Set a default export path to prevent errors if the user doesn't select one.
    $defaultFileName = switch -Regex ($SyncHash.ScanType)
    {
        "RetrieveOnly"           { "RetrieveOnly.csv" }
        "GetSingleVideoDetails"     { "SingleVideoDetails.csv" }
        "CompareContainersAndExtensions" { "ContainersVsExtensions.csv" }
        "GetMoovAtom"               { "MoovAtom.csv" }
        "GetSomeVideoCorruption" { "SomeVideoCorruption.csv" }
        default                  { "VideoCorruptionReport.csv" }
    }
    $ExportPath.Text = Join-Path $SyncHash.ScriptRoot $defaultFileName

    $ResultsForm.Controls.Add($ExportPath)

    $ExportWarning = New-Object System.Windows.Forms.Label
    $ExportWarning.Location = New-Object System.Drawing.Point(540, 5)
    # $ExportWarning.Size = New-Object System.Drawing.Size(150, 20) # No longer needed
    $ExportWarning.AutoSize = $true
    $ExportWarning.Text = "Only checked items will be exported."
    $ExportWarning.BackColor = [System.Drawing.Color]::Transparent
    $ExportWarning.ForeColor = "Yellow"
    $ExportWarning.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $ResultsForm.Controls.Add($ExportWarning)

    $RowCount = New-Object System.Windows.Forms.Label
    $RowCount.Location = New-Object System.Drawing.Point(1230, 15)
    $RowCount.Size = New-Object System.Drawing.Size(50, 20)
    $RowCount.Text = "0"
    $RowCount.ForeColor = "Yellow"
    $RowCount.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $ResultsForm.Controls.Add($RowCount)

    $SelectRepairPath = New-Object System.Windows.Forms.Button
    $SelectRepairPath.Text = "Folder to output repair attempts"
    $SelectRepairPath.Location = New-Object System.Drawing.Point(20, 710)
    $SelectRepairPath.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($SelectRepairPath)

    $FixAttemptPath = New-Object System.Windows.Forms.TextBox
    $FixAttemptPath.Location = New-Object System.Drawing.Point(155, 720)
    $FixAttemptPath.Size = New-Object System.Drawing.Size(450, 40)
    $FixAttemptPath.ReadOnly = $true # Make it read-only for display

    if($null -eq $SyncHash.RepairOutputPath)
    {
        $SyncHash.RepairOutputPath = $SyncHash.ScriptRoot + "\RepairAttempts"
    }

    $FixAttemptPath.Text = $SyncHash.RepairOutputPath
    $ResultsForm.Controls.Add($FixAttemptPath)

    $AttemptRepair = New-Object System.Windows.Forms.Button
    $AttemptRepair.Text = "Attempt Repair on Selected"
    $AttemptRepair.Location = New-Object System.Drawing.Point(1140, 50)
    $AttemptRepair.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($AttemptRepair)

    $PlayVideos = New-Object System.Windows.Forms.Button
    $PlayVideos.Text = "Play Selected Videos (1-4)"
    $PlayVideos.Location = New-Object System.Drawing.Point(1140, 100)
    $PlayVideos.Size = New-Object System.Drawing.Size(120, 40)
    $ResultsForm.Controls.Add($PlayVideos)

    $ButtonDelete = New-Object System.Windows.Forms.Button
    $ButtonDelete.Text = "Delete Selected"
    $ButtonDelete.Location = New-Object System.Drawing.Point(1100, 708)
    $ButtonDelete.Size = New-Object System.Drawing.Size(160, 50)
    $ResultsForm.Controls.Add($ButtonDelete)

    $PermDelete = New-Object System.Windows.Forms.CheckBox
    $PermDelete.Text = "Permanently Delete"
    $PermDelete.Location = New-Object System.Drawing.Size(980, 720)
    $PermDelete.Size = New-Object System.Drawing.Size(13, 13)
    $PermDelete.Visible = $true
    $PermDelete.ForeColor = [System.Drawing.Color]::White
    $PermDelete.AutoSize = $true
    $ResultsForm.Controls.Add($PermDelete)

    # The label and textbox for repair progress are now below the DataGridView.
    $ProcessingLabel = New-Object System.Windows.Forms.Label
    $ProcessingLabel.Text = "File Being Processed:"
    $ProcessingLabel.Location = New-Object System.Drawing.Point(10, 628)
    $ProcessingLabel.AutoSize = $true
    $ProcessingLabel.ForeColor = "White"
    $ResultsForm.Controls.Add($ProcessingLabel)

    $RepairProgressRTB = New-Object System.Windows.Forms.RichTextBox
    $RepairProgressRTB.Multiline = $true
    $RepairProgressRTB.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $RepairProgressRTB.WordWrap = $true
    $RepairProgressRTB.Location = New-Object System.Drawing.Point(10, 648)
    $RepairProgressRTB.Size = New-Object System.Drawing.Size(1124, 50)
    $RepairProgressRTB.ReadOnly = $true
    $RepairProgressRTB.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ResultsForm.Controls.Add($RepairProgressRTB)

    $SyncHash.RepairProgressRTB = $RepairProgressRTB
    # $SyncHash.Add("RepairProgressRTB", $RepairProgressRTB)

    $ButtonCopy = New-Object System.Windows.Forms.Button
    $ButtonCopy.Text = "Copy Repaired"
    $ButtonCopy.Location = New-Object System.Drawing.Point(625, 708)
    $ButtonCopy.Size = New-Object System.Drawing.Size(160, 50)
    $ResultsForm.Controls.Add($ButtonCopy)

    $ButtonCopy.Add_Click({

            $FixAttemptPath = $FixAttemptPath.Text
            $OriginalScanPath = $SyncHash.OriginalScanPath # This might be null if no scan was run

            # Only run the check if both paths are valid
            $PathStatus = $false
            if ($OriginalScanPath -and $FixAttemptPath) { $PathStatus = Test-PathContains -ParentPath $OriginalScanPath -ChildPath $FixAttemptPath }

            if($PathStatus)
            {
                [System.Windows.Forms.MessageBox]::Show("The output path is a sub-folder of the original`nscanned path. This function will not work.`nOriginals and Fixes will collide." , "Oops!")
            }
            else
            {
                $Continue = [System.Windows.Forms.MessageBox]::Show(
                    "-Copy will be added to the repaired name`nand copied to the original's folder.`nEx: video.mp4 --> video-Copy.mp4`n`nContinue?" ,
                    "Heads Up", 
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning) 

                if($Continue -eq 'Yes')
                {
                    foreach ($row in $dataGridView.Rows)
                    {
                        if (($row.Cells["CheckBoxColumn"].Value) -and ($row.Visible))
                        {
                            $SelectedFiles += (($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value))
                        }
                    }

                    if($SelectedFiles.Count -le 0)
                    {
                        [System.Windows.Forms.MessageBox]::Show("No files selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                    else
                    {
                        $SelectedFiles = @()
                        $NewNameFiles = @()
                        foreach ($row in $dataGridView.Rows)
                        {
                            $Extension = ($row.Cells["Extension"].Value)
                            $CopyName = ($row.Cells["File"].Value).Replace("$Extension", "-Copy$Extension")
                
                            if (($row.Cells["CheckBoxColumn"].Value) -and ($row.Visible))
                            {
                                $SelectedFiles += ($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value)
                                $NewNameFile += (($row.Cells["Folder"].Value) + "\" + $CopyName)
                            }
                        }

                        foreach ($SF in $SelectedFiles)
                        {
                            $FileObject = Get-Item $SF
                            $Folder = $FileObject.Directory
                            $File = $FileObject.Name

                            $FoundFile = Get-ChildItem -Path $OriginalScanPath -Filter $File -Recurse -ErrorAction SilentlyContinue
                            $FoundFilePath = $FoundFile.DirectoryName

                            $Extension = $FileObject.Extension

                            $DestinationFile = "$($FoundFilePath)\$(($File).Replace("$Extension","-Copy$Extension"))"

                            Copy-Item -Path $FileObject -Destination $DestinationFile --force-window-position

                            [System.Windows.Forms.MessageBox]::Show("Copy Complete" , "Done" )
                        }
                    }
                }
            }
        })

    $ButtonOverWrite = New-Object System.Windows.Forms.Button
    $ButtonOverWrite.Text = "Replace Original(s)"
    $ButtonOverWrite.Location = New-Object System.Drawing.Point(805, 708)
    $ButtonOverWrite.Size = New-Object System.Drawing.Size(160, 50)
    $ResultsForm.Controls.Add($ButtonOverWrite)

    $ButtonOverWrite.Add_Click({

            $FixAttemptPath = $FixAttemptPath.Text
            $OriginalScanPath = $SyncHash.OriginalScanPath # This might be null if no scan was run

            # Only run the check if both paths are valid
            $PathStatus = $false
            if ($OriginalScanPath -and $FixAttemptPath) { $PathStatus = Test-PathContains -ParentPath $OriginalScanPath -ChildPath $FixAttemptPath }

            if($PathStatus)
            {
                [System.Windows.Forms.MessageBox]::Show("The output path is a sub-folder of the original`nscanned path. This function will not work.`nOriginals and Fixes will collide." , "Oops!")
            }
            else
            {
                $Continue = [System.Windows.Forms.MessageBox]::Show(
                    "Videos in the original scan path will`nreplaced/over-written!!!`n`nContinue?" ,
                    "Heads Up", 
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning) 

                if($Continue -eq 'Yes')
                {
                    foreach ($row in $dataGridView.Rows)
                    {
                        if (($row.Cells["CheckBoxColumn"].Value) -and ($row.Visible))
                        {
                            $SelectedFiles += (($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value))
                        }
                    }

                    if($SelectedFiles.Count -le 0)
                    {
                        [System.Windows.Forms.MessageBox]::Show("No files selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                    else
                    {
                        $SelectedFiles = @()
                        $NewNameFiles = @()
                        foreach ($row in $dataGridView.Rows)
                        {
                            if (($row.Cells["CheckBoxColumn"].Value) -and ($row.Visible))
                            {
                                $SelectedFiles += ($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value)
                            }
                        }

                        foreach ($SF in $SelectedFiles)
                        {
                            $FileObject = Get-Item $SF
                            $File = $FileObject.Name

                            $FoundFile = Get-ChildItem -Path $OriginalScanPath -Filter $File -Recurse -ErrorAction SilentlyContinue
                            $FoundFilePath = $FoundFile.DirectoryName

                            $DestinationFile = "$($FoundFilePath)\$($File)"

                            Copy-Item -Path $FileObject -Destination $DestinationFile -Force
                        }

                        [System.Windows.Forms.MessageBox]::Show("Replace Complete" , "Done" )
                    }
                }
            }
        })

    # Set a default background color for all buttons on the form
    $ResultsForm.Controls | Where-Object { $_.GetType().Name -eq "Button" } | ForEach-Object {
        $_.BackColor = [System.Drawing.Color]::RoyalBlue
        $_.ForeColor = [System.Drawing.Color]::WhiteSmoke
        $_.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    }

    if($SyncHash.ScanType -match "RetrieveOnly|TwoFolders")
    {
        $ErrorsOnly.Visible = $false
        $AllResults.Visible = $false
        $AttemptRepair.Visible = $false
        $ProcessingLabel.Visible = $false
        $RepairProgressRTB.Visible = $false
        $SelectRepairPath.Visible = $false
        $FixAttemptPath.Visible = $false
        $ShowDuplicates.Visible = $true
        $ShowUnique.Visible = $true
    }

    if($SyncHash.IsRescan -eq $false)
    {
        $ButtonCopy.Visible = $false
        $ButtonOverWrite.Visible = $false
    }
    else
    {
        $AttemptRepair.Visible = $false
    }

    $DataGridView = New-Object System.Windows.Forms.DataGridView
    $DataGridView.Size = New-Object System.Drawing.Size(1124, 568)
    $DataGridView.Location = New-Object System.Drawing.Point(10, 50)
    $DataGridView.AutoGenerateColumns = $true
    # Enable text wrapping for all cells in the DataGridView
    $DataGridView.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    # Set AutoSizeRowsMode to AllCells to adjust row height automatically
    $DataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $DataGridView.RowHeadersVisible = $true
    $DataGridView.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::TopLeft
    $DataGridView.EnableHeadersVisualStyles = $true
    $DataGridView.ColumnHeadersHeight = 40
    $DataGridView.AllowUserToOrderColumns = $true
    $DataGridView.AllowUserToAddRows = $false

    $checkbox1 = New-Object System.Windows.Forms.CheckBox
    $checkbox1.Location = New-Object System.Drawing.Size(113, 73)
    $checkbox1.Size = New-Object System.Drawing.Size(13, 13)
    $checkbox1.Visible = $true
    $checkbox1.add_Click({
            if($checkbox1.Checked)
            {
                for($i = 0; $i -lt $dataGridView.RowCount; $i++)
                {
                    $dataGridView.Rows[$i].Cells[0].Value = $true
                }
            }
            else
            {
                for($i = 0; $i -lt $dataGridView.RowCount; $i++)
                {
                    $dataGridView.Rows[$i].Cells[0].Value = $false
                }
            }
        })
    $ResultsForm.Controls.Add($checkbox1)
    $checkbox1.BringToFront()

    #Handle Events:
    $dataGridView.add_CellContentClick({
            $dataGridView.EndEdit() #otherwise the cell value won't have changed yet
            [System.Windows.Forms.DataGridViewCellEventArgs]$e = $args[1]
            if($e.columnIndex -eq 0)
            {
                if($dataGridView.rows[$e.RowIndex].Cells[$e.ColumnIndex].value -eq $false)
                {
                    $checkbox1.CheckState = 'unchecked'
                }
            }
        })

    $CheckBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $CheckBoxColumn.HeaderText = "Select All`n`n" # Header text for the column
    $CheckBoxColumn.Name = "CheckBoxColumn" # Name for programmatic access
    $CheckBoxColumn.Width = 50 # Set the width of the column
    $CheckBoxColumn.TrueValue = $true # Value when checked
    $CheckBoxColumn.FalseValue = $false # Value when unchecked

    # Add the checkbox column to the DataGridView
    $dataGridView.Columns.Add($CheckBoxColumn)
    $SyncHash.CheckBoxColumn = $CheckBoxColumn

    # Create a new DataTable
    $dataTable = New-Object System.Data.DataTable


    # Define columns in the DataTable
    $dataTable.Columns.Add("Error", [System.String])
    $dataTable.Columns.Add("File", [System.String])
    $dataTable.Columns.Add("Folder", [System.String])
    $dataTable.Columns.Add("Extension", [System.String])
    $dataTable.Columns.Add("SizeMB", [System.String])
    $dataTable.Columns.Add("Duration", [System.String])
    $dataTable.Columns.Add("Section", [System.String])
    $dataTable.Columns.Add("Name", [System.String])
    $dataTable.Columns.Add("Value", [System.String])
    $dataTable.Columns.Add("Results", [System.String])

    # Populate the DataTable from the ArrayList
    foreach ($item in $FfmpegResults)
    {
        $row = $dataTable.NewRow()
        $row.Error = $item.Error
        $row.File = $item.File
        $row.Folder = $item.Folder
        $row.Extension = $item.Extension
        $row.SizeMB = $item.SizeMB
        $row.Duration = $item.Duration
        $row.Section = $item.Section
        $row.Name = $item.Name
        $row.Value = $item.Value
        $row.Results = $item.Results
        $dataTable.Rows.Add($row)
    }
    # Set the DataTable as the DataSource for the DataGridView
    $dataGridView.DataSource = $dataTable

    # $DataGridView.DataSource = $arrayList

    $ResultsForm.Controls.Add($DataGridView)

    $SyncHash.DataGridView = $DataGridView

    $ResultsForm.Add_Load({

            if($SyncHash.ScanType -match "GetSingleVideoDetails")
            {
                $checkbox1.Visible = $false
                $dataGridView.Columns["CheckBoxColumn"].Visible = $false
                $DataGridView.Columns["Error"].Visible = $false
                $DataGridView.Columns["File"].Width = 210
                $DataGridView.Columns["Folder"].Width = 270
                $DataGridView.Columns["Extension"].Width = 60
                $DataGridView.Columns["SizeMB"].Width = 50
                $DataGridView.Columns["Duration"].Width = 60
                $DataGridView.Columns["Section"].Width = 90
                $DataGridView.Columns["Name"].Width = 130 
                $DataGridView.Columns["Value"].Width = 450
                $DataGridView.Columns["Results"].Visible = $false

                $ErrorsOnly.Visible = $false
                $AllResults.Visible = $false
                $PlayVideos.Visible = $false
                $AttemptRepair.Text = "Attempt Repair"
                $ButtonDelete.Text = "Delete Video"
                $ExportWarning.Visible = $false
            }
            else
            {
                $DataGridView.Columns["CheckBoxColumn"].Width = 60
                $DataGridView.Columns["Error"].Width = 40
                $DataGridView.Columns["File"].Width = 210
                $DataGridView.Columns["Folder"].Width = 270
                $DataGridView.Columns["Extension"].Width = 60
                $DataGridView.Columns["SizeMB"].Width = 50
                $DataGridView.Columns["Duration"].Width = 60
                $DataGridView.Columns["Section"].Visible = $false
                $DataGridView.Columns["Name"].Visible = $false
                $DataGridView.Columns["Value"].Visible = $false
                $DataGridView.Columns["Results"].Width = 470                    
            }

            for ($x = 0; $x -lt $SyncHash.DataGridView.RowCount; $x++)
            {
                $ErrorValue = $SyncHash.DataGridView.Rows[$x].Cells["Error"].Value
                $ResultsValue = $SyncHash.DataGridView.Rows[$x].Cells["Results"].Value
                $SingleFileName = $SyncHash.DataGridView.Rows[$x].Cells["Name"].Value
                $SingleFileValue = $SyncHash.DataGridView.Rows[$x].Cells["Value"].Value

                if(($SingleFileName -match "Error Detection -->") -and (-not [string]::IsNullOrEmpty($SingleFileValue)) -and ($SingleFileValue -notmatch "No errors found."))
                { 
                    $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 0) # Red color
                }

                if($ErrorValue -eq "True")
                { 
                    $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 0) # Red color
                }

                if($ResultsValue -match "is before")
                { 
                    $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 0, 255, 0) # Green color
                    # $DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0) # White color
                }
                        
                if($ResultsValue -match "is after")
                {
                    $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 0) # Yellow color
                }
                        
            }

            $HeaderLabel = New-Object System.Windows.Forms.Label
            $HeaderLabel.Text = "Play Video"
            $HeaderLabel.Location = New-Object System.Drawing.Point(10, 10) # Position at top-left
            $HeaderLabel.AutoSize = $true # Ensures the label resizes to fit the text
            $HeaderLabel.BackColor = [System.Drawing.Color]::Transparent
            $SyncHash.DataGridView.Controls.Add($HeaderLabel)

            foreach ($row in  $SyncHash.DataGridView.Rows)
            {
                if ($row.IsNewRow) { continue } # Skip the blank row at the bottom
                $row.HeaderCell.Value = "Play"
            }

            $SyncHash.DataGridView.AutoResizeRowHeadersWidth("AutoSizeToAllHeaders") | Out-Null

            $ErrorsOnly.Add_Click({
                    $SyncHash.DataGridView.CurrentCell = $null
                    foreach ($row in $SyncHash.DataGridView.Rows)
                    {
                        # Check if the row is a new row template, and skip if it is
                        if ($row.IsNewRow)
                        {
                            continue
                        }

                        $cellValue = $row.Cells['Error'].Value

                        if ($cellValue -eq 'True')
                        {

                            $row.Visible = $true
                        }
                        else
                        {
                            $SyncHash.DataGridView.CurrentCell = $null
                            $row.Visible = $false
                        }
                    }                            
                })
                        
            $AllResults.Add_Click({
                    $SyncHash.DataGridView.CurrentCell = $null
                    foreach ($row in $SyncHash.DataGridView.Rows)
                    {
                        # Check if the row is a new row template, and skip if it is
                        if ($row.IsNewRow)
                        {
                            continue
                        }

                        $row.Visible = $true
                    }
                })
                        
            $ExportToCSV.Add_Click({

                    $SelectedFiles = @()
                    foreach ($row in $SyncHash.DataGridView.Rows)
                    {
                        if ($row.Cells["CheckBoxColumn"].Value)
                        {
                            $SelectedFiles += (($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value))
                        }
                    }
                    
                    if ($SelectedFiles.Count -eq 0)
                    {
                        [System.Windows.Forms.MessageBox]::Show("No items selected to export." , "Oops!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        return
                    }

                    $visibleData = @()

                    # Get column headers
                    $columnHeaders = @()
                    foreach ($column in $SyncHash.DataGridView.Columns)
                    {
                        $columnHeaders += $column.HeaderText
                    }

                    # Iterate through rows and add visible rows to the array
                    foreach ($row in $SyncHash.DataGridView.Rows)
                    {
                        $checkboxCell = $row.Cells['CheckBoxColumn']

                        # Get the value of the checkbox cell
                        # The value will be $true if checked, $false if unchecked
                        $isChecked = $checkboxCell.Value

                        if($SyncHash.ScanType -match "GetSingleVideoDetails")
                        {
                            $rowData = @{}
                            for ($i = 0; $i -lt $SyncHash.DataGridView.Columns.Count; $i++)
                            {
                                $columnName = $columnHeaders[$i]
                                $cellValue = $row.Cells[$i].Value
                                $rowData.$columnName = $cellValue
                            }
                            $visibleData += New-Object PSObject -Property $rowData
                        }
                        else
                        {
                            if (($row.Visible) -and ($isChecked -eq $true))
                            {
                                $rowData = @{}
                                for ($i = 0; $i -lt $SyncHash.DataGridView.Columns.Count; $i++)
                                {
                                    $columnName = $columnHeaders[$i]
                                    $cellValue = $row.Cells[$i].Value
                                    $rowData.$columnName = $cellValue
                                }
                                $visibleData += New-Object PSObject -Property $rowData
                            }
                        }
                    }

                    if($SyncHash.ScanType -match "GetSingleVideoDetails")
                    {
                        $orderedProperties = "File", "Folder", "Extension", "SizeMB", "Duration", "Section", "Name", "Value" # Replace with your actual headers
                    }
                    else
                    {
                        $orderedProperties = "Error", "File", "Folder", "Extension", "SizeMB", "Duration", "Results" # Replace with your actual headers
                    }
                            

                    # Export the visible data to a CSV file
                    $visibleData | Select-Object -Property $orderedProperties | Export-Csv -Path ($ExportPath.Text) -NoTypeInformation -Force

                    # Create the main form
                    $MessageForm = New-Object System.Windows.Forms.Form
                    $MessageForm.Text = "CSV Export Complete"
                    $MessageForm.Size = New-Object System.Drawing.Size(350, 150)
                    $MessageForm.StartPosition = "CenterScreen"
                    $MessageForm.FormBorderStyle = "FixedDialog" # Prevent resizing
                    $MessageForm.MaximizeBox = $false
                    $MessageForm.MinimizeBox = $false

                    # Add a label for the message
                    $label = New-Object System.Windows.Forms.Label
                    $label.Text = "Choose an option:"
                    $label.Location = New-Object System.Drawing.Point(20, 20)
                    $label.AutoSize = $true
                    $MessageForm.Controls.Add($label)

                    # Create the first custom button
                    $OpenFile = New-Object System.Windows.Forms.Button
                    $OpenFile.Text = "Open File"
                    $OpenFile.Location = New-Object System.Drawing.Point(20, 70)
                    $OpenFile.Size = New-Object System.Drawing.Size(90, 30)
                    $OpenFile.Add_Click({ $MessageForm.DialogResult = [System.Windows.Forms.DialogResult]::Yes }) # Assign a DialogResult
                    $MessageForm.Controls.Add($OpenFile)

                    # Create the second custom button
                    $OpenFolder = New-Object System.Windows.Forms.Button
                    $OpenFolder.Text = "Open Folder"
                    $OpenFolder.Location = New-Object System.Drawing.Point(120, 70)
                    $OpenFolder.Size = New-Object System.Drawing.Size(90, 30)
                    $OpenFolder.Add_Click({ $MessageForm.DialogResult = [System.Windows.Forms.DialogResult]::No }) # Assign a DialogResult
                    $MessageForm.Controls.Add($OpenFolder)

                    # Create the third custom button
                    $Cancel = New-Object System.Windows.Forms.Button
                    $Cancel.Text = "Cancel"
                    $Cancel.Location = New-Object System.Drawing.Point(220, 70)
                    $Cancel.Size = New-Object System.Drawing.Size(90, 30)
                    $Cancel.Add_Click({ $MessageForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel }) # Assign a DialogResult
                    $MessageForm.Controls.Add($Cancel)

                    # Show the form and capture the result
                    $result = $MessageForm.ShowDialog()

                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes)
                    {
                        Start-Process -FilePath ($ExportPath.Text)
                    }
                    elseif ($result -eq [System.Windows.Forms.DialogResult]::No)
                    {
                        $folderPath = Split-Path -Parent ($ExportPath.Text)
                        Start-Process -FilePath $folderPath
                    }
                    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
                    {
                        $MessageForm.Dispose()
                    }
                })

            $SelectExportPath.Add_Click({
                    # Create a new SaveFileDialog object
                    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

                    # Configure the dialog properties (optional)
                    # $SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments') # Set initial directory
                    $SaveFileDialog.InitialDirectory = $SyncHash.ScriptRoot
                    $SaveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*" # Set file type filter
                    $SaveFileDialog.Title = "Select a file to save" # Set dialog title

                    if($SyncHash.ScanType -match "RetrieveOnly")
                    {
                        $SaveFileDialog.FileName = "RetrieveOnly.csv"
                    }
                    elseif($SyncHash.ScanType -match "GetSingleVideoDetails")
                    {
                        $SaveFileDialog.FileName = "SingleVideoDetails.csv"
                    }
                    elseif($SyncHash.ScanType -match "CompareContainersAndExtensions")
                    {
                        $SaveFileDialog.FileName = "ContainersVsExtensions.csv"
                    }
                    elseif($SyncHash.ScanType -match "GetMoovAtom")
                    {
                        $SaveFileDialog.FileName = "MoovAtom.csv"
                    }
                    elseif($SyncHash.ScanType -match "GetSomeVideoCorruption")
                    {
                        $SaveFileDialog.FileName = "SomeVideoCorruption.csv"
                    }
                    else
                    {
                        $SaveFileDialog.FileName = "YourVideoInformation.csv"
                    }

                    # Show the dialog and get the result
                    $DialogResult = $SaveFileDialog.ShowDialog()

                    # Check if the user clicked OK
                    if ($DialogResult -eq [System.Windows.Forms.DialogResult]::OK)
                    {
                        # Get the selected file path
                        $CsvExportPath = ($SaveFileDialog.FileName)
                        $ExportPath.Text = $CsvExportPath
                        # You can now use $SelectedFilePath to save your data
                        # For example, to save a simple string to the selected file:
                        # "This is some content to save." | Out-File -FilePath $SelectedFilePath
                    }
                    # Dispose of the dialog object to free up resources
                    $SaveFileDialog.Dispose()
                })

            $SelectRepairPath.Add_Click({
                    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                    $FolderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop

                    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
                    {
                        $SyncHash.RepairOutputPath = $folderBrowser.SelectedPath
                        $FixAttemptPath.Text = $SyncHash.RepairOutputPath
                    }

                    # Check if the new repair path is a subfolder of the original scan path
                    $currentRepairPath = $FixAttemptPath.Text
                    $originalScanPath = $SyncHash.OriginalScanPath # This might be null
                    $PathStatus = $false
                    if ($originalScanPath -and $currentRepairPath) { $PathStatus = Test-PathContains -ParentPath $originalScanPath -ChildPath $currentRepairPath }

                    if($PathStatus)
                    {
                        [System.Windows.Forms.MessageBox]::Show(
                            "The output path is a sub-folder of the original`nscanned path. Some function will not work.`nOriginals and Fixes will collide." ,
                            "Heads Up!",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                })

            $AttemptRepair.Add_Click({
                    # This now calls the async wrapper function
                    & $SyncHash.StartAsyncRepair
                })

            $ButtonDelete.Add_Click({

                    if($SyncHash.ScanType -match "GetSingleVideoDetails")
                    {

                        $selectedRow = $SyncHash.DataGridView.SelectedRows[0] 

                        $cell1Value = $SyncHash.DataGridView.Rows[0].Cells[3].Value
                        $cell2Value = $SyncHash.DataGridView.Rows[0].Cells[2].Value

                        $ViewedFile = $cell1Value + "\" + $cell2Value

                        if($PermDelete.Checked)
                        {
                            # Define message box properties
                            $MessageBody = "The following video will be permanently deleted!!`n`n$ViewedFile`n`n                   Continue?"
                            $MessageTitle = "Confirmation"
                            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::YesNo
                            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Warning # Optional: Add an icon

                            # Display the message box and capture the result
                            $Result = [System.Windows.Forms.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)

                            # Check the user's response
                            if ($Result -eq [System.Windows.Forms.DialogResult]::Yes)
                            {
                                Remove-Item -Path $ViewedFile -Force
                                [System.Windows.Forms.MessageBox]::Show("The following video was permanently deleted:`n`n$ViewedFile", "Video Deleted")
                                $ResultsForm.Dispose()
                            }
                        }
                        else
                        {
                            Add-Type -AssemblyName Microsoft.VisualBasic
                            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($ViewedFile, 'OnlyErrorDialogs', 'SendToRecycleBin')
                            [System.Windows.Forms.MessageBox]::Show("The following video was moved to the Recycle Bin:`n`n$ViewedFile", "Video Deleted")
                            $ResultsForm.Dispose()
                        }
                    }
                    else
                    {
                        $SelectedFiles = @()
                        foreach ($row in $SyncHash.DataGridView.Rows)
                        {
                            if ($row.Cells["CheckBoxColumn"].Value)
                            {
                                $SelectedFiles += (($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value))
                            }
                        }
                    
                        if ($SelectedFiles.Count -eq 0)
                        {
                            [System.Windows.Forms.MessageBox]::Show("No items selected." , "Oops!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                            return
                        }

                        if($PermDelete.Checked)
                        {
                            # Define message box properties
                            $MessageBody = "Selected video(s) will be permanently deleted!!`n`n                   Continue?"
                            $MessageTitle = "Confirmation"
                            $ButtonType = [System.Windows.Forms.MessageBoxButtons]::YesNo
                            $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Warning # Optional: Add an icon

                            # Display the message box and capture the result
                            $Result = [System.Windows.Forms.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)

                            # Check the user's response
                            if ($Result -eq [System.Windows.Forms.DialogResult]::Yes)
                            {
                                foreach ($F in $SelectedFiles)
                                {
                                    Remove-Item -Path $F -Force
                                }
                            }
                        }
                        else
                        {
                            foreach ($F in $SelectedFiles)
                            {
                                Add-Type -AssemblyName Microsoft.VisualBasic
                                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($F, 'OnlyErrorDialogs', 'SendToRecycleBin')
                            }
                        }

                        if ($Result -ne [System.Windows.Forms.DialogResult]::No)
                        {
                            for ($i = $SyncHash.DataGridView.Rows.Count - 1; $i -ge 0; $i--)
                            {
                                $row = $SyncHash.DataGridView.Rows[$i]
                                $checkBoxCell = $row.Cells[0] # Assuming CheckBox column is at index 0

                                # Check if the CheckBox is checked
                                if ($checkBoxCell.Value -eq $true)
                                {
                                    $SyncHash.DataGridView.Rows.RemoveAt($i)
                                }
                            }
                        }
                    }
                })

            $SyncHash.DataGridView.Add_RowHeaderMouseClick({
                    # Access the row index from the event arguments
                    $rowIndex = $_.RowIndex

                    # Get the row object
                    $row = $SyncHash.DataGridView.Rows[$rowIndex]

                    # Now, you can access the cell value by its column header (name)
                    # Replace "YourColumnHeader" with the actual header of the column you want
                    $FileValue = ($row.Cells["File"].Value)
                    $FolderValue = ($row.Cells["Folder"].Value)
                    $PlayFile = ($FolderValue + "\" + $FileValue)
                    $ErrorColumn = $SyncHash.DataGridView.Columns["Error"]
                    if (($ErrorColumn.Visible))
                    {
                        $ErrorValue = $row.Cells['Error'].Value
                    }
                    $BooleanVariable = [System.Convert]::ToBoolean($ErrorValue)

                    if($SyncHash.ScanType -match "GetSingleVideoDetails")
                    {
                        $ffmpegOutput = $null
                        $ffmpegOutput = ffmpeg.exe -hwaccel auto -v error -i "$PlayFile" -f null - 2>&1
                        if ($ffmpegOutput -match ($SyncHash.ErrorsToMatch)){ $BooleanVariable = $true }
                    }
                            
                    if($BooleanVariable -eq $true)
                    {
                        [System.Windows.Forms.MessageBox]::Show("An error is indicated. `nIt could lock the applicaion." , "Heads Up" , [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                
                    }                    
                    Start-Process -FilePath "ffplay.exe" -ArgumentList "-i -loglevel quiet -nostats `"$PlayFile`"" -NoNewWindow -RedirectStandardError 'NUL'
                })

            $PlayVideos.add_Click({ & $SyncHash.ShowMultipleVideos })

            $SyncHash.DataGridView.Add_ColumnHeaderMouseClick({
                    param($sender, $e)
                    foreach ($row in $SyncHash.DataGridView.Rows)
                    {
                        if ($row.IsNewRow) { continue } # Skip the blank row at the bottom
                        $row.HeaderCell.Value = "Play"
                    }
                    for ($x = 0; $x -lt $SyncHash.DataGridView.RowCount; $x++)
                    {
                        $ErrorValue = $SyncHash.DataGridView.Rows[$x].Cells["Error"].Value
                        $ResultsValue = $SyncHash.DataGridView.Rows[$x].Cells["Results"].Value
                        $SingleFileName = $SyncHash.DataGridView.Rows[$x].Cells["Name"].Value
                        $SingleFileValue = $SyncHash.DataGridView.Rows[$x].Cells["Value"].Value

                        if(($SingleFileName -match "Some Detected Errors") -and (-not [string]::IsNullOrEmpty($SingleFileValue)))
                        { 
                            $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 0) # Red color
                        }

                        if($ErrorValue -eq "True")
                        { 
                            $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 255, 0, 0) # Red color
                        }

                        if($ResultsValue -match "is before")
                        { 
                            $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 0, 255, 0) # Green color
                            # $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 0, 0, 0) # White color
                        }
                        
                        if($ResultsValue -match "is after")
                        {
                            $SyncHash.DataGridView.Rows[$x].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255, 0) # Yellow color
                        }
                        
                    }
                })
        })

    $StatusTextBox.AppendText("Complete`r`n")

    if (-not ($SyncHash.ContainsKey("DataGridView")))
    {
        $SyncHash.Add("DataGridView", $DataGridView)
    }
    else
    {
        Set-Variable -Name "SyncHash.DataGridView" -Value $DataGridView -Force
    }
    # Get number of visible rows
    $visibleRows = $dataGridView.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible)

    # Update the Label text
    $RowCount.Text = "$visibleRows"
    # Set-Variable -Name "SyncHash.DataGridView" -Value $DataGridView -Force
    # $SyncHash.DataGridView = $DataGridView
    $ResultsForm.ShowDialog()
    $ResultsForm.BringToFront()
    $ResultsForm.Activate()

    # Get number of visible rows
    $visibleRows = $dataGridView.Rows.GetRowCount([System.Windows.Forms.DataGridViewElementStates]::Visible)

    # Update the Label text
    $RowCount.Text = "$visibleRows"
}

$SyncHash.GetTwoFolders = 
{
    param (
        $ScanPath = $SyncHash.ScanPath,
        $SyncHash = $SyncHash,
        $StatusTextBox = $SyncHash.StatusTextBox,
        $ScriptRoot = $SyncHash.ScriptRoot
    )

    $TwoFolderForm = New-Object System.Windows.Forms.Form
    $TwoFolderForm.Text = "List Two Folders Together"
    $TwoFolderForm.Size = New-Object System.Drawing.Size(600, 200)
    $TwoFolderForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $FolderOne = New-Object System.Windows.Forms.Button
    $FolderOne.Text = "Folder One"
    $FolderOne.Location = New-Object System.Drawing.Point(10, 10)
    $FolderOne.Size = New-Object System.Drawing.Size(120, 40)
    $TwoFolderForm.Controls.Add($FolderOne)

    $LabelOne = New-Object System.Windows.Forms.Label
    $LabelOne.Text = ""
    $LabelOne.Location = New-Object System.Drawing.Point(135, 20)
    $LabelOne.Size = New-Object System.Drawing.Size(430, 20)
    $LabelOne.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $LabelOne.Font = New-Object System.Drawing.Font("Arial", 12)
    $TwoFolderForm.Controls.Add($LabelOne)

    $FolderTwo = New-Object System.Windows.Forms.Button
    $FolderTwo.Text = "Folder Two"
    $FolderTwo.Location = New-Object System.Drawing.Point(10, 60)
    $FolderTwo.Size = New-Object System.Drawing.Size(120, 40)
    $TwoFolderForm.Controls.Add($FolderTwo)

    $LabelTwo = New-Object System.Windows.Forms.Label
    $LabelTwo.Text = ""
    $LabelTwo.Location = New-Object System.Drawing.Point(135, 70)
    $LabelTwo.Size = New-Object System.Drawing.Size(430, 20)
    $LabelTwo.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $LabelTwo.Font = New-Object System.Drawing.Font("Arial", 12)
    $TwoFolderForm.Controls.Add($LabelTwo)

    $GoButton = New-Object System.Windows.Forms.Button
    $GoButton.Text = "GO"
    $GoButton.Location = New-Object System.Drawing.Point(170, 110)
    $GoButton.Size = New-Object System.Drawing.Size(120, 40)
    $GoButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $TwoFolderForm.Controls.Add($GoButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Location = New-Object System.Drawing.Point(310, 110)
    $CancelButton.Size = New-Object System.Drawing.Size(120, 40)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $TwoFolderForm.Controls.Add($CancelButton)

    $FolderOneDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderTwoDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    # ... (Other form elements and their properties)

    # Define actions for the "Browse" buttons
    $FolderOne.Add_Click({
            if ($FolderOneDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $LabelOne.Text = $FolderOneDialog.SelectedPath
            }
        })

    $FolderTwo.Add_Click({
            if ($FolderTwoDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $LabelTwo.Text = $FolderTwoDialog.SelectedPath
            }
        })

    # Show the form and retrieve the results
    $result = $TwoFolderForm.ShowDialog()

    # Process selected paths if the user clicks OK
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        $TwoFolderForm.Dispose()

    }
    elseif($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $FolderOne = $LabelOne.Text
        $FolderTwo = $LabelTwo.Text

        $StatusTextBox.Clear()
        $FfmpegResults = @()
        $CsvExportPath = $SyncHash.ScriptRoot + "\TwoFolders.csv"

        $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
        "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
        "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
        "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
        "*.ts", "*.vob"#, "*.swf"

        $StatusTextBox.AppendText("Getting Files/Folders...`r`n")

        $AllFiles = Get-ChildItem -Path $FolderOne, $FolderTwo -Include $VideoExtensions -Recurse -File

        foreach($A in $AllFiles)
        {
            $StatusTextBox.AppendText("Processing: $($A.FullName)`r`n")
            $StatusTextBox.ScrollToCaret() 

            $Folder = (Split-Path $A.FullName)
            $File = (Split-Path $A.FullName -Leaf)
            
            if ($SyncHash.RescanForm.Visible )
            {
                $SyncHash.RescanForm.Invoke([action]{
                        $SyncHash.RescanTextBox.AppendText("Rescanning: $($File)`r`n")
                        [System.Windows.Forms.Application]::DoEvents()
                    })
            }
                
            $ErrorResult = $false
            $ProbeResult = $null
            $Result = $null
            $FfprobeOutput = $null

            $ffprobeOutput = (ffprobe.exe -v quiet -print_format json -show_entries format -i ($A.FullName) | ConvertFrom-Json)
            $FormatEntries = $FfprobeOutput.format
            $ProbeResult = "$($FormatEntries.format_name)"
                
            $VideoDuration = $FormatEntries.duration
            $TimeSpan = [TimeSpan]::FromSeconds($VideoDuration)
            $FormattedTime = $TimeSpan.ToString("hh\:mm\:ss")

            # $SizeKB = [math]::Round($A.Length / 1KB, 2) # Size in Kilobytes
            $SizeMB = [math]::Round($A.Length / 1MB, 2) # Size in Megabytes
            # $SizeGB = [math]::Round($A.Length / 1GB, 2) # Size in Gigabytes

            $Result = "Scan not performed."   

            $FfmpegResults += [PSCustomObject]@{
                Error     = $ErrorResult
                File      = $File
                Folder    = $Folder
                Extension = $A.Extension
                SizeMB    = $SizeMB
                Duration  = $FormattedTime
                Results   = $Result
            }
        }
 
        # New-ResultsForm -ScanType "ContainersVsExtensions" -SyncHash $SyncHash
        $SyncHash.ScanType = "GetTwoFolders"
        & $SyncHash.NewResultsForm $SyncHash.ScanPath $SyncHash $SyncHash.StatusTextBox $SyncHash.ScriptRoot $SyncHash.DataGridView $FfmpegResults
    }
}

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Find Video Corruption"
$MainForm.Size = New-Object System.Drawing.Size(1024, 768)
$MainForm.Font = New-Object System.Drawing.Font("Times New Roman", 12, [System.Drawing.FontStyle]::Regular)
$MainForm.StartPosition = "CenterScreen"
$MainForm.FormBorderStyle = "Fixed3D"
$MainForm.BackColor = "SeaGreen"

$RetrieveOnly = New-Object System.Windows.Forms.Button
$RetrieveOnly.Text = "Retrieve Only"
$RetrieveOnly.Location = New-Object System.Drawing.Point(824, 20)
$RetrieveOnly.Size = New-Object System.Drawing.Size(160, 50)

$CompareCnE = New-Object System.Windows.Forms.Button
$CompareCnE.Text = "Compare Containers / Extensions"
$CompareCnE.Location = New-Object System.Drawing.Point(824, 90)
$CompareCnE.Size = New-Object System.Drawing.Size(160, 50)

$MoovAtom = New-Object System.Windows.Forms.Button
$MoovAtom.Text = "Check Moov Atom"
$MoovAtom.Location = New-Object System.Drawing.Point(824, 160)
$MoovAtom.Size = New-Object System.Drawing.Size(160, 50)

$GeneralCorruption = New-Object System.Windows.Forms.Button
$GeneralCorruption.Text = "Check for General Corruption"
$GeneralCorruption.Location = New-Object System.Drawing.Point(824, 230)
$GeneralCorruption.Size = New-Object System.Drawing.Size(160, 50)

$SingleFile = New-Object System.Windows.Forms.Button
$SingleFile.Text = "Get Single Video Details"
$SingleFile.Location = New-Object System.Drawing.Point(824, 300)
$SingleFile.Size = New-Object System.Drawing.Size(160, 50)

$TwoFolders = New-Object System.Windows.Forms.Button
$TwoFolders.Text = "Get / Show Two Folders"
$TwoFolders.Location = New-Object System.Drawing.Point(824, 370)
$TwoFolders.Size = New-Object System.Drawing.Size(160, 50)

$FileCounterLabel = New-Object System.Windows.Forms.Label
$FileCounterLabel.Text = ""
$FileCounterLabel.Location = New-Object System.Drawing.Point(824, 450)
$FileCounterLabel.Size = New-Object System.Drawing.Size(160, 120)
$FileCounterLabel.Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Bold)
$FileCounterLabel.ForeColor = [System.Drawing.Color]::White
$FileCounterLabel.BackColor = [System.Drawing.Color]::Transparent
$FileCounterLabel.TextAlign = "MiddleCenter"
$SyncHash.FileCounterLabel = $FileCounterLabel

$ButtonQuit = New-Object System.Windows.Forms.Button
$ButtonQuit.Text = "Quit"
$ButtonQuit.Location = New-Object System.Drawing.Point(824, 658)
$ButtonQuit.Size = New-Object System.Drawing.Size(160, 50)

$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Text = "Cancel Scan"
$ButtonCancel.Location = New-Object System.Drawing.Point(824, 598)
$ButtonCancel.Size = New-Object System.Drawing.Size(160, 50)
$ButtonCancel.Enabled = $false # Disabled by default

$SelectScanPath = New-Object System.Windows.Forms.Button
$SelectScanPath.Text = "Path to scan"
$SelectScanPath.Location = New-Object System.Drawing.Point(20, 20)
$SelectScanPath.Size = New-Object System.Drawing.Size(120, 40)

$PathToScan = New-Object System.Windows.Forms.TextBox
$PathToScan.Location = New-Object System.Drawing.Point(160, 30)
$PathToScan.Size = New-Object System.Drawing.Size(400, 40)
$PathToScan.ReadOnly = $true # Make it read-only for display
$PathToScan.Text = $null
$SyncHash.PathToScan = $PathToScan

$RecursiveCheckBox = New-Object System.Windows.Forms.CheckBox
$RecursiveCheckBox.Text = "Include Subfolders"
$RecursiveCheckBox.Location = New-Object System.Drawing.Point(570, 35)
$RecursiveCheckBox.AutoSize = $true
$SyncHash.RecursiveCheckBox = $RecursiveCheckBox

$StatusTextBox = New-Object System.Windows.Forms.RichTextBox
$StatusTextBox.Multiline = $true
$StatusTextBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$StatusTextBox.WordWrap = $false
$StatusTextBox.Size = New-Object System.Drawing.Size(764, 628)
$StatusTextBox.Location = New-Object System.Drawing.Point(20, 80)
$StatusTextBox.ReadOnly = $true
$StatusTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$SyncHash.Add("StatusTextBox", $StatusTextBox)

$MainForm.Controls.Add($RetrieveOnly)
$MainForm.Controls.Add($CompareCnE)
$MainForm.Controls.Add($GeneralCorruption)
$MainForm.Controls.Add($MoovAtom)
$MainForm.Controls.Add($SingleFile)
$MainForm.Controls.Add($TwoFolders)
$MainForm.Controls.Add($FileCounterLabel)
$MainForm.Controls.Add($ButtonQuit)
$MainForm.Controls.Add($ButtonCancel)
$MainForm.Controls.Add($PathToScan)
$MainForm.Controls.Add($SelectScanPath)
$MainForm.Controls.Add($RecursiveCheckBox)
$MainForm.Controls.Add($StatusTextBox)

    $HelpLabel = New-Object System.Windows.Forms.Label
    $HelpLabel.Text = "F1 - Help"
    $HelpLabel.AutoSize = $true
    $HelpLabel.Location = New-Object System.Drawing.Point(750, 35)
    $HelpLabel.ForeColor = [System.Drawing.Color]::White
    $MainForm.Controls.Add($HelpLabel)

[xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help" Height="450" Width="600" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <RichTextBox x:Name="MyRichTextBox" Grid.Row="0" Margin="5" IsReadOnly="True" VerticalScrollBarVisibility="Auto">
            <FlowDocument>
                <FlowDocument.Resources>
                    <Style TargetType="{x:Type Paragraph}">
                        <Setter Property="Margin" Value="0,0,0,5"/>
                    </Style>
                </FlowDocument.Resources>
                <Paragraph>
                    <Bold><Run Text="Scan Types:"/></Bold>
                </Paragraph>
                <Paragraph>
                    <Bold><Run Text="Retrieve Only:"/></Bold> Just lists the video files found in the selected path without performing any checks.
                </Paragraph>
                <Paragraph>
                    <Bold><Run Text="Compare Containers / Extensions:"/></Bold> Checks if the file extension (e.g., .mp4) matches the actual container format inside the file. Mismatches can cause playback issues.
                </Paragraph>
                <Paragraph>
                    <Bold><Run Text="Check Moov Atom:"/></Bold> For MP4/MOV files, checks if the 'moov atom' (a critical part of the file's index) is at the beginning. If it's at the end, it can cause delays or issues with streaming.
                </Paragraph>
                <Paragraph>
                    <Bold><Run Text="Check for General Corruption:"/></Bold> A comprehensive scan using ffmpeg to decode the entire file and report various errors, such as missing headers, invalid data packets, and more.
                </Paragraph>
                <Paragraph>
                    <Bold><Run Text="Get Single Video Details:"/></Bold> Provides a detailed breakdown of a single video file's format, video stream, audio stream, and any detected corruption errors.
                </Paragraph>
                 <Paragraph>
                    <Bold><Run Text="Get / Show Two Folders:"/></Bold> Combines the file listings from two separate folders into a single view, which is useful for comparing contents.
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$MainForm.KeyPreview = $True
$MainForm.Add_KeyDown({
    param($Sender, $e)
    if ($e.KeyCode -eq 'F1') {
        $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
        $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
        $OkButton = $PopupWindow.FindName("OKButton")
        $OkButton.Add_Click({ $PopupWindow.Close() })
        $PopupWindow.ShowDialog() | Out-Null
    }
})

# --- Asynchronous Runspace and Timer Setup ---
$runspaceTimer = New-Object System.Windows.Forms.Timer
$runspaceTimer.Interval = 250 # Check every 250ms

# This hashtable will hold the active runspace details for the timer to monitor.
$activeRunspace = @{
    PowerShell  = $null
    AsyncResult = $null
}

$runspaceTimer.Add_Tick({
        # Check for progress updates
        if ($activeRunspace.PowerShell -and $activeRunspace.PowerShell.Streams.Progress.Count -gt 0)
        {
            $progressRecords = $activeRunspace.PowerShell.Streams.Progress.ReadAll()
            foreach ($record in $progressRecords)
            {
                # Update the file counter label if it exists and has data
                if ($SyncHash.FileCounterLabel -and $SyncHash.FileCounterLabel.IsHandleCreated -and $record.CurrentOperation)
                {
                    $SyncHash.FileCounterLabel.Invoke([action]{
                        $SyncHash.FileCounterLabel.Text = $record.CurrentOperation
                    })
                }
                # Determine which textbox to update based on whether it's a rescan or repair.
                $targetTextBox = if ($SyncHash.IsRescan) { $SyncHash.RescanTextBox } elseif ($SyncHash.IsRepair) { $SyncHash.RepairProgressRTB } else { $StatusTextBox }

                # Check if the target textbox and its parent form's handle have been created before invoking.
                if ($targetTextBox -and $targetTextBox.IsHandleCreated)
                {
                    # Use Invoke to safely update the UI from the timer's thread
                    $targetTextBox.Invoke([action]{
                            if ($record.Activity -eq "Repairing")
                            {
                                # For repair, clear the box and write the new file name.
                                # This happens only when a new Write-Progress is issued.
                                $targetTextBox.Text = $record.StatusDescription
                            }
                            else
                            {
                                # For scans, we append a new line
                                $targetTextBox.AppendText("$($record.Activity) - $($record.StatusDescription)`r`n")
                            }
                            $targetTextBox.ScrollToCaret()
                        })
                }
            }
        }

        # --- This is the new logic for streaming periods ---
        # If a repair is active and the background job is still running, append a period.
        if ($SyncHash.IsRepair -and $activeRunspace.AsyncResult -and -not $activeRunspace.AsyncResult.IsCompleted)
        {
            $targetTextBox = $SyncHash.RepairProgressRTB
            $targetTextBox.Add_TextChanged({
                    # Set the caret position to the end of the text
                    $this.SelectionStart = $this.Text.Length
                    # Scroll the RichTextBox to the caret position
                    $this.ScrollToCaret()
                })
            if ($targetTextBox -and $targetTextBox.IsHandleCreated)
            {
                $targetTextBox.Invoke([action]{
                        $targetTextBox.AppendText(".")
                    })
            }
        }
        # --- End of new logic ---

        # Check if the async operation is complete
        if ($activeRunspace.AsyncResult -and $activeRunspace.AsyncResult.IsCompleted)
        {
            # Stop the timer and immediately capture the runspace details into local variables.
            # Then, nullify the shared variables to prevent this block from running again on a subsequent tick (race condition).
            $runspaceTimer.Stop()
            $completedPowerShell = $activeRunspace.PowerShell
            $completedAsyncResult = $activeRunspace.AsyncResult
            $isRepair = $SyncHash.IsRepair
            $activeRunspace.PowerShell = $null
            $activeRunspace.AsyncResult = $null
            $SyncHash.IsRepair = $false # Reset flag

            # Determine which textbox to update for completion message.
            $targetTextBox = if ($SyncHash.IsRescan) { $SyncHash.RescanTextBox } elseif ($isRepair) { $SyncHash.RepairProgressRTB } else { $StatusTextBox }
            if ($targetTextBox -and $targetTextBox.IsHandleCreated)
            {
                $targetTextBox.Invoke([action]{ $targetTextBox.AppendText("`r`nReceiving results...`r`n") })
            }
        
            # End the invocation to get results and check for errors
            try
            {
                $results = $completedPowerShell.EndInvoke($completedAsyncResult)
                # Also check the error stream from the background task
                if ($completedPowerShell.Streams.Error.Count -gt 0)
                {
                    $backgroundErrors = $completedPowerShell.Streams.Error.ReadAll()
                    $backgroundErrors | ForEach-Object { $StatusTextBox.AppendText("`r`nBACKGROUND ERROR: $($_.ToString())`r`n") }
                }
            }
            catch [System.Management.Automation.PipelineStoppedException]
            {
                # This block executes specifically when the pipeline is stopped by .Stop()
                $StatusTextBox.AppendText("`r`nOperation cancelled by user.`r`n")
            }
            catch
            {
                $StatusTextBox.AppendText("`r`nERROR in background task: $($_.Exception.Message)`r`n")
            }

            # Re-enable UI controls
            $RetrieveOnly.Enabled = $true
            $CompareCnE.Enabled = $true
            $MoovAtom.Enabled = $true
            $GeneralCorruption.Enabled = $true
            $SingleFile.Enabled = $true
            $TwoFolders.Enabled = $true
            $ButtonCancel.Enabled = $false # Disable cancel button when scan is done

            # Clear the counter label when the scan is finished
            if ($SyncHash.FileCounterLabel -and $SyncHash.FileCounterLabel.IsHandleCreated)
            {
                $SyncHash.FileCounterLabel.Invoke([action]{ $SyncHash.FileCounterLabel.Text = "" })
            }

            if ($SyncHash.DataGridView)
            {
                $AttemptRepair = $SyncHash.DataGridView.FindForm().Controls.Find('AttemptRepair', $true)[0]
                if ($AttemptRepair) { $AttemptRepair.Enabled = $true }
            }
            if ($results)
            {
                if ($isRepair)
                {
                    # Handle repair results
                    $repairResult = $results[0]
                    $SyncHash.RepairProgressRTB.AppendText("Complete`r`n")
                    if ($repairResult.SkippedFiles.Count -ne 0)
                    {
                        # Show skipped files form (this needs to be invoked on UI thread)
                        $SyncHash.DataGridView.FindForm().Invoke([action]{
                                $SkippedFilesForm = New-Object System.Windows.Forms.Form
                                $SkippedFilesForm.Text = "Skipped Files"
                                $SkippedFilesForm.Size = New-Object System.Drawing.Size(600, 400)
                                $SkippedFilesForm.StartPosition = "CenterScreen"        

                                $SkippedFilesTB = New-Object System.Windows.Forms.RichTextBox
                                $SkippedFilesTB.Location = New-Object System.Drawing.Point(10, 10)
                                $SkippedFilesTB.Size = New-Object System.Drawing.Size(560, 300)
                                $SkippedFilesTB.Multiline = $true
                                $SkippedFilesTB.ReadOnly = $true
                                $SkippedFilesTB.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
                                $SkippedFilesForm.Controls.Add($SkippedFilesTB)

                                $OkButton = New-Object System.Windows.Forms.Button
                                $OkButton.Text = "OK"
                                $OkButton.Location = New-Object System.Drawing.Point(160, 310)
                                $OkButton.Size = New-Object System.Drawing.Size(60, 40)
                                $SkippedFilesForm.Controls.Add($OkButton)

                                $OkButton.Add_Click({ $SkippedFilesForm.Dispose() })

                                switch ($SyncHash.ScanType)
                                {
                                    "CompareContainersAndExtensions" { $SkippedFilesTB.Text = "The following files were skipped for the reason listed:`n`n$($repairResult.SkippedFiles -join "`n")" }
                                    "GetMoovAtom" { $SkippedFilesTB.Text = "The following files were skipped for the reason listed:`n`n$($repairResult.SkippedFiles -join "`n")" }
                                    { ($_ -match "GetSomeVideoCorruption" ) -or ($_ -match "GetSingleVideoDetails") } { $SkippedFilesTB.Text = "The following files were skipped because a repair for the detected error is not yet known:`n`n$($repairResult.SkippedFiles -join "`n")" }
                                }
                                $SkippedFilesForm.ShowDialog()
                            })
                    }
                    if($repairResult.FixedCount -ne 0)
                    {
                        $SyncHash.DataGridView.FindForm().Invoke([action]{
                                & $SyncHash.GetAfterRepairOption -ReScanPath $SyncHash.RepairOutputPath -SyncHash $SyncHash
                            })
                    }
                }
                else
                {
                    # Handle scan results
                    $scanResultData = $results[0]
                    $SyncHash.ScanType = $scanResultData.ScanType
                    # Clear results-form specific keys from SyncHash before calling the form
                    $SyncHash.Remove("RescanTextBox")
                    $SyncHash.Remove("DataGridView")
                    $SyncHash.Remove("CheckBoxColumn")
                    # Also remove the RescanForm to ensure it's fresh on the next run
                    if ($SyncHash.RescanForm -and $SyncHash.RescanForm.Visible)
                    {
                        $SyncHash.RescanForm.Invoke([action]{
                                $SyncHash.RescanForm.Close()
                            })
                    }
                    $SyncHash.Remove("RescanForm")
                    $SyncHash.IsRescan = $false # Reset the rescan flag
                    & $SyncHash.NewResultsForm -FfmpegResults $scanResultData.Results
                    $StatusTextBox.AppendText("`r`nScan results are ready.`r`n")
                }
            }
            else
            {
                $StatusTextBox.AppendText("`r`nOperation finished, but no results were returned.`r`n")
            }
        
            # --- Guaranteed Cleanup for this specific operation ---
            $completedPowerShell.Dispose()
            $completedPowerShell.Runspace.Dispose()
        }
    })


function Start-AsyncScan
{
    param(
        [string]$ScanType,
        [string]$ScanPath,
        [switch]$IsRescan
    )

    # --- Create a new, clean runspace for this specific scan ---
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $newRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($initialSessionState)
    $newRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $newRunspace.Open()

    # Create a new PowerShell instance for this runspace
    $newPowerShell = [System.Management.Automation.PowerShell]::Create()
    $newPowerShell.Runspace = $newRunspace

    # Store the new PowerShell instance and its result handle for the timer to monitor
    $activeRunspace.PowerShell = $newPowerShell
    # ----------------------------------------------------------------

    # Disable UI controls to prevent multiple scans
    $RetrieveOnly.Enabled = $false
    $CompareCnE.Enabled = $false
    $MoovAtom.Enabled = $false
    $GeneralCorruption.Enabled = $false
    $SingleFile.Enabled = $false
    $TwoFolders.Enabled = $false
    $ButtonCancel.Enabled = $true # Enable the cancel button

    $StatusTextBox.Clear()
    $StatusTextBox.Text = "Starting background scan for '$ScanType'..."

    # This is the most robust method: build a single, self-contained script block
    # that includes the full text of all required functions.
    $scriptBlockContent = @"
param(`$ScanType, `$ScanPath, `$IsRescan, `$ErrorsToMatch, `$Recursive)

function RetrieveOnly {
    $( (Get-Command RetrieveOnly -CommandType Function).Definition )
}
function CompareContainersAndExtensions {
    $( (Get-Command CompareContainersAndExtensions -CommandType Function).Definition )
}
function GetMoovAtom {
    $( (Get-Command GetMoovAtom -CommandType Function).Definition )
}
function GetSomeVideoCorruption {
    $( (Get-Command GetSomeVideoCorruption -CommandType Function).Definition )
}
function GetSingleVideoDetails {
    $( (Get-Command GetSingleVideoDetails -CommandType Function).Definition )
}
function Start-RepairAttempt { # This function is called from the UI thread, not the background thread
    $( (Get-Command Start-RepairAttempt -CommandType Function).Definition )
}
function Start-AsyncRepair {
    $( (Get-Command Start-AsyncRepair -CommandType Function).Definition )
}
function Test-PathContains {
    $( (Get-Command Test-PathContains -CommandType Function).Definition )
}

& `$ScanType -ScanPath `$ScanPath -IsRescan:`$IsRescan -ErrorsToMatch `$ErrorsToMatch -Recursive:`$Recursive
"@
    # Add the script to the pipeline
    $newPowerShell.AddScript($scriptBlockContent)

    # Add the parameters safely using a hashtable
    $newPowerShell.AddParameters(@{ ScanType = $ScanType; ScanPath = $ScanPath; IsRescan = $IsRescan.IsPresent; ErrorsToMatch = $SyncHash.ErrorsToMatch; Recursive = $SyncHash.RecursiveCheckBox.Checked }) | Out-Null

    # Begin the asynchronous invocation
    $activeRunspace.AsyncResult = $newPowerShell.BeginInvoke()

    # Start the timer to monitor progress and completion
    $runspaceTimer.Start()
}

function Start-AsyncRepair
{
    # This function is called from the UI thread to start the repair in the background.

    # --- 1. Gather data from the UI thread first ---
    $CheckedItems = @()
    if ($SyncHash.ScanType -match "GetSingleVideoDetails")
    {
        $errorRow = $SyncHash.DataGridView.Rows | Where-Object { $_.Cells["Name"].Value -eq "Error Detection -->" } | Select-Object -First 1
        if ($errorRow.Cells["Value"].Value -eq "No errors found.")
        {
            [System.Windows.Forms.MessageBox]::Show("No corruption errors found to repair.", "Heads Up", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $firstRowData = $SyncHash.DataGridView.Rows[0].DataBoundItem
        $CheckedItems = @([PSCustomObject]@{
                File      = $firstRowData.File
                Folder    = $firstRowData.Folder
                Extension = $firstRowData.Extension
                Results   = $errorRow.Cells["Value"].Value
                Error     = $true
            })
    }
    else
    {
        foreach ($row in $SyncHash.DataGridView.Rows)
        {
            if ((-not $row.IsNewRow) -and ($row.Visible) -and ($row.Cells['CheckBoxColumn'].Value -eq $true))
            {
                $CheckedItems += $row.DataBoundItem
            }
        }
    }

    if ($CheckedItems.Count -eq 0)
    {
        [System.Windows.Forms.MessageBox]::Show("No items selected.", "Oops!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # --- 1.5. Confirmation Dialog ---
    $continue = [System.Windows.Forms.DialogResult]::No
    switch ($SyncHash.ScanType)
    {
        "CompareContainersAndExtensions"
        {
            $continue = [System.Windows.Forms.MessageBox]::Show(
                "Be aware this will not address corruption.`n`nIt will only attempt to set correct file extension.`n`nFiles in repair destination of same name will be overwritten.`n`nContinue?" ,
                "Heads Up", 
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning) 
        }
        "GetMoovAtom"
        {
            $continue = [System.Windows.Forms.MessageBox]::Show(
                "Be aware this will not address corruption.`n`nIt will only attempt to move the Moov Atom.`n`nFiles in repair destination of same name will be overwritten.`n`nContinue?" , "Heads Up",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        { $SyncHash.ScanType -match "GetSomeVideoCorruption" -or $SyncHash.ScanType -match "GetSingleVideoDetails" }
        {
            $continue = [System.Windows.Forms.MessageBox]::Show(
                "Be aware that severly corrupted videos may be shorter, `nbut playable.`n`nFor Severe corruption, 2 repaired files `nwill be created as a choice which to keep.`n`nFiles in repair destination of same name will `nbe overwritten.`n`nContinue?" , "Heads Up",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    }
    if ($continue -ne [System.Windows.Forms.DialogResult]::Yes)
    {
        return
    }

    # --- 2. Create and configure the background runspace ---
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $newRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($initialSessionState)
    $newRunspace.ApartmentState = [System.Threading.ApartmentState]::STA
    $newRunspace.Open()

    $newPowerShell = [System.Management.Automation.PowerShell]::Create()
    $newPowerShell.Runspace = $newRunspace

    # Store the PowerShell instance for the timer to monitor
    $activeRunspace.PowerShell = $newPowerShell

    # --- 3. Disable UI and start the process ---
    $AttemptRepair = $SyncHash.DataGridView.FindForm().Controls.Find('AttemptRepair', $true)[0]
    if ($AttemptRepair) { $AttemptRepair.Enabled = $false }
    $ButtonCancel.Enabled = $true

    $SyncHash.RepairProgressRTB.Clear()
    $SyncHash.RepairProgressRTB.Text = "Starting background repair for '$($SyncHash.ScanType)'..."
    $SyncHash.IsRepair = $true # Set flag for the timer

    # Build the script block to execute in the background
    $scriptBlockContent = @"
param(`$SyncHash, `$CheckedItems)

# Define the function within the runspace's scope
function Start-RepairAttempt {
    $( (Get-Command Start-RepairAttempt -CommandType Function).Definition )
}

# Execute the function
Start-RepairAttempt -SyncHash `$SyncHash -CheckedItems `$CheckedItems
"@

    $newPowerShell.AddScript($scriptBlockContent)

    # Pass the synchronized hashtable and the collected items as parameters
    $newPowerShell.AddParameters(@{
            SyncHash     = $SyncHash
            CheckedItems = $CheckedItems
        }) | Out-Null

    # Begin the asynchronous invocation
    $activeRunspace.AsyncResult = $newPowerShell.BeginInvoke()

    # Start the timer to monitor for completion and progress
    $runspaceTimer.Start()

    # The timer's Add_Tick event will handle the rest:
    # - Reading progress streams
    # - Detecting completion
    # - Calling EndInvoke
    # - Re-enabling UI
    # - Disposing the runspace
}

function Show-MultipleVideos 
{
    $ffplayPath = "ffplay.exe"
    $SelectedFiles = @()
    foreach ($row in $SyncHash.DataGridView.Rows)
    {
        if (($row.Cells["CheckBoxColumn"].Value) -and $row.Visible)
        {
            $SelectedFiles += (($row.Cells["Folder"].Value) + "\" + ($row.Cells["File"].Value))
        }
    }

    if ($SelectedFiles.Count -gt 4)
    {
        [System.Windows.Forms.MessageBox]::Show("Too many videos selected.`n`rYou may select 1-4 videos.", "Oops!")
    }
    elseif($SelectedFiles.Count -le 0)
    {
        [System.Windows.Forms.MessageBox]::Show("No files selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
    else
    {
        # This is the start of the 'else' block.
        # Create and show a "Loading" popup form.
        $LoadingForm = New-Object System.Windows.Forms.Form -Property @{ Text = "Loading"; Size = New-Object System.Drawing.Size(200, 100); StartPosition = "CenterScreen"; FormBorderStyle = "FixedDialog"; ControlBox = $false; TopMost = $true }
        $LoadingLabel = New-Object System.Windows.Forms.Label -Property @{ Text = "Loading Videos..."; Location = New-Object System.Drawing.Point(40, 30); AutoSize = $true }
        $LoadingForm.Controls.Add($LoadingLabel)
        $LoadingForm.Show()
        $LoadingForm.Refresh()

        # PInvoke for SetParent and MoveWindow.
        Add-Type -MemberDefinition @"
            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
"@ -Namespace Win32 -Name API -PassThru | Out-Null

        $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
        $workingAreaWidth = $primaryScreen.WorkingArea.Width
        $workingAreaHeight = $primaryScreen.WorkingArea.Height

        $PlayerForm = New-Object System.Windows.Forms.Form -Property @{ Text = "Multi-Video Player"; Size = New-Object System.Drawing.Size(1280, 768); StartPosition = "CenterScreen"; TopMost = $true }

        $PlayerProcesses = New-Object System.Collections.ArrayList
        $Timers = New-Object System.Collections.ArrayList
        $sharedState = [PSCustomObject]@{ PlayersCompleted = 0 }

        function Start-Player
        {
            param(
                [string]$VideoFile,
                [int]$Width,
                [int]$Height
            )
            # Launch ffplay with the exact dimensions calculated, and still off-screen.
            $ffplayArgs = "-x $Width -y $Height -loglevel quiet -nostats -noborder -left $($workingAreaWidth) -top $($workingAreaHeight) -i `"$VideoFile`""
            $Process = New-Object System.Diagnostics.Process
            $Process.StartInfo.FileName = $ffplayPath
            $Process.StartInfo.Arguments = $ffplayArgs
            $Process.StartInfo.RedirectStandardOutput = $true
            $Process.StartInfo.RedirectStandardError = $true
            $Process.StartInfo.UseShellExecute = $false
            $Process.StartInfo.CreateNoWindow = $true
            $Process.Start()
            $PlayerProcesses.Add($Process) | Out-Null
            return $Process
        }

        $Panels = @()
        for ($i = 0; $i -lt [System.Math]::Min($SelectedFiles.Count, 4); $i++)
        {
            $panel = New-Object System.Windows.Forms.Panel -Property @{ Size = New-Object System.Drawing.Size(625, 325); BorderStyle = "FixedSingle"; Location = New-Object System.Drawing.Point((5 + (630 * ($i % 2))), (10 + (360 * [math]::Floor($i / 2)))) }
            $PlayerForm.Controls.Add($panel)
            $Panels += $panel

            $videoFile = $SelectedFiles[$i]

            # --- Pre-calculate the correct size for ffplay ---
            $ffprobeOutput = ffprobe.exe -v error -select_streams v:0 -show_entries stream=width, height -of csv=s=x:p=0 -i $videoFile 2> $null
            $videoWidth, $videoHeight = $ffprobeOutput -split 'x'
            
            $newWidth = $panel.Width
            $newHeight = $panel.Height

            if ($videoWidth -gt 0 -and $videoHeight -gt 0)
            {
                $aspectRatio = [double]$videoWidth / [double]$videoHeight
                if (($panel.Width / $panel.Height) -gt $aspectRatio)
                {
                    $newWidth = [int]($panel.Height * $aspectRatio)
                }
                else
                {
                    $newHeight = [int]($panel.Width / $aspectRatio)
                }
            }
            # --- End of pre-calculation ---
            
            $process = Start-Player -VideoFile $videoFile -Width $newWidth -Height $newHeight

            $timer = New-Object System.Windows.Forms.Timer
            $timer.Interval = 100
            $timer.Tag = @{ Process = $process; Panel = $panel; VideoFile = $videoFile; StartTime = (Get-Date); LoadingForm = $LoadingForm; PlayerForm = $PlayerForm; SharedState = $sharedState; PlayerCount = [System.Math]::Min($SelectedFiles.Count, 4) }
            [void]$Timers.Add($timer)

            $timer.Add_Tick({
                    param($sender, $e)
                    $timerData = $sender.Tag
                    # Re-read all variables from the Tag at the start of every tick.
                    $timerProcess = $timerData.Process
                    $timerPanel = $timerData.Panel
                    $timerVideoFile = $timerData.VideoFile
                    $timerLoadingForm = $timerData.LoadingForm
                    $timerPlayerForm = $timerData.PlayerForm
                    $timerSharedState = $timerData.SharedState
                    $timerPlayerCount = $timerData.PlayerCount

                    # --- Step 1: Check the status of the video process ---
                    $processHandle = $null
                    $timedOut = $false
                    $processCrashed = $false

                    # Check for a valid window handle
                    $processHandle = (Get-Process -Id $timerProcess.Id -ErrorAction SilentlyContinue).MainWindowHandle

                    # If no handle, check for other failure conditions
                    if (-not ($processHandle -and $processHandle -ne [System.IntPtr]::Zero))
                    {
                        if ($timerProcess.HasExited)
                        {
                            $processCrashed = $true
                        }
                        elseif (((Get-Date) - $timerData.StartTime).TotalSeconds -gt 10) # 10-second timeout
                        {
                            $timedOut = $true
                        }
                    }

                    # --- Step 2: Act based on the status. Only proceed if the timer should stop. ---
                    if (($processHandle -and $processHandle -ne [System.IntPtr]::Zero) -or $timedOut -or $processCrashed)
                    {
                        $sender.Stop() # Stop the timer, we have a final result.

                        if ($processHandle -and $processHandle -ne [System.IntPtr]::Zero)
                        {
                            # SUCCESS: We have a valid window handle.
                            $ffprobeOutput = ffprobe.exe -v error -select_streams v:0 -show_entries stream=width, height -of csv=s=x:p=0 -i $timerVideoFile 2> $null
                            $videoWidth, $videoHeight = $ffprobeOutput -split 'x'
                            if ($videoWidth -gt 0 -and $videoHeight -gt 0)
                            {
                                $aspectRatio = [double]$videoWidth / [double]$videoHeight
                                
                                # Determine the scaling factor by comparing the video and panel aspect ratios
                                if (($timerPanel.Width / $timerPanel.Height) -gt $aspectRatio)
                                {
                                    # Panel is wider than the video, so fit to height
                                    $newHeight = $timerPanel.Height                                    
                                    $newWidth = [int]($timerPanel.Height * $aspectRatio)
                                }
                                else
                                {
                                    # Panel is taller than or same ratio as the video, so fit to width
                                    $newWidth = $timerPanel.Width
                                    $newHeight = [int]($timerPanel.Width / $aspectRatio)
                                }
                                $xOffset = [int](($timerPanel.Width - $newWidth) / 2)
                                $yOffset = [int](($timerPanel.Height - $newHeight) / 2)

                                [Win32.API]::SetParent($processHandle, $timerPanel.Handle) | Out-Null
                                [Win32.API]::MoveWindow($processHandle, $xOffset, $yOffset, $newWidth, $newHeight, $true) | Out-Null
                            }
                            else
                            {
                                # Fallback if ffprobe fails
                                [Win32.API]::SetParent($processHandle, $timerPanel.Handle) | Out-Null
                                [Win32.API]::MoveWindow($processHandle, 0, 0, $timerPanel.Width, $timerPanel.Height, $true) | Out-Null
                            }
                        }
                        else
                        {
                            # FAILURE: The process timed out or crashed.
                            $errorLabel = New-Object System.Windows.Forms.Label -Property @{
                                Text = "Video Unplayable`n`n$(Split-Path $timerVideoFile -Leaf)"; Dock = 'Fill'; TextAlign = 'MiddleCenter'
                                ForeColor = [System.Drawing.Color]::White; BackColor = [System.Drawing.Color]::Black
                                Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
                            }
                            $timerPanel.Controls.Add($errorLabel)
                            # This check is now safe because $timerProcess is guaranteed to be a process object.
                            if ($timerProcess -and -not $timerProcess.HasExited) { $timerProcess.Kill() }
                        }

                        # --- Step 3: Final completion logic (runs only once after timer stops) ---
                        $timerSharedState.PlayersCompleted++
                        if ($timerSharedState.PlayersCompleted -ge $timerPlayerCount)
                        {
                            $timerLoadingForm.Close(); $timerLoadingForm.Dispose()
                            $timerPlayerForm.ShowDialog() | Out-Null
                        }
                    }
                    # If we reach here, the timer is still running (video is loading), so we do nothing and wait for the next tick.
                })
            $timer.Start()
        }

        $PlayerForm.Add_FormClosed({
                # Stop all timers
                foreach ($T in $Timers) { $T.Stop(); $T.Dispose() }
                # Kill all ffplay processes
                foreach ($P in $PlayerProcesses)
                {
                    if (-not $P.HasExited)
                    {
                        $P.Kill()
                    }
                }
            })

    }
}


# Store the Show-MultipleVideos function in SyncHash for compatibility with the results form
$SyncHash.ShowMultipleVideos = ${function:Show-MultipleVideos}

# Store the Start-AsyncRepair function in SyncHash
$SyncHash.StartAsyncRepair = ${function:Start-AsyncRepair}

# Store the Start-RepairAttempt function in SyncHash
$SyncHash.StartRepairAttempt = ${function:Start-RepairAttempt}

$SyncHash.GetAfterRepairOption = {
    param (
        $ReScanPath,
        $SyncHash
    )

    $AfterScanForm = New-Object System.Windows.Forms.Form
    $AfterScanForm.Text = 'Repair Attempt Complete'
    $AfterScanForm.Size = New-Object System.Drawing.Size(440, 140)
    $AfterScanForm.StartPosition = 'CenterScreen'

    $OpenFolder = New-Object System.Windows.Forms.Button
    $OpenFolder.Location = New-Object System.Drawing.Point(10, 50)
    $OpenFolder.Size = New-Object System.Drawing.Size(120, 40)
    $OpenFolder.Text = 'Open Folder'
    $OpenFolder.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $AfterScanForm.AcceptButton = $OpenFolder
    $AfterScanForm.Controls.Add($OpenFolder)

    $Rescan = New-Object System.Windows.Forms.Button
    $Rescan.Location = New-Object System.Drawing.Point(150, 50)
    $Rescan.Size = New-Object System.Drawing.Size(120, 40)
    $Rescan.Text = 'Scan Output Folder'
    $Rescan.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $AfterScanForm.Controls.Add($Rescan)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(290, 50)
    $CancelButton.Size = New-Object System.Drawing.Size(120, 40)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $AfterScanForm.CancelButton = $CancelButton
    $AfterScanForm.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(150, 20)
    $label.Size = New-Object System.Drawing.Size(150, 50)
    $label.Text = "Repair Attempt Complete."

    $AfterScanForm.Controls.Add($label)
    $AfterScanForm.Topmost = $true
    $AfterScanForm.TopLevel = $true

    $result = $AfterScanForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        Start-Process -FilePath $SyncHash.RepairOutputPath
    }
    elseif($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        return
    }
    elseif($result -eq [System.Windows.Forms.DialogResult]::Retry)
    {
        $SyncHash.IsRescan = $true
        $syncHash.RescanForm.Show()

        # Determine the correct scan type for the rescan.
        # If the original scan was for a single file, we should now use a general corruption scan for the folder.
        $rescanType = if ($SyncHash.ScanType -eq "GetSingleVideoDetails") { "GetSomeVideoCorruption" } else { $SyncHash.ScanType }

        # Since this runs on the main thread, we can call Start-AsyncScan directly
        # We also need to ensure the Recursive checkbox is not checked for this specific rescan, as we only want to scan the output folder.
        $SyncHash.RecursiveCheckBox.Checked = $false
        Start-AsyncScan -ScanType $rescanType -ScanPath $SyncHash.RepairOutputPath -IsRescan
    }
}

# Set a default background color for all buttons on the form
$MainForm.Controls | Where-Object { $_.GetType().Name -eq "Button" } | ForEach-Object {
    $_.BackColor = [System.Drawing.Color]::LawnGreen
    $_.ForeColor = [System.Drawing.Color]::Black
    $_.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
}

$RetrieveOnly.Add_Click({
        if($null -ne ($SyncHash.ScanPath))
        { Start-AsyncScan -ScanType "RetrieveOnly" -ScanPath $SyncHash.ScanPath }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Please select a path to scan first.", "Oops", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
$CompareCnE.Add_Click({
        if($null -ne ($SyncHash.ScanPath))
        { Start-AsyncScan -ScanType "CompareContainersAndExtensions" -ScanPath $SyncHash.ScanPath }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Please select a path to scan first.", "Oops", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
$MoovAtom.Add_Click({
        if($null -ne ($SyncHash.ScanPath))
        { Start-AsyncScan -ScanType "GetMoovAtom" -ScanPath $SyncHash.ScanPath }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Please select a path to scan first.", "Oops", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
$GeneralCorruption.Add_Click({
        if($null -ne ($SyncHash.ScanPath))
        { Start-AsyncScan -ScanType "GetSomeVideoCorruption" -ScanPath $SyncHash.ScanPath }
        else
        {
            [System.Windows.Forms.MessageBox]::Show("Please select a path to scan first.", "Oops", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
$SingleFile.Add_Click({
        # Single file selection is quick, so we can run it directly without a job.
        $resultData = GetSingleVideoDetails -ErrorsToMatch $SyncHash.ErrorsToMatch
        if ($resultData)
        {
            $SyncHash.ScanType = $resultData.ScanType
            & $SyncHash.NewResultsForm -FfmpegResults $resultData.Results
        }
    })
$TwoFolders.Add_Click({
        # This function needs to be refactored to work with background jobs if it's a long process.
        # For now, calling it directly.
        & $SyncHash.GetTwoFolders $SyncHash.ScanPath $SyncHash $SyncHash.StatusTextBox $SyncHash.ScriptRoot
    })
    
$SelectScanPath.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.ShowNewFolderButton = $false
        $FolderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $PathToScan.Text = $folderBrowser.SelectedPath
            $SyncHash.PathToScan.Text = $folderBrowser.SelectedPath
            $SyncHash.OriginalScanPath = $folderBrowser.SelectedPath
        }
        $SyncHash.ScanPath = $folderBrowser.SelectedPath
    })

$ButtonQuit.Add_Click({
        # Simply close the main form. The Add_Closed event will handle all cleanup.
        $MainForm.Close() | Out-Null
    })

$ButtonCancel.Add_Click({
        if ($activeRunspace.PowerShell -and $activeRunspace.PowerShell.InvocationStateInfo.State -eq 'Running')
        {
            $StatusTextBox.AppendText("`r`nAttempting to cancel scan...`r`n")
            $activeRunspace.PowerShell.Stop()
        }
    })

$MainForm.Add_Closed({
        # Centralized cleanup logic. This runs when the form is closed for any reason.
        # Stop the timer first to prevent it from firing during cleanup.
        if ($runspaceTimer -and $runspaceTimer.Enabled) { $runspaceTimer.Stop() }
        if ($runspaceTimer) { $runspaceTimer.Dispose() }

        # If a PowerShell pipeline is running, stop it gracefully before disposing.
        if ($activeRunspace.PowerShell -and $activeRunspace.PowerShell.InvocationStateInfo.State -eq 'Running')
        {
            $activeRunspace.PowerShell.Stop()
        }
        if ($activeRunspace.PowerShell) { $activeRunspace.PowerShell.Dispose() }
        if ($activeRunspace.PowerShell -and $activeRunspace.PowerShell.Runspace) { $activeRunspace.PowerShell.Runspace.Dispose() }

        # Kill any lingering ffplay processes to ensure a clean exit.
        Get-Process -Name "ffplay" -ErrorAction SilentlyContinue | Stop-Process -Force
    })

$MainForm.ShowDialog() | Out-Null
$MainForm.BringToFront()
$MainForm.Activate()
