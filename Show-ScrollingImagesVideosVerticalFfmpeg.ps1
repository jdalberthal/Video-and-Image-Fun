<#
.SYNOPSIS
    Creates a continuous vertical-scrolling display of selected images and videos using FFmpeg.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them in a continuous,
    auto-scrolling vertical display within a WPF window.

    It uses FFmpeg to decode video frames in real-time, which provides broad format support beyond
    native Windows capabilities. The display is interactive, with controls to pause, reverse
    direction, and adjust the scrolling speed.

.EXAMPLE
    PS C:\> .\Show-ScrollingImagesVideosVerticalFfmpeg.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the script will
    launch the vertical scrolling window.

.NOTES
    Name:           Show-ScrollingImagesVideosVerticalFfmpeg.ps1
    Version:        1.0.0, 10/18/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET/WPF access. The following executables must be in
                    the system's PATH or in the same directory as the script:
                    - FFmpeg (ffmpeg.exe, ffplay.exe): https://www.ffmpeg.org/download.html
#>
Clear-Host
# Step 1: Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase | Out-Null
Add-Type -AssemblyName System.Windows.Forms, System.Drawing | Out-Null
Add-Type -AssemblyName WindowsFormsIntegration, System.Xaml | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$ExternalButtonName = "Scrolling Images/Videos `n Vertcal - Uses Ffmpeg"
$ScriptDescription = "Creates a continuous vertical-scrolling display of selected images and videos. Uses FFmpeg for video decoding."
$RequiredExecutables = @("ffmpeg.exe", "ffplay.exe")

# --- Dependency Check ---
if ($RequiredExecutables) {
    $dependencyStatus = @()
    $allDependenciesMet = $true

    # First, check all dependencies without writing to the console yet
    foreach ($exe in $RequiredExecutables) {
        # Check in the PATH and in the script's local directory ($PSScriptRoot)
        $localPath = Join-Path $PSScriptRoot $exe
        if ((Get-Command $exe -ErrorAction SilentlyContinue) -or (Test-Path -Path $localPath)) {
            $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'Found' }
        } else {
            $dependencyStatus += [PSCustomObject]@{ Name = $exe; Status = 'NOT FOUND' }
            $allDependenciesMet = $false
        }
    }

    # If any dependency is missing, then write the status of all of them
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
}
    
$SyncHash = [hashtable]::Synchronized(@{}) # For passing data between runspaces

$SyncHash.Paused = $false
$SyncHash.ControlsHidden = $false
$SyncHash.RbSelection = ""

$MyFont = New-Object System.Drawing.Font("Arial", 12)

# Create the form
$SelectFolderForm = New-Object System.Windows.Forms.Form
$SelectFolderForm.Text = "Video/Image Selector"
$SelectFolderForm.Size = New-Object System.Drawing.Size(800, 800)
$SelectFolderForm.StartPosition = "CenterScreen"

# Create a Button for browsing folders
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse Folder"
$browseButton.Location = New-Object System.Drawing.Point(10, 10)
$browseButton.Size = New-Object System.Drawing.Size(100, 25)
$SelectFolderForm.Controls.Add($browseButton)

# Create a TextBox to display the selected folder path
$folderPathTextBox = New-Object System.Windows.Forms.TextBox
$folderPathTextBox.Location = New-Object System.Drawing.Point(120, 10)
$folderPathTextBox.Size = New-Object System.Drawing.Size(450, 25)
$folderPathTextBox.ReadOnly = $true
$SelectFolderForm.Controls.Add($folderPathTextBox)

# Create a CheckBox for recursive scanning
$recursiveCheckBox = New-Object System.Windows.Forms.CheckBox
$recursiveCheckBox.Text = "Include Subfolders"
$recursiveCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$recursiveCheckBox.AutoSize = $true
$recursiveCheckBox.Checked = $false
$SelectFolderForm.Controls.Add($recursiveCheckBox)

# Create a CheckBox for transparency
$TransparentCheckbox = New-Object System.Windows.Forms.CheckBox
$TransparentCheckbox.Text = "Make Semi-Transparent"
$TransparentCheckbox.Location = New-Object System.Drawing.Point(440, 47)
$TransparentCheckbox.AutoSize = $true
$TransparentCheckbox.Checked = $false
$SelectFolderForm.Controls.Add($TransparentCheckbox)

# Create a DataGridView to display files
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 80)
$dataGridView.Size = New-Object System.Drawing.Size(760, 350)
$dataGridView.Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right)
$dataGridView.AutoGenerateColumns = $false # We'll define columns manually
$dataGridView.AllowUserToAddRows = $false
$SelectFolderForm.Controls.Add($dataGridView)

# Add a CheckBox column to the DataGridView
$checkBoxColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$checkBoxColumn.Name = "Select"
$checkBoxColumn.HeaderText = ""
$checkBoxColumn.Width = 30
$dataGridView.Columns.Add($checkBoxColumn) | Out-Null

# Add a column for file names
$fileNameColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$fileNameColumn.Name = "FileName"
$fileNameColumn.HeaderText = "File Name"
$fileNameColumn.Width = 300
$fileNameColumn.ReadOnly = $true
$dataGridView.Columns.Add($fileNameColumn) | Out-Null

# Add a column for full file paths
$filePathColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$filePathColumn.Name = "FilePath"
$filePathColumn.HeaderText = "Full Path"
$filePathColumn.Width = 430
$filePathColumn.ReadOnly = $true
$dataGridView.Columns.Add($filePathColumn) | Out-Null

$MaximizeCheckBox = New-Object System.Windows.Forms.CheckBox
$MaximizeCheckBox.Text = "Maximize"
$MaximizeCheckBox.Location = New-Object System.Drawing.Point(350, 47)
$MaximizeCheckBox.AutoSize = $true
$MaximizeCheckBox.Checked = $false
$SelectFolderForm.Controls.Add($MaximizeCheckBox)

# Create a Button to perform an action on selected files
$PlayButton = New-Object System.Windows.Forms.Button
$PlayButton.Text = "Play Selected Files"
$PlayButton.Location = New-Object System.Drawing.Point(600, 40)
$PlayButton.Size = New-Object System.Drawing.Size(170, 30)
$SelectFolderForm.Controls.Add($PlayButton)

# --- Text Overlay Controls ---
# GroupBox for Radio Buttons
$GroupBox = New-Object System.Windows.Forms.GroupBox
$GroupBox.Text = "Text Overlay"
$GroupBox.Location = New-Object System.Drawing.Point(10, 440)
$GroupBox.Size = New-Object System.Drawing.Size(125, 130)

$RadioButton1 = New-Object System.Windows.Forms.RadioButton
$RadioButton1.Text = "Hide Text Overlay"
$RadioButton1.Location = New-Object System.Drawing.Point(10, 30)
$RadioButton1.Width = 114
$RadioButton1.Checked = $True # Default selected

$RadioButton2 = New-Object System.Windows.Forms.RadioButton
$RadioButton2.Text = "Filename"
$RadioButton2.Location = New-Object System.Drawing.Point(10, 60)

$RadioButton3 = New-Object System.Windows.Forms.RadioButton
$RadioButton3.Text = "Custom Text"
$RadioButton3.Location = New-Object System.Drawing.Point(10, 90)

$GroupBox.Controls.Add($RadioButton1)
$GroupBox.Controls.Add($RadioButton2)
$GroupBox.Controls.Add($RadioButton3)
$SelectFolderForm.Controls.Add($GroupBox)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Point(140, 440)
$TextBox.Size = New-Object System.Drawing.Size(455, 310)
$TextBox.Multiline = $True
$TextBox.Visible = $False
$TextBox.ScrollBars = "Vertical"
$TextBox.Font = $MyFont
$SelectFolderForm.Controls.Add($TextBox)
$SyncHash.TextBox = $TextBox

$SyncHash.TextColor = [PSCustomObject]@{ A = 255; R = 0; G = 0; B = 0 }

$CurrentColor = New-Object System.Windows.Forms.Label
$CurrentColor.Text = "Text Color:"
$CurrentColor.Location = New-Object System.Drawing.Point(600, 477)
$CurrentColor.AutoSize = $True
$CurrentColor.BackColor = [System.Drawing.Color]::Transparent
$CurrentColor.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
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

$selectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckbox.Text = "Select All"
$selectAllCheckbox.Location = New-Object System.Drawing.Point(10, 60)
$selectAllCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$selectAllCheckbox.Checked = $false
$SelectFolderForm.Controls.Add($selectAllCheckbox) # Add to your form or panel

$SizeLabel = New-Object System.Windows.Forms.Label
$SizeLabel.Text = "Font Size:"
$SizeLabel.AutoSize = $True
$SizeLabel.Location = New-Object System.Drawing.Point(600, 522)
$SizeLabel.Size = New-Object System.Drawing.Size(25, 20)
$SizeLabel.Visible = $False
$SelectFolderForm.Controls.Add($SizeLabel)

$NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$NumericUpDown.Location = New-Object System.Drawing.Point(660, 520)
$NumericUpDown.Size = New-Object System.Drawing.Size(40, 20)
$NumericUpDown.Visible = $False
$NumericUpDown.Minimum = 0
$NumericUpDown.Maximum = 600
$NumericUpDown.Increment = 1
$NumericUpDown.DecimalPlaces = 0
$NumericUpDown.Value = 12
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
$BoldCheckbox.Checked = $False
$BoldCheckbox.Visible = $False
$SelectFolderForm.Controls.Add($BoldCheckbox)
$SyncHash.BoldCheckbox = $BoldCheckbox

$Event = {
    if ($RadioButton1.Checked) { $SyncHash.RbSelection = "Hidden" }
    elseif ($RadioButton2.Checked) { $SyncHash.RbSelection = "Filename" }
    elseif ($RadioButton3.Checked) { $SyncHash.RbSelection = "Custom" }

    $TextBox.Visible = $RadioButton3.Checked
    $controlsVisibility = ($RadioButton2.Checked -or $RadioButton3.Checked)
    $CurrentColor.Visible = $controlsVisibility
    $ColorExample.Visible = $controlsVisibility
    $SelectColorButton.Visible = $controlsVisibility
    $SizeLabel.Visible = $controlsVisibility
    $NumericUpDown.Visible = $controlsVisibility
    $FontButton.Visible = $controlsVisibility
    $ItalicCheckbox.Visible = $controlsVisibility
    $BoldCheckbox.Visible = $controlsVisibility
}
$RadioButton1.Add_Click($Event)
$RadioButton2.Add_Click($Event)
$RadioButton3.Add_Click($Event)

$selectAllCheckbox.Add_CheckedChanged({
        $checkedState = $selectAllCheckbox.Checked
        foreach ($row in $dataGridView.Rows)
        {
            $row.Cells["Select"].Value = $checkedState
        }
        $dataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    })

# Copy the XAML for the font picker popup
[xml]$XamlFontPicker = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Font Selector" Height="500" Width="400" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="14">
    <Window.Resources>
        <DataTemplate x:Key="FontItemTemplate">
            <TextBlock Text="{Binding}" FontFamily="{Binding}" Margin="2"/>
        </DataTemplate>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <ListBox x:Name="fontListBox" Grid.Row="0" Margin="5" 
                 ItemTemplate="{StaticResource FontItemTemplate}" />
        <Button x:Name="SelectButton" Grid.Row="1" Content="Select" HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$FontButton.Add_Click({
        $ReaderFontPicker = (New-Object System.Xml.XmlNodeReader $XamlFontPicker)
        $FontPopupWindow = [Windows.Markup.XamlReader]::Load($ReaderFontPicker)
        $fontListBox = $FontPopupWindow.FindName("fontListBox")
        $SelectButton = $FontPopupWindow.FindName("SelectButton")

        $installedFonts = New-Object System.Drawing.Text.InstalledFontCollection
        foreach ($fontFamily in $installedFonts.Families)
        {
            [void]$fontListBox.Items.Add($fontFamily.Name)
        }

        if ($fontListBox.Items.Contains($SyncHash.SelectedFont)) {
            $fontListBox.SelectedItem = $SyncHash.SelectedFont
        } else {
            $fontListBox.SelectedIndex = 0
        }

        $SelectButton.Add_Click({
            if ($fontListBox.SelectedItem) {
                $SyncHash.SelectedFont = $fontListBox.SelectedItem
                $newFont = New-Object System.Drawing.Font($SyncHash.SelectedFont, $SyncHash.TextBox.Font.Size, $SyncHash.TextBox.Font.Style)
                $SyncHash.TextBox.Font = $newFont
            }
            $FontPopupWindow.Close()
        })
        $FontPopupWindow.ShowDialog() | Out-Null
    })

$colorDialog = New-Object System.Windows.Forms.ColorDialog
$SelectColorButton.Add_Click({
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $ColorExample.BackColor = $colorDialog.Color
            $TextBox.ForeColor = $colorDialog.Color
            $SyncHash.TextColor = $colorDialog.Color
        }
    }) | Out-Null

$NumericUpDown.Add_ValueChanged({
        param($sender, $e)
        $SyncHash.SelectedFontSize = $sender.Value
        $newFont = New-Object System.Drawing.Font($SyncHash.TextBox.Font.FontFamily, $sender.Value, $SyncHash.TextBox.Font.Style)
        $SyncHash.TextBox.Font = $newFont
    })

$fontStyleChangeHandler = {
    $fontStyle = [System.Drawing.FontStyle]::Regular
    if ($SyncHash.ItalicCheckbox.Checked) { $fontStyle = $fontStyle -bor [System.Drawing.FontStyle]::Italic }
    if ($SyncHash.BoldCheckbox.Checked) { $fontStyle = $fontStyle -bor [System.Drawing.FontStyle]::Bold }
    $newFont = New-Object System.Drawing.Font($SyncHash.TextBox.Font.FontFamily, $SyncHash.TextBox.Font.Size, $fontStyle)
    $SyncHash.TextBox.Font = $newFont
}
$ItalicCheckbox.Add_CheckedChanged($fontStyleChangeHandler)
$BoldCheckbox.Add_CheckedChanged($fontStyleChangeHandler)

$SelectFolderForm.KeyPreview = $True
$SelectFolderForm.Add_KeyDown({
        param($Sender, $e)
        if ($_.KeyCode -eq "F1")
        {
            $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
            $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
            $OkButton = $PopupWindow.FindName("OKButton")
            $OkButton.Add_Click({ $PopupWindow.Close() })
            $PopupWindow.ShowDialog() | Out-Null
        }
    })

$dataGridView.Add_CellContentClick({
        if ($_.Column.Name -eq "Select" -and $_.RowIndex -ne -1)
        {
            if (-not $_.Value)
            {
                $selectAllCheckbox.Checked = $false
            }
        }
    })

$dataGridView.Add_RowHeaderMouseClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0) {
        $row = $dataGridView.Rows[$e.RowIndex]
        $filePath = $row.Cells["FilePath"].Value
        if ([System.IO.File]::Exists($filePath)) {
            Start-Process -FilePath "ffplay.exe" -ArgumentList "-loglevel quiet -nostats -i `"$filePath`"" -NoNewWindow
        } else {
            [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Event handler for the Browse Folder button
$browseButton.Add_Click({

        $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
        "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
        "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
        "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
        "*.ts", "*.vob"#, "*.swf"

        $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"
        $AllowedExtension = $VideoExtensions + $ImageExtensions

        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select the folder to scan."

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $selectedPath = "$($folderBrowser.SelectedPath)"
            $folderPathTextBox.Text = $selectedPath

            $dataGridView.Rows.Clear()

            $files = if ($recursiveCheckBox.Checked) {
                Get-ChildItem -Path $selectedPath -File -Include $AllowedExtension -Recurse
            } else {
                Get-ChildItem -Path "$($selectedPath)\*" -File -Include $AllowedExtension
            }

            foreach ($file in $files)
            {
                $dataGridView.Rows.Add($false, $file.Name, $file.FullName)
            }
        }
    })

# Event handler for the Process Selected Files button
$PlayButton.Add_Click({
        $selectedFiles = @()
        foreach ($row in $dataGridView.Rows)
        {
            if ($row.Cells["Select"].Value)
            {
                $selectedFiles += $row.Cells["FilePath"].Value
            }
        }

        if($selectedFiles.Count -le 0)
        {
            [System.Windows.Forms.MessageBox]::Show("No files selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        else
        {
            # --- Loading Form in a Separate Runspace ---
            $SyncHash.CubeReady = $false

            $loadingScriptBlock = {
                param($SyncHash)
                Add-Type -AssemblyName System.Windows.Forms, System.Drawing
                [System.Windows.Forms.Application]::EnableVisualStyles()

                $loadingForm = New-Object System.Windows.Forms.Form
                $loadingForm.Text = "Loading..."
                $loadingForm.Size = New-Object System.Drawing.Size(300, 120)
                $loadingForm.StartPosition = "CenterScreen"
                $loadingForm.FormBorderStyle = "FixedDialog"
                $loadingForm.ControlBox = $false

                $loadingLabel = New-Object System.Windows.Forms.Label
                $loadingLabel.Text = "Loading media, please wait..."
                $loadingLabel.Location = New-Object System.Drawing.Point(20, 20)
                $loadingLabel.AutoSize = $true
                $loadingForm.Controls.Add($loadingLabel)

                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Style = "Marquee"
                $progressBar.Location = New-Object System.Drawing.Point(20, 50)
                $progressBar.Size = New-Object System.Drawing.Size(250, 20)
                $loadingForm.Controls.Add($progressBar)

                $loadingForm.Show()

                while (-not $SyncHash.CubeReady) {
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 50
                }
                $loadingForm.Close()
                $loadingForm.Dispose()
            }

            # Create a new runspace and set its ApartmentState to STA *before* it's opened.
            $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $runspace.ApartmentState = "STA"
            $runspace.Open()

            # Create a PowerShell object and associate it with our pre-configured runspace.
            $ps = [PowerShell]::Create().AddScript($loadingScriptBlock).AddArgument($SyncHash)
            $ps.Runspace = $runspace
            $loadingJob = $ps.BeginInvoke()

            $SyncHash.VideoFiles = $selectedFiles
            $SyncHash.CurrentIndex = $selectedFiles.Count # Start index for cycling
            if ($RadioButton1.Checked) { $SyncHash.RbSelection = "Hidden" }

            [xml]$xaml = @"
<Window x:Name="ImageWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PowerShell WPF Auto-Scrolling Media"
        WindowStartupLocation="CenterScreen"
        Width="384" Height="768"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <ScrollViewer x:Name="ImageScrollViewer" Grid.Row="0"
                      HorizontalScrollBarVisibility="Hidden"
                      VerticalScrollBarVisibility="Hidden">
            <StackPanel x:Name="ImagePanel"
                        Orientation="Vertical"
                        HorizontalAlignment="Center"
                        VerticalAlignment="Top">
                <!-- Media elements will be added by PowerShell -->
            </StackPanel>
        </ScrollViewer>
        
        <Button Content="X" 
                Canvas.Left="999" 
                Canvas.Top="0"
                FontSize="16"
                Width="25"
                Height="25"
                Name="CloseButton"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"/>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Background="#22000000">
            <Button Content="Scroll Slower" Name="SlowDown" Margin="2" Padding="3"/>
            <Button Content="Scroll Faster" Name="SpeedUp" Margin="2" Padding="3"/>
            <Button Content="Pause" Name="PauseButton" Margin="2" Padding="3"/>
            <Button Content="Reverse" Name="ReverseButton" Margin="2" Padding="3"/>
            <Button Content="Hide Controls" Name="HideControlsButton" Margin="2" Padding="3"/>
            <Button Content="Redo" Name="ReDoButton" Margin="2" Padding="3"/>
        </StackPanel>
    </Grid>
</Window>
"@

            $primaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
            $workingAreaWidth = $primaryScreen.WorkingArea.Width
            $workingAreaHeight = $primaryScreen.WorkingArea.Height

            $reader = (New-Object System.Xml.XmlNodeReader $xaml)
            $Window = [Windows.Markup.XamlReader]::Load($reader)
            $SyncHash.Window = $Window

            $imageScrollViewer = $Window.FindName("ImageScrollViewer")
            $imagePanel = $Window.FindName("ImagePanel")
            $CloseButton = $Window.FindName("CloseButton")
            $SyncHash.CloseButton = $CloseButton
            $ReDoButton = $Window.FindName("ReDoButton")
            $SlowDown = $Window.FindName("SlowDown")
            $SpeedUp = $Window.FindName("SpeedUp")
            $HideControlsButton = $Window.FindName("HideControlsButton")
            $ReverseButton = $Window.FindName("ReverseButton")
            $PauseButton = $Window.FindName("PauseButton")

            # Set the window to be the full vertical height of the screen and centered.
            $Window.Height = $workingAreaHeight
            $Window.Width = 384 # A fixed width for the vertical scroller
            if($MaximizeCheckBox.Checked -eq $true) {
                $Window.Width = $workingAreaWidth
                $Window.Height = $workingAreaHeight
            }

            # Apply transparency if the checkbox is checked
            if ($TransparentCheckbox.Checked) {
                $imagePanel.Opacity = 0.7
            }
            $SyncHash.Paused = $false
                
            $CloseButton.Add_Click({ $Window.Close() })

            $SyncHash.PlayerState = [hashtable]::Synchronized(@{})
            $ImageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")
            $SyncHash.ImageExtensions = $ImageExtensions
            $SyncHash.VideoUris = $SyncHash.VideoFiles | ForEach-Object { New-Object System.Uri($_, [System.UriKind]::Absolute) }

            $SyncHash.HandleMediaFailure = {
                param($ErrorElement, [string]$Reason = "Unknown Error")

                $SyncHash.Window.Dispatcher.Invoke([action]{
                    $playerState = $SyncHash.PlayerState[$ErrorElement.Name] # ErrorElement is the Grid
                    if ($playerState.IsFailed) { return } 
                    $playerState.IsFailed = $true

                    $fileName = if ($playerState.CurrentFile) { [System.IO.Path]::GetFileName($playerState.CurrentFile) } else { "an unknown media file" }
                    $errorText = "Error: $($fileName)`n$Reason"

                    $errorTextBlock = $ErrorElement.Children | Where-Object { $_.Name -like "errorTextBlock*" }
                    
                    if ($errorTextBlock) {
                        $errorTextBlock.Text = $errorText
                        $errorTextBlock.Visibility = "Visible"
                    }

                    if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }

                    $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $recoveryTimer.Interval = [TimeSpan]::FromSeconds(10)
                    $recoveryTimer.Tag = $ErrorElement 

                    $recoveryTick = {
                        $timer = $args[0]; $failedElement = $timer.Tag; $timer.Stop()
                        $SyncHash.PlayerState[$failedElement.Name].IsFailed = $false
                        & $SyncHash.MediaEndedHandler -Sender $failedElement -e $null -IsRecovery
                    }
                    $recoveryTimer.Add_Tick($recoveryTick)
                    $playerState.RecoveryTimer = $recoveryTimer
                    $recoveryTimer.Start()
                })
            }

            $MediaEndedHandler = {
                param(
                    [System.Windows.Controls.Grid]$Sender, # The sender is now the Grid
                    $e,
                    [switch]$IsRecovery
                )
                $FinishedElement = $Sender
                
                $playerState = $SyncHash.PlayerState[$FinishedElement.Name]
                if ($playerState.IsFailed) { return }

                if (-not $IsRecovery) {
                    $playerState.PlaybackStopwatch.Stop()
                    $elapsedMilliseconds = $playerState.PlaybackStopwatch.Elapsed.TotalMilliseconds
                    if (($elapsedMilliseconds -lt 2000) -and (-not $playerState.IsImage)) {
                        & $SyncHash.HandleMediaFailure -ErrorElement $FinishedElement -Reason "Playback failed or ended instantly."
                        return
                    }
                }
                if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) {
                    $playerState.FfmpegProcess.Kill()
                    $playerState.FfmpegTimer.Stop()
                }
                if ($playerState.FfmpegEventId) {
                    # Unregister the event handler to prevent conflicts when the slot is reused
                    Unregister-Event -SourceIdentifier $playerState.FfmpegEventId -ErrorAction SilentlyContinue
                    $playerState.FfmpegEventId = $null
                }
                
                # $FinishedElement is the Grid, so we search its children.
                $errorTextBlock = $FinishedElement.Children | Where-Object { $_.Name -like "errorTextBlock*" }
                if ($errorTextBlock) { $errorTextBlock.Visibility = "Collapsed" }

                if($SyncHash.CurrentIndex -ge $SyncHash.VideoUris.Count) {
                    $SyncHash.CurrentIndex = 0
                }

                $NewUri = $SyncHash.VideoUris[$SyncHash.CurrentIndex]
                $extension = [System.IO.Path]::GetExtension($NewUri.LocalPath).ToLower()
                $isImage = $SyncHash.ImageExtensions -contains $extension

                $playerState.CurrentFile = $NewUri.LocalPath
                $imageElement = $FinishedElement.Children | Where-Object { $_ -is [System.Windows.Controls.Image] }
                $mediaTextBlock = $FinishedElement.Children | Where-Object { $_.Name -like "mediaText_*" }
                
                if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }

                $playerState.IsImage = $isImage

                if ($isImage) {
                    $imageElement.Visibility = "Visible"
                    
                    $bitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bitmapImage.BeginInit()
                    $bitmapImage.add_DownloadCompleted({
                        param($s, $ev)
                        $img = $s
                        if ($img.PixelHeight -gt 0) {
                            $aspectRatio = $img.PixelWidth / $img.PixelHeight
                            if ($MaximizeCheckBox.Checked -eq $true) {
                                if ($img.PixelWidth -ge $img.PixelHeight) { # Landscape or square
                                    $FinishedElement.Width = $workingAreaWidth
                                    $FinishedElement.Height = $workingAreaWidth / $aspectRatio
                                } else { # Portrait
                                    $FinishedElement.Height = $workingAreaHeight
                                    $FinishedElement.Width = $workingAreaHeight * $aspectRatio
                                }
                            } else { # Not maximized
                                $FinishedElement.Width = 384
                                $FinishedElement.Height = 384 / $aspectRatio
                            }
                        }
                    })
                    $bitmapImage.UriSource = $NewUri
                    $bitmapImage.EndInit()
                    $imageElement.Source = $bitmapImage
                    
                    $playerState.ImageTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $playerState.ImageTimer.Interval = [TimeSpan]::FromSeconds(10)
                    $playerState.ImageTimer.Tag = $FinishedElement
                    
                    $tickScriptBlock = {
                        $timer = $args[0]; $element = $timer.Tag; $timer.Stop()
                        & $SyncHash.MediaEndedHandler -Sender $element -e $null
                    }
                    $playerState.ImageTimer.Add_Tick($tickScriptBlock)
                    $playerState.ImageTimer.Start()
                } else {
                    $imageElement.Visibility = "Collapsed" # Hide image element to show video
                    & $SyncHash.StartFfmpegStream -Grid $FinishedElement -FilePath $NewUri.LocalPath
                }

                switch ($SyncHash.RbSelection) {
                    "Hidden" { $mediaTextBlock.Visibility = "Collapsed" }
                    "Filename" {
                        $mediaTextBlock.Text = $NewUri.Segments[-1]
                        $mediaTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B))
                        $mediaTextBlock.FontSize = $SyncHash.SelectedFontSize
                        $mediaTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.SelectedFont)
                        $mediaTextBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
                        $mediaTextBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
                        $mediaTextBlock.Visibility = "Visible"
                    }
                    "Custom" {
                        $mediaTextBlock.Text = $SyncHash.TextBox.Text
                        $mediaTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B))
                        $mediaTextBlock.FontSize = $SyncHash.SelectedFontSize
                        $mediaTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.SelectedFont)
                        $mediaTextBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
                        $mediaTextBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
                        $mediaTextBlock.Visibility = "Visible"
                    }
                }

                $SyncHash.CurrentIndex++
            }
            $SyncHash.MediaEndedHandler = $MediaEndedHandler

            $SyncHash.StartFfmpegStream = {
                param(
                    [System.Windows.Controls.Grid]$Grid,
                    [string]$FilePath
                )

                $imageElement = $Grid.Children | Where-Object { $_ -is [System.Windows.Controls.Image] }
                $playerState = $SyncHash.PlayerState[$Grid.Name]

                $width = [int]$Grid.Width; $height = [int]$Grid.Height
                $frameSize = $width * $height * 3
                $bitmap = [Windows.Media.Imaging.WriteableBitmap]::new($width, $height, 96, 96, [Windows.Media.PixelFormats]::Bgr24, $null)
                
                $imageElement.Source = $bitmap
                $imageElement.Visibility = "Visible"

                $args = "-hide_banner -loglevel error -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "ffmpeg.exe"
                $psi.Arguments = $args
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $proc = [System.Diagnostics.Process]::Start($psi)

                # Enable event to handle process exit
                $proc.EnableRaisingEvents = $true
                $exitAction = {
                    if ($proc.ExitCode -ne 0) {
                        & $SyncHash.HandleMediaFailure -ErrorElement $Grid -Reason "ffmpeg process exited with code $($proc.ExitCode)."
                    }
                }
                Register-ObjectEvent -InputObject $proc -EventName Exited -Action $exitAction | Out-Null

                $stream = $proc.StandardOutput.BaseStream
                $bytes = New-Object byte[] $frameSize
                $rect = [System.Windows.Int32Rect]::new(0, 0, $width, $height)
                $stride = $width * 3

                $timer = New-Object Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(33) # ~30fps

                $tickScriptBlock = {
                    $total = 0
                    while ($total -lt $frameSize) {
                        $n = $stream.Read($bytes, $total, $frameSize - $total)
                        if ($n -le 0) { # End of stream
                            $timer.Stop()
                            & $SyncHash.MediaEndedHandler -Sender $Grid -e $null
                            return 
                        }
                        $total += $n
                    }
                    if ($total -eq $frameSize) {
                        $bitmap.Lock()
                        $bitmap.WritePixels($rect, $bytes, $stride, 0)
                        $bitmap.Unlock()
                    }
                }.GetNewClosure()

                $timer.Add_Tick($tickScriptBlock)
                $timer.Start()

                $playerState.FfmpegProcess = $proc
                $playerState.FfmpegTimer = $timer
            }

            $mediaIndex = 0
            foreach ($file in $SyncHash.VideoFiles)
            {
                $uri = New-Object System.Uri($file, [System.UriKind]::Absolute)
                $extension = [System.IO.Path]::GetExtension($file).ToLower()
                $isImage = $ImageExtensions -contains $extension

                $grid = New-Object System.Windows.Controls.Grid -Property @{ Name = "mediaGrid_$mediaIndex" }
                # Initial size is set here, but will be adjusted in MediaOpened/DownloadCompleted
                if ($MaximizeCheckBox.Checked -eq $true) {
                    $grid.Width = $workingAreaWidth
                    $grid.Height = $workingAreaHeight
                } else {
                    $grid.Width = 384 
                    $grid.Height = 384 # Default height, will be adjusted
                }
                $grid.Margin = New-Object System.Windows.Thickness(5)

                $imageElement = New-Object System.Windows.Controls.Image
                $imageElement.Name = "imageElement_$mediaIndex"
                $imageElement.Stretch = "Fill"
                $grid.Children.Add($imageElement)

                $mediaTextBlock = New-Object System.Windows.Controls.TextBlock
                $mediaTextBlock.Name = "mediaText_$mediaIndex"
                $mediaTextBlock.HorizontalAlignment = "Center"
                $mediaTextBlock.VerticalAlignment = "Top"
                $grid.Children.Add($mediaTextBlock)

                $errorTextBlock = New-Object System.Windows.Controls.TextBlock
                $errorTextBlock.Name = "errorTextBlock_$mediaIndex"
                $errorTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
                $errorTextBlock.VerticalAlignment = "Center"
                $errorTextBlock.HorizontalAlignment = "Center"
                $errorTextBlock.FontSize = 24
                $errorTextBlock.Visibility = "Collapsed"
                $grid.Children.Add($errorTextBlock)

                $SyncHash.PlayerState[$grid.Name] = @{
                    IsImage = $isImage
                    IsFailed = $false
                    ImageTimer = $null
                    RecoveryTimer = $null
                    FfmpegProcess     = $null
                    FfmpegTimer       = $null
                    CurrentFile       = $file
                    PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
                }

                if ($isImage) {
                    $bitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bitmapImage.BeginInit()
                    $bitmapImage.add_DownloadCompleted({
                        param($s, $ev)
                        $img = $s
                        if ($img.PixelHeight -gt 0) {
                            $aspectRatio = $img.PixelWidth / $img.PixelHeight
                            if ($MaximizeCheckBox.Checked -eq $true) {
                                if ($img.PixelWidth -ge $img.PixelHeight) { # Landscape or square
                                    $grid.Width = $workingAreaWidth
                                    $grid.Height = $workingAreaWidth / $aspectRatio
                                } else { # Portrait
                                    $grid.Height = $workingAreaHeight
                                    $grid.Width = $workingAreaHeight * $aspectRatio
                                }
                            } else { # Not maximized
                                $grid.Width = 384
                                $grid.Height = 384 / $aspectRatio
                            }
                        }
                    })
                    $bitmapImage.UriSource = $uri
                    $bitmapImage.EndInit()
                    $imageElement.Source = $bitmapImage
                    
                    $playerState = $SyncHash.PlayerState[$grid.Name]
                    $playerState.ImageTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $playerState.ImageTimer.Interval = [TimeSpan]::FromSeconds(10)
                    $playerState.ImageTimer.Tag = $grid
                    $tickScriptBlock = {
                        $timer = $args[0]; $element = $timer.Tag; $timer.Stop()
                        & $SyncHash.MediaEndedHandler -Sender $element -e $null
                    }
                    $playerState.ImageTimer.Add_Tick($tickScriptBlock)
                    $playerState.ImageTimer.Start()
                } else {
                    $imageElement.Visibility = "Collapsed"
                    & $SyncHash.StartFfmpegStream -Grid $grid -FilePath $file
                }

                switch ($SyncHash.RbSelection) {
                    "Hidden" { $mediaTextBlock.Visibility = "Collapsed" }
                    "Filename" {
                        $mediaTextBlock.Text = $uri.Segments[-1]
                        $mediaTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B))
                        $mediaTextBlock.FontSize = $SyncHash.SelectedFontSize
                        $mediaTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.SelectedFont)
                        $mediaTextBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
                        $mediaTextBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
                        $mediaTextBlock.Visibility = "Visible"
                    }
                    "Custom" {
                        $mediaTextBlock.Text = $SyncHash.TextBox.Text
                        $mediaTextBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B))
                        $mediaTextBlock.FontSize = $SyncHash.SelectedFontSize
                        $mediaTextBlock.FontFamily = New-Object System.Windows.Media.FontFamily($SyncHash.SelectedFont)
                        $mediaTextBlock.FontStyle = if ($SyncHash.ItalicCheckbox.Checked) { [System.Windows.FontStyles]::Italic } else { [System.Windows.FontStyles]::Normal }
                        $mediaTextBlock.FontWeight = if ($SyncHash.BoldCheckbox.Checked) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
                        $mediaTextBlock.Visibility = "Visible"
                    }
                    Default { $mediaTextBlock.Visibility = "Collapsed" }
                }

                $imagePanel.Children.Add($grid)
                $mediaIndex++
            }
            
            $SyncHash.Storyboard = $null
            $UpdateAnimation = {
                if ($SyncHash.Storyboard) { $SyncHash.Storyboard.Stop() }

                $totalHeight = 0
                foreach($child in $imagePanel.Children) {
                    # Use ActualHeight for a more reliable measurement after rendering
                    $totalHeight += $child.ActualHeight + $child.Margin.Top + $child.Margin.Bottom
                }
                
                if ($totalHeight -eq 0) { return } # Don't start if there's no content

                $from = if ($SyncHash.IsReversed) { -$totalHeight } else { 0 }
                $to = if ($SyncHash.IsReversed) { 0 } else { -$totalHeight }

                $myAnimation = $SyncHash.Storyboard.Children[0]
                $myAnimation.From = $from
                $myAnimation.To = $to
                
                $SyncHash.Storyboard.Begin($imagePanel, $true)
                if ($SyncHash.Paused) { $SyncHash.Storyboard.Pause($imagePanel) }
            }
            $SyncHash.UpdateAnimation = $UpdateAnimation

            $SlowDown.Add_Click({
                    $myAnimation = $SyncHash.Storyboard.Children | Where-Object { $_.Name -eq "myAnimation" }
                    $SyncHash.Storyboard.Stop()
                    $GetAnimation = $SyncHash.Storyboard.Children[0]
                    $GetDuration = $GetAnimation.Duration
                    $DurationSeconds = ($GetDuration.TimeSpan.TotalSeconds + 5)
                    $newDuration = New-Object System.Windows.Duration (New-Object System.TimeSpan (0, 0, $DurationSeconds))
                    $myAnimation.Duration = $newDuration
                    $SyncHash.Storyboard.Begin()
                })

            $SpeedUp.Add_Click({
                    $myAnimation = $SyncHash.Storyboard.Children | Where-Object { $_.Name -eq "myAnimation" }
                    $SyncHash.Storyboard.Stop()
                    $GetAnimation = $SyncHash.Storyboard.Children[0]
                    $GetDuration = $GetAnimation.Duration
                    $DurationSeconds = ($GetDuration.TimeSpan.TotalSeconds - 5)
                    if($DurationSeconds -le 0){ $DurationSeconds = 1 }
                    $newDuration = New-Object System.Windows.Duration (New-Object System.TimeSpan (0, 0, $DurationSeconds))
                    $myAnimation.Duration = $newDuration
                    $SyncHash.Storyboard.Begin()
                })
            
            # Create a reusable function to toggle pause/resume
            $TogglePause = {
                if ($SyncHash.Paused) {
                    $SyncHash.Storyboard.Resume($imagePanel)
                    $SyncHash.Paused = $false
                    $PauseButton.Content = "Pause"
                } else {
                    $SyncHash.Storyboard.Pause($imagePanel)
                    $SyncHash.Paused = $true
                    $PauseButton.Content = "Resume"
                }
            }

            # Attach the toggle function to the mouse click event
            $imageScrollViewer.add_PreviewMouseLeftButtonDown({ & $TogglePause })
            $PauseButton.Add_Click({ & $TogglePause })

            $SyncHash.IsReversed = $false
            $ReverseScrollDirection = {
                # Stop the animation to capture its current state
                $SyncHash.Storyboard.Stop($imagePanel)

                $myAnimation = $SyncHash.Storyboard.Children[0]
                $transform = $imagePanel.RenderTransform.Children[0]
                $currentY = $transform.Y

                # Recalculate total height, as it might have changed if media loaded late
                $totalHeight = 0
                foreach($child in $imagePanel.Children) {
                    $totalHeight += $child.ActualHeight + $child.Margin.Top + $child.Margin.Bottom
                }
                if ($totalHeight -eq 0) { return } # Safety check

                # Toggle the direction
                $SyncHash.IsReversed = -not $SyncHash.IsReversed

                # Set the new start and end points based on the current position
                $myAnimation.From = $currentY
                $myAnimation.To = if ($SyncHash.IsReversed) { 0 } else { -$totalHeight }

                # Recalculate the duration for the new path to maintain consistent speed
                $totalDistance = [Math]::Abs($myAnimation.To - $myAnimation.From)
                # Use the speed value stored in the synchronized hashtable
                $fullDurationSeconds = $SyncHash.AnimationSpeed * $totalHeight
                $newDurationSeconds = ($totalDistance / $totalHeight) * $fullDurationSeconds 
                $myAnimation.Duration = [TimeSpan]::FromSeconds($newDurationSeconds)

                $SyncHash.Storyboard.Begin($imagePanel, $true) # Begin animation from its current state
                if ($SyncHash.Paused) { $SyncHash.Storyboard.Pause($imagePanel) }
            }
            $ReverseButton.Add_Click($ReverseScrollDirection)

            $ToggleControlsVisibility = {
                if($SyncHash.ControlsHidden -eq $false)
                {
                    $SyncHash.ControlsHidden = $true
                    $ReDoButton.Visibility = "Hidden"
                    $CloseButton.Visibility = "Hidden"
                    $SlowDown.Visibility = "Hidden"
                    $SpeedUp.Visibility = "Hidden"
                    $HideControlsButton.Visibility = "Hidden"
                    $ReverseButton.Visibility = "Hidden"
                    $PauseButton.Visibility = "Hidden"
                }
                else
                {
                    $SyncHash.ControlsHidden = $false
                    $ReDoButton.Visibility = [System.Windows.Visibility]::Visible
                    $CloseButton.Visibility = [System.Windows.Visibility]::Visible
                    $SlowDown.Visibility = [System.Windows.Visibility]::Visible
                    $SpeedUp.Visibility = [System.Windows.Visibility]::Visible
                    $HideControlsButton.Visibility = [System.Windows.Visibility]::Visible
                    $ReverseButton.Visibility = [System.Windows.Visibility]::Visible
                    $ReverseButton.Visibility = [System.Windows.Visibility]::Visible
                    $PauseButton.Visibility = [System.Windows.Visibility]::Visible
                }
            }
            $HideControlsButton.Add_Click($ToggleControlsVisibility)

            $ReDoButton.Add_Click({
                    $Window.Close()
                    $SelectFolderForm.Show()
                })

            $Window.Add_KeyDown({
                    param($sender, $e)
                    switch ($e.Key)
                    {
                        "Left" { $SpeedUp.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                        "Right" { $SlowDown.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
                        "Escape"{ $SelectFolderForm.Dispose(); $Window.Close() }
                        "p"  { & $TogglePause }
                        "r"  { $Window.Close(); $SelectFolderForm.Show() }
                        "h"  { & $ToggleControlsVisibility }
                        "c"  { & $ReverseScrollDirection }
                        "F1"
                        { 
                            $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                            $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
                            $OkButton = $PopupWindow.FindName("OKButton")
                            $OkButton.Add_Click({ $PopupWindow.Close() })
                            $PopupWindow.ShowDialog() | Out-Null
                        }
                    }
                })

            $Window.Add_Loaded({
                $Window.Dispatcher.InvokeAsync([action]{
                    #region Animation setup
                    $scrollSpeedSeconds = 30
                    if ($MaximizeCheckBox.Checked -eq $true)
                    {
                        $ScrollSpeed = (30 * 2.5) # Default value of 30, scaled for maximized view
                        $duration = New-Object System.TimeSpan(0, 0, $ScrollSpeed)
                    }
                    else
                    {
                        # Use a default scroll duration of 30 seconds
                        $duration = New-Object System.TimeSpan(0, 0, $scrollSpeedSeconds)
                    }
                    
                    $Animation = New-Object System.Windows.Media.Animation.DoubleAnimation
                    $Animation.Duration = New-Object System.Windows.Duration($duration)
                    $Animation.Name = "myAnimation"
                    $Animation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever

                    # Store speed for dynamic duration calculation on reverse
                    $totalHeight = 0
                    foreach($child in $imagePanel.Children) {
                        $totalHeight += $child.ActualHeight + $child.Margin.Top + $child.Margin.Bottom
                    }
                    if ($totalHeight -gt 0) {
                        # Store the calculated speed in the synchronized hashtable instead of on the animation object
                        $SyncHash.AnimationSpeed = $duration.TotalSeconds / $totalHeight
                    }

                    $storyboard = New-Object System.Windows.Media.Animation.Storyboard
                    $storyboard.Children.Add($Animation)

                    [System.Windows.Media.Animation.Storyboard]::SetTargetProperty($Animation,
                        (New-Object System.Windows.PropertyPath("(UIElement.RenderTransform).(TransformGroup.Children)[0].(TranslateTransform.Y)")))
                    [System.Windows.Media.Animation.Storyboard]::SetTarget($Animation, $imagePanel)

                    $imagePanel.RenderTransform = New-Object System.Windows.Media.TransformGroup
                    $translateTransform = New-Object System.Windows.Media.TranslateTransform(0, 0)
                    $imagePanel.RenderTransform.Children.Add($translateTransform)
                    $SyncHash.Storyboard = $storyboard
                    #endregion

                    # Initial animation calculation
                    & $UpdateAnimation
                })
                })
            
            $Window.Add_Closed({
                foreach($playerState in $SyncHash.PlayerState.Values) {
                    if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
                    if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                    if ($playerState.FfmpegTimer) { $playerState.FfmpegTimer.Stop() }
                    if ($playerState.FfmpegProcess -and -not $playerState.FfmpegProcess.HasExited) {
                        $playerState.FfmpegProcess.Kill()
                    }
                    # Clean up any lingering event subscriptions
                    Get-EventSubscriber | Where-Object { $_.SourceIdentifier -like 'ffmpegExit_*' } | Unregister-Event
                }
                # Clear the synchronized hashtable to ensure a clean state for the next run
                $SyncHash.Clear()
            })

            $SelectFolderForm.Hide()

            # Signal the loading form's runspace that the main window is ready.
            $SyncHash.CubeReady = $true
            # Give the loading form a moment to close before showing the main window.
            Start-Sleep -Milliseconds 200

            $Window.ShowDialog() | Out-Null
        }
    })

[xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rich Text Popup" Height="340" Width="450" WindowStartupLocation="CenterScreen">
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
                    <Run Text="Hopefully selection dialog is self explanatory. :-)"/><LineBreak/>
                    <Run Text=" "/>
                </Paragraph>
                <Paragraph>
                    <Run Text="Commands for after scrolling starts:"/><LineBreak/>
                </Paragraph>    
                <Paragraph TextAlignment="Left" FontFamily="Consolas">
                    <Bold>
                        <Run Text="Button            : Key : Action" TextDecorations="Underline"/><LineBreak/>
                    </Bold>
                    <Run Text="X                 : Esc : Exit"/><LineBreak/>
                    <Run Text="Pause             :  P  : Pause Scrolling"/><LineBreak/>
                    <Run Text="Redo              :  R  : Reselect Images/Videos"/><LineBreak/>
                    <Run Text="Reverse           :  C  : Change Scrolling Direction"/><LineBreak/>
                    <Run Text="Hide Controls     :  H  : Hide/Show Controls"/><LineBreak/>
                    <Run Text="Scroll Slower     :  &#x2192;  : Slow Down Scrolling"/><LineBreak/>
                    <Run Text="Scroll Faster     :  &#x2190;  : Speed Up Scrolling"/><LineBreak/><LineBreak/>
                    <Run Text="*Click media to Pause/Resume"/><LineBreak/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" 
                HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$HelpLabel = New-Object System.Windows.Forms.Label
$HelpLabel.Text = "F1 - Help"
$HelpLabel.AutoSize = $True
$HelpLabel.Location = New-Object System.Drawing.Point(700, 0)
$HelpLabel.Size = New-Object System.Drawing.Size(150, 20)
$SelectFolderForm.Controls.Add($HelpLabel)

$SelectFolderForm.ShowDialog() | Out-Null
