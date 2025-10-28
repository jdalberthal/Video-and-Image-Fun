<#
.SYNOPSIS
    Displays selected images and videos on the faces of a rotating 3D cube using FFmpeg.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto the faces
    of a rotating 3D cube in a WPF window.

    It uses a sophisticated approach for video playback by leveraging FFmpeg to decode video
    frames in real-time and stream them to a WriteableBitmap. This bitmap is then applied as a
    texture to the cube's faces, allowing for broad video format support beyond what the standard
    Windows MediaElement can handle.

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays on the cube
    faces.

.EXAMPLE
    PS C:\> .\Show-RotatingImageVideoCubeFfmpeg.ps1

    Launches the file selection GUI. After selecting at least 6 files and clicking "Play", the
    script will launch the 3D cube window.

.NOTES
    Name:           Show-RotatingImageVideoCubeFfmpeg.ps1
    Version:        1.0.0, 10/18/2025
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

$ExternalButtonName = "Rotating Images/Videos Cube `n Uses Ffmpeg"
$ScriptDescription = "Loops through and displays 6 or more selected images or videos on the faces of a rotating 3D cube. Uses FFmpeg for video decoding, providing broad format support."
$RequiredExecutables = @("ffmpeg.exe","ffplay.exe")

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

[System.Windows.Forms.Application]::EnableVisualStyles()

$SyncHash = [hashtable]::Synchronized(@{}) # For passing data between runspaces
$SyncHash.ControlsHidden = $False
$SyncHash.Paused = $False
$SyncHash.RbSelection = ""

$PrimaryScreen = [System.Windows.Forms.Screen]::PrimaryScreen
                            
# Get the screen's total bounds (resolution)
$ScreenWidth = $PrimaryScreen.Bounds.Width
$ScreenHeight = $PrimaryScreen.Bounds.Height

# Get the screen's working area (excluding taskbars and docked windows)
$WorkingAreaWidth = $PrimaryScreen.WorkingArea.Width
$WorkingAreaHeight = $PrimaryScreen.WorkingArea.Height

$MyFont = New-Object System.Drawing.Font("Arial", 12)

# Create the form
$SelectFolderForm = New-Object System.Windows.Forms.Form
$SelectFolderForm.Text = "Video/Image Selector"
$SelectFolderForm.Size = New-Object System.Drawing.Size(800, 800)
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
$TransparentFacesCheckbox.Location = New-Object System.Drawing.Point(450, 47)
$TransparentFacesCheckbox.Checked = $False
$SelectFolderForm.Controls.Add($TransparentFacesCheckbox)

# Create a DataGridView to display files
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(10, 80)
$DataGridView.Size = New-Object System.Drawing.Size(760, 350)
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
$SelectAllCheckbox.Location = New-Object System.Drawing.Point(10, 60)
$SelectAllCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$SelectAllCheckbox.Checked = $False
$SelectFolderForm.Controls.Add($SelectAllCheckbox) # Add to your form or panel

$HeaderLabel = New-Object System.Windows.Forms.Label
$HeaderLabel.Text = "Play Video"
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

$Event = {
    if ($RadioButton1.Checked)
    {
        $SyncHash.RbSelection = "Hidden"
        $TextBox.Visible = $False
        $CurrentColor.Visible = $False
        $ColorExample.Visible = $False
        $SelectColorButton.Visible = $False
        $SizeLabel.Visible = $False
        $NumericUpDown.Visible = $False
        $FontButton.Visible = $False
        $ItalicCheckbox.Visible = $False
        $BoldCheckbox.Visible = $False
    }
    elseif ($RadioButton2.Checked)
    {
        $SyncHash.RbSelection = "Filename"
        $TextBox.Visible = $False
        $CurrentColor.Visible = $True
        $ColorExample.Visible = $True
        $SelectColorButton.Visible = $True
        $SizeLabel.Visible = $True
        $NumericUpDown.Visible = $True
        $FontButton.Visible = $True
        $ItalicCheckbox.Visible = $True
        $BoldCheckbox.Visible = $True
    }
    elseif ($RadioButton3.Checked)
    {
        $SyncHash.RbSelection = "Custom"
        $TextBox.Visible = $True
        $CurrentColor.Visible = $True
        $ColorExample.Visible = $True
        $SelectColorButton.Visible = $True
        $SizeLabel.Visible = $True
        $NumericUpDown.Visible = $True
        $FontButton.Visible = $True
        $ItalicCheckbox.Visible = $True
        $BoldCheckbox.Visible = $True
    }
}

$RadioButton1.Add_Click($Event)
$RadioButton2.Add_Click($Event)
$RadioButton3.Add_Click($Event)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Point(140, 440) ### Location of the text box
$TextBox.Size = New-Object System.Drawing.Size(455, 310) ### Size of the text box
$TextBox.Multiline = $True ### Allows multiple lines of data
$TextBox.Visible = $False ### By hitting enter it creates a new line
$TextBox.ScrollBars = "Vertical" ### Allows for a vertical scroll bar if the list of text is too big for the window
$TextBox.Font = $MyFont
$SelectFolderForm.Controls.Add($TextBox)

$SyncHash.TextBox = $TextBox

# Method One
$SyncHash.TextColor = [PSCustomObject]@{
    A = 255
    R = 0
    G = 0
    B = 0
}

# Method Two
# $SyncHash.TextColor = [System.Drawing.Color]::Black

$CurrentColor = New-Object System.Windows.Forms.Label
$CurrentColor.Text = "Text Color:"
$CurrentColor.Location = New-Object System.Drawing.Point(600, 477) # Position at top-left
$CurrentColor.AutoSize = $True # Ensures the label resizes to fit the text
$CurrentColor.BackColor = [System.Drawing.Color]::Transparent
$CurrentColor.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$CurrentColor.Visible = $False
$SelectFolderForm.Controls.Add($CurrentColor)

$ColorExample = New-Object System.Windows.Forms.Label
$ColorExample.Text = "     "
$ColorExample.Location = New-Object System.Drawing.Point(660, 477) # Position at top-left
$ColorExample.AutoSize = $True # Ensures the label resizes to fit the text
$ColorExample.BackColor = [System.Drawing.Color]::Black
$ColorExample.Visible = $False
$SelectFolderForm.Controls.Add($ColorExample)

# Create a button to open the color dialog
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
$SizeLabel.Size = New-Object System.Drawing.Size(25, 20)
$SizeLabel.Visible = $False
$SelectFolderForm.Controls.Add($SizeLabel)

# Create a NumericUpDown control
$NumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$NumericUpDown.Location = New-Object System.Drawing.Point(660, 520)
$NumericUpDown.Size = New-Object System.Drawing.Size(40, 20)
$NumericUpDown.Visible = $False

# Configure NumericUpDown properties
$NumericUpDown.Minimum = 0           # Set the minimum allowed value
$NumericUpDown.Maximum = 600          # Set the maximum allowed value
$NumericUpDown.Increment = 1         # Set the increment/decrement step
$NumericUpDown.DecimalPlaces = 0      # Set the number of decimal places (0 for integers)
$NumericUpDown.Value = 12           # Set the initial value
$SelectFolderForm.Controls.Add($NumericUpDown)
$SyncHash.NumericUpDown = $NumericUpDown
$SyncHash.SelectedFontSize = $SyncHash.NumericUpDown.Value

$NumericUpDown.Add_ValueChanged({
        param($sender, $e)
        $SyncHash.SelectedFontSize = $sender.Value
        # $SyncHash.TextBox.FontSize = $sender.Value

        $newFont = New-Object System.Drawing.Font($SyncHash.TextBox.Font.FontFamily, $sender.Value, $SyncHash.TextBox.Font.Style)
        $SyncHash.TextBox.Font = $newFont

    })

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
$SelectFolderForm.Controls.Add($ItalicCheckbox) # Add to your form or panel

$SyncHash.ItalicCheckbox = $ItalicCheckbox

$BoldCheckbox = New-Object System.Windows.Forms.CheckBox
$BoldCheckbox.Text = "Bold"
$BoldCheckbox.Location = New-Object System.Drawing.Point(680, 620)
$BoldCheckbox.Size = New-Object System.Drawing.Size(75, 20)
$BoldCheckbox.Checked = $False
$BoldCheckbox.Visible = $False
$SelectFolderForm.Controls.Add($BoldCheckbox) # Add to your form or panel

$SyncHash.BoldCheckbox = $BoldCheckbox

$HelpLabel = New-Object System.Windows.Forms.Label
$HelpLabel.Text = "F1 - Help"
$HelpLabel.AutoSize = $True
$HelpLabel.Location = New-Object System.Drawing.Point(700, 0)
$HelpLabel.Size = New-Object System.Drawing.Size(150, 20)
$SelectFolderForm.Controls.Add($HelpLabel)

$FontButton.Add_Click({
        $SyncHash.SelectedFont = "Arial"

        # Create the main form
        $FontForm = New-Object System.Windows.Forms.Form
        $FontForm.Text = "Font Selector"
        $FontForm.Size = New-Object System.Drawing.Size(400, 500)
        $FontForm.StartPosition = 'CenterScreen'

        # Create a ListBox and set its properties
        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Location = New-Object System.Drawing.Point(10, 10)
        $listBox.Size = New-Object System.Drawing.Size(360, 400)
        $listBox.Anchor = 'Top, Left, Bottom, Right'
        $listBox.DrawMode = 'OwnerDrawFixed' # Set to owner-drawn mode
        $listBox.ItemHeight = 20 # Adjust item height

        # Create a text box to preview the selected font
        $SelectButton = New-Object System.Windows.Forms.Button
        $SelectButton.Text = "Select"
        $SelectButton.Location = New-Object System.Drawing.Point(130, 420)
        $SelectButton.Size = New-Object System.Drawing.Size(100, 25)

        $SelectButton.Add_Click({
                param($sender, $e)

                if ($listBox.SelectedItem)
                {
                    # Update preview text box font (with error handling)
                    try { $SelectButton.Font = New-Object System.Drawing.Font($listBox, 12, [System.Drawing.FontStyle]::Regular) }
                    catch { $SelectButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular) }
                }
                $SyncHash.SelectedFont = $listBox.SelectedItem
                $SyncHash.FontButton.Font = $listBox.SelectedItem
                $TextBox.Font = $listBox.SelectedItem
                $FontForm.Dispose()
            })

        # --- DrawItem event handler ---
        $listBox.add_DrawItem({
                param($sender, $e)

                $e.DrawBackground()
                $e.DrawFocusRectangle()

                # Get item text and create font object (with error handling)
                $fontName = $sender.Items[$e.Index]
                try { $font = New-Object System.Drawing.Font($fontName, 12, [System.Drawing.FontStyle]::Regular) }
                catch { $font = $e.Font }

                # Draw the string
                $brush = New-Object System.Drawing.SolidBrush($e.ForeColor)
                $e.Graphics.DrawString($fontName, $font, $brush, $e.Bounds.Left + 2, $e.Bounds.Top + 2)
            })

        # --- SelectedIndexChanged event handler ---
        $listBox.add_SelectedIndexChanged({
                param($sender, $e)

                if ($sender.SelectedItem)
                {

                    # Update preview text box font (with error handling)
                    try { $SelectButton.Font = New-Object System.Drawing.Font($SyncHash.SelectedFont, 12, [System.Drawing.FontStyle]::Regular) }
                    catch { $SelectButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular) }
                }
            })

        # Populate the ListBox with installed font families
        $installedFonts = New-Object System.Drawing.Text.InstalledFontCollection
        foreach ($fontFamily in $installedFonts.Families)
        {
            [void]$listBox.Items.Add($fontFamily.Name)
        }

        # Optional: Select the first item by default
        if ($listBox.Items.Count -gt 0)
        {
            $defaultFontName = "Arial"
            if ($listBox.Items.Contains($defaultFontName))
            {
                $listBox.SelectedItem = $defaultFontName
            }
            else
            {
                $listBox.SelectedIndex = 0
            }
        }

        # Add controls to the form
        $FontForm.Controls.Add($listBox)
        $FontForm.Controls.Add($SelectButton)

        # Show the form and clean up
        [void]$FontForm.ShowDialog()
        $FontForm.Dispose()
    })

# Create a ColorDialog object
$colorDialog = New-Object System.Windows.Forms.ColorDialog

# Add an event handler for the button click
$SelectColorButton.Add_Click({
        # Show the color dialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            # If a color is selected, set the form's background color
            $ColorExample.BackColor = $colorDialog.Color
            $TextBox.ForeColor = $colorDialog.Color #[System.Drawing.Color]::Blue
            $SyncHash.TextColor = $colorDialog.Color
        }
    })

$SelectAllCheckbox.Add_CheckedChanged({
        $CheckedState = $SelectAllCheckbox.Checked
        foreach ($Row in $DataGridView.Rows)
        {
            # Assuming 'SelectColumn' is the name of your checkbox column
            $Row.Cells["Select"].Value = $CheckedState
        }
        # Commit the changes immediately if needed
        $DataGridView.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    })
    
# Event handler for the Browse Folder button
$BrowseButton.Add_Click({

        $VideoExtensions = "*.webm", "*.mkv", "*.flv", "*.vob", "*.ogv", "*.ogg", "*.rrc", "*.gifv", "*.mng", "*.mov",
        "*.avi", "*.qt", "*.wmv", "*.yuv", "*.rm", "*.asf", "*.amv", "*.mp4", "*.m4p", "*.m4v", "*.mpg", "*.mp2",
        "*.mpeg", "*.mpe", "*.mpv", "*.m4v", "*.svi", "*.3gp", "*.3g2", "*.mxf", "*.roq", "*.nsv", "*.flv", "*.f4v",
        "*.f4p", "*.f4a", "*.f4b", "*.mod", "*.wtv", "*.hevc", "*.m2ts", "*.m2v", "*.m4v", "*.mjpeg", "*.mts", "*.rm",
        "*.ts", "*.vob"#, "*.swf"

        $ImageExtensions = "*.bmp", "*.jpeg", "*.jpg", "*.png", "*.tif", "*.tiff", "*.gif", "*.wmp", "*.ico"

        $AllowedExtension = $VideoExtensions + $ImageExtensions

        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select the folder to scan."

        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $SelectedPath = "$($FolderBrowser.SelectedPath)"
            $FolderPathTextBox.Text = $SelectedPath

            # Clear existing rows in DataGridView
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

            foreach ($row in  $DataGridView.Rows)
            {
                if ($row.IsNewRow) { continue } # Skip the blank row at the bottom
                $row.HeaderCell.Value = "Play"
            }
        }
    })

$ItalicCheckbox.Add_CheckedChanged({
        if($ItalicCheckbox.Checked -and $BoldCheckbox.Checked)
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
        }
        elseif(($ItalicCheckbox.Checked))
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Italic)
        }
        elseif((-not $ItalicCheckbox.Checked -and $BoldCheckbox.Checked))
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Bold)
        }
        else
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Regular)
        }
    })

$BoldCheckbox.Add_CheckedChanged({
        if($ItalicCheckbox.Checked -and $BoldCheckbox.Checked)
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
        }
        elseif(($BoldCheckbox.Checked))
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Bold)
        }
        elseif((-not $BoldCheckbox.Checked -and $ItalicCheckbox.Checked))
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Italic)
        }
        else
        {
            $TextBox.Font = New-Object System.Drawing.Font($TextBox.Font.FontFamily, $TextBox.Font.Size, [System.Drawing.FontStyle]::Regular)
        }
    })

[xml]$XamlHelpPopup = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Rich Text Popup" Height="340" Width="400" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <RichTextBox x:Name="MyRichTextBox" Grid.Row="0" Margin="5" AcceptsReturn="False" VerticalScrollBarVisibility="Auto">
            <FlowDocument>
                <FlowDocument.Resources>
                    <Style TargetType="{x:Type Paragraph}">
                        <Setter Property="Margin" Value="0"/>
                    </Style>
                </FlowDocument.Resources>
                <Paragraph>
                    <Run Text="Hopefully selection diialog is self explanatory. :-)"/><LineBreak/>
                    <Run Text=" "/>
                </Paragraph>
                <Paragraph>
                    <Run Text="Commands for after video(s) are playing:"/><LineBreak/>
                </Paragraph>    
                <Paragraph TextAlignment="Left" FontFamily="Consolas">
                    <Bold>
                        <Run Text="Button        : Key : Action                     ." TextDecorations="Underline"/><LineBreak/>
                    </Bold>
                    <Run Text="X             : Esc : Exit"/><LineBreak/>
                    <Run Text="Pause         :  P  : Pause Cube Spinning"/><LineBreak/>
                    <Run Text="Redo          :  R  : Reselect videos"/><LineBreak/>
                    <Run Text="Random Axis   :  A  : Change Rotation Axis"/><LineBreak/>
                    <Run Text="Hide Controls :  H  : Hide Controls/Show Controls"/><LineBreak/>
                    <Run Text="Left Arrow    :  &#x2190;  : Slow Down Spinning"/><LineBreak/>
                    <Run Text="Right Arrow   :  &#x2192;  : Speed Up Spinning"/><LineBreak/><LineBreak/>
                    <Run Text="*Click Cube to Pause"/><LineBreak/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        
        <Button x:Name="OKButton" Grid.Row="1" Content="OK" 
                HorizontalAlignment="Right" Width="80" Height="30" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$SelectFolderForm.KeyPreview = $True
$SelectFolderForm.Add_KeyDown({
        param($Sender, $e)
        switch ($_.KeyCode)
        {
            "F1"
            {
                $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)

                # Get controls from the popup window
                $OkButton = $PopupWindow.FindName("OKButton")

                # Define OK button click event for the popup
                $OkButton.Add_Click({
                        # Closes the popup window
                        $PopupWindow.Close()
                    })

                # Show the popup window as a modal dialog
                $PopupWindow.ShowDialog() | Out-Null
            }
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

# Event handler for the Process Selected Files button
$PlayButton.Add_Click({
        $SyncHash.SelectedFiles = @()
        foreach ($Row in $DataGridView.Rows)
        {
            if ($Row.Cells["Select"].Value)
            {
                $SyncHash.SelectedFiles += $Row.Cells["FilePath"].Value
            }
        }

        if($SyncHash.SelectedFiles.Count -le 0)
        {
            [System.Windows.Forms.MessageBox]::Show("No Videos selected.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
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

            $SyncHash.VideoFiles = $SyncHash.SelectedFiles

            $SyncHash.CurrentIndex = 6

            # ---------------- XAML cube -----------------
            [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="3D Cube with Videos" Width="900" Height="700"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        SizeToContent="Manual"
        WindowState="Normal"
        AllowsTransparency="True"
        Background="Transparent">
    <Grid x:Name="MainGrid">
        <Viewport3D>
            <!-- Camera -->
            <Viewport3D.Camera>
                <PerspectiveCamera Position="4,4,8" LookDirection="-4,-4,-8"
                                   UpDirection="0,1,0" FieldOfView="60"/>
            </Viewport3D.Camera>

            <!-- Lighting -->
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <AmbientLight Color="Gray"/>
                        <DirectionalLight Color="White" Direction="-1,-1,-2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
            </ModelVisual3D>

            <!-- Cube with ScaleTransform3D -->
            <ModelVisual3D>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <!-- Scale factor to enlarge cube -->
                        <ScaleTransform3D ScaleX="3.5" ScaleY="3.5" ScaleZ="3.5"/>
                        <!-- Rotations -->
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
                            <RotateTransform3D>
                                <RotateTransform3D.Rotation>
                                    <AxisAngleRotation3D x:Name="AxisAngleZ" Axis="0,0,1" Angle="0" />
                                </RotateTransform3D.Rotation>
                            </RotateTransform3D>
                    </Transform3DGroup>
                </ModelVisual3D.Transform>

                <ModelVisual3D.Children>
                    <!-- FRONT (video) -->
                    <Viewport2DVisual3D x:Name="V2DV3D_0">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.5,-0.5,0.5  0.5,-0.5,0.5  0.5,0.5,0.5  -0.5,0.5,0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid0" Background="Transparent">
                            <ContentPresenter Name="Content0"/>
                            <TextBlock x:Name="FrontFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder0"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock0" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>

                    <!-- BACK (video) -->
                    <Viewport2DVisual3D x:Name="V2DV3D_1">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="0.5,-0.5,-0.5 -0.5,-0.5,-0.5 -0.5,0.5,-0.5 0.5,0.5,-0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid1" Background="Transparent">
                            <ContentPresenter Name="Content1"/>
                            <TextBlock x:Name="BackFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder1"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock1" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>

                    <!-- LEFT -->
                    <Viewport2DVisual3D x:Name="V2DV3D_2">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.5,-0.5,-0.5 -0.5,-0.5,0.5 -0.5,0.5,0.5 -0.5,0.5,-0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid2" Background="Transparent">
                            <ContentPresenter Name="Content2"/>
                            <TextBlock x:Name="LeftFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder3"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock3" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>

                    <!-- RIGHT -->
                    <Viewport2DVisual3D x:Name="V2DV3D_3">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="0.5,-0.5,0.5 0.5,-0.5,-0.5 0.5,0.5,-0.5 0.5,0.5,0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid3" Background="Transparent">
                            <ContentPresenter Name="Content3"/>
                            <TextBlock x:Name="RightFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder2"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock2" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>

                    <!-- TOP -->
                    <Viewport2DVisual3D x:Name="V2DV3D_4">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.5,0.5,0.5 0.5,0.5,0.5 0.5,0.5,-0.5 -0.5,0.5,-0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid4" Background="Transparent">
                            <ContentPresenter Name="Content4"/>
                            <TextBlock x:Name="TopFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder4"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock4" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>

                    <!-- BOTTOM -->
                    <Viewport2DVisual3D x:Name="V2DV3D_5">
                        <Viewport2DVisual3D.Geometry>
                            <MeshGeometry3D Positions="-0.5,-0.5,-0.5 0.5,-0.5,-0.5 0.5,-0.5,0.5 -0.5,-0.5,0.5"
                                            TextureCoordinates="0,1 1,1 1,0 0,0"
                                            TriangleIndices="0,1,2 0,2,3"/>
                        </Viewport2DVisual3D.Geometry>
                        <Viewport2DVisual3D.Material>
                            <DiffuseMaterial Viewport2DVisual3D.IsVisualHostMaterial="True" Brush="White"/>
                        </Viewport2DVisual3D.Material>
                        <Grid x:Name="FaceGrid5" Background="Transparent">
                            <ContentPresenter Name="Content5"/>
                            <TextBlock x:Name="BottomFaceText" 
                                        Text=""
                                        FontSize="24"
                                        FontWeight="Bold"
                                        Foreground="White"
                                        HorizontalAlignment="Center"
                                        VerticalAlignment="Top"
                                        IsHitTestVisible="False" />
                            <Border Name="ErrorBorder5"
                                    Panel.ZIndex="1"
                                    VerticalAlignment="Center"
                                    HorizontalAlignment="Center"
                                    Visibility="Collapsed"
                                    Width="200" Height="200">
                                <TextBlock Name="ErrorTextBlock5" TextWrapping="Wrap" TextAlignment="Center" Foreground="Red" VerticalAlignment="Center"/>
                            </Border>
                        </Grid>
                    </Viewport2DVisual3D>
                </ModelVisual3D.Children>
            </ModelVisual3D>
        </Viewport3D>
    <StackPanel  Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
            <Button
                Name="PauseButton"
                Content="Pause"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0"
                Width="90">
            </Button>
            <Button
                Name="RandomAxis"
                Content="Random Axis"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
            <!--Left: &#x2190; | Down: &#x2191; | Right: &#x2192; | Up: &#x2193;-->
            <Button
                Name="SlowDown"
                Content="&#x2190;"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
            <Button
                Name="SpeedUp"
                Content="&#x2192;"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
            <Button
                Name="ReDoButton"
                Content="Redo"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
            <Button
                Name="HideControls"
                Content="Hide Controls"
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                Padding="15,5"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
            <Button Content="X" 
                HorizontalAlignment="Right"
                VerticalAlignment="Top"
                FontSize="16"
                Width="28"
                Height="28"
                Name="CloseButton"
                Grid.Column="0"
                Grid.Row="0">
            </Button>
        </StackPanel>
    </Grid>
</Window>
"@
            $Xaml.Window.Width = "$WorkingAreaWidth"
            $Xaml.Window.Height = "$WorkingAreaHeight"

            $reader = New-Object System.Xml.XmlNodeReader $xaml
            $window = [Windows.Markup.XamlReader]::Load($reader)

            $SyncHash.CloseButton = $Window.FindName("CloseButton")
            $SyncHash.SlowDown = $Window.FindName("SlowDown")
            $SyncHash.SpeedUp = $Window.FindName("SpeedUp")
            $SyncHash.ReDoButton = $Window.FindName("ReDoButton")
            $SyncHash.HideControls = $Window.FindName("HideControls")
            $SyncHash.PauseButton = $Window.FindName("PauseButton")
            $SyncHash.RandomAxis = $Window.FindName("RandomAxis")
            $SyncHash.MainGrid = $Window.FindName("MainGrid")
            $SyncHash.Window = $Window
            $SyncHash.AxisAngleX = $Window.FindName("AxisAngleX")
            $SyncHash.AxisAngleY = $Window.FindName("AxisAngleY")
            $SyncHash.AxisAngleZ = $Window.FindName("AxisAngleZ")
            $SyncHash.FrontFaceText = $Window.FindName("FrontFaceText")
            $SyncHash.BackFaceText = $Window.FindName("BackFaceText")
            $SyncHash.RightFaceText = $Window.FindName("RightFaceText")
            $SyncHash.LeftFaceText = $Window.FindName("LeftFaceText")
            $SyncHash.TopFaceText = $Window.FindName("TopFaceText")
            $SyncHash.BottomFaceText = $Window.FindName("BottomFaceText")
            $SyncHash.ErrorTextBlock0 = $Window.FindName("ErrorTextBlock0")
            $SyncHash.ErrorTextBlock1 = $Window.FindName("ErrorTextBlock1")
            $SyncHash.ErrorTextBlock2 = $Window.FindName("ErrorTextBlock2")
            $SyncHash.ErrorTextBlock3 = $Window.FindName("ErrorTextBlock3")
            $SyncHash.ErrorTextBlock4 = $Window.FindName("ErrorTextBlock4")
            $SyncHash.ErrorTextBlock5 = $Window.FindName("ErrorTextBlock5")
            $SyncHash.ErrorBorder0 = $Window.FindName("ErrorBorder0")
            $SyncHash.ErrorBorder1 = $Window.FindName("ErrorBorder1")
            $SyncHash.ErrorBorder2 = $Window.FindName("ErrorBorder2")
            $SyncHash.ErrorBorder3 = $Window.FindName("ErrorBorder3")
            $SyncHash.ErrorBorder4 = $Window.FindName("ErrorBorder4")
            $SyncHash.ErrorBorder5 = $Window.FindName("ErrorBorder5")
            $SyncHash.V2DV3D_0 = $Window.FindName("V2DV3D_0")
            $SyncHash.V2DV3D_1 = $Window.FindName("V2DV3D_1")
            $SyncHash.V2DV3D_2 = $Window.FindName("V2DV3D_2")
            $SyncHash.V2DV3D_3 = $Window.FindName("V2DV3D_3")
            $SyncHash.V2DV3D_4 = $Window.FindName("V2DV3D_4")
            $SyncHash.V2DV3D_5 = $Window.FindName("V2DV3D_5")
            $SyncHash.FaceGrid0 = $Window.FindName("FaceGrid0")
            $SyncHash.FaceGrid1 = $Window.FindName("FaceGrid1")
            $SyncHash.FaceGrid2 = $Window.FindName("FaceGrid2")
            $SyncHash.FaceGrid3 = $Window.FindName("FaceGrid3")
            $SyncHash.FaceGrid4 = $Window.FindName("FaceGrid4")
            $SyncHash.FaceGrid5 = $Window.FindName("FaceGrid5")

            # ---------------- Animate cube on Y + Z -----------------
            $animX = New-Object Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(15))
            $animX.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
            $AxisAngleX = $window.FindName("AxisAngleX")
            $AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

            $animY = New-Object Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(15))
            $animY.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
            $AxisAngleY = $window.FindName("AxisAngleY")
            $AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

            $animZ = New-Object Windows.Media.Animation.DoubleAnimation(0, 360, [TimeSpan]::FromSeconds(30)) # slower
            $animZ.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
            $AxisAngleZ = $window.FindName("AxisAngleZ")
            $AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)

            $ChangeMaterial = {
                param($Viewport, $Index)
 
                # Generate two fully random RGB colors
                $Color1 = [System.Windows.Media.Color]::FromRgb((Get-Random -Maximum 256), (Get-Random -Maximum 256), (Get-Random -Maximum 256))
                $Color2 = [System.Windows.Media.Color]::FromRgb((Get-Random -Maximum 256), (Get-Random -Maximum 256), (Get-Random -Maximum 256))

                # Create GradientStops for the brush
                $GradientStop1 = New-Object Windows.Media.GradientStop($Color1, 0.0)
                $GradientStop2 = New-Object Windows.Media.GradientStop($Color2, 1.0)

                # Create a LinearGradientBrush for a prettier effect
                $GradientBrush = New-Object Windows.Media.LinearGradientBrush
                $GradientBrush.StartPoint = New-Object Windows.Point(0, 0) # Top-left
                $GradientBrush.EndPoint = New-Object Windows.Point(1, 1)   # Bottom-right
                $GradientBrush.GradientStops.Add($GradientStop1)
                $GradientBrush.GradientStops.Add($GradientStop2)
                
                # Create a new DiffuseMaterial with the gradient brush and apply it
                $DiffuseMaterial = New-Object Windows.Media.Media3D.DiffuseMaterial($GradientBrush)
                $Viewport.Material = $DiffuseMaterial
            }

            function Start-NextMediaOnFace {
                param(
                    [int]$FaceIndex
                )

                if ($SyncHash.CurrentIndex -ge $SyncHash.SelectedFiles.Count) {
                    $SyncHash.CurrentIndex = 0
                }
                $nextMedia = $SyncHash.SelectedFiles[$SyncHash.CurrentIndex]
                $SyncHash.CurrentIndex++

                # Update text if needed
                if ($SyncHash.RbSelection -eq "Filename") {
                    $faceTextBlockName = switch($FaceIndex) {
                        0 {"FrontFaceText"}
                        1 {"BackFaceText"}
                        2 {"LeftFaceText"}
                        3 {"RightFaceText"}
                        4 {"TopFaceText"}
                        5 {"BottomFaceText"}
                    }
                    $faceTextBlock = $SyncHash.Window.FindName($faceTextBlockName)
                    if ($faceTextBlock) {
                        $faceTextBlock.Text = (Split-Path -Path $nextMedia -Leaf)
                    }
                }

                # Determine if next media is image or video and start it
                $extension = [System.IO.Path]::GetExtension($nextMedia).ToLower()
                if ($ImageExtensions -contains $extension) {
                    & $SyncHash.StartImageOnFace -FaceIndex $FaceIndex -FilePath $nextMedia
                }
                else {
                    & $SyncHash.StartVideoOnFace -FaceIndex $FaceIndex -FilePath $nextMedia
                }
            }

            function Start-ImageOnFace {
                param(
                    [int]$FaceIndex,
                    [string]$FilePath
                )

                # Clean up previous media on this face if any
                if ($SyncHash.FaceData.ContainsKey($FaceIndex)) {
                    $oldFaceData = $SyncHash.FaceData[$FaceIndex]
                    if ($oldFaceData.Timer) { $oldFaceData.Timer.Stop() }
                    if ($oldFaceData.Process -and -not $oldFaceData.Process.HasExited) { $oldFaceData.Process.Kill() }
                }

                $image = New-Object Windows.Controls.Image
                $image.Source = [Windows.Media.Imaging.BitmapImage]::new(
                    [Uri]$FilePath
                )
                $image.Stretch = "Fill"
                ($window.FindName("Content$FaceIndex")).Content = $image

                # If more than 6 files, set a timer to cycle to the next one
                if ($SyncHash.SelectedFiles.Count -gt 6) {
                    $timer = New-Object Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]::FromSeconds(10)

                    $tickScriptBlock = {
                        $timer.Stop()
                        & $SyncHash.StartNextMediaOnFace -FaceIndex $FaceIndex
                    }
                    $timer.Add_Tick($tickScriptBlock.GetNewClosure())
                    $timer.Start()

                    $SyncHash.FaceData[$FaceIndex] = @{
                        Timer = $timer
                        Process = $null
                    }
                }
            }

            function Start-VideoOnFace {
                param(
                    [int]$FaceIndex,
                    [string]$FilePath
                )

                # Clean up previous media on this face if any
                if ($SyncHash.FaceData.ContainsKey($FaceIndex)) {
                    $oldFaceData = $SyncHash.FaceData[$FaceIndex]
                    if ($oldFaceData.Timer) { $oldFaceData.Timer.Stop() }
                    if ($oldFaceData.Process -and -not $oldFaceData.Process.HasExited) { $oldFaceData.Process.Kill() }
                }

                $width = 640
                $height = 480
                $frameSize = $width * $height * 3
                $bitmap = [Windows.Media.Imaging.WriteableBitmap]::new($width, $height, 96, 96, [Windows.Media.PixelFormats]::Bgr24, $null)
                $image = New-Object Windows.Controls.Image
                $image.Source = $bitmap
                $image.Stretch = "Fill"
                ($window.FindName("Content$FaceIndex")).Content = $image

                # Determine loop argument
                $loopArg = if ($SyncHash.SelectedFiles.Count -le 6) { "-stream_loop -1" } else { "" }
                $args = "-hide_banner -loglevel error $loopArg -i `"$FilePath`" -f rawvideo -pix_fmt bgr24 -vf scale=${width}:${height} -"

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "ffmpeg.exe"
                $psi.Arguments = $args
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $true
                $proc = [System.Diagnostics.Process]::Start($psi)

                $stream = $proc.StandardOutput.BaseStream
                $bytes = New-Object byte[] $frameSize
                $rect = [System.Windows.Int32Rect]::new(0, 0, $width, $height)
                $stride = $width * 3

                $timer = New-Object Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(33) # ~30fps

                # Define the script block for the timer's tick event.
                $tickScriptBlock = {
                    $total = 0
                    while ($total -lt $frameSize)
                    {
                        $n = $stream.Read($bytes, $total, $frameSize - $total)
                        if ($n -le 0) { # End of stream
                            if ($SyncHash.SelectedFiles.Count -gt 6) {
                                # Stop timer and kill process
                                $timer.Stop()
                                & $SyncHash.StartNextMediaOnFace -FaceIndex $FaceIndex
                            }
                            return # Exit the scriptblock for this tick
                        }
                        $total += $n
                    }
                    if ($total -eq $frameSize)
                    {
                        $bitmap.Lock()
                        $bitmap.WritePixels($rect, $bytes, $stride, 0)
                        $bitmap.Unlock()
                    }
                }

                $timer.Add_Tick($tickScriptBlock.GetNewClosure())
                $timer.Start()

                # Store for cleanup
                $SyncHash.FaceData[$FaceIndex] = @{
                    Timer = $timer
                    Process = $proc
                }
            }

            $SyncHash.StartNextMediaOnFace = ${function:Start-NextMediaOnFace}
            $SyncHash.StartVideoOnFace = ${function:Start-VideoOnFace}
            $SyncHash.StartImageOnFace = ${function:Start-ImageOnFace}

            # Arrays to hold timers and processes for cleanup
            $timers = @()
            $procs = @()

            # Store face-specific data
            $SyncHash.FaceData = [hashtable]::Synchronized(@{})

            # Determine which material to use based on the checkbox
            $materialType = if ($TransparentFacesCheckbox.Checked) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }

            # Apply the chosen material type to all faces
            for ($i = 0; $i -lt 6; $i++) {
                $viewport = $window.FindName("V2DV3D_$i")
                
                # Create a new material of the chosen type
                $newMaterial = New-Object $materialType
                
                # This is the crucial part: setting the attached property in code.
                # It tells the material to use the 2D visual as its texture.
                [System.Windows.Media.Media3D.Viewport2DVisual3D]::SetIsVisualHostMaterial($newMaterial, $true)
                
                # Assign the new material to the viewport
                $viewport.Material = $newMaterial
            }

            $ImageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico") # Make this global for Start-NextMediaOnFace

            for ($i = 0; $i -lt 6; $i++)
            {
                if($SyncHash.VideoFiles[$i] -eq $Null)
                {
                    & $ChangeMaterial ($window.FindName("V2DV3D_$i")) $i
                }
                else
                {
                    $filePath = $SyncHash.VideoFiles[$i]
                    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()

                    if ($ImageExtensions -contains $extension) {
                        # It's an image
                        & $SyncHash.StartImageOnFace -FaceIndex $i -FilePath $filePath
                    }
                    else {
                        # Assume it's a video and check it
                        $ErrorCheck = ffmpeg.exe -v error -xerror -i $filePath -f null - 2>&1

                        if ($ErrorCheck -match "error|invalid data")
                        {
                            $VideoPlayable = $False
                        }
                        else
                        {
                            $VideoPlayable = $True
                        }

                        if (-not $VideoPlayable)
                        {
                            # This is the definitive fix: replace the Visual of the Viewport2DVisual3D.
                            $viewport = $window.FindName("V2DV3D_$i")
                            $originalGrid = $viewport.Visual # Save the original grid with the ContentPresenter

                            # Create a new Grid just for the error message
                            $errorGrid = New-Object System.Windows.Controls.Grid
                            $errorGrid.Background = [System.Windows.Media.Brushes]::Black

                            # Create a fixed-size container for the text
                            $containerBorder = New-Object System.Windows.Controls.Border
                            $containerBorder.Width = 150
                            $containerBorder.Height = 150

                            $errorTextBlock = New-Object System.Windows.Controls.TextBlock
                            $errorTextBlock.Text = "Error playing video: `n$(Split-Path -Path $filePath -Leaf)"
                            $errorTextBlock.Foreground = [System.Windows.Media.Brushes]::Red
                            $errorTextBlock.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
                            $errorTextBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                            $errorTextBlock.TextAlignment = [System.Windows.TextAlignment]::Center
                            $errorTextBlock.TextWrapping = "Wrap"

                            # Place the TextBlock inside the fixed-size border, and the border inside the main grid
                            $containerBorder.Child = $errorTextBlock
                            $errorGrid.Children.Add($containerBorder)

                            # Set the new error grid as the visual source for the viewport
                            $viewport.Visual = $errorGrid

                            # If there are more files to cycle through, start a timer to show the next one
                            if ($SyncHash.SelectedFiles.Count -gt 6) {
                                $errorTimer = New-Object Windows.Threading.DispatcherTimer
                                $errorTimer.Interval = [TimeSpan]::FromSeconds(10)
                                $tickScriptBlock = {
                                    $errorTimer.Stop()
                                    $viewport.Visual = $originalGrid # Restore the original grid
                                    & $SyncHash.StartNextMediaOnFace -FaceIndex $i
                                }.GetNewClosure()
                                $errorTimer.Add_Tick($tickScriptBlock)
                                $errorTimer.Start()
                            }
                        }
                        else {
                            & $SyncHash.StartVideoOnFace -FaceIndex $i -FilePath $filePath
                        }
                    }
                }
            }

            switch ($SyncHash.RbSelection)
            {
                "Hidden"
                {
                    $SyncHash.FrontFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.BackFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.RightFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.LeftFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.TopFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.BottomFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                }
                "Filename"
                {
                    # Convert System.Drawing.Color to System.Windows.Media.Color
                    $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)

                    # Create a SolidColorBrush from the System.Windows.Media.Color
                    $SyncHash.FrontFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.BackFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.RightFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.LeftFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.TopFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.BottomFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)

                    $SyncHash.SelectedFontSize = $SyncHash.NumericUpDown.Value

                    $SyncHash.FrontFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.BackFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.RightFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.LeftFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.TopFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.BottomFaceText.FontSize = $SyncHash.SelectedFontSize

                    $SyncHash.FrontFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.BackFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.RightFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.LeftFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.TopFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.BottomFaceText.FontFamily = $SyncHash.SelectedFont

                    if($SyncHash.ItalicCheckbox.Checked -and $BoldCheckbox.Checked)
                    {
                        $StyleFont = [System.Windows.FontStyles]::Italic
                        $WeightFont = [System.Windows.FontWeights]::Bold
                    }
                    elseif(($SyncHash.ItalicCheckbox.Checked))
                    {
                        $StyleFont = [System.Windows.FontStyles]::Italic
                        $WeightFont = $Null
                    }
                    elseif(($SyncHash.BoldCheckbox.Checked))
                    {
                        $StyleFont = $Null
                        $WeightFont = [System.Windows.FontWeights]::Bold
                    }
                    else
                    {
                        $StyleFont = [System.Windows.FontStyles]::Normal
                        $WeightFont = $Null
                    }

                    if($Null -ne $WeightFont)
                    {
                        $SyncHash.FrontFaceText.FontWeight = $WeightFont
                        $SyncHash.BackFaceText.FontWeight = $WeightFont
                        $SyncHash.RightFaceText.FontWeight = $WeightFont
                        $SyncHash.LeftFaceText.FontWeight = $WeightFont
                        $SyncHash.TopFaceText.FontWeight = $WeightFont
                        $SyncHash.BottomFaceText.FontWeight = $WeightFont
                    }

                    if($Null -ne $StyleFont)
                    {
                        $SyncHash.FrontFaceText.FontStyle = $StyleFont
                        $SyncHash.BackFaceText.FontStyle = $StyleFont
                        $SyncHash.RightFaceText.FontStyle = $StyleFont
                        $SyncHash.LeftFaceText.FontStyle = $StyleFont
                        $SyncHash.TopFaceText.FontStyle = $StyleFont
                        $SyncHash.BottomFaceText.FontStyle = $StyleFont
                    }
        
                    try
                    {
                        $SyncHash.FrontFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[0] -Leaf)
                        $SyncHash.BackFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[1] -Leaf)
                        $SyncHash.RightFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[2] -Leaf)
                        $SyncHash.LeftFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[3] -Leaf)
                        $SyncHash.TopFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[4] -Leaf)
                        $SyncHash.BottomFaceText.Text = (Split-Path -Path $SyncHash.VideoFiles[5] -Leaf)
                    }
                    catch
                    {
                        # Silently ignore the null videoUri Error. No output will be displayed.
                    }
                }
                "Custom"
                {
                    $SyncHash.FrontFaceText.Text = $TextBox.Text
                    $SyncHash.BackFaceText.Text = $TextBox.Text
                    $SyncHash.RightFaceText.Text = $TextBox.Text
                    $SyncHash.LeftFaceText.Text = $TextBox.Text
                    $SyncHash.TopFaceText.Text = $TextBox.Text
                    $SyncHash.BottomFaceText.Text = $TextBox.Text

                    # Convert System.Drawing.Color to System.Windows.Media.Color
                    $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)

                    # Create a SolidColorBrush from the System.Windows.Media.Color
                    $SyncHash.FrontFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.BackFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.RightFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.LeftFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.TopFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                    $SyncHash.BottomFaceText.Foreground = New-Object System.Windows.Media.SolidColorBrush($mediaColor)

                    $SyncHash.SelectedFontSize = $SyncHash.NumericUpDown.Value

                    $SyncHash.FrontFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.BackFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.RightFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.LeftFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.TopFaceText.FontSize = $SyncHash.SelectedFontSize
                    $SyncHash.BottomFaceText.FontSize = $SyncHash.SelectedFontSize

                    $SyncHash.FrontFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.BackFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.RightFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.LeftFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.TopFaceText.FontFamily = $SyncHash.SelectedFont
                    $SyncHash.BottomFaceText.FontFamily = $SyncHash.SelectedFont

                    if($SyncHash.ItalicCheckbox.Checked -and $BoldCheckbox.Checked)
                    {
                        $StyleFont = [System.Windows.FontStyles]::Italic
                        $WeightFont = [System.Windows.FontWeights]::Bold
                    }
                    elseif(($SyncHash.ItalicCheckbox.Checked))
                    {
                        $StyleFont = [System.Windows.FontStyles]::Italic
                        $WeightFont = $Null
                    }
                    elseif(($SyncHash.BoldCheckbox.Checked))
                    {
                        $StyleFont = $Null
                        $WeightFont = [System.Windows.FontWeights]::Bold
                    }
                    else
                    {
                        $StyleFont = [System.Windows.FontStyles]::Normal
                        $WeightFont = $Null
                    }

                    if($Null -ne $WeightFont)
                    {
                        $SyncHash.FrontFaceText.FontWeight = $WeightFont
                        $SyncHash.BackFaceText.FontWeight = $WeightFont
                        $SyncHash.RightFaceText.FontWeight = $WeightFont
                        $SyncHash.LeftFaceText.FontWeight = $WeightFont
                        $SyncHash.TopFaceText.FontWeight = $WeightFont
                        $SyncHash.BottomFaceText.FontWeight = $WeightFont
                    }

                    if($Null -ne $StyleFont)
                    {
                        $SyncHash.FrontFaceText.FontStyle = $StyleFont
                        $SyncHash.BackFaceText.FontStyle = $StyleFont
                        $SyncHash.RightFaceText.FontStyle = $StyleFont
                        $SyncHash.LeftFaceText.FontStyle = $StyleFont
                        $SyncHash.TopFaceText.FontStyle = $StyleFont
                        $SyncHash.BottomFaceText.FontStyle = $StyleFont
                    }
                }
                Default
                {
                    $SyncHash.FrontFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.BackFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.RightFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.LeftFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.TopFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                    $SyncHash.BottomFaceText.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }

            $Window.Add_Closed({
                    # ---------------- Cleanup -----------------
                    foreach ($faceIndex in $SyncHash.FaceData.Keys) {
                        $face = $SyncHash.FaceData[$faceIndex]
                        if ($face.Timer) { $face.Timer.Stop() }
                        if ($face.Process -and -not $face.Process.HasExited) { $face.Process.Kill() }
                        if ($face.Process) { $face.Process.Dispose() }
                    }
                    Get-Process -Name "ffmpeg" -ErrorAction SilentlyContinue | Stop-Process -Force
                })

            $Window.Add_KeyDown({
                    param($Sender, $e)
                    
                    switch ($e.Key)
                    {
                        "A"
                        {
                            for ($i = 1; $i -le 9; $i++)
                            {
                                $X1 = Get-Random -Minimum 0 -Maximum 2
                                $X2 = Get-Random -Minimum 0 -Maximum 2
                                $X3 = Get-Random -Minimum 0 -Maximum 2

                                $Y1 = Get-Random -Minimum 0 -Maximum 2
                                $Y2 = Get-Random -Minimum 0 -Maximum 2
                                $Y3 = Get-Random -Minimum 0 -Maximum 2

                                $Z1 = Get-Random -Minimum 0 -Maximum 2
                                $Z2 = Get-Random -Minimum 0 -Maximum 2
                                $Z3 = Get-Random -Minimum 0 -Maximum 2
    
                                $SyncHash.AxisAngleX.Axis = "$X1,$X2,$X3"
                                $SyncHash.AxisAngleY.Axis = "$Y1,$Y2,$Y3"
                                $SyncHash.AxisAngleZ.Axis = "$Z1,$Z2,$Z3"

                                $SyncHash.AxisAngleX.Axis = "$X1,$X2,$X3"
                                $SyncHash.AxisAngleY.Axis = "$Y1,$Y2,$Y3"
                                $SyncHash.AxisAngleZ.Axis = "$Z1,$Z2,$Z3"

                                $SyncHash.AxisAngleX.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                                $SyncHash.AxisAngleY.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                                $SyncHash.AxisAngleZ.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                            } 
                        }
                        "F1"
                        {
                            $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                            $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)

                            # Get controls from the popup window
                            $OkButton = $PopupWindow.FindName("OKButton")

                            # Define OK button click event for the popup
                            $OkButton.Add_Click({
                                    # Closes the popup window
                                    $PopupWindow.Close()
                                })

                            # Show the popup window as a modal dialog
                            $PopupWindow.ShowDialog() | Out-Null
                        }
                        "Left"
                        {
                            $CurrentDuration = $(($animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds) / 3)
                            $NewDuration = $CurrentDuration * 2
                            if($NewDuration -le 0){ $NewDuration = 1 }

                            $animX.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $animY.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $animZ.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))

                            $animX.From = $SyncHash.AxisAngleX.Angle
                            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                            $animY.From = $SyncHash.AxisAngleY.Angle
                            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                            $animZ.From = $SyncHash.AxisAngleZ.Angle
                            $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)
                        }
                        "Right"
                        {
                            $CurrentDuration = $(($animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds) / 3)
                            $NewDuration = $CurrentDuration / 2
                            if($NewDuration -le 0){ $NewDuration = 1 }

                            $animX.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $animY.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $animZ.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))

                            $animX.From = $SyncHash.AxisAngleX.Angle
                            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                            $animY.From = $SyncHash.AxisAngleY.Angle
                            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                            $animZ.From = $SyncHash.AxisAngleZ.Angle
                            $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)
                        }
                        "Escape"{ $SelectFolderForm.Dispose(); $Window.Close() }
                        "p"
                        {
                            if($SyncHash.Paused -eq $False)
                            {
                                # Get the current angle from the animation's current value
                                $Current_angleX = $SyncHash.AxisAngleX.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                                $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                                $Current_angleZ = $SyncHash.AxisAngleZ.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)

                                # Use the current angle as the new base value
                                $SyncHash.AxisAngleX.Angle = $Current_angleX
                                $SyncHash.AxisAngleY.Angle = $Current_angleY
                                $SyncHash.AxisAngleZ.Angle = $Current_angleZ

                                # Clear the ongoing animation by calling BeginAnimation with null
                                $SyncHash.AxisAngleX.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                                $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                                $SyncHash.AxisAngleZ.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
    
                                $SyncHash.PauseButton.Content = "Resume"
                                $SyncHash.Paused = $True
                            }
                            else
                            {
                                $animX.From = $SyncHash.AxisAngleX.Angle
                                $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                                $animY.From = $SyncHash.AxisAngleY.Angle
                                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                                $animZ.From = $SyncHash.AxisAngleZ.Angle
                                $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)

                                $SyncHash.PauseButton.Content = "Pause"
                                $SyncHash.Paused = $False
                            }
                        }
                        "R"  { $Window.Close(); $SelectFolderForm.Show() }
                        "h"
                        {
                            if ($SyncHash.ControlsHidden -eq $True)
                            {
                                $SyncHash.ControlsHidden = $False
                                $SyncHash.ReDoButton.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.CloseButton.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.SlowDown.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.SpeedUp.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.PauseButton.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.HideControls.Visibility = [System.Windows.Visibility]::Visible
                                $SyncHash.RandomAxis.Visibility = [System.Windows.Visibility]::Visible
                            }
                            else
                            {
                                $SyncHash.ControlsHidden = $True
                                $SyncHash.ReDoButton.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.CloseButton.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.SlowDown.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.SpeedUp.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.PauseButton.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.HideControls.Visibility = [System.Windows.Visibility]::Hidden
                                $SyncHash.RandomAxis.Visibility = [System.Windows.Visibility]::Hidden
                            }
                        }
                    }
                })

            $SyncHash.CloseButton.Add_Click({ $Window.Close() })

            $SyncHash.SlowDown.Add_Click({
                    $CurrentDuration = $(($animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds) / 3)
                    $NewDuration = $CurrentDuration * 2
                    if($NewDuration -le 0){ $NewDuration = 1 }

                    $animX.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $animY.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $animZ.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))

                    $animX.From = $SyncHash.AxisAngleX.Angle
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                    $animY.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                    $animZ.From = $SyncHash.AxisAngleZ.Angle
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)
                })

            $SyncHash.SpeedUp.Add_Click({
                    $CurrentDuration = $(($animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds + $animX.Duration.TimeSpan.TotalSeconds) / 3)
                    $NewDuration = $CurrentDuration / 2
                    if($NewDuration -le 0){ $NewDuration = 1 }

                    $animX.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $animY.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $animZ.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))

                    $animX.From = $SyncHash.AxisAngleX.Angle
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                    $animY.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                    $animZ.From = $SyncHash.AxisAngleZ.Angle
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)
                })

            # Define the event handler script block
            $SyncHash.PauseButton.Add_Click({
                    if($SyncHash.Paused -eq $False)
                    {
                        # Get the current angle from the animation's current value
                        $Current_angleX = $SyncHash.AxisAngleX.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                        $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                        $Current_angleZ = $SyncHash.AxisAngleZ.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)

                        # Use the current angle as the new base value
                        $SyncHash.AxisAngleX.Angle = $Current_angleX
                        $SyncHash.AxisAngleY.Angle = $Current_angleY
                        $SyncHash.AxisAngleZ.Angle = $Current_angleZ

                        # Clear the ongoing animation by calling BeginAnimation with null
                        $SyncHash.AxisAngleX.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                        $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                        $SyncHash.AxisAngleZ.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
    
                        $SyncHash.PauseButton.Content = "Resume"
                        $SyncHash.Paused = $True
                    }
                    else
                    {
                        $animX.From = $SyncHash.AxisAngleX.Angle
                        $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                        $animY.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                        $animZ.From = $SyncHash.AxisAngleZ.Angle
                        $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)

                        $SyncHash.PauseButton.Content = "Pause"
                        $SyncHash.Paused = $False
                    }
                })

            $SyncHash.ReDoButton.Add_Click({
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                    $Window.Close()
                    $SelectFolderForm.Show()
                })

            $SyncHash.RandomAxis.Add_Click({
                    for ($i = 1; $i -le 9; $i++)
                    {
                        $X1 = Get-Random -Minimum 0 -Maximum 2
                        $X2 = Get-Random -Minimum 0 -Maximum 2
                        $X3 = Get-Random -Minimum 0 -Maximum 2

                        $Y1 = Get-Random -Minimum 0 -Maximum 2
                        $Y2 = Get-Random -Minimum 0 -Maximum 2
                        $Y3 = Get-Random -Minimum 0 -Maximum 2

                        $Z1 = Get-Random -Minimum 0 -Maximum 2
                        $Z2 = Get-Random -Minimum 0 -Maximum 2
                        $Z3 = Get-Random -Minimum 0 -Maximum 2
    
                        $SyncHash.AxisAngleX.Axis = "$X1,$X2,$X3"
                        $SyncHash.AxisAngleY.Axis = "$Y1,$Y2,$Y3"
                        $SyncHash.AxisAngleZ.Axis = "$Z1,$Z2,$Z3"

                        $SyncHash.AxisAngleX.Axis = "$X1,$X2,$X3"
                        $SyncHash.AxisAngleY.Axis = "$Y1,$Y2,$Y3"
                        $SyncHash.AxisAngleZ.Axis = "$Z1,$Z2,$Z3"

                        $SyncHash.AxisAngleX.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                        $SyncHash.AxisAngleY.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                        $SyncHash.AxisAngleZ.Angle = "$(Get-Random -Minimum -360 -Maximum 360)"
                    } 
                })

            $SyncHash.HideControls.Add_Click({
                    if ($SyncHash.ControlsHidden -eq $True)
                    {
                        $SyncHash.ControlsHidden = $False
                        $SyncHash.ReDoButton.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.CloseButton.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.SlowDown.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.SpeedUp.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.PauseButton.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.HideControls.Visibility = [System.Windows.Visibility]::Visible
                        $SyncHash.RandomAxis.Visibility = [System.Windows.Visibility]::Visible
                    }
                    else
                    {
                        $SyncHash.ControlsHidden = $True
                        $SyncHash.ReDoButton.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.CloseButton.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.SlowDown.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.SpeedUp.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.PauseButton.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.HideControls.Visibility = [System.Windows.Visibility]::Hidden
                        $SyncHash.RandomAxis.Visibility = [System.Windows.Visibility]::Hidden
                    }
                })

            # Define the MouseDown event handler
            $SyncHash.MainGrid.Add_MouseDown({
                    param($Sender, $e)
                    if($SyncHash.Paused -eq $False)
                    {
                        # Get the current angle from the animation's current value
                        $Current_angleX = $SyncHash.AxisAngleX.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                        $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                        $Current_angleZ = $SyncHash.AxisAngleZ.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)

                        # Use the current angle as the new base value
                        $SyncHash.AxisAngleX.Angle = $Current_angleX
                        $SyncHash.AxisAngleY.Angle = $Current_angleY
                        $SyncHash.AxisAngleZ.Angle = $Current_angleZ

                        # Clear the ongoing animation by calling BeginAnimation with null
                        $SyncHash.AxisAngleX.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                        $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                        $SyncHash.AxisAngleZ.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
    
                        $SyncHash.PauseButton.Content = "Resume"
                        $SyncHash.Paused = $True
                    }
                    else
                    {
                        $animX.From = $SyncHash.AxisAngleX.Angle
                        $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animX)

                        $animY.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animY)

                        $animZ.From = $SyncHash.AxisAngleZ.Angle
                        $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $animZ)

                        $SyncHash.PauseButton.Content = "Pause"
                        $SyncHash.Paused = $False
                    }
                })


            # ---------------- Show -----------------
            $SelectFolderForm.Hide()
            
            # Signal the loading form's runspace that the cube is ready.
            $SyncHash.CubeReady = $true
            # Give the loading form a moment to close before showing the main window.
            Start-Sleep -Milliseconds 200

            $null = $window.ShowDialog()
        }

    })
$SelectFolderForm.ShowDialog() | Out-Null
