<#
.SYNOPSIS
    Displays selected images and videos on the faces of a rotating 3D cube using
    MediaElement.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them onto the faces
    of a rotating 3D cube in a WPF window.

    This version uses the built-in Windows MediaElement for video playback. As a result, video
    format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI).
    For broader format support, use the FFmpeg version of this script.

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays on the cube
    faces.

.EXAMPLE
    PS C:\> .\Show-RotatingImageVideoCubeMediaElement.ps1

    Launches the file selection GUI. After selecting at least 6 files and clicking "Play", the
    script will launch the 3D cube window.

.NOTES
    Name:           Show-RotatingImageVideoCubeMediaElement.ps1
    Version:        1.0.0, 10/18/2025
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

$ExternalButtonName = "Rotating Images/Videos Cube `n No Ffmpeg"
$ScriptDescription = "Loops through and displays 6 or more selected images or videos on the faces of a rotating 3D cube. Uses the built-in Windows MediaElement, which may have more limited video format support."
$RequiredExecutables = @() # No external executables needed

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

        $VideoExtensions = "*.wmv", "*.asf", "*.mpg", "*.mpeg", "*.mp2", "*.mpe", "*.mpv", "*.avi", "*.mp4", "*.mov", "*.webm", "*.mkv"

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

$HelpLabel = New-Object System.Windows.Forms.Label
$HelpLabel.Text = "F1 - Help"
$HelpLabel.AutoSize = $True
$HelpLabel.Location = New-Object System.Drawing.Point(700, 0)
$HelpLabel.Size = New-Object System.Drawing.Size(150, 20)
$SelectFolderForm.Controls.Add($HelpLabel)

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


            # 1. Load the XAML from a here-string
            [xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Video Cube" Height="600" Width="800"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    SizeToContent="Manual"
    WindowState="Normal"
    AllowsTransparency="True"
    Background="Transparent">
    <Window.Resources>
        <!-- Define VisualBrushes as resources to be shared -->
        <VisualBrush x:Key="FrontFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerFront" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="FrontFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="FrontErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
        <VisualBrush x:Key="BackFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerBack" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="BackFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="BackErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
        <VisualBrush x:Key="RightFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerRight" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="RightFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="RightErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
        <VisualBrush x:Key="LeftFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerLeft" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="LeftFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="LeftErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
        <VisualBrush x:Key="TopFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerTop" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="TopFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="TopErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
        <VisualBrush x:Key="BottomFaceVisualBrush" Stretch="UniformToFill" Viewbox="0,0,1,1">
            <VisualBrush.Visual>
                <Grid Background="LightBlue" Height="600" Width="800">
                    <MediaElement x:Name="videoPlayerBottom" Stretch="Fill" LoadedBehavior="Manual"/>
                    <TextBlock x:Name="BottomFaceText" Text="" FontSize="10" Foreground="Black" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="10" IsHitTestVisible="False" />
                    <TextBlock Name="BottomErrorTextBlock" Foreground="Red" VerticalAlignment="Center" HorizontalAlignment="Center" Visibility="Collapsed" Width="400" Height="400"/>
                </Grid>
            </VisualBrush.Visual>
        </VisualBrush>
    </Window.Resources>
  <Grid x:Name="MainGrid">
    <Viewport3D Name="mainViewport" ClipToBounds="True">
      <Viewport3D.Camera>
        <PerspectiveCamera Position="0, 0, 4" LookDirection="0, 0, -1" UpDirection="0, 1, 0" FieldOfView="60" />
      </Viewport3D.Camera>
            <ModelVisual3D>
                <ModelVisual3D.Content>
                    <Model3DGroup x:Name="CubeModel">
                        <!-- Lights -->
                        <AmbientLight Color="Gray" />
                        <DirectionalLight Color="#ffffffff" Direction="-1, -1, -3" />
                        <DirectionalLight Color="#ffffffff" Direction="1, 1, 3" />
                        <DirectionalLight Color="#ffffffff" Direction="-3, -1, -1" />
                        <DirectionalLight Color="#ffffffff" Direction="3, 1, 1" />
                        <!-- The cube faces are defined as MeshGeometry3D with dynamically assigned materials -->
                        <!-- You can have more than 6 faces, but this example uses 6 for a standard cube. -->
                        
                        <!-- Front Face -->
                        <GeometryModel3D x:Name="FrontFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="-1,-1,1  1,-1,1  1,1,1  -1,1,1" 
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="0,0,1 0,0,1 0,0,1 0,0,1" 
                                                TextureCoordinates="0,1 1,1 1,0 0,0" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="FrontScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource FrontFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Back Face -->
                        <GeometryModel3D x:Name="BackFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="-1,-1,-1  -1,1,-1  1,1,-1  1,-1,-1"
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="0,0,-1 0,0,-1 0,0,-1 0,0,-1" 
                                                TextureCoordinates="1,1 1,0 0,0 0,1" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="BackScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource BackFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Right Face -->
                        <GeometryModel3D x:Name="RightFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="1,-1,1  1,-1,-1  1,1,-1  1,1,1" 
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="1,0,0 1,0,0 1,0,0 1,0,0" 
                                                TextureCoordinates="0,1 1,1 1,0 0,0" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="RightScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource RightFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Left Face -->
                        <GeometryModel3D x:Name="LeftFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="-1,-1,-1  -1,-1,1  -1,1,1  -1,1,-1" 
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="-1,0,0 -1,0,0 -1,0,0 -1,0,0" 
                                                TextureCoordinates="0,1 1,1 1,0 0,0" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="LeftScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource LeftFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Top Face -->
                        <GeometryModel3D x:Name="TopFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="-1,1,1  1,1,1  1,1,-1  -1,1,-1" 
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="0,1,0 0,1,0 0,1,0 0,1,0" 
                                                TextureCoordinates="0,1 1,1 1,0 0,0" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="TopScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource TopFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Bottom Face -->
                        <GeometryModel3D x:Name="BottomFaceModel">
                            <GeometryModel3D.Geometry>
                                <MeshGeometry3D Positions="-1,-1,-1  1,-1,-1  1,-1,1  -1,-1,1" 
                                                TriangleIndices="0,1,2 0,2,3" 
                                                Normals="0,-1,0 0,-1,0 0,-1,0 0,-1,0" 
                                                TextureCoordinates="0,1 1,1 1,0 0,0" />
                            </GeometryModel3D.Geometry>
                            <GeometryModel3D.Transform>
                                <ScaleTransform3D x:Name="BottomScaleTransform" ScaleX=".7" ScaleY=".7" ScaleZ=".7" />
                            </GeometryModel3D.Transform>
                                <GeometryModel3D.Material>
                                    <DiffuseMaterial Brush="{StaticResource BottomFaceVisualBrush}" />
                                </GeometryModel3D.Material>
                        </GeometryModel3D>
                        
                        <!-- Rotation Transform -->
                        <Model3DGroup.Transform>
                            <Transform3DGroup x:Name="CubeTransformGroup">
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
                        </Model3DGroup.Transform>
                    </Model3DGroup>
                </ModelVisual3D.Content>
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
            
            # 2. Load XAML
            $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
            $Window = [Windows.Markup.XamlReader]::Load($Reader)

            # Define face names to enable looping for MediaPlayers
            $faceNames = @("Front", "Back", "Right", "Left", "Top", "Bottom")

            $SyncHash.Window = $Window

            # 3. Get controls from XAML
            $SyncHash.CubeModel = $Window.FindName("CubeModel")
            $SyncHash.MediaPlayers = @(0..5 | ForEach-Object { $Window.FindName("videoPlayer$($faceNames[$_])") })
            $SyncHash.AxisAngleX = $Window.FindName("AxisAngleX")
            $SyncHash.AxisAngleY = $Window.FindName("AxisAngleY")
            $SyncHash.AxisAngleZ = $Window.FindName("AxisAngleZ")
            $SyncHash.CloseButton = $Window.FindName("CloseButton")
            $SyncHash.SlowDown = $Window.FindName("SlowDown")
            $SyncHash.SpeedUp = $Window.FindName("SpeedUp")
            $SyncHash.ReDoButton = $Window.FindName("ReDoButton")
            $SyncHash.HideControls = $Window.FindName("HideControls")
            $SyncHash.PauseButton = $Window.FindName("PauseButton")
            $SyncHash.RandomAxis = $Window.FindName("RandomAxis")
            $SyncHash.FrontFaceModel = $Window.FindName("FrontFaceModel")
            $SyncHash.BackFaceModel = $Window.FindName("BackFaceModel")
            $SyncHash.RightFaceModel = $Window.FindName("RightFaceModel")
            $SyncHash.LeftFaceModel = $Window.FindName("LeftFaceModel")
            $SyncHash.TopFaceModel = $Window.FindName("TopFaceModel")
            $SyncHash.BottomFaceModel = $Window.FindName("BottomFaceModel")
            $SyncHash.FrontFaceText = $Window.FindName("FrontFaceText")
            $SyncHash.BackFaceText = $Window.FindName("BackFaceText")
            $SyncHash.RightFaceText = $Window.FindName("RightFaceText")
            $SyncHash.LeftFaceText = $Window.FindName("LeftFaceText")
            $SyncHash.TopFaceText = $Window.FindName("TopFaceText")
            $SyncHash.BottomFaceText = $Window.FindName("BottomFaceText")
            $SyncHash.FrontErrorTextBlock = $Window.FindName("FrontErrorTextBlock")
            $SyncHash.BackErrorTextBlock = $Window.FindName("BackErrorTextBlock")
            $SyncHash.RightErrorTextBlock = $Window.FindName("RightErrorTextBlock")
            $SyncHash.LeftErrorTextBlock = $Window.FindName("LeftErrorTextBlock")
            $SyncHash.TopErrorTextBlock = $Window.FindName("TopErrorTextBlock")
            $SyncHash.BottomErrorTextBlock = $Window.FindName("BottomErrorTextBlock")
            $SyncHash.MainGrid = $Window.FindName("MainGrid")

            # If the checkbox is checked, replace DiffuseMaterial with EmissiveMaterial
            if ($TransparentFacesCheckbox.Checked)
            {
                $faceNames = @("Front", "Back", "Right", "Left", "Top", "Bottom")
                foreach ($faceName in $faceNames)
                {
                    $faceModel = $SyncHash."$($faceName)FaceModel"
                    # Find the VisualBrush resource from the window's resources
                    $visualBrush = $Window.Resources["$($faceName)FaceVisualBrush"]
                    
                    if ($faceModel -and $visualBrush)
                    {
                        # Create a new EmissiveMaterial
                        $emissiveMaterial = New-Object System.Windows.Media.Media3D.EmissiveMaterial
                        
                        # Set the material's color to White. This is crucial for the VisualBrush to render correctly.
                        $emissiveMaterial.Color = [System.Windows.Media.Colors]::White

                        # Assign the VisualBrush resource to the new material
                        $emissiveMaterial.Brush = $visualBrush
                        
                        # Replace the material on the face model
                        $faceModel.Material = $emissiveMaterial
                    }
                }
            }

            # Initialize a state tracker for each player
            $SyncHash.PlayerState = [hashtable]::Synchronized(@{})
            foreach ($player in $SyncHash.MediaPlayers) {
                $SyncHash.PlayerState[$player.Name] = @{
                    IsImage = $false
                    ImageTimer = $null # To hold the dedicated timer for an image
                    RecoveryTimer = $null # To hold a timer for recovering from an error
                    PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch
                    IsFailed = $false
                }
            }

            # Define image extensions to identify image files
            $ImageExtensions = @(".bmp", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".gif", ".wmp", ".ico")
            $SyncHash.ImageExtensions = $ImageExtensions

            # Centralized function to handle any media failure
            $SyncHash.HandleMediaFailure = {
                param($ErrorElement, [string]$Reason = "Unknown Error")

                $SyncHash.Window.Dispatcher.Invoke([action]{
                    $playerState = $SyncHash.PlayerState[$ErrorElement.Name]
                    if ($playerState.IsFailed) { return } # Prevent this from running twice for the same error
                    $playerState.IsFailed = $true

                    $fileName = if ($ErrorElement.Source) { $ErrorElement.Source.Segments[-1] } else { "an unknown media file" }
                    $errorText = "Error: $($fileName)`n$Reason"

                    # Display the error on the correct face
                    switch ($ErrorElement.Name) {
                        "videoPlayerFront"  { $SyncHash.FrontErrorTextBlock.Text = $errorText; $SyncHash.FrontErrorTextBlock.Visibility = "Visible" }
                        "videoPlayerBack"   { $SyncHash.BackErrorTextBlock.Text = $errorText; $SyncHash.BackErrorTextBlock.Visibility = "Visible" }
                        "videoPlayerRight"  { $SyncHash.RightErrorTextBlock.Text = $errorText; $SyncHash.RightErrorTextBlock.Visibility = "Visible" }
                        "videoPlayerLeft"   { $SyncHash.LeftErrorTextBlock.Text = $errorText; $SyncHash.LeftErrorTextBlock.Visibility = "Visible" }
                        "videoPlayerTop"    { $SyncHash.TopErrorTextBlock.Text = $errorText; $SyncHash.TopErrorTextBlock.Visibility = "Visible" }
                        "videoPlayerBottom" { $SyncHash.BottomErrorTextBlock.Text = $errorText; $SyncHash.BottomErrorTextBlock.Visibility = "Visible" }
                    }
                    $ErrorElement.Visibility = "Collapsed"
                    $ErrorElement.Stop() # Stop the failed media

                    if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }

                    $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $recoveryTimer.Interval = [TimeSpan]::FromSeconds(10)
                    $recoveryTimer.Tag = $ErrorElement # Pass the failed element

                    $recoveryTick = {
                        $timer = $args[0]; $failedElement = $timer.Tag; $timer.Stop()
                        # Reset the failed flag before calling the next handler to allow it to proceed.
                        $SyncHash.PlayerState[$failedElement.Name].IsFailed = $false
                        # Now, try to load the next media.
                        & $SyncHash.MediaEndedHandler -Sender $failedElement -e $null -IsRecovery
                    }
                    $recoveryTimer.Add_Tick($recoveryTick)
                    $playerState.RecoveryTimer = $recoveryTimer
                    $recoveryTimer.Start()
                })
            }


            # Handler for explicit media failures (e.g., file not found, network error)
            $MediaFailedHandler = {
                param($Sender, $EventArgs)
                $reason = if ($EventArgs.ErrorException) { $EventArgs.ErrorException.Message } else { "MediaFailed event fired." }
                & $SyncHash.HandleMediaFailure -ErrorElement $Sender -Reason $reason
            }

            # Handler for when media is successfully opened
            $MediaOpenedHandler = {
                param($Sender, $EventArgs)
                $playerState = $SyncHash.PlayerState[$Sender.Name]
                $playerState.IsFailed = $false # Reset failure flag on a new open
                $playerState.PlaybackStopwatch.Restart() # Start timing how long it plays

                # When a new media opens successfully, hide any previous error message on that face.
                $SyncHash.Window.Dispatcher.Invoke([action]{
                    $Sender.Visibility = "Visible"
                    switch ($Sender.Name) {
                        "videoPlayerFront"  { $SyncHash.FrontErrorTextBlock.Visibility = "Collapsed" }
                        "videoPlayerBack"   { $SyncHash.BackErrorTextBlock.Visibility = "Collapsed" }
                        "videoPlayerRight"  { $SyncHash.RightErrorTextBlock.Visibility = "Collapsed" }
                        "videoPlayerLeft"   { $SyncHash.LeftErrorTextBlock.Visibility = "Collapsed" }
                        "videoPlayerTop"    { $SyncHash.TopErrorTextBlock.Visibility = "Collapsed" }
                        "videoPlayerBottom" { $SyncHash.BottomErrorTextBlock.Visibility = "Collapsed" }
                    }
                })

                # If it's an image, it won't have a duration and won't fire MediaEnded naturally.
                # We'll start a dedicated timer to cycle it.
                if ($playerState.IsImage) {
                    # Explicitly pause the MediaElement to prevent it from "playing" the image
                    # and prematurely firing MediaEnded if it somehow detects a short duration.
                    $Sender.Pause() 

                    # If there are more files than faces, start a timer to cycle the image.
                    if ($SyncHash.SelectedFiles.Count -gt 6) {
                        # Stop any existing timer for this player to prevent conflicts
                        if ($playerState.ImageTimer) {
                            $playerState.ImageTimer.Stop()
                        }

                        # Create a new timer specific to this player state
                        $playerState.ImageTimer = New-Object System.Windows.Threading.DispatcherTimer
                        $playerState.ImageTimer.Interval = [TimeSpan]::FromSeconds(10)
                        $playerState.ImageTimer.Tag = $Sender # Pass the MediaElement to the handler
                        
                        $tickScriptBlock = {
                            # This runs on the UI thread
                            $timer = $args[0]
                            $mediaElement = $timer.Tag
                            $timer.Stop()
                            # Simulate MediaEnded for the image
                            & $SyncHash.MediaEndedHandler -Sender $mediaElement -e $null
                        }
                        $playerState.ImageTimer.Add_Tick($tickScriptBlock)
                        $playerState.ImageTimer.Start()
                    }
                }
                # A valid video will have a duration. If it doesn't, it's a silent failure.
                elseif (-not $Sender.NaturalDuration.HasTimeSpan) {
                    # This is a silent failure (e.g. bad codec). Trigger the failure handler.
                    & $SyncHash.HandleMediaFailure -ErrorElement $Sender -Reason "No duration found (silent failure)."
                }
                else {
                    # This is a video, so ensure any lingering image timer for this face is stopped.
                    if ($playerState.ImageTimer) {
                        $playerState.ImageTimer.Stop()
                    }
                }
            }

            # 6. Event handler for MediaEnded
            $MediaEndedHandler = {
                param(
                    $Sender, 
                    $e,
                    [switch]$IsRecovery
                )
                # The $Sender variable refers to the MediaElement that finished playing
                $FinishedElement = $Sender

                # Gracefully exit if the sender is null, preventing a crash.
                if (-not $FinishedElement -or -not $FinishedElement.Name) {
                    return
                }
                
                $playerState = $SyncHash.PlayerState[$FinishedElement.Name]
                if ($playerState.IsFailed) { return } # The recovery timer is in control.

                # This block is for normal media ended events, not for recovery calls.
                if (-not $IsRecovery) {
                    $playerState.PlaybackStopwatch.Stop()
                    $elapsedMilliseconds = $playerState.PlaybackStopwatch.Elapsed.TotalMilliseconds

                    # If media "ends" in under 2 seconds, it's a failure (e.g., bad codec, renamed DLL).
                    # We also check if it's an image, as images are paused and would have an elapsed time of nearly 0.
                    if (($elapsedMilliseconds -lt 2000) -and (-not $playerState.IsImage)) {
                        & $SyncHash.HandleMediaFailure -ErrorElement $FinishedElement -Reason "Playback failed or ended instantly."
                        return # IMPORTANT: Stop processing here. HandleMediaFailure is now in control.
                    }
                }
                
                # --- Normal playback completion or recovery ---
                if(($SyncHash.SelectedFiles.Count -le 6))
                {
                    # Restart the video or image
                    $FinishedElement.Position = [TimeSpan]::FromSeconds(0)
                    $FinishedElement.Play()
                }
                else
                {
                    if($SyncHash.CurrentIndex -ge $SyncHash.SelectedFiles.Count)
                    {
                        $SyncHash.CurrentIndex = 0
                    }

                    # Stop the current media playback before changing the source
                    $FinishedElement.Stop()

                    # Define the path to the new video file
                    $NewVideoPath = $SyncHash.VideoUris[$SyncHash.CurrentIndex]
        
                    # Create a new Uri object for the video
                    $NewUri = New-Object System.Uri($NewVideoPath)

                    # PRE-SET THE IsImage FLAG HERE TO AVOID RACE CONDITIONS
                    $playerState = $SyncHash.PlayerState[$FinishedElement.Name]
                    $extension = [System.IO.Path]::GetExtension($NewUri.LocalPath).ToLower()
                    $playerState.IsImage = ($SyncHash.ImageExtensions -contains $extension)
                    # Set the new Source for the MediaElement
                    $FinishedElement.Source = $NewUri
                    if( $FinishedElement.Name -match "Front" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.FrontFaceText.Text = $NewUri.Segments[-1] }

                    if( $FinishedElement.Name -match "Back" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.BackFaceText.Text = $NewUri.Segments[-1] }

                    if( $FinishedElement.Name -match "Right" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.RightFaceText.Text = $NewUri.Segments[-1] }
                                    
                    if( $FinishedElement.Name -match "Left" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.LeftFaceText.Text = $NewUri.Segments[-1] }
                                    
                    if( $FinishedElement.Name -match "Top" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.TopFaceText.Text = $NewUri.Segments[-1] }

                    if( $FinishedElement.Name -match "Bottom" -and $SyncHash.RbSelection -notmatch 'Hidden|Custom' )
                    { $SyncHash.BottomFaceText.Text = $NewUri.Segments[-1] }

                    # Play the new media (will be paused by MediaOpened if it's an image)
                    $FinishedElement.Play()

                    $SyncHash.CurrentIndex++                            
                }
            }

            # Store the handler in the synchronized hashtable to make it accessible from other scopes (like timers)
            $SyncHash.MediaEndedHandler = $MediaEndedHandler

            # Attach the event handlers to all six MediaElements
            foreach ($player in $SyncHash.MediaPlayers) {
                $player.Add_MediaFailed($MediaFailedHandler)
                $player.Add_MediaOpened($MediaOpenedHandler)
                $player.Add_MediaEnded($MediaEndedHandler)
            }

            # 4. Prepare videos
            
            # Function to stop the MediaElement and replace the material
            $ChangeMaterial = {
                param($FaceModel)
                
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

                # Assign the new material to the GeometryModel3D
                $FaceModel.Material = $DiffuseMaterial
            }

            $SyncHash.CurrentVideoIndex = 0
            $VideoUris = $SyncHash.VideoFiles | ForEach-Object { New-Object System.Uri($_, [System.UriKind]::Absolute) }
            $SyncHash.VideoUris = $VideoUris
            for ($i = 0; $i -le 5; $i++)
            {
                if ($i -lt $VideoUris.Count) {
                    # This is where the initial media is set for each face
                    $player = $SyncHash.MediaPlayers[$i]
                    $uri = $VideoUris[$i]
                    
                    # PRE-SET THE IsImage FLAG HERE TO AVOID RACE CONDITIONS
                    $playerState = $SyncHash.PlayerState[$player.Name]
                    $extension = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
                    $playerState.IsImage = ($SyncHash.ImageExtensions -contains $extension)

                    # Now set the source, which will trigger MediaOpened event later
                    $player.Source = $uri
                }
                else {
                    # This still works because FaceModel variables are still individual
                    $faceModelName = "$($faceNames[$i])FaceModel"
                    & $ChangeMaterial $SyncHash.$faceModelName
                }
            }

            if($RadioButton1.Checked){ $SyncHash.RbSelection = "Hidden" }

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
                        $SyncHash.FrontFaceText.Text = $VideoUris[0].Segments[-1]
                        $SyncHash.BackFaceText.Text = $VideoUris[1].Segments[-1]
                        $SyncHash.RightFaceText.Text = $VideoUris[2].Segments[-1]
                        $SyncHash.LeftFaceText.Text = $VideoUris[3].Segments[-1]
                        $SyncHash.TopFaceText.Text = $VideoUris[4].Segments[-1]
                        $SyncHash.BottomFaceText.Text = $VideoUris[5].Segments[-1]
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

            # 5. Create the rotation animation
            $RotateAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation
            $RotateAnimation.From = 0
            $RotateAnimation.To = 360
            $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds(30))
            $RotateAnimation.RepeatBehavior = "Forever"
            $RotateAnimation.IsCumulative = $True

            # 7. Add rotation and start video when the window is loaded
            $Window.Add_Loaded({
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)

                    # Start video playback for all players
                    foreach ($player in $SyncHash.MediaPlayers) {
                        $player.Play()
                    }
                })

            $Window.Add_Closed({
                    # Stop all timers to prevent them from firing after the window is closed
                    if ($SyncHash.Contains("MediaPlayers")) {
                        $playerNames = $SyncHash.MediaPlayers | ForEach-Object { $_.Name }
                        foreach ($name in $PlayerNames) {
                            if ($SyncHash.PlayerState.ContainsKey($name)) {
                                $playerState = $SyncHash.PlayerState[$name]
                                if ($playerState -and $playerState.ImageTimer) {
                                    $playerState.ImageTimer.Stop()
                                }
                                if ($playerState -and $playerState.RecoveryTimer) {
                                    $playerState.RecoveryTimer.Stop()
                                }
                            }
                        }

                        # Stop all media elements, remove event handlers, and release file locks
                        foreach ($player in $SyncHash.MediaPlayers) {
                            $player.Stop()
                            try {
                                $player.Remove_MediaEnded($SyncHash.MediaEndedHandler)
                                $player.Remove_MediaFailed($MediaFailedHandler)
                                $player.Remove_MediaOpened($MediaOpenedHandler)
                            } catch {
                                # Ignore errors if handlers were already removed or never attached
                            }
                            $player.Source = $null
                            $player.Close()
                        }
                    }
                    # Clear out state from the synchronized hashtable to ensure a clean start on next run
                    $SyncHash.Clear()
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
                            $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                            $NewDuration = $CurrentDuration * 2
                            if($NewDuration -le 0){ $NewDuration = 1 }
                            $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                            $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                            $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                            $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        }
                        "Right"
                        {
                            $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                            $NewDuration = $CurrentDuration / 2
                            $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                            $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                            $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                            $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                            $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                            $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
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
                                $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                                $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                                $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                                $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                                $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
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
                    $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                    $NewDuration = $CurrentDuration * 2
                    if($NewDuration -le 0){ $NewDuration = 1 }
                    $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                })

            $SyncHash.SpeedUp.Add_Click({
                    $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                    $NewDuration = $CurrentDuration / 2
                    $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                    $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                    $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                    $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
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
                        $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                        $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                        $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
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
                        $RotateAnimation.From = $SyncHash.AxisAngleX.Angle
                        $SyncHash.AxisAngleX.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        $RotateAnimation.From = $SyncHash.AxisAngleZ.Angle
                        $SyncHash.AxisAngleZ.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                        $SyncHash.PauseButton.Content = "Pause"
                        $SyncHash.Paused = $False
                    }
                })

            # 8. Show the window
            $SelectFolderForm.Hide()

            # Signal the loading form's runspace that the cube is ready.
            $SyncHash.CubeReady = $true
            # Give the loading form a moment to close before showing the main window.
            Start-Sleep -Milliseconds 200

            $Window.ShowDialog() | Out-Null
        }
    })
$SelectFolderForm.ShowDialog() | Out-Null
