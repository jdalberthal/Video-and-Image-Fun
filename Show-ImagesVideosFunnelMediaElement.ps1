<#
.SYNOPSIS
    Displays selected images and videos on the surface of a rotating 3D funnel
    using MediaElement.

.DESCRIPTION
    This script launches a GUI to select image and video files, then renders them one at a time
    onto the surface of a rotating 3D funnel (cone) in a WPF window.

    This version uses the built-in Windows MediaElement for video playback. As a result, video
    format support is limited to the codecs installed on the local system (e.g., MP4, WMV, AVI).

    The 3D view is interactive, with controls to pause the rotation, change the rotation axis and
    speed, and hide the UI for an unobstructed view. It also supports text overlays.

.EXAMPLE
    PS C:\> .\Show-ImagesVideosFunnelMediaElement.ps1

    Launches the file selection GUI. After selecting files and clicking "Play", the
    script will launch the 3D funnel window.

.NOTES
    Name:           Show-ImagesVideosFunnelMediaElement.ps1
    Version:        1.0.0, 11/22/2025
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

$ExternalButtonName = "Rotating Images/Videos Funnel `n No Ffmpeg"
$ScriptDescription = "Loops through and displays selected images or videos on the faces of a rotating 3D funnel. Uses the built-in Windows MediaElement, which may have more limited video format support."
$RequiredExecutables = @() # No external executables needed

$SyncHash = [hashtable]::Synchronized(@{}) # For passing data between runspaces
$SyncHash.ControlsHidden = $False
$SyncHash.Paused = $False
$SyncHash.RbSelection = ""
$SyncHash.RedoClicked = $false

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
$TextBox.TextAlign = 'Center'
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
                    <Run Text="Pause         :  P  : Pause Funnel Spinning"/><LineBreak/>
                    <Run Text="Redo          :  R  : Reselect videos"/><LineBreak/>
                    <Run Text="Random Axis   :  A  : Change Rotation Axis"/><LineBreak/>
                    <Run Text="Hide Controls :  H  : Hide Controls/Show Controls"/><LineBreak/>
                    <Run Text="Left Arrow    :  &#x2190;  : Slow Down Spinning"/><LineBreak/>
                    <Run Text="Right Arrow   :  &#x2192;  : Speed Up Spinning"/><LineBreak/><LineBreak/>
                    <Run Text="*Click Funnel to Pause"/><LineBreak/>
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

$HelpLabel = New-Object System.Windows.Forms.Label
$HelpLabel.Text = "F1 - Help"
$HelpLabel.AutoSize = $True
$HelpLabel.Location = New-Object System.Drawing.Point(700, 0)
$HelpLabel.Size = New-Object System.Drawing.Size(150, 20)
$SelectFolderForm.Controls.Add($HelpLabel)

# Event handler for the Process Selected Files button
$PlayButton.Add_Click({
    # This event handler will now just close the form.
    # The main script loop will then proceed to launch the 3D window.
    $SelectFolderForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $SelectFolderForm.Close()
})

while ($true) {
    # Show the selection form and wait for it to be closed.
    $dialogResult = $SelectFolderForm.ShowDialog()

    # If the form was closed by the 'X' button or another non-OK way, exit the script.
    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Selection form closed without pressing 'Play'. Exiting."
        break
    }

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
        # Continue the loop to show the selection form again
        continue
    }
    else
    {
        $SyncHash.RedoClicked = $false
        $panelCount = 8

        # --- Loading Form in a Separate Runspace ---
        $SyncHash.FunnelReady = $false

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

            while (-not $SyncHash.FunnelReady) {
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

        $SyncHash.CurrentIndex = 0
        
        # --- Spiraling Panel Mesh Generation Function (from FFmpeg version) ---
        function New-SpiralingPanelMesh {
            param(
                [double]$startRadius,
                [double]$endRadius,
                [double]$height,
                [double]$arcAngle, # The angular width of the panel in degrees
                [double]$twistAngle, # The total twist from top to bottom in degrees
                [int]$stacks = 50, # Vertical subdivisions
                [int]$slices = 2   # Horizontal subdivisions
            )

            $mesh = New-Object System.Windows.Media.Media3D.MeshGeometry3D
            $arcAngleRad = $arcAngle * [Math]::PI / 180.0
            $twistAngleRad = $twistAngle * [Math]::PI / 180.0

            for ($j = 0; $j -le $stacks; $j++) {
                $v = $j / $stacks # Vertical progress (0 to 1)

                # Interpolate radius, Y position, and twist for the current stack
                $v_eased = $v * $v # Use an ease-in quadratic function for a proper funnel curve
                $currentRadius = $startRadius - $v * ($startRadius - $endRadius)
                $currentY = ($height / 2) - $v_eased * $height
                $currentTwist = $v * $twistAngleRad

                for ($i = 0; $i -le $slices; $i++) {
                    $u = $i / $slices # Horizontal progress (0 to 1)

                    # Calculate the angle for this slice of the panel
                    $theta = $currentTwist + ($u * $arcAngleRad) - ($arcAngleRad / 2.0)

                    # Calculate vertex position
                    $x = $currentRadius * [Math]::Cos($theta)
                    $z = $currentRadius * [Math]::Sin($theta)

                    $mesh.Positions.Add([System.Windows.Media.Media3D.Point3D]::new($x, $currentY, $z))
                    $mesh.TextureCoordinates.Add([System.Windows.Point]::new($u, $v))
                }
            }

            # Create triangle indices
            for ($j = 0; $j -lt $stacks; $j++) {
                for ($i = 0; $i -lt $slices; $i++) {
                    $row1 = $j * ($slices + 1)
                    $row2 = ($j + 1) * ($slices + 1)

                    $mesh.TriangleIndices.Add($row1 + $i); $mesh.TriangleIndices.Add($row2 + $i); $mesh.TriangleIndices.Add($row1 + $i + 1)
                    $mesh.TriangleIndices.Add($row1 + $i + 1); $mesh.TriangleIndices.Add($row2 + $i); $mesh.TriangleIndices.Add($row2 + $i + 1)
                }
            }

            $mesh.Freeze()
            return $mesh
        }

        # 1. Load the XAML from a here-string
        [xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    Title="Video Funnel" Height="600" Width="800"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    SizeToContent="Manual"
    WindowState="Normal"
    AllowsTransparency="True"
    Background="Transparent">    
  <Grid x:Name="MainGrid">
    <Viewport3D Name="mainViewport" ClipToBounds="True">
      <Viewport3D.Camera> 
        <PerspectiveCamera Position="0,0,15" LookDirection="0,0,-1" UpDirection="0,1,0" FieldOfView="70"/>
      </Viewport3D.Camera>
            <ModelVisual3D x:Name="ObjectContainer">
                <ModelVisual3D.Content>
                    <Model3DGroup>
                        <!-- Lights -->
                        <AmbientLight Color="Gray" />
                        <DirectionalLight Color="#FFFFFF" Direction="-1,-1,-2"/>
                        <DirectionalLight Color="#FFFFFF" Direction="1,1,2"/>
                    </Model3DGroup>
                </ModelVisual3D.Content>
                <ModelVisual3D.Transform>
                    <Transform3DGroup>
                        <!-- SPIN: This animated rotation happens first, spinning the object on its own Y-axis. -->
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D x:Name="AxisAngleY" Axis="0,1,0" Angle="0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <!-- TILT: This static rotation happens second, tilting the spinning object forward. -->
                        <RotateTransform3D>
                            <RotateTransform3D.Rotation>
                                <AxisAngleRotation3D Angle="40" Axis="1,0,0" />
                            </RotateTransform3D.Rotation>
                        </RotateTransform3D>
                        <!-- MOVE: This static translation happens last, moving the tilted, spinning object down. -->
                        <TranslateTransform3D OffsetY="-2.0" />
                    </Transform3DGroup>
                </ModelVisual3D.Transform>
            </ModelVisual3D>

    </Viewport3D>
    <StackPanel  Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
            <Button Name="PauseButton" Content="Pause" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="15,5" Grid.Column="0" Grid.Row="0" Width="90" />
            <Button Name="SlowDown" Content="&#x2190;" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="15,5" Grid.Column="0" Grid.Row="0" />
            <Button Name="SpeedUp" Content="&#x2192;" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="15,5" Grid.Column="0" Grid.Row="0" />
            <Button Name="ReDoButton" Content="Redo" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="15,5" Grid.Column="0" Grid.Row="0" />
            <Button Name="HideControls" Content="Hide Controls" HorizontalAlignment="Right" VerticalAlignment="Top" Padding="15,5" Grid.Column="0" Grid.Row="0" />
            <Button Content="X" HorizontalAlignment="Right" VerticalAlignment="Top" FontSize="16" Width="28" Height="28" Name="CloseButton" Grid.Column="0" Grid.Row="0" />
        </StackPanel>
    </Grid>
</Window>
"@
        $Xaml.Window.Width = "$WorkingAreaWidth"; $Xaml.Window.Height = "$WorkingAreaHeight"
        
        # 2. Load XAML
        $Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
        $Window = [Windows.Markup.XamlReader]::Load($Reader)

        $SyncHash.Window = $Window
        $objectContainer = $Window.FindName("ObjectContainer")

        # 3. Get controls from XAML
        $SyncHash.AxisAngleY = $Window.FindName("AxisAngleY") 
        $SyncHash.CloseButton = $Window.FindName("CloseButton") 
        $SyncHash.SlowDown = $Window.FindName("SlowDown") 
        $SyncHash.SpeedUp = $Window.FindName("SpeedUp") 
        $SyncHash.ReDoButton = $Window.FindName("ReDoButton") 
        $SyncHash.HideControls = $Window.FindName("HideControls") 
        $SyncHash.PauseButton = $Window.FindName("PauseButton") 
        $SyncHash.MainGrid = $Window.FindName("MainGrid") 

        # Initialize arrays to hold the dynamically created controls for each panel
        $SyncHash.MediaPlayers = [System.Collections.ArrayList]::new()
        $SyncHash.MediaPlayersBack = [System.Collections.ArrayList]::new()
        $SyncHash.OverlayTextBlocks = [System.Collections.ArrayList]::new()
        $SyncHash.OverlayTextBlocksBack = [System.Collections.ArrayList]::new()
        $SyncHash.MediaHostGrids = [System.Collections.ArrayList]::new()
        $SyncHash.MediaHostGridsBack = [System.Collections.ArrayList]::new()

        # Define the mesh for one spiraling panel
        $panelMesh = New-SpiralingPanelMesh -startRadius 5.0 -endRadius 1.0 -height 10.0 -arcAngle (360.0 / $panelCount) -twistAngle 90 -stacks 50 -slices 2

        for ($i = 0; $i -lt $panelCount; $i++) {
            # --- Create 2D content host for this panel ---
            $mediaHostGridFront = New-Object System.Windows.Controls.Grid
            $mediaHostGridBack = New-Object System.Windows.Controls.Grid
            
            $mediaPlayer = New-Object System.Windows.Controls.MediaElement -Property @{ Name="player$i"; Stretch="Fill"; LoadedBehavior="Manual" }
            $overlayTextBlock = New-Object System.Windows.Controls.TextBlock -Property @{ Name="overlay$i"; HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGridFront.Children.Add($mediaPlayer) | Out-Null
            $mediaHostGridFront.Children.Add($overlayTextBlock) | Out-Null

            # The back media player is a clone that just plays; its events aren't used for logic.
            $mediaPlayerBack = New-Object System.Windows.Controls.MediaElement -Property @{ Name="playerBack$i"; Stretch="Fill"; LoadedBehavior="Manual" }
            $overlayTextBlockBack = New-Object System.Windows.Controls.TextBlock -Property @{ Name="overlayBack$i"; HorizontalAlignment='Center'; VerticalAlignment='Center'; TextWrapping='Wrap'; TextAlignment='Center'; IsHitTestVisible=$false }
            $mediaHostGridBack.Children.Add($mediaPlayerBack) | Out-Null
            $mediaHostGridBack.Children.Add($overlayTextBlockBack) | Out-Null

            # --- Create the 3D Model for this panel using GeometryModel3D ---
            $panelModel = New-Object System.Windows.Media.Media3D.GeometryModel3D
            $panelModel.Geometry = $panelMesh
            
            # Create brushes that will render our 2D grids onto the 3D model
            $frontBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGridFront }
            $backBrush = New-Object System.Windows.Media.VisualBrush -Property @{ Visual = $mediaHostGridBack }

            # Create materials from the brushes
            $materialType = if ($TransparentFacesCheckbox.Checked) { [System.Windows.Media.Media3D.EmissiveMaterial] } else { [System.Windows.Media.Media3D.DiffuseMaterial] }
            $frontMaterial = New-Object $materialType -Property @{ Brush = $frontBrush }
            $backMaterial = New-Object $materialType -Property @{ Brush = $backBrush }

            # Assign materials to the front and back of the model
            $panelModel.Material = $frontMaterial
            $panelModel.BackMaterial = $backMaterial

            # Rotate this panel into its position in the circle
            $angle = $i * (360.0 / $panelCount)
            $rotation = New-Object System.Windows.Media.Media3D.AxisAngleRotation3D([System.Windows.Media.Media3D.Vector3D]::new(0,1,0), $angle)
            $panelModel.Transform = New-Object System.Windows.Media.Media3D.RotateTransform3D($rotation)

            # Add the single, two-sided model to the scene
            $objectContainer.Content.Children.Add($panelModel) | Out-Null

            # Store controls for later access
            $null = $SyncHash.MediaPlayers.Add($mediaPlayer)
            $null = $SyncHash.MediaPlayersBack.Add($mediaPlayerBack)
            $null = $SyncHash.OverlayTextBlocks.Add($overlayTextBlock)
            $null = $SyncHash.OverlayTextBlocksBack.Add($overlayTextBlockBack)
            $null = $SyncHash.MediaHostGrids.Add($mediaHostGridFront)
            $null = $SyncHash.MediaHostGridsBack.Add($mediaHostGridBack)
        }

        # Initialize a state tracker for each player
        $SyncHash.PlayerState = [hashtable]::Synchronized(@{})
        foreach ($player in $SyncHash.MediaPlayers) {
            $SyncHash.PlayerState[$player.Name] = @{
                IsImage = $false; ImageTimer = $null; RecoveryTimer = $null
                PlaybackStopwatch = New-Object System.Diagnostics.Stopwatch; IsFailed = $false
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
                if ($playerState.IsFailed) { return }
                $playerState.IsFailed = $true

                $fileName = if ($ErrorElement.Source) { $ErrorElement.Source.Segments[-1] } else { "an unknown media file" }
                Write-Warning "Media failed for player $($ErrorElement.Name) (File: '$fileName'). Reason: $Reason."

                # Find the corresponding MediaHostGrid and set its background to black for visual feedback
                $playerIndex = $SyncHash.MediaPlayers.IndexOf($ErrorElement)
                if ($playerIndex -ge 0) { 
                    $SyncHash.MediaHostGrids[$playerIndex].Background = [System.Windows.Media.Brushes]::Black
                    $SyncHash.MediaHostGridsBack[$playerIndex].Background = [System.Windows.Media.Brushes]::Black
                }

                $ErrorElement.Visibility = "Collapsed"
                $ErrorElement.Stop()

                if ($playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                # Shorten recovery time to 1 second for a better user experience
                $recoveryTimer = New-Object System.Windows.Threading.DispatcherTimer -Property @{ Interval = [TimeSpan]::FromSeconds(1); Tag = $ErrorElement }

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

        # Handler for explicit media failures
        $MediaFailedHandler = {
            param($Sender, $EventArgs)
            $reason = if ($EventArgs.ErrorException) { $EventArgs.ErrorException.Message } else { "MediaFailed event fired." }
            & $SyncHash.HandleMediaFailure -ErrorElement $Sender -Reason $reason
        }

        # Handler for when media is successfully opened
        $MediaOpenedHandler = {
            param($Sender, $EventArgs)
            $playerState = $SyncHash.PlayerState[$Sender.Name]
            $playerState.IsFailed = $false
            $playerState.PlaybackStopwatch.Restart()

            # Reset the background in case it was set to black from a previous error
            $playerIndex = $SyncHash.MediaPlayers.IndexOf($Sender)
            if ($playerIndex -ge 0) { 
                $SyncHash.MediaHostGrids[$playerIndex].Background = [System.Windows.Media.Brushes]::Transparent
                $SyncHash.MediaHostGridsBack[$playerIndex].Background = [System.Windows.Media.Brushes]::Transparent
            }

            $SyncHash.Window.Dispatcher.Invoke([action]{
                # Find the index of the player
                $playerIndex = $SyncHash.MediaPlayers.IndexOf($Sender)
                $Sender.Visibility = "Visible"
            })

            if ($playerState.IsImage) {
                $Sender.Pause()
                if ($SyncHash.SelectedFiles.Count -gt $panelCount) {
                    if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
                    $playerState.ImageTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $playerState.ImageTimer.Interval = [TimeSpan]::FromSeconds(10) 
                    $playerState.ImageTimer.Tag = $Sender # Pass the MediaElement to the handler
                    
                    $tickScriptBlock = {
                        $timer = $args[0]; $timer.Stop()
                        $mediaElement = $timer.Tag
                        & $SyncHash.MediaEndedHandler -Sender $mediaElement -e $null
                    }
                    $playerState.ImageTimer.Add_Tick($tickScriptBlock) # No .GetNewClosure() needed here
                    $playerState.ImageTimer.Start()
                }
            }
            elseif (-not $Sender.NaturalDuration.HasTimeSpan) {
                & $SyncHash.HandleMediaFailure -ErrorElement $Sender -Reason "No duration found (silent failure)."
            }
            else {
                if ($playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
            }
        }

        # Event handler for MediaEnded
        $MediaEndedHandler = {
            param(
                $Sender,
                $e,
                [switch]$IsRecovery
            )
            $FinishedElement = $Sender
            if (-not $FinishedElement -or -not $FinishedElement.Name) { return }

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
            
            if ($SyncHash.SelectedFiles.Count -le $panelCount) {
                $FinishedElement.Position = [TimeSpan]::FromSeconds(0)
                $FinishedElement.Play()
            }
            else {
                # This is the correct way to increment and loop the global index
                $SyncHash.CurrentIndex = ($SyncHash.CurrentIndex + 1) % $SyncHash.SelectedFiles.Count

                $NewVideoPath = $SyncHash.VideoUris[$SyncHash.CurrentIndex]
                $NewUri = New-Object System.Uri($NewVideoPath)

                $playerState = $SyncHash.PlayerState[$FinishedElement.Name]
                $extension = [System.IO.Path]::GetExtension($NewUri.LocalPath).ToLower()
                $playerState.IsImage = ($SyncHash.ImageExtensions -contains $extension)
                
                $FinishedElement.Source = $NewUri
                if ($SyncHash.RbSelection -notmatch 'Hidden|Custom') {
                    # Find the index of the player that finished
                    $playerIndex = $SyncHash.MediaPlayers.IndexOf($FinishedElement)
                    $SyncHash.OverlayTextBlocks[$playerIndex].Text = $NewUri.Segments[-1]
                    $SyncHash.OverlayTextBlocksBack[$playerIndex].Text = $NewUri.Segments[-1]
                }
                $FinishedElement.Play()
            }
        }

        $SyncHash.MediaEndedHandler = $MediaEndedHandler

        foreach ($player in $SyncHash.MediaPlayers) {
            $player.Add_MediaFailed($MediaFailedHandler)
            $player.Add_MediaOpened($MediaOpenedHandler)
            $player.Add_MediaEnded($MediaEndedHandler)
        }

        $SyncHash.VideoUris = $SyncHash.SelectedFiles | ForEach-Object { New-Object System.Uri($_, [System.UriKind]::Absolute) }
        
        # Set initial media
        for ($i = 0; $i -lt $panelCount; $i++) {
            if ($i -lt $SyncHash.VideoUris.Count) {
                $player = $SyncHash.MediaPlayers[$i]
                $uri = $SyncHash.VideoUris[$i]
                
                $playerState = $SyncHash.PlayerState[$player.Name]
                $extension = [System.IO.Path]::GetExtension($uri.LocalPath).ToLower()
                $playerState.IsImage = ($SyncHash.ImageExtensions -contains $extension)

                $player.Source = $uri
                
                # Also set the source for the back player
                $playerBack = $SyncHash.MediaPlayersBack[$i]
                $playerBack.Source = $uri
            }
            else {
                # If not enough files, leave panel blank or show an error/placeholder
                # Since we removed the error textblocks, we'll just log to the console.
                # This will only be visible if you run the script from a console.
                Write-Warning "Not enough media files for panel $i. It will remain blank."
            }
        }
        $SyncHash.CurrentIndex = [math]::Min($panelCount, $SyncHash.SelectedFiles.Count) -1

        if($RadioButton1.Checked){ $SyncHash.RbSelection = "Hidden" }

        switch ($SyncHash.RbSelection)
        {
            "Hidden"
            {
                $SyncHash.OverlayTextBlocks.ForEach({ $_.Visibility = 'Collapsed' })
                $SyncHash.OverlayTextBlocksBack.ForEach({ $_.Visibility = 'Collapsed' })
            }
            "Filename"
            {
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $fontWeight = if ($SyncHash.BoldCheckbox.Checked) { 'Bold' } else { 'Normal' }
                $fontStyle = if ($SyncHash.ItalicCheckbox.Checked) { 'Italic' } else { 'Normal' }

                for ($i = 0; $i -lt [math]::Min($panelCount, $SyncHash.VideoUris.Count); $i++) {
                    $overlay = $SyncHash.OverlayTextBlocks[$i]
                    $overlay.Foreground = $brush.Clone(); $overlay.FontSize = $SyncHash.SelectedFontSize
                    $overlay.FontFamily = $SyncHash.SelectedFont; $overlay.FontWeight = $fontWeight
                    $overlay.FontStyle = $fontStyle; $overlay.Text = $SyncHash.VideoUris[$i].Segments[-1]
                    $overlayBack = $SyncHash.OverlayTextBlocksBack[$i]
                    $overlayBack.Foreground = $brush.Clone(); $overlayBack.FontSize = $SyncHash.SelectedFontSize
                    $overlayBack.FontFamily = $SyncHash.SelectedFont; $overlayBack.FontWeight = $fontWeight
                    $overlayBack.FontStyle = $fontStyle; $overlayBack.Text = $SyncHash.VideoUris[$i].Segments[-1]
                }
            }
            "Custom"
            {
                $mediaColor = [System.Windows.Media.Color]::FromArgb($SyncHash.TextColor.A, $SyncHash.TextColor.R, $SyncHash.TextColor.G, $SyncHash.TextColor.B)
                $brush = New-Object System.Windows.Media.SolidColorBrush($mediaColor)
                $fontWeight = if ($SyncHash.BoldCheckbox.Checked) { 'Bold' } else { 'Normal' }
                $fontStyle = if ($SyncHash.ItalicCheckbox.Checked) { 'Italic' } else { 'Normal' }

                foreach ($overlay in $SyncHash.OverlayTextBlocks) {
                    $overlay.Foreground = $brush.Clone(); $overlay.FontSize = $SyncHash.SelectedFontSize
                    $overlay.FontFamily = $SyncHash.SelectedFont; $overlay.FontWeight = $fontWeight
                    $overlay.FontStyle = $fontStyle; $overlay.Text = $TextBox.Text
                }
                foreach ($overlay in $SyncHash.OverlayTextBlocksBack) {
                    $overlay.Foreground = $brush.Clone(); $overlay.FontSize = $SyncHash.SelectedFontSize
                    $overlay.FontFamily = $SyncHash.SelectedFont; $overlay.FontWeight = $fontWeight
                    $overlay.FontStyle = $fontStyle; $overlay.Text = $TextBox.Text
                }
            }
            Default
            {
                $SyncHash.OverlayTextBlocks.ForEach({ $_.Visibility = 'Collapsed' })
                $SyncHash.OverlayTextBlocksBack.ForEach({ $_.Visibility = 'Collapsed' })
            }
        }

        # Create the rotation animation
        $RotateAnimation = New-Object System.Windows.Media.Animation.DoubleAnimation
        $RotateAnimation.From = 0
        $RotateAnimation.To = 360
        $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds(30))
        $RotateAnimation.RepeatBehavior = "Forever"
        $RotateAnimation.IsCumulative = $True

        $Window.Add_Loaded({
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                foreach ($player in $SyncHash.MediaPlayers) { $player.Play() }
                foreach ($player in $SyncHash.MediaPlayersBack) {
                    $player.Play()
                }
            })

        $Window.Add_Closed({
                if ($SyncHash.Contains("MediaPlayers")) {
                    $playerNames = $SyncHash.MediaPlayers | ForEach-Object { $_.Name }
                    foreach ($name in $PlayerNames) {
                        if ($SyncHash.PlayerState.ContainsKey($name)) {
                            $playerState = $SyncHash.PlayerState[$name]
                            if ($playerState -and $playerState.ImageTimer) { $playerState.ImageTimer.Stop() }
                            if ($playerState -and $playerState.RecoveryTimer) { $playerState.RecoveryTimer.Stop() }
                        }
                    }

                    foreach ($player in $SyncHash.MediaPlayers) {
                        $player.Stop()
                        try {
                            $player.Remove_MediaEnded($SyncHash.MediaEndedHandler)
                            $player.Remove_MediaFailed($MediaFailedHandler)
                            $player.Remove_MediaOpened($MediaOpenedHandler)
                        } catch {}
                        $player.Source = $null
                        $player.Close()
                    }
                    foreach ($player in $SyncHash.MediaPlayersBack) {
                        $player.Stop()
                        $player.Close()
                    }
                }
                # Do not clear SyncHash here if we are redoing
                if (-not $SyncHash.RedoClicked) {
                    $SyncHash.Clear()
                }
            })

        $Window.Add_KeyDown({
                param($Sender, $e)
                
                switch ($e.Key)
                {
                    "F1"
                    {
                        $ReaderPopup = (New-Object System.Xml.XmlNodeReader $XamlHelpPopup)
                        $PopupWindow = [Windows.Markup.XamlReader]::Load($ReaderPopup)
                        $OkButton = $PopupWindow.FindName("OKButton")
                        $OkButton.Add_Click({ $PopupWindow.Close() })
                        $PopupWindow.ShowDialog() | Out-Null
                    }
                    "Left"
                    {
                        $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                        $NewDuration = $CurrentDuration * 2
                        if($NewDuration -le 0){ $NewDuration = 1 }
                        $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                        $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    }
                    "Right"
                    {
                        $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                        $NewDuration = $CurrentDuration / 2
                        $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                        $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                        $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    }
                    "Escape"{ $SyncHash.RedoClicked = $false; $Window.Close() }
                    "p"
                    {
                        if($SyncHash.Paused -eq $False)
                        {
                            $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                            $SyncHash.AxisAngleY.Angle = $Current_angleY
                            $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                            $SyncHash.PauseButton.Content = "Resume"; $SyncHash.Paused = $True
                        }
                        else
                        {
                            $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                            $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                            $SyncHash.PauseButton.Content = "Pause"; $SyncHash.Paused = $False
                        }
                    }
                    "R"  { $SyncHash.RedoClicked = $true; $Window.Close() }
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
                        }
                    }
                }
            })

        $SyncHash.CloseButton.Add_Click({ $SyncHash.RedoClicked = $false; $Window.Close() })

        $SyncHash.SlowDown.Add_Click({
                $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                $NewDuration = $CurrentDuration * 2
                if($NewDuration -le 0){ $NewDuration = 1 }
                $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
            })

        $SyncHash.SpeedUp.Add_Click({
                $CurrentDuration = $RotateAnimation.Duration.TimeSpan.TotalSeconds
                $NewDuration = $CurrentDuration / 2
                $RotateAnimation.Duration = New-Object System.Windows.Duration([TimeSpan]::FromSeconds($NewDuration))
                $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
            })

        $SyncHash.PauseButton.Add_Click({
                if($SyncHash.Paused -eq $False)
                {
                    $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                    $SyncHash.AxisAngleY.Angle = $Current_angleY
                    $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                    $SyncHash.PauseButton.Content = "Resume"; $SyncHash.Paused = $True
                }
                else
                {
                    $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $SyncHash.PauseButton.Content = "Pause"; $SyncHash.Paused = $False
                }
            })

        $SyncHash.ReDoButton.Add_Click({
                $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                $SyncHash.RedoClicked = $true
                $Window.Close()
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
                }
            })

        $SyncHash.MainGrid.Add_MouseDown({
                param($Sender, $e)
                if($SyncHash.Paused -eq $False)
                {
                    $Current_angleY = $SyncHash.AxisAngleY.GetValue([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty)
                    $SyncHash.AxisAngleY.Angle = $Current_angleY
                    $SyncHash.AxisAngleY.BeginAnimation([Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $Null)
                    $SyncHash.PauseButton.Content = "Resume"; $SyncHash.Paused = $True
                }
                else
                {
                    $RotateAnimation.From = $SyncHash.AxisAngleY.Angle
                    $SyncHash.AxisAngleY.BeginAnimation([System.Windows.Media.Media3D.AxisAngleRotation3D]::AngleProperty, $RotateAnimation)
                    $SyncHash.PauseButton.Content = "Pause"; $SyncHash.Paused = $False
                }
            })

        # Show the window
        $SelectFolderForm.Hide()

        $SyncHash.FunnelReady = $true
        Start-Sleep -Milliseconds 200

        $Window.ShowDialog() | Out-Null

        # After the 3D window closes, check if we need to loop back to the selection form.
        if (-not $SyncHash.RedoClicked) {
            break # Exit the while loop
        }
    }
}

$SelectFolderForm.Dispose()
