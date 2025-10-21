<#
.SYNOPSIS
    A dynamic GUI that scans its directory for other PowerShell scripts and creates launch buttons for them.

.DESCRIPTION
    This script generates a Windows Forms GUI that discovers all other PowerShell (.ps1) scripts
    located in the same directory. For each script found, it creates a clickable button to launch it.

    The launcher intelligently groups scripts into collapsible sections based on their dependencies,
    which it determines by parsing the content of each script for a '$RequiredExecutables' variable.
    If a dependency is not found in the system's PATH, the button for that script is disabled and shows
    which executables are missing.

    It also features a 'Help' button that aggregates descriptions from all discovered scripts (parsed from
    a '$ScriptDescription' variable within each) into a single, formatted view.

.EXAMPLE
    PS C:\> .\Show-ScriptLauncher.ps1

    Launches the GUI. The application will then scan the current directory, populate the window with
    buttons for each script, and await user interaction.

.NOTES
    Name:           Show-ScriptLauncher.ps1
    Version:        1.0.0, 10/18/2025
    Author:         JD Alberthal (jd@jdalberthal.com)
    Website:        https://www.jdalberthal.com
    GitHub:         https://github.com/jdalberthal
    Dependencies:   Requires PowerShell with .NET Framework access for Windows Forms and WPF assemblies.
#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add WPF and Automation assemblies to prevent parsing errors when the launcher reads child scripts that use them.
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, WindowsFormsIntegration, System.Xaml, System.Management.Automation

[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptName = $MyInvocation.MyCommand.Name

# --- Form Definition ---
$mainForm = New-Object System.Windows.Forms.Form -Property @{
    Text          = "PowerShell Script Launcher"
    Size          = New-Object System.Drawing.Size(800, 600)
    StartPosition = "CenterScreen"
}

# --- Controls ---

# Main layout panel to hold the groups
$mainTableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    Dock        = [System.Windows.Forms.DockStyle]::Fill
    ColumnCount = 1
    AutoScroll  = $true
}

# Help Button
$helpButton = New-Object System.Windows.Forms.Button -Property @{
    Text   = "Help"
    Anchor = ([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right)
    Size   = New-Object System.Drawing.Size(75, 23)
}
# Position it at the top right corner of the form
$helpButton.Location = New-Object System.Drawing.Point(($mainForm.ClientSize.Width - $helpButton.Width - 10), 10)

# --- Functions ---

function Create-ScriptButtons {
    param (
        [string]$FolderPath,
        [System.Windows.Forms.TableLayoutPanel]$LayoutTable,
        [string]$ScriptName
    )

    # This hashtable will store the FlowLayoutPanels for each dependency group.
    $groupPanels = @{}

    # Clear any previously created groups from the main layout.
    $LayoutTable.Controls.Clear()
    $LayoutTable.RowStyles.Clear()
    $LayoutTable.RowCount = 0

    # Find all .ps1 files in the selected folder
    $scripts = Get-ChildItem -Path $FolderPath -Filter "*.ps1" -ErrorAction SilentlyContinue

    if (-not $scripts) {
        [System.Windows.Forms.MessageBox]::Show("No PowerShell scripts (.ps1) found in the selected folder.", "No Scripts Found", "OK", "Information")
        return
    }

    # Create a button for each script
    foreach ($script in $scripts) {
        # Skip the launcher script itself
        if ($script.Name -eq $ScriptName) {
            continue
        }

        # --- Dependency and Button Text Logic ---
        $buttonText = $script.Name
        $dependenciesMet = $true
        $scriptDependencies = [System.Collections.Generic.List[string]]::new()
        $missingDependencies = [System.Collections.Generic.List[string]]::new()

        $scriptContent = ""
        try {
            $scriptContent = [System.IO.File]::ReadAllText($script.FullName)
        } catch {
            Write-Warning "Could not read or parse '$($script.Name)'. Using default name and assuming no dependencies."
        }

        # Check for $ExternalButtonName
        if ($scriptContent -match '\$ExternalButtonName\s*=\s*["''](.*?)["'']') {
            $buttonText = $matches[1].Replace('`n', "`n")
        }

        # Check for $ScriptDescription
        $scriptDescription = "No description available."
        if ($scriptContent -match '\$ScriptDescription\s*=\s*["''](.*?)["'']') {
            $scriptDescription = $matches[1]
        }

        # Treat "MediaElement" as a dependency for grouping
        if ($script.Name -match "MediaElement") {
            $scriptDependencies.Add("MediaElement")
        }

        # Check for $RequiredExecutables
        if ($scriptContent -match '\$RequiredExecutables\s*=\s*@\((.*?)\)') {
            # Extract the array content and split it into individual executable names
            $executablesString = $matches[1].Trim()
                $requiredExecutables = [string[]]($executablesString -split ',\s*' | ForEach-Object { $_.Trim('"'' ') })
            $scriptDependencies.AddRange($requiredExecutables)

            # Check if each required executable is available in the PATH
            foreach ($exe in $requiredExecutables) {
                # Check in the PATH and in the script's local directory
                $localPath = Join-Path $script.DirectoryName $exe
                if (-not (Get-Command $exe -ErrorAction SilentlyContinue) -and -not (Test-Path -Path $localPath)) {
                    $dependenciesMet = $false
                    $missingDependencies.Add($exe)
                }
            }
        }

        # --- Dynamic Grouping Logic ---
        # Create a unique, sorted key for the dependency group.
        $dependencyKey = if ($scriptDependencies.Count -gt 0) {
            ($scriptDependencies | Sort-Object) -join ', '
        } else {
            "No Dependencies"
        }

        # Check if a panel for this group already exists. If not, create it.
        if (-not $groupPanels.ContainsKey($dependencyKey)) {
            $LayoutTable.RowCount++
            $LayoutTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

            $groupBox = New-Object System.Windows.Forms.GroupBox -Property @{
                Text = "Dependencies: $dependencyKey"
                Dock = [System.Windows.Forms.DockStyle]::Fill
                Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }
            $flowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel -Property @{ Dock = [System.Windows.Forms.DockStyle]::Fill; AutoScroll = $true; Padding = New-Object System.Windows.Forms.Padding(10) }
            $groupBox.Controls.Add($flowLayoutPanel)
            
            $LayoutTable.Controls.Add($groupBox, 0, ($LayoutTable.RowCount - 1))
            $groupPanels[$dependencyKey] = $flowLayoutPanel
        }
        $targetPanel = $groupPanels[$dependencyKey]

        if (-not $dependenciesMet) {
            # Create a label to show missing dependencies instead of a button
            $labelText = "$buttonText`n(Missing: $($missingDependencies -join ', '))"
            $dependencyLabel = New-Object System.Windows.Forms.Label
            $dependencyLabel.Text = $labelText
            $dependencyLabel.Size = New-Object System.Drawing.Size(200, 50)
            $dependencyLabel.Margin = New-Object System.Windows.Forms.Padding(5)
            $dependencyLabel.ForeColor = "GrayText"
            $dependencyLabel.BackColor = "ControlLight"
            $dependencyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $dependencyLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

            # Add the label to the correct panel
            $targetPanel.Controls.Add($dependencyLabel)
            continue # Move to the next script
        }

        $scriptButton = New-Object System.Windows.Forms.Button -Property @{
            Text   = $buttonText          # Display the determined name
            # Store a custom object in the Tag property
            Tag    = [PSCustomObject]@{ Path = $script.FullName; Description = $scriptDescription }
            Size   = New-Object System.Drawing.Size(200, 50)
            Margin = New-Object System.Windows.Forms.Padding(5)
        }

        # Add the click event handler
        $scriptButton.Add_Click({
            $scriptPath = $this.Tag
            # The mainForm is in the script scope, so we can access it directly.
            try { # The Tag is now a PSCustomObject, so we access the Path property
                # Hide the main launcher window before running the script.
                $mainForm.Hide()

                # Execute the script. This is a blocking call because the
                # target scripts use .ShowDialog(), so the code here will
                # wait until the script's window is closed.
                & $scriptPath.Path 2>&1 | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("An error occurred while running:`n$scriptPath`n`n$($_.Exception.Message)", "Execution Error", "OK", "Error")
            }
            finally {
                # Ensure the main form is always shown again, even if an error occurs.
                $mainForm.Show()
                $mainForm.Activate() # Bring the launcher back to the front.
            }
        })

        # Add the button to the correct panel based on its name
        $targetPanel.Controls.Add($scriptButton)
    }
}

# --- Event Handlers ---
$mainForm.Add_Load({
    $scriptRoot = $PSScriptRoot
    Create-ScriptButtons -FolderPath $scriptRoot -LayoutTable $mainTableLayoutPanel -ScriptName $ScriptName
})

$helpButton.Add_Click({
    $helpText = New-Object System.Text.StringBuilder

    # Helper function to process a panel's buttons
    $processPanel = {
        param($GroupBox)
        $GroupName = $GroupBox.Text
        $Panel = $GroupBox.Controls[0] # The FlowLayoutPanel is the first control

        if ($Panel.Controls.Count -gt 0 -and ($Panel.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] })) {
            # Use RTF codes for formatting: \par is a new paragraph, \b is bold, \b0 is no bold.
            $helpText.AppendLine("\par\b --- $($GroupName) ---\b0\par") | Out-Null
            foreach ($control in $Panel.Controls) {
                # Only process controls that are Buttons and have a Tag with a Description
                if ($control -isnot [System.Windows.Forms.Button] -or -not $control.Tag -or -not $control.Tag.Description) {
                    continue
                }

                $buttonInfo = $control.Tag
                $buttonTitle = $control.Text.Replace("`n", " ") # Make button text single-line for the title

                # Escape backslashes and curly braces in the text to prevent RTF parsing errors
                $escapedTitle = $buttonTitle.Replace('\', '\\').Replace('{', '\{').Replace('}', '\}')
                $escapedDescription = $buttonInfo.Description.Replace('\', '\\').Replace('{', '\{').Replace('}', '\}')

                $helpText.AppendLine("\b $($escapedTitle)\b0\par") | Out-Null # Bold title
                $helpText.AppendLine(" $($escapedDescription)\par\par") | Out-Null
            }
        }
    }

    # Process each panel
    foreach ($groupBox in $mainTableLayoutPanel.Controls) {
        & $processPanel $groupBox
    }

    # Wrap the generated content in a valid RTF document structure
    $rtfHeader = "{\rtf1\ansi\deff0{\fonttbl{\f0 Segoe UI;}}"
    $rtfContent = "$rtfHeader $($helpText.ToString()) }"

    # Create and show the help form
    $helpForm = New-Object System.Windows.Forms.Form -Property @{
        Text          = "Script Descriptions"
        Size          = New-Object System.Drawing.Size(600, 400)
        StartPosition = "CenterParent"
        ShowInTaskbar = $false
        FormBorderStyle = "SizableToolWindow"
    }
    $richTextBox = New-Object System.Windows.Forms.RichTextBox -Property @{
        Dock     = [System.Windows.Forms.DockStyle]::Fill
        ReadOnly = $true
        Rtf      = $rtfContent # Use the fully-formed RTF string
        Font     = New-Object System.Drawing.Font("Segoe UI", 10)
    }
    $helpForm.Controls.Add($richTextBox)
    $helpForm.ShowDialog($mainForm) | Out-Null
    $helpForm.Dispose()
})

# --- Form Setup and Display ---
$mainForm.Controls.Add($mainTableLayoutPanel)
# Add the help button on top of the TableLayoutPanel
$mainForm.Controls.Add($helpButton)
$helpButton.BringToFront()

[void]$mainForm.ShowDialog()

$mainForm.Dispose()