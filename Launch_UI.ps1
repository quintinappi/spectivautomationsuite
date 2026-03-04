<#
Inventor Automation Suite - Professional UI Launcher
Compatible with Windows 7+ PowerShell 2.0+
Version: 1.0.0
#>

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName System.Drawing | Out-Null

# Form setup
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Inventor Automation Suite"
$MainForm.Size = New-Object System.Drawing.Size(1200, 850)
$MainForm.StartPosition = "CenterScreen"
$MainForm.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$MainForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$MainForm.ClientSize = New-Object System.Drawing.Size(1200, 850)

# Menu bar
$menuBar = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.Add_Click({ $MainForm.Close() })
$fileMenu.DropDownItems.Add($exitMenuItem)

$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "View"
$refreshMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$refreshMenuItem.Text = "Refresh Status"
$refreshMenuItem.Add_Click({ Update-Status "Ready"; Show-CategoryButtons })
$viewMenu.DropDownItems.Add($refreshMenuItem)

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
$aboutMenuItem.Add_Click({ Show-About })
$helpMenu.DropDownItems.Add($aboutMenuItem)

$menuBar.Items.Add($fileMenu)
$menuBar.Items.Add($viewMenu)
$menuBar.Items.Add($helpMenu)
$MainForm.MainMenuStrip = $menuBar
$MainForm.Controls.Add($menuBar)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Text = "Ready"
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$MainForm.Controls.Add($statusBar)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 675)
$progressBar.Size = New-Object System.Drawing.Size(1170, 20)
$progressBar.Style = "Continuous"
$progressBar.Visible = $false
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$MainForm.Controls.Add($progressBar)

# Label for searching
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(10, 35)
$searchLabel.Size = New-Object System.Drawing.Size(60, 23)
$searchLabel.Text = "Search:"
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$MainForm.Controls.Add($searchLabel)

# Search box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(75, 33)
$searchBox.Size = New-Object System.Drawing.Size(200, 25)
$searchBox.Add_TextChanged({ Filter-Buttons })
$searchBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$MainForm.Controls.Add($searchBox)

# Category sidebar (TreeView)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Location = New-Object System.Drawing.Point(10, 65)
$treeView.Size = New-Object System.Drawing.Size(220, 520)
$treeView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$treeView.BackColor = [System.Drawing.Color]::White
$treeView.ForeColor = [System.Drawing.Color]::Black
$treeView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left

$treeView.Nodes.Add("All Items") | Out-Null

# Add category nodes
$categoryNames = @(
    "Core Production Workflow",
    "Management & Utilities",
    "Rescue & Synchronization",
    "Cloning Tools",
    "iLogic & Analysis",
    "Sheet Metal Conversion",
    "Drawing Customization",
    "View Management",
    "Parts List and BOM",
    "Parameter Management"
)

foreach ($cat in $categoryNames) {
    $treeView.Nodes.Add($cat) | Out-Null
}

$treeView.Add_AfterSelect({ Show-CategoryButtons }) | Out-Null
$MainForm.Controls.Add($treeView) | Out-Null

# Main button panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(240, 65)
$buttonPanel.Size = New-Object System.Drawing.Size(940, 520)
$buttonPanel.BackColor = [System.Drawing.Color]::White
$buttonPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$buttonPanel.AutoScroll = $true
$MainForm.Controls.Add($buttonPanel)

# Log window label
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(10, 595)
$logLabel.Size = New-Object System.Drawing.Size(50, 23)
$logLabel.Text = "Log:"
$logLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$logLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$MainForm.Controls.Add($logLabel)

# Splash image
$splashImage = New-Object System.Windows.Forms.PictureBox
$splashImage.Location = New-Object System.Drawing.Point(10, 65)
$splashImage.Size = New-Object System.Drawing.Size(1170, 520)
$splashImage.SizeMode = "Zoom"
$splashImage.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$splashImage.BackColor = [System.Drawing.Color]::White
$splashImage.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$splashPath = Join-Path $PSScriptRoot "assets\splash.png"
if (Test-Path $splashPath) {
    $splashImage.ImageLocation = $splashPath
}
$MainForm.Controls.Add($splashImage)

# Log window
$logTextBox = New-Object System.Windows.Forms.TextBox
$logTextBox.Location = New-Object System.Drawing.Point(10, 620)
$logTextBox.Size = New-Object System.Drawing.Size(1170, 50)
$logTextBox.Multiline = $true
$logTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logTextBox.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 240)
$logTextBox.ReadOnly = $true
$logTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$logTextBox.Height = 50
$MainForm.Controls.Add($logTextBox)

# Define all menu items with their categories
$script:allMenuItems = @{
    "Core Production Workflow" = @(
        @{ Name = "Assembly Renamer"; Script = "Part_Renaming\Launch_Assembly_Renamer.bat"; Description = "RENAME WORKFLOW STEP 1: Rename parts with heritage method. Creates STEP_1_MAPPING.txt automatically. RUN FIRST." },
        @{ Name = "Update Same-Folder Derived Parts"; Script = "Part_Renaming\Launch_Update_SameFolder_Derived_Parts.bat"; Description = "RENAME WORKFLOW STEP 2: Fix local derived parts. ONLY if you have derived parts in same folder. OPTIONAL." },
        @{ Name = "iLogic Patcher"; Script = "Part_Renaming\Launch_iLogic_Patcher.bat"; Description = "RENAME WORKFLOW STEP 3: Fix iLogic rules. ONLY if assembly has iLogic rules. OPTIONAL." },
        @{ Name = "IDW Updates"; Script = "IDW_Updates\Launch_IDW_Reference_Updater.bat"; Description = "RENAME WORKFLOW STEP 4: Update ALL IDW drawings to new part names using mapping file. ALWAYS REQUIRED." },
        @{ Name = "Title Automation"; Script = "Title_Automation\Launch_Title_Updater.bat"; Description = "Update IDW drawing titles. Run after IDW updates complete. OPTIONAL." }
    )
    "Management & Utilities" = @(
        @{ Name = "Registry Management"; Script = "Registry_Management\Launch_Registry_Manager.bat"; Description = "Manage numbering counters and database" },
        @{ Name = "File Utilities"; Script = "File_Utilities\Launch_Duplicate_File_Finder.bat"; Description = "Find duplicate files and conflicts" },
        @{ Name = "Unused Part Finder"; Script = "File_Utilities\Launch_Unused_Part_Finder.bat"; Description = "Cleanup: Find and backup IPT files not used in open assembly. Scan assembly + folders to identify orphaned parts." },
        @{ Name = "Deploy Inventor Add-In"; Script = "InventorAddIn\Deploy_AddIn.bat"; Description = "Install/Update Assembly Cloner Add-In for Inventor"; CheckInstalled = $true }
    )
    "Rescue & Synchronization" = @(
        @{ Name = "Smart Prefix Scanner"; Script = "Part_Renaming\Launch_Smart_Prefix_Scanner.bat"; Description = "BEFORE renaming: Scan model to detect and prevent duplicate part numbers" },
        @{ Name = "Emergency IDW Fixer"; Script = "IDW_Updates\Launch_Emergency_IDW_Fixer.bat"; Description = "TROUBLESHOOTING: Fix specific IDW folders missed by Step 4. Use when IDW Updates didn't catch all drawings" },
        @{ Name = "IDW-Assembly Synchronizer"; Script = "IDW_Updates\Launch_IDW_Assembly_Synchronizer.bat"; Description = "TROUBLESHOOTING: Sync scattered IDWs when they use old assembly names. Use when IDWs reference renamed assembly" }
    )
    "Cloning Tools" = @(
        @{ Name = "Assembly Cloner"; Script = "Part_Renaming\Launch_Assembly_Cloner.bat"; Description = "Clone assembly + parts to new folder. CLONING WORKFLOW STEP 1" },
        @{ Name = "Cloner (Prefix Changer Only)"; Script = "Part_Renaming\Launch_Prefix_Cloner.bat"; Description = "Clone assembly replacing ONLY the prefix (e.g., N1SCR04-780- to N2SCR04-780-). Keeps part suffixes intact." },
        @{ Name = "Part Cloner"; Script = "Part_Renaming\Launch_Part_Cloner.bat"; Description = "Clone individual part to new folder" },
        @{ Name = "Fix Derived Parts (Post-Clone)"; Script = "Part_Renaming\Launch_Fix_Derived_Parts.bat"; Description = "[POST-CLONE] Fix EXTERNAL derived part references after cloning. CLONING WORKFLOW STEP 2" }
    )
    "iLogic & Analysis" = @(
        @{ Name = "iLogic Scanner"; Script = "iLogic_Tools\Launch_iLogic_Scanner.bat"; Description = "Scan/export iLogic rules from document" },
        @{ Name = "iLogic Patcher"; Script = "Part_Renaming\Launch_iLogic_Patcher.bat"; Description = "[POST-RENAME] Rename component references in iLogic rules after part renaming. Use after Assembly_Renamer to fix broken iLogic references." },
        @{ Name = "Find Missing Detailed Parts"; Script = "iLogic_Tools\Launch_Find_Missing_Detailed_Parts.bat"; Description = "Check which assembly parts haven't been detailed" }
    )
    "Sheet Metal Conversion" = @(
        @{ Name = "Sheet Metal Converter (Assembly)"; Script = "iLogic_Tools\Launch_Sheet_Metal_Converter.bat"; Description = "Convert all plate parts in assembly" },
        @{ Name = "Sheet Metal Converter (Part)"; Script = "iLogic_Tools\Launch_Sheet_Metal_Part_Converter.bat"; Description = "Convert single part to sheet metal" }
    )
    "Drawing Customization" = @(
        @{ Name = "Change Balloon Style"; Script = "Title_Automation\Launch_Change_Balloon_Style.bat"; Description = "Replace all balloons with selected style" },
        @{ Name = "Change Dimension Style"; Script = "Title_Automation\Launch_Change_Dimension_Style.bat"; Description = "Replace all dimensions with selected style" },
        @{ Name = "Export IDW Sheets to PDF"; Script = "iLogic_Tools\Launch_Export_IDW_Sheets_to_PDF.bat"; Description = "Export each sheet as separate PDF" },
        @{ Name = "Master Style Replicator"; Script = "View_Style_Manager\Launch_Master_Style_Replicator.bat"; Description = "Copy view styling from Master View to other views" }
    )
    "Parts List and BOM" = @(
        @{ Name = "Create Sheet Parts List"; Script = "iLogic_Tools\Launch_Create_Sheet_Parts_List.bat"; Description = "Create parts list for components visible on current sheet" },
        @{ Name = "Clean Up unused Files"; Script = "File_Utilities\Launch_IDW_Parts_List_Scanner.bat"; Description = "Cleanup: After renaming parts, move old/unreferenced IPT files to Unrenamed Parts folder" },
        @{ Name = "Populate DWG REF from Parts Lists"; Script = "File_Utilities\Launch_Populate_DWG_REF_From_Parts_Lists.bat"; Description = "Scan all IDW sheets, collect parts from parts lists, and populate DWG REF." },
        @{ Name = "Populate DWG REF + Auto-place Missing Parts"; Script = "File_Utilities\Launch_Populate_DWG_REF_From_Parts_Lists_AutoPlace_Sheet5.bat"; Description = "Runs DWG REF update and places all missing parts; plate parts get folded+flat views. Prompts for target non-DXF sheet." },
        @{ Name = "CREATE DXF FOR MODEL PLATES"; Script = "File_Utilities\Launch_Create_DXF_For_Model_Plates.bat"; Description = "Select source sheet, find assembly, create DXF FOR {MODEL} sheet, place all plate parts at 1:1, and create plate-only parts list." }
    )
    "View Management" = @(
        @{ Name = "Master Style Replicator"; Script = "View_Style_Manager\Launch_Master_Style_Replicator.bat"; Description = "Copy view styling from Master View to other views" }
    )
    "Parameter Management" = @(
        @{ Name = "Length Parameter Exporter"; Script = "iLogic_Tools\Launch_Length_Parameter_Exporter.bat"; Description = "Enable export for Length params on non-plate parts" },
        @{ Name = "Length2 Parameter Exporter"; Script = "iLogic_Tools\Launch_Length2_Parameter_Exporter.bat"; Description = "Enable export for Length2 user params on non-plate parts" },
        @{ Name = "Thickness Parameter Exporter"; Script = "iLogic_Tools\Launch_Thickness_Parameter_Exporter.bat"; Description = "Enable export for Thickness params on plate parts" },
        @{ Name = "Fix Non-Plate Parts"; Script = "iLogic_Tools\Launch_Fix_Non_Plate_Parts.bat"; Description = "Add Length2 parameter to parts missing Length" },
        @{ Name = "Fix Single Part Length2"; Script = "iLogic_Tools\Launch_Fix_Non_Plate_Parts.bat"; Description = "Add Length2 parameter to active part (longest dimension)" },
        @{ Name = "Fix BOM Plate Dimensions"; Script = "iLogic_Tools\Launch_Fix_BOM_Plate_Dimensions.bat"; Description = "Add WIDTH and LENGTH columns to BOM for plate parts (PL/VRN/S355JR) using sheet metal flat pattern dimensions" },
        @{ Name = "Refresh BOM Precision"; Script = "iLogic_Tools\Force_BOM_Precision_Robust.vbs"; Description = "Fix BOM decimal precision display - ROBUST version with user checkpoint after each part" }
    )
}

# Category colors for professional appearance
$script:categoryColors = @{
    "Core Production Workflow" = [System.Drawing.Color]::FromArgb(0, 120, 215)
    "Management & Utilities" = [System.Drawing.Color]::FromArgb(16, 110, 190)
    "Rescue & Synchronization" = [System.Drawing.Color]::FromArgb(0, 90, 180)
    "Cloning Tools" = [System.Drawing.Color]::FromArgb(88, 89, 91)
    "iLogic & Analysis" = [System.Drawing.Color]::FromArgb(16, 110, 190)
    "Sheet Metal Conversion" = [System.Drawing.Color]::FromArgb(202, 80, 16)
    "Drawing Customization" = [System.Drawing.Color]::FromArgb(127, 127, 127)
    "View Management" = [System.Drawing.Color]::FromArgb(0, 153, 153)
    "Parts List and BOM" = [System.Drawing.Color]::FromArgb(33, 150, 243)
    "Parameter Management" = [System.Drawing.Color]::FromArgb(255, 140, 0)
}

# Show buttons for selected category
function Show-CategoryButtons {
    $buttonPanel.Controls.Clear()

    # Hide splash image when category is selected
    $splashImage.Visible = $false
    $buttonPanel.Visible = $true

    if ($treeView.SelectedNode -eq $null) {
        # Show splash image
        $splashImage.Visible = $true
        $buttonPanel.Visible = $false
        return
    }

    $selectedCategory = $treeView.SelectedNode.Text

    if ([string]::IsNullOrWhiteSpace($selectedCategory)) {
        return
    }

    if ($selectedCategory -eq "All Items") {
        $y = 10
        foreach ($category in $script:allMenuItems.Keys) {
            # Add category header
            $categoryLabel = New-Object System.Windows.Forms.Label
            $categoryLabel.Location = New-Object System.Drawing.Point(10, $y)
            $categoryLabel.Size = New-Object System.Drawing.Size(680, 25)
            $categoryLabel.Text = "> " + $category
            $categoryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $categoryLabel.BackColor = $script:categoryColors[$category]
            $categoryLabel.ForeColor = [System.Drawing.Color]::White
            $buttonPanel.Controls.Add($categoryLabel)
            $y += 35

            # Add buttons for items in this category
            foreach ($item in $script:allMenuItems[$category]) {
                $button = Create-Button $item $y 680
                $buttonPanel.Controls.Add($button)
                $y += 60

                # Add description label below the button
                $descriptionLabel = New-Object System.Windows.Forms.Label
                $descriptionLabel.Location = New-Object System.Drawing.Point(10, $y)
                $descriptionLabel.Size = New-Object System.Drawing.Size(680, 50)
                $descriptionLabel.Text = $item.Description
                $descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
                $descriptionLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
                $buttonPanel.Controls.Add($descriptionLabel)
                $y += 55
            }
            $y += 15
        }
    }
    else {
        $y = 10

        # Check if category exists in dictionary
        if (-Not $script:allMenuItems.ContainsKey($selectedCategory)) {
            return
        }

        # Add category header
        $categoryLabel = New-Object System.Windows.Forms.Label
        $categoryLabel.Location = New-Object System.Drawing.Point(10, $y)
        $categoryLabel.Size = New-Object System.Drawing.Size(680, 25)
        $categoryLabel.Text = "> " + $selectedCategory
        $categoryLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $categoryLabel.BackColor = $script:categoryColors[$selectedCategory]
        $categoryLabel.ForeColor = [System.Drawing.Color]::White
        $buttonPanel.Controls.Add($categoryLabel)
        $y += 35

        # Add buttons for items in this category
        foreach ($item in $script:allMenuItems[$selectedCategory]) {
            $button = Create-Button $item $y 680
            $buttonPanel.Controls.Add($button)
            $y += 60

            # Add description label below the button
            $descriptionLabel = New-Object System.Windows.Forms.Label
            $descriptionLabel.Location = New-Object System.Drawing.Point(10, $y)
            $descriptionLabel.Size = New-Object System.Drawing.Size(680, 50)
            $descriptionLabel.Text = $item.Description
            $descriptionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
            $descriptionLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
            $buttonPanel.Controls.Add($descriptionLabel)
            $y += 55
        }
    }

    Apply-Filter
}

# Function to check if Inventor Add-In is installed
function Test-InventorAddInInstalled {
    try {
        $addinName = "AssemblyClonerAddIn"
        $addinDll = "$addinName.dll"
        $addinManifest = "$addinName.addin"

        # Possible installation paths for different Inventor versions
        $addinPaths = @(
            "$env:APPDATA\Autodesk\Inventor Addins",
            "$env:ProgramData\Autodesk\Inventor 2026\Addins",
            "$env:ProgramData\Autodesk\Inventor 2025\Addins",
            "$env:ProgramData\Autodesk\Inventor 2024\Addins",
            "$env:ProgramData\Autodesk\Inventor 2023\Addins"
        )

        foreach ($path in $addinPaths) {
            if (Test-Path $path -ErrorAction SilentlyContinue) {
                if (Test-Path "$path\$addinDll" -ErrorAction SilentlyContinue -and Test-Path "$path\$addinManifest" -ErrorAction SilentlyContinue) {
                    return $true
                }
            }
        }

        return $false
    }
    catch {
        return $false
    }
}

# Create professional button
function Create-Button {
    param($item, $y, $width)

    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(10, $y)
    $button.Size = New-Object System.Drawing.Size($width, 55)
    $button.Text = $item.Name
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $button.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $button.ForeColor = [System.Drawing.Color]::FromArgb(0, 0, 0)
    $button.FlatStyle = "Flat"
    $button.FlatAppearance.BorderSize = 1
    $button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $button.UseVisualStyleBackColor = $false

    # Tooltip
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($button, $item.Description)

    # Check if add-in installation status should be checked and displayed
    if ($item.ContainsKey('CheckInstalled') -and $item.CheckInstalled -eq $true) {
        $isInstalled = Test-InventorAddInInstalled
        if ($isInstalled) {
            $button.Text = "$($item.Name) - Installed"
            $button.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
            $tooltip.SetToolTip($button, "$($item.Description) - Currently installed")
            $button.BackColor = [System.Drawing.Color]::FromArgb(230, 255, 230)
        }
        else {
            $button.Text = "$($item.Name) (Click to Install)"
            $button.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
            $tooltip.SetToolTip($button, "$($item.Description) - Not installed")
            $button.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240)
        }
    }

    # Click handler
    $button.Add_Click({
        param($sender, $e)
        $scriptPath = $sender.Tag
        $scriptName = $sender.Text

        # Save working directory
        $originalDir = Get-Location

        try {
            Update-Status "Running: $scriptName..."
            $logTextBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] Starting: $scriptName`n")
            $statusBar.Text = "Running: $scriptName..."
            $progressBar.Value = 0
            $progressBar.Visible = $true
            $progressBar.Maximum = 100
            $progressBar.Value = 50
            $statusBar.Refresh()

            # Change to script directory and run
            $scriptDir = Split-Path (Join-Path $PSScriptRoot $scriptPath) -Parent
            if (Test-Path $scriptDir) {
                Set-Location $scriptDir
                $scriptFile = Split-Path $scriptPath -Leaf

                # Handle VBS files differently - use cscript
                if ($scriptFile -match '\.vbs$') {
                    $process = Start-Process cscript.exe -ArgumentList "//nologo `"$scriptFile`"" -Wait -PassThru
                }
                else {
                    $process = Start-Process cmd.exe -ArgumentList "/c $scriptFile" -Wait -PassThru
                }

                if ($process.ExitCode -eq 0) {
                    $logTextBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $scriptName completed successfully`n")
                    Update-Status "Ready"
                    $progressBar.Value = 100
                    Start-Sleep -Milliseconds 500
                }
                else {
                    $logTextBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] $scriptName completed with errors (Exit Code: $($process.ExitCode))`n")
                    Update-Status "Completed with errors"
                    $progressBar.BackColor = [System.Drawing.Color]::FromArgb(255, 200, 200)
                }
            }
            else {
                $logTextBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] ERROR: Script directory not found: $scriptDir`n")
                Update-Status "Error"
                [System.Windows.Forms.MessageBox]::Show("Script directory not found:`n$scriptDir", "Error", "OK", "Error")
            }
        }
        catch {
            $logTextBox.AppendText("[$(Get-Date -Format 'HH:mm:ss')] ERROR: $($_.Exception.Message)`n")
            Update-Status "Error"
            [System.Windows.Forms.MessageBox]::Show("An error occurred:`n$($_.Exception.Message)", "Error", "OK", "Error")
        }
        finally {
            # Restore working directory
            Set-Location $originalDir
            $progressBar.Visible = $false
            $progressBar.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        }
    })

    # Script path stored in Tag
    $button.Tag = $item.Script

    return $button
}

# Filter buttons based on search
function Filter-Buttons {
    Apply-Filter
}

function Apply-Filter {
    $searchText = $searchBox.Text.ToLower()

    foreach ($control in $buttonPanel.Controls) {
        if ($control -is [System.Windows.Forms.Button]) {
            if ($searchText -and -not $control.Text.ToLower().Contains($searchText)) {
                $control.Visible = $false
            }
            else {
                $control.Visible = $true
            }
        }
        elseif ($control -is [System.Windows.Forms.Label]) {
            $hasVisibleButtons = $false
            foreach ($btn in $buttonPanel.Controls) {
                if ($btn -is [System.Windows.Forms.Button] -and $btn.Location.Y -gt $control.Location.Y -and $btn.Location.Y -lt ($control.Location.Y + 100) -and $btn.Visible) {
                    $hasVisibleButtons = $true
                    break
                }
            }
            if ($searchText) {
                $control.Visible = $hasVisibleButtons
            }
            else {
                $control.Visible = $true
            }
        }
    }
}

# Update status
function Update-Status {
    param($status)
    $statusBar.Text = $status
}

# Show About dialog
function Show-About {
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About Inventor Automation Suite"
    $aboutForm.Size = New-Object System.Drawing.Size(450, 300)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(400, 200)
    $label.Text = "Inventor Automation Suite`n`nVersion: 1.0.0`n`nA comprehensive tool suite for Autodesk Inventor automation`n`nFeatures include:`n- Part and assembly renaming`n- IDW reference updates`n- Title automation`n- Cloning tools`n- Sheet metal conversion`n- iLogic analysis`n- And much more...`n`n(c) 2025 Spectiv"
    $aboutForm.Controls.Add($label)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(175, 230)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = "OK"
    $aboutForm.Controls.Add($okButton)

    $aboutForm.ShowDialog()
}

# Initialize
Show-CategoryButtons
Update-Status "Ready"

# Show form
$MainForm.ShowDialog()
