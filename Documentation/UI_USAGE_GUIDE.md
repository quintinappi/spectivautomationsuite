# Inventor Automation Suite - Professional PowerShell UI

## Installation & Usage

### Quick Start

1. **Run the UI:**
   - Double-click `Launch_UI.bat` - this will launch the professional PowerShell interface
   - Or right-click `Launch_UI.ps1` and select "Run with PowerShell"

2. **Using the Launcher:**
   - The UI displays all automation tools organized by category
   - Use the **left sidebar** to navigate between categories
   - Use the **search bar** to filter tools by name
   - Click any button to run the corresponding automation script
   - Watch the **status bar** and **log window** for progress and results

### Features

- **Smart Search:** Type in the search box to filter visible buttons
- **Category Navigation:** Click categories in the tree view to show only tools in that group
- **"All Items" View:** Shows all tools from all categories
- **Progress Tracking:** Status bar shows which script is running
- **Log Window:** Timestamped log of all operations
- **Error Handling:** Clear error messages if scripts fail

### Menu Items

The launcher contains all the same tools as the original MAIN_LAUNCHER.bat:

**Core Production Workflow:**
- Part Renaming
- IDW Updates
- Title Automation

**Management & Utilities:**
- Registry Management
- File Utilities

**Rescue & Synchronization:**
- Smart Prefix Scanner
- Emergency IDW Fixer
- IDW-Assembly Synchronizer

**Cloning Tools:**
- Assembly Cloner
- Part Cloner

**iLogic & Analysis:**
- iLogic Scanner
- Find Missing Detailed Parts

**Sheet Metal Conversion:**
- Sheet Metal Converter (Assembly)
- Sheet Metal Converter (Part)

**Drawing Customization:**
- Change Balloon Style
- Change Dimension Style
- Export IDW Sheets to PDF

**Parts List and BOM:**
- Create Sheet Parts List

**Parameter Management:**
- Length Parameter Exporter
- Fix Non-Plate Parts
- Fix Single Part Length2

### Troubleshooting

**UI doesn't open:**
- Right-click `Launch_UI.bat` and select "Run as Administrator"
- Check that PowerShell execution policy allows scripts: Run `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

**Scripts don't run:**
- Make sure you're running from the main project directory
- Check that the subdirectories (Part_Renaming, IDW_Updates, etc.) exist
- Verify that the .bat files exist in their respective folders

**Small window size:**
- The window is resizable - drag the corner to enlarge it
- The form has a minimum size to ensure all controls are visible

### Compatibility

- **Windows 7+** (PowerShell 2.0+)
- **Windows 10/11 recommended** (PowerShell 5.1+)
- .NET Framework 3.0+ (standard on Windows 7+)
- No additional software required

### Technical Notes

The PowerShell UI is a wrapper that:
- Executes the existing BAT files
- Maintains the same functionality as the text-based launcher
- Adds a professional, user-friendly interface
- Provides better visual organization of tools

All automation logic remains in the existing scripts - this is just a modern UI frontend.

### Support

For issues with specific automation tools, refer to the documentation in each tool's folder.
For UI issues, check the PowerShell error messages in the log window.
