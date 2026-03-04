# View Style Manager

## Overview
This tool helps you manage drawing view styles in Autodesk Inventor IDW files. It's particularly useful when you've copied views from other IDW files that carry over their original styles.

## Problem It Solves
When you copy a view from one IDW to another, the view often retains the style from the source drawing. This can cause inconsistencies in your drawings. This tool allows you to:
- Scan all views to see what styles are currently applied
- Change views from one style to another
- Batch change all views to a consistent style

## Features

### 1. Scan View Styles
- Lists all available styles in the current drawing
- Shows which style is applied to each view on each sheet
- Generates a detailed report saved to a text file
- Displays view type (Standard, Projected, Section, Detail, etc.)

### 2. Change View Styles
- Change specific views from one style to another
- Change ALL views to a single style
- Interactive style selection (by number or name)
- Automatic document saving after changes

## How to Use

### Prerequisites
1. Autodesk Inventor must be running
2. Open the IDW file you want to work with
3. Make sure you have backup copies (always recommended)

### Running the Tool

#### Method 1: Double-click the Batch File
1. Navigate to the `View_Style_Manager` folder
2. Double-click `Launch_View_Style_Manager.bat`
3. Follow the on-screen prompts

#### Method 2: Run from Command Line
```batch
cd "c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\View_Style_Manager"
cscript //nologo View_Style_Manager.vbs
```

### Workflow

#### Scanning View Styles
1. Run the tool
2. Select "YES" when asked what you want to do
3. Review the on-screen report
4. Optionally open the detailed text report file

#### Changing View Styles
1. Run the tool
2. Select "NO" when asked what you want to do
3. You'll see a list of available styles
4. Choose the source style (or leave blank to change ALL views)
5. Choose the target style
6. Confirm the change
7. The tool will update all matching views and save the document

## Example Scenarios

### Scenario 1: Fix Copied Views
You copied views from another IDW and they're using "PEHD25 A1-3RD ANGLE" but you want them to use "Default Standard (ANSI)":

1. Run the tool → Select "NO" (Change styles)
2. Enter "PEHD25 A1-3RD ANGLE" as the source style
3. Enter "Default Standard (ANSI)" as the target style
4. Confirm → Done!

### Scenario 2: Standardize All Views
You want all views in the drawing to use the same style:

1. Run the tool → Select "NO" (Change styles)
2. Leave the source style BLANK (to change all views)
3. Enter your desired target style
4. Confirm → All views updated!

### Scenario 3: Audit Your Drawing
You want to see what styles are being used:

1. Run the tool → Select "YES" (Scan styles)
2. Review the report
3. Optionally open the saved text file for detailed analysis

## Understanding View Styles

In Inventor, view styles control:
- Line weights and types
- Hidden line display
- Dimension appearance
- Annotation standards
- Drawing standards (ANSI, ISO, etc.)

Common style names you might see:
- Default Standard (ANSI)
- Default Standard (ISO)
- PEHD25 A1-3RD ANGLE
- PENTALIN 25
- RDE
- DESIGN SOLVE (ISO)

## Log Files

The tool creates log files in the same directory:
- `ViewStyleManager_[timestamp].log` - Detailed operation log
- `ViewStyleReport_[timestamp].txt` - Scan results (when scanning)

## Troubleshooting

### "No IDW/DWG file is open!"
- Make sure you have an Inventor drawing file open before running the tool

### "No styles found in this drawing!"
- This is unusual - check that your drawing file is valid
- Try opening a different drawing

### "Invalid style selection!"
- Make sure you enter either a valid number or exact style name
- Style names are case-sensitive

### Changes don't appear
- Make sure you saved the document (tool does this automatically)
- Try closing and reopening the drawing
- Check the log file for errors

## Technical Details

### Supported View Types
- Standard views
- Projected views
- Auxiliary views
- Section views
- Detail views
- Drafting views
- Overlay views

### API Used
- `DrawingDocument.StylesManager.DrawingViewStyles` - Access available styles
- `DrawingView.Style` - Get/set view style
- `DrawingView.ViewType` - Determine view type

### Compatibility
- Autodesk Inventor 2015 and later
- Windows 7/8/10/11
- VBScript (built into Windows)

## Safety Features
- Read-only scanning mode available
- Confirmation required before making changes
- Automatic document saving after changes
- Detailed logging of all operations
- Skips views that can't be changed

## Integration with Other Scripts

This tool follows the same pattern as other scripts in the FINAL_PRODUCTION_SCRIPTS collection:
- Similar logging format
- Compatible batch file launcher
- Follows Inventor API best practices
- Error handling and recovery

## Future Enhancements

Possible future additions:
- Batch process multiple IDW files
- Import/export style definitions
- Style comparison between drawings
- Automatic style standardization rules

## Support

For issues or questions:
1. Check the log file for detailed error messages
2. Review the INVENTOR_API_REFERENCE.md in the parent directory
3. Consult Autodesk Inventor API documentation

## Version History

### Version 1.0 (2026-01-09)
- Initial release
- Scan view styles functionality
- Change view styles functionality
- Interactive menu system
- Detailed logging and reporting
