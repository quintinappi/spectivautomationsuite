# Part Placer - Add-In Feature Guide

## Overview

The **Part Placer** is a new feature in the AssemblyClonerAddIn that automatically scans an assembly for parts containing specific material designations ("PL" or "S355JR") and places them as base views in a new IDW drawing at 1:1 scale.

## Purpose

This tool was created to replace a problematic VBScript that had issues with:
- Unreliable view placement
- Limited error handling
- No logging/debugging capability
- Difficult to troubleshoot

The Add-In approach provides:
- Full error handling and logging
- Direct Inventor API access
- Better UI integration
- Comprehensive debugging output

## How It Works

### 1. Assembly Selection
- User opens an assembly in Inventor
- Clicks the "Place Parts in IDW" button on the ribbon

### 2. Part Scanning
- Scans all occurrences in the assembly (including sub-assemblies)
- Identifies unique parts (doesn't duplicate if same part appears multiple times)
- Checks for matching criteria:
  - Part Number or Description contains "S355JR"
  - Part Number or Description contains "PL" (as material designation)

### 3. IDW Creation
- Prompts user for save location
- Creates new IDW from template
- Places each matching part as a base view at 1:1 scale
- Arranges views in a grid layout on the sheet

### 4. Logging
- Every step is logged to a timestamped file
- Log location: `%USERPROFILE%\Documents\InventorAutomationSuite\Logs\`

## Usage

### Prerequisites
1. Inventor 2026 must be installed
2. AssemblyClonerAddIn must be deployed
3. An assembly (.iam) must be open

### Steps
1. **Open an assembly** in Inventor
2. **Navigate to the Tools tab** on the ribbon
3. **Look for the "Cloner Tools" panel**
4. **Click the "Place Parts in IDW" button** (icon with "V")
5. **Select save location** for the new IDW file
6. **Wait for processing** - a log file will be created
7. **Review the created IDW** with all matching parts

### Search Criteria

Parts are matched if their Part Number OR Description contains:

| Term | Match Reason |
|------|-------------|
| `S355JR` | Steel grade designation |
| `PL` | Plate material (e.g., "PL 10", "PLATE") |

Examples of matches:
- Part Number: "PLATE_100x100", Description: "Steel plate" ✓
- Part Number: "BRACKET_S355JR", Description: "" ✓
- Part Number: "BEAM", Description: "S355JR steel beam" ✓
- Part Number: "PL 10x200", Description: "Plate 10mm" ✓

Examples of non-matches:
- Part Number: "SAMPLE", Description: "" ✗
- Part Number: "APPLY", Description: "" ✗ (PL part of word)

## Technical Details

### File Structure
```
InventorAddIn/
├── AssemblyClonerAddIn/
│   ├── PartPlacer.vb          # Main module
│   └── StandardAddInServer.vb # Updated with button
├── DEPLOY_NOW.bat             # Updated deployment script
└── PART_PLACER_GUIDE.md       # This file
```

### Key Classes

#### `PartPlacer`
Main class that orchestrates the operation:
- `Execute()` - Main entry point
- `ScanAssemblyForMatchingParts()` - Scans for matching parts
- `CreateIDWWithParts()` - Creates IDW and places views

#### `PartOccurrenceInfo`
Data class for part information:
- `FilePath` - Full path to part file
- `PartNumber` - Stock Number from iProperties
- `Description` - Description from iProperties
- `MatchReason` - Why this part was matched

### Logging System

Every operation is logged with timestamps:

```
================================================================================
PART PLACER - LOG FILE
================================================================================
Start Time: 2025-02-04 10:30:15
Inventor Version: Autodesk Inventor 2026
Log File: C:\Users\...\Documents\InventorAutomationSuite\Logs\PartPlacer_20250204_103015.log
================================================================================

[10:30:15.123] ----------------------------------------------------------
[10:30:15.123] SECTION: STEP 1: VALIDATE ACTIVE DOCUMENT
[10:30:15.123] ----------------------------------------------------------
[10:30:15.234] Active document type: kAssemblyDocumentObject
[10:30:15.234] Assembly file: C:\...\MyAssembly.iam
[10:30:15.234] VALIDATION PASSED: Assembly document is active
```

### Error Handling

Errors are caught and logged with:
- Context of the error
- Full exception message
- Stack trace
- User-friendly message box

## Troubleshooting

### "No matching parts found"
- Check that parts have "PL" or "S355JR" in Part Number or Description
- Verify iProperties are properly set
- Check the log file for which parts were scanned

### "Could not create IDW"
- Verify you have write permissions to the selected folder
- Check that drawing templates exist in the Templates folder
- Review log for specific error details

### Views not placed correctly
- Check log for each part placement attempt
- Verify part files exist and are accessible
- Check if parts are in a valid state (not corrupted)

### Add-in button not visible
1. Close Inventor
2. Run `DEPLOY_NOW.bat` as Administrator
3. Restart Inventor
4. Check Tools -> Add-Ins for "Assembly Cloner with iLogic Patcher"

## Log File Location

```
%USERPROFILE%\Documents\InventorAutomationSuite\Logs\
```

Example:
```
C:\Users\Quintin\Documents\InventorAutomationSuite\Logs\PartPlacer_20250204_103015.log
```

## Deployment

### Build Requirements
- Visual Studio 2022 (or later)
- .NET Framework 4.8
- Inventor 2026 Interop assemblies

### Build Steps
1. Open `InventorAddIn\AssemblyClonerAddIn.sln`
2. Set Configuration: `Release`
3. Set Platform: `x64`
4. Build -> Build Solution

### Deploy Steps
1. Right-click `InventorAddIn\DEPLOY_NOW.bat`
2. Select "Run as administrator"
3. Wait for deployment confirmation
4. Close and restart Inventor

## Future Enhancements

Possible improvements:
- [ ] Configurable search terms (user-defined patterns)
- [ ] Configurable scale factor
- [ ] Multiple sheet support for large assemblies
- [ ] Automatic view alignment and spacing
- [ ] Part list/balloon annotation
- [ ] Support for additional view orientations

## Support

For issues or questions:
1. Check the log file first
2. Review this guide
3. Check Windows Event Viewer for .NET errors
4. Verify Inventor and .NET Framework versions

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2025-02-04 | Initial release |
