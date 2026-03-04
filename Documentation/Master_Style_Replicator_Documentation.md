# Master Style Replicator - Complete Documentation

## Overview
The **Master Style Replicator** is an interactive tool that copies view styling (hidden/visible line formatting, **center lines**, and **center marks**) from a user-specified Master View to Target Views in Autodesk Inventor drawings.

## What's New - Version 2.0
🎉 **NOW INCLUDES CENTER LINES AND CENTER MARKS REPLICATION!**

The tool now replicates:
- ✅ Visible line styling
- ✅ Hidden line styling  
- ✅ **Center Lines** (layer assignment)
- ✅ **Center Marks** (layer assignment)

## Problem Solved
When copying views between sheets in Inventor:
- Hidden lines often lose their proper styling (cyan color) and appear as regular visible lines (black)
- **Center lines and center marks may end up on incorrect layers**
- Annotation styling becomes inconsistent across views

This tool replicates the correct styling from a properly-formatted Master View to other views that need the same formatting.

## How It Works

### Key Discovery: Layer Name vs Color Detection
**Critical technical breakthrough:** The script identifies hidden/visible lines by **LAYER NAME**, not color.

**Why this matters:**
- Hidden lines on layers named "Hidden" may have BLACK color (incorrect styling)
- Script checks `InStr(UCase(layerName), "HIDDEN")` to detect hidden layers
- This works regardless of current layer color settings
- Master view layers (PEN25 Hidden, PEN25 Visible) are identified by name and then applied to target views

### Center Lines & Center Marks Detection
The script now analyzes:
- `DrawingView.Centerlines` collection - for all center line annotations
- `DrawingView.CenterMarks` collection - for all center mark annotations

It identifies the layer used for center lines/center marks in the Master View and applies that same layer to target views.

### Process Flow
1. **Scan all views** in the drawing and display them
2. **User selects Master View** (the view with correct styling)
3. **Analyze Master View** to identify:
   - Hidden Layer (e.g., "PEN25 Hidden (ISO)")
   - Visible Layer (e.g., "PEN25 Visible (ISO)")
   - **Center Line Layer** (e.g., "PEN25 Centerline (ISO)")
   - **Center Mark Layer** (e.g., "PEN25 Centerline (ISO)")
   - Count of hidden and visible curves
   - **Count of center lines and center marks**
4. **User selects Target Views** (ALL, specific, or pattern-matched)
5. **Analyze each Target View** to show:
   - Total curves found
   - Line breakdown by layer name and color
   - **Center line and center mark counts**
   - Sample conversions (first 3 changes)
6. **Apply styling** to each target view:
   - Lines on "Hidden" layers → Move to Master Hidden Layer
   - Lines on "Visible" layers → Move to Master Visible Layer
   - **Center lines → Move to Master Center Line Layer**
   - **Center marks → Move to Master Center Mark Layer**
7. **Save document** with updated styling

## Usage Instructions

### Prerequisites
- Autodesk Inventor must be **running**
- An **IDW drawing file** must be **open**

### Running the Tool
```batch
Launch_Master_Style_Replicator.bat
```

### Step-by-Step Workflow

**Step 1: Review Available Views**
The script displays all views in your drawing:
```
==========================================
  AVAILABLE VIEWS IN DRAWING
==========================================

SHEET: Sheet:1
  - 1
  - ELEVATION

Found 2 view(s) in drawing
```

**Step 2: Select Master View**
Enter the name of the view that has **correct styling** (proper cyan hidden lines, correct center line layer):
```
Enter the NAME of the Master View to copy style FROM:
  [1 or ELEVATION]
```

**Step 3: Analyze Master View**
Script analyzes the Master View and reports:
```
==========================================
  MASTER VIEW STYLE ANALYSIS
==========================================
View: 1

Geometry Curves found: 759

Line breakdown:
  PEN25 Hidden (ISO) [R0 G255 B255]: 469 curves
  PEN25 Visible (ISO) [R0 G0 B0]: 290 curves

--- Center Lines Analysis ---
  Found 5 center line(s) on layer: PEN25 Centerline (ISO)

--- Center Marks Analysis ---
  Found 3 center mark(s) on layer: PEN25 Centerline (ISO)

--- Identified Master Layers ---
  Visible Layer: PEN25 Visible (ISO)
  Hidden Layer: PEN25 Hidden (ISO)
  Center Line Layer: PEN25 Centerline (ISO)
  Center Mark Layer: PEN25 Centerline (ISO)
```

**Step 4: Select Target Views**
Choose how to select views to update:
```
Choose which views to apply Master View's style TO:
  1 - ALL views (except Master View)
  2 - Specific views (enter names separated by commas)
  3 - Views matching pattern (use wildcards like ELEVATION*, SECTION*)

NOTE: Style will include Center Lines & Center Marks!
```

**Step 5: View Application Results**
For each target view, script shows:
```
==========================================
  TARGET VIEW: ELEVATION
==========================================
Geometry curves found: 781
  [SAMPLE] Visible layer: Visible (ISO) -> PEN25 Visible (ISO)
  [SAMPLE] Visible layer: Visible (ISO) -> PEN25 Visible (ISO)
  [SAMPLE] Hidden layer: Hidden (ISO) -> PEN25 Hidden (ISO)
  -> Processed 345 visible lines
  -> Processed 500 hidden lines

  -> Processed 4 center line(s), 2 updated to layer: PEN25 Centerline (ISO)
  -> Processed 2 center mark(s), 2 updated to layer: PEN25 Centerline (ISO)
```

**Step 6: Verification**
Final summary:
```
==========================================
  SUMMARY
==========================================

Total views updated: 1
Total visible lines fixed: 345
Total hidden lines fixed: 500
Total center lines updated: 2
Total center marks updated: 2

Updating document...
Document updated successfully!

Saving document...
Document saved successfully!

==========================================
  COMPLETE
==========================================

Style from '1' has been applied to 1 view(s).

GEOMETRY STYLING:
  - Visible lines: 345 segments updated
  - Hidden lines: 500 segments updated

ANNOTATION STYLING:
  - Center lines: 2 updated
  - Center marks: 2 updated
```

## Technical Details

### Layer Detection Logic
```vbscript
' Check if current layer indicates hidden geometry (by layer NAME, not color)
Dim isHiddenLayer
isHiddenLayer = False
If InStr(UCase(curLayer.Name), "HIDDEN") > 0 Then
    isHiddenLayer = True
End If
```

### Center Lines Processing
```vbscript
' Process Center Lines
Set targetCenterLines = targetView.Centerlines
If targetCenterLines.Count > 0 And Not masterCenterLineLayer Is Nothing Then
    For Each cl In targetCenterLines
        If UCase(cl.Layer.Name) <> UCase(masterCenterLineLayer.Name) Then
            cl.Layer = masterCenterLineLayer
        End If
    Next
End If
```

### Center Marks Processing
```vbscript
' Process Center Marks
Set targetCenterMarks = targetView.CenterMarks
If targetCenterMarks.Count > 0 And Not masterCenterMarkLayer Is Nothing Then
    For Each cm In targetCenterMarks
        If UCase(cm.Layer.Name) <> UCase(masterCenterMarkLayer.Name) Then
            cm.Layer = masterCenterMarkLayer
        End If
    Next
End If
```

### Layer Assignment
```vbscript
' If currently on a Hidden layer --> Move to Master Hidden Layer
If isHiddenLayer Then
    If Not masterHidLayer Is Nothing Then
        s.Layer = masterHidLayer
        countViewHid = countViewHid + 1
    End If
Else
    ' If currently on a Visible layer --> Move to Master Visible Layer
    If Not masterVisLayer Is Nothing Then
        s.Layer = masterVisLayer
        countViewVis = countViewVis + 1
    End If
End If
```

### Key Features
1. **Layer name detection** - Works regardless of current layer colors
2. **Center line/center mark replication** - Now includes annotation styling
3. **Detailed analysis** - Shows line breakdown for both master and target views
4. **Sample logging** - First 3 conversions shown for verification
5. **Multiple selection modes** - ALL, specific views, or pattern matching
6. **Comprehensive reporting** - Full summary of changes made including annotations

## File Locations
- **Script:** `Documentation\Master_Style_Replicator.vbs`
- **Launcher:** `View_Style_Manager\Launch_Master_Style_Replicator.bat`
- **Category:** View Management in SpectivLauncher.exe

## Troubleshooting

### Issue: "0 hidden lines fixed"
**Cause:** Target view has no layers with "Hidden" in the name
**Solution:** Check layer names in target view. Ensure hidden geometry is on layers containing "Hidden" in the name

### Issue: "No center lines found in Master View"
**Cause:** Master view doesn't have any center lines to use as reference
**Solution:** Add center lines to your Master View first, or choose a different Master View that has center lines

### Issue: "Script doesn't find the view"
**Cause:** View name doesn't match exactly (case-sensitive)
**Solution:** Enter view name exactly as shown in the available views list

### Issue: Hidden lines still black after running
**Cause:** Master View's hidden layer doesn't have "Hidden" in the name
**Solution:** Choose a Master View where hidden lines are on layers named "Hidden" or verify layer naming convention

### Issue: Center lines not being updated
**Cause:** No center lines in Master View, or they're on a layer without "Center" in the name
**Solution:** Ensure your Master View has center lines on a layer with "Center" or "Centerline" in the name

## Example Use Cases

### Use Case 1: Fixing Copied Views
You copied a view from another sheet and hidden lines appear as regular black lines.
1. Run Master Style Replicator
2. Select the original (correctly styled) view as Master
3. Select ALL views to update
4. All views now have proper cyan hidden lines

### Use Case 2: Standardizing Drawing Style
Your drawing has multiple views with inconsistent hidden line styling.
1. Identify one view with correct styling as Master
2. Select specific views that need fixing
3. Apply consistent styling across all selected views

### Use Case 3: Pattern-Based Updates
You have multiple "ELEVATION" views (ELEVATION A, ELEVATION B, ELEVATION LEFT) that all need fixing.
1. Select Master View (ELEVATION FRONT)
2. Choose option 3 (pattern matching)
3. Enter pattern: `ELEVATION*`
4. All elevation views updated automatically

### Use Case 4: Center Line Standardization
You have multiple detail views and want all center lines on the same layer.
1. Select a Master View with center lines on the correct layer
2. Select target views that have center lines on wrong layers
3. Run the tool - all center lines will be moved to the Master View's center line layer

## Version History
- **v1.0** - Initial version using color-based detection (R0 G255 B255 for hidden)
- **v1.1** - Critical fix: Changed to layer name detection ("Hidden" in layer name)
  - Fixes issue where hidden layers have black color
  - Makes tool work regardless of current layer color settings
  - Added detailed line breakdown analysis
  - Added sample logging for verification
- **v2.0** - Major enhancement: Added Center Lines and Center Marks replication
  - Analyzes `DrawingView.Centerlines` collection
  - Analyzes `DrawingView.CenterMarks` collection
  - Replicates center line/center mark layer assignment
  - Updated summary to show annotation counts

## Development Notes

### Critical Implementation Details
1. **All variables declared at script level** (Option Explicit compliance)
2. **No duplicate Dim statements** (prevents VBScript compilation errors)
3. **Layer detection by name** (not color) - most important technical decision
4. **Center line/center mark layer detection** - Uses the layer from the first center line/mark found
5. **Comprehensive error handling** with clear status messages
6. **Sample logging** (first 3 conversions) for user verification

### What Made This Work
The breakthrough was realizing that:
- Original approach: Check for cyan color (R0 G255 B255) to identify hidden lines
- Problem: Hidden layers can have any color (including black)
- Solution: Check for "Hidden" in layer name instead
- Result: Works regardless of layer color, just needs proper layer naming

For center lines/center marks:
- Access through `DrawingView.Centerlines` and `DrawingView.CenterMarks` collections
- Each has a `.Layer` property that can be read and set
- Apply the same layer-based approach for consistency

This aligns with Inventor's layer management philosophy - layers are named by purpose (Hidden, Visible, Centerline, etc.), and color/dash pattern are display properties that can be customized per drawing template.
