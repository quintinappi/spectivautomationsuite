# FLAT PATTERN ORIENTATION FIX - TECHNICAL DOCUMENTATION

**Date:** January 6, 2026  
**Issue:** Sheet metal flat patterns displaying edge view (6mm thickness) instead of top/bottom face  
**Status:** RESOLVED

---

## PROBLEM DESCRIPTION

When converting standard parts to sheet metal using the Inventor API, the resulting flat pattern was oriented incorrectly:

- **Expected:** Flat pattern showing the large plate face (e.g., 2500mm x 4700mm)
- **Actual:** Flat pattern showing the thin edge (e.g., 3850mm x 6mm)

This occurred specifically with:
- Derived parts (parts derived from "DM Underpan Development for plates.ipt")
- Parts converted from standard to sheet metal via API
- The API was selecting Face #1 (23,100 mm²) instead of Face #5 (6,174,892 mm²)

---

## ROOT CAUSE ANALYSIS

### API Limitations Discovered:

1. **BaseFace Property is READ-ONLY**
   - `FlatPattern.BaseFace` cannot be set programmatically
   - Attempting to set it returns "Object doesn't support this property or method"

2. **ConvertToSheetMetalFeatures API Not Available**
   - `Features.ConvertToSheetMetalFeatures` does not exist on `PartFeatures` collection
   - Only available on `SheetMetalFeatures` (when part is already sheet metal)
   - `CreateDefinition()` returns "Empty" type on derived parts

3. **Unfold() Does Not Accept Face Parameter**
   - `ComponentDefinition.Unfold(face)` gives "Wrong number of arguments"
   - Only `Unfold()` with no parameters works
   - Always uses the default base face (often the wrong one)

4. **FlipBaseFace Method Fails Silently**
   - `FlatPattern.FlipBaseFace` executes but does nothing
   - No error returned, but orientation remains unchanged

5. **Pre-Selection Ignored by Command**
   - Adding face to `SelectSet` before executing command does not affect face selection
   - The "Convert to Sheet Metal" command clears the selection and waits for user input

### Tested and Failed Approaches:

- Setting `BaseFace` property directly → Read-only
- Using `CreateDefinition(face)` → Returns Empty type
- Using `Unfold(face)` → Wrong parameter count
- Using `FlipBaseFace` → No effect
- Mouse click simulation (PowerShell) → Unreliable, position-dependent
- NameValueMap parameters to command → Type mismatch error

---

## SOLUTION

### Working Method: Pre-Selection + SendKeys Automation

The solution combines face pre-selection with keyboard automation to simulate user confirmation:

1. **Revert if Already Sheet Metal**
   - Delete existing flat pattern
   - Execute "PartConvertToStandardPartCmd" command
   - Ensures clean starting state

2. **Find Largest Face**
   ```vbscript
   Set faces = compDef.SurfaceBodies.Item(1).Faces
   For Each face In faces
       area = face.Evaluator.Area * 100 ' Convert to mm²
       If area > largestArea Then
           largestArea = area
           Set largestFace = face
       End If
   Next
   ```

3. **Pre-Select the Largest Face**
   ```vbscript
   Set selectSet = partDoc.SelectSet
   selectSet.Clear
   selectSet.Select largestFace
   ```

4. **Execute Convert Command**
   ```vbscript
   Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
   convertCmd.Execute
   ```

5. **Automate User Confirmation with SendKeys**
   ```vbscript
   Set WshShell = CreateObject("WScript.Shell")
   WshShell.AppActivate "Autodesk Inventor"
   WScript.Sleep 1000
   WshShell.SendKeys "{ENTER}"  ' Confirm face selection
   WScript.Sleep 500
   WshShell.SendKeys "{ENTER}"  ' Accept sheet metal defaults
   ```

6. **Create Flat Pattern**
   ```vbscript
   compDef.Unfold
   ```

### Key Timing Considerations:

- 500ms after activating part document
- 1000ms between command execution and first Enter
- 500ms between first and second Enter
- These delays ensure Inventor UI has time to display dialogs

---

## RESULTS

### Before Fix:
- Flat Pattern: **3,850mm × 6mm** (edge view)
- BaseFace Area: **23,100 mm²** (Face #1 - edge)
- Status: ❌ WRONG ORIENTATION

### After Fix:
- Flat Pattern: **4,727.9mm × 2,500.7mm** (top view)
- BaseFace Area: **6,174,892 mm²** (Face #5 - large face)
- Status: ✅ CORRECT ORIENTATION

---

## IMPLEMENTATION FILES

### Updated Files:
1. **Sheet_Metal_Converter.vbs** - Main conversion script
   - Updated `ConvertPartToSheetMetal()` function
   - Added largest face detection logic
   - Integrated SendKeys automation

### Test Files Created:
1. **TEST_SendKeys_Convert.vbs** - Proof of concept test
2. **TEST_Verify_Dimensions.vbs** - Dimension verification
3. **TEST_New_Convert_Single_Part.vbs** - Standalone single-part test
4. **Launch_Test_Single_Part_New.bat** - Test launcher

---

## VALIDATION CHECKLIST

✅ Revert existing sheet metal parts to standard  
✅ Detect largest face by area calculation  
✅ Pre-select face before command execution  
✅ Execute convert command successfully  
✅ Automate face confirmation with SendKeys  
✅ Automate defaults dialog acceptance  
✅ Create flat pattern after conversion  
✅ Verify flat pattern dimensions > 100mm on both axes  
✅ Confirm BaseFace is the large face (> 6 million mm²)  

---

## KNOWN LIMITATIONS

1. **Requires User Presence**
   - SendKeys requires Inventor window to be active
   - Cannot run completely headless/unattended
   - User must not interact with keyboard during execution

2. **Timing Dependent**
   - Fixed sleep delays may need adjustment on slower systems
   - Dialog appearance timing varies by system performance

3. **Window Focus Required**
   - AppActivate must successfully bring Inventor to foreground
   - May fail if other applications force focus

4. **Current Scope: Single Part Processing**
   - Solution validated for individual part conversion
   - Assembly batch processing requires additional integration

---

## NEXT STEPS

1. ✅ Single part conversion with correct orientation - **COMPLETE**
2. ⏳ Assembly-level batch processing - **PENDING**
   - Process all plate parts in assembly sequentially
   - Handle timing for multiple parts
   - Update assembly parameters (PLATE LENGTH, PLATE WIDTH)
3. ⏳ Error handling and recovery - **PENDING**
   - Handle conversion failures gracefully
   - Log parts requiring manual intervention
   - Generate summary report

---

## TECHNICAL NOTES

### Why SendKeys Works When API Doesn't:

The "Convert to Sheet Metal" command internally:
1. Waits for user to select a face (interactive mode)
2. Highlights the selected face
3. Creates the sheet metal definition using that face
4. **This face selection happens DURING command execution, not before**

The API provides no way to:
- Pass the face as a parameter to the command
- Programmatically confirm the selection while command is active
- Create the ConvertToSheetMetal definition directly with a specific face

SendKeys simulates the user pressing Enter, which:
- Confirms the face that's currently in SelectSet (if any)
- Or confirms the default face if SelectSet is empty
- Triggers the command to proceed to the defaults dialog

### Inventor API Gap:

This reveals a gap in the Inventor API where:
- The command-line approach requires user interaction
- The feature-creation API (CreateDefinition/Add) is not functional for derived parts
- No programmatic bridge exists between the two

Our solution uses the only available middle ground: UI automation via SendKeys.

---

**Author:** Development Team  
**Last Updated:** January 6, 2026
