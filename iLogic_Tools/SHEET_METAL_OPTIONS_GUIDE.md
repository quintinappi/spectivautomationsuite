# SHEET METAL CONVERTER - TWO OPTIONS

**Date:** January 6, 2026  
**Status:** ✅ BOTH OPTIONS FULLY WORKING

**Option 13** - Single part conversion  
**Option 12** - Assembly batch conversion

---

## OPTION 13: Single Part Converter (WORKING)

**Purpose:** Convert a single open part to sheet metal with correct flat pattern orientation

**Location:** Main Menu → Option 13  
**Launcher:** `iLogic_Tools\Launch_Sheet_Metal_Part_Converter.bat`  
**Script:** `iLogic_Tools\Sheet_Metal_Part_Converter.vbs`

### How to Use:
1. Open a part (.ipt) in Inventor
2. Run Main_Launcher.bat
3. Select Option 13
4. Script automatically:
   - Finds the largest face (top/bottom, not 6mm edge)
   - Converts to sheet metal
   - Creates flat pattern with CORRECT orientation
   - Adds custom iProperties with formulas:
     - PLATE LENGTH = `=<SHEET METAL LENGTH>`
     - PLATE WIDTH = `=<SHEET METAL WIDTH>`
   - Saves the part

### Verified Working:
✅ Correct face selection (6,174,892 mm² not 23,100 mm²)  
✅ Flat pattern shows large face (4,727.9mm x 2,500.7mm)  
✅ Formula-based custom iProperties  
✅ Automatic save  

---

## OPTION 12: Assembly Batch Converter (WORKING ✅)

**Purpose:** Scan assembly and convert ALL plate parts in one batch operation

**Location:** Main Menu → Option 12  
**Launcher:** `iLogic_Tools\Launch_Sheet_Metal_Converter.bat`  
**Script:** `iLogic_Tools\Sheet_Metal_Converter.vbs`

### How to Use:
1. Open an assembly (.iam) in Inventor
2. Run Main_Launcher.bat
3. Select Option 12
4. Confirm the upfront warning showing total parts to process
5. **For each part that needs conversion:**
   - Script pre-selects the largest face (highlighted green)
   - MsgBox appears: "Click green face"
   - Click OK in MsgBox (face already selected)
   - Conversion completes automatically
6. Script will:
   - Scan assembly BOM for parts with "PL" or "S355JR" in description
   - Group parts by thickness (6mm, 10mm, 20mm, 25mm)
   - Close assembly before processing
   - Open each unique part VISIBLE in UI
   - Convert to sheet metal with correct orientation
   - Add formula-based custom properties
   - Save each part
   - Reopen assembly
   - Add PLATE LENGTH/WIDTH parameters to assembly

### Key Features:
- ✅ Automatic assembly scanning for plate parts
- ✅ Grouping by thickness for organized processing
- ✅ Parts open VISIBLE in UI (critical for Convert dialog)
- ✅ Semi-automated: pre-selects face, user clicks OK per part
- ✅ Custom iProperty formulas added automatically
- ✅ Assembly parameters created automatically
- ✅ Full error handling and logging

### Tested Successfully:
- 17 parts in DM Underpan assembly
- Multiple thickness groups (6mm, 10mm, 20mm, 25mm)
- Mixed standard and already-sheet-metal parts
- Parts already sheet metal: properties updated only
- Standard parts: full conversion with user interaction

---

## TECHNICAL DETAILS

### The Core Solution (SendKeys Method)

Both options use the same proven conversion technique:

```vbscript
' 1. Find largest face
For Each face In faces
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

' 2. Pre-select the face
selectSet.Clear
selectSet.Select largestFace

' 3. Execute convert command
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
convertCmd.Execute

' 4. Use SendKeys to confirm
Set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate "Autodesk Inventor"
WScript.Sleep 1000
WshShell.SendKeys "{ENTER}"  ' Confirm face
WScript.Sleep 500
WshShell.SendKeys "{ENTER}"  ' Accept defaults

' 5. Create flat pattern
compDef.Unfold

' 6. Add formulas
customPropSet.Add "=<SHEET METAL LENGTH>", "PLATE LENGTH"
customPropSet.Add "=<SHEET METAL WIDTH>", "PLATE WIDTH"
```

### Critical Implementation Details

**VISIBILITY FIX (Critical):**
```vbscript
' Parts MUST be opened with visible=True for Convert dialog to work
Set partDoc = m_InventorApp.Documents.Open(fullPath, True)  ' TRUE = visible!
```

**Assembly Workflow:**
1. Close assembly before processing (`asmDoc.Close True`)
2. Open each part independently and visibly
3. Process conversions with user interaction
4. Reopen assembly after all parts processed
5. Add parameters to assembly

### Key Differences Between Options

| Feature | Option 13 (Part) | Option 12 (Assembly) |
|---------|------------------|----------------------|
| Input | Single open part | Assembly with multiple parts |
| Processing | Immediate | Sequential batch |
| User Interaction | 1 click per part | 1 click per unconverted part |
| Automation Level | Semi-automated | Semi-automated |
| Assembly Parameters | No | Yes (PLATE LENGTH/WIDTH added) |
| User Interaction | Minimal | Hands-off after start |
| Use Case | Quick single conversion | Bulk processing |
| Timing | ~5 seconds | ~5 seconds per part |

---

## USAGE RECOMMENDATIONS

**Use Option 13 when:**
- Converting one part manually
- Testing/verifying conversion
- Part is not in an assembly context
- Quick one-off conversion needed

**Use Option 12 when:**
- Processing entire underpan assembly
- Converting 10+ plate parts at once
- Want consistent batch processing
- Need summary report of all conversions

---

## FILES STRUCTURE

```
iLogic_Tools/
├── Launch_Sheet_Metal_Part_Converter.bat    (Option 13 launcher)
├── Sheet_Metal_Part_Converter.vbs           (Option 13 script)
├── Launch_Sheet_Metal_Converter.bat         (Option 12 launcher)
├── Sheet_Metal_Converter.vbs                (Option 12 script)
├── Verify_Custom_Properties.vbs             (Verification tool)
└── FLAT_PATTERN_ORIENTATION_FIX.md         (Technical docs)
```

---

**Last Updated:** January 6, 2026
