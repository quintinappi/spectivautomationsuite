' SEMI_AUTO_Sheet_Metal_Convert.vbs
' Semi-automated conversion - handles everything except the face click
' User needs to click the large flat face once when prompted
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, selectSet, WshShell

WScript.Echo "==========================================="
WScript.Echo "  SEMI-AUTOMATED SHEET METAL CONVERSION"
WScript.Echo "==========================================="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set WshShell = CreateObject("WScript.Shell")
Set partDoc = invApp.ActiveDocument

If partDoc Is Nothing Then
    WScript.Echo "ERROR: No document open in Inventor"
    WScript.Quit
End If

Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Step 1: Find and highlight the largest face
WScript.Echo "=== STEP 1: IDENTIFYING LARGEST FACE ==="

Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 And area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"

' Check if already sheet metal with correct flat pattern
Dim isSheetMetal
isSheetMetal = (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")

If isSheetMetal And compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** ALREADY CONVERTED WITH CORRECT ORIENTATION ***"
        WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
        WScript.Echo ""
        WScript.Echo "=== DONE ==="
        WScript.Quit
    End If
End If

' Step 2: Zoom and orient view
WScript.Echo ""
WScript.Echo "=== STEP 2: ZOOMING TO FIT ==="

Dim view
Set view = invApp.ActiveView
view.Fit
WScript.Echo "View zoomed to fit"

' Step 3: Check if we need to run convert command
WScript.Echo ""
WScript.Echo "=== STEP 3: STARTING CONVERSION ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

If convertCmd Is Nothing Then
    WScript.Echo "Convert command not found"
ElseIf convertCmd.Enabled Then
    ' Pre-select the largest face
    selectSet.Clear
    selectSet.Select largestFace
    WScript.Echo "Largest face pre-selected (green highlight)"
    
    ' Execute convert command
    WScript.Echo ""
    WScript.Echo "Launching Convert to Sheet Metal..."
    convertCmd.Execute
    
    WScript.Echo ""
    WScript.Echo "================================================"
    WScript.Echo "  ACTION REQUIRED: Click the green face!"
    WScript.Echo "================================================"
    WScript.Echo ""
    WScript.Echo "1. The largest face should be highlighted green"
    WScript.Echo "2. CLICK on the green face to confirm selection"
    WScript.Echo "3. Click OK on the Sheet Metal Defaults dialog"
    WScript.Echo ""
    WScript.Echo "Waiting for user action..."
    
    ' Wait for user to complete
    WScript.Sleep 10000 ' Wait 10 seconds
    
    ' Check if conversion completed
    partDoc.Update
    Set compDef = partDoc.ComponentDefinition
    
    If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
        WScript.Echo "Sheet metal conversion detected!"
    Else
        WScript.Echo "Conversion may not be complete - please check Inventor"
    End If
Else
    WScript.Echo "Convert command disabled - part may already be sheet metal"
End If

' Step 4: Create flat pattern if not exists
WScript.Echo ""
WScript.Echo "=== STEP 4: CREATING FLAT PATTERN ==="

Set compDef = partDoc.ComponentDefinition

If Not compDef.HasFlatPattern Then
    WScript.Echo "Creating flat pattern..."
    compDef.Unfold
    partDoc.Update
    WScript.Echo "Flat pattern created"
Else
    WScript.Echo "Flat pattern already exists"
End If

' Step 5: Final check
WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="

Set compDef = partDoc.ComponentDefinition

If compDef.HasFlatPattern Then
    Set fp = compDef.FlatPattern
    WScript.Echo "Flat pattern dimensions:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** SUCCESS! ORIENTATION IS CORRECT! ***"
    Else
        WScript.Echo ""
        WScript.Echo "*** WARNING: May need manual orientation fix ***"
        WScript.Echo "Use: Edit Flat Pattern Definition > Select large face"
    End If
Else
    WScript.Echo "No flat pattern created"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
