' TEST_Full_Auto_Convert.vbs
' Full automated conversion: Convert to Sheet Metal with face selection + Create Flat Pattern
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, selectSet, WshShell

WScript.Echo "=== FULL AUTOMATED SHEET METAL CONVERSION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set WshShell = CreateObject("WScript.Shell")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

WScript.Echo "Part: " & partDoc.DisplayName

' Step 1: Find largest face
WScript.Echo ""
WScript.Echo "=== STEP 1: FIND LARGEST FACE ==="

Dim body, faces, face, largestFace, largestArea, faceNum
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    faceNum = faceNum + 1
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number <> 0 Then
        area = 0
        Err.Clear
    End If
    
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Step 2: Get the convert command
WScript.Echo ""
WScript.Echo "=== STEP 2: GET CONVERT COMMAND ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

Dim needsConvert
needsConvert = True

If convertCmd Is Nothing Then
    WScript.Echo "Convert command not found"
    needsConvert = False
ElseIf Not convertCmd.Enabled Then
    WScript.Echo "Convert command not enabled - part may already be sheet metal"
    needsConvert = False
Else
    WScript.Echo "Convert command found and enabled"
End If

If needsConvert Then
    ' Step 3: Pre-select the face
    WScript.Echo ""
    WScript.Echo "=== STEP 3: PRE-SELECT FACE ==="
    
    selectSet.Clear
    selectSet.Select largestFace
    
    If selectSet.Count > 0 Then
        WScript.Echo "Face pre-selected!"
    Else
        WScript.Echo "Face selection failed"
    End If
    
    ' Step 4: Execute convert command
    WScript.Echo ""
    WScript.Echo "=== STEP 4: EXECUTE CONVERT COMMAND ==="
    WScript.Echo "Executing PartConvertToSheetMetalCmd..."
    
    convertCmd.Execute
    WScript.Sleep 500
    
    ' Step 5: Confirm face selection with simulated click
    WScript.Echo ""
    WScript.Echo "=== STEP 5: CONFIRM FACE SELECTION ==="
    
    ' Activate Inventor window
    WshShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 300
    
    ' Re-select the face (in case selection was cleared)
    selectSet.Clear
    selectSet.Select largestFace
    WScript.Sleep 200
    
    ' Send Enter to confirm selection
    WScript.Echo "Sending Enter to confirm face..."
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 500
    
    ' Step 6: Handle Sheet Metal Defaults dialog
    WScript.Echo ""
    WScript.Echo "=== STEP 6: SHEET METAL DEFAULTS DIALOG ==="
    WScript.Echo "Dialog should appear. Sending OK..."
    
    ' Wait for dialog
    WScript.Sleep 500
    
    ' Just press Enter to accept defaults (thickness should already be detected)
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 1000
    
    ' Update the document
    partDoc.Update
    WScript.Sleep 500
End If

' Step 7: Create Flat Pattern
WScript.Echo ""
WScript.Echo "=== STEP 7: CREATE FLAT PATTERN ==="

' Refresh compDef
Set compDef = partDoc.ComponentDefinition

' Check if it's now sheet metal
WScript.Echo "Checking sheet metal status..."
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern
If Err.Number <> 0 Then Err.Clear

If Not compDef.HasFlatPattern Then
    WScript.Echo "Creating flat pattern..."
    
    ' Try Unfold
    compDef.Unfold
    
    If Err.Number = 0 Then
        WScript.Echo "Unfold succeeded!"
    Else
        WScript.Echo "Unfold failed: " & Err.Description
        Err.Clear
        
        ' Try via command
        Dim unfoldCmd
        Set unfoldCmd = cmdMgr.ControlDefinitions.Item("PartUnfoldCmd")
        If Not unfoldCmd Is Nothing And unfoldCmd.Enabled Then
            WScript.Echo "Trying PartUnfoldCmd..."
            unfoldCmd.Execute
            WScript.Sleep 500
            WshShell.SendKeys "{ENTER}"
        End If
    End If
    
    partDoc.Update
End If

' Step 8: Check result
WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="

Set compDef = partDoc.ComponentDefinition
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo ""
    WScript.Echo "Flat pattern dimensions:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** SUCCESS! ORIENTATION IS CORRECT! ***"
    Else
        WScript.Echo ""
        WScript.Echo "Edge view - orientation needs fix"
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
