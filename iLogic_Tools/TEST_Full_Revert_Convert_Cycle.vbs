' TEST_Full_Revert_Convert_Cycle.vbs
' Revert to standard, then convert to sheet metal with correct face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr

WScript.Echo "=== FULL REVERT AND CONVERT CYCLE ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Step 1: Check current state
WScript.Echo "=== STEP 1: CURRENT STATE ==="
WScript.Echo "SubType: " & compDef.SubType

Dim isSheetMetal
isSheetMetal = (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo "Has Flat Pattern: True"
    WScript.Echo "Current dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

' Step 2: Revert to standard if sheet metal
If isSheetMetal Then
    WScript.Echo ""
    WScript.Echo "=== STEP 2: REVERTING TO STANDARD ==="
    
    ' Delete flat pattern first
    If compDef.HasFlatPattern Then
        WScript.Echo "Deleting flat pattern..."
        compDef.FlatPattern.Delete
        If Err.Number <> 0 Then
            WScript.Echo "  Error: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  Flat pattern deleted"
        End If
        partDoc.Update
    End If
    
    ' Execute revert command
    Dim revertCmd
    Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
    
    If revertCmd Is Nothing Then
        WScript.Echo "Revert command not found!"
    ElseIf Not revertCmd.Enabled Then
        WScript.Echo "Revert command disabled"
    Else
        WScript.Echo "Executing: " & revertCmd.DisplayName
        revertCmd.Execute
        
        WScript.Sleep 1000
        partDoc.Update
        
        Set compDef = partDoc.ComponentDefinition
        WScript.Echo "New SubType: " & compDef.SubType
        isSheetMetal = (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
        WScript.Echo "Is Sheet Metal: " & isSheetMetal
    End If
End If

' Step 3: Find largest face
WScript.Echo ""
WScript.Echo "=== STEP 3: FINDING LARGEST FACE ==="

Set compDef = partDoc.ComponentDefinition

Dim body, faces, face, largestFace, largestArea, largestIndex
largestArea = 0
largestIndex = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

Dim i
i = 0
For Each face In faces
    i = i + 1
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 And area > largestArea Then
        largestArea = area
        Set largestFace = face
        largestIndex = i
    End If
    Err.Clear
Next

WScript.Echo "Largest face: Face " & largestIndex & " = " & FormatNumber(largestArea, 0) & " mm²"

' Step 4: Pre-select face and convert
If Not isSheetMetal Then
    WScript.Echo ""
    WScript.Echo "=== STEP 4: CONVERTING TO SHEET METAL ==="
    
    ' Pre-select the largest face
    selectSet.Clear
    selectSet.Select largestFace
    WScript.Echo "Pre-selected face (SelectSet.Count = " & selectSet.Count & ")"
    
    Dim convertCmd
    Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
    
    If convertCmd Is Nothing Then
        WScript.Echo "Convert command not found!"
    ElseIf Not convertCmd.Enabled Then
        WScript.Echo "Convert command disabled"
    Else
        WScript.Echo "Executing: " & convertCmd.DisplayName
        WScript.Echo ""
        WScript.Echo ">>> WATCH INVENTOR - IF DIALOG APPEARS, CLICK OK <<<"
        WScript.Echo ">>> THE FACE SHOULD ALREADY BE SELECTED! <<<"
        WScript.Echo ""
        
        convertCmd.Execute
        
        WScript.Echo "Waiting 5 seconds for user action..."
        WScript.Sleep 5000
    End If
End If

' Step 5: Check result
WScript.Echo ""
WScript.Echo "=== STEP 5: CHECKING RESULT ==="

partDoc.Update
Set compDef = partDoc.ComponentDefinition

WScript.Echo "SubType: " & compDef.SubType
isSheetMetal = (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern

If isSheetMetal Then
    ' Create flat pattern if needed
    If Not compDef.HasFlatPattern Then
        WScript.Echo "Creating flat pattern..."
        compDef.Unfold
        partDoc.Update
    End If
    
    If compDef.HasFlatPattern Then
        Set fp = compDef.FlatPattern
        WScript.Echo ""
        WScript.Echo "=== FINAL FLAT PATTERN ==="
        WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        If fp.Length * 10 > 100 And fp.Width * 10 > 100 Then
            WScript.Echo ""
            WScript.Echo "**********************************"
            WScript.Echo "*** SUCCESS! CORRECT ORIENTATION! ***"
            WScript.Echo "**********************************"
        Else
            WScript.Echo ""
            WScript.Echo "*** STILL WRONG ORIENTATION ***"
        End If
    End If
Else
    WScript.Echo "Conversion not complete"
    WScript.Echo "ActiveCommand: " & cmdMgr.ActiveCommand
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
