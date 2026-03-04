' TEST_Revert_Then_Convert.vbs
' Use the Revert command first, then convert with correct face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr

WScript.Echo "=== REVERT THEN CONVERT ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "Document SubType: " & partDoc.SubType
WScript.Echo "ComponentDef Type: " & TypeName(compDef)
WScript.Echo ""

' Step 1: Revert to standard part
WScript.Echo "=== STEP 1: REVERT TO STANDARD PART ==="

' Delete flat pattern first if exists
If TypeName(compDef) = "SheetMetalComponentDefinition" Then
    If compDef.HasFlatPattern Then
        WScript.Echo "Deleting flat pattern..."
        compDef.FlatPattern.Delete
        partDoc.Update
        WScript.Echo "  Done"
    End If
End If

Dim revertCmd
Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")

If revertCmd.Enabled Then
    WScript.Echo "Executing: " & revertCmd.DisplayName
    revertCmd.Execute
    
    WScript.Sleep 2000
    partDoc.Update
    
    Set compDef = partDoc.ComponentDefinition
    WScript.Echo "New ComponentDef Type: " & TypeName(compDef)
    WScript.Echo "New SubType: " & partDoc.SubType
Else
    WScript.Echo "Revert command not enabled"
End If

WScript.Echo ""

' Step 2: Find largest face
WScript.Echo "=== STEP 2: FINDING LARGEST FACE ==="

Set compDef = partDoc.ComponentDefinition

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
WScript.Echo ""

' Step 3: Pre-select and convert
WScript.Echo "=== STEP 3: CONVERT TO SHEET METAL ==="

' Check if convert command is now enabled
Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

WScript.Echo "Convert command enabled: " & convertCmd.Enabled

If convertCmd.Enabled Then
    ' Pre-select the largest face
    selectSet.Clear
    selectSet.Select largestFace
    WScript.Echo "Face pre-selected (Count=" & selectSet.Count & ")"
    
    WScript.Echo ""
    WScript.Echo "Executing Convert to Sheet Metal..."
    WScript.Echo ">>> LOOK AT INVENTOR - IF THE FACE IS GREEN, CLICK OK <<<"
    WScript.Echo ""
    
    convertCmd.Execute
    
    WScript.Echo "Waiting for user action..."
    WScript.Sleep 5000
    
    partDoc.Update
End If

' Step 4: Create flat pattern
WScript.Echo ""
WScript.Echo "=== STEP 4: CHECK RESULT ==="

Set compDef = partDoc.ComponentDefinition
WScript.Echo "ComponentDef Type: " & TypeName(compDef)
WScript.Echo "SubType: " & partDoc.SubType

If TypeName(compDef) = "SheetMetalComponentDefinition" Then
    If Not compDef.HasFlatPattern Then
        WScript.Echo "Creating flat pattern..."
        compDef.Unfold
        partDoc.Update
    End If
    
    If compDef.HasFlatPattern Then
        Dim fp
        Set fp = compDef.FlatPattern
        WScript.Echo ""
        WScript.Echo "=== FLAT PATTERN DIMENSIONS ==="
        WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        If fp.Length * 10 > 100 And fp.Width * 10 > 100 Then
            WScript.Echo ""
            WScript.Echo "******************************"
            WScript.Echo "*** SUCCESS! CORRECT FACE! ***"
            WScript.Echo "******************************"
        Else
            WScript.Echo ""
            WScript.Echo "*** STILL WRONG ORIENTATION ***"
        End If
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
