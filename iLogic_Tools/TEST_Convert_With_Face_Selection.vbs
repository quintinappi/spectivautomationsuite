' TEST_Convert_With_Face_Selection.vbs
' Run Convert to Sheet Metal command with face pre-selection
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, selectSet

WScript.Echo "=== CONVERT TO SHEET METAL WITH FACE SELECTION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
WScript.Echo "Part: " & partDoc.DisplayName

Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

' Step 1: List all faces and find the largest planar face
WScript.Echo ""
WScript.Echo "=== STEP 1: ANALYZE FACES ==="

Dim body, faces, face
Dim largestFace, largestArea, largestFaceIndex
Dim secondLargestFace, secondLargestArea
largestArea = 0
secondLargestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
If Err.Number <> 0 Then
    WScript.Echo "Error getting body: " & Err.Description
    Err.Clear
End If

Set faces = body.Faces
WScript.Echo "Total faces: " & faces.Count

Dim faceNum, faceInfo
faceInfo = ""
faceNum = 0

For Each face In faces
    faceNum = faceNum + 1
    
    Dim surfType, area, faceTypeName
    surfType = face.SurfaceType
    
    Select Case surfType
        Case 1: faceTypeName = "Cone"
        Case 2: faceTypeName = "Cylinder"
        Case 3: faceTypeName = "Plane"
        Case 4: faceTypeName = "Sphere"
        Case 5: faceTypeName = "Torus"
        Case 6: faceTypeName = "BSpline"
        Case Else: faceTypeName = "Other(" & surfType & ")"
    End Select
    
    area = 0
    On Error Resume Next
    area = face.Evaluator.Area * 100 ' Convert to mm²
    If Err.Number <> 0 Then
        area = 0
        Err.Clear
    End If
    
    WScript.Echo "  Face " & faceNum & ": " & faceTypeName & " - " & FormatNumber(area, 0) & " mm²"
    
    ' Track largest planar faces
    If surfType = 3 Then ' Plane
        If area > largestArea Then
            secondLargestArea = largestArea
            Set secondLargestFace = largestFace
            
            largestArea = area
            Set largestFace = face
            largestFaceIndex = faceNum
        ElseIf area > secondLargestArea Then
            secondLargestArea = area
            Set secondLargestFace = face
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "Largest planar face: Face " & largestFaceIndex & " (" & FormatNumber(largestArea, 0) & " mm²)"

' Step 2: Find the Convert to Sheet Metal command
WScript.Echo ""
WScript.Echo "=== STEP 2: FIND CONVERT COMMAND ==="

Dim convertCmds
convertCmds = Array( _
    "PartConvertToSheetMetalCmd", _
    "SheetMetalConvertCmd", _
    "ConvertToSheetMetalCmd", _
    "PartSheetMetalConvertCmd", _
    "SMConvertCmd" _
)

Dim convertCmd, cmdName
Set convertCmd = Nothing

For Each cmdName In convertCmds
    Set convertCmd = cmdMgr.ControlDefinitions.Item(cmdName)
    If Not convertCmd Is Nothing Then
        WScript.Echo "Found: " & cmdName & " (Enabled: " & convertCmd.Enabled & ")"
        If convertCmd.Enabled Then Exit For
    End If
    Err.Clear
Next

' Also scan for any command with "convert" and "sheet"
WScript.Echo ""
WScript.Echo "Scanning for convert commands..."

Dim ctrlDef
For Each ctrlDef In cmdMgr.ControlDefinitions
    Dim name
    name = LCase(ctrlDef.InternalName)
    
    If InStr(name, "convert") > 0 And InStr(name, "sheet") > 0 Then
        WScript.Echo "  " & ctrlDef.InternalName & " (Enabled: " & ctrlDef.Enabled & ")"
        If ctrlDef.Enabled And convertCmd Is Nothing Then
            Set convertCmd = ctrlDef
        End If
    End If
    Err.Clear
Next

' Step 3: Pre-select the largest face and run command
WScript.Echo ""
WScript.Echo "=== STEP 3: SELECT FACE AND RUN COMMAND ==="

If convertCmd Is Nothing Then
    WScript.Echo "No convert command found!"
    WScript.Quit
End If

If Not convertCmd.Enabled Then
    WScript.Echo "Convert command is not enabled (part may already be sheet metal)"
    WScript.Quit
End If

' Clear any existing selection
selectSet.Clear
WScript.Echo "Selection cleared"

' Select the largest face
WScript.Echo "Selecting largest face..."
selectSet.Select largestFace

If Err.Number = 0 Then
    WScript.Echo "Face selected! Selection count: " & selectSet.Count
Else
    WScript.Echo "Face selection failed: " & Err.Description
    Err.Clear
End If

' Verify what's selected
If selectSet.Count > 0 Then
    WScript.Echo "Selected item type: " & TypeName(selectSet.Item(1))
End If

' Execute the convert command
WScript.Echo ""
WScript.Echo "Executing: " & convertCmd.InternalName & "..."
WScript.Echo ""
WScript.Echo ">>> A DIALOG MAY APPEAR IN INVENTOR <<<"
WScript.Echo ">>> The face should already be selected <<<"

convertCmd.Execute

If Err.Number = 0 Then
    WScript.Echo "Command executed!"
Else
    WScript.Echo "Command failed: " & Err.Description
    Err.Clear
End If

' Wait for dialog interaction
WScript.Sleep 3000
partDoc.Update

' Check result
WScript.Echo ""
WScript.Echo "=== CHECKING RESULT ==="

Set compDef = partDoc.ComponentDefinition
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern
If Err.Number <> 0 Then Err.Clear

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
