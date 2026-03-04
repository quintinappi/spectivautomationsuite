' TEST_Convert_With_Preselection.vbs
' Pre-select the largest face, then execute convert command
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, cmdMgr

WScript.Echo "=== CONVERT WITH PRE-SELECTION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo ""

' Check if already sheet metal - need to revert first
If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Part is sheet metal - reverting to standard..."
    
    ' Delete flat pattern
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        WScript.Echo "  Flat pattern deleted"
    End If
    
    ' Change subtype back to standard
    compDef.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}"
    If Err.Number <> 0 Then
        WScript.Echo "  SubType change: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  SubType changed to standard"
    End If
    
    partDoc.Update
    Set compDef = partDoc.ComponentDefinition
    WScript.Echo ""
End If

' Find the largest face
WScript.Echo "=== FINDING LARGEST FACE ==="

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
WScript.Echo ""

' Pre-select the face
WScript.Echo "=== PRE-SELECTING FACE ==="
selectSet.Clear
selectSet.Select largestFace

WScript.Echo "SelectSet.Count: " & selectSet.Count

If selectSet.Count > 0 Then
    WScript.Echo "Face type in SelectSet: " & TypeName(selectSet.Item(1))
End If

' Now execute the convert command
WScript.Echo ""
WScript.Echo "=== EXECUTING CONVERT COMMAND ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

If convertCmd Is Nothing Then
    WScript.Echo "Convert command not found!"
    WScript.Quit
End If

WScript.Echo "Command: " & convertCmd.DisplayName
WScript.Echo "Enabled: " & convertCmd.Enabled

If Not convertCmd.Enabled Then
    WScript.Echo "Command is disabled!"
    WScript.Quit
End If

' Execute with face pre-selected
WScript.Echo ""
WScript.Echo "Executing with pre-selected face..."
WScript.Echo ">>> IF A DIALOG APPEARS, JUST CLICK OK <<<"
WScript.Echo ""

convertCmd.Execute

' Wait and check status
WScript.Echo "Waiting 3 seconds..."
WScript.Sleep 3000

WScript.Echo ""
WScript.Echo "=== CHECKING RESULT ==="

partDoc.Update
Set compDef = partDoc.ComponentDefinition

WScript.Echo "SubType: " & compDef.SubType
WScript.Echo "Is Sheet Metal: " & (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "HasFlatPattern: " & compDef.HasFlatPattern

If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    If Not compDef.HasFlatPattern Then
        WScript.Echo ""
        WScript.Echo "Creating flat pattern..."
        compDef.Unfold
        partDoc.Update
    End If
    
    If compDef.HasFlatPattern Then
        Dim fp
        Set fp = compDef.FlatPattern
        WScript.Echo ""
        WScript.Echo "=== FLAT PATTERN RESULT ==="
        WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        If fp.Length * 10 > 100 And fp.Width * 10 > 100 Then
            WScript.Echo ""
            WScript.Echo "*** SUCCESS! CORRECT ORIENTATION! ***"
        Else
            WScript.Echo ""
            WScript.Echo "*** WRONG ORIENTATION - still on edge face ***"
        End If
    End If
Else
    WScript.Echo ""
    WScript.Echo "Conversion incomplete - command may still be active"
    WScript.Echo "ActiveCommand: " & cmdMgr.ActiveCommand
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
