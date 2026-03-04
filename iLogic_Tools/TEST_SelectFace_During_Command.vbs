' TEST_SelectFace_During_Command.vbs
' Tries to select a face while a command is waiting for selection
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet

WScript.Echo "=== SELECT FACE DURING COMMAND ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Find the largest face
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

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""

' Clear any existing selection
selectSet.Clear

' Method 1: Try SelectSet.Select
WScript.Echo "=== METHOD 1: SelectSet.Select ==="
selectSet.Select largestFace
If Err.Number <> 0 Then
    WScript.Echo "SelectSet.Select: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Face added to SelectSet"
    WScript.Echo "SelectSet.Count: " & selectSet.Count
End If

' Give Inventor a moment
WScript.Sleep 500

' Method 2: Try through TransientBRep / HighlightSet
WScript.Echo ""
WScript.Echo "=== METHOD 2: HighlightSet ==="

Dim highlightSet
Set highlightSet = partDoc.HighlightSet

If Err.Number <> 0 Then
    WScript.Echo "HighlightSet: " & Err.Description
    Err.Clear
Else
    WScript.Echo "HighlightSet Type: " & TypeName(highlightSet)
    highlightSet.Clear
    highlightSet.AddItem largestFace
    If Err.Number <> 0 Then
        WScript.Echo "AddItem: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Face added to HighlightSet"
    End If
End If

' Method 3: Try InteractionEvents
WScript.Echo ""
WScript.Echo "=== METHOD 3: Check Active Environment ==="

Dim cmdMgr
Set cmdMgr = invApp.CommandManager

WScript.Echo "ActiveCommand: " & cmdMgr.ActiveCommand
WScript.Echo "ActiveEnvironment: " & invApp.ActiveEnvironment.DisplayName

' Method 4: Try to fire a selection event
WScript.Echo ""
WScript.Echo "=== METHOD 4: TransientObjects ObjectCollection ==="

Dim transObjs, objColl
Set transObjs = invApp.TransientObjects
Set objColl = transObjs.CreateObjectCollection

objColl.Add largestFace
WScript.Echo "ObjectCollection created with face"
WScript.Echo "Collection Count: " & objColl.Count

' Try to use this with command
WScript.Echo ""
WScript.Echo "=== CURRENT STATE ==="
WScript.Echo "SelectSet.Count: " & selectSet.Count
WScript.Echo "ActiveCommand: " & cmdMgr.ActiveCommand

' If a command is waiting, try DoEvents
WScript.Echo ""
WScript.Echo "Calling invApp.CommandManager.StopActiveCommand..."
' cmdMgr.StopActiveCommand  ' This would cancel the command

WScript.Echo ""
WScript.Echo "The face should now be highlighted/selected."
WScript.Echo "If a dialog is waiting, try clicking OK."
WScript.Echo ""
WScript.Echo "=== DONE ==="
