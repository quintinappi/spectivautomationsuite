' TEST_Trigger_Pick.vbs
' Use different methods to trigger face pick acceptance
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, selectSet

WScript.Echo "=== TRIGGER FACE PICK ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

' Find largest face
Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
For Each face In body.Faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 And area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Method 1: Try CommandManager.Pick
WScript.Echo ""
WScript.Echo "Method 1: CommandManager.Pick..."

Dim pickedObj
Set pickedObj = cmdMgr.Pick(4096, "Select face") ' 4096 = kPartFaceFilter
If Err.Number = 0 And Not pickedObj Is Nothing Then
    WScript.Echo "Pick returned: " & TypeName(pickedObj)
Else
    WScript.Echo "Pick failed: " & Err.Description
    Err.Clear
End If

' Method 2: Try to use DoSelect
WScript.Echo ""
WScript.Echo "Method 2: invApp.ActiveView.Update then select..."

invApp.ActiveView.Update
selectSet.Clear
selectSet.Select largestFace
WScript.Echo "Selected: " & selectSet.Count

' Method 3: Use AcceptInput command
WScript.Echo ""
WScript.Echo "Method 3: Looking for input acceptance commands..."

Dim cmdNames, cmdName, cmd
cmdNames = Array("AcceptInputCmd", "PartAcceptInputCmd", "SMAcceptInputCmd", _
                 "ConfirmSelectionCmd", "ApplySelectionCmd", "SelectionAcceptCmd", _
                 "PartSelectFaceAcceptCmd", "SheetMetalAcceptFaceCmd")

For Each cmdName In cmdNames
    Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
    If Not cmd Is Nothing Then
        WScript.Echo "  Found: " & cmdName & " (Enabled: " & cmd.Enabled & ")"
        If cmd.Enabled Then
            cmd.Execute
            WScript.Echo "    Executed!"
        End If
    End If
    Err.Clear
Next

' Method 4: Try to directly set the input
WScript.Echo ""
WScript.Echo "Method 4: Check InteractionEvents..."

Dim intEvents
Set intEvents = cmdMgr.ActiveInteractionEvents
If Not intEvents Is Nothing Then
    WScript.Echo "Active InteractionEvents found!"
    WScript.Echo "  Name: " & intEvents.Name
    
    ' Try to trigger selection
    Dim selEvents
    Set selEvents = intEvents.SelectEvents
    If Not selEvents Is Nothing Then
        WScript.Echo "  SelectEvents available"
        selEvents.AddToSelectedEntities largestFace
        WScript.Echo "  Added face to selection"
        
        ' Fire selection
        intEvents.SetCursor 1 ' kCursorArrow
        WScript.Echo "  Cursor set"
    End If
Else
    WScript.Echo "No active InteractionEvents"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
