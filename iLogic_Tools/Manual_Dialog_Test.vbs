On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== MANUAL DIALOG TEST ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition
Dim cmdMgr
Set cmdMgr = invApp.CommandManager

' Find largest face
Dim faces, face, largestFace, largestArea, area
Set faces = compDef.SurfaceBodies.Item(1).Faces
largestArea = 0

For Each face In faces
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""

' Pre-select it
Dim selectSet
Set selectSet = doc.SelectSet
selectSet.Clear
WScript.Sleep 300
selectSet.Select largestFace

WScript.Echo "Pre-selected face. SelectSet.Count = " & selectSet.Count
WScript.Echo ""
WScript.Echo "Now I will execute the Convert command."
WScript.Echo "Watch Inventor carefully:"
WScript.Echo "  1. Is the large face highlighted/green when command starts?"
WScript.Echo "  2. Does the dialog show a face already selected?"
WScript.Echo "  3. What happens to the SelectSet?"
WScript.Echo ""
WScript.Echo "Starting in 3 seconds..."
WScript.Sleep 3000

' Execute convert command
Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

convertCmd.Execute

WScript.Sleep 2000

WScript.Echo ""
WScript.Echo "Command executed. Dialog should be open now."
WScript.Echo "SelectSet.Count after command = " & selectSet.Count
WScript.Echo ""
WScript.Echo "*** DO NOT PRESS ANYTHING YET ***"
WScript.Echo "Look at the dialog - is there a face highlighted in green?"
WScript.Echo "If not, the pre-selection was cleared."
WScript.Echo ""
WScript.Echo "Press Enter when ready to analyze..."
WScript.StdIn.ReadLine
