On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== ANALYZING CONVERT COMMAND BEHAVIOR ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition

' Get SelectSet
Dim selectSet
Set selectSet = doc.SelectSet

' Find largest face
Dim faces
Set faces = compDef.SurfaceBodies.Item(1).Faces

Dim largestFace, largestArea, area, face, i
largestArea = 0

For i = 1 To faces.Count
    Set face = faces.Item(i)
    area = face.Evaluator.Area * 100
    
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face area: " & Round(largestArea, 0) & " mm²"
WScript.Echo ""

' Pre-select the face
selectSet.Clear
WScript.Sleep 200

selectSet.Select largestFace
WScript.Echo "Pre-selected face, SelectSet.Count = " & selectSet.Count
WScript.Sleep 500

' Activate Inventor window
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate "Autodesk Inventor"
WScript.Sleep 500

' Get the convert command
Dim convertCmd
Set convertCmd = invApp.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

WScript.Echo ""
WScript.Echo "*** WATCH CAREFULLY - IS THE LARGE FACE HIGHLIGHTED/GREEN? ***"
WScript.Sleep 3000

WScript.Echo "Executing command..."
convertCmd.Execute

WScript.Sleep 2000

WScript.Echo "Checking SelectSet after command starts..."
WScript.Echo "SelectSet.Count = " & selectSet.Count

If selectSet.Count > 0 Then
    WScript.Echo "SelectSet STILL HAS SELECTION"
    
    Dim selectedFace
    Set selectedFace = selectSet.Item(1)
    WScript.Echo "Selected face area: " & Round(selectedFace.Evaluator.Area * 100, 0) & " mm²"
Else
    WScript.Echo "*** SELECTSET WAS CLEARED BY THE COMMAND ***"
End If

WScript.Echo ""
WScript.Echo "Press Enter in the dialog to continue..."
