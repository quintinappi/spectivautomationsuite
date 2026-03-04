On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "Part: " & doc.DisplayName

Dim compDef
Set compDef = doc.ComponentDefinition

Dim flatPattern
Set flatPattern = compDef.FlatPattern

If flatPattern Is Nothing Then
    WScript.Echo "No flat pattern"
    WScript.Quit
End If

Dim length, width, baseArea
length = Round(flatPattern.Length / 10, 1)
width = Round(flatPattern.Width / 10, 1)

If Not flatPattern.BaseFace Is Nothing Then
    baseArea = Round(flatPattern.BaseFace.Evaluator.Area * 100, 0)
Else
    baseArea = 0
End If

WScript.Echo "Length: " & length & " mm"
WScript.Echo "Width: " & width & " mm"
WScript.Echo "BaseFace Area: " & baseArea & " mm²"
WScript.Echo ""

If baseArea > 1000000 Then
    WScript.Echo "*** CORRECT - Large face (2,653,726 mm²) is base ***"
ElseIf baseArea < 30000 Then
    WScript.Echo "*** WRONG - Edge face (23,100 mm²) is base ***"
Else
    WScript.Echo "*** UNKNOWN - BaseFace area: " & baseArea & " mm² ***"
End If
