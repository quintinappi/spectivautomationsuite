On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "Part: " & doc.DisplayName

Dim compDef
Set compDef = doc.ComponentDefinition

WScript.Echo "Creating flat pattern..."
compDef.Unfold

Dim flatPattern
Set flatPattern = compDef.FlatPattern

If flatPattern Is Nothing Then
    WScript.Echo "ERROR: No flat pattern created"
    WScript.Quit 1
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

' Determine if this is correct orientation
If baseArea > 1000000 Then
    WScript.Echo "*** CORRECT ORIENTATION - Large face is base ***"
Else
    WScript.Echo "*** WRONG ORIENTATION - Edge view ***"
End If
