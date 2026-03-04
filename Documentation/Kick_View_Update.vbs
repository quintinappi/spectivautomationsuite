' Kick_View_Update.vbs
' Toggles view style to force regeneration of curve data.

Option Explicit

Dim invApp, doc
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

Dim sheet
Set sheet = doc.Sheets.Item(2)
Dim v
Set v = sheet.DrawingViews.Item(1) ' ELEVATION

WScript.Echo "Kicking View: " & v.Name & " (Current Style: " & v.ViewStyle & ")"

' 1. Toggle to Hidden Line
v.ViewStyle = 32258 ' kHiddenLineDrawingViewStyle
doc.Update
WScript.Echo "Switched to Hidden Line."

' 2. Toggle back to Standard
v.ViewStyle = 32257 ' kFromBaseDrawingViewStyle
doc.Update
WScript.Echo "Switched back to Standard."

' 3. Check result
WScript.Echo "Checking curve layers..."
Dim curves
Set curves = v.DrawingCurves
If curves.Count > 0 Then
    If curves.Item(1).Segments.Count > 0 Then
        WScript.Echo "First segment layer: " & curves.Item(1).Segments.Item(1).Layer.Name
    End If
End If

WScript.Echo "Done."
