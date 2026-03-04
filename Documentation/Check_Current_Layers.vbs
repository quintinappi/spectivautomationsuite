' Check_Current_Layers.vbs
Option Explicit
Dim invApp, doc, sheet, v
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.Sheets.Item(2)
Set v = sheet.DrawingViews.Item(1) ' ELEVATION

WScript.Echo "View: " & v.Name
Dim curves
Set curves = v.DrawingCurves
WScript.Echo "Curves: " & curves.Count
If curves.Count > 0 Then
    Dim c
    Set c = curves.Item(1)
    If c.Segments.Count > 0 Then
        WScript.Echo "First Segment Layer: " & c.Segments.Item(1).Layer.Name
    End If
End If
