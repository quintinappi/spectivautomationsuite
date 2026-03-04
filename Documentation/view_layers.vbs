On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.ActiveSheet
Set v = sheet.DrawingViews.Item(1)

WScript.Echo "View: " & v.Name
WScript.Echo "Checking curve layers..."

Set curves = v.DrawingCurves
If curves.Count > 0 Then
    Set c = curves.Item(1)
    
    ' DrawingCurve objects don't have Style/Layer directly, 
    ' but DrawingCurveSegment objects DO.
    Set segments = c.Segments
    If segments.Count > 0 Then
        Set seg = segments.Item(1)
        WScript.Echo "  Segment Layer: " & seg.Layer.Name
        WScript.Echo "  Segment Layer Color: " & seg.Layer.Color.Red & "," & seg.Layer.Color.Green & "," & seg.Layer.Color.Blue
    End If
End If
