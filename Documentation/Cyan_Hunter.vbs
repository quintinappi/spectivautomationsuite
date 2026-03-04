' Cyan_Hunter.vbs
Option Explicit

Dim invApp, doc, sheet
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.Sheets.Item(2)

WScript.Echo "=== CYAN LAYER HUNTER ==="

' 1. FIND VIEW1
Dim viewCorrect
Dim v
For Each v In sheet.DrawingViews
    If InStr(UCase(v.Name), "VIEW1") > 0 Then
        Set viewCorrect = v
        Exit For
    End If
Next

If viewCorrect Is Nothing Then WScript.Echo "VIEW1 not found": WScript.Quit

' 2. FIND THE CYAN LAYER IN VIEW1
Dim cyanLayer
Set cyanLayer = Nothing

Dim curves
Set curves = viewCorrect.DrawingCurves

Dim c
For Each c In curves
    If c.Segments.Count > 0 Then
        Dim l
        Set l = c.Segments.Item(1).Layer
        
        ' Check for Cyan (0, 255, 255)
        If l.Color.Red = 0 And l.Color.Green = 255 And l.Color.Blue = 255 Then
            Set cyanLayer = l
            WScript.Echo "FOUND CYAN LAYER: " & l.Name & " (Type: " & c.EdgeType & ")"
            Exit For
        End If
    End If
Next

If cyanLayer Is Nothing Then
    WScript.Echo "Could not find any Cyan lines in VIEW1."
    ' Fallback: Try to find a layer with 'Hidden' in the name that is NOT the default
    WScript.Quit
End If

' 3. APPLY CYAN LAYER TO ALL TARGET VIEWS
Dim targetNames, viewName
targetNames = Array("ELEVATION", "C", "PLAN")

For Each viewName In targetNames
    Dim viewTarget
    Set viewTarget = Nothing
    
    For Each v In sheet.DrawingViews
        If InStr(UCase(v.Name), viewName) > 0 Then
            Set viewTarget = v
            Exit For
        End If
    Next

    If Not viewTarget Is Nothing Then
        WScript.Echo "Updating " & viewTarget.Name & "..."
        Dim count
        count = 0
        
        Set curves = viewTarget.DrawingCurves
        For Each c In curves
            ' Apply to kHiddenEdge (32258) AND Mystery Type (82695)
            If c.EdgeType = 32258 Or c.EdgeType = 82695 Then
                Dim s
                For Each s In c.Segments
                    s.Layer = cyanLayer
                    count = count + 1
                Next
            End If
        Next
        WScript.Echo "  -> Forced " & count & " segments to " & cyanLayer.Name
    End If
Next

doc.Update
WScript.Echo "Done."
