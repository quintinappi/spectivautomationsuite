On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.Sheets.Item(2) ' Sheet 2

WScript.Echo "Comparing View Styles on Sheet 2"
WScript.Echo "================================"

For Each v In sheet.DrawingViews
    WScript.Echo "View: " & v.Name
    
    Set curves = v.DrawingCurves
    If Err.Number <> 0 Then
        WScript.Echo "  Error getting curves: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Curve Count: " & curves.Count
        
        ' Find a hidden curve to check its layer
        Dim foundHidden, foundVisible
        foundHidden = False
        foundVisible = False
        
        For Each c In curves
            ' kHiddenEdge = 32258, kVisibleEdge = 32257
            If c.EdgeType = 32258 And Not foundHidden Then
                If c.Segments.Count > 0 Then
                    WScript.Echo "  Hidden Line Layer: " & c.Segments.Item(1).Layer.Name
                    foundHidden = True
                End If
            End If
            If c.EdgeType = 32257 And Not foundVisible Then
                 If c.Segments.Count > 0 Then
                    WScript.Echo "  Visible Line Layer: " & c.Segments.Item(1).Layer.Name
                    foundVisible = True
                End If
            End If
            If foundHidden And foundVisible Then Exit For
        Next
        
        If Not foundHidden Then WScript.Echo "  Hidden Line Layer: Not Found"
        If Not foundVisible Then WScript.Echo "  Visible Line Layer: Not Found"
    End If
    WScript.Echo "--------------------------------"
Next
