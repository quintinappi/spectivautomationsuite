' specific_view_fix.vbs
Option Explicit

Dim invApp, doc, sheet
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.Sheets.Item(2) ' Sheet 2

WScript.Echo "=== VIEW STYLE TRANSFER ==="

' 1. FIND THE CORRECT VIEW AND WRONG VIEWS BY NAME
Dim viewCorrect, viewWrong
Set viewCorrect = Nothing

Dim v
For Each v In sheet.DrawingViews
    If InStr(UCase(v.Name), "VIEW1") > 0 Then
        Set viewCorrect = v
        WScript.Echo "Found Correct View: " & v.Name
    End If
Next

If viewCorrect Is Nothing Then
    WScript.Echo "ERROR: Could not find VIEW1"
    WScript.Quit
End If

' 2. GET THE HIDDEN LINE LAYER FROM VIEW1
Dim correctLayer
Set correctLayer = Nothing

Dim curves
Set curves = viewCorrect.DrawingCurves
WScript.Echo "Scanning " & curves.Count & " lines in " & viewCorrect.Name & "..."

Dim c
For Each c In curves
    ' kHiddenEdge = 32258
    If c.EdgeType = 32258 Then
        If c.Segments.Count > 0 Then
            Set correctLayer = c.Segments.Item(1).Layer
            WScript.Echo "SUCCESS: Found Hidden Layer -> " & correctLayer.Name
            WScript.Echo "Color: R=" & correctLayer.Color.Red & " G=" & correctLayer.Color.Green & " B=" & correctLayer.Color.Blue
            Exit For
        End If
    End If
Next

If correctLayer Is Nothing Then
    WScript.Echo "ERROR: VIEW1 has no hidden lines to copy from!"
    WScript.Quit
End If

' 3. APPLY TO ELEVATION, C, and PLAN
Dim targetNames
targetNames = Array("ELEVATION", "C", "PLAN")

Dim name
For Each name In targetNames
    Dim targetView
    Set targetView = Nothing
    
    ' Find view by name
    For Each v In sheet.DrawingViews
        If InStr(UCase(v.Name), name) > 0 Then
            Set targetView = v
            Exit For
        End If
    Next
    
    If Not targetView Is Nothing Then
        WScript.Echo "Updating View: " & targetView.Name
        
        Dim updateCount
        updateCount = 0
        
        Set curves = targetView.DrawingCurves
        For Each c In curves
            ' If it's a hidden line (32258), force the layer
            If c.EdgeType = 32258 Then
                Dim s
                For Each s In c.Segments
                    s.Layer = correctLayer
                    updateCount = updateCount + 1
                Next
            End If
        Next
        WScript.Echo "  -> Fixed " & updateCount & " lines."
    End If
Next

doc.Update
WScript.Echo "Done!"
