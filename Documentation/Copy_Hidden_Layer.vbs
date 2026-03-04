' Copy_Hidden_Layer.vbs
' Finds the hidden line layer used in VIEW1 (View 5) and applies it to ELEVATION (View 1)

Option Explicit

Dim invApp, doc, sheet
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.Sheets.Item(2) ' Specifically target Sheet 2

WScript.Echo "Style Transfer: VIEW1 -> Other Views"
WScript.Echo "===================================="

' 1. Get the Source Layer from VIEW1 (Index 5)
Dim sourceView
Set sourceView = sheet.DrawingViews.Item(5)
WScript.Echo "Source View: " & sourceView.Name

Dim targetLayerName
targetLayerName = ""
Dim targetLayer
Set targetLayer = Nothing

Dim curves
Set curves = sourceView.DrawingCurves
WScript.Echo "  Source Curve Count: " & curves.Count

Dim c
For Each c In curves
    ' kHiddenEdge = 32258
    If c.EdgeType = 32258 Then
        Dim segs
        Set segs = c.Segments
        If segs.Count > 0 Then
            Set targetLayer = segs.Item(1).Layer
            targetLayerName = targetLayer.Name
            WScript.Echo "  FOUND correct hidden layer: '" & targetLayerName & "'"
            WScript.Echo "  Color: R=" & targetLayer.Color.Red & " G=" & targetLayer.Color.Green & " B=" & targetLayer.Color.Blue
            Exit For
        End If
    End If
Next

If targetLayerName = "" Then
    WScript.Echo "Error: Could not find any hidden lines in VIEW1 to copy from!"
    WScript.Quit
End If

WScript.Echo ""

' 2. Apply this layer to the other views
Dim i
For i = 1 To 4 ' Views 1 through 4 (ELEVATION, C, D, PLAN)
    Dim destView
    Set destView = sheet.DrawingViews.Item(i)
    WScript.Echo "Processing: " & destView.Name
    
    Dim dCurves
    Set dCurves = destView.DrawingCurves
    
    Dim count
    count = 0
    
    For Each c In dCurves
        ' Find Hidden Edges (32258)
        If c.EdgeType = 32258 Then
            Dim s
            For Each s In c.Segments
                s.Layer = targetLayer
                count = count + 1
            Next
        End If
    Next
    
    WScript.Echo "  Updated " & count & " segments."
Next

WScript.Echo ""
WScript.Echo "Updating sheet..."
doc.Update
WScript.Echo "Done."
