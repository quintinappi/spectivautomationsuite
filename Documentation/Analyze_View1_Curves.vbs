' Analyze_View1_Curves.vbs
Option Explicit

Dim invApp, doc, sheet, view
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.Sheets.Item(2)
Set view = sheet.DrawingViews.Item(5) ' VIEW1

WScript.Echo "Analyzing VIEW1 Curves..."
WScript.Echo "View Name: " & view.Name
WScript.Echo "Total Curves: " & view.DrawingCurves.Count

Dim i, c
Dim count
count = 0

For Each c In view.DrawingCurves
    count = count + 1
    If count > 50 Then Exit For ' Just check the first 50
    
    Dim typeName
    Select Case c.EdgeType
        Case 32257: typeName = "Visible"
        Case 32258: typeName = "Hidden"
        Case 32259: typeName = "Tangent"
        Case 32260: typeName = "Bend"
        Case Else: typeName = "Other (" & c.EdgeType & ")"
    End Select
    
    Dim layerName, colorInfo
    If c.Segments.Count > 0 Then
        layerName = c.Segments.Item(1).Layer.Name
        Dim color
        Set color = c.Segments.Item(1).Layer.Color
        colorInfo = "R:" & color.Red & " G:" & color.Green & " B:" & color.Blue
    Else
        layerName = "No Segments"
    End If
    
    WScript.Echo count & ". " & typeName & " | Layer: " & layerName & " | " & colorInfo
Next
