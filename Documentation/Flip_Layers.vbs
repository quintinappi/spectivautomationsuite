' Flip_Layers.vbs
' Your colors are swapped (Visible is Cyan, Hidden is Black).
' This script simply swaps them back.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

' 1. GET CORRECT LAYERS
Dim activeStd, objDefs
Set activeStd = doc.StylesManager.ActiveStandardStyle
Set objDefs = activeStd.ActiveObjectDefaults

Dim correctVis, correctHid
Set correctVis = objDefs.VisibleEdgeLayer
Set correctHid = objDefs.HiddenEdgeLayer

WScript.Echo "FLIP OPERATION:"
WScript.Echo "Turning 'Hidden/Cyan' lines -> " & correctVis.Name & " (Black)"
WScript.Echo "Turning 'Visible/Black' lines -> " & correctHid.Name & " (Cyan)"

' 2. PROCESS VIEWS
Dim sheet
Set sheet = doc.Sheets.Item(2)
Dim targetNames, viewName
targetNames = Array("ELEVATION", "C", "D", "PLAN")

For Each viewName In targetNames
    Dim targetView
    Set targetView = Nothing
    For Each v In sheet.DrawingViews
        If InStr(UCase(v.Name), viewName) > 0 Then
            Set targetView = v
            Exit For
        End If
    Next
    
    If Not targetView Is Nothing Then
        WScript.Echo "Flipping Layers in: " & targetView.Name
        
        Dim curves
        Set curves = targetView.DrawingCurves
        Dim c, s
        Dim countFlippedToVis, countFlippedToHid
        countFlippedToVis = 0
        countFlippedToHid = 0
        
        For Each c In curves
            For Each s In c.Segments
                Dim currentName
                currentName = UCase(s.Layer.Name)
                
                ' If currently Hidden/Cyan -> Make Visible
                If InStr(currentName, "HIDDEN") > 0 Or InStr(currentName, "DASHED") > 0 Then
                    s.Layer = correctVis
                    countFlippedToVis = countFlippedToVis + 1
                    
                ' If currently Visible/Black -> Make Hidden
                ElseIf InStr(currentName, "VISIBLE") > 0 Or InStr(currentName, "CONT") > 0 Then
                    s.Layer = correctHid
                    countFlippedToHid = countFlippedToHid + 1
                End If
            Next
        Next
        WScript.Echo "  -> To Visible: " & countFlippedToVis & " | To Hidden: " & countFlippedToHid
    End If
Next

doc.Update
WScript.Echo "Done!"
