' Smart_Layer_Migrator.vbs
' Migrates lines to the correct standard layers based on their CURRENT layer name text.

Option Explicit

Dim invApp, doc
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

' 1. GET CORRECT LAYERS FROM STANDARD
Dim activeStd, objDefs
Set activeStd = doc.StylesManager.ActiveStandardStyle
Set objDefs = activeStd.ActiveObjectDefaults

Dim correctVis, correctHid, correctTan, correctCen
Set correctVis = objDefs.VisibleEdgeLayer
Set correctHid = objDefs.HiddenEdgeLayer
Set correctTan = objDefs.TangentEdgeLayer
' Attempt to find centerline layer (often not a direct property, assume name)
On Error Resume Next
Set correctCen = doc.StylesManager.Layers.Item("PEN25 Centerline (ISO)")
If correctCen Is Nothing Then Set correctCen = correctVis ' Fallback

WScript.Echo "MIGRATION MAP:"
WScript.Echo "Current 'Visible' -> " & correctVis.Name
WScript.Echo "Current 'Hidden'  -> " & correctHid.Name
WScript.Echo "Current 'Center'  -> " & correctCen.Name

' 2. MIGRATE VIEWS
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
        WScript.Echo "Migrating: " & targetView.Name
        
        Dim curves
        Set curves = targetView.DrawingCurves
        Dim c, s, countVis, countHid, countCen
        countVis = 0
        countHid = 0
        countCen = 0
        
        For Each c In curves
            For Each s In c.Segments
                Dim currentName
                currentName = UCase(s.Layer.Name)
                
                Dim newLayer
                Set newLayer = Nothing
                
                If InStr(currentName, "HIDDEN") > 0 Or InStr(currentName, "DASHED") > 0 Then
                    Set newLayer = correctHid
                    countHid = countHid + 1
                ElseIf InStr(currentName, "CENTER") > 0 Then
                    Set newLayer = correctCen
                    countCen = countCen + 1
                ElseIf InStr(currentName, "VISIBLE") > 0 Or InStr(currentName, "CONT") > 0 OR InStr(currentName, "DEFAULT") > 0 Then
                    Set newLayer = correctVis
                    countVis = countVis + 1
                Else
                    ' Default unknown layers to Visible (safest)
                    Set newLayer = correctVis
                    countVis = countVis + 1
                End If
                
                If Not newLayer Is Nothing Then
                    s.Layer = newLayer
                End If
            Next
        Next
        WScript.Echo "  -> Visible: " & countVis & " | Hidden: " & countHid & " | Center: " & countCen
    End If
Next

doc.Update
WScript.Echo "Done!"
