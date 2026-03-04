' Debug_Flip.vbs
Option Explicit
Dim invApp, doc
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

Dim sheet
Set sheet = doc.Sheets.Item(2)
Dim v
Set v = sheet.DrawingViews.Item(1) ' ELEVATION

WScript.Echo "Targeting View: " & v.Name

Dim activeStd
Set activeStd = doc.StylesManager.ActiveStandardStyle
Dim objDefs
Set objDefs = activeStd.ActiveObjectDefaults

Dim correctVis
Set correctVis = objDefs.VisibleEdgeLayer
WScript.Echo "Target Visible Layer: " & correctVis.Name

Dim curves
Set curves = v.DrawingCurves
Dim c, s
Dim count
count = 0

For Each c In curves
    For Each s In c.Segments
        Dim curName
        curName = UCase(s.Layer.Name)
        
        If count < 5 Then
            WScript.Echo "Found Layer: " & curName
        End If
        
        ' If it is Hidden, change to Visible
        If InStr(curName, "HIDDEN") > 0 Then
            s.Layer = correctVis
            If count < 5 Then WScript.Echo " -> Changed to " & correctVis.Name
            count = count + 1
        End If
    Next
Next

doc.Update
WScript.Echo "Done. Updated " & count & " segments."
