' Scan_View_Styles.vbs
' Diagnostic tool to scan all styles/layers in a drawing view

Option Explicit

Dim invApp, doc, activeSheet
Dim viewName, targetView
Dim v, c, s, curves
Dim layer, layerName
Dim foundLayers
Dim i, anno

On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running!"
    WScript.Quit 1
End If

Set doc = invApp.ActiveDocument
If doc Is Nothing Then
    WScript.Echo "ERROR: No active document!"
    WScript.Quit 1
End If

If doc.DocumentType <> 12292 Then
    WScript.Echo "ERROR: Active document is not an IDW!"
    WScript.Quit 1
End If

Set activeSheet = doc.ActiveSheet

WScript.Echo "=========================================="
WScript.Echo "  VIEW STYLE SCANNER - DIAGNOSTIC"
WScript.Echo "=========================================="
WScript.Echo ""

' List all views on active sheet
WScript.Echo "Views on active sheet (" & activeSheet.Name & "):"
For Each v In activeSheet.DrawingViews
    WScript.Echo "  - " & v.Name
Next
WScript.Echo ""

' Ask which view to scan
viewName = InputBox("Enter view name to scan:", "Select View", "VIEW1")
If viewName = "" Then
    WScript.Echo "Cancelled."
    WScript.Quit
End If

' Find the view
Set targetView = Nothing
For Each v In activeSheet.DrawingViews
    If UCase(v.Name) = UCase(viewName) Then
        Set targetView = v
        Exit For
    End If
Next

If targetView Is Nothing Then
    WScript.Echo "ERROR: View '" & viewName & "' not found!"
    WScript.Quit 1
End If

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SCANNING VIEW: " & targetView.Name
WScript.Echo "=========================================="
WScript.Echo ""

' Create dictionary to track unique layers
Set foundLayers = CreateObject("Scripting.Dictionary")

' ===== SCAN DRAWING CURVES (Geometry) =====
WScript.Echo "--- DRAWING CURVES (Geometry Lines) ---"
Set curves = targetView.DrawingCurves
WScript.Echo "Total curves: " & curves.Count
WScript.Echo ""

For Each c In curves
    For Each s In c.Segments
        Set layer = s.Layer
        layerName = layer.Name
        
        If Not foundLayers.Exists(layerName) Then
            foundLayers.Add layerName, 1
        Else
            foundLayers(layerName) = foundLayers(layerName) + 1
        End If
    Next
Next

If foundLayers.Count > 0 Then
    WScript.Echo "Unique layers found in geometry:"
    Dim key
    For Each key In foundLayers.Keys
        WScript.Echo "  - " & key & " (" & foundLayers(key) & " segments)"
    Next
Else
    WScript.Echo "  No geometry curves found"
End If

' ===== SCAN ALL ANNOTATIONS ON SHEET =====
WScript.Echo ""
WScript.Echo "--- SCANNING ALL SHEET ANNOTATIONS ---"

Dim allAnno, annoCount
annoCount = 0

' Try to iterate through all annotations on the sheet
On Error Resume Next
For Each anno In activeSheet.Annotations
    annoCount = annoCount + 1
    
    ' Check if this annotation belongs to our target view
    On Error Resume Next
    Dim parentView
    Set parentView = anno.Parent
    
    If Err.Number = 0 And Not parentView Is Nothing Then
        If UCase(parentView.Name) = UCase(targetView.Name) Then
            WScript.Echo ""
            WScript.Echo "Annotation #" & annoCount & ":"
            WScript.Echo "  Type: " & TypeName(anno)
            
            On Error Resume Next
            If anno.Layer Is Nothing Then
                WScript.Echo "  Layer: (none)"
            Else
                WScript.Echo "  Layer: " & anno.Layer.Name
                
                ' Check if it's a centerline-type layer
                If InStr(UCase(anno.Layer.Name), "CENTER") > 0 Then
                    WScript.Echo "  *** FOUND CENTERLINE LAYER: " & anno.Layer.Name & " ***"
                End If
            End If
            Err.Clear
        End If
    End If
    Err.Clear
Next

If annoCount = 0 Then
    WScript.Echo "  No annotations found on sheet"
End If

' ===== TRY TO FIND CENTERLINES VIA SHEET SKETCHES =====
WScript.Echo ""
WScript.Echo "--- SCANNING SHEET SKETCHES ---"
On Error Resume Next
Dim sk, sketchEntity
For Each sk In activeSheet.Sketches
    WScript.Echo "Sketch: " & sk.Name
    
    For Each sketchEntity In sk.SketchEntities
        On Error Resume Next
        WScript.Echo "  Entity: " & TypeName(sketchEntity)
        
        ' Check for centerline sketch entities
        If InStr(UCase(TypeName(sketchEntity)), "CENTER") > 0 Then
            WScript.Echo "    *** FOUND: " & TypeName(sketchEntity) & " ***"
            If Not sketchEntity.Layer Is Nothing Then
                WScript.Echo "    Layer: " & sketchEntity.Layer.Name
            End If
        End If
        Err.Clear
    Next
Next
Err.Clear

' ===== SCAN FOR ALL LAYERS IN DOCUMENT =====
WScript.Echo ""
WScript.Echo "--- ALL LAYERS IN DOCUMENT ---"
On Error Resume Next
Dim allLayers, lyr
Set allLayers = doc.StylesManager.Layers
If Err.Number = 0 And Not allLayers Is Nothing Then
    For Each lyr In allLayers
        WScript.Echo "  - " & lyr.Name
    Next
Else
    WScript.Echo "  Could not access layers"
End If
Err.Clear

WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SCAN COMPLETE"
WScript.Echo "=========================================="
