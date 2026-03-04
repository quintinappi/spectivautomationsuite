' Quick_Scan.vbs
' Quick diagnostic for centerlines

Option Explicit
Dim invApp, doc, sheet, view, anno
Dim foundCenterLayers

On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If

Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Scanning view '1' on sheet: " & sheet.Name
WScript.Echo ""

' Find view "1"
For Each view In sheet.DrawingViews
    If view.Name = "1" Then
        WScript.Echo "Found view 1"
        WScript.Echo "Curves: " & view.DrawingCurves.Count
        Exit For
    End If
Next

' Check annotations
WScript.Echo ""
WScript.Echo "Checking sheet annotations..."
Dim count
count = 0
For Each anno In sheet.Annotations
    count = count + 1
    If TypeName(anno) = "Centerline" OR TypeName(anno) = "CenterMark" Then
        WScript.Echo TypeName(anno) & ": " & anno.Layer.Name
    End If
Next
WScript.Echo "Total annotations checked: " & count
WScript.Echo ""

' List all layers
WScript.Echo "All layers in document:"
Dim lyr
For Each lyr In doc.StylesManager.Layers
    If InStr(UCase(lyr.Name), "CENTER") > 0 Then
        WScript.Echo "  * " & lyr.Name & " [CENTERLINE LAYER]"
    End If
Next

WScript.Echo "Done"
