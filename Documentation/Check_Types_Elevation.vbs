' Check_Types_Elevation.vbs
Option Explicit
Dim invApp, doc, sheet, v
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.Sheets.Item(2)
Set v = sheet.DrawingViews.Item(1) ' ELEVATION

WScript.Echo "Types in ELEVATION:"
Dim counts
Set counts = CreateObject("Scripting.Dictionary")

Dim c
For Each c In v.DrawingCurves
    If Not counts.Exists(c.EdgeType) Then counts.Add c.EdgeType, 0
    counts(c.EdgeType) = counts(c.EdgeType) + 1
Next

Dim k
For Each k In counts.Keys
    WScript.Echo "Type " & k & ": " & counts(k)
Next
