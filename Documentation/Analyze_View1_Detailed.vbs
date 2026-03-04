' Analyze_View1_Detailed.vbs
Option Explicit

Dim invApp, doc, sheet, v
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.Sheets.Item(2)
Dim view1
Set view1 = Nothing

For Each v In sheet.DrawingViews
    If InStr(UCase(v.Name), "VIEW1") > 0 Then
        Set view1 = v
        Exit For
    End If
Next

WScript.Echo "Analyzing VIEW1 Types..."
Dim curves
Set curves = view1.DrawingCurves

Dim typeCounts
Set typeCounts = CreateObject("Scripting.Dictionary")

Dim c
For Each c In curves
    If Not typeCounts.Exists(c.EdgeType) Then
        typeCounts.Add c.EdgeType, 0
    End If
    typeCounts(c.EdgeType) = typeCounts(c.EdgeType) + 1
Next

Dim k
For Each k In typeCounts.Keys
    WScript.Echo "Type " & k & ": " & typeCounts(k) & " curves"
Next
