' Quick_Check_Part.vbs
' Check current part status
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument

WScript.Echo "=== CURRENT PART STATUS ==="
WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "SubType: " & partDoc.SubType

Set compDef = partDoc.ComponentDefinition

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    WScript.Echo ""
    WScript.Echo "Flat Pattern:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    Dim baseFace
    Set baseFace = fp.BaseFace
    If Not baseFace Is Nothing Then
        Dim area
        area = baseFace.Evaluator.Area * 100
        WScript.Echo "  BaseFace Area: " & FormatNumber(area, 0) & " mm²"
    End If
Else
    WScript.Echo "No flat pattern"
End If

' Check custom properties
Dim customPropSet, lengthProp, widthProp
Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

WScript.Echo ""
WScript.Echo "Custom Properties:"

On Error Resume Next
Set lengthProp = customPropSet.Item("PLATE LENGTH")
If Err.Number = 0 Then
    WScript.Echo "  PLATE LENGTH: " & lengthProp.Expression
Else
    WScript.Echo "  PLATE LENGTH: NOT FOUND"
End If
Err.Clear

Set widthProp = customPropSet.Item("PLATE WIDTH")
If Err.Number = 0 Then
    WScript.Echo "  PLATE WIDTH: " & widthProp.Expression
Else
    WScript.Echo "  PLATE WIDTH: NOT FOUND"
End If
