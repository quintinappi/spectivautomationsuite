' Verify_Custom_Properties.vbs
' Check that PLATE LENGTH and PLATE WIDTH have the correct formulas
Option Explicit
On Error Resume Next

Dim invApp, partDoc, customPropSet

WScript.Echo "=== VERIFY CUSTOM PROPERTIES ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not get custom property set"
    WScript.Quit
End If

WScript.Echo "=== CUSTOM PROPERTIES ==="

' Check PLATE LENGTH
Dim lengthProp
Set lengthProp = customPropSet.Item("PLATE LENGTH")

If Err.Number = 0 Then
    WScript.Echo "PLATE LENGTH:"
    WScript.Echo "  Value: " & lengthProp.Value
    WScript.Echo "  Expression: " & lengthProp.Expression
    WScript.Echo ""
Else
    WScript.Echo "PLATE LENGTH: NOT FOUND"
    WScript.Echo ""
    Err.Clear
End If

' Check PLATE WIDTH
Dim widthProp
Set widthProp = customPropSet.Item("PLATE WIDTH")

If Err.Number = 0 Then
    WScript.Echo "PLATE WIDTH:"
    WScript.Echo "  Value: " & widthProp.Value
    WScript.Echo "  Expression: " & widthProp.Expression
    WScript.Echo ""
Else
    WScript.Echo "PLATE WIDTH: NOT FOUND"
    WScript.Echo ""
    Err.Clear
End If

' Check sheet metal dimensions
Dim compDef
Set compDef = partDoc.ComponentDefinition

If compDef.HasFlatPattern Then
    Dim fp
    Set fp = compDef.FlatPattern
    
    WScript.Echo "=== FLAT PATTERN DIMENSIONS ==="
    WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

WScript.Echo ""
WScript.Echo "=== VERIFICATION COMPLETE ==="
