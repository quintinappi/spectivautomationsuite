Option Explicit

' Diagnostic script to dump all AttributeSets in a document

Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")

If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running!"
    WScript.Quit 1
End If

Dim activeDoc
Set activeDoc = invApp.ActiveDocument

WScript.Echo "Document: " & activeDoc.DisplayName
WScript.Echo ""
WScript.Echo "ATTRIBUTESETS IN DOCUMENT:"
WScript.Echo "================================================================"

Dim attrSets
Set attrSets = activeDoc.AttributeSets

WScript.Echo "Total AttributeSets: " & attrSets.Count
WScript.Echo ""

Dim attrSet
For Each attrSet In attrSets
    WScript.Echo "AttributeSet: " & attrSet.Name
    WScript.Echo "  Attributes count: " & attrSet.Count

    Dim attr
    For Each attr In attrSet
        On Error Resume Next
        WScript.Echo "    Attribute: " & attr.Name & " = " & CStr(Left(attr.Value, 100))
        Err.Clear
    Next
    WScript.Echo ""
Next

WScript.Echo "================================================================"
WScript.Echo "DONE"
