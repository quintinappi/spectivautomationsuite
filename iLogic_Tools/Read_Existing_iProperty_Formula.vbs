' Read the formula from existing PLATE LENGTH / PLATE WIDTH iProperties
Option Explicit

Sub Main()
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp.ActiveDocument Is Nothing Then
        MsgBox "No active document"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument

    WScript.Echo "=== READING iPROPERTIES FROM: " & partDoc.DisplayName & " ==="
    WScript.Echo ""

    ' Get custom property set
    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    WScript.Echo "=== ALL CUSTOM iPROPERTIES ==="
    Dim i
    For i = 1 To customPropSet.Count
        Dim prop
        Set prop = customPropSet.Item(i)
        WScript.Echo "Name: " & prop.Name
        WScript.Echo "  Value: " & prop.Value
        WScript.Echo "  Expression: " & prop.Expression
        WScript.Echo ""
    Next

End Sub

Main
