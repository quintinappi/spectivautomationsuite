' Quick check - what are the sheet metal parameter names?
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

    WScript.Echo "=== PARAMETER LIST FOR: " & partDoc.DisplayName & " ==="
    WScript.Echo ""

    Dim params
    Set params = partDoc.ComponentDefinition.Parameters

    Dim i
    For i = 1 To params.Count
        Dim param
        Set param = params.Item(i)
        WScript.Echo param.Name & " = " & param.Value & " " & param.Units
    Next

    WScript.Echo ""
    WScript.Echo "=== USER PARAMETERS ==="
    Dim userParams
    Set userParams = params.UserParameters
    For i = 1 To userParams.Count
        Set param = userParams.Item(i)
        WScript.Echo param.Name & " = " & param.Value & " " & param.Units
    Next
End Sub

Main
