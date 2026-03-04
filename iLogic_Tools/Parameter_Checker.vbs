' Parameter Checker - Check assembly parameters
' Author: Quintin de Bruin © 2026

Option Explicit

Sub Main()
    On Error Resume Next

    ' Get Inventor application
    Dim inventorApp
    Set inventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "Inventor not running", vbCritical
        Exit Sub
    End If

    ' Get active document
    Dim doc
    Set doc = inventorApp.ActiveDocument

    If doc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        MsgBox "Please open an assembly document", vbExclamation
        Exit Sub
    End If

    ' Get user parameters
    Dim compDef
    Set compDef = doc.ComponentDefinition

    Dim userParams
    Set userParams = compDef.Parameters.UserParameters

    Dim paramList
    paramList = "Assembly Parameters:" & vbCrLf & vbCrLf

    Dim i
    For i = 1 To userParams.Count
        Dim param
        Set param = userParams.Item(i)
        paramList = paramList & param.Name & " = " & param.Value & " " & param.Units & vbCrLf
    Next

    ' Check for our specific parameters
    Dim lengthParam
    Dim widthParam

    On Error Resume Next
    Set lengthParam = userParams.Item("PLATE LENGTH")
    If Err.Number = 0 Then
        paramList = paramList & vbCrLf & "PLATE LENGTH found: " & lengthParam.Value & " " & lengthParam.Units
    Else
        paramList = paramList & vbCrLf & "PLATE LENGTH NOT FOUND"
    End If
    Err.Clear

    Set widthParam = userParams.Item("PLATE WIDTH")
    If Err.Number = 0 Then
        paramList = paramList & vbCrLf & "PLATE WIDTH found: " & widthParam.Value & " " & widthParam.Units
    Else
        paramList = paramList & vbCrLf & "PLATE WIDTH NOT FOUND"
    End If
    Err.Clear

    ' Output to console instead of MsgBox
    WScript.Echo paramList
End Sub

Main