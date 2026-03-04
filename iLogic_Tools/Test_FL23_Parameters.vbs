' Quick Test - Check if FL23 has Length parameter

Option Explicit

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim asmDoc
Set asmDoc = invApp.ActiveDocument

' Find FL23
Dim occ
For Each occ In asmDoc.ComponentDefinition.Occurrences
    If InStr(occ.Name, "FL23") > 0 Then
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        WScript.Echo "Found: " & refDoc.DisplayName
        
        Dim userParams
        Set userParams = refDoc.ComponentDefinition.Parameters.UserParameters
        
        WScript.Echo "User parameters count: " & userParams.Count
        
        Dim i, param
        For i = 1 To userParams.Count
            Set param = userParams.Item(i)
            WScript.Echo "  " & param.Name & " = " & param.Value
        Next
        
        ' Try to get Length
        On Error Resume Next
        Dim lengthParam
        Set lengthParam = userParams.Item("Length")
        
        If Err.Number = 0 And Not lengthParam Is Nothing Then
            WScript.Echo "HAS Length parameter!"
        Else
            WScript.Echo "NO Length parameter (Error: " & Err.Number & " - " & Err.Description & ")"
        End If
        Err.Clear
        
        ' Try to get Length2
        Dim length2Param
        Set length2Param = userParams.Item("Length2")
        
        If Err.Number = 0 And Not length2Param Is Nothing Then
            WScript.Echo "HAS Length2 parameter!"
        Else
            WScript.Echo "NO Length2 parameter (Error: " & Err.Number & " - " & Err.Description & ")"
        End If
        
        Exit For
    End If
Next
