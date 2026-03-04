' =========================================================
' DEBUG - CHECK ALL PARAMETER COLLECTIONS
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== DEBUG ALL PARAMETER COLLECTIONS ==="
    WScript.Echo ""
    
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    If activeDoc.DocumentType <> 12290 Then ' kPartDocumentObject
        WScript.Echo "ERROR: Not a part document (.ipt)"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & activeDoc.DisplayName
    WScript.Echo ""
    
    Dim compDef
    Set compDef = activeDoc.ComponentDefinition
    
    ' Check User Parameters
    WScript.Echo "=== USER PARAMETERS ==="
    Dim userParams
    Set userParams = compDef.Parameters.UserParameters
    
    Dim i
    For i = 1 To userParams.Count
        Dim userParam
        Set userParam = userParams.Item(i)
        WScript.Echo "  " & userParam.Name & " (Type: " & TypeName(userParam) & ")"
        
        ' Try to access ExportParameter
        On Error Resume Next
        If userParam.Name = "Thickness" Then
            Dim exportVal
            exportVal = userParam.ExportParameter
            If Err.Number = 0 Then
                WScript.Echo "    - ExportParameter = " & exportVal
            Else
                WScript.Echo "    - ExportParameter = (not available - Err: " & Err.Description & ")"
            End If
            Err.Clear
        End If
        On Error Goto 0
    Next
    
    WScript.Echo ""
    WScript.Echo "=== MODEL PARAMETERS ==="
    Dim modelParams
    Set modelParams = compDef.Parameters.ModelParameters
    
    For i = 1 To modelParams.Count
        Dim modelParam
        Set modelParam = modelParams.Item(i)
        WScript.Echo "  " & modelParam.Name & " (Type: " & TypeName(modelParam) & ")"
    Next
    
    WScript.Echo ""
    WScript.Echo "=== MODEL REFERENCE PARAMETERS ==="
    Dim refParams
    Set refParams = compDef.Parameters.ReferenceParameters
    
    For i = 1 To refParams.Count
        Dim refParam
        Set refParam = refParams.Item(i)
        WScript.Echo "  " & refParam.Name & " (Type: " & TypeName(refParam) & ")"
    Next
    
End Sub
