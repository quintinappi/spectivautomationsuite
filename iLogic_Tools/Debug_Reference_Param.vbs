' =========================================================
' DEBUG - CHECK REFERENCE PARAMETER PROPERTIES
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== DEBUG REFERENCE PARAMETER PROPERTIES ==="
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
    
    ' Check if Thickness exists as Reference Parameter
    Dim refParams
    Set refParams = compDef.Parameters.ReferenceParameters
    
    Dim thicknessRefParam
    Set thicknessRefParam = Nothing
    
    Dim i
    For i = 1 To refParams.Count
        If refParams.Item(i).Name = "Thickness" Then
            Set thicknessRefParam = refParams.Item(i)
            Exit For
        End If
    Next
    
    If Not (thicknessRefParam Is Nothing) Then
        WScript.Echo "Found Thickness as ReferenceParameter!"
        WScript.Echo "Type: " & TypeName(thicknessRefParam)
        
        ' Try to access ExportParameter
        On Error Resume Next
        Dim exportVal
        exportVal = thicknessRefParam.ExportParameter
        If Err.Number = 0 Then
            WScript.Echo "ExportParameter = " & exportVal
        Else
            WScript.Echo "ERROR accessing ExportParameter: " & Err.Description
        End If
        Err.Clear
        On Error Goto 0
    Else
        WScript.Echo "Thickness is NOT a ReferenceParameter"
        
        ' Try as ModelParameter
        Dim modelParams
        Set modelParams = compDef.Parameters.ModelParameters
        
        For i = 1 To modelParams.Count
            Dim modelParam
            Set modelParam = modelParams.Item(i)
            If modelParam.Name = "Thickness" Then
                WScript.Echo ""
                WScript.Echo "Found Thickness as ModelParameter!"
                WScript.Echo "Type: " & TypeName(modelParam)
                
                ' Check what properties it has
                WScript.Echo ""
                WScript.Echo "Checking common parameter properties:"
                
                Dim props
                props = Array("Name", "Value", "Expression", "Precision", "Comment", "ExportParameter", "IsOptional", "IsVisible", "IsParametricAttribute")
                
                Dim j
                For j = LBound(props) To UBound(props)
                    On Error Resume Next
                    Dim val
                    Set val = Nothing
                    val = modelParam.GetProperty(props(j))
                    If Err.Number = 0 Then
                        WScript.Echo "  " & props(j) & " = " & val
                    Else
                        WScript.Echo "  " & props(j) & " = (error: " & Err.Description & ")"
                    End If
                    Err.Clear
                    On Error Goto 0
                Next
                
                Exit For
            End If
        Next
    End If
    
End Sub
