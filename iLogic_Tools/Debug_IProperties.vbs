' =========================================================
' CHECK THICKNESS IN iPROPERTIES
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== CHECK THICKNESS IN iPROPERTIES ==="
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
    
    ' Check all PropertySets
    Dim propertySets
    Set propertySets = activeDoc.PropertySets
    
    WScript.Echo "Available PropertySets:"
    
    Dim i
    For i = 1 To propertySets.Count
        Dim propSet
        Set propSet = propertySets.Item(i)
        WScript.Echo "  " & propSet.DisplayName
        
        ' Check if "Thickness" property exists in any set
        On Error Resume Next
        Dim thickProp
        Set thickProp = propSet.Item("Thickness")
        If Err.Number = 0 Then
            WScript.Echo "    -> Found Thickness property!"
            WScript.Echo "       Value: " & thickProp.Value
        End If
        Err.Clear
        On Error Goto 0
    Next
    
    WScript.Echo ""
    WScript.Echo "Checking parameter export through UserParameters:"
    
    Dim userParams
    Set userParams = activeDoc.ComponentDefinition.Parameters.UserParameters
    
    Dim j
    For j = 1 To userParams.Count
        Dim userParam
        Set userParam = userParams.Item(j)
        WScript.Echo "  UserParam: " & userParam.Name
        
        ' Try ExportParameter property
        On Error Resume Next
        Dim exportVal
        exportVal = userParam.ExportParameter
        If Err.Number = 0 Then
            WScript.Echo "    ExportParameter = " & exportVal
        Else
            WScript.Echo "    ExportParameter = (property doesn't exist)"
        End If
        Err.Clear
        On Error Goto 0
    Next
    
End Sub
