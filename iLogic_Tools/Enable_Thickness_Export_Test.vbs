' =========================================================
' ENABLE THICKNESS EXPORT PARAMETER - TEST SINGLE PART
' =========================================================
' Enables the "Export Parameter" checkbox for Thickness
' parameter in plate parts
' =========================================================

Option Explicit

Dim m_InventorApp
Dim m_LogFile

' Initialize
Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - SINGLE PART TEST ==="
    WScript.Echo ""
    
    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not connect to Inventor."
        WScript.Echo "Please make sure Inventor is running."
        WScript.Quit 1
    End If
    
    Dim activeDoc
    Set activeDoc = m_InventorApp.ActiveDocument
    
    If activeDoc Is Nothing Then
        WScript.Echo "ERROR: No active document."
        WScript.Quit 1
    End If
    
    ' Check if it's a part document
    If activeDoc.DocumentType <> 12290 Then ' kPartDocumentObject
        WScript.Echo "ERROR: Not a part document (.ipt)"
        WScript.Quit 1
    End If
    
    WScript.Echo "Part: " & activeDoc.DisplayName
    WScript.Echo ""
    
    ' Enable export for Thickness parameter
    WScript.Echo "Enabling Export Parameter for Thickness..."
    
    Dim params
    Set params = activeDoc.ComponentDefinition.Parameters
    
    Dim thicknessParam
    Set thicknessParam = Nothing
    
    ' Try to find Thickness parameter
    On Error Resume Next
    Set thicknessParam = params.Item("Thickness")
    On Error Goto 0
    
    If thicknessParam Is Nothing Then
        WScript.Echo "ERROR: Thickness parameter not found"
        WScript.Quit 1
    End If
    
    WScript.Echo "Current ExportParameter value: " & thicknessParam.ExportParameter
    
    ' Enable export
    thicknessParam.ExportParameter = True
    
    WScript.Echo "New ExportParameter value: " & thicknessParam.ExportParameter
    
    ' Save the part
    WScript.Echo ""
    WScript.Echo "Saving part..."
    activeDoc.Save
    
    WScript.Echo ""
    WScript.Echo "SUCCESS! Thickness export parameter enabled."
    WScript.Echo ""
    
End Sub
