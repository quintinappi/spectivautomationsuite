' =========================================================
' ENABLE THICKNESS EXPORT - API METHOD
' =========================================================
' Uses the ExposedAsProperty property to enable
' the Export Parameter checkbox for Thickness
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== ENABLE THICKNESS EXPORT - API METHOD ==="
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
    
    ' Get Parameters
    Dim params
    Set params = activeDoc.ComponentDefinition.Parameters
    
    ' Find Thickness in ModelParameters
    Dim modelParams
    Set modelParams = params.ModelParameters
    
    Dim thicknessParam
    Set thicknessParam = Nothing
    
    Dim i
    For i = 1 To modelParams.Count
        If modelParams.Item(i).Name = "Thickness" Then
            Set thicknessParam = modelParams.Item(i)
            Exit For
        End If
    Next
    
    If thicknessParam Is Nothing Then
        WScript.Echo "ERROR: Thickness parameter not found in ModelParameters"
        WScript.Quit 1
    End If
    
    WScript.Echo "Found Thickness parameter"
    WScript.Echo "Type: " & TypeName(thicknessParam)
    WScript.Echo ""
    
    ' Check current ExposedAsProperty value
    WScript.Echo "Current ExposedAsProperty: " & thicknessParam.ExposedAsProperty
    
    ' Enable export
    WScript.Echo "Setting ExposedAsProperty = True..."
    thicknessParam.ExposedAsProperty = True
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: " & Err.Description
        Err.Clear
        WScript.Quit 1
    End If
    
    WScript.Echo "New ExposedAsProperty: " & thicknessParam.ExposedAsProperty
    
    ' Save
    WScript.Echo ""
    WScript.Echo "Saving part..."
    activeDoc.Save
    
    WScript.Echo ""
    WScript.Echo "SUCCESS! Thickness ExposedAsProperty enabled."
    WScript.Echo ""
    
End Sub
