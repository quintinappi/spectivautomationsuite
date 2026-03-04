' =========================================================
' DEBUG - LIST ALL PARAMETER PROPERTIES
' =========================================================

Option Explicit

Dim m_InventorApp

Main()

Sub Main()
    On Error Resume Next
    
    WScript.Echo "=== DEBUG PARAMETER PROPERTIES ==="
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
    
    Dim params
    Set params = activeDoc.ComponentDefinition.Parameters
    
    Dim length2Param
    Set length2Param = Nothing
    
    On Error Resume Next
    Set length2Param = params.UserParameters.Item("Length2")
    On Error Goto 0
    
    If length2Param Is Nothing Then
        WScript.Echo "ERROR: Length2 parameter not found"
        WScript.Quit 1
    End If
    
    WScript.Echo "Length2 parameter found!"
    WScript.Echo "Type: " & TypeName(length2Param)
    WScript.Echo ""
    WScript.Echo "Available properties and methods:"
    WScript.Echo ""
    
    ' List common properties
    Dim props
    props = Array("Name", "Expression", "Value", "DisplayValue", "Tolerance", "Comment", "Precision", "Unit", "ParsedExpression")
    
    Dim i
    For i = LBound(props) To UBound(props)
        On Error Resume Next
        Dim val
        val = Eval("length2Param." & props(i))
        If Err.Number = 0 Then
            WScript.Echo "  " & props(i) & " = " & val
        Else
            WScript.Echo "  " & props(i) & " = (not available)"
        End If
        Err.Clear
        On Error Goto 0
    Next
    
End Sub
