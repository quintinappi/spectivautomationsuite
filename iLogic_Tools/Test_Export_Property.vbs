' Test Export Parameter property
Dim app, doc, cd, up, param

Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set cd = doc.ComponentDefinition
Set up = cd.Parameters.UserParameters

On Error Resume Next
Set param = up.Item("Length2")
If Err.Number = 0 Then
    WScript.Echo "Length2 found"
    WScript.Echo "Trying different property names..."
    
    Err.Clear
    param.ExportParameter = True
    If Err.Number = 0 Then
        WScript.Echo "ExportParameter = True : SUCCESS"
    Else
        WScript.Echo "ExportParameter : FAILED - " & Err.Description
    End If
    
    Err.Clear
    param.Export = True
    If Err.Number = 0 Then
        WScript.Echo "Export = True : SUCCESS"
    Else
        WScript.Echo "Export : FAILED - " & Err.Description
    End If
    
    Err.Clear
    param.ExportToPartsList = True
    If Err.Number = 0 Then
        WScript.Echo "ExportToPartsList = True : SUCCESS"
    Else
        WScript.Echo "ExportToPartsList : FAILED - " & Err.Description
    End If
End If
