' Check what's currently open in Inventor
Dim app, doc
Set app = GetObject(, "Inventor.Application")

If app.ActiveDocument Is Nothing Then
    WScript.Echo "No document is currently open in Inventor."
Else
    Set doc = app.ActiveDocument
    WScript.Echo "Currently open document:"
    WScript.Echo "  Name: " & doc.DisplayName
    WScript.Echo "  Full Path: " & doc.FullFileName
    WScript.Echo "  Type: " & doc.DocumentType
    
    If doc.DocumentType = 12291 Then
        WScript.Echo "  Document is an ASSEMBLY (.iam)"
    ElseIf doc.DocumentType = 12290 Then
        WScript.Echo "  Document is a PART (.ipt)"
    ElseIf doc.DocumentType = 12292 Then
        WScript.Echo "  Document is a DRAWING (.idw)"
    Else
        WScript.Echo "  Document is an unknown type"
    End If
    
    WScript.Echo ""
    WScript.Echo "Please open the assembly (DM Underpan.iam) to run the scan."
End If
