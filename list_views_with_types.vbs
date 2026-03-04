Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Active Sheet: " & sheet.Name
WScript.Echo "===================="
WScript.Echo ""

For Each v In sheet.DrawingViews
    viewType = ""
    On Error Resume Next
    viewType = v.ViewType
    If Err.Number <> 0 Then
        viewType = "Unknown"
        Err.Clear
    End If
    
    ' Determine if it's a base view or child view
    parentInfo = ""
    On Error Resume Next
    Set parent = v.ParentView
    If Err.Number = 0 And Not parent Is Nothing Then
        parentInfo = " (Parent: " & parent.Name & ")"
    End If
    Err.Clear
    
    WScript.Echo "View Name: " & v.Name
    WScript.Echo "  Type: " & viewType & parentInfo
    WScript.Echo ""
Next
