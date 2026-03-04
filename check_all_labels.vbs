Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Active Sheet: " & sheet.Name
WScript.Echo "===================="
WScript.Echo ""

For Each v In sheet.DrawingViews
    viewType = ""
    On Error Resume Next
    viewTypeNum = v.ViewType
    If Err.Number = 0 Then
        Select Case viewTypeNum
            Case 10501: viewType = "Base View"
            Case 10502: viewType = "Projected View"
            Case 10505: viewType = "Detail View"
            Case Else:  viewType = "Type " & viewTypeNum
        End Select
    End If
    
    WScript.Echo "View Name: " & v.Name
    WScript.Echo "  Type: " & viewType
    
    On Error Resume Next
    Set label = v.Label
    If Err.Number = 0 And Not label Is Nothing Then
        WScript.Echo "  FormattedText: " & label.FormattedText
    Else
        WScript.Echo "  FormattedText: [No Label]"
    End If
    Err.Clear
    
    WScript.Echo ""
Next
