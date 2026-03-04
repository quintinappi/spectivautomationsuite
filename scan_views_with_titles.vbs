Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Active Sheet: " & sheet.Name
WScript.Echo "===================="
WScript.Echo ""

For Each v In sheet.DrawingViews
    ' Get view type
    viewType = ""
    On Error Resume Next
    viewTypeNum = v.ViewType
    If Err.Number <> 0 Then
        viewType = "Unknown"
        Err.Clear
    Else
        Select Case viewTypeNum
            Case 10501: viewType = "Base View"
            Case 10502: viewType = "Projected View"
            Case 10503: viewType = "Auxiliary View"
            Case 10504: viewType = "Section View"
            Case 10505: viewType = "Detail View"
            Case 10506: viewType = "Draft View"
            Case Else:  viewType = "Type " & viewTypeNum
        End Select
    End If
    
    ' Determine if it's a base view or child view
    parentInfo = ""
    On Error Resume Next
    Set parent = v.ParentView
    If Err.Number = 0 And Not parent Is Nothing Then
        parentInfo = " (Parent: " & parent.Name & ")"
    End If
    Err.Clear
    
    ' Get title/label content
    titleText = "[No Label]"
    On Error Resume Next
    Set viewLabel = v.Label
    If Err.Number = 0 And Not viewLabel Is Nothing Then
        titleText = viewLabel.Text
        If titleText = "" Then titleText = "[Empty Label]"
    End If
    Err.Clear
    
    ' Output view info
    WScript.Echo "View Name: " & v.Name
    WScript.Echo "  Type: " & viewType & parentInfo
    WScript.Echo "  Title: " & titleText
    WScript.Echo ""
Next

WScript.Echo "Total views: " & sheet.DrawingViews.Count
