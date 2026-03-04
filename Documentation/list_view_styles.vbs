On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet

WScript.Echo "Active Sheet Styles Report"
WScript.Echo "=========================="
WScript.Echo "Document Active Standard: " & doc.StylesManager.ActiveStandardStyle.Name
WScript.Echo ""

For Each v In sheet.DrawingViews
    Dim styleName
    styleName = "Unknown"
    
    ' Clear error and try direct Style property
    Err.Clear
    styleName = v.Style.Name
    
    If Err.Number <> 0 Then
        ' Fallback: Some views might use nested styles or different property names
        Err.Clear
        ' Try to see if it's a managed style via the Style property without .Name first
        Set s = v.Style
        If Not s Is Nothing Then
            styleName = s.Name
        Else
            ' Check if it's a "display style" enum instead
            styleName = "No Style Object (Enum: " & v.ViewStyle & ")"
        End If
    End If
    
    WScript.Echo "View: " & v.Name & " | Style: " & styleName
Next
