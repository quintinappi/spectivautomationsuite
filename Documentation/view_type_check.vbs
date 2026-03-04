On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set sheet = invApp.ActiveDocument.ActiveSheet
Set v = sheet.DrawingViews.Item(1)

WScript.Echo "View 1 Name: " & v.Name
WScript.Echo "View 1 TypeName: " & TypeName(v)
WScript.Echo "View Style Value: " & v.ViewStyle

Err.Clear
Set s = v.Style
If Err.Number <> 0 Then
    WScript.Echo "v.Style Error: " & Err.Description
Else
    WScript.Echo "v.Style Name: " & s.Name
End If

Err.Clear
WScript.Echo "DisplayStyle: " & v.DisplayStyle
If Err.Number <> 0 Then WScript.Echo "DisplayStyle Error: " & Err.Description
