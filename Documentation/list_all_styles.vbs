On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

WScript.Echo "Styles Manager Content"
WScript.Echo "======================"

' List all standard styles
WScript.Echo "Standard Styles:"
For Each std In doc.StylesManager.StandardStyles
    WScript.Echo " - " & std.Name
Next

' List all Drawing View Styles
WScript.Echo ""
WScript.Echo "Drawing View Styles:"
Set dvs = doc.StylesManager.DrawingViewStyles
If dvs Is Nothing Then
    WScript.Echo " Could not get DrawingViewStyles collection"
Else
    WScript.Echo " Count: " & dvs.Count
    For Each s In dvs
        WScript.Echo " - " & s.Name
    Next
End If

' Check Active Standard
WScript.Echo ""
WScript.Echo "Active Standard: " & doc.StylesManager.ActiveStandardStyle.Name
