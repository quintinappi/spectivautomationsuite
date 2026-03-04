Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet
WScript.Echo "Active Sheet: " & sheet.Name
WScript.Echo "===================="
For Each v In sheet.DrawingViews
    WScript.Echo "View Name: " & v.Name
Next
