Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set sheet = doc.ActiveSheet
Set view = sheet.DrawingViews.Item(1)
Set label = view.Label

WScript.Echo "View Name: " & view.Name
WScript.Echo ""
WScript.Echo "=== LABEL FORMATTED TEXT ==="
WScript.Echo label.FormattedText
WScript.Echo ""
WScript.Echo "=== LABEL TEXT ==="
WScript.Echo label.Text
