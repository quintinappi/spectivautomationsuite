On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument

WScript.Echo "Layer Diagnostic"
WScript.Echo "================"

Set layers = doc.StylesManager.Layers
WScript.Echo "Total Layers: " & layers.Count

For Each l In layers
    WScript.Echo " - " & l.Name & " (Internal Name: " & l.InternalName & ")"
Next
