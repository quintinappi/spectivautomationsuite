' Check active part parameters
Dim app, doc, cd, mp, p
Set app = GetObject(, "Inventor.Application")
Set doc = app.ActiveDocument
Set cd = doc.ComponentDefinition
Set mp = cd.Parameters.ModelParameters

WScript.Echo "Part: " & doc.DisplayName
WScript.Echo "Model Parameters:"

Dim i
For i = 1 To mp.Count
    Set p = mp.Item(i)
    WScript.Echo "  " & p.Name & " = " & p.Value & " " & p.Units & " (ModelValue: " & p.ModelValue & ")"
Next

WScript.Echo ""
WScript.Echo "d2 details:"
Set p = mp.Item("d2")
WScript.Echo "  ModelValue in BASE units: " & p.ModelValue
WScript.Echo "  Converted to mm (x10): " & (p.ModelValue * 10)
