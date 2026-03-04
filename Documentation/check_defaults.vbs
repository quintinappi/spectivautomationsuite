On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
Set doc = invApp.ActiveDocument
Set activeStd = doc.StylesManager.ActiveStandardStyle
Set objDef = activeStd.ActiveObjectDefaults

WScript.Echo "Active Standard: " & activeStd.Name
WScript.Echo "Object Defaults Style: " & objDef.Name

' Use ObjectDefaults style properties
' The property names are slightly different
' We'll try to get them by calling specific methods if needed, 
' but normally they are properties.

' Check if we can get the Visible Edge Layer
Set l = objDef.VisibleEdgeLayer
If Not l Is Nothing Then
    WScript.Echo "Visible Edge Layer Default: " & l.Name
End If

Set l = objDef.HiddenEdgeLayer
If Not l Is Nothing Then
    WScript.Echo "Hidden Edge Layer Default: " & l.Name
End If
