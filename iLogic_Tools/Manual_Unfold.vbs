' Manual_Unfold.vbs
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Creating flat pattern manually..."

compDef.Unfold

If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Description
Else
    partDoc.Update
    WScript.Echo "Flat pattern created"
    
    If compDef.HasFlatPattern Then
        Dim fp
        Set fp = compDef.FlatPattern
        WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        Dim baseFace
        Set baseFace = fp.BaseFace
        Dim area
        area = baseFace.Evaluator.Area * 100
        WScript.Echo "BaseFace Area: " & FormatNumber(area, 0) & " mm²"
    End If
End If
