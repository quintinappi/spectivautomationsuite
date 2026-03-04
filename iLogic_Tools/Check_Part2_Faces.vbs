' Check_Part2_Faces.vbs
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, body, faces, face

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument

WScript.Echo "=== PART2 FACE ANALYSIS ==="
WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Revert to standard first
If partDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Reverting to standard..."
    Set compDef = partDoc.ComponentDefinition
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        partDoc.Update
    End If
    
    Dim cmdMgr, revertCmd
    Set cmdMgr = invApp.CommandManager
    Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
    If revertCmd.Enabled Then
        revertCmd.Execute
        WScript.Sleep 1500
        partDoc.Update
    End If
    WScript.Echo ""
End If

Set compDef = partDoc.ComponentDefinition
Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

WScript.Echo "Total faces: " & faces.Count
WScript.Echo ""

Dim i
For i = 1 To faces.Count
    Set face = faces.Item(i)
    Dim area, surfType
    area = face.Evaluator.Area * 100
    surfType = face.SurfaceType
    
    WScript.Echo "Face " & i & ":"
    WScript.Echo "  Area: " & FormatNumber(area, 0) & " mm²"
    WScript.Echo "  Type: " & surfType & " (0=Plane)"
    
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: " & Err.Description
        Err.Clear
    End If
Next

WScript.Echo ""
WScript.Echo "=== DONE ==="
