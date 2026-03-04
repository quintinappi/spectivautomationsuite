' TEST_Set_BaseFace.vbs
' Try to SET the BaseFace property to the largest face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp

WScript.Echo "=== SET BASE FACE ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set smDef = compDef

WScript.Echo "Part: " & partDoc.DisplayName

' Ensure flat pattern exists
If Not smDef.HasFlatPattern Then
    WScript.Echo "Creating flat pattern..."
    smDef.Unfold
    partDoc.Update
End If

Set fp = smDef.FlatPattern

WScript.Echo ""
WScript.Echo "Current flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"

' Get current base face
Dim currentBaseFace
Set currentBaseFace = fp.BaseFace
WScript.Echo "Current BaseFace area: " & FormatNumber(currentBaseFace.Evaluator.Area * 100, 0) & " mm²"

' Find the largest face
WScript.Echo ""
WScript.Echo "Finding largest face..."

Dim body, faces, face, largestFace, largestArea
largestArea = 0
Set body = smDef.SurfaceBodies.Item(1)

For Each face In body.Faces
    Dim area
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
    Err.Clear
Next

WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"

' Try different ways to set the base face
WScript.Echo ""
WScript.Echo "=== TRYING TO SET BASE FACE ==="

WScript.Echo ""
WScript.Echo "Method 1: fp.BaseFace = largestFace"
Set fp.BaseFace = largestFace
If Err.Number = 0 Then
    WScript.Echo "  SUCCESS!"
    partDoc.Update
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "Method 2: fp.SetBaseFace(largestFace)"
fp.SetBaseFace largestFace
If Err.Number = 0 Then
    WScript.Echo "  SUCCESS!"
    partDoc.Update
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "Method 3: fp.ChangeBaseFace(largestFace)"
fp.ChangeBaseFace largestFace
If Err.Number = 0 Then
    WScript.Echo "  SUCCESS!"
    partDoc.Update
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "Method 4: fp.RedefineBaseFace(largestFace)"
fp.RedefineBaseFace largestFace
If Err.Number = 0 Then
    WScript.Echo "  SUCCESS!"
    partDoc.Update
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "Method 5: fp.Edit - then change"
fp.Edit
If Err.Number = 0 Then
    WScript.Echo "  Edit mode entered!"
    
    ' Try setting in edit mode
    Set fp.BaseFace = largestFace
    If Err.Number = 0 Then
        WScript.Echo "  BaseFace set in edit mode!"
    Else
        WScript.Echo "  Error: " & Err.Description
        Err.Clear
    End If
    
    fp.ExitEdit
Else
    WScript.Echo "  Error: " & Err.Description
    Err.Clear
End If

' Check final result
partDoc.Update

WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="
Set fp = smDef.FlatPattern
WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
WScript.Echo "BaseFace area: " & FormatNumber(fp.BaseFace.Evaluator.Area * 100, 0) & " mm²"

If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
    WScript.Echo ""
    WScript.Echo "*** SUCCESS! ***"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
