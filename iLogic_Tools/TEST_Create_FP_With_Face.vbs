' TEST_Create_FP_With_Face.vbs
' Create flat pattern with largest face as base
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp

WScript.Echo "=== CREATE FLAT PATTERN WITH CORRECT FACE ==="
WScript.Echo ""

' Connect to Inventor
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
If partDoc Is Nothing Then
    WScript.Echo "No active document"
    WScript.Quit
End If

WScript.Echo "Part: " & partDoc.DisplayName

Set compDef = partDoc.ComponentDefinition
WScript.Echo "Component type: " & compDef.Type

If compDef.Type <> 99588099 Then
    WScript.Echo "Not a sheet metal part"
    WScript.Quit
End If

Set smDef = compDef
WScript.Echo "Sheet metal confirmed"
WScript.Echo "Has flat pattern: " & smDef.HasFlatPattern

If smDef.HasFlatPattern Then
    WScript.Echo "Flat pattern already exists!"
    Set fp = smDef.FlatPattern
    WScript.Echo "Dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
    WScript.Quit
End If

' Find largest face
WScript.Echo ""
WScript.Echo "Finding largest face..."

Dim body, faces, face
Dim largestFace, largestArea, largestFaceNum
largestArea = 0

Set body = smDef.SurfaceBodies.Item(1)
Set faces = body.Faces

WScript.Echo "Total faces: " & faces.Count

Dim faceNum
faceNum = 0
For Each face In faces
    faceNum = faceNum + 1
    If face.SurfaceType = 3 Then ' kPlaneSurface
        Dim area
        area = face.Evaluator.Area * 100
        WScript.Echo "  Face " & faceNum & " (plane): " & FormatNumber(area, 0) & " mm²"
        
        If area > largestArea Then
            largestArea = area
            Set largestFace = face
            largestFaceNum = faceNum
        End If
    End If
Next

WScript.Echo ""
WScript.Echo "Largest is Face " & largestFaceNum & ": " & FormatNumber(largestArea, 0) & " mm²"

' Create flat pattern with this face
WScript.Echo ""
WScript.Echo "=== CREATING FLAT PATTERN ==="

' Method 1: Unfold with face parameter
WScript.Echo "Method 1: smDef.Unfold(largestFace)..."
smDef.Unfold largestFace

If Err.Number = 0 And smDef.HasFlatPattern Then
    WScript.Echo "SUCCESS with Unfold(face)!"
Else
    WScript.Echo "Failed: " & Err.Description
    Err.Clear
    
    ' Method 2: Simple Unfold
    WScript.Echo ""
    WScript.Echo "Method 2: smDef.Unfold..."
    smDef.Unfold
    
    If Err.Number = 0 And smDef.HasFlatPattern Then
        WScript.Echo "Created with simple Unfold"
    Else
        WScript.Echo "Failed: " & Err.Description
        Err.Clear
    End If
End If

partDoc.Update

' Check result
WScript.Echo ""
WScript.Echo "=== RESULT ==="
WScript.Echo "Has flat pattern: " & smDef.HasFlatPattern

If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo "Dimensions: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** ORIENTATION CORRECT! ***"
    Else
        WScript.Echo ""
        WScript.Echo "Edge view - need to fix orientation"
        
        ' Try FlipBaseFace
        WScript.Echo ""
        WScript.Echo "Trying FlipBaseFace..."
        fp.FlipBaseFace
        
        If Err.Number = 0 Then
            partDoc.Update
            WScript.Echo "After flip: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
        Else
            WScript.Echo "FlipBaseFace failed: " & Err.Description
            Err.Clear
        End If
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
