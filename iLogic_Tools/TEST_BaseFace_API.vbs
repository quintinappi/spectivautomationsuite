' TEST_BaseFace_API.vbs
' Explore all properties and methods related to base face in flat pattern
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, fp

WScript.Echo "=== EXPLORE BASE FACE API ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Part: " & partDoc.DisplayName

' Check if has flat pattern
If Not compDef.HasFlatPattern Then
    WScript.Echo "No flat pattern - creating one first..."
    compDef.Unfold
    partDoc.Update
End If

Set fp = compDef.FlatPattern

WScript.Echo ""
WScript.Echo "=== FLAT PATTERN OBJECT ==="
WScript.Echo "Type: " & TypeName(fp)
WScript.Echo "Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
WScript.Echo "Width: " & FormatNumber(fp.Width * 10, 1) & " mm"

' Try to access BaseFace property
WScript.Echo ""
WScript.Echo "=== TRYING BASE FACE PROPERTIES ==="

WScript.Echo "1. fp.BaseFace..."
Dim baseFace
Set baseFace = fp.BaseFace
If Err.Number = 0 And Not baseFace Is Nothing Then
    WScript.Echo "   Found! Type: " & TypeName(baseFace)
    WScript.Echo "   Area: " & FormatNumber(baseFace.Evaluator.Area * 100, 0) & " mm²"
Else
    WScript.Echo "   Error: " & Err.Description
    Err.Clear
End If

WScript.Echo "2. fp.Definition..."
Dim fpDef
Set fpDef = fp.Definition
If Err.Number = 0 And Not fpDef Is Nothing Then
    WScript.Echo "   Found! Type: " & TypeName(fpDef)
    
    ' Try Definition.BaseFace
    WScript.Echo "   fpDef.BaseFace..."
    Set baseFace = fpDef.BaseFace
    If Err.Number = 0 And Not baseFace Is Nothing Then
        WScript.Echo "      Found! Type: " & TypeName(baseFace)
    Else
        WScript.Echo "      Error: " & Err.Description
        Err.Clear
    End If
    
    ' Try Definition.StaticFace
    WScript.Echo "   fpDef.StaticFace..."
    Set baseFace = fpDef.StaticFace
    If Err.Number = 0 And Not baseFace Is Nothing Then
        WScript.Echo "      Found! Type: " & TypeName(baseFace)
    Else
        WScript.Echo "      Error: " & Err.Description
        Err.Clear
    End If
Else
    WScript.Echo "   Error: " & Err.Description
    Err.Clear
End If

WScript.Echo "3. fp.Parent..."
Dim parent
Set parent = fp.Parent
If Not parent Is Nothing Then
    WScript.Echo "   Type: " & TypeName(parent)
End If
Err.Clear

' Try accessing features
WScript.Echo ""
WScript.Echo "=== SHEET METAL FEATURES ==="

Dim smDef
Set smDef = compDef

WScript.Echo "Features count: " & smDef.Features.Count
Dim feat
For Each feat In smDef.Features
    WScript.Echo "  " & feat.Name & " (" & TypeName(feat) & ")"
    Err.Clear
Next

' Try FlatPatternFeature specifically
WScript.Echo ""
WScript.Echo "=== FLAT PATTERN FEATURE ==="

Dim fpFeature
For Each feat In smDef.Features
    If TypeName(feat) = "FlatPatternFeature" Or InStr(LCase(feat.Name), "flat") > 0 Then
        Set fpFeature = feat
        WScript.Echo "Found: " & feat.Name
        Exit For
    End If
    Err.Clear
Next

If Not fpFeature Is Nothing Then
    WScript.Echo "FlatPatternFeature properties:"
    
    WScript.Echo "  .BaseFace..."
    Set baseFace = fpFeature.BaseFace
    If Err.Number = 0 Then
        WScript.Echo "    Type: " & TypeName(baseFace)
    Else
        WScript.Echo "    Error: " & Err.Description
        Err.Clear
    End If
    
    WScript.Echo "  .Definition..."
    Set fpDef = fpFeature.Definition
    If Err.Number = 0 And Not fpDef Is Nothing Then
        WScript.Echo "    Type: " & TypeName(fpDef)
        
        WScript.Echo "    .BaseFace..."
        Set baseFace = fpDef.BaseFace
        If Err.Number = 0 Then
            WScript.Echo "      Type: " & TypeName(baseFace)
        Else
            WScript.Echo "      Error: " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "    Error: " & Err.Description
        Err.Clear
    End If
End If

' Look at SheetMetalComponentDefinition specific methods
WScript.Echo ""
WScript.Echo "=== SHEETMETAL COMPONENT DEFINITION ==="

WScript.Echo "smDef.UseSheetMetalStyleThickness: " & smDef.UseSheetMetalStyleThickness
Err.Clear

WScript.Echo "smDef.Thickness.Value: " & smDef.Thickness.Value * 10 & " mm"
Err.Clear

WScript.Echo "smDef.FlatPattern (same as compDef.FlatPattern): " & TypeName(smDef.FlatPattern)
Err.Clear

' Try to find Unfold method signature
WScript.Echo ""
WScript.Echo "=== EXPLORING UNFOLD ==="

' Delete flat pattern and try unfold with parameters
WScript.Echo "Deleting flat pattern..."
fp.Delete
partDoc.Update
WScript.Echo "Deleted"

' Find largest face
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
WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Try Unfold with face parameter
WScript.Echo ""
WScript.Echo "Trying smDef.Unfold(largestFace)..."
smDef.Unfold largestFace
If Err.Number = 0 Then
    WScript.Echo "SUCCESS!"
Else
    WScript.Echo "Error: " & Err.Description
    Err.Clear
End If

partDoc.Update

' Check result
If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo ""
    WScript.Echo "Flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
