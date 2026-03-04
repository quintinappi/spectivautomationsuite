' TEST_ConvertDefinition.vbs
' Look for ConvertToSheetMetalDefinition or similar
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, cmdMgr

WScript.Echo "=== EXPLORE CONVERT DEFINITIONS ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo ""

' Find largest face
Dim body, face, largestFace, largestArea
largestArea = 0
Set body = compDef.SurfaceBodies.Item(1)
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

' Check Features collections
WScript.Echo ""
WScript.Echo "=== FEATURES COLLECTIONS ==="

WScript.Echo "Features count: " & compDef.Features.Count

' List all feature collections
WScript.Echo ""
WScript.Echo "Looking for ConvertToSheetMetal in Features..."

Dim feat
For Each feat In compDef.Features
    WScript.Echo "  " & feat.Name & " - " & TypeName(feat)
Next
Err.Clear

' Try accessing specific feature types
WScript.Echo ""
WScript.Echo "=== SPECIFIC FEATURE COLLECTIONS ==="

Dim colls
colls = Array("ExtrudeFeatures", "RevolveFeatures", "ShellFeatures", _
              "FaceFeatures", "ThickenFeatures", "LoftFeatures", _
              "SheetMetalFeatures", "FlangeFeatures", "BendFeatures", _
              "ConvertToSheetMetalFeatures", "FlatPatternFeatures")

Dim collName, coll
For Each collName In colls
    Err.Clear
    Set coll = Nothing
    
    ' Use CallByName to access property
    On Error Resume Next
    Set coll = CallByName(compDef.Features, collName, 2)
    
    If Not coll Is Nothing Then
        WScript.Echo collName & ": Count = " & coll.Count
        Err.Clear
    Else
        ' Didn't work, skip
        Err.Clear
    End If
Next

' Try to access SheetMetalComponentDefinition specific
WScript.Echo ""
WScript.Echo "=== SHEET METAL DEFINITION ==="

Set smDef = compDef

WScript.Echo "Type: " & TypeName(smDef)
WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern
Err.Clear

' Try to access FlatPatternDefinition
WScript.Echo ""
WScript.Echo "smDef.FlatPatternDefinition..."
Dim fpDef
Set fpDef = smDef.FlatPatternDefinition
If Not fpDef Is Nothing Then
    WScript.Echo "  Found! Type: " & TypeName(fpDef)
Else
    WScript.Echo "  Not found"
End If
Err.Clear

' Try smDef.FlatPatternFeature
WScript.Echo "smDef.FlatPatternFeature..."
Dim fpFeat
Set fpFeat = smDef.FlatPatternFeature
If Not fpFeat Is Nothing Then
    WScript.Echo "  Found! Type: " & TypeName(fpFeat)
Else
    WScript.Echo "  Not found"
End If
Err.Clear

' Try smDef.FlatPatterns
WScript.Echo "smDef.FlatPatterns..."
Dim fps
Set fps = smDef.FlatPatterns
If Not fps Is Nothing Then
    WScript.Echo "  Found! Count: " & fps.Count
Else
    WScript.Echo "  Not found"
End If
Err.Clear

' Look for method to create flat pattern with face
WScript.Echo ""
WScript.Echo "=== TRY CREATE WITH FACE ==="

' Make sure no flat pattern
If smDef.HasFlatPattern Then
    smDef.FlatPattern.Delete
    partDoc.Update
    WScript.Echo "Deleted existing flat pattern"
End If

' Try various create methods
WScript.Echo ""
WScript.Echo "smDef.AddFlatPattern largestFace..."
smDef.AddFlatPattern largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

WScript.Echo "smDef.CreateFlatPattern largestFace..."
smDef.CreateFlatPattern largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

WScript.Echo "smDef.Flatten largestFace..."
smDef.Flatten largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Try Features.FlatPatternFeatures.Add
WScript.Echo ""
WScript.Echo "compDef.Features.FlatPatternFeatures.Add largestFace..."
Dim fpFeats
Set fpFeats = compDef.Features.FlatPatternFeatures
If Not fpFeats Is Nothing Then
    fpFeats.Add largestFace
    WScript.Echo "  Result: " & Err.Description
Else
    WScript.Echo "  FlatPatternFeatures not found"
End If
Err.Clear

' Create without face
WScript.Echo ""
WScript.Echo "Creating flat pattern normally..."
smDef.Unfold
partDoc.Update

If smDef.HasFlatPattern Then
    WScript.Echo "Created! Dimensions: " & FormatNumber(smDef.FlatPattern.Length * 10, 1) & " x " & FormatNumber(smDef.FlatPattern.Width * 10, 1) & " mm"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
