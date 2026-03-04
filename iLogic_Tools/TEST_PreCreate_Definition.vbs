' TEST_PreCreate_Definition.vbs
' Try to create a FlatPatternDefinition with face BEFORE creating flat pattern
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef

WScript.Echo "=== PRE-CREATE DEFINITION TEST ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set smDef = compDef

WScript.Echo "Part: " & partDoc.DisplayName

' Delete any existing flat pattern
If smDef.HasFlatPattern Then
    smDef.FlatPattern.Delete
    partDoc.Update
    WScript.Echo "Deleted existing flat pattern"
End If

' Find largest face
Dim body, face, largestFace, largestArea
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

' Try to use TransientObjects to create a definition
WScript.Echo ""
WScript.Echo "=== TRYING TRANSIENT OBJECTS ==="

Dim transObj
Set transObj = invApp.TransientObjects
WScript.Echo "TransientObjects: " & TypeName(transObj)

' Try CreateFlatPatternDefinition
WScript.Echo ""
WScript.Echo "transObj.CreateFlatPatternDefinition(largestFace)..."
Dim fpDef
Set fpDef = transObj.CreateFlatPatternDefinition(largestFace)
If Not fpDef Is Nothing Then
    WScript.Echo "  Created! Type: " & TypeName(fpDef)
Else
    WScript.Echo "  Error: " & Err.Description
End If
Err.Clear

' Try with no args first
WScript.Echo "transObj.CreateFlatPatternDefinition()..."
Set fpDef = transObj.CreateFlatPatternDefinition()
If Not fpDef Is Nothing Then
    WScript.Echo "  Created! Type: " & TypeName(fpDef)
    
    ' Try to set base face on definition
    WScript.Echo "  Setting BaseFace..."
    Set fpDef.BaseFace = largestFace
    WScript.Echo "  Result: " & Err.Description
    Err.Clear
    
    fpDef.BaseFace = largestFace
    WScript.Echo "  Result2: " & Err.Description
    Err.Clear
Else
    WScript.Echo "  Error: " & Err.Description
End If
Err.Clear

' Check Features.FlatPatternFeatures for CreateDefinition
WScript.Echo ""
WScript.Echo "=== FLATPATTERN FEATURES ==="

Dim fpFeatures
Set fpFeatures = smDef.Features.FlatPatternFeatures
If Not fpFeatures Is Nothing Then
    WScript.Echo "Found FlatPatternFeatures"
    WScript.Echo "Count: " & fpFeatures.Count
    
    ' Try CreateDefinition
    WScript.Echo ""
    WScript.Echo "fpFeatures.CreateDefinition()..."
    Set fpDef = fpFeatures.CreateDefinition()
    If Not fpDef Is Nothing Then
        WScript.Echo "  Created! Type: " & TypeName(fpDef)
    Else
        WScript.Echo "  Error: " & Err.Description
    End If
    Err.Clear
    
    WScript.Echo "fpFeatures.CreateDefinition(largestFace)..."
    Set fpDef = fpFeatures.CreateDefinition(largestFace)
    If Not fpDef Is Nothing Then
        WScript.Echo "  Created! Type: " & TypeName(fpDef)
    Else
        WScript.Echo "  Error: " & Err.Description
    End If
    Err.Clear
Else
    WScript.Echo "FlatPatternFeatures not found: " & Err.Description
End If
Err.Clear

' Look at the Inventor type library for FlatPattern
WScript.Echo ""
WScript.Echo "=== INVENTOR TYPE INFO ==="

' Get FlatPattern object and inspect
smDef.Unfold
partDoc.Update

Dim fp
Set fp = smDef.FlatPattern

WScript.Echo "FlatPattern object methods/properties:"

' Check for Redefine method
WScript.Echo "fp.Redefine..."
fp.Redefine
WScript.Echo "  Result: " & Err.Description
Err.Clear

WScript.Echo "fp.Redefine largestFace..."
fp.Redefine largestFace
WScript.Echo "  Result: " & Err.Description
Err.Clear

' Check SetEndOfPart position influence
WScript.Echo ""
WScript.Echo "Checking if face order matters..."
WScript.Echo "fp.BaseFace index or info..."

Dim baseFaceInBody, idx
idx = 0
For Each face In body.Faces
    idx = idx + 1
    If face Is fp.BaseFace Then
        WScript.Echo "BaseFace is Face #" & idx & " in body"
        Exit For
    End If
Next

' Find which index is the largest face
idx = 0
For Each face In body.Faces
    idx = idx + 1
    If face Is largestFace Then
        WScript.Echo "LargestFace is Face #" & idx & " in body"
        Exit For
    End If
Next

WScript.Echo ""
WScript.Echo "Final flat pattern: " & FormatNumber(fp.Length * 10, 1) & " x " & FormatNumber(fp.Width * 10, 1) & " mm"

WScript.Echo ""
WScript.Echo "=== DONE ==="
