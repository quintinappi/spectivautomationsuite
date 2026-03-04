' TEST_Convert_Feature.vbs
' Look for ConvertToSheetMetalFeature and its definition
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, cmdMgr

WScript.Echo "=== CONVERT TO SHEET METAL FEATURE ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager

WScript.Echo "Part: " & partDoc.DisplayName

' First, let's see what features exist
WScript.Echo ""
WScript.Echo "=== CURRENT FEATURES ==="

Dim feat
For Each feat In compDef.Features
    WScript.Echo "  " & feat.Name & " - " & TypeName(feat)
Next

' Delete flat pattern if exists
If compDef.HasFlatPattern Then
    compDef.FlatPattern.Delete
    partDoc.Update
    WScript.Echo ""
    WScript.Echo "Deleted flat pattern"
End If

' Try to access ConvertToSheetMetalFeatures
WScript.Echo ""
WScript.Echo "=== CONVERT TO SHEET METAL FEATURES ==="

Dim ctsm
Set ctsm = compDef.Features.ConvertToSheetMetalFeatures
If Not ctsm Is Nothing Then
    WScript.Echo "Found! Count: " & ctsm.Count
    
    ' List them
    Dim ctsmFeat
    For Each ctsmFeat In ctsm
        WScript.Echo "  " & ctsmFeat.Name
        
        ' Get definition
        Dim ctsmDef
        Set ctsmDef = ctsmFeat.Definition
        If Not ctsmDef Is Nothing Then
            WScript.Echo "    Definition type: " & TypeName(ctsmDef)
            WScript.Echo "    BaseFace: " & TypeName(ctsmDef.BaseFace)
            Err.Clear
        End If
    Next
Else
    WScript.Echo "Not found: " & Err.Description
End If
Err.Clear

' Check for ReferenceFeatures 
WScript.Echo ""
WScript.Echo "=== REFERENCE FEATURES ==="

Dim refFeats
Set refFeats = compDef.Features.ReferenceFeatures
If Not refFeats Is Nothing Then
    WScript.Echo "Found! Count: " & refFeats.Count
    
    Dim refFeat
    For Each refFeat In refFeats
        WScript.Echo "  " & refFeat.Name & " - " & TypeName(refFeat)
        
        ' Check definition
        Dim refDef
        Set refDef = refFeat.Definition
        If Not refDef Is Nothing Then
            WScript.Echo "    Definition: " & TypeName(refDef)
        End If
        Err.Clear
    Next
Else
    WScript.Echo "Not found: " & Err.Description
End If
Err.Clear

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
WScript.Echo ""
WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Let's try to use ConvertToSheetMetalFeatures.Add with a definition
WScript.Echo ""
WScript.Echo "=== TRY CREATE CONVERT FEATURE ==="

' First check if already sheet metal
WScript.Echo "Current SubType: " & compDef.SubType

If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Already sheet metal - try reverting first"
    
    ' Revert to standard part
    WScript.Echo "Reverting to standard part..."
    compDef.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}"
    partDoc.Update
    WScript.Echo "New SubType: " & compDef.SubType
    Err.Clear
End If

' Try to create ConvertToSheetMetalDefinition
WScript.Echo ""
WScript.Echo "Creating ConvertToSheetMetalDefinition..."

Dim ctsmDef2
Set ctsm = compDef.Features.ConvertToSheetMetalFeatures

If Not ctsm Is Nothing Then
    WScript.Echo "ConvertToSheetMetalFeatures found"
    
    ' Try CreateDefinition
    Set ctsmDef2 = ctsm.CreateDefinition(largestFace, 0.6) ' face and thickness
    If Not ctsmDef2 Is Nothing Then
        WScript.Echo "  Definition created!"
        WScript.Echo "  Type: " & TypeName(ctsmDef2)
        
        ' Add the feature
        WScript.Echo "  Adding feature..."
        Dim newFeat
        Set newFeat = ctsm.Add(ctsmDef2)
        If Not newFeat Is Nothing Then
            WScript.Echo "  Feature added!"
        Else
            WScript.Echo "  Add failed: " & Err.Description
        End If
        Err.Clear
    Else
        WScript.Echo "  CreateDefinition failed: " & Err.Description
    End If
    Err.Clear
    
    ' Try CreateDefinition with different params
    WScript.Echo ""
    WScript.Echo "Trying CreateDefinition()..."
    Set ctsmDef2 = ctsm.CreateDefinition()
    If Not ctsmDef2 Is Nothing Then
        WScript.Echo "  Created!"
        
        ' Set properties
        WScript.Echo "  Setting BaseFace..."
        ctsmDef2.BaseFace = largestFace
        WScript.Echo "  Result: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Failed: " & Err.Description
    End If
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
