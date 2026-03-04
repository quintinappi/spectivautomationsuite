' TEST_ConvertToSheetMetal_API.vbs
' Deep exploration of ConvertToSheetMetalFeatures API
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef

WScript.Echo "=== CONVERT TO SHEET METAL API EXPLORATION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition

WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' If already sheet metal, delete the conversion first
If compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Part is already sheet metal"
    WScript.Echo "Deleting flat pattern and conversion feature..."
    
    ' Delete flat pattern first
    If compDef.HasFlatPattern Then
        compDef.FlatPattern.Delete
        WScript.Echo "  Flat pattern deleted"
    End If
    
    ' Find and delete the ConvertToSheetMetal feature
    Dim feat
    For Each feat In compDef.Features
        If TypeName(feat) = "ConvertToSheetMetalFeature" Then
            WScript.Echo "  Found ConvertToSheetMetalFeature"
            feat.Delete
            WScript.Echo "  ConvertToSheetMetalFeature deleted"
            Exit For
        End If
    Next
    
    partDoc.Update
    Set compDef = partDoc.ComponentDefinition
    WScript.Echo ""
End If

WScript.Echo "SubType GUID: " & compDef.SubType
WScript.Echo ""

' Find the largest face
Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

WScript.Echo "=== ANALYZING FACES ==="
Dim faceIndex
faceIndex = 0
For Each face In faces
    faceIndex = faceIndex + 1
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number = 0 Then
        WScript.Echo "Face " & faceIndex & ": " & FormatNumber(area, 0) & " mm²"
        If area > largestArea Then
            largestArea = area
            Set largestFace = face
        End If
    End If
    Err.Clear
Next

WScript.Echo ""
WScript.Echo "Largest face has " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""

' Check what's available in SheetMetalComponentDefinition
WScript.Echo "=== SHEET METAL COMPONENT DEFINITION ACCESS ==="

' Try to cast/access SheetMetalComponentDefinition features
Dim smFeatures
Set smFeatures = compDef.Features.ConvertToSheetMetalFeatures

If Err.Number <> 0 Then
    WScript.Echo "ConvertToSheetMetalFeatures: " & Err.Description
    Err.Clear
Else
    WScript.Echo "ConvertToSheetMetalFeatures Type: " & TypeName(smFeatures)
    WScript.Echo "ConvertToSheetMetalFeatures Count: " & smFeatures.Count
    
    ' Try CreateDefinition with face parameter
    WScript.Echo ""
    WScript.Echo "=== TRYING CREATE DEFINITION ==="
    
    Dim cstDef
    Set cstDef = smFeatures.CreateDefinition(largestFace)
    
    If Err.Number <> 0 Then
        WScript.Echo "CreateDefinition(face): " & Err.Description
        Err.Clear
        
        ' Try without parameters
        Set cstDef = smFeatures.CreateDefinition()
        
        If Err.Number <> 0 Then
            WScript.Echo "CreateDefinition(): " & Err.Description
            Err.Clear
        Else
            WScript.Echo "CreateDefinition() Type: " & TypeName(cstDef)
        End If
    Else
        WScript.Echo "CreateDefinition(face) Type: " & TypeName(cstDef)
    End If
    
    If Not cstDef Is Nothing Then
        WScript.Echo ""
        WScript.Echo "=== DEFINITION PROPERTIES ==="
        
        WScript.Echo "Trying to set BaseFace..."
        Set cstDef.BaseFace = largestFace
        If Err.Number <> 0 Then
            WScript.Echo "  Set BaseFace: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  BaseFace SET SUCCESSFULLY!"
        End If
        
        WScript.Echo "Trying to access Thickness..."
        Dim thick
        thick = cstDef.Thickness
        If Err.Number <> 0 Then
            WScript.Echo "  Thickness: " & Err.Description
            Err.Clear
        Else
            WScript.Echo "  Thickness: " & thick
        End If
    End If
End If

' Try Add method
WScript.Echo ""
WScript.Echo "=== TRYING ADD METHOD ==="

If Not smFeatures Is Nothing Then
    If smFeatures.Count = 0 And Not cstDef Is Nothing And TypeName(cstDef) <> "Empty" Then
        WScript.Echo "Calling smFeatures.Add(cstDef)..."
        
        Dim cstFeat
        Set cstFeat = smFeatures.Add(cstDef)
        
        If Err.Number <> 0 Then
            WScript.Echo "Add(): " & Err.Description
            Err.Clear
        Else
            WScript.Echo "Add() successful! Type: " & TypeName(cstFeat)
        End If
    Else
        WScript.Echo "Cannot add - smFeatures.Count=" & smFeatures.Count & " or cstDef is invalid"
    End If
End If

WScript.Echo ""
WScript.Echo "=== FINAL STATE ==="
Set compDef = partDoc.ComponentDefinition
WScript.Echo "SubType: " & compDef.SubType
WScript.Echo "Is Sheet Metal: " & (compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")

WScript.Echo ""
WScript.Echo "=== DONE ==="
