On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== COMPARING PART2 vs PART3 FACE GEOMETRY ==="
WScript.Echo ""
WScript.Echo "Part: " & doc.DisplayName

Dim compDef
Set compDef = doc.ComponentDefinition

Dim faces
Set faces = compDef.SurfaceBodies.Item(1).Faces

WScript.Echo "Total Faces: " & faces.Count
WScript.Echo ""

' Analyze each face
Dim i, face, area, faceType
For i = 1 To faces.Count
    Set face = faces.Item(i)
    area = Round(face.Evaluator.Area * 100, 0)
    faceType = face.SurfaceType
    
    WScript.Echo "Face " & i & ":"
    WScript.Echo "  Area: " & area & " mm²"
    WScript.Echo "  Type: " & faceType
    
    ' Get geometry details
    If faceType = 3 Then ' Plane
        Dim plane
        Set plane = face.Geometry
        WScript.Echo "  Geometry: Planar"
    End If
    
    WScript.Echo ""
Next

' Find largest
Dim largestArea, largestIndex
largestArea = 0
largestIndex = 0

For i = 1 To faces.Count
    Set face = faces.Item(i)
    area = face.Evaluator.Area * 100
    
    If area > largestArea Then
        largestArea = area
        largestIndex = i
    End If
Next

WScript.Echo "=== LARGEST FACE ==="
WScript.Echo "Face " & largestIndex & " with area " & Round(largestArea, 0) & " mm²"
