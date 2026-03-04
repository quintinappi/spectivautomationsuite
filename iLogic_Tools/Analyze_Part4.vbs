On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== DETAILED PART ANALYSIS ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition

Dim faces
Set faces = compDef.SurfaceBodies.Item(1).Faces

WScript.Echo "Total Faces: " & faces.Count
WScript.Echo ""

' Show each face with detailed info
Dim face, i, area
For i = 1 To faces.Count
    Set face = faces.Item(i)
    area = Round(face.Evaluator.Area * 100, 0)
    
    WScript.Echo "Face " & i & ": " & FormatNumber(area, 0) & " mm²"
    
    ' Try to get face normal/geometry
    If face.SurfaceType = 3 Then ' Plane
        Dim plane
        Set plane = face.Geometry
        
        Dim normal
        Set normal = plane.Normal
        
        WScript.Echo "  Normal: (" & Round(normal.X, 3) & ", " & Round(normal.Y, 3) & ", " & Round(normal.Z, 3) & ")"
    End If
Next

WScript.Echo ""
WScript.Echo "=== EXPECTED RESULT ==="
WScript.Echo "Largest face should be base: ~7,546,148 mm²"
WScript.Echo ""

' Check current state
If doc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
    WScript.Echo "Current state: Sheet Metal"
    
    Dim flatPattern
    Set flatPattern = compDef.FlatPattern
    
    If Not flatPattern Is Nothing Then
        Dim baseArea
        If Not flatPattern.BaseFace Is Nothing Then
            baseArea = Round(flatPattern.BaseFace.Evaluator.Area * 100, 0)
            WScript.Echo "Current BaseFace: " & FormatNumber(baseArea, 0) & " mm²"
            
            If baseArea > 5000000 Then
                WScript.Echo "*** CORRECT ***"
            Else
                WScript.Echo "*** WRONG - Should be 7,546,148 mm² ***"
            End If
        End If
    End If
Else
    WScript.Echo "Current state: Standard Part"
End If
