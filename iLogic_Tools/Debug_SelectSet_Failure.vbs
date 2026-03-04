On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== DIAGNOSING SELECTSET FAILURE ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition

' Get faces
Dim faces, face, i
Set faces = compDef.SurfaceBodies.Item(1).Faces

WScript.Echo "Total faces: " & faces.Count
WScript.Echo ""

' Try selecting each face individually
Dim selectSet
Set selectSet = doc.SelectSet

For i = 1 To faces.Count
    Set face = faces.Item(i)
    
    WScript.Echo "Face " & i & ":"
    WScript.Echo "  Area: " & FormatNumber(Round(face.Evaluator.Area * 100, 0), 0) & " mm²"
    
    ' Clear and try to select
    selectSet.Clear
    Err.Clear
    
    selectSet.Select face
    
    If Err.Number <> 0 Then
        WScript.Echo "  Selection FAILED: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  SelectSet.Count after Select: " & selectSet.Count
        
        If selectSet.Count > 0 Then
            WScript.Echo "  Selection SUCCESS"
        Else
            WScript.Echo "  Selection FAILED - Count is 0"
        End If
    End If
    WScript.Echo ""
Next

WScript.Echo "=== TESTING ALTERNATIVE METHOD ==="
WScript.Echo ""

' Try AddWithTest
selectSet.Clear

Dim largestFace, largestArea, area
largestArea = 0

For Each face In faces
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
WScript.Echo ""
WScript.Echo "Trying AddWithTest method..."

Err.Clear
Dim filter
Set filter = invApp.TransientObjects.CreateObjectCollection

selectSet.AddWithTest largestFace, filter

If Err.Number <> 0 Then
    WScript.Echo "AddWithTest FAILED: " & Err.Description
Else
    WScript.Echo "AddWithTest called. SelectSet.Count = " & selectSet.Count
End If
