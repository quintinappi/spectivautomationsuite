' TEST_Select_Large_Face.vbs
' Select the large face for the waiting dialog
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet

WScript.Echo "=== SELECT LARGE FACE ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet

WScript.Echo "Part: " & partDoc.DisplayName

' Get all faces
Dim body, faces, face
Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

WScript.Echo "Total faces: " & faces.Count

' Find the largest face
Dim largestFace, largestArea, faceNum, bestFaceNum
largestArea = 0
faceNum = 0

For Each face In faces
    faceNum = faceNum + 1
    
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number <> 0 Then
        area = 0
        Err.Clear
    End If
    
    WScript.Echo "Face " & faceNum & ": " & FormatNumber(area, 0) & " mm²"
    
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
        bestFaceNum = faceNum
    End If
Next

WScript.Echo ""
WScript.Echo "Largest is Face " & bestFaceNum & " (" & FormatNumber(largestArea, 0) & " mm²)"

' Clear selection and select the face
WScript.Echo ""
WScript.Echo "Selecting face..."

selectSet.Clear

' Try different selection methods
WScript.Echo "Method 1: selectSet.Select(face)..."
selectSet.Select largestFace
If Err.Number = 0 And selectSet.Count > 0 Then
    WScript.Echo "Success! Count: " & selectSet.Count
Else
    WScript.Echo "Failed: " & Err.Description
    Err.Clear
    
    ' Method 2: Use SelectSet.SelectAdd
    WScript.Echo "Method 2: selectSet.SelectAdd..."
    selectSet.Clear
    selectSet.SelectAdd largestFace
    If Err.Number = 0 And selectSet.Count > 0 Then
        WScript.Echo "Success! Count: " & selectSet.Count
    Else
        WScript.Echo "Failed: " & Err.Description
        Err.Clear
    End If
End If

' Method 3: Use Highlight
WScript.Echo ""
WScript.Echo "Method 3: Highlighting face..."
largestFace.Highlighted = True
If Err.Number = 0 Then
    WScript.Echo "Face highlighted!"
Else
    WScript.Echo "Highlight failed: " & Err.Description
    Err.Clear
End If

' Method 4: Use InteractionEvents to simulate click
WScript.Echo ""
WScript.Echo "Method 4: Using HighlightSet..."

Dim highlightSet
Set highlightSet = partDoc.HighlightSets.Add
highlightSet.AddItem largestFace
If Err.Number = 0 Then
    WScript.Echo "Added to HighlightSet!"
Else
    WScript.Echo "HighlightSet failed: " & Err.Description
    Err.Clear
End If

WScript.Echo ""
WScript.Echo "Current selection count: " & selectSet.Count
WScript.Echo ""
WScript.Echo ">>> Please click Face " & bestFaceNum & " in Inventor <<<"
WScript.Echo ">>> (The large flat face on top or bottom) <<<"
WScript.Echo ""
WScript.Echo "=== DONE ==="
