On Error Resume Next

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
Dim doc
Set doc = invApp.ActiveDocument

WScript.Echo "=== TESTING CLICK SIMULATION ==="
WScript.Echo "Part: " & doc.DisplayName
WScript.Echo ""

Dim compDef
Set compDef = doc.ComponentDefinition

' Find largest face
Dim faces, face, largestFace, largestArea, area
Set faces = compDef.SurfaceBodies.Item(1).Faces
largestArea = 0

For Each face In faces
    area = face.Evaluator.Area * 100
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"

' Get face centroid
Dim evaluator
Set evaluator = largestFace.Evaluator

Dim point
Set point = evaluator.Centroid

WScript.Echo "Centroid:"
WScript.Echo "  X: " & point.X
WScript.Echo "  Y: " & point.Y  
WScript.Echo "  Z: " & point.Z
WScript.Echo ""

' Convert to screen coordinates
Dim camera
Set camera = invApp.ActiveView.Camera

Dim screenPt
Set screenPt = camera.WorldToViewCoordinates(point)

WScript.Echo "View Coordinates:"
WScript.Echo "  X: " & screenPt.X
WScript.Echo "  Y: " & screenPt.Y
WScript.Echo ""

WScript.Echo "This approach would require mouse click simulation,"
WScript.Echo "which you previously said was 'too fucked up'."
WScript.Echo ""
WScript.Echo "Alternative: Semi-automated approach where script"
WScript.Echo "pre-selects the face, starts the command, and pauses"
WScript.Echo "for you to click once on the highlighted green face."
