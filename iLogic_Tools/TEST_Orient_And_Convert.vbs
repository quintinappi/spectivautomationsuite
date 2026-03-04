' TEST_Orient_And_Convert.vbs
' Zoom extents, orient to largest face, then convert with click
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, cmdMgr, selectSet, WshShell, camera

WScript.Echo "=== ORIENT AND CONVERT ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set WshShell = CreateObject("WScript.Shell")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set cmdMgr = invApp.CommandManager
Set selectSet = partDoc.SelectSet

WScript.Echo "Part: " & partDoc.DisplayName

' Step 1: Find largest face and get its normal
WScript.Echo ""
WScript.Echo "=== STEP 1: FIND LARGEST FACE ==="

Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number <> 0 Then
        area = 0
        Err.Clear
    End If
    
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Step 2: Zoom to fit
WScript.Echo ""
WScript.Echo "=== STEP 2: ZOOM EXTENTS ==="

Dim view
Set view = invApp.ActiveView

' Zoom to fit
view.Fit
WScript.Echo "Zoomed to fit"

' Step 3: Orient view to look at the largest face
WScript.Echo ""
WScript.Echo "=== STEP 3: ORIENT VIEW TO FACE ==="

' Get face normal at center point
Dim evaluator, paramRange, uMin, uMax, vMin, vMax, uMid, vMid
Dim params(1), normal, point

Set evaluator = largestFace.Evaluator

' Get parameter range
Dim minParams, maxParams
evaluator.ParamRangeRect minParams, maxParams

uMid = (minParams(0) + maxParams(0)) / 2
vMid = (minParams(1) + maxParams(1)) / 2

params(0) = uMid
params(1) = vMid

' Get normal at center of face
Dim normals, points
evaluator.GetNormals 1, params, normals
evaluator.GetPointsAtParams 1, params, points

If Err.Number = 0 Then
    WScript.Echo "Face center point: " & FormatNumber(points(0)*10, 1) & ", " & FormatNumber(points(1)*10, 1) & ", " & FormatNumber(points(2)*10, 1)
    WScript.Echo "Face normal: " & FormatNumber(normals(0), 3) & ", " & FormatNumber(normals(1), 3) & ", " & FormatNumber(normals(2), 3)
    
    ' Set camera to look at face
    Set camera = view.Camera
    
    ' Position camera along the normal direction
    Dim eyeX, eyeY, eyeZ, targetX, targetY, targetZ, dist
    dist = 50 ' 500mm distance from center
    
    targetX = points(0)
    targetY = points(1)
    targetZ = points(2)
    
    eyeX = targetX + normals(0) * dist
    eyeY = targetY + normals(1) * dist
    eyeZ = targetZ + normals(2) * dist
    
    ' Set camera properties
    Dim tg, transGeom
    Set transGeom = invApp.TransientGeometry
    
    camera.Eye = transGeom.CreatePoint(eyeX, eyeY, eyeZ)
    camera.Target = transGeom.CreatePoint(targetX, targetY, targetZ)
    camera.UpVector = transGeom.CreateUnitVector(0, 1, 0)
    camera.ApplyWithoutTransition
    
    WScript.Echo "Camera oriented to face"
Else
    WScript.Echo "Could not get face normal: " & Err.Description
    Err.Clear
End If

' Zoom to fit again after reorientation
view.Fit
WScript.Echo "Zoomed to fit"

WScript.Sleep 500

' Step 4: Pre-select face and run convert command
WScript.Echo ""
WScript.Echo "=== STEP 4: RUN CONVERT COMMAND ==="

Dim convertCmd
Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

If convertCmd Is Nothing Or Not convertCmd.Enabled Then
    WScript.Echo "Convert command not available"
    WScript.Quit
End If

' Pre-select face
selectSet.Clear
selectSet.Select largestFace
WScript.Echo "Face pre-selected"

' Execute command
WScript.Echo "Executing PartConvertToSheetMetalCmd..."
convertCmd.Execute
WScript.Sleep 500

' Re-select face
selectSet.Clear
selectSet.Select largestFace
WScript.Sleep 200

WScript.Echo ""
WScript.Echo ">>> NOW RUN ClickInventor.ps1 TO CLICK THE FACE <<<"
WScript.Echo ""
WScript.Echo "=== WAITING ==="
