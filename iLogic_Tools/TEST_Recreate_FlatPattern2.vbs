' TEST_Recreate_FlatPattern2.vbs
' Open Part2 and recreate flat pattern
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp
Dim partPath

WScript.Echo "=== RECREATE FLAT PATTERN ==="
WScript.Echo ""

' Connect to Inventor
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

' Open the part file
partPath = "C:\Users\Quintin\Documents\Spectiv\3. Working\DM Underpan\Part2 DM-UP.ipt"
WScript.Echo "Opening: " & partPath

Set partDoc = invApp.Documents.Open(partPath, False)
If Err.Number <> 0 Then
    WScript.Echo "Open failed: " & Err.Description
    WScript.Quit
End If

WScript.Echo "Part opened: " & partDoc.DisplayName

Set compDef = partDoc.ComponentDefinition

' Check if sheet metal
WScript.Echo "Component type: " & compDef.Type
WScript.Echo "(Expected 99588099 for SheetMetalComponentDefinition)"

If compDef.Type <> 99588099 Then
    WScript.Echo "Not a sheet metal part"
    WScript.Quit
End If

Set smDef = compDef
WScript.Echo "Sheet metal part confirmed"

' Check for existing flat pattern
WScript.Echo ""
WScript.Echo "Has flat pattern: " & smDef.HasFlatPattern

If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo "Current flat pattern:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    ' Find the largest face BEFORE deleting
    WScript.Echo ""
    WScript.Echo "Finding largest face on folded model..."
    
    Dim body, faces, face
    Dim largestFace, largestArea
    largestArea = 0
    
    Set body = smDef.SurfaceBodies.Item(1)
    Set faces = body.Faces
    
    WScript.Echo "Total faces: " & faces.Count
    
    Dim faceNum, faceAreas
    faceAreas = ""
    faceNum = 0
    For Each face In faces
        faceNum = faceNum + 1
        If face.SurfaceType = 3 Then ' kPlaneSurface
            Dim area
            area = face.Evaluator.Area * 100 ' Convert to mm²
            faceAreas = faceAreas & "  Face " & faceNum & ": " & FormatNumber(area, 0) & " mm²" & vbCrLf
            
            If area > largestArea Then
                largestArea = area
                Set largestFace = face
            End If
        End If
    Next
    
    WScript.Echo "Plane faces:"
    WScript.Echo faceAreas
    WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Store reference to the face
    WScript.Echo ""
    WScript.Echo "Deleting flat pattern..."
    
    fp.Delete
    
    If Err.Number = 0 Then
        WScript.Echo "Flat pattern deleted!"
    Else
        WScript.Echo "Delete failed: " & Err.Description
        Err.Clear
        
        ' Try alternative delete method
        WScript.Echo "Trying Fold..."
        smDef.Fold
        
        If Err.Number = 0 Then
            WScript.Echo "Fold succeeded!"
        Else
            WScript.Echo "Fold failed: " & Err.Description
            Err.Clear
        End If
    End If
    
    partDoc.Update
    WScript.Echo "Has flat pattern after delete: " & smDef.HasFlatPattern
    
    ' Now recreate
    WScript.Echo ""
    WScript.Echo "=== RECREATING FLAT PATTERN ==="
    
    ' Get fresh reference to the body and largest face after fold
    Set body = smDef.SurfaceBodies.Item(1)
    Set faces = body.Faces
    
    largestArea = 0
    Set largestFace = Nothing
    
    For Each face In faces
        If face.SurfaceType = 3 Then
            area = face.Evaluator.Area * 100
            If area > largestArea Then
                largestArea = area
                Set largestFace = face
            End If
        End If
    Next
    
    WScript.Echo "Largest face (fresh): " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Try Unfold with face
    WScript.Echo ""
    WScript.Echo "Trying Unfold(face)..."
    smDef.Unfold largestFace
    
    If Err.Number = 0 Then
        WScript.Echo "Unfold with face succeeded!"
    Else
        WScript.Echo "Failed: " & Err.Description
        Err.Clear
        
        ' Simple Unfold
        WScript.Echo ""
        WScript.Echo "Trying simple Unfold..."
        smDef.Unfold
        
        If Err.Number = 0 Then
            WScript.Echo "Simple Unfold succeeded!"
        Else
            WScript.Echo "Failed: " & Err.Description
            Err.Clear
        End If
    End If
    
    partDoc.Update
End If

' Check final result
WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="
WScript.Echo "Has flat pattern: " & smDef.HasFlatPattern

If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo "Flat pattern:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** SUCCESS! Orientation looks correct! ***"
    Else
        WScript.Echo ""
        WScript.Echo "Still showing edge view - 6mm dimension is the thickness"
    End If
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
