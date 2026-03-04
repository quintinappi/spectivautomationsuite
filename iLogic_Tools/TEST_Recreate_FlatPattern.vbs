' TEST_Recreate_FlatPattern.vbs
' Delete and recreate flat pattern with correct base face
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef, fp
Dim selectSet, cmdMgr

WScript.Echo "=== RECREATE FLAT PATTERN ==="
WScript.Echo ""

' Connect to Inventor
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
If partDoc Is Nothing Or partDoc.DocumentType <> 12291 Then
    WScript.Echo "Please open a part document"
    WScript.Quit
End If

WScript.Echo "Part: " & partDoc.DisplayName

Set compDef = partDoc.ComponentDefinition

' Check if sheet metal
If compDef.Type <> 99588099 Then ' kSheetMetalComponentDefinitionObject
    WScript.Echo "Not a sheet metal part"
    WScript.Quit
End If

Set smDef = compDef
WScript.Echo "Sheet metal part confirmed"

' Check for existing flat pattern
If smDef.HasFlatPattern Then
    Set fp = smDef.FlatPattern
    WScript.Echo ""
    WScript.Echo "Current flat pattern dimensions:"
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
    
    Dim faceNum
    faceNum = 0
    For Each face In faces
        faceNum = faceNum + 1
        If face.SurfaceType = 3 Then ' kPlaneSurface
            Dim area
            area = face.Evaluator.Area * 100 ' Convert to mm²
            
            If area > largestArea Then
                largestArea = area
                Set largestFace = face
            End If
        End If
    Next
    
    WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Delete the flat pattern
    WScript.Echo ""
    WScript.Echo "Deleting flat pattern..."
    smDef.FlatPattern.Delete
    
    If Err.Number = 0 Then
        WScript.Echo "Flat pattern deleted!"
    Else
        WScript.Echo "Delete failed: " & Err.Description
        Err.Clear
    End If
    
    partDoc.Update
    WScript.Echo "Part updated"
    
    ' Now recreate with the largest face
    WScript.Echo ""
    WScript.Echo "Creating new flat pattern with largest face..."
    
    ' Try Unfold method with face parameter
    WScript.Echo "Method 1: Unfold with face parameter..."
    smDef.Unfold largestFace
    
    If Err.Number = 0 Then
        WScript.Echo "Unfold succeeded!"
    Else
        WScript.Echo "Unfold failed: " & Err.Description
        Err.Clear
        
        ' Try CreateFlatPattern with face parameter
        WScript.Echo ""
        WScript.Echo "Method 2: CreateFlatPattern with face..."
        
        ' Check for CreateFlatPattern method
        Dim fpDefs
        Set fpDefs = smDef.FlatPatternDefinitions
        
        If Not fpDefs Is Nothing Then
            Dim fpDef
            Set fpDef = fpDefs.Add(largestFace)
            
            If Err.Number = 0 Then
                WScript.Echo "FlatPatternDefinition created!"
            Else
                WScript.Echo "FlatPatternDefinitions.Add failed: " & Err.Description
                Err.Clear
            End If
        Else
            WScript.Echo "No FlatPatternDefinitions collection"
            Err.Clear
        End If
        
        ' Method 3: Just Unfold without parameters
        WScript.Echo ""
        WScript.Echo "Method 3: Unfold without parameters..."
        smDef.Unfold
        
        If Err.Number = 0 Then
            WScript.Echo "Unfold succeeded!"
        Else
            WScript.Echo "Unfold failed: " & Err.Description
            Err.Clear
        End If
    End If
    
    partDoc.Update
    
    ' Check result
    WScript.Echo ""
    WScript.Echo "=== RESULT ==="
    
    If smDef.HasFlatPattern Then
        Set fp = smDef.FlatPattern
        WScript.Echo "Flat pattern dimensions:"
        WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
        WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
        
        If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
            WScript.Echo ""
            WScript.Echo "SUCCESS! Orientation looks correct!"
        Else
            WScript.Echo ""
            WScript.Echo "Still showing edge view"
        End If
    Else
        WScript.Echo "No flat pattern exists"
    End If
Else
    WScript.Echo "No flat pattern to delete"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
