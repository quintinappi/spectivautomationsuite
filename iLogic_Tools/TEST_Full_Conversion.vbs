' TEST_Full_Conversion.vbs
' Full workflow: Break link -> Convert to Sheet Metal -> Create Flat Pattern
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, smDef

WScript.Echo "=== FULL CONVERSION WORKFLOW ==="
WScript.Echo ""

' Connect to Inventor
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "Cannot connect to Inventor"
    WScript.Quit
End If

Set partDoc = invApp.ActiveDocument
WScript.Echo "Part: " & partDoc.DisplayName
WScript.Echo ""

' Step 1: Check current state
WScript.Echo "=== STEP 1: CHECK CURRENT STATE ==="
Set compDef = partDoc.ComponentDefinition

' For derived parts, use partDoc directly
Dim subType
On Error Resume Next
subType = compDef.SubType
If Err.Number <> 0 Then
    Err.Clear
    ' Try alternative access
    subType = partDoc.ComponentDefinition.SubType
    If Err.Number <> 0 Then
        WScript.Echo "Cannot get SubType, checking features..."
        Err.Clear
    End If
End If

WScript.Echo "SubType: " & subType

' Check if it's already sheet metal
Dim isSheetMetal
isSheetMetal = (subType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

' Check for derived feature
WScript.Echo ""
WScript.Echo "Checking for derived feature..."

Dim feat, isDerived, derivedFeature
isDerived = False
For Each feat In compDef.Features
    If TypeName(feat) = "ReferenceFeature" Or InStr(feat.Name, "::") > 0 Then
        isDerived = True
        Set derivedFeature = feat
        WScript.Echo "Found derived reference: " & feat.Name
        Exit For
    End If
Next
If Err.Number <> 0 Then Err.Clear

If Not isDerived Then
    WScript.Echo "Not a derived part"
End If

' Step 2: Break derived link if needed
If isDerived Then
    WScript.Echo ""
    WScript.Echo "=== STEP 2: BREAK DERIVED LINK ==="
    
    Dim cmdMgr, breakCmd
    Set cmdMgr = invApp.CommandManager
    Set breakCmd = cmdMgr.ControlDefinitions.Item("PartBreakLinkDerivedPartCtxCmd")
    
    If Not breakCmd Is Nothing And breakCmd.Enabled Then
        WScript.Echo "Executing PartBreakLinkDerivedPartCtxCmd..."
        breakCmd.Execute
        
        If Err.Number = 0 Then
            WScript.Echo "Link broken!"
            partDoc.Update
        Else
            WScript.Echo "Failed: " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "Break link command not available"
    End If
    
    ' Re-check SubType after breaking
    WScript.Sleep 500
    Set compDef = partDoc.ComponentDefinition
    subType = compDef.SubType
    WScript.Echo "SubType after break: " & subType
End If

' Step 3: Convert to Sheet Metal if not already
WScript.Echo ""
WScript.Echo "=== STEP 3: CONVERT TO SHEET METAL ==="

isSheetMetal = (subType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
WScript.Echo "Is Sheet Metal: " & isSheetMetal

If Not isSheetMetal Then
    WScript.Echo "Converting to sheet metal..."
    
    ' Detect thickness from geometry
    Dim thickness
    thickness = 0.6 ' Default 6mm in cm
    
    ' Get bounding box dimensions
    Dim rBox, dims(2)
    Set rBox = compDef.RangeBox
    dims(0) = Abs(rBox.MaxPoint.X - rBox.MinPoint.X)
    dims(1) = Abs(rBox.MaxPoint.Y - rBox.MinPoint.Y)
    dims(2) = Abs(rBox.MaxPoint.Z - rBox.MinPoint.Z)
    
    ' Sort to find smallest (thickness)
    Dim temp, i, j
    For i = 0 To 1
        For j = i + 1 To 2
            If dims(j) < dims(i) Then
                temp = dims(i)
                dims(i) = dims(j)
                dims(j) = temp
            End If
        Next
    Next
    
    thickness = dims(0) ' Smallest dimension
    WScript.Echo "Detected thickness: " & FormatNumber(thickness * 10, 1) & " mm"
    
    ' Set SubType to Sheet Metal
    compDef.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    If Err.Number = 0 Then
        WScript.Echo "SubType changed to Sheet Metal"
        
        ' Set thickness parameter
        Set smDef = compDef
        smDef.Thickness.Value = thickness
        WScript.Echo "Thickness set to: " & FormatNumber(smDef.Thickness.Value * 10, 1) & " mm"
        
        partDoc.Update
    Else
        WScript.Echo "Failed to change SubType: " & Err.Description
        Err.Clear
    End If
Else
    WScript.Echo "Already sheet metal"
    Set smDef = compDef
End If

' Step 4: Create Flat Pattern
WScript.Echo ""
WScript.Echo "=== STEP 4: CREATE FLAT PATTERN ==="

If smDef Is Nothing Then
    Set smDef = compDef
End If

WScript.Echo "HasFlatPattern: " & smDef.HasFlatPattern
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
    Err.Clear
End If

If Not smDef.HasFlatPattern Then
    WScript.Echo "Creating flat pattern..."
    
    ' Find largest planar face for base
    Dim body, faces, face
    Dim largestFace, largestArea
    largestArea = 0
    
    Set body = smDef.SurfaceBodies.Item(1)
    Set faces = body.Faces
    WScript.Echo "Total faces: " & faces.Count
    
    For Each face In faces
        If face.SurfaceType = 3 Then ' kPlaneSurface
            Dim area
            area = face.Evaluator.Area * 100
            If area > largestArea Then
                largestArea = area
                Set largestFace = face
            End If
        End If
    Next
    
    WScript.Echo "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Create flat pattern
    smDef.Unfold
    
    If Err.Number = 0 Then
        WScript.Echo "Flat pattern created!"
        partDoc.Update
    Else
        WScript.Echo "Unfold failed: " & Err.Description
        Err.Clear
    End If
End If

' Step 5: Check result
WScript.Echo ""
WScript.Echo "=== FINAL RESULT ==="

If smDef.HasFlatPattern Then
    Dim fp
    Set fp = smDef.FlatPattern
    WScript.Echo "Flat pattern dimensions:"
    WScript.Echo "  Length: " & FormatNumber(fp.Length * 10, 1) & " mm"
    WScript.Echo "  Width: " & FormatNumber(fp.Width * 10, 1) & " mm"
    
    If fp.Length * 10 > 50 And fp.Width * 10 > 50 Then
        WScript.Echo ""
        WScript.Echo "*** SUCCESS! Orientation looks correct! ***"
    Else
        WScript.Echo ""
        WScript.Echo "Edge view detected - may need manual orientation fix"
    End If
Else
    WScript.Echo "No flat pattern"
End If

WScript.Echo ""
WScript.Echo "=== DONE ==="
