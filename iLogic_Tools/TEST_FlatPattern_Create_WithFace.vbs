' TEST SCRIPT - Create Flat Pattern with Correct Base Face
' Author: Quintin de Bruin © 2026

Option Explicit

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

Sub Main()
    On Error Resume Next

    WScript.Echo "=== CREATE FLAT PATTERN WITH CORRECT FACE ==="
    WScript.Echo ""

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument
    
    WScript.Echo "Part: " & partDoc.DisplayName

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    ' Step 1: Delete existing flat pattern
    WScript.Echo ""
    WScript.Echo "=== STEP 1: DELETE EXISTING FLAT PATTERN ==="
    
    If smDef.HasFlatPattern Then
        WScript.Echo "Deleting existing flat pattern..."
        smDef.FlatPattern.Delete
        If Err.Number = 0 Then
            WScript.Echo "Deleted successfully"
        Else
            WScript.Echo "Delete failed: " & Err.Description
            Err.Clear
        End If
        WScript.Sleep 500
    Else
        WScript.Echo "No flat pattern exists"
    End If

    ' Step 2: Find the largest face
    WScript.Echo ""
    WScript.Echo "=== STEP 2: FIND LARGEST FACE ==="
    
    Dim body
    Set body = smDef.SurfaceBodies.Item(1)
    
    Dim largestFace, secondLargestFace
    Dim largestArea, secondLargestArea
    largestArea = 0
    secondLargestArea = 0
    
    Dim i
    For i = 1 To body.Faces.Count
        Dim face
        Set face = body.Faces.Item(i)
        
        Dim faceArea
        faceArea = face.Evaluator.Area
        
        WScript.Echo "Face " & i & ": " & FormatNumber(faceArea * 100, 0) & " mm²"
        
        If faceArea > largestArea Then
            Set secondLargestFace = largestFace
            secondLargestArea = largestArea
            Set largestFace = face
            largestArea = faceArea
        ElseIf faceArea > secondLargestArea Then
            Set secondLargestFace = face
            secondLargestArea = faceArea
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "Largest face: " & FormatNumber(largestArea * 100, 0) & " mm²"

    ' Step 3: Try different methods to create flat pattern with the correct face
    WScript.Echo ""
    WScript.Echo "=== STEP 3: CREATE FLAT PATTERN WITH LARGEST FACE ==="
    
    ' METHOD A: Try using FlatPattern feature Add with definition
    WScript.Echo ""
    WScript.Echo "METHOD A: FlatPatternFeatures.Add with definition..."
    
    Dim fpFeatures
    Set fpFeatures = smDef.Features.FlatPatternFeatures
    
    If Err.Number = 0 And Not fpFeatures Is Nothing Then
        WScript.Echo "Got FlatPatternFeatures collection"
        
        ' Create the definition
        Dim fpDef
        Set fpDef = fpFeatures.CreateFlatPatternDefinition(largestFace)
        
        If Err.Number = 0 And Not fpDef Is Nothing Then
            WScript.Echo "Created definition with largest face"
            
            ' Add the feature
            Dim fpFeature
            Set fpFeature = fpFeatures.Add(fpDef)
            
            If Err.Number = 0 Then
                WScript.Echo "SUCCESS! Flat pattern created with largest face!"
            Else
                WScript.Echo "Add failed: " & Err.Description
                Err.Clear
            End If
        Else
            WScript.Echo "CreateFlatPatternDefinition failed: " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "Could not get FlatPatternFeatures: " & Err.Description
        Err.Clear
    End If
    
    ' Check if we have flat pattern now
    If Not smDef.HasFlatPattern Then
        ' METHOD B: Try SheetMetalFeatures.FlatPatterns
        WScript.Echo ""
        WScript.Echo "METHOD B: SheetMetalFeatures.FlatPatterns..."
        
        Dim smFeatures
        Set smFeatures = smDef.Features.SheetMetalFeatures
        
        If Err.Number = 0 And Not smFeatures Is Nothing Then
            WScript.Echo "Got SheetMetalFeatures"
            
            Dim smFlatPatterns
            Set smFlatPatterns = smFeatures.FlatPatterns
            
            If Err.Number = 0 And Not smFlatPatterns Is Nothing Then
                WScript.Echo "Got FlatPatterns from SheetMetalFeatures"
            Else
                WScript.Echo "Could not get FlatPatterns: " & Err.Description
                Err.Clear
            End If
        Else
            WScript.Echo "Could not get SheetMetalFeatures: " & Err.Description
            Err.Clear
        End If
    End If
    
    ' Check if we have flat pattern now
    If Not smDef.HasFlatPattern Then
        ' METHOD C: Use Unfold() then try to edit
        WScript.Echo ""
        WScript.Echo "METHOD C: Create with Unfold, then edit definition..."
        
        smDef.Unfold
        If Err.Number = 0 Then
            WScript.Echo "Unfold succeeded"
        Else
            WScript.Echo "Unfold failed: " & Err.Description
            Err.Clear
        End If
    End If
    
    ' Step 4: If we have flat pattern, check dimensions and try to fix
    WScript.Echo ""
    WScript.Echo "=== STEP 4: CHECK AND FIX ORIENTATION ==="
    
    If smDef.HasFlatPattern Then
        Dim fp
        Set fp = smDef.FlatPattern
        
        Dim fpLength, fpWidth
        fpLength = fp.Length * 10
        fpWidth = fp.Width * 10
        
        WScript.Echo "Flat pattern dimensions: " & FormatNumber(fpLength, 2) & " x " & FormatNumber(fpWidth, 2) & " mm"
        
        If fpLength < 50 Or fpWidth < 50 Then
            WScript.Echo ""
            WScript.Echo "Dimensions incorrect - trying to fix..."
            
            ' Try to get and modify the flat pattern feature
            WScript.Echo ""
            WScript.Echo "Looking for FlatPattern in Features..."
            
            Dim allFeatures
            Set allFeatures = smDef.Features
            
            ' Iterate through all features to find the flat pattern
            WScript.Echo "Total features: " & allFeatures.Count
            
            Dim j
            For j = 1 To allFeatures.Count
                Dim feat
                Set feat = allFeatures.Item(j)
                
                Dim featName, featType
                featName = ""
                featType = ""
                
                On Error Resume Next
                featName = feat.Name
                featType = TypeName(feat)
                Err.Clear
                
                If InStr(LCase(featName), "flat") > 0 Or InStr(LCase(featType), "flat") > 0 Then
                    WScript.Echo "Found: " & featName & " (Type: " & featType & ")"
                    
                    ' Try to access definition
                    Dim featDef
                    Set featDef = feat.Definition
                    
                    If Err.Number = 0 And Not featDef Is Nothing Then
                        WScript.Echo "  Got Definition"
                        WScript.Echo "  Definition Type: " & TypeName(featDef)
                        
                        ' List properties of the definition
                        WScript.Echo "  Checking for SetBaseFace or StaticFace..."
                        
                        ' Try SetBaseFace method
                        featDef.SetBaseFace largestFace
                        If Err.Number = 0 Then
                            WScript.Echo "  SetBaseFace succeeded!"
                            partDoc.Update
                        Else
                            WScript.Echo "  SetBaseFace failed: " & Err.Description
                            Err.Clear
                        End If
                    Else
                        WScript.Echo "  Could not get Definition: " & Err.Description
                        Err.Clear
                    End If
                End If
            Next
            
            ' Re-check dimensions
            fpLength = fp.Length * 10
            fpWidth = fp.Width * 10
            WScript.Echo ""
            WScript.Echo "Final dimensions: " & FormatNumber(fpLength, 2) & " x " & FormatNumber(fpWidth, 2) & " mm"
        Else
            WScript.Echo "Dimensions look correct!"
        End If
    Else
        WScript.Echo "No flat pattern exists"
    End If
    
    WScript.Echo ""
    WScript.Echo "=== DONE ==="
End Sub

Main
