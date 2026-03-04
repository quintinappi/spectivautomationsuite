' TEST SCRIPT - Fix Flat Pattern Orientation
' Specifically tests fixing the flat pattern base face issue
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kPartDocumentObject = 12290
Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

' Global variables
Dim m_Log

Sub Main()
    On Error Resume Next

    m_Log = ""
    LogMessage "=== FLAT PATTERN ORIENTATION FIX TEST ==="
    LogMessage ""

    ' Get Inventor application
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        LogMessage "ERROR: Inventor is not running"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        Exit Sub
    End If

    LogMessage "Connected to Inventor"

    ' Check if we have an active document
    If invApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "Please open a sheet metal PART document (.ipt file) in Inventor first.", vbCritical, "Error"
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument

    ' Verify it's a part document
    If partDoc.DocumentType <> kPartDocumentObject Then
        LogMessage "ERROR: Active document is not a part (type: " & partDoc.DocumentType & ")"
        MsgBox "Please open a PART document (.ipt file), not an assembly.", vbCritical, "Error"
        Exit Sub
    End If

    LogMessage "Part: " & partDoc.DisplayName

    ' Verify it's sheet metal
    If partDoc.SubType <> kSheetMetalSubType Then
        LogMessage "ERROR: Part is not sheet metal type"
        MsgBox "This part is not sheet metal. Please convert to sheet metal first.", vbCritical, "Error"
        Exit Sub
    End If

    LogMessage "Part is sheet metal type - good"

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    ' Check current flat pattern state
    LogMessage ""
    LogMessage "=== CURRENT STATE ==="
    If smDef.HasFlatPattern Then
        Dim fp
        Set fp = smDef.FlatPattern
        LogMessage "Flat pattern EXISTS"
        LogMessage "Current dimensions: " & FormatNumber(fp.Length * 10, 2) & "mm x " & FormatNumber(fp.Width * 10, 2) & "mm"
        
        ' Check if one dimension is suspiciously small (< 50mm = likely thickness/edge view)
        If fp.Length * 10 < 50 Or fp.Width * 10 < 50 Then
            LogMessage "WARNING: One dimension < 50mm - this is likely EDGE VIEW (showing thickness)"
        End If
    Else
        LogMessage "No flat pattern exists"
    End If

    ' STEP 1: Check if this is a derived part
    LogMessage ""
    LogMessage "=== STEP 1: ANALYZE PART TYPE ==="
    
    ' Check for derived part features
    Dim features
    Set features = smDef.Features
    
    Dim derivedPartFeatures
    On Error Resume Next
    Set derivedPartFeatures = features.DerivePartFeatures
    
    Dim isDerivedPart
    isDerivedPart = False
    
    If Err.Number = 0 And Not derivedPartFeatures Is Nothing Then
        If derivedPartFeatures.Count > 0 Then
            LogMessage "This is a DERIVED PART (has " & derivedPartFeatures.Count & " derived features)"
            isDerivedPart = True
        Else
            LogMessage "No DerivePartFeatures found - checking for ShrinkwrapSubstitute..."
        End If
    Else
        Err.Clear
    End If
    
    ' Also check for ReferenceFeatures which is another sign of derived parts
    Dim refFeatures
    Set refFeatures = features.ReferenceFeatures
    
    If Err.Number = 0 And Not refFeatures Is Nothing Then
        If refFeatures.Count > 0 Then
            LogMessage "Part has " & refFeatures.Count & " reference features (may be derived)"
            isDerivedPart = True
        End If
    End If
    Err.Clear
    
    ' For derived parts, try a workaround: use the "Flat Pattern Orientation" feature
    If isDerivedPart Then
        LogMessage ""
        LogMessage "*** DERIVED PART DETECTED ***"
        LogMessage "Derived parts have limited API control for flat pattern base face."
        LogMessage "Workaround: Try using FlatPatternOrientation feature or manual intervention."
    End If
    
    ' STEP 2: If flat pattern exists, delete it so we can start fresh
    LogMessage ""
    LogMessage "=== STEP 2: DELETE FLAT PATTERN ==="
    
    If smDef.HasFlatPattern Then
        LogMessage "Attempting to delete current flat pattern..."
        
        ' Method 1: Try Refold
        On Error Resume Next
        smDef.Refold
        If Err.Number <> 0 Then
            LogMessage "Refold method failed: " & Err.Description
            Err.Clear
            
            ' Method 2: Try FlatPattern.Delete
            LogMessage "Trying FlatPattern.Delete..."
            Dim fp2
            Set fp2 = smDef.FlatPattern
            fp2.Delete
            If Err.Number <> 0 Then
                LogMessage "FlatPattern.Delete failed: " & Err.Description
                Err.Clear
                
                ' Method 3: Try deleting via Features
                LogMessage "Trying to suppress flat pattern feature..."
                Dim fpFeature
                Set fpFeature = smDef.FlatPattern
                fpFeature.Suppressed = True
                If Err.Number <> 0 Then
                    LogMessage "Suppress also failed: " & Err.Description
                    Err.Clear
                Else
                    LogMessage "Flat pattern suppressed"
                End If
            Else
                LogMessage "FlatPattern.Delete succeeded"
            End If
        Else
            LogMessage "Refold succeeded"
        End If
        
        WScript.Sleep 500
        partDoc.Update
    Else
        LogMessage "No flat pattern to delete"
    End If

    ' Verify refold worked
    If smDef.HasFlatPattern Then
        LogMessage "WARNING: Flat pattern still exists after refold"
    Else
        LogMessage "Flat pattern removed - model is now folded"
    End If

    ' STEP 2: Find all faces and identify the largest face (for base)
    LogMessage ""
    LogMessage "=== STEP 2: ANALYZE FACES ==="
    
    ' Get all surface bodies
    Dim surfaceBodies
    Set surfaceBodies = smDef.SurfaceBodies
    
    LogMessage "Number of surface bodies: " & surfaceBodies.Count
    
    If surfaceBodies.Count = 0 Then
        LogMessage "ERROR: No surface bodies found"
        MsgBox "No surface bodies found in the part.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Analyze all faces across all bodies
    ' Find the LARGEST face by area - for a plate, this will be the top or bottom face
    Dim largestFace
    Set largestFace = Nothing
    Dim largestArea
    largestArea = 0
    Dim secondLargestFace
    Set secondLargestFace = Nothing
    Dim secondLargestArea
    secondLargestArea = 0
    Dim totalFaces
    totalFaces = 0
    
    Dim bodyIndex
    For bodyIndex = 1 To surfaceBodies.Count
        Dim body
        Set body = surfaceBodies.Item(bodyIndex)
        
        LogMessage "Body " & bodyIndex & " has " & body.Faces.Count & " faces"
        
        Dim faceIndex
        For faceIndex = 1 To body.Faces.Count
            totalFaces = totalFaces + 1
            
            Dim face
            Set face = body.Faces.Item(faceIndex)
            
            ' Get face area using different methods
            Dim faceArea
            faceArea = 0
            
            On Error Resume Next
            ' Try Evaluator.Area first
            faceArea = face.Evaluator.Area
            If Err.Number <> 0 Then
                Err.Clear
                ' Try direct Area property
                faceArea = face.Area
                If Err.Number <> 0 Then
                    Err.Clear
                    faceArea = 0
                End If
            End If
            
            ' Convert from cm² to mm²
            Dim areaInMM2
            areaInMM2 = faceArea * 100
            
            ' Get surface type for logging
            Dim surfType
            surfType = "Unknown"
            On Error Resume Next
            
            ' Try to determine if it's planar by checking geometry
            Dim geom
            Set geom = face.Geometry
            If Err.Number = 0 Then
                If TypeName(geom) = "Plane" Then
                    surfType = "PLANAR"
                Else
                    surfType = TypeName(geom)
                End If
            End If
            Err.Clear
            
            LogMessage "  Face " & faceIndex & ": " & surfType & ", Area = " & FormatNumber(areaInMM2, 0) & " mm²"
            
            ' Track largest and second largest faces
            If faceArea > largestArea Then
                ' Current largest becomes second largest
                Set secondLargestFace = largestFace
                secondLargestArea = largestArea
                ' New largest
                largestArea = faceArea
                Set largestFace = face
            ElseIf faceArea > secondLargestArea Then
                secondLargestArea = faceArea
                Set secondLargestFace = face
            End If
        Next
    Next
    
    LogMessage ""
    LogMessage "Total faces: " & totalFaces
    
    If largestFace Is Nothing Then
        LogMessage "ERROR: No faces found"
        MsgBox "Could not find any faces on the part.", vbCritical, "Error"
        Exit Sub
    End If
    
    LogMessage "Largest face area: " & FormatNumber(largestArea * 100, 0) & " mm²"
    If Not secondLargestFace Is Nothing Then
        LogMessage "Second largest face area: " & FormatNumber(secondLargestArea * 100, 0) & " mm²"
    End If

    ' STEP 3: Create flat pattern using the largest face
    LogMessage ""
    LogMessage "=== STEP 3: CREATE FLAT PATTERN WITH LARGEST FACE ==="
    
    On Error Resume Next
    
    ' Method 1: Try creating FlatPatternDefinition with specific face
    LogMessage "Method 1: Trying FlatPatternDefinition.Create with face..."
    
    Dim flatPatternFeatures
    Set flatPatternFeatures = smDef.Features.FlatPatterns
    
    If Err.Number = 0 And Not flatPatternFeatures Is Nothing Then
        LogMessage "Got FlatPatterns collection"
        
        ' Try to create definition with the largest face
        Dim fpDef
        Set fpDef = flatPatternFeatures.CreateFlatPatternDefinition(largestFace)
        
        If Err.Number = 0 And Not fpDef Is Nothing Then
            LogMessage "Created FlatPatternDefinition with face"
            
            ' Create the flat pattern
            Dim newFP
            Set newFP = flatPatternFeatures.Add(fpDef)
            
            If Err.Number = 0 Then
                LogMessage "Flat pattern created with Add!"
            Else
                LogMessage "Add failed: " & Err.Description
                Err.Clear
            End If
        Else
            LogMessage "CreateFlatPatternDefinition failed: " & Err.Description
            Err.Clear
        End If
    Else
        LogMessage "Could not get FlatPatterns collection: " & Err.Description
        Err.Clear
    End If
    
    ' Check if we have a flat pattern now
    If Not smDef.HasFlatPattern Then
        ' Method 2: Try Unfold with ObjectCollection
        LogMessage "Method 2: Trying Unfold with ObjectCollection..."
        
        Dim objCol
        Set objCol = invApp.TransientObjects.CreateObjectCollection
        objCol.Add largestFace
        
        smDef.Unfold objCol
        
        If Err.Number <> 0 Then
            LogMessage "Unfold with ObjectCollection failed: " & Err.Description
            Err.Clear
        Else
            LogMessage "Unfold with ObjectCollection succeeded"
        End If
    End If
    
    ' Check if we have a flat pattern now
    If Not smDef.HasFlatPattern Then
        ' Method 3: Simple Unfold (no parameters)
        LogMessage "Method 3: Trying simple Unfold..."
        smDef.Unfold
        
        If Err.Number <> 0 Then
            LogMessage "Simple Unfold failed: " & Err.Description
            Err.Clear
        Else
            LogMessage "Simple Unfold succeeded"
        End If
    End If
    
    WScript.Sleep 500
    partDoc.Update

    ' STEP 4: Check results and try to fix orientation
    LogMessage ""
    LogMessage "=== STEP 4: CHECK RESULTS ==="
    
    If smDef.HasFlatPattern Then
        Set fp = smDef.FlatPattern
        Dim newLength, newWidth
        newLength = fp.Length * 10
        newWidth = fp.Width * 10
        
        LogMessage "Flat pattern created successfully"
        LogMessage "New dimensions: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
        
        ' Check if dimensions look correct now
        If newLength >= 50 And newWidth >= 50 Then
            LogMessage "SUCCESS: Both dimensions are reasonable - orientation is correct!"
        Else
            LogMessage "WARNING: One dimension is still small - trying to fix..."
            
            ' Try to change the base face
            LogMessage ""
            LogMessage "=== ATTEMPTING TO FIX BASE FACE ==="
            
            ' Enter Edit mode
            LogMessage "Entering flat pattern Edit mode..."
            fp.Edit
            
            If Err.Number <> 0 Then
                LogMessage "Could not enter Edit mode: " & Err.Description
                Err.Clear
            Else
                LogMessage "In Edit mode"
                
                ' Try FlipBaseFace
                LogMessage "Trying FlipBaseFace..."
                fp.FlipBaseFace
                
                If Err.Number <> 0 Then
                    LogMessage "FlipBaseFace failed: " & Err.Description
                    Err.Clear
                Else
                    LogMessage "FlipBaseFace executed (no error)"
                End If
                
                ' Exit Edit mode
                fp.ExitEdit
                partDoc.Update
                
                ' Check dimensions again
                newLength = fp.Length * 10
                newWidth = fp.Width * 10
                LogMessage "After FlipBaseFace: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
            End If
            
            ' If still wrong, try accessing the FlatPattern FEATURE (not object)
            If newLength < 50 Or newWidth < 50 Then
                LogMessage ""
                LogMessage "Trying to access FlatPattern as a FEATURE..."
                
                On Error Resume Next
                
                ' The FlatPattern is also a feature in the Features collection
                Dim flatPatternFeature
                Set flatPatternFeature = smDef.Features.FlatPattern
                
                If Err.Number = 0 And Not flatPatternFeature Is Nothing Then
                    LogMessage "Got FlatPattern feature"
                    LogMessage "Feature name: " & flatPatternFeature.Name
                    
                    ' Try to get the definition
                    Dim fpDefinition
                    Set fpDefinition = flatPatternFeature.Definition
                    
                    If Err.Number = 0 And Not fpDefinition Is Nothing Then
                        LogMessage "Got FlatPattern definition"
                        
                        ' Try to access/set StaticFace property
                        Dim staticFace
                        Set staticFace = fpDefinition.StaticFace
                        
                        If Err.Number = 0 Then
                            LogMessage "Current StaticFace retrieved"
                            Dim sfArea
                            sfArea = staticFace.Evaluator.Area * 100
                            LogMessage "StaticFace area: " & FormatNumber(sfArea, 0) & " mm²"
                            
                            If sfArea < 100000 Then
                                LogMessage "StaticFace is edge - trying to change it..."
                                Set fpDefinition.StaticFace = largestFace
                                
                                If Err.Number = 0 Then
                                    LogMessage "StaticFace changed!"
                                    partDoc.Update
                                    
                                    newLength = fp.Length * 10
                                    newWidth = fp.Width * 10
                                    LogMessage "New dimensions: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
                                Else
                                    LogMessage "Could not set StaticFace: " & Err.Description
                                    Err.Clear
                                End If
                            End If
                        Else
                            LogMessage "Could not get StaticFace: " & Err.Description
                            Err.Clear
                        End If
                    Else
                        LogMessage "Could not get definition: " & Err.Description
                        Err.Clear
                    End If
                Else
                    LogMessage "Could not get FlatPattern feature: " & Err.Description
                    Err.Clear
                End If
            End If
            
            ' Try BaseFace property
            If newLength < 50 Or newWidth < 50 Then
                LogMessage ""
                LogMessage "Trying BaseFace property..."
                
                On Error Resume Next
                
                ' Get current base face
                Dim currentBaseFace
                Set currentBaseFace = fp.BaseFace
                
                If Err.Number = 0 And Not currentBaseFace Is Nothing Then
                    LogMessage "Current base face retrieved"
                    
                    Dim currentBaseArea
                    currentBaseArea = currentBaseFace.Evaluator.Area * 100
                    LogMessage "Current base face area: " & FormatNumber(currentBaseArea, 0) & " mm²"
                    
                    ' The base face should be the large face (2.6M mm²)
                    ' If it's small (edge face), we need to change it
                    If currentBaseArea < 100000 Then
                        LogMessage "Current base face is an EDGE - need to change to a large face"
                        
                        ' Try to set BaseFace property
                        LogMessage "Attempting to set fp.BaseFace = largestFace..."
                        Set fp.BaseFace = largestFace
                        
                        If Err.Number <> 0 Then
                            LogMessage "Setting BaseFace failed: " & Err.Description
                            Err.Clear
                        Else
                            LogMessage "BaseFace set successfully!"
                            partDoc.Update
                            
                            newLength = fp.Length * 10
                            newWidth = fp.Width * 10
                            LogMessage "New dimensions: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
                        End If
                    Else
                        LogMessage "Base face appears to be a large face - strange"
                    End If
                Else
                    LogMessage "Could not get BaseFace: " & Err.Description
                    Err.Clear
                End If
            End If
            
            ' Last resort: Try using command to change view
            If newLength < 50 Or newWidth < 50 Then
                LogMessage ""
                LogMessage "Trying PartChangeFlatPatternBaseFaceCmd command..."
                LogMessage "Will try to pre-select the largest face first..."
                
                ' Try to select the largest face before executing command
                Dim selectSet
                Set selectSet = partDoc.SelectSet
                
                ' Clear any existing selection
                selectSet.Clear
                
                ' Add largest face to selection
                selectSet.Select largestFace
                
                If Err.Number = 0 Then
                    LogMessage "Largest face pre-selected"
                Else
                    LogMessage "Could not pre-select face: " & Err.Description
                    Err.Clear
                End If
                
                Dim cmdMgr
                Set cmdMgr = invApp.CommandManager
                
                Dim changeBaseFaceCmd
                Set changeBaseFaceCmd = cmdMgr.ControlDefinitions.Item("PartChangeFlatPatternBaseFaceCmd")
                
                If Err.Number = 0 And Not changeBaseFaceCmd Is Nothing Then
                    LogMessage "Found PartChangeFlatPatternBaseFaceCmd - executing..."
                    changeBaseFaceCmd.Execute
                    
                    ' Wait for command to complete
                    WScript.Sleep 1000
                    partDoc.Update
                    
                    ' Check dimensions after command
                    newLength = fp.Length * 10
                    newWidth = fp.Width * 10
                    LogMessage "After command: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
                    
                    If newLength >= 50 And newWidth >= 50 Then
                        LogMessage "SUCCESS: Base face changed!"
                    Else
                        LogMessage "Command executed but dimensions still wrong"
                        LogMessage "USER ACTION REQUIRED: Please select a large face in the Inventor dialog"
                    End If
                Else
                    LogMessage "Command not found: " & Err.Description
                    Err.Clear
                    
                    ' Try other command names
                    Dim cmdNames
                    cmdNames = Array("SheetMetalChangeBaseFaceCmd", "ChangeBaseFaceCmd", "FlatPatternBaseFaceCmd")
                    
                    Dim cmdName
                    For Each cmdName In cmdNames
                        Set changeBaseFaceCmd = cmdMgr.ControlDefinitions.Item(cmdName)
                        If Err.Number = 0 And Not changeBaseFaceCmd Is Nothing Then
                            LogMessage "Found " & cmdName & " - executing..."
                            changeBaseFaceCmd.Execute
                            Exit For
                        End If
                        Err.Clear
                    Next
                End If
            End If
            
            ' FINAL METHOD: Try changing Punch Tool Direction which can affect orientation
            If newLength < 50 Or newWidth < 50 Then
                LogMessage ""
                LogMessage "=== TRYING PUNCH TOOL DIRECTION METHOD ==="
                
                On Error Resume Next
                
                ' Access the sheet metal component definition
                Dim smCompDef
                Set smCompDef = partDoc.ComponentDefinition
                
                ' Check if we can access PunchToolDirection
                LogMessage "Checking PunchToolDirection property..."
                
                Dim punchDir
                punchDir = smCompDef.PunchToolDirection
                
                If Err.Number = 0 Then
                    LogMessage "Current PunchToolDirection: " & punchDir
                    
                    ' Try to flip it (1 = top, 2 = bottom typically)
                    If punchDir = 1 Then
                        smCompDef.PunchToolDirection = 2
                    Else
                        smCompDef.PunchToolDirection = 1
                    End If
                    
                    If Err.Number = 0 Then
                        LogMessage "PunchToolDirection changed!"
                        partDoc.Update
                        
                        ' Check dimensions
                        newLength = fp.Length * 10
                        newWidth = fp.Width * 10
                        LogMessage "After direction change: " & FormatNumber(newLength, 2) & "mm x " & FormatNumber(newWidth, 2) & "mm"
                    Else
                        LogMessage "Could not change PunchToolDirection: " & Err.Description
                        Err.Clear
                    End If
                Else
                    LogMessage "Could not get PunchToolDirection: " & Err.Description
                    Err.Clear
                End If
                
                ' Try accessing FlatPattern.PunchRepresentation
                LogMessage ""
                LogMessage "Checking FlatPattern.PunchRepresentation..."
                
                Dim punchRep
                punchRep = fp.PunchRepresentation
                
                If Err.Number = 0 Then
                    LogMessage "Current PunchRepresentation: " & punchRep
                Else
                    LogMessage "Could not get PunchRepresentation: " & Err.Description
                    Err.Clear
                End If
            End If
        End If
    Else
        LogMessage "ERROR: Flat pattern was not created"
    End If

    LogMessage ""
    LogMessage "=== TEST COMPLETE ==="
    
    ' Show results
    Dim resultMsg
    If smDef.HasFlatPattern Then
        Set fp = smDef.FlatPattern
        resultMsg = "Flat pattern dimensions:" & vbCrLf & _
                    FormatNumber(fp.Length * 10, 2) & "mm x " & FormatNumber(fp.Width * 10, 2) & "mm" & vbCrLf & vbCrLf
        
        If fp.Length * 10 >= 50 And fp.Width * 10 >= 50 Then
            resultMsg = resultMsg & "Orientation appears CORRECT!"
        Else
            resultMsg = resultMsg & "Orientation may still be wrong - check manually"
        End If
    Else
        resultMsg = "No flat pattern was created"
    End If
    
    MsgBox resultMsg, vbInformation, "Flat Pattern Fix Test"
    
    ' Print log
    WScript.Echo m_Log
End Sub

Sub LogMessage(msg)
    m_Log = m_Log & msg & vbCrLf
    WScript.Echo msg
End Sub

' Run main
Main
