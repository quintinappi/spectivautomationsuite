' Sheet_Metal_Converter.vbs - DETAILING WORKFLOW STEP 3: Convert PL parts to sheet metal
' DETAILING WORKFLOW - STEP 3: Convert PL parts to sheet metal in assemblies
' Sheet Metal Converter - Standalone VBScript
' Converts plate parts to sheet metal with flat patterns
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kAssemblyDocumentObject = 12291
Const kMillimeterLengthUnits = 11269
Const kNumberParameterType = 1

' Global variables
Dim m_InventorApp
Dim m_Log
Dim m_LogPath
Dim m_SkippedParts ' Collection of parts that need manual conversion
Dim m_Apprentice

Sub Main()
    On Error Resume Next

    ' Initialize logging
    m_Log = ""
    
    ' Initialize skipped parts collection
    Set m_SkippedParts = CreateObject("Scripting.Dictionary")
    
    LogMessage "=== SHEET METAL CONVERTER STARTED ==="

' Get Inventor application (for UI and active document)
    LogMessage "Attempting to get Inventor application..."
    On Error Resume Next

    ' First try to get existing instance
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "No existing Inventor instance found, trying to create new instance..."
        Err.Clear

        ' Try to create new instance
        Set m_InventorApp = CreateObject("Inventor.Application")
        If Err.Number <> 0 Then
            LogMessage "ERROR: Failed to connect to Inventor - " & Err.Description
            MsgBox "Failed to connect to Inventor. Please make sure Inventor is installed and try starting it manually." & vbCrLf & vbCrLf & _
                   "Error: " & Err.Description, vbCritical, "Connection Failed"
            SaveLog
            Exit Sub
        Else
            LogMessage "Created new Inventor instance"
            ' Make Inventor visible
            m_InventorApp.Visible = True
        End If
    Else
        LogMessage "Connected to existing Inventor instance"
    End If

    ' Verify we have a valid application object
    If m_InventorApp Is Nothing Then
        LogMessage "ERROR: Inventor application object is null"
        MsgBox "Failed to get valid Inventor application object.", vbCritical, "Connection Failed"
        SaveLog
        Exit Sub
    End If
    On Error GoTo 0

    If m_InventorApp Is Nothing Then
        LogMessage "ERROR: Inventor application object is Nothing"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        Exit Sub
    End If

    LogMessage "Inventor application found successfully"

    ' Get Inventor Apprentice Server (for read-only part analysis)
    LogMessage "Attempting to connect to Inventor Apprentice Server..."
    On Error Resume Next

    ' Try different ProgIDs for Apprentice Server
    Dim apprenticeProgIDs
    apprenticeProgIDs = Array("Inventor.ApprenticeServerComponent", "Inventor.ApprenticeServer", "ApprenticeServer.Component")

    Dim progID
    For Each progID In apprenticeProgIDs
        LogMessage "Trying Apprentice ProgID: " & progID
        Set m_Apprentice = CreateObject(progID)
        If Err.Number = 0 And Not m_Apprentice Is Nothing Then
            LogMessage "Successfully connected to Apprentice Server with: " & progID
            Exit For
        Else
            Err.Clear
        End If
    Next

    If m_Apprentice Is Nothing Then
        LogMessage "WARNING: Failed to connect to Apprentice Server - will use basic detection"
        LogMessage "Parts will be checked for sheet metal capabilities when opened"
    Else
        LogMessage "Connected to Inventor Apprentice Server successfully"
    End If
    On Error GoTo 0

    ' Check if we have an active document (for assembly path)
    If m_InventorApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document found"
        MsgBox "No active document! Please open an assembly in Inventor.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Active document found: " & m_InventorApp.ActiveDocument.FullFileName
    LogMessage "Document type: " & m_InventorApp.ActiveDocument.DocumentType

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogMessage "ERROR: Document is not an assembly - found: " & m_InventorApp.ActiveDocument.DocumentType
        MsgBox "Please open an ASSEMBLY document (.iam file), not a part." & vbCrLf & vbCrLf & _
               "The sheet metal converter needs to scan an assembly to find plate parts." & vbCrLf & vbCrLf & _
               "Current document: " & m_InventorApp.ActiveDocument.DisplayName, vbExclamation, "Assembly Required"
        SaveLog
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogMessage "Processing assembly: " & asmDoc.FullFileName

    ' Step 1: Scan assembly for plates (using same method as renamer)
    LogMessage "STEP 1: Scanning assembly for plates containing 'PL' or 'S355JR'"
    Dim plateParts
    Set plateParts = ScanAssemblyForPlates(asmDoc)

    If plateParts.Count = 0 Then
        LogMessage "No plate parts found in BOM"
        MsgBox "No parts containing 'PL' or 'S355JR' found in the BOM." & vbCrLf & vbCrLf & _
               "Make sure your parts have 'PL' or 'S355JR' in their Part Number or Description field," & vbCrLf & _
               "and that the thickness is specified (e.g., '10mm', '5 mm', etc.).", vbInformation, "No Plates Found"
        SaveLog
        Exit Sub
    End If

    LogMessage "Found " & plateParts.Count & " thickness groups:"
    Dim thickness
    For Each thickness In plateParts.Keys
        LogMessage "  Thickness " & thickness & "mm: " & plateParts(thickness).Count & " parts"
    Next
    
    ' Count total parts to process
    Dim totalParts
    totalParts = 0
    For Each thickness In plateParts.Keys
        totalParts = totalParts + plateParts(thickness).Count
    Next
    
    ' Show upfront warning about interactive process
    LogMessage "Total parts to process: " & totalParts
    
    Dim userResponse
    userResponse = MsgBox("SHEET METAL BATCH CONVERTER" & vbCrLf & vbCrLf & _
                          "Found " & totalParts & " plate parts to convert." & vbCrLf & vbCrLf & _
                          "IMPORTANT: This is a SEMI-AUTOMATED process." & vbCrLf & vbCrLf & _
                          "For EACH part, you will need to:" & vbCrLf & _
                          "  1. Wait for the part to open" & vbCrLf & _
                          "  2. Wait for the Convert dialog to appear" & vbCrLf & _
                          "  3. CLICK ONCE on the GREEN highlighted face" & vbCrLf & _
                          "  4. Click OK in the prompt to continue" & vbCrLf & vbCrLf & _
                          "This ensures correct flat pattern orientation." & vbCrLf & vbCrLf & _
                          "Ready to process " & totalParts & " parts?", vbOKCancel + vbInformation, "Batch Conversion")
    
    If userResponse = vbCancel Then
        LogMessage "User cancelled batch conversion"
        MsgBox "Batch conversion cancelled by user.", vbInformation, "Cancelled"
        SaveLog
        Exit Sub
    End If
    
    LogMessage "User confirmed - starting batch conversion"
    
    ' CRITICAL: Close the assembly so parts can open independently
    LogMessage ""
    LogMessage "Closing assembly to allow parts to open independently..."
    Dim asmPath
    asmPath = asmDoc.FullFileName
    
    asmDoc.Close True ' Close and save assembly
    LogMessage "Assembly closed: " & asmPath
    WScript.Sleep 1000

    ' Step 2: Process each plate group
    LogMessage "STEP 2: Processing plate groups"
    Dim processedCount
    processedCount = 0
    Dim failedCount
    failedCount = 0

    For Each thickness In plateParts.Keys
        Dim bomRows
        Set bomRows = plateParts(thickness)

        LogMessage "Processing thickness group: " & thickness & "mm (" & bomRows.Count & " parts)"

        Dim i
        For i = 0 To bomRows.Count - 1
            On Error Resume Next
            Dim success
            success = ProcessPlatePart(bomRows(i), thickness)
            If success Then
                processedCount = processedCount + 1
            Else
                failedCount = failedCount + 1
                LogMessage "FAILED to process part in thickness group " & thickness & "mm"
            End If
        Next
    Next

    ' Check if all parts were successfully processed
    If failedCount > 0 Then
        LogMessage "ERROR: " & failedCount & " parts failed to convert. Aborting parameter creation."
        MsgBox "Sheet metal conversion FAILED!" & vbCrLf & vbCrLf & _
               "Successfully converted: " & processedCount & " parts" & vbCrLf & _
               "Failed conversions: " & failedCount & " parts" & vbCrLf & vbCrLf & _
               "Assembly parameters will NOT be created until all parts convert successfully." & vbCrLf & vbCrLf & _
               "Check the log file for details: " & m_LogPath, vbCritical, "Conversion Failed"
        SaveLog
        Exit Sub
    End If

    LogMessage "All " & processedCount & " parts converted successfully"
    
    ' Reopen the assembly
    LogMessage ""
    LogMessage "Reopening assembly: " & asmPath
    Set asmDoc = m_InventorApp.Documents.Open(asmPath, True)
    
    If asmDoc Is Nothing Then
        LogMessage "ERROR: Could not reopen assembly"
        MsgBox "Parts converted successfully, but could not reopen assembly to add parameters." & vbCrLf & vbCrLf & _
               "Assembly: " & asmPath, vbExclamation, "Assembly Error"
        SaveLog
        Exit Sub
    End If
    
    LogMessage "Assembly reopened successfully"

    ' Step 3: Add parameters to assembly (only if all conversions succeeded)
    LogMessage "STEP 3: Adding PLATE LENGTH and PLATE WIDTH parameters to assembly"
    AddPlateParametersToAssembly asmDoc, plateParts, plateParts

    ' Save the assembly
    LogMessage "Saving assembly..."
    asmDoc.Save
    LogMessage "Assembly saved"

    SaveLog

    MsgBox "Sheet metal conversion completed!" & vbCrLf & vbCrLf & _
           "Processed " & processedCount & " plate parts." & vbCrLf & _
           "Added PLATE LENGTH and PLATE WIDTH parameters to assembly." & vbCrLf & vbCrLf & _
           "Check the Parameters dialog in Inventor to verify the parameter values." & vbCrLf & vbCrLf & _
           "Log saved to: " & m_LogPath, vbInformation, "Conversion Complete"

End Sub

Function ScanAssemblyForPlates(asmDoc)
    Dim plateParts
    Set plateParts = CreateObject("Scripting.Dictionary")

    LogMessage "Starting recursive assembly scan for plates..."

    ' Create a dictionary to track unique parts (prevent duplicates)
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    ' Start recursive traversal from root assembly
    Call ProcessAssemblyForPlates(asmDoc, uniqueParts, plateParts, "ROOT")

    LogMessage "Assembly scanning completed. Found " & plateParts.Count & " thickness groups"
    Set ScanAssemblyForPlates = plateParts
End Function

Sub ProcessAssemblyForPlates(asmDoc, uniqueParts, plateParts, asmLevel)
    LogMessage "Processing assembly - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")"

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "Found " & occurrences.Count & " occurrences in " & asmDoc.DisplayName

    ' Process each occurrence in this assembly
    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' SKIP suppressed occurrences
        If occ.Suppressed Then
            LogMessage "SKIPPING suppressed occurrence in " & asmDoc.DisplayName
        Else
            Dim doc
            Set doc = occ.Definition.Document

            Dim fileName
            fileName = GetFileNameFromPath(doc.FullFileName)

            Dim fullPath
            fullPath = doc.FullFileName

            ' Check if this is a part file
            If LCase(Right(fileName, 4)) = ".ipt" Then
                ' Process part file - check for uniqueness
                If Not uniqueParts.Exists(fullPath) Then
                    ' Mark as processed to prevent duplicates
                    uniqueParts.Add fullPath, True

                    ' Read Description from Design Tracking Properties
                    Dim description
                    description = GetDescriptionFromIProperty(doc)

                    If description = "" Then
                        LogMessage "WARNING - No description found for " & fileName
                    Else
                        ' Check if part contains "PL" or "S355JR"
                        If (InStr(1, UCase(description), "PL", vbTextCompare) > 0 Or _
                            InStr(1, UCase(description), "S355JR", vbTextCompare) > 0) Then

                            ' Extract thickness from description
                            Dim thickness
                            thickness = ExtractThickness(description)

                            If thickness = "" Then
                                LogMessage "WARNING: Could not extract thickness from: " & fileName & " - " & description
                            Else
                                ' Group by thickness
                                If Not plateParts.Exists(thickness) Then
                                    plateParts.Add thickness, CreateObject("System.Collections.ArrayList")
                                End If

                                ' Create a simple object to hold part info
                                Dim partInfo
                                Set partInfo = CreateObject("Scripting.Dictionary")
                                partInfo.Add "fullPath", fullPath
                                partInfo.Add "fileName", fileName
                                partInfo.Add "description", description
                                partInfo.Add "document", doc

                                plateParts(thickness).Add partInfo

                                LogMessage "Found plate: " & fileName & " - " & description & " (Thickness: " & thickness & "mm)"
                            End If
                        End If
                    End If
                Else
                    LogMessage "DUPLICATE PART SKIPPED - " & fileName & " (already processed)"
                End If

            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                ' This is a sub-assembly - recurse into it
                LogMessage "RECURSING into sub-assembly - " & fileName
                Call ProcessAssemblyForPlates(doc, uniqueParts, plateParts, asmLevel & ">" & fileName)
            End If
        End If
    Next
End Sub

Function GetFileNameFromPath(fullPath)
    GetFileNameFromPath = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Function GetDescriptionFromIProperty(doc)
    ' Read Description from Design Tracking Properties (same as renamer)
    On Error Resume Next

    Dim propertySet
    Set propertySet = doc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    Dim descriptionProp
    Set descriptionProp = propertySet.Item("Description")

    If Err.Number <> 0 Then
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    GetDescriptionFromIProperty = Trim(descriptionProp.Value)
    Err.Clear
End Function

Function ExtractThickness(text)
    ' Look for thickness patterns in the text
    Dim patterns
    patterns = Array("(\d+(?:\.\d+)?)\s*mm", "(\d+(?:\.\d+)?)\s*MM", "THK\s*(\d+(?:\.\d+)?)", "THICKNESS\s*(\d+(?:\.\d+)?)")

    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False

    Dim i
    For i = 0 To UBound(patterns)
        regex.Pattern = patterns(i)
        Dim matches
        Set matches = regex.Execute(text)

        If matches.Count > 0 Then
            ExtractThickness = matches(0).SubMatches(0)
            Exit Function
        End If
    Next

    ExtractThickness = ""
End Function

' ============================================================================
' HELPER FUNCTIONS FOR PART PROCESSING
' ============================================================================

Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

Function CheckIfPartIsSheetMetal(partDoc)
    CheckIfPartIsSheetMetal = False

    On Error Resume Next
    LogMessage "Checking SubType GUID: " & partDoc.SubType

    If partDoc.SubType = kSheetMetalSubType Then
        CheckIfPartIsSheetMetal = True
        LogMessage "Part is SHEET METAL type"
    Else
        LogMessage "Part is STANDARD type (not sheet metal)"
    End If
End Function

Function GetPartThicknessFromGeometry(partDoc)
    GetPartThicknessFromGeometry = 0

    On Error Resume Next

    LogMessage "Analyzing part bounding box to detect thickness..."

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get ComponentDefinition: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Get the bounding box
    Dim rangeBox
    Set rangeBox = compDef.RangeBox

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get RangeBox: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Calculate dimensions in all three axes (in cm)
    Dim dimX, dimY, dimZ
    dimX = rangeBox.MaxPoint.X - rangeBox.MinPoint.X
    dimY = rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y
    dimZ = rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z

    LogMessage "Bounding box dimensions: X=" & FormatNumber(dimX * 10, 2) & "mm, Y=" & FormatNumber(dimY * 10, 2) & "mm, Z=" & FormatNumber(dimZ * 10, 2) & "mm"

    ' Find the smallest dimension (this is the thickness)
    Dim thickness
    thickness = dimX
    If dimY < thickness Then thickness = dimY
    If dimZ < thickness Then thickness = dimZ

    LogMessage "Smallest dimension (thickness): " & FormatNumber(thickness * 10, 2) & "mm"

    GetPartThicknessFromGeometry = thickness ' Return in cm
End Function

Sub SetSheetMetalThickness(partDoc, thicknessInCm)
    On Error Resume Next

    ' Get sheet metal component definition
    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get ComponentDefinition: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    ' CRITICAL: Disable "use style thickness" first
    LogMessage "Disabling UseSheetMetalStyleThickness..."
    smDef.UseSheetMetalStyleThickness = False

    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not set UseSheetMetalStyleThickness: " & Err.Description
        Err.Clear
        ' Continue anyway
    End If

    ' Now set thickness parameter directly on component definition
    Dim thicknessParam
    Set thicknessParam = smDef.Thickness

    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not get Thickness parameter: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    If thicknessParam Is Nothing Then
        LogMessage "WARNING: Thickness parameter is Nothing"
        Exit Sub
    End If

    LogMessage "Setting thickness: " & FormatNumber(thicknessInCm * 10, 2) & "mm (" & thicknessInCm & " cm internal)"
    thicknessParam.Value = thicknessInCm

    If Err.Number <> 0 Then
        LogMessage "WARNING: Failed to set thickness: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    ' Update document to apply change
    partDoc.Update

    LogMessage "Thickness parameter value: " & FormatNumber(thicknessParam.Value * 10, 2) & " mm"
End Sub

Function CreateFlatPatternWithDimensions(partDoc, ByRef outLength, ByRef outWidth)
    CreateFlatPatternWithDimensions = False
    outLength = 0
    outWidth = 0

    On Error Resume Next

    ' Check if flat pattern already exists
    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If smDef.HasFlatPattern Then
        LogMessage "Flat pattern already exists"
    Else
        LogMessage "Creating new flat pattern..."

        ' Try using Unfold method
        LogMessage "Attempting ComponentDefinition.Unfold()..."
        smDef.Unfold

        If Err.Number <> 0 Then
            LogMessage "ERROR: Unfold failed: " & Err.Description
            Err.Clear

            ' Try using CommandManager
            LogMessage "Attempting to execute 'Create Flat Pattern' command..."

            Dim cmdMgr
            Set cmdMgr = m_InventorApp.CommandManager

            Dim flatPatternCmd
            Set flatPatternCmd = cmdMgr.ControlDefinitions.Item("PartFlatPatternCmd")

            If Err.Number <> 0 Then
                LogMessage "ERROR: Could not find PartFlatPatternCmd: " & Err.Description
                Err.Clear

                ' Try alternative command IDs
                Dim cmdIds
                cmdIds = Array("SheetMetalFlatPatternCmd", "FlatPatternCmd", "SMFlatPatternCmd")

                Dim cmdId
                For Each cmdId In cmdIds
                    LogMessage "Trying: " & cmdId
                    Set flatPatternCmd = cmdMgr.ControlDefinitions.Item(cmdId)
                    If Err.Number = 0 And Not flatPatternCmd Is Nothing Then
                        LogMessage "Found command: " & cmdId
                        Exit For
                    End If
                    Err.Clear
                Next
            End If

            If Not flatPatternCmd Is Nothing Then
                flatPatternCmd.Execute
                If Err.Number <> 0 Then
                    LogMessage "ERROR: Flat pattern command execution failed: " & Err.Description
                    Err.Clear
                    Exit Function
                End If

                ' Wait for completion
                WScript.Sleep 1000
            Else
                LogMessage "ERROR: Could not find flat pattern command"
                Exit Function
            End If
        Else
            LogMessage "Unfold method succeeded"
        End If
    End If

    ' Get flat pattern dimensions
    If smDef.HasFlatPattern Then
        Dim flatPattern
        Set flatPattern = smDef.FlatPattern

        If Err.Number <> 0 Then
            LogMessage "ERROR: Could not get FlatPattern object: " & Err.Description
            Err.Clear
            Exit Function
        End If

        ' Get dimensions (convert cm to mm)
        outLength = flatPattern.Length * 10.0
        outWidth = flatPattern.Width * 10.0

        LogMessage "Flat pattern dimensions retrieved: " & FormatNumber(outLength, 2) & "mm x " & FormatNumber(outWidth, 2) & "mm"
        CreateFlatPatternWithDimensions = True
    Else
        LogMessage "ERROR: HasFlatPattern is still False after creation attempt"
        Exit Function
    End If
End Function

Sub FixFlatPatternOrientation(partDoc, ByRef outLength, ByRef outWidth)
    On Error Resume Next

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Not smDef.HasFlatPattern Then
        LogMessage "No flat pattern exists - cannot fix orientation"
        Exit Sub
    End If

    Dim flatPattern
    Set flatPattern = smDef.FlatPattern

    ' Check if one dimension is suspiciously small (edge view)
    ' For platework, both dimensions should be reasonable
    ' If one dimension matches the thickness (< 50mm), it's probably showing the edge
    Dim minDim, maxDim
    If outWidth < outLength Then
        minDim = outWidth
        maxDim = outLength
    Else
        minDim = outLength
        maxDim = outWidth
    End If

    If minDim < 50 Then
        LogMessage "WARNING: One dimension is only " & FormatNumber(minDim, 2) & "mm - this appears to be EDGE VIEW"
        LogMessage "The flat pattern is showing the thickness edge instead of the top face"
        LogMessage "Attempting to fix orientation..."

        ' Method 1: Try Refold and re-Unfold with largest planar face selection
        LogMessage "Method 1: Trying Refold and re-Unfold..."
        
        ' First, refold the flat pattern
        smDef.Refold
        Err.Clear
        
        ' Wait for refold to complete
        WScript.Sleep 500
        
        ' Now try to find the largest planar face and unfold from it
        Dim faces
        Set faces = smDef.SurfaceBodies.Item(1).Faces
        
        Dim largestFace
        Dim largestArea
        largestArea = 0
        
        Dim i
        For i = 1 To faces.Count
            Dim face
            Set face = faces.Item(i)
            
            ' Check if this is a planar face
            Dim surfType
            surfType = face.SurfaceType
            
            ' SurfaceType 0 = Plane
            If surfType = 0 Then
                Dim faceArea
                faceArea = face.Evaluator.Area
                
                If faceArea > largestArea Then
                    largestArea = faceArea
                    Set largestFace = face
                End If
            End If
            Err.Clear
        Next
        
        If Not largestFace Is Nothing Then
            LogMessage "Found largest planar face with area: " & FormatNumber(largestArea * 100, 2) & " mm²"
            
            ' Try to unfold using the largest face
            smDef.Unfold largestFace
            Err.Clear
            
            WScript.Sleep 500
            
            ' Re-read dimensions
            If smDef.HasFlatPattern Then
                Set flatPattern = smDef.FlatPattern
                outLength = flatPattern.Length * 10
                outWidth = flatPattern.Width * 10
                LogMessage "New flat pattern dimensions: " & FormatNumber(outLength, 2) & "mm x " & FormatNumber(outWidth, 2) & "mm"
                
                ' Check if it's fixed
                If outWidth >= 50 And outLength >= 50 Then
                    LogMessage "Orientation fixed successfully!"
                    Exit Sub
                End If
            End If
        End If
        
        ' Method 2: Try FlipBaseFace (original method)
        LogMessage "Method 2: Trying FlipBaseFace..."
        
        If smDef.HasFlatPattern Then
            Set flatPattern = smDef.FlatPattern
            
            flatPattern.Edit
            If Err.Number = 0 Then
                flatPattern.FlipBaseFace
                If Err.Number = 0 Then
                    flatPattern.ExitEdit
                    partDoc.Update
                    
                    outLength = flatPattern.Length * 10
                    outWidth = flatPattern.Width * 10
                    LogMessage "After FlipBaseFace: " & FormatNumber(outLength, 2) & "mm x " & FormatNumber(outWidth, 2) & "mm"
                Else
                    LogMessage "FlipBaseFace failed: " & Err.Description
                    Err.Clear
                    flatPattern.ExitEdit
                End If
            Else
                LogMessage "Could not enter edit mode: " & Err.Description
                Err.Clear
            End If
        End If
        
        ' If still wrong, log warning
        If outWidth < 50 Or outLength < 50 Then
            LogMessage "WARNING: Could not automatically fix flat pattern orientation"
            LogMessage "The flat pattern may need manual adjustment in Inventor"
            LogMessage "Current dimensions: " & FormatNumber(outLength, 2) & "mm x " & FormatNumber(outWidth, 2) & "mm"
        End If
    Else
        LogMessage "Flat pattern dimensions OK: " & FormatNumber(outLength, 2) & "mm x " & FormatNumber(outWidth, 2) & "mm"
    End If
End Sub

Sub AddPlateCustomProperties(partDoc)
    On Error Resume Next

    LogMessage "Adding PLATE LENGTH and PLATE WIDTH custom iProperties..."

    ' Get custom property set
    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get custom property set: " & Err.Description
        Err.Clear
        Exit Sub
    End If

    ' Add or update PLATE LENGTH with formula
    Dim lengthProp
    On Error Resume Next
    Set lengthProp = customPropSet.Item("PLATE LENGTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it with formula
        Err.Clear
        LogMessage "Adding PLATE LENGTH = =<SHEET METAL LENGTH>"
        customPropSet.Add "=<SHEET METAL LENGTH>", "PLATE LENGTH"

        If Err.Number <> 0 Then
            LogMessage "ERROR: Could not add PLATE LENGTH: " & Err.Description
            Err.Clear
        Else
            LogMessage "PLATE LENGTH added successfully"
            ' Set precision to 0 decimal places
            Set lengthProp = customPropSet.Item("PLATE LENGTH")
            If Not lengthProp Is Nothing Then
                lengthProp.DisplayString = "0"
                LogMessage "Set PLATE LENGTH precision to 0 decimals"
            End If
        End If
    Else
        ' Property exists - update with formula
        LogMessage "Updating PLATE LENGTH with formula"
        lengthProp.Value = "=<SHEET METAL LENGTH>"
        If Err.Number <> 0 Then
            LogMessage "WARNING: Could not update PLATE LENGTH: " & Err.Description
            Err.Clear
        Else
            LogMessage "PLATE LENGTH updated successfully"
            ' Set precision to 0 decimal places
            lengthProp.DisplayString = "0"
            LogMessage "Set PLATE LENGTH precision to 0 decimals"
        End If
    End If
    Err.Clear

    ' Add or update PLATE WIDTH with formula
    Dim widthProp
    On Error Resume Next
    Set widthProp = customPropSet.Item("PLATE WIDTH")

    If Err.Number <> 0 Then
        ' Property doesn't exist, add it with formula
        Err.Clear
        LogMessage "Adding PLATE WIDTH = =<SHEET METAL WIDTH>"
        customPropSet.Add "=<SHEET METAL WIDTH>", "PLATE WIDTH"

        If Err.Number <> 0 Then
            LogMessage "ERROR: Could not add PLATE WIDTH: " & Err.Description
            Err.Clear
        Else
            LogMessage "PLATE WIDTH added successfully"
            ' Set precision to 0 decimal places
            Set widthProp = customPropSet.Item("PLATE WIDTH")
            If Not widthProp Is Nothing Then
                widthProp.DisplayString = "0"
                LogMessage "Set PLATE WIDTH precision to 0 decimals"
            End If
        End If
    Else
        ' Property exists - update with formula
        LogMessage "Updating PLATE WIDTH with formula"
        widthProp.Value = "=<SHEET METAL WIDTH>"
        If Err.Number <> 0 Then
            LogMessage "WARNING: Could not update PLATE WIDTH: " & Err.Description
            Err.Clear
        Else
            LogMessage "PLATE WIDTH updated successfully"
            ' Set precision to 0 decimal places
            widthProp.DisplayString = "0"
            LogMessage "Set PLATE WIDTH precision to 0 decimals"
        End If
    End If
    Err.Clear
    
    ' CRITICAL FIX: Set Document Settings for zero decimal precision
    LogMessage "Setting Document Settings to eliminate decimals..."
    Call SetDocumentSettingsForZeroDecimals(partDoc)
End Sub

Sub SetDocumentSettingsForZeroDecimals(partDoc)
    ' Sets all 3 critical Document Settings to show dimensions without decimals
    On Error Resume Next
    
    Dim params
    Set params = partDoc.ComponentDefinition.Parameters
    
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not access Parameters: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' Setting 1: Linear Dimension Precision = 0 decimals
    LogMessage "  Setting LinearDimensionPrecision = 0"
    params.LinearDimensionPrecision = 0
    If Err.Number <> 0 Then
        LogMessage "  WARNING: Could not set LinearDimensionPrecision: " & Err.Description
        Err.Clear
    End If
    
    ' Setting 2: Modeling Dimension Display = "Display as value" (34821)
    LogMessage "  Setting DimensionDisplayType = 34821 (Display as value)"
    params.DimensionDisplayType = 34821
    If Err.Number <> 0 Then
        LogMessage "  WARNING: Could not set DimensionDisplayType: " & Err.Description
        Err.Clear
    End If
    
    ' Setting 3: Default Parameter Input Display = "Display as expression" (True)
    LogMessage "  Setting DisplayParameterAsExpression = True"
    params.DisplayParameterAsExpression = True
    If Err.Number <> 0 Then
        LogMessage "  WARNING: Could not set DisplayParameterAsExpression: " & Err.Description
        Err.Clear
    End If
    
    LogMessage "  Document Settings applied successfully"
End Sub

Function ProcessPlatePart(partInfo, thickness)
    ProcessPlatePart = False ' Default to failure

    On Error Resume Next

    LogMessage "Processing part with thickness: " & thickness & "mm"

    ' Get the full path from the part info
    Dim fullPath
    fullPath = partInfo("fullPath")

    If fullPath = "" Then
        LogMessage "ERROR: Could not get part path"
        Exit Function
    End If

    LogMessage "Opening part for editing: " & fullPath

    ' CRITICAL FIX: Open the part explicitly for full editing access
    ' The reference from assembly occurrence is read-only - we need an editable document
    ' TRUE parameter makes the document VISIBLE in the UI (required for Convert dialog interaction)
    Dim partDoc
    Set partDoc = m_InventorApp.Documents.Open(fullPath, True)

    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to open part: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If partDoc Is Nothing Then
        LogMessage "ERROR: Could not open part document"
        Exit Function
    End If

    LogMessage "Part opened successfully for editing: " & partDoc.FullFileName

    ' Check if already sheet metal
    Dim isSheetMetal
    isSheetMetal = CheckIfPartIsSheetMetal(partDoc)

    If isSheetMetal Then
        LogMessage "Part is already sheet metal - processing..."
    Else
        ' Standard part - needs conversion
        LogMessage "Part is STANDARD type - converting to sheet metal..."
        
        ' Call the conversion function (like TEST script does)
        ConvertPartToSheetMetal partDoc, CDbl(thickness)
        
        ' Verify conversion succeeded
        If Not CheckIfPartIsSheetMetal(partDoc) Then
            LogMessage "ERROR: Part is still not sheet metal after conversion attempt"
            LogMessage "This part may require manual conversion - please convert manually in Inventor"
            
            ' Add to skipped parts list for reporting
            If m_SkippedParts Is Nothing Then
                Set m_SkippedParts = CreateObject("Scripting.Dictionary")
            End If
            m_SkippedParts.Add partDoc.FullFileName, "Conversion failed - not sheet metal"
            
            partDoc.Close True
            Exit Function
        End If
        
        LogMessage "Part successfully converted to sheet metal"
    End If

    ' Detect actual thickness from geometry
    Dim actualThickness
    actualThickness = GetPartThicknessFromGeometry(partDoc)
    If actualThickness > 0 Then
        LogMessage "Detected thickness from geometry: " & FormatNumber(actualThickness * 10, 2) & "mm"
    Else
        LogMessage "Using description thickness: " & thickness & "mm"
        actualThickness = CDbl(thickness) / 10.0 ' Convert mm to cm
    End If

    ' Set thickness
    SetSheetMetalThickness partDoc, actualThickness

    ' Create flat pattern
    Dim flatLength, flatWidth
    Dim flatPatternSuccess
    flatPatternSuccess = CreateFlatPatternWithDimensions(partDoc, flatLength, flatWidth)
    If Not flatPatternSuccess Then
        LogMessage "ERROR: Flat pattern creation failed"
        Err.Clear
        partDoc.Close True ' Close without saving
        Exit Function
    End If

    ' Check and fix orientation if needed
    FixFlatPatternOrientation partDoc, flatLength, flatWidth

    ' Add custom iProperties
    AddPlateCustomProperties partDoc

    ' Save the part
    partDoc.Save
    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to save part: " & Err.Description
        Err.Clear
        partDoc.Close True
        Exit Function
    End If

    LogMessage "Part processed and saved successfully: " & partDoc.FullFileName

    ' Close the part to free memory (it's saved, assembly will reload it)
    partDoc.Close False

    ProcessPlatePart = True ' Success
End Function

Sub ConvertPartToSheetMetal(partDoc, thickness)
    On Error Resume Next

    LogMessage "Converting to sheet metal with thickness: " & thickness & "mm"

    ' Part must be the active document for command to work
    LogMessage "Activating part document..."
    partDoc.Activate
    
    ' Bring Inventor to front and make fully visible
    m_InventorApp.WindowState = 1
    m_InventorApp.Visible = True
    
    ' Force view update and wait for activation to complete
    If Not m_InventorApp.ActiveView Is Nothing Then
        m_InventorApp.ActiveView.Update
    End If
    
    ' CRITICAL: Wait longer for part to fully activate and UI to update
    LogMessage "Waiting for part activation and UI update..."
    WScript.Sleep 2000
    
    ' Create WshShell for SendKeys
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim compDef
    Set compDef = partDoc.ComponentDefinition
    
    ' Step 1: Revert to standard part if already sheet metal
    If partDoc.SubType = kSheetMetalSubType Then
        LogMessage "Part is already sheet metal - reverting to standard first..."
        
        ' Delete flat pattern if exists
        If compDef.HasFlatPattern Then
            compDef.FlatPattern.Delete
            If Err.Number <> 0 Then
                LogMessage "Warning: Could not delete flat pattern: " & Err.Description
                Err.Clear
            Else
                partDoc.Update
                LogMessage "Flat pattern deleted"
            End If
        End If
        
        ' Execute revert command
        Dim cmdMgr
        Set cmdMgr = m_InventorApp.CommandManager
        
        Dim revertCmd
        Set revertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToStandardPartCmd")
        
        If Not revertCmd Is Nothing And revertCmd.Enabled Then
            LogMessage "Executing revert to standard part..."
            revertCmd.Execute
            WScript.Sleep 1500
            partDoc.Update
            Set compDef = partDoc.ComponentDefinition
            LogMessage "Reverted to standard part"
        Else
            LogMessage "Warning: Could not revert - command not available"
        End If
    End If
    
    ' Re-get component definition after possible revert
    Set compDef = partDoc.ComponentDefinition
    
    ' Step 2: Save document to prevent dirty state issues with SelectSet
    LogMessage ""
    LogMessage "Saving document to ensure SelectSet works..."
    
    On Error Resume Next
    partDoc.Save
    
    If Err.Number <> 0 Then
        LogMessage "Warning: Could not save - " & Err.Description
        Err.Clear
    Else
        LogMessage "Document saved"
    End If
    
    ' Step 3: Find the largest face for correct flat pattern orientation
    LogMessage ""
    LogMessage "Finding largest face for correct flat pattern orientation..."
    
    Dim body, faces, face, largestFace, largestArea
    largestArea = 0
    
    Set body = compDef.SurfaceBodies.Item(1)
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not access surface body: " & Err.Description
        Exit Sub
    End If
    
    Set faces = body.Faces
    
    For Each face In faces
        Dim area
        area = face.Evaluator.Area * 100 ' Convert to mm²
        If Err.Number = 0 And area > largestArea Then
            largestArea = area
            Set largestFace = face
        End If
        Err.Clear
    Next
    
    If largestFace Is Nothing Then
        LogMessage "ERROR: Could not find largest face"
        Exit Sub
    End If
    
    LogMessage "Largest face area: " & FormatNumber(largestArea, 0) & " mm²"
    
    ' Step 4: Pre-select the largest face
    LogMessage ""
    LogMessage "Pre-selecting largest face..."
    
    Dim selectSet
    Set selectSet = partDoc.SelectSet
    selectSet.Clear
    selectSet.Select largestFace
    
    If Err.Number <> 0 Then
        LogMessage "Warning: Could not pre-select face: " & Err.Description
        Err.Clear
    Else
        LogMessage "Face pre-selected (SelectSet.Count = " & selectSet.Count & ")"
        
        ' Force view update to show green selection
        If Not m_InventorApp.ActiveView Is Nothing Then
            m_InventorApp.ActiveView.Update
        End If
        WScript.Sleep 1000
    End If
    
    ' Step 5: Execute Convert to Sheet Metal command
    LogMessage ""
    LogMessage "Executing Convert to Sheet Metal command..."
    
    Set cmdMgr = m_InventorApp.CommandManager
    
    Dim convertCmd
    Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
    
    If convertCmd Is Nothing Or Not convertCmd.Enabled Then
        LogMessage "ERROR: Convert command not available"
        Exit Sub
    End If
    
    ' Bring Inventor to foreground
    m_InventorApp.WindowState = 1 ' kNormalWindow
    m_InventorApp.Visible = True
    
    ' Activate the Inventor window
    WshShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    ' Execute the command
    LogMessage "Executing: " & convertCmd.DisplayName
    convertCmd.Execute
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: Command execution failed: " & Err.Description
        Exit Sub
    End If
    
    WScript.Sleep 1500
    
    ' Step 6: User interaction required
    LogMessage "Waiting for user to click face..."
    
    ' Ensure Inventor is in front and visible
    m_InventorApp.WindowState = 1
    m_InventorApp.Visible = True
    WshShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    Dim userResponse
    userResponse = MsgBox("ACTION REQUIRED" & vbCrLf & vbCrLf & _
                          "Part: " & partDoc.DisplayName & vbCrLf & vbCrLf & _
                          "1. The Convert to Sheet Metal dialog is now open" & vbCrLf & _
                          "2. CLICK ONCE on the GREEN highlighted face in Inventor" & vbCrLf & _
                          "3. Then click OK here to continue" & vbCrLf & vbCrLf & _
                          "(The large face should be highlighted GREEN)", vbOKCancel + vbExclamation, "Click Face - " & partDoc.DisplayName)
    
    If userResponse = vbCancel Then
        LogMessage "User cancelled conversion for this part"
        partDoc.Close True
        Exit Sub
    End If
    
    ' Step 7: Complete the conversion
    LogMessage "User confirmed face click - completing conversion..."
    WshShell.AppActivate "Autodesk Inventor"
    WScript.Sleep 500
    
    ' Send Enter for the Sheet Metal Defaults dialog
    WshShell.SendKeys "{ENTER}"
    WScript.Sleep 1500
    
    ' Step 8: Verify conversion
    partDoc.Update
    Set compDef = partDoc.ComponentDefinition
    
    If partDoc.SubType = kSheetMetalSubType Then
        LogMessage "VERIFIED: Part is now sheet metal type"
        
        ' Re-get component definition as SheetMetalComponentDefinition
        Set compDef = partDoc.ComponentDefinition
        
        ' Step 9: Create flat pattern if not exists
        If Not compDef.HasFlatPattern Then
            LogMessage "Creating flat pattern..."
            Err.Clear
            compDef.Unfold
            If Err.Number <> 0 Then
                LogMessage "ERROR creating flat pattern: " & Err.Description
                Err.Clear
            Else
                partDoc.Update
                LogMessage "Flat pattern created"
            End If
        Else
            LogMessage "Flat pattern already exists"
        End If
        
        ' Verify flat pattern dimensions
        If compDef.HasFlatPattern Then
            Dim fp
            Set fp = compDef.FlatPattern
            Dim fpLength, fpWidth
            fpLength = fp.Length * 10
            fpWidth = fp.Width * 10
            
            LogMessage "Flat pattern dimensions: " & FormatNumber(fpLength, 1) & " x " & FormatNumber(fpWidth, 1) & " mm"
            
            ' Check if orientation is correct
            If fpLength > 100 And fpWidth > 100 Then
                LogMessage "SUCCESS: Flat pattern shows correct orientation (large face)"
            Else
                LogMessage "WARNING: Flat pattern may still show edge view - check manually"
            End If
        End If
    Else
        LogMessage "ERROR: Conversion failed - part SubType is: " & partDoc.SubType
    End If
End Sub

' OLD CreateFlatPattern - KEPT FOR COMPATIBILITY (not used by new flow)
Sub CreateFlatPattern(partDoc)
    On Error Resume Next

    LogMessage "Creating flat pattern"

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    ' Create flat pattern
    compDef.Unfold

    ' Get flat pattern and verify it was created
    Dim flatPattern
    Set flatPattern = compDef.FlatPattern

    If flatPattern Is Nothing Then
        LogMessage "ERROR: Flat pattern creation failed - no flat pattern object"
        Err.Raise vbObjectError + 1002, "CreateFlatPattern", "Flat pattern creation failed"
    End If

    Dim length
    length = flatPattern.Length * 10.0 ' Convert to mm
    Dim width
    width = flatPattern.Width * 10.0 ' Convert to mm

    ' Validate dimensions are reasonable (not zero)
    If length <= 0 Or width <= 0 Then
        LogMessage "ERROR: Flat pattern has invalid dimensions: " & length & "mm x " & width & "mm"
        Err.Raise vbObjectError + 1003, "CreateFlatPattern", "Flat pattern has invalid dimensions"
    End If

    ' Store the dimensions globally for the last processed part
    m_LastSheetMetalLength = length
    m_LastSheetMetalWidth = width

    LogMessage "Flat pattern created successfully - Length: " & FormatNumber(length, 2) & "mm, Width: " & FormatNumber(width, 2) & "mm"
End Sub

Sub AddPlateParametersToAssembly(asmDoc, plateParts)
    On Error Resume Next

    LogMessage "Adding PLATE LENGTH and PLATE WIDTH parameters"

    ' Find maximum dimensions from all processed flat patterns
    Dim maxLength
    maxLength = 0
    Dim maxWidth
    maxWidth = 0

    Dim compDef

    Dim thickness
    For Each thickness In plateParts.Keys
        Dim parts
        Set parts = plateParts(thickness)

        Dim i
        For i = 0 To parts.Count - 1
            Dim partInfo
            Set partInfo = parts(i)

            ' Check if this part was successfully processed (has dimensions > 0)
            ' We need to check the part document to see if it has a flat pattern
            On Error Resume Next
            Dim partDoc
            Set partDoc = partInfo("document")

            If Not partDoc Is Nothing Then
                Set compDef = partDoc.ComponentDefinition

                Dim flatPattern
                Set flatPattern = compDef.FlatPattern

                If Not flatPattern Is Nothing Then
                    Dim length
                    length = flatPattern.Length * 10.0 ' Convert to mm
                    Dim width
                    width = flatPattern.Width * 10.0 ' Convert to mm

                    If length > maxLength Then maxLength = length
                    If width > maxWidth Then maxWidth = width

                    LogMessage "Found flat pattern dimensions: " & FormatNumber(length, 2) & "mm x " & FormatNumber(width, 2) & "mm"
                End If
            End If
            Err.Clear
        Next
    Next

    LogMessage "Maximum dimensions found: Length=" & FormatNumber(maxLength, 2) & "mm, Width=" & FormatNumber(maxWidth, 2) & "mm"

    Set compDef = asmDoc.ComponentDefinition

    Dim userParams
    Set userParams = compDef.Parameters.UserParameters

    ' Add or update PLATE LENGTH parameter
    Dim lengthParam
    On Error Resume Next

    ' Try to get existing parameter first
    Set lengthParam = userParams.Item("PLATE LENGTH")
    If Err.Number <> 0 Then
        ' Parameter doesn't exist, create it using Add method
        Err.Clear
        Set lengthParam = userParams.Add("PLATE LENGTH", kNumberParameterType)
        lengthParam.Value = maxLength
        lengthParam.Units = "mm"
        LogMessage "Created parameter: PLATE LENGTH = " & FormatNumber(maxLength, 2) & " mm"
    Else
        ' Parameter exists, update its value
        lengthParam.Value = maxLength
        LogMessage "Updated parameter: PLATE LENGTH = " & FormatNumber(maxLength, 2) & " mm"
    End If
    Err.Clear

    ' Add or update PLATE WIDTH parameter
    Dim widthParam
    On Error Resume Next

    ' Try to get existing parameter first
    Set widthParam = userParams.Item("PLATE WIDTH")
    If Err.Number <> 0 Then
        ' Parameter doesn't exist, create it using Add method
        Err.Clear
        Set widthParam = userParams.Add("PLATE WIDTH", kNumberParameterType)
        widthParam.Value = maxWidth
        widthParam.Units = "mm"
        LogMessage "Created parameter: PLATE WIDTH = " & FormatNumber(maxWidth, 2) & " mm"
    Else
        ' Parameter exists, update its value
        widthParam.Value = maxWidth
        LogMessage "Updated parameter: PLATE WIDTH = " & FormatNumber(maxWidth, 2) & " mm"
    End If
    Err.Clear

    ' Force a document update to ensure parameters are committed
    asmDoc.Update
    LogMessage "Document updated to commit parameter changes"
End Sub

Sub LogMessage(message)
    Dim timestamp
    timestamp = FormatDateTime(Now, vbShortTime)
    m_Log = m_Log & timestamp & " | " & message & vbCrLf
    ' Also output to console for debugging
    WScript.Echo message
End Sub

Sub SaveLog()
    On Error Resume Next

    m_LogPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & _
                "\SheetMetalConverter_Log_" & Format(Now, "yyyyMMdd_HHmmss") & ".txt"
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFile
    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.Write m_Log
    logFile.Close

    LogMessage "Log saved to: " & m_LogPath
End Sub

' Run the main function
Main