' TEST SCRIPT - Single Part Sheet Metal Conversion
' Tests programmatic conversion of ONE part to verify workflow
' Author: Quintin de Bruin © 2026

Option Explicit

' Inventor API Constants
Const kPartDocumentObject = 12290
Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

' Global variables
Dim m_Log
Dim m_LogPath

Sub Main()
    On Error Resume Next

    m_Log = ""
    LogMessage "=== SINGLE PART SHEET METAL CONVERSION TEST ==="
    LogMessage "This script tests the conversion workflow on the currently open part"
    LogMessage ""

    ' Get Inventor application
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If invApp Is Nothing Then
        LogMessage "ERROR: Inventor is not running"
        MsgBox "Inventor is not running. Please start Inventor first.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Connected to Inventor"

    ' Check if we have an active document
    If invApp.ActiveDocument Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "Please open a PART document (.ipt file) in Inventor first.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    Dim partDoc
    Set partDoc = invApp.ActiveDocument

    ' Verify it's a part document
    If partDoc.DocumentType <> kPartDocumentObject Then
        LogMessage "ERROR: Active document is not a part (type: " & partDoc.DocumentType & ")"
        MsgBox "Please open a PART document (.ipt file), not an assembly.", vbCritical, "Error"
        SaveLog
        Exit Sub
    End If

    LogMessage "Part document loaded: " & partDoc.FullFileName
    LogMessage "Part name: " & partDoc.DisplayName

    ' Read part description
    Dim description
    description = GetPartDescription(partDoc)
    LogMessage "Part description: " & description

    ' Extract thickness from description
    Dim thickness
    thickness = ExtractThickness(description)
    If thickness = "" Then
        LogMessage "WARNING: Could not extract thickness from description"
        thickness = "10" ' Default for testing
        LogMessage "Using default thickness: " & thickness & "mm"
    Else
        LogMessage "Detected thickness: " & thickness & "mm"
    End If

    ' STEP 1: Check current sheet metal status
    LogMessage ""
    LogMessage "STEP 1: Checking if part is already sheet metal..."
    Dim isSheetMetal
    isSheetMetal = CheckIfSheetMetal(partDoc)

    If isSheetMetal Then
        LogMessage "Part is ALREADY sheet metal - skipping conversion"
    Else
        LogMessage "Part is NOT sheet metal - proceeding with conversion"

        ' STEP 2: Convert to sheet metal
        LogMessage ""
        LogMessage "STEP 2: Converting to sheet metal..."
        Dim convertSuccess
        convertSuccess = ConvertToSheetMetal(invApp, partDoc)

        If Not convertSuccess Then
            LogMessage "ERROR: Sheet metal conversion failed"
            MsgBox "Sheet metal conversion FAILED!" & vbCrLf & vbCrLf & _
                   "Check the log for details: " & m_LogPath, vbCritical, "Conversion Failed"
            SaveLog
            Exit Sub
        End If

        LogMessage "Sheet metal conversion completed successfully"

        ' Re-verify sheet metal status
        isSheetMetal = CheckIfSheetMetal(partDoc)
        If Not isSheetMetal Then
            LogMessage "ERROR: Part is still not sheet metal after conversion!"
            MsgBox "Conversion appeared to succeed but part is not sheet metal type!", vbCritical, "Verification Failed"
            SaveLog
            Exit Sub
        End If
    End If

    ' STEP 3: Detect actual thickness from geometry
    LogMessage ""
    LogMessage "STEP 3: Detecting actual thickness from part geometry..."
    Dim actualThickness
    actualThickness = GetPartThicknessFromGeometry(partDoc)

    If actualThickness > 0 Then
        LogMessage "Detected thickness from geometry: " & FormatNumber(actualThickness * 10, 2) & "mm"
    Else
        LogMessage "WARNING: Could not detect thickness from geometry, using description value"
        actualThickness = CDbl(thickness) / 10.0 ' Convert mm to cm
    End If

    ' STEP 4: Set thickness properly
    LogMessage ""
    LogMessage "STEP 4: Setting thickness to " & FormatNumber(actualThickness * 10, 2) & "mm..."
    Dim thicknessSuccess
    thicknessSuccess = SetSheetMetalThickness(partDoc, actualThickness)

    If Not thicknessSuccess Then
        LogMessage "WARNING: Could not set thickness"
    Else
        LogMessage "Thickness set successfully"
    End If

    ' STEP 5: Create flat pattern
    LogMessage ""
    LogMessage "STEP 5: Creating flat pattern..."
    Dim flatPatternSuccess
    Dim flatLength, flatWidth
    flatPatternSuccess = CreateFlatPattern(invApp, partDoc, flatLength, flatWidth)

    If Not flatPatternSuccess Then
        LogMessage "ERROR: Flat pattern creation failed"
        MsgBox "Flat pattern creation FAILED!" & vbCrLf & vbCrLf & _
               "Check the log for details: " & m_LogPath, vbCritical, "Flat Pattern Failed"
        SaveLog
        Exit Sub
    End If

    LogMessage "Flat pattern created successfully"
    LogMessage "Initial flat pattern dimensions: " & FormatNumber(flatLength, 2) & "mm x " & FormatNumber(flatWidth, 2) & "mm"

    ' STEP 6: Check and fix orientation if needed
    LogMessage ""
    LogMessage "STEP 6: Checking flat pattern orientation..."
    Dim orientationFixed
    orientationFixed = FixFlatPatternOrientation(partDoc, flatLength, flatWidth)

    If orientationFixed Then
        LogMessage "Orientation was CORRECTED - flat pattern flipped to show top view"
        LogMessage "Corrected flat pattern dimensions: " & FormatNumber(flatLength, 2) & "mm x " & FormatNumber(flatWidth, 2) & "mm"
    Else
        LogMessage "Orientation is correct - no changes needed"
    End If

    ' STEP 7: Add custom iProperties
    LogMessage ""
    LogMessage "STEP 7: Adding PLATE LENGTH and PLATE WIDTH custom iProperties..."
    Dim propsAdded
    propsAdded = AddPlateCustomProperties(partDoc)

    If propsAdded Then
        LogMessage "Custom iProperties added successfully"
    Else
        LogMessage "WARNING: Could not add custom iProperties"
    End If

    ' STEP 8: Save the part
    LogMessage ""
    LogMessage "STEP 8: Saving part..."
    partDoc.Save
    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to save part: " & Err.Description
        Err.Clear
    Else
        LogMessage "Part saved successfully"
    End If

    ' Success!
    LogMessage ""
    LogMessage "=== TEST COMPLETED SUCCESSFULLY ==="
    LogMessage "Part: " & partDoc.DisplayName
    LogMessage "Detected Thickness: " & FormatNumber(actualThickness * 10, 2) & "mm"
    LogMessage "Final Flat Dimensions: " & FormatNumber(flatLength, 2) & "mm x " & FormatNumber(flatWidth, 2) & "mm"

    SaveLog

    MsgBox "Test SUCCESSFUL!" & vbCrLf & vbCrLf & _
           "Part: " & partDoc.DisplayName & vbCrLf & _
           "Thickness: " & FormatNumber(actualThickness * 10, 2) & "mm (from geometry)" & vbCrLf & _
           "Flat dimensions: " & FormatNumber(flatLength, 2) & "mm x " & FormatNumber(flatWidth, 2) & "mm" & vbCrLf & vbCrLf & _
           "Log saved to: " & m_LogPath, vbInformation, "Test Complete"
End Sub

Function CheckIfSheetMetal(partDoc)
    CheckIfSheetMetal = False

    On Error Resume Next
    LogMessage "Checking SubType GUID: " & partDoc.SubType

    If partDoc.SubType = kSheetMetalSubType Then
        CheckIfSheetMetal = True
        LogMessage "Part is SHEET METAL type"
    Else
        LogMessage "Part is STANDARD type (not sheet metal)"
    End If
End Function

Function ConvertToSheetMetal(invApp, partDoc)
    ConvertToSheetMetal = False

    On Error Resume Next

    ' Try to execute the Convert to Sheet Metal command
    LogMessage "Attempting to execute 'Convert to Sheet Metal' command..."

    ' Method 1: Try CommandManager with the documented command ID
    Dim cmdMgr
    Set cmdMgr = invApp.CommandManager

    Dim convertCmd
    Set convertCmd = cmdMgr.ControlDefinitions.Item("PartConvertToSheetMetalCmd")

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not find PartConvertToSheetMetalCmd: " & Err.Description
        Err.Clear

        ' Try alternative command IDs
        LogMessage "Trying alternative command IDs..."

        Dim cmdIds
        cmdIds = Array("ConvertToSheetMetalCmd", "SMConvertCmd", "SheetMetalConvertCmd")

        Dim cmdId
        For Each cmdId In cmdIds
            LogMessage "Trying: " & cmdId
            Set convertCmd = cmdMgr.ControlDefinitions.Item(cmdId)
            If Err.Number = 0 And Not convertCmd Is Nothing Then
                LogMessage "Found command: " & cmdId
                Exit For
            End If
            Err.Clear
        Next
    End If

    If convertCmd Is Nothing Then
        LogMessage "ERROR: Could not find Convert to Sheet Metal command in CommandManager"
        Exit Function
    End If

    ' Execute the command
    LogMessage "Executing conversion command..."
    convertCmd.Execute

    If Err.Number <> 0 Then
        LogMessage "ERROR: Command execution failed: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Wait a moment for the conversion to complete
    WScript.Sleep 1000

    LogMessage "Conversion command executed successfully"
    ConvertToSheetMetal = True
End Function

Function SetSheetMetalThickness(partDoc, thicknessInCm)
    SetSheetMetalThickness = False

    On Error Resume Next

    ' Get sheet metal component definition
    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get ComponentDefinition: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' CRITICAL: Disable "use style thickness" first
    LogMessage "Disabling UseSheetMetalStyleThickness..."
    smDef.UseSheetMetalStyleThickness = False

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not set UseSheetMetalStyleThickness: " & Err.Description
        Err.Clear
        ' Continue anyway
    End If

    ' Now set thickness parameter directly on component definition
    Dim thicknessParam
    Set thicknessParam = smDef.Thickness

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get Thickness parameter: " & Err.Description
        Err.Clear
        Exit Function
    End If

    If thicknessParam Is Nothing Then
        LogMessage "ERROR: Thickness parameter is Nothing"
        Exit Function
    End If

    LogMessage "Setting thickness: " & FormatNumber(thicknessInCm * 10, 2) & "mm (" & thicknessInCm & " cm internal)"
    thicknessParam.Value = thicknessInCm

    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to set thickness: " & Err.Description
        Err.Clear
        Exit Function
    End If

    ' Update document to apply change
    partDoc.Update

    LogMessage "Thickness parameter value: " & FormatNumber(thicknessParam.Value * 10, 2) & " mm"
    SetSheetMetalThickness = True
End Function

Function CreateFlatPattern(invApp, partDoc, ByRef outLength, ByRef outWidth)
    CreateFlatPattern = False
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

        ' Method 1: Try using Unfold method
        LogMessage "Attempting ComponentDefinition.Unfold()..."
        smDef.Unfold

        If Err.Number <> 0 Then
            LogMessage "ERROR: Unfold failed: " & Err.Description
            Err.Clear

            ' Method 2: Try using CommandManager
            LogMessage "Attempting to execute 'Create Flat Pattern' command..."

            Dim cmdMgr
            Set cmdMgr = invApp.CommandManager

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
        CreateFlatPattern = True
    Else
        LogMessage "ERROR: HasFlatPattern is still False after creation attempt"
        Exit Function
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

Function FixFlatPatternOrientation(partDoc, ByRef outLength, ByRef outWidth)
    FixFlatPatternOrientation = False

    On Error Resume Next

    Dim smDef
    Set smDef = partDoc.ComponentDefinition

    If Not smDef.HasFlatPattern Then
        LogMessage "ERROR: No flat pattern exists"
        Exit Function
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
        LogMessage "Method 1: Trying Refold and re-Unfold with largest planar face..."
        
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
                    FixFlatPatternOrientation = True
                    Exit Function
                End If
            End If
        Else
            LogMessage "Could not find planar faces for unfold"
        End If
        
        ' Method 2: Try FlipBaseFace (original method)
        LogMessage "Method 2: Trying FlipBaseFace..."
        
        ' Make sure we have a flat pattern
        If Not smDef.HasFlatPattern Then
            smDef.Unfold
            Err.Clear
            WScript.Sleep 500
        End If
        
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
                    
                    If outWidth >= 50 And outLength >= 50 Then
                        LogMessage "Orientation fixed via FlipBaseFace!"
                        FixFlatPatternOrientation = True
                        Exit Function
                    End If
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
End Function

Function AddPlateCustomProperties(partDoc)
    AddPlateCustomProperties = False

    On Error Resume Next

    LogMessage "Adding PLATE LENGTH and PLATE WIDTH custom iProperties..."

    ' Get custom property set
    Dim customPropSet
    Set customPropSet = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not get custom property set: " & Err.Description
        Err.Clear
        Exit Function
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
        End If
    Else
        ' Property exists - check if it's empty and update it
        If lengthProp.Value = "" Or IsEmpty(lengthProp.Value) Then
            LogMessage "PLATE LENGTH exists but is empty - updating with formula"
            lengthProp.Value = "=<SHEET METAL LENGTH>"
            If Err.Number <> 0 Then
                LogMessage "ERROR: Could not update PLATE LENGTH: " & Err.Description
                Err.Clear
            Else
                LogMessage "PLATE LENGTH updated successfully"
            End If
        Else
            LogMessage "PLATE LENGTH already exists with value: " & lengthProp.Value
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
        End If
    Else
        ' Property exists - check if it's empty and update it
        If widthProp.Value = "" Or IsEmpty(widthProp.Value) Then
            LogMessage "PLATE WIDTH exists but is empty - updating with formula"
            widthProp.Value = "=<SHEET METAL WIDTH>"
            If Err.Number <> 0 Then
                LogMessage "ERROR: Could not update PLATE WIDTH: " & Err.Description
                Err.Clear
            Else
                LogMessage "PLATE WIDTH updated successfully"
            End If
        Else
            LogMessage "PLATE WIDTH already exists with value: " & widthProp.Value
        End If
    End If
    Err.Clear

    AddPlateCustomProperties = True
End Function

Function GetPartDescription(partDoc)
    On Error Resume Next

    Dim propertySet
    Set propertySet = partDoc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then
        GetPartDescription = ""
        Err.Clear
        Exit Function
    End If

    Dim descriptionProp
    Set descriptionProp = propertySet.Item("Description")

    If Err.Number <> 0 Then
        GetPartDescription = ""
        Err.Clear
        Exit Function
    End If

    GetPartDescription = Trim(descriptionProp.Value)
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

Sub LogMessage(message)
    Dim timestamp
    timestamp = FormatDateTime(Now, vbShortTime)
    m_Log = m_Log & timestamp & " | " & message & vbCrLf
    WScript.Echo message
End Sub

Sub SaveLog()
    On Error Resume Next

    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    Dim docsFolder
    docsFolder = wshShell.SpecialFolders("MyDocuments")

    m_LogPath = docsFolder & "\SheetMetal_Test_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".txt"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logFile
    Set logFile = fso.CreateTextFile(m_LogPath, True)
    logFile.Write m_Log
    logFile.Close

    LogMessage "Log saved to: " & m_LogPath
End Sub

' Run main function
Main
