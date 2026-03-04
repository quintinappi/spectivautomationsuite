Option Explicit

' ==============================================================================
' DYNAMIC HERITAGE-BASED SOLUTION - User-Friendly Interactive Version
' ==============================================================================
' Author: Quintin de Bruin © 2025
' 
' Dynamic version that:
' 1. Detects currently open assembly in Inventor
' 2. Flattens hierarchy and groups similar components  
' 3. Asks user for naming scheme per component group
' 4. Uses proven heritage method for model + IDW updates
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_ComponentGroups ' Dictionary to store component groups
Dim g_NamingSchemes   ' Dictionary to store user-defined naming schemes
Dim g_FileNameMapping ' Dictionary to store original -> new filename mappings
Dim g_PlantSection    ' User-defined plant section prefix
Dim g_ComprehensiveMapping ' Master mapping: originalPath -> "newPath|originalFile|newFile|group|description"
Dim g_SkipExisting    ' Flag to skip existing parts or rename from 1

' Inventor document type constants
Const kPartDocumentObject = 12290
Const kAssemblyDocumentObject = 12291

' iLogic Add-in GUID for audit capture
Const ILOGIC_ADDIN_GUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

Sub SCAN_PARTS_AND_DESCRIPTIONS()
    Call StartLogging
    LogMessage "=== SCAN PARTS AND DESCRIPTIONS ==="
    LogMessage "Scanning open model for parts and their iProperties descriptions"

    Dim result
    result = MsgBox("SCAN PARTS AND DESCRIPTIONS" & vbCrLf & vbCrLf & _
                    "This will scan your currently open assembly and log each part's filename and description." & vbCrLf & vbCrLf & _
                    "Make sure your target assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Scan Parts")

    If result = vbNo Then
        LogMessage "User cancelled scan"
        Exit Sub
    End If

    ' Connect to existing Inventor application
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open your assembly first.", vbCritical
        Exit Sub
    End If

    LogMessage "SUCCESS: Connected to existing Inventor instance"
    Err.Clear

    ' Step 1: Detect open assembly
    Dim activeDoc
    Set activeDoc = DetectOpenAssembly(invApp)
    If activeDoc Is Nothing Then
        MsgBox "ERROR: No assembly is currently open in Inventor!" & vbCrLf & _
               "Please open your target assembly first.", vbCritical
        Exit Sub
    End If

    LogMessage "ASSEMBLY: Detected - " & activeDoc.DisplayName

    ' Step 2: Scan parts and log descriptions
    Call ScanPartsForDescriptions(activeDoc)

    LogMessage "=== SCAN COMPLETED ==="
    MsgBox "Scan completed. Check the log file for the list of parts and descriptions.", vbInformation
End Sub

Sub ScanPartsForDescriptions(asmDoc)
    LogMessage "SCAN: Starting recursive scan for parts and descriptions"

    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    Call RecursivelyScanAssemblyForParts(asmDoc, uniqueParts)

    LogMessage "SCAN: Found " & uniqueParts.Count & " unique parts"
End Sub

Sub RecursivelyScanAssemblyForParts(asmDoc, uniqueParts)
    Dim asmCompDef
    Set asmCompDef = asmDoc.ComponentDefinition

    Dim occ
    For Each occ In asmCompDef.Occurrences
        Dim doc
        Set doc = occ.Definition.Document

        If doc.DocumentType = kPartDocumentObject Then
            Dim fullPath
            fullPath = doc.FullFileName

            If Not uniqueParts.Exists(fullPath) Then
                uniqueParts.Add fullPath, True

                Dim fileName
                fileName = GetFileNameFromPath(fullPath)

                Dim description
                description = GetDescriptionFromIProperty(doc)

                LogMessage "PART: " & fileName & " | Description: " & description
            End If
        ElseIf doc.DocumentType = kAssemblyDocumentObject Then
            ' Recurse into sub-assembly
            Call RecursivelyScanAssemblyForParts(doc, uniqueParts)
        End If
    Next
End Sub

Call DYNAMIC_HERITAGE_SOLUTION()

Sub DYNAMIC_HERITAGE_SOLUTION()
    Call StartLogging
    LogMessage "=== DYNAMIC HERITAGE-BASED SOLUTION ==="
    LogMessage "Auto-detecting open model and creating dynamic renaming workflow"
    
    Dim result
    result = MsgBox("DYNAMIC INVENTOR RENAMING TOOL" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Detect your currently open assembly" & vbCrLf & _
                    "2. Group similar components automatically" & vbCrLf & _
                    "3. Let you define naming schemes per group" & vbCrLf & _
                    "4. Update models AND drawings automatically" & vbCrLf & vbCrLf & _
                    "Make sure your target assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Dynamic Heritage Solution")
    
    If result = vbNo Then
        LogMessage "User cancelled workflow"
        Exit Sub
    End If
    
    ' Initialize collections
    Set g_ComponentGroups = CreateObject("Scripting.Dictionary")
    Set g_NamingSchemes = CreateObject("Scripting.Dictionary")
    Set g_FileNameMapping = CreateObject("Scripting.Dictionary")
    Set g_ComprehensiveMapping = CreateObject("Scripting.Dictionary")
    
    ' Connect to existing Inventor application
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open your assembly first.", vbCritical
        Exit Sub
    End If
    
    LogMessage "SUCCESS: Connected to existing Inventor instance"
    Err.Clear
    
    ' Step 1: Detect open assembly and analyze components
    LogMessage "STEP 1: Detecting open assembly and analyzing components"
    Dim activeDoc
    Set activeDoc = DetectOpenAssembly(invApp)
    If activeDoc Is Nothing Then
        MsgBox "ERROR: No assembly is currently open in Inventor!" & vbCrLf & _
               "Please open your target assembly first.", vbCritical
        Exit Sub
    End If

    ' Step 1.5: Create audit log BEFORE any changes (capture current state)
    LogMessage "STEP 1.5: Creating audit log - capturing current state BEFORE renaming"
    Call CreateAuditLog(invApp, activeDoc)

    ' Step 2: Flatten hierarchy and group similar components
    LogMessage "STEP 2: Flattening hierarchy and grouping components"
    Call FlattenAndGroupComponents(activeDoc)
    
    ' Step 3: Get plant section naming convention
    LogMessage "STEP 3: Getting plant section naming convention"
    Call GetPlantSectionNaming()
    
    ' Step 4: Show groups summary and get user confirmation
    LogMessage "STEP 4: Showing component groups summary"
    Call ShowGroupsSummary()
    
    ' Step 5: Get user input for naming schemes
    LogMessage "STEP 5: Getting user naming schemes for component groups"
    Call GetUserNamingSchemes()
    
    ' Step 6: Create heritage-based copies with user names
    LogMessage "STEP 6: Creating heritage-based copies with user naming"
    Call CreateDynamicHeritageBasedCopies(invApp, activeDoc)
    
    ' Step 7: Update assembly references
    LogMessage "STEP 7: Updating assembly references"
    Call UpdateDynamicAssemblyReferences(invApp, activeDoc)
    
    ' Step 8: Update IDW files automatically
    LogMessage "STEP 8: Auto-detecting and updating IDW files"
    Call UpdateAllIDWsInDirectory(invApp, activeDoc)

    ' Step 8.5: Save mapping file for external IDW updates
    LogMessage "STEP 8.5: Saving mapping file for external IDW updater"
    Call SaveMappingFile(activeDoc)

    ' Step 9: Keep original files for safety - skip cleanup
    LogMessage "STEP 9: Keeping original files for safety (skipping cleanup)"
    
    LogMessage "=== DYNAMIC HERITAGE-BASED SOLUTION COMPLETED ==="
    Call StopLogging
    
    MsgBox "DYNAMIC SOLUTION COMPLETED!" & vbCrLf & vbCrLf & _
           "✅ Components analyzed and grouped" & vbCrLf & _
           "✅ Heritage-based copies created" & vbCrLf & _
           "✅ Assembly references updated" & vbCrLf & _
           "✅ IDW drawings updated automatically" & vbCrLf & _
           "✅ Mapping file saved for STEP 2" & vbCrLf & _
           "✅ Fully automated workflow!" & vbCrLf & vbCrLf & _
           "Log: " & g_LogPath, vbInformation, "Success!"
End Sub

Function DetectOpenAssembly(invApp)
    On Error Resume Next
    
    ' Check if there's an active document
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument
    
    If Err.Number <> 0 Or activeDoc Is Nothing Then
        LogMessage "No active document found"
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If
    
    ' Debug - log document info
    LogMessage "DETECTED: Document type: " & activeDoc.DocumentType
    On Error Resume Next
    LogMessage "DETECTED: Document subtype: " & activeDoc.DocumentSubType.DisplayName
    Err.Clear
    
    ' Just check by file extension - skip document type checks
    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        LogMessage "File extension is not .iam: " & activeDoc.FullFileName
        MsgBox "FILE TYPE ISSUE" & vbCrLf & vbCrLf & _
               "Current file: " & activeDoc.DisplayName & vbCrLf & _
               "Extension: " & Right(activeDoc.FullFileName, 4) & vbCrLf & vbCrLf & _
               "Need: Assembly file (.iam extension)", vbExclamation
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If
    
    LogMessage "DETECTED: Assembly by .iam extension - proceeding"

    LogMessage "DETECTED: Active assembly - " & activeDoc.DisplayName
    LogMessage "DETECTED: Full path - " & activeDoc.FullFileName

    ' Count total occurrences for validation
    Dim occCount
    On Error Resume Next
    occCount = activeDoc.ComponentDefinition.Occurrences.Count
    If Err.Number <> 0 Then occCount = 0
    Err.Clear

    ' Extract folder path for display
    Dim fso2
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    Dim folderPath
    folderPath = fso2.GetParentFolderName(activeDoc.FullFileName)

    ' Show validation prompt to user
    Dim confirmMsg
    confirmMsg = "ASSEMBLY DETECTED" & vbCrLf & vbCrLf & _
                 "Assembly: " & activeDoc.DisplayName & vbCrLf & _
                 "Parts Count: " & occCount & " occurrences" & vbCrLf & _
                 "Location: " & folderPath & vbCrLf & vbCrLf & _
                 "Is this the correct assembly to process?" & vbCrLf & vbCrLf & _
                 "⚠️  This will create heritage files for all parts!"

    Dim confirmResult
    confirmResult = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Assembly")

    If confirmResult = vbNo Then
        LogMessage "USER CANCELLED: Assembly validation failed"
        MsgBox "Operation cancelled. Please open the correct assembly and try again.", vbInformation
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If

    LogMessage "USER CONFIRMED: Proceeding with assembly processing"
    LogMessage "CONFIRMED: " & occCount & " occurrences to process"

    Set DetectOpenAssembly = activeDoc
    Err.Clear
End Function

Sub FlattenAndGroupComponents(asmDoc)
    LogMessage "ANALYZE: Recursively flattening ENTIRE model hierarchy and reading iProperty descriptions"

    ' Create a dictionary to track unique parts (prevent duplicates)
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    ' Start recursive traversal from root assembly
    Call ProcessAssemblyRecursively(asmDoc, uniqueParts, "ROOT")

    LogMessage "ANALYZE: Recursive processing completed - Total unique parts processed: " & uniqueParts.Count

    LogMessage "ANALYZE: Created " & g_ComponentGroups.Count & " component groups"
    
    ' DEBUG: Show all component groups and their contents
    Dim debugGroupKeys
    debugGroupKeys = g_ComponentGroups.Keys
    
    Dim debugI
    For debugI = 0 To UBound(debugGroupKeys)
        Dim debugGroupName
        debugGroupName = debugGroupKeys(debugI)
        
        Dim debugGroupDict
        Set debugGroupDict = g_ComponentGroups.Item(debugGroupName)
        
        LogMessage "DEBUG GROUP: '" & debugGroupName & "' contains " & debugGroupDict.Count & " components:"
        
        Dim debugCompKeys
        debugCompKeys = debugGroupDict.Keys
        
        Dim debugJ
        For debugJ = 0 To UBound(debugCompKeys)
            LogMessage "  - " & debugCompKeys(debugJ)
        Next
    Next
End Sub

Sub ProcessAssemblyRecursively(asmDoc, uniqueParts, asmLevel)
    LogMessage "ANALYZE: Processing assembly - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")"

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "ANALYZE: Found " & occurrences.Count & " occurrences in " & asmDoc.DisplayName

    ' Process each occurrence in this assembly
    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' SKIP suppressed occurrences
        If occ.Suppressed Then
            LogMessage "ANALYZE: SKIPPING (suppressed occurrence in " & asmDoc.DisplayName & ")"
        Else
            Dim doc
            Set doc = occ.Definition.Document

            Dim fileName
            fileName = GetFileNameFromPath(doc.FullFileName)

            Dim fullPath
            fullPath = doc.FullFileName

            ' Check if this is a part file
            If LCase(Right(fileName, 4)) = ".ipt" Then
                ' CRITICAL SAFETY CHECK: Skip if already heritage-renamed with SAME prefix
                ' Prevents corrupting mapping file by running STEP 1 twice with same prefix
                If InStr(fileName, g_PlantSection) > 0 And g_PlantSection <> "" Then
                    LogMessage "ANALYZE: ⚠️ SKIPPING (already heritage-renamed with same prefix) - " & fileName
                    LogMessage "         This file has already been processed by STEP 1 with prefix '" & g_PlantSection & "'. Skipping to prevent duplicate numbering."
                ' Process part file - check for uniqueness
                ElseIf Not uniqueParts.Exists(fullPath) Then
                    ' Mark as processed to prevent duplicates
                    uniqueParts.Add fullPath, True

                    ' Read Description from Design Tracking Properties
                    Dim description
                    description = GetDescriptionFromIProperty(doc)

                    If description = "" Then
                        LogMessage "ANALYZE: WARNING - No description found for " & fileName & " (in " & asmDoc.DisplayName & ")"
                    Else
                        ' Group by description using client's logic
                        Dim groupCode
                        groupCode = ClassifyByDescription(description)

                        If groupCode = "SKIP" Then
                            LogMessage "ANALYZE: SKIPPING " & fileName & " (hardware/bolts) - Description: " & description
                        Else
                            LogMessage "ANALYZE: PART - " & fileName & " -> Description: " & description & " -> Group: " & groupCode & " (from " & asmDoc.DisplayName & ")"

                            ' Add to component groups
                            If Not g_ComponentGroups.Exists(groupCode) Then
                                g_ComponentGroups.Add groupCode, CreateObject("Scripting.Dictionary")
                                LogMessage "ANALYZE: Created new group - " & groupCode
                            End If

                            ' Add this part to the group using full path as key to ensure uniqueness
                            Dim groupDict
                            Set groupDict = g_ComponentGroups.Item(groupCode)

                            ' Use full path as key to prevent filename collisions
                            If Not groupDict.Exists(fullPath) Then
                                groupDict.Add fullPath, fullPath & "|" & description & "|" & fileName
                            End If
                        End If
                    End If
                Else
                    LogMessage "ANALYZE: DUPLICATE PART SKIPPED - " & fileName & " (already processed)"
                End If ' Heritage-renamed / uniqueParts check

            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                ' This is a sub-assembly - recurse into it (unless it's a bolted connection)
                If InStr(LCase(fileName), "bolted connection") > 0 Then
                    LogMessage "ANALYZE: SKIPPING " & fileName & " (bolted connection assembly)"
                Else
                    LogMessage "ANALYZE: RECURSING into sub-assembly - " & fileName
                    Call ProcessAssemblyRecursively(doc, uniqueParts, asmLevel & ">" & fileName)
                End If
            Else
                LogMessage "ANALYZE: SKIPPING " & fileName & " (unknown file type)"
            End If
        End If
    Next
End Sub

Sub ShowGroupsSummary()
    LogMessage "SUMMARY: Showing component groups to user"

    If g_ComponentGroups.Count = 0 Then
        MsgBox "No component groups found!" & vbCrLf & vbCrLf & _
               "Make sure your assembly contains part files (.ipt) with Description properties", vbExclamation
        Exit Sub
    End If

    ' Build summary message
    Dim summaryMsg
    summaryMsg = "STEP 4: IDENTIFY QTY OF SECTIONS" & vbCrLf & vbCrLf
    summaryMsg = summaryMsg & "Found " & g_ComponentGroups.Count & " component groups:" & vbCrLf & vbCrLf

    Dim debugGroupKeys
    debugGroupKeys = g_ComponentGroups.Keys

    Dim debugI
    For debugI = 0 To UBound(debugGroupKeys)
        Dim debugGroupName
        debugGroupName = debugGroupKeys(debugI)

        Dim debugGroupDict
        Set debugGroupDict = g_ComponentGroups.Item(debugGroupName)

        ' Show group with client-friendly naming
        Dim groupDescription
        Select Case debugGroupName
            Case "B"
                groupDescription = "I and H sections (UB/UC)"
            Case "PL"
                groupDescription = "Platework (PL + S355JR)"
            Case "LPL"
                groupDescription = "Liners (PL + other materials)"
            Case "A"
                groupDescription = "Angles (L sections)"
            Case "CH"
                groupDescription = "Channels (PFC/TFC)"
            Case "P"
                groupDescription = "Pipes & Circular hollow (CHS/PIPE)"
            Case "FLG"
                groupDescription = "Flanges (FLANGE)"
            Case "R"
                groupDescription = "Roundbar (R with diameter)"
            Case "SQ"
                groupDescription = "Square/rectangular hollow (SHS)"
            Case "FL"
                groupDescription = "Flatbar (FL)"
            Case "IPE"
                groupDescription = "European I-beams (IPE)"
            Case Else
                groupDescription = "Other components"
        End Select

        summaryMsg = summaryMsg & "[" & debugGroupName & "] " & groupDescription & vbCrLf
        summaryMsg = summaryMsg & "   Quantity: " & debugGroupDict.Count & " components" & vbCrLf

        ' Show first few component descriptions as examples
        Dim debugCompKeys
        debugCompKeys = debugGroupDict.Keys

        Dim exampleCount
        exampleCount = 3 ' Show max 3 examples
        If UBound(debugCompKeys) + 1 < exampleCount Then
            exampleCount = UBound(debugCompKeys) + 1
        End If

        Dim debugJ
        For debugJ = 0 To exampleCount - 1
            Dim pathAndDescAndFile
            pathAndDescAndFile = debugGroupDict.Item(debugCompKeys(debugJ))
            Dim parts
            parts = Split(pathAndDescAndFile, "|")
            Dim description
            description = parts(1)
            Dim fileName
            fileName = parts(2)
            summaryMsg = summaryMsg & "   - " & fileName & " (" & description & ")" & vbCrLf
        Next

        If UBound(debugCompKeys) + 1 > exampleCount Then
            summaryMsg = summaryMsg & "   - (and " & (UBound(debugCompKeys) + 1 - exampleCount) & " more...)" & vbCrLf
        End If

        summaryMsg = summaryMsg & vbCrLf
    Next

    summaryMsg = summaryMsg & "Continue with renaming each group?"

    Dim result
    result = MsgBox(summaryMsg, vbYesNo + vbQuestion, "Component Groups - Step 4")

    If result = vbNo Then
        LogMessage "SUMMARY: User cancelled after reviewing groups"
        WScript.Quit
    End If

    LogMessage "SUMMARY: User approved component groups"
End Sub

Sub GetPlantSectionNaming()
    LogMessage "PLANT: Getting plant section naming convention from user"

    Dim plantInput
    plantInput = InputBox("STEP 3: DEFINE PREFIX" & vbCrLf & vbCrLf & _
                         "Enter the project prefix (as per drawing register):" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "  PLANT1-000-    (for Plant 1)" & vbCrLf & _
                         "  AREA2-000-     (for Area 2)" & vbCrLf & _
                         "  SEC-A-000-     (for Section A)" & vbCrLf & _
                         "  BLOCK3-000-    (for Block 3)" & vbCrLf & vbCrLf & _
                         "This will create part numbers like:" & vbCrLf & _
                         "  PLANT1-000-B1, PLANT1-000-PL1, PLANT1-000-CH1, etc." & vbCrLf & vbCrLf & _
                         "NOTE: Uses single digit numbering (1, 2, 3...)" & vbCrLf & vbCrLf & _
                         "REQUIRED: Enter project prefix (e.g., N1SCR04-730-)", _
                         "Define Project Prefix", "")

    If plantInput = "" Then
        MsgBox "ERROR: Project prefix is required!" & vbCrLf & vbCrLf & _
               "Please enter a valid project prefix.", vbCritical, "Input Required"
        LogMessage "ERROR: User did not provide required project prefix"
        Exit Sub
    Else
        ' Trim whitespace
        plantInput = Trim(plantInput)

        ' Validate prefix contains only safe characters (alphanumeric, dash, underscore)
        Dim isValid
        isValid = True
        Dim charCheck
        For charCheck = 1 To Len(plantInput)
            Dim char
            char = Mid(plantInput, charCheck, 1)
            If Not (char >= "A" And char <= "Z") And _
               Not (char >= "a" And char <= "z") And _
               Not (char >= "0" And char <= "9") And _
               char <> "-" And char <> "_" Then
                isValid = False
                Exit For
            End If
        Next

        If Not isValid Then
            MsgBox "Invalid prefix!" & vbCrLf & vbCrLf & _
                   "Prefix can only contain:" & vbCrLf & _
                   "• Letters (A-Z)" & vbCrLf & _
                   "• Numbers (0-9)" & vbCrLf & _
                   "• Dash (-)" & vbCrLf & _
                   "• Underscore (_)" & vbCrLf & vbCrLf & _
                   "Please restart and enter a valid prefix.", vbCritical
            LogMessage "ERROR: Invalid prefix characters: " & plantInput
            WScript.Quit
        End If

        ' Ensure it ends with a dash
        If Right(plantInput, 1) <> "-" Then
            plantInput = plantInput & "-"
        End If

        g_PlantSection = UCase(plantInput) ' Normalize to uppercase
        LogMessage "PLANT: Using custom naming convention: " & g_PlantSection
    End If

    ' Check if prefix exists in registry
    If CheckIfPrefixExistsInRegistry(g_PlantSection) Then
        Dim choice
        choice = MsgBox("Prefix '" & g_PlantSection & "' already exists in registry." & vbCrLf & vbCrLf & _
                        "Do you want to SKIP existing parts and continue numbering from current counters?" & vbCrLf & vbCrLf & _
                        "YES = Skip existing parts, continue numbering" & vbCrLf & _
                        "NO = Rename everything from 1", vbYesNo + vbQuestion, "Registry Check")
        g_SkipExisting = (choice = vbYes)
        LogMessage "REGISTRY: Prefix exists, user chose to " & IIf(g_SkipExisting, "skip existing", "rename from 1")
    Else
        g_SkipExisting = False
        LogMessage "REGISTRY: Prefix does not exist, starting fresh"
    End If

    ' Show confirmation
    MsgBox "PROJECT PREFIX SET" & vbCrLf & vbCrLf & _
           "Your project prefix: " & g_PlantSection & vbCrLf & vbCrLf & _
           "Example part numbers will be:" & vbCrLf & _
           "  " & g_PlantSection & "B1 (I/H sections)" & vbCrLf & _
           "  " & g_PlantSection & "PL1 (Platework)" & vbCrLf & _
           "  " & g_PlantSection & "CH1 (Channels)" & vbCrLf & _
           "  " & g_PlantSection & "A1 (Angles)" & vbCrLf & vbCrLf & _
           "Continue to component analysis...", vbInformation, "Prefix Set"
End Sub

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function GetDescriptionFromIProperty(doc)
    ' Read Description from Design Tracking Properties
    On Error Resume Next

    Dim propertySet
    Set propertySet = doc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then
        LogMessage "WARNING: Cannot access Design Tracking Properties for " & doc.DisplayName
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    Dim descriptionProp
    Set descriptionProp = propertySet.Item("Description")

    If Err.Number <> 0 Then
        LogMessage "WARNING: No Description property found for " & doc.DisplayName
        GetDescriptionFromIProperty = ""
        Err.Clear
        Exit Function
    End If

    GetDescriptionFromIProperty = Trim(descriptionProp.Value)
    Err.Clear
End Function

Function ClassifyByDescription(description)
    ' Classify components based on Description iProperty using client's exact requirements

    Dim desc
    desc = UCase(Trim(description))

    ' Skip hardware and bolts first
    If InStr(desc, "BOLT") > 0 Or InStr(desc, "SCREW") > 0 Or InStr(desc, "WASHER") > 0 Or InStr(desc, "NUT") > 0 Then
        ClassifyByDescription = "SKIP"
        Exit Function
    End If

    ' NEW: Check for FLANGE in description (description-only as requested)
    If InStr(desc, "FLANGE") > 0 Then
        ClassifyByDescription = "FLG"  ' Flanges
        Exit Function
    End If

    ' NEW: Check for PIPE
    If InStr(desc, "PIPE") > 0 Then
        ClassifyByDescription = "P"  ' Pipes
        Exit Function
    End If

    ' NEW: Check for Roundbar R followed by digits
    If Len(desc) >= 2 And Left(desc, 1) = "R" Then
        Dim secondChar
        secondChar = Mid(desc, 2, 1)
        If IsNumeric(secondChar) Then
            ClassifyByDescription = "R"  ' Roundbar
            Exit Function
        End If
    End If

    ' Client's grouping logic - exact requirements
    If Left(desc, 2) = "UB" Then
        ClassifyByDescription = "B"  ' I and H sections - UB beams
    ElseIf Left(desc, 2) = "UC" Then
        ClassifyByDescription = "B"  ' I and H sections - UC columns
    ElseIf Left(desc, 2) = "PL" Then
        ' Check if it's platework (PL + S355JR) or liners (PL + NOT S355JR)
        If InStr(desc, "S355JR") > 0 Then
            ClassifyByDescription = "PL"  ' Platework
        Else
            ClassifyByDescription = "LPL" ' Liners
        End If
    ElseIf Left(desc, 1) = "L" And (InStr(desc, "X") > 0 Or InStr(desc, " X ") > 0) Then
        ClassifyByDescription = "A"   ' Angles - L50x50x6, L70 x 70 x 6 etc.
    ElseIf Left(desc, 3) = "PFC" Then
        ClassifyByDescription = "CH"  ' Parallel flange channels
    ElseIf Left(desc, 3) = "TFC" Then
        ClassifyByDescription = "CH"  ' Taper flange channels
    ElseIf Left(desc, 3) = "CHS" Then
        ClassifyByDescription = "P"   ' Circular hollow sections
    ElseIf Left(desc, 3) = "SHS" Then
        ClassifyByDescription = "SQ"  ' Square/rectangular hollow sections
    ElseIf Left(desc, 2) = "FL" And Not InStr(desc, "FLOOR") > 0 Then
        ClassifyByDescription = "FL"  ' Flatbar (but not floor grating)
    ElseIf Left(desc, 3) = "IPE" Then
        ClassifyByDescription = "IPE"  ' European I-beams (separate group)
    Else
        ' Default - unclassified part
        ClassifyByDescription = "OTHER"
    End If
End Function

Function RemoveLeadingNumbers(inputName)
    ' Remove numbers at the start of words: "Part1 TFC" -> "Part TFC"
    Dim result
    result = inputName
    
    ' Split by spaces and process each word
    Dim words
    words = Split(result, " ")
    
    Dim i
    For i = 0 To UBound(words)
        Dim word
        word = words(i)
        
        ' If word starts with letters followed by numbers, keep only letters
        If Len(word) > 1 Then
            Dim j
            For j = 1 To Len(word)
                If IsNumeric(Mid(word, j, 1)) Then
                    words(i) = Left(word, j - 1)
                    Exit For
                End If
            Next
        End If
    Next
    
    RemoveLeadingNumbers = Join(words, " ")
End Function

Function RemoveNumbersFromName(inputName)
    ' Fallback method - remove all standalone numbers
    Dim result
    result = inputName
    
    ' Remove common number patterns
    result = Replace(result, "1 ", " ")
    result = Replace(result, "2 ", " ")
    result = Replace(result, "3 ", " ")
    result = Replace(result, "4 ", " ")
    result = Replace(result, "5 ", " ")
    result = Replace(result, "6 ", " ")
    result = Replace(result, "7 ", " ")
    result = Replace(result, "8 ", " ")
    result = Replace(result, "9 ", " ")
    result = Replace(result, "0 ", " ")
    
    ' Clean up extra spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    RemoveNumbersFromName = Trim(result)
End Function

Sub GetUserNamingSchemes()
    LogMessage "INPUT: Getting user naming schemes for " & g_ComponentGroups.Count & " groups"

    ' Show user the groups and get naming schemes
    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys

    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)

        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)

        ' Build component list with descriptions
        Dim componentList
        componentList = ""
        Dim compKeys
        compKeys = groupDict.Keys

        Dim j
        For j = 0 To UBound(compKeys)
            Dim pathAndDescAndFile
            pathAndDescAndFile = groupDict.Item(compKeys(j))
            Dim parts
            parts = Split(pathAndDescAndFile, "|")
            Dim description
            description = parts(1)
            Dim fileName
            fileName = parts(2)
            componentList = componentList & "  • " & fileName & " (" & description & ")" & vbCrLf
        Next

        ' Generate default scheme based on client's group codes
        Dim defaultScheme
        Select Case groupName
            Case "B"
                defaultScheme = g_PlantSection & "B{N}"
            Case "PL"
                defaultScheme = g_PlantSection & "PL{N}"
            Case "LPL"
                defaultScheme = g_PlantSection & "LPL{N}"
            Case "A"
                defaultScheme = g_PlantSection & "A{N}"
            Case "CH"
                defaultScheme = g_PlantSection & "CH{N}"
            Case "P"
                defaultScheme = g_PlantSection & "P{N}"
            Case "FLG"
                defaultScheme = g_PlantSection & "FLG{N}"
            Case "R"
                defaultScheme = g_PlantSection & "R{N}"
            Case "SQ"
                defaultScheme = g_PlantSection & "SQ{N}"
            Case "FL"
                defaultScheme = g_PlantSection & "FL{N}"
            Case "IPE"
                defaultScheme = g_PlantSection & "IPE{N}"
            Case Else
                defaultScheme = g_PlantSection & "PART{N}"
        End Select

        Dim userInput
        userInput = InputBox("COMPONENT GROUP: " & groupName & " (" & groupDict.Count & " components)" & vbCrLf & vbCrLf & _
                            "Plant Section: " & g_PlantSection & vbCrLf & vbCrLf & _
                            "Components in this group:" & vbCrLf & componentList & vbCrLf & _
                            "Enter naming scheme:" & vbCrLf & _
                            "IMPORTANT: Use {N} for auto-numbering!" & vbCrLf & vbCrLf & _
                            "Examples with your plant section:" & vbCrLf & _
                            "  " & g_PlantSection & "B{N}   -> " & g_PlantSection & "B1, " & g_PlantSection & "B2..." & vbCrLf & _
                            "  " & g_PlantSection & "PL{N}  -> " & g_PlantSection & "PL1, " & g_PlantSection & "PL2..." & vbCrLf & _
                            "  " & g_PlantSection & "CH{N}  -> " & g_PlantSection & "CH1, " & g_PlantSection & "CH2..." & vbCrLf & vbCrLf & _
                            "WITHOUT {N}, all parts get the SAME name!", _
                            "Naming Scheme for Group: " & groupName, defaultScheme)

        If userInput = "" Then
            LogMessage "INPUT: User cancelled - using default scheme: " & defaultScheme
            userInput = defaultScheme
        End If

        g_NamingSchemes.Add groupName, userInput
        LogMessage "INPUT: Group '" & groupName & "' -> Scheme: " & userInput
    Next
End Sub

Sub CreateDynamicHeritageBasedCopies(invApp, asmDoc)
    LogMessage "HERITAGE: Creating dynamic heritage-based copies"

    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys

    Dim asmDir
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    asmDir = fso.GetParentFolderName(asmDoc.FullFileName) & "\"

    ' Scan mapping file for existing heritage files to continue numbering
    LogMessage "HERITAGE: Scanning mapping file for existing heritage files to continue numbering"
    Dim existingCounters
    Set existingCounters = CreateObject("Scripting.Dictionary")
    If g_SkipExisting Then
        Call ScanRegistryForCounters(existingCounters, g_PlantSection)
    Else
        LogMessage "HERITAGE: Starting fresh - not loading existing counters"
    End If

    ' Process each group
    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)

        Dim namingScheme
        namingScheme = g_NamingSchemes.Item(groupName)

        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)

        ' Get starting counter for this group with this prefix
        Dim prefixGroupKey
        prefixGroupKey = ExtractPrefixFromScheme(namingScheme) & groupName
        LogMessage "HERITAGE: Checking for existing counter with key: '" & prefixGroupKey & "'"

        Dim startingCounter
        If existingCounters.Exists(prefixGroupKey) Then
            Dim highestExisting
            highestExisting = existingCounters.Item(prefixGroupKey)
            startingCounter = highestExisting + 1
            LogMessage "HERITAGE: Group '" & groupName & "' continuing from number " & startingCounter & " (found existing files with highest number " & highestExisting & ")"
        Else
            startingCounter = 1
            LogMessage "HERITAGE: Group '" & groupName & "' starting from number 1 (new prefix or group - key '" & prefixGroupKey & "' not found)"

            ' Debug - show all existing keys
            If existingCounters.Count > 0 Then
                LogMessage "HERITAGE: DEBUG - Existing counter keys:"
                Dim debugKeys
                debugKeys = existingCounters.Keys
                Dim debugI
                For debugI = 0 To UBound(debugKeys)
                    LogMessage "HERITAGE: DEBUG -   '" & debugKeys(debugI) & "' = " & existingCounters.Item(debugKeys(debugI))
                Next
            Else
                LogMessage "HERITAGE: DEBUG - No existing counters found"
            End If
        End If

        LogMessage "HERITAGE: Processing group '" & groupName & "' with scheme: " & namingScheme
        
        ' Create heritage copies for each component in group
        Dim compKeys
        compKeys = groupDict.Keys

        Dim j
        Dim componentCounter
        componentCounter = startingCounter
        
        For j = 0 To UBound(compKeys)
            Dim pathAndDescAndFile
            pathAndDescAndFile = groupDict.Item(compKeys(j))
            Dim parts
            parts = Split(pathAndDescAndFile, "|")
            Dim originalPath
            originalPath = parts(0)
            Dim fileName
            fileName = parts(2)

            ' Generate new filename using scheme with proper counter
            Dim newFileName
            newFileName = GenerateNewFileName(namingScheme, componentCounter)
            componentCounter = componentCounter + 1

            ' Create heritage file in same directory as original part
            Dim fso2
            Set fso2 = CreateObject("Scripting.FileSystemObject")
            Dim originalDir
            originalDir = fso2.GetParentFolderName(originalPath) & "\"

            Dim newPath
            newPath = originalDir & newFileName

            LogMessage "HERITAGE: " & fileName & " -> " & newFileName
            
            ' Always store the mapping (regardless of file existence)
            g_FileNameMapping.Add originalPath, newFileName

            ' Store comprehensive mapping: originalPath -> "newPath|originalFile|newFile|group|description"
            Dim pathParts
            pathParts = Split(pathAndDescAndFile, "|")
            Dim description
            description = pathParts(1)
            Dim originalFileName
            originalFileName = pathParts(2)

            Dim mappingValue
            mappingValue = newPath & "|" & originalFileName & "|" & newFileName & "|" & groupName & "|" & description
            g_ComprehensiveMapping.Add originalPath, mappingValue

            LogMessage "MAPPING: " & originalPath & " -> " & newPath

            ' Check if file already exists (safety check)
            Dim heritageFileSystem
            Set heritageFileSystem = CreateObject("Scripting.FileSystemObject")
            If heritageFileSystem.FileExists(newPath) Then
                LogMessage "HERITAGE: File already exists: " & newFileName & " (mapping still recorded)"
            Else
                ' Create heritage file
                LogMessage "HERITAGE: Creating new file: " & newFileName

                ' Open document and create heritage copy
                On Error Resume Next
                Dim partDoc
                Set partDoc = invApp.Documents.Open(originalPath, False)

                If Err.Number = 0 Then
                    partDoc.SaveAs newPath, True
                    If Err.Number = 0 Then
                        LogMessage "HERITAGE: SUCCESS - Created " & newFileName
                    Else
                        LogMessage "HERITAGE: ERROR - SaveAs failed for " & newFileName & ": " & Err.Description
                    End If
                    partDoc.Close
                Else
                    LogMessage "HERITAGE: ERROR - Could not open " & originalPath & ": " & Err.Description
                End If

                Err.Clear
            End If
        Next

        ' Save final counter to Registry for this group
        Dim finalCounter
        finalCounter = componentCounter - 1  ' Last used number
        Call SaveCounterToRegistry(prefixGroupKey, finalCounter)
    Next
End Sub

Function GenerateNewFileName(scheme, number)
    Dim result
    result = scheme

    ' Replace {N} with simple number (no padding)
    result = Replace(result, "{N}", CStr(number))

    ' Add .ipt extension if not present
    If Right(LCase(result), 4) <> ".ipt" Then
        result = result & ".ipt"
    End If

    GenerateNewFileName = result
End Function

Sub UpdateDynamicAssemblyReferences(invApp, asmDoc)
    LogMessage "ASSEMBLY: Recursively updating assembly references across entire model hierarchy"

    ' Start recursive reference updating from root assembly
    Call UpdateAssemblyReferencesRecursively(asmDoc, "ROOT")

    ' Save main assembly with error handling
    LogMessage "ASSEMBLY: Saving assembly with updated references..."
    On Error Resume Next
    asmDoc.Save
    
    If Err.Number = 0 Then
        LogMessage "ASSEMBLY: Successfully saved assembly"
    Else
        LogMessage "ASSEMBLY: ERROR saving assembly: " & Err.Description
        MsgBox "WARNING: Assembly save failed!" & vbCrLf & _
               "Error: " & Err.Description & vbCrLf & vbCrLf & _
               "You may need to save manually.", vbExclamation
    End If
    Err.Clear
End Sub

Sub UpdateAssemblyReferencesRecursively(asmDoc, asmLevel)
    LogMessage "ASSEMBLY: Updating references in - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")"

    ' Get assembly directory for new file paths
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim asmDir
    asmDir = fso.GetParentFolderName(asmDoc.FullFileName) & "\"

    ' Process each occurrence in this assembly
    Dim i
    For i = 1 To asmDoc.ComponentDefinition.Occurrences.Count
        Dim occ
        Set occ = asmDoc.ComponentDefinition.Occurrences.Item(i)

        ' SKIP suppressed occurrences
        If occ.Suppressed Then
            LogMessage "ASSEMBLY: SKIPPING (suppressed occurrence in " & asmDoc.DisplayName & ")"
        Else
            Dim doc
            Set doc = occ.Definition.Document

            Dim fullPath
            fullPath = doc.FullFileName

            Dim fileName
            fileName = GetFileNameFromPath(fullPath)

            ' Check if this is a part file that needs updating
            If LCase(Right(fileName, 4)) = ".ipt" Then
                ' Use comprehensive mapping to find heritage file
                If g_ComprehensiveMapping.Exists(fullPath) Then
                    Dim mappingValue
                    mappingValue = g_ComprehensiveMapping.Item(fullPath)

                    ' Parse mapping: "newPath|originalFile|newFile|group|description"
                    Dim mappingParts
                    mappingParts = Split(mappingValue, "|")

                    Dim newPath
                    newPath = mappingParts(0)
                    Dim originalFileName
                    originalFileName = mappingParts(1)
                    Dim newFileName
                    newFileName = mappingParts(2)
                    Dim groupName
                    groupName = mappingParts(3)

                    LogMessage "ASSEMBLY: Replacing " & originalFileName & " -> " & newFileName & " [" & groupName & "] (in " & asmDoc.DisplayName & ")"

                    On Error Resume Next
                    occ.Replace newPath, True

                    If Err.Number = 0 Then
                        LogMessage "ASSEMBLY: SUCCESS - Updated to " & newFileName
                    Else
                        LogMessage "ASSEMBLY: ERROR - Replace failed: " & Err.Description
                    End If
                    Err.Clear
                Else
                    LogMessage "ASSEMBLY: INFO - No mapping found for " & fileName & " (not renamed)"
                End If

            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                ' This is a sub-assembly - recurse into it (unless it's a bolted connection)
                If InStr(LCase(fileName), "bolted connection") > 0 Then
                    LogMessage "ASSEMBLY: SKIPPING " & fileName & " (bolted connection assembly)"
                Else
                    LogMessage "ASSEMBLY: RECURSING into sub-assembly - " & fileName
                    Call UpdateAssemblyReferencesRecursively(doc, asmLevel & ">" & fileName)

                    ' Save the sub-assembly after updating its references
                    On Error Resume Next
                    doc.Save
                    If Err.Number = 0 Then
                        LogMessage "ASSEMBLY: Saved sub-assembly - " & fileName
                    Else
                        LogMessage "ASSEMBLY: ERROR saving sub-assembly " & fileName & ": " & Err.Description
                    End If
                    Err.Clear
                End If
            End If
        End If
    Next
End Sub

Function FindHeritageFileForOriginal(originalFullPath)
    ' Find the heritage file created for the original file
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim originalDir
    originalDir = fso.GetParentFolderName(originalFullPath) & "\"

    ' Search through all component groups to find this original path
    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys

    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)

        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)

        ' Check if this original path exists in this group
        If groupDict.Exists(originalFullPath) Then
            ' Found the group - now find the index and generate heritage file path
            Dim compKeys
            compKeys = groupDict.Keys

            Dim j
            Dim componentIndex
            componentIndex = 1

            For j = 0 To UBound(compKeys)
                If compKeys(j) = originalFullPath Then
                    ' Found the component - generate its heritage name
                    Dim namingScheme
                    namingScheme = g_NamingSchemes.Item(groupName)

                    Dim heritageFileName
                    heritageFileName = GenerateNewFileName(namingScheme, componentIndex)

                    Dim heritagePath
                    heritagePath = originalDir & heritageFileName

                    ' Check if heritage file exists
                    If fso.FileExists(heritagePath) Then
                        FindHeritageFileForOriginal = heritagePath
                        Exit Function
                    End If
                End If
                componentIndex = componentIndex + 1
            Next
        End If
    Next

    FindHeritageFileForOriginal = "" ' Not found
End Function

' ❌ DELETED: FindHeritagePathByOriginalFilename() - Obsolete function removed (October 1, 2025)
' This function caused the filename-only matching bug. Now using direct dictionary lookup with full paths.

Function FindNewPathForComponent(originalFileName, asmDir)
    ' Find which group this component belongs to and return new path
    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys
    
    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)
        
        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)
        
        If groupDict.Exists(originalFileName) Then
            ' Found the group - now find the index and generate new name
            Dim compKeys
            compKeys = groupDict.Keys
            
            Dim j
            Dim componentIndex
            componentIndex = 1
            
            For j = 0 To UBound(compKeys)
                If compKeys(j) = originalFileName Then
                    Dim namingScheme
                    namingScheme = g_NamingSchemes.Item(groupName)
                    
                    Dim newFileName
                    newFileName = GenerateNewFileName(namingScheme, componentIndex)
                    
                    FindNewPathForComponent = asmDir & newFileName
                    Exit Function
                End If
                componentIndex = componentIndex + 1
            Next
        End If
    Next
    
    FindNewPathForComponent = "" ' Not found
End Function

Sub UpdateAllIDWsInDirectory(invApp, asmDoc)
    LogMessage "IDW: Auto-detecting IDW files in assembly directory"
    
    ' Get assembly directory
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim asmDir
    asmDir = fso.GetParentFolderName(asmDoc.FullFileName)
    
    ' Find all IDW files in the directory
    Dim folder
    Set folder = fso.GetFolder(asmDir)
    
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            idwFiles.Add file.Name, file.Path
            LogMessage "IDW: Found drawing file - " & file.Name
        End If
    Next
    
    If idwFiles.Count = 0 Then
        LogMessage "IDW: No IDW files found in directory"
        Exit Sub
    End If
    
    ' Update each IDW file
    Dim idwKeys
    idwKeys = idwFiles.Keys
    
    Dim i
    For i = 0 To UBound(idwKeys)
        Dim idwPath
        idwPath = idwFiles.Item(idwKeys(i))
        
        LogMessage "IDW: Processing " & idwKeys(i)
        Call UpdateSingleIDWWithDynamicReferences(invApp, idwPath, asmDir)
    Next
End Sub

Sub UpdateSingleIDWWithDynamicReferences(invApp, idwPath, asmDir)
    On Error Resume Next
    
    ' Close all documents first
    invApp.Documents.CloseAll
    
    LogMessage "IDW: Opening " & GetFileNameFromPath(idwPath)
    
    Dim idwDoc
    Set idwDoc = invApp.Documents.Open(idwPath, False)
    
    If Err.Number <> 0 Then
        LogMessage "IDW: ERROR - Could not open: " & Err.Description
        Exit Sub
    End If
    
    ' Access file descriptors
    Dim fileDescriptors
    Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
    
    LogMessage "IDW: Found " & fileDescriptors.Count & " referenced files"
    
    ' Update each reference dynamically
    Dim i
    For i = 1 To fileDescriptors.Count
        Dim fd
        Set fd = fileDescriptors.Item(i)

        ' ✅ FIX: Use FULL PATH from IDW reference (not just filename!)
        Dim currentFullPath
        currentFullPath = fd.FullFileName

        Dim currentFileName
        currentFileName = GetFileNameFromPath(currentFullPath)

        ' Direct dictionary lookup using full path as key (same as STEP 2 fix)
        Dim newPath
        newPath = ""

        If g_ComprehensiveMapping.Exists(currentFullPath) Then
            Dim mappingValue
            mappingValue = g_ComprehensiveMapping.Item(currentFullPath)

            ' Parse mapping: originalPath|newPath|originalFile|newFile|group|description
            ' BUT stored as: newPath|originalFile|newFile|group|description (in dictionary value)
            Dim mappingParts
            mappingParts = Split(mappingValue, "|")

            If UBound(mappingParts) >= 0 Then
                newPath = mappingParts(0) ' newPath is field #0 in dictionary value
            End If
        End If

        If newPath <> "" Then
            Dim newFileName
            newFileName = GetFileNameFromPath(newPath)
            LogMessage "IDW: Updating reference " & currentFileName & " -> " & newFileName

            fd.ReplaceReference newPath

            If Err.Number = 0 Then
                LogMessage "IDW: SUCCESS - Reference updated"
            Else
                LogMessage "IDW: ERROR - ReplaceReference failed: " & Err.Description
                Err.Clear
            End If
        Else
            LogMessage "IDW: INFO - No mapping found for " & currentFileName
        End If
    Next
    
    ' Save IDW
    idwDoc.Save
    LogMessage "IDW: Saved " & GetFileNameFromPath(idwPath)
    
    Err.Clear
End Sub

Sub OptionalCleanupOldFiles(asmDoc)
    Dim result
    result = MsgBox("CLEANUP ORIGINAL FILES?" & vbCrLf & vbCrLf & _
                    "Do you want to delete the original files?" & vbCrLf & _
                    "This cannot be undone!" & vbCrLf & vbCrLf & _
                    "✅ Heritage copies created successfully" & vbCrLf & _
                    "✅ Assembly and IDW files updated" & vbCrLf & vbCrLf & _
                    "Delete originals?", vbYesNo + vbQuestion, "Cleanup Original Files")
    
    If result = vbNo Then
        LogMessage "CLEANUP: User chose to keep original files"
        Exit Sub
    End If
    
    LogMessage "CLEANUP: User chose to delete original files"
    
    ' Delete original files from all groups
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys
    
    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)
        
        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)
        
        Dim compKeys
        compKeys = groupDict.Keys
        
        Dim j
        For j = 0 To UBound(compKeys)
            Dim pathAndDescAndFile
            pathAndDescAndFile = groupDict.Item(compKeys(j))
            Dim parts
            parts = Split(pathAndDescAndFile, "|")
            Dim originalPath
            originalPath = parts(0)
            Dim fileName
            fileName = parts(2)

            If fso.FileExists(originalPath) Then
                On Error Resume Next
                fso.DeleteFile originalPath
                
                If Err.Number = 0 Then
                    LogMessage "CLEANUP: Deleted " & fileName
                Else
                    LogMessage "CLEANUP: ERROR - Could not delete " & fileName & ": " & Err.Description
                End If
                Err.Clear
            End If
        Next
    Next
    
    LogMessage "CLEANUP: Original file cleanup completed"
End Sub

Sub CreateAuditLog(invApp, activeDoc)
    On Error GoTo 0
    On Error Resume Next

    LogMessage "AUDIT: === START AUDIT LOG CREATION ==="
    LogMessage "AUDIT: Capturing assembly state BEFORE any renaming"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Extract assembly directory
    Dim asmDir
    asmDir = fso.GetParentFolderName(activeDoc.FullFileName)
    Dim auditFilePath
    auditFilePath = asmDir & "\STEP_0_AUDIT.txt"

    LogMessage "AUDIT: Assembly directory: " & asmDir
    LogMessage "AUDIT: Audit file path: " & auditFilePath

    ' Create audit file
    Dim auditFile
    Set auditFile = fso.CreateTextFile(auditFilePath, True)

    ' Header
    auditFile.WriteLine "================================================================"
    auditFile.WriteLine "STEP 0: ASSEMBLY AUDIT LOG - BEFORE STATE"
    auditFile.WriteLine "================================================================"
    auditFile.WriteLine "Created: " & Now()
    auditFile.WriteLine "Assembly: " & activeDoc.DisplayName
    auditFile.WriteLine "Full Path: " & activeDoc.FullFileName
    auditFile.WriteLine "================================================================"
    auditFile.WriteLine ""

    ' Capture component list with occurrence numbers
    auditFile.WriteLine "COMPONENT LIST (with occurrence numbers):"
    auditFile.WriteLine "================================================================"

    Dim asmCompDef
    Set asmCompDef = activeDoc.ComponentDefinition

    Dim occ
    For Each occ In asmCompDef.Occurrences
        Dim compName
        compName = occ.Name

        Dim occNumber
        occNumber = ""

        ' Check if this component has occurrence numbering
        If InStr(compName, ":") > 0 Then
            occNumber = Mid(compName, InStr(compName, ":"))
            compName = Left(compName, InStr(compName, ":") - 1)
        End If

        Dim doc
        Set doc = occ.Definition.Document

        auditFile.WriteLine compName & occNumber & " | " & doc.DisplayName & " | " & doc.FullFileName
    Next

    auditFile.WriteLine ""
    auditFile.WriteLine "================================================================"
    auditFile.WriteLine ""

    ' Capture iLogic rules and their component references
    auditFile.WriteLine "ILOGIC RULES (and component name references):"
    auditFile.WriteLine "================================================================"

    ' Connect to iLogic Add-in using proven pattern
    Dim addins
    Set addins = invApp.ApplicationAddIns

    Dim iLogicAddin
    Set iLogicAddin = addins.ItemById(ILOGIC_ADDIN_GUID)

    If Err.Number <> 0 Or iLogicAddin Is Nothing Then
        auditFile.WriteLine "ERROR: Could not connect to iLogic Add-in"
        auditFile.WriteLine "  Err.Number: " & Err.Number
        auditFile.WriteLine "  Err.Description: " & Err.Description
        auditFile.WriteLine "  iLogicAddin Is Nothing: " & (iLogicAddin Is Nothing)
        auditFile.WriteLine ""
        auditFile.WriteLine "ATTEMPTING FALLBACK - Searching add-ins by name..."

        ' Try fallback method - search by name
        Dim addin, addinName
        For Each addin In addins
            On Error Resume Next
            addinName = addin.DisplayName
            If Err.Number = 0 Then
                If InStr(1, addinName, "iLogic", vbTextCompare) > 0 Then
                    auditFile.WriteLine "  FOUND BY NAME: " & addinName
                    auditFile.WriteLine "  GUID: " & addin.ClassId
                    Set iLogicAddin = addin
                    Exit For
                End If
            End If
            Err.Clear
        Next

        If iLogicAddin Is Nothing Then
            auditFile.WriteLine "  iLogic add-in not found - assembly may not have iLogic rules"
            LogMessage "AUDIT: ERROR - Could not connect to iLogic Add-in: " & Err.Number & " - " & Err.Description
            auditFile.WriteLine ""
            auditFile.WriteLine "================================================================"
            auditFile.WriteLine "ASSEMBLY MAY NOT HAVE ILOGIC RULES OR ADD-IN NOT INSTALLED"
            auditFile.WriteLine "================================================================"
            auditFile.Close
            Exit Sub
        Else
            auditFile.WriteLine "  Using fallback connection method..."
        End If
    Else
        auditFile.WriteLine "SUCCESS: Connected to iLogic add-in via GUID"
        auditFile.WriteLine "  Add-in name: " & iLogicAddin.DisplayName
    End If

    ' Activate the add-in and get automation interface
    iLogicAddin.Activate
    Dim iLogicAuto
    Set iLogicAuto = iLogicAddin.Automation

    ' Get iLogic rules from assembly
        ' Get iLogic rules from assembly
    On Error Resume Next
    Dim ruleNames
    Set ruleNames = iLogicAuto.Rules(activeDoc)

    If Err.Number <> 0 Then
        auditFile.WriteLine "ERROR: Could not get rules from assembly - " & Err.Description
        LogMessage "AUDIT: ERROR - Could not get rules: " & Err.Description
        Err.Clear
    Else
        auditFile.WriteLine "ASSEMBLY: " & activeDoc.DisplayName & " (" & ruleNames.Count & " rules)"
        auditFile.WriteLine "----------------------------------------------------------------"

        Dim ruleName
        For Each ruleName In ruleNames
            Dim ruleText
            ruleText = iLogicAuto.RuleText(activeDoc, ruleName)
            auditFile.WriteLine ""
            auditFile.WriteLine "RULE: " & ruleName
            auditFile.WriteLine "----------------------------------------------------------------"

            ' Extract component names from rule text
            ' Look for patterns like "ComponentName:1" or "ComponentName:"
            Dim compRefPattern
            compRefPattern = "[A-Za-z0-9_ \-]+:[0-9]+"

            Dim regex
            Set regex = CreateObject("VBScript.RegExp")
            regex.Global = True
            regex.IgnoreCase = False
            regex.Pattern = compRefPattern

                Dim matches
            Set matches = regex.Execute(ruleText)

            If matches.Count > 0 Then
                auditFile.WriteLine "COMPONENT REFERENCES FOUND:"
                Dim match
                For Each match In matches
                    auditFile.WriteLine "  --> " & match.Value
                Next
            Else
                auditFile.WriteLine "No component name references detected"
            End If

            auditFile.WriteLine "----------------------------------------------------------------"
        Next
    End If

    ' Get iLogic rules from referenced parts
    auditFile.WriteLine ""
    auditFile.WriteLine "REFERENCED PARTS:"
    auditFile.WriteLine "================================================================"

    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    For Each occ In asmCompDef.Occurrences
        Set doc = occ.Definition.Document
        If doc.DocumentType = kPartDocumentObject Then
            Dim uniqueKey
            uniqueKey = doc.FullFileName
            If Not uniqueParts.Exists(uniqueKey) Then
                uniqueParts.Add uniqueKey, doc.DisplayName
            End If
        End If
    Next

    Dim partKey
    For Each partKey In uniqueParts.Keys
        Dim partPath
        partPath = CStr(partKey)
        Dim partDoc
        On Error Resume Next
        Set partDoc = invApp.Documents.Open(partPath, False)
        If Err.Number = 0 And Not partDoc Is Nothing Then
            Dim partRuleNames
            Set partRuleNames = iLogicAuto.Rules(partDoc)

            If partRuleNames.Count > 0 Then
                auditFile.WriteLine ""
                auditFile.WriteLine "PART: " & uniqueParts(partKey) & " (" & partRuleNames.Count & " rules)"
                auditFile.WriteLine "----------------------------------------------------------------"

                Dim partRuleName
                For Each partRuleName In partRuleNames
                    Dim partRuleText
                    partRuleText = iLogicAuto.RuleText(partDoc, partRuleName)
                    auditFile.WriteLine ""
                    auditFile.WriteLine "RULE: " & partRuleName
                    auditFile.WriteLine "----------------------------------------------------------------"

                    ' Extract component names from part rule
                    Dim partMatches
                    Set partMatches = regex.Execute(partRuleText)

                    If partMatches.Count > 0 Then
                        auditFile.WriteLine "COMPONENT REFERENCES FOUND:"
                        Dim partMatch
                        For Each partMatch In partMatches
                            auditFile.WriteLine "  --> " & partMatch.Value
                        Next
                    Else
                        auditFile.WriteLine "No component name references detected"
                    End If

                    auditFile.WriteLine "----------------------------------------------------------------"
                Next
            End If

            partDoc.Close True
        End If
            Err.Clear
        Next
    End If

    auditFile.WriteLine ""
    auditFile.WriteLine "================================================================"
    auditFile.WriteLine "AUDIT LOG COMPLETED"
    auditFile.WriteLine "================================================================"

    auditFile.Close

    LogMessage "AUDIT: Audit log created successfully: " & auditFilePath
    LogMessage "AUDIT: === END AUDIT LOG CREATION ==="

    Err.Clear
    On Error GoTo 0
End Sub

Sub SaveMappingFile(activeDoc)
    On Error GoTo 0 ' Clear any previous error handling
    On Error Resume Next ' Enable error handling
    
    LogMessage "MAPPING: === START SAVE MAPPING FILE ==="
    LogMessage "MAPPING: Saving comprehensive mapping file for STEP 2"
    
    ' Check for errors during setup
    If Err.Number <> 0 Then
        LogMessage "MAPPING: SETUP ERROR - " & Err.Description & " (Error " & Err.Number & ")"
        Exit Sub
    End If

    ' Log input validation
    If activeDoc Is Nothing Then
        LogMessage "MAPPING: ERROR - activeDoc is Nothing!"
        Exit Sub
    End If
    LogMessage "MAPPING: Active document: " & activeDoc.DisplayName

    ' Log mapping dictionary count
    LogMessage "MAPPING: Comprehensive mapping count = " & g_ComprehensiveMapping.Count
    If g_ComprehensiveMapping.Count = 0 Then
        LogMessage "MAPPING: WARNING - No mappings to save!"
    End If

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Save in the assembly's directory for safety
    Dim asmDir
    ' Extract directory from first mapping key instead of activeDoc (which points to project IPJ)
    Dim mappingKeys
    mappingKeys = g_ComprehensiveMapping.Keys
    If UBound(mappingKeys) >=0 Then
        Dim firstOriginalPath
        firstOriginalPath = mappingKeys(0)
        asmDir = fso.GetParentFolderName(firstOriginalPath)
        LogMessage "MAPPING: Extracted directory from first mapping key: " & firstOriginalPath
    Else
        ' Fallback if mapping is empty (shouldn't happen in normal workflow)
        asmDir = fso.GetParentFolderName(activeDoc.FullFileName)
        LogMessage "MAPPING: WARNING - Using activeDoc fallback (mapping empty!)"
    End If
    Dim mappingFilePath
    mappingFilePath = asmDir & "\STEP_1_MAPPING.txt"

    LogMessage "MAPPING: Assembly directory: " & asmDir
    LogMessage "MAPPING: Full mapping file path: " & mappingFilePath

    ' Check directory exists
    If Not fso.FolderExists(asmDir) Then
        LogMessage "MAPPING: ERROR - Assembly directory does not exist: " & asmDir
        Exit Sub
    End If
    LogMessage "MAPPING: Assembly directory exists"

    Dim mappingFile
    Dim fileExists
    fileExists = fso.FileExists(mappingFilePath)
    LogMessage "MAPPING: Mapping file exists? " & fileExists

    On Error Resume Next

    If fileExists Then
        ' Append to existing file
        LogMessage "MAPPING: Opening existing file for append..."
        Set mappingFile = fso.OpenTextFile(mappingFilePath, 8) ' 8 = ForAppending
        If Err.Number <>0 Then
            LogMessage "MAPPING: ERROR - Could not open existing file: " & Err.Description
            Exit Sub
        End If
        LogMessage "MAPPING: Appending to existing mapping file"
    Else
        LogMessage "MAPPING: Creating new mapping file..."
        ' Create new file with header
        Set mappingFile = fso.CreateTextFile(mappingFilePath, True)
        If Err.Number <>0 Then
            LogMessage "MAPPING: ERROR - Could not create file: " & Err.Description
            Exit Sub
        End If
        
        ' Write comprehensive header
        mappingFile.WriteLine "# STEP 1 MAPPING FILE - Generated: " & Now
        mappingFile.WriteLine "# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description"
        mappingFile.WriteLine ""
        LogMessage "MAPPING: Created new mapping file with header"
    End If

    If Err.Number <>0 Then
        LogMessage "MAPPING: ERROR - File operation failed: " & Err.Description & " (Error " & Err.Number & ")"
        Exit Sub
    End If

    ' Write comprehensive mappings
    ' mappingKeys was already declared and populated above for directory extraction

    LogMessage "MAPPING: Writing " & (UBound(mappingKeys) + 1) & " mappings to file..."

    Dim i, writeCount
    writeCount =0
    For i =0 To UBound(mappingKeys)
        Dim originalPath
        originalPath = mappingKeys(i)
        Dim mappingValue
        mappingValue = g_ComprehensiveMapping.Item(originalPath)

        ' Parse mapping: "newPath|originalFile|newFile|group|description"
        Dim parts
        parts = Split(mappingValue, "|")
        If UBound(parts) >=4 Then
            Dim newPath, originalFile, newFile, groupName, description
            newPath = parts(0)       ' newPath
            originalFile = parts(1)  ' originalFile
            newFile = parts(2)       ' newFile
            groupName = parts(3)     ' group
            description = parts(4)   ' description

            ' Write full mapping format for STEP 2
            mappingFile.WriteLine originalPath & "|" & newPath & "|" & originalFile & "|" & newFile & "|" & groupName & "|" & description
            writeCount = writeCount + 1
            LogMessage "MAPPING: Written: " & originalFile & " -> " & newFile
        Else
            LogMessage "MAPPING: WARNING - Invalid mapping format for: " & originalPath
        End If
    Next

    mappingFile.WriteLine ""
    mappingFile.WriteLine "# End of mapping file"

    mappingFile.Close
    LogMessage "MAPPING: File closed"

    ' Verify file was created
    If fso.FileExists(mappingFilePath) Then
        Dim fileObj
        Set fileObj = fso.GetFile(mappingFilePath)
        Dim fileSize
        fileSize = fileObj.Size
        LogMessage "MAPPING: VERIFICATION - File exists! Size: " & fileSize & " bytes"
    Else
        LogMessage "MAPPING: VERIFICATION - ERROR - File does not exist after save!"
    End If

    LogMessage "MAPPING: Save complete. Wrote " & writeCount & " mappings"
    LogMessage "MAPPING: Saved comprehensive mapping to: " & mappingFilePath
    LogMessage "MAPPING: === END SAVE MAPPING FILE ==="
End Sub

Sub ScanRegistryForCounters(existingCounters, userPrefix)
    ' Scan Windows Registry for existing counters to continue numbering
    ' Much more reliable than file-based approaches
    ' Now dynamically generates counter keys based on user's prefix

    LogMessage "SCAN: Scanning Registry for existing counters to continue numbering with prefix: " & userPrefix

    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    ' Generate dynamic counter keys based on user's prefix
    Dim counterKeys
    counterKeys = Array(userPrefix & "CH", userPrefix & "PL", userPrefix & "B", userPrefix & "A", _
                       userPrefix & "P", userPrefix & "SQ", userPrefix & "FL", userPrefix & "LPL", _
                       userPrefix & "IPE", userPrefix & "OTHER", userPrefix & "FLG", userPrefix & "R")

    Dim foundCount
    foundCount = 0

    Dim i
    For i = 0 To UBound(counterKeys)
        Dim keyName
        keyName = counterKeys(i)

        On Error Resume Next
        Dim currentValue
        currentValue = shell.RegRead(regPath & keyName)

        If Err.Number = 0 Then
            ' Key exists - add to dictionary
            existingCounters.Add keyName, currentValue
            LogMessage "SCAN: Found existing counter: " & keyName & " = " & currentValue
            foundCount = foundCount + 1
        Else
            LogMessage "SCAN: No existing counter for: " & keyName & " (will start from 1)"
        End If

        On Error GoTo 0
    Next

    If foundCount > 0 Then
        LogMessage "SCAN: Loaded " & foundCount & " existing counters from Registry"
    Else
        LogMessage "SCAN: No existing counters found in Registry - starting fresh"
    End If
End Sub

Sub SaveCounterToRegistry(prefixGroupKey, finalCounter)
    ' Save counter to Registry for persistence across runs

    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    On Error Resume Next
    shell.RegWrite regPath & prefixGroupKey, finalCounter, "REG_DWORD"

    If Err.Number = 0 Then
        LogMessage "REGISTRY: Saved " & prefixGroupKey & " = " & finalCounter
    Else
        LogMessage "REGISTRY: ERROR - Could not save " & prefixGroupKey & ": " & Err.Description
    End If

    On Error GoTo 0
End Sub

Function CheckIfPrefixExistsInRegistry(userPrefix)
    ' Check if any registry keys exist for the given prefix
    CheckIfPrefixExistsInRegistry = False

    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    ' Generate dynamic counter keys based on user's prefix
    Dim counterKeys
    counterKeys = Array(userPrefix & "CH", userPrefix & "PL", userPrefix & "B", userPrefix & "A", _
                       userPrefix & "P", userPrefix & "SQ", userPrefix & "FL", userPrefix & "LPL", _
                       userPrefix & "IPE", userPrefix & "OTHER", userPrefix & "FLG", userPrefix & "R")

    Dim i
    For i = 0 To UBound(counterKeys)
        Dim keyName
        keyName = counterKeys(i)

        On Error Resume Next
        Dim currentValue
        currentValue = shell.RegRead(regPath & keyName)

        If Err.Number = 0 Then
            CheckIfPrefixExistsInRegistry = True
            Exit Function
        End If

        On Error GoTo 0
    Next
End Function

Sub ExtractAndUpdateCounter(fileName, existingCounters)
    ' Extract prefix+group and number from filename
    ' Expected format: PREFIX-###-GROUP###.ipt (e.g., NCRH01-000-PL173.ipt)

    LogMessage "SCAN: Analyzing fileName: " & fileName

    ' Handle .ipt extension
    Dim baseName
    If LCase(Right(fileName, 4)) = ".ipt" Then
        baseName = Left(fileName, Len(fileName) - 4) ' Remove .ipt
    Else
        baseName = fileName ' Already without extension
    End If

    LogMessage "SCAN: BaseName after extension removal: " & baseName

    ' Find the last group of letters followed by numbers at the end
    Dim i, groupStart, numberStart
    groupStart = 0
    numberStart = 0

    ' Scan backwards to find where numbers start
    For i = Len(baseName) To 1 Step -1
        Dim char
        char = Mid(baseName, i, 1)
        If IsNumeric(char) And numberStart = 0 Then
            numberStart = i
        ElseIf Not IsNumeric(char) And numberStart > 0 Then
            groupStart = i + 1
            Exit For
        End If
    Next

    LogMessage "SCAN: groupStart=" & groupStart & ", numberStart=" & numberStart

    If groupStart > 0 And numberStart > 0 Then
        Dim groupPart
        groupPart = Mid(baseName, groupStart, numberStart - groupStart)
        Dim numberPart
        numberPart = Mid(baseName, numberStart)
        Dim prefixPart
        prefixPart = Left(baseName, groupStart - 1)

        LogMessage "SCAN: Parsed - Prefix: '" & prefixPart & "', Group: '" & groupPart & "', Number: '" & numberPart & "'"

        If IsNumeric(numberPart) And Len(groupPart) > 0 Then
            Dim prefixGroupKey
            prefixGroupKey = prefixPart & groupPart
            Dim currentNumber
            currentNumber = CInt(numberPart)

            LogMessage "SCAN: Valid parse - Key: '" & prefixGroupKey & "', Number: " & currentNumber

            ' Update highest number for this prefix+group combination
            If existingCounters.Exists(prefixGroupKey) Then
                Dim existingNumber
                existingNumber = existingCounters.Item(prefixGroupKey)
                If currentNumber > existingNumber Then
                    existingCounters.Item(prefixGroupKey) = currentNumber
                    LogMessage "SCAN: Updated " & prefixGroupKey & " from " & existingNumber & " to " & currentNumber
                Else
                    LogMessage "SCAN: Kept existing " & prefixGroupKey & " at " & existingNumber & " (current " & currentNumber & " is lower)"
                End If
            Else
                existingCounters.Add prefixGroupKey, currentNumber
                LogMessage "SCAN: Added new " & prefixGroupKey & " = " & currentNumber
            End If
        Else
            LogMessage "SCAN: WARNING - Invalid number or group: numberPart='" & numberPart & "', groupPart='" & groupPart & "'"
        End If
    Else
        LogMessage "SCAN: WARNING - Could not parse group and number from: " & baseName
    End If
End Sub

Function ExtractPrefixFromScheme(namingScheme)
    ' Extract prefix from naming scheme (everything before the group part)
    ' e.g., "NCRH01-000-PL{N}" -> "NCRH01-000-"

    LogMessage "EXTRACT: Processing naming scheme: " & namingScheme

    Dim lastDashPos
    lastDashPos = InStrRev(namingScheme, "-")

    Dim result
    If lastDashPos > 0 Then
        ' The prefix is everything up to and including the last dash
        result = Left(namingScheme, lastDashPos)
    Else
        ' Fallback - find where {N} starts and take everything before it
        Dim nPos
        nPos = InStr(namingScheme, "{N}")
        If nPos > 0 Then
            ' Find the start of the group by scanning backwards from {N}
            Dim i
            For i = nPos - 1 To 1 Step -1
                If Not IsLetter(Mid(namingScheme, i, 1)) Then
                    result = Left(namingScheme, i)
                    Exit For
                End If
            Next
            If result = "" Then result = "" ' If no delimiter found, no prefix
        Else
            result = namingScheme ' No {N} found, use whole scheme
        End If
    End If

    LogMessage "EXTRACT: Extracted prefix: '" & result & "'"
    ExtractPrefixFromScheme = result
End Function

Function IsLetter(char)
    Dim asciiValue
    asciiValue = Asc(UCase(char))
    IsLetter = (asciiValue >= 65 And asciiValue <= 90) ' A-Z
End Function

' === LOGGING FUNCTIONS ===
Sub StartLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    ' Go up one level to FINAL_PRODUCTION_SCRIPTS root, then to Logs folder
    Dim rootDir
    rootDir = fso.GetParentFolderName(scriptDir)
    Dim logsDir
    logsDir = rootDir & "\Logs"
    ' Ensure Logs directory exists
    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder(logsDir)
    End If
    g_LogPath = logsDir & "\DynamicHeritage_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFileNum = fso.CreateTextFile(g_LogPath, True)
End Sub

Sub LogMessage(message)
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
    End If
    WScript.Echo message
End Sub

Sub StopLogging()
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.Close
    End If
End Sub