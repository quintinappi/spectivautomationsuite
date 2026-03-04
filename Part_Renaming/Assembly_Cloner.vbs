Option Explicit

' ==============================================================================
' ASSEMBLY CLONER - Copy Assembly with All Sub-Assemblies and Parts to New Location
' ==============================================================================
' Author: Quintin de Bruin © 2025-2026
' Last Updated: January 20, 2026
' 
' This script:
' 1. Detects currently open assembly in Inventor
' 2. Asks for destination folder
' 3. Copies assembly, ALL sub-assemblies, parts, and IDW files to destination
' 4. Updates all references in copied assemblies to use local copies
' 5. Recursively updates ALL IDW files in ALL subfolders
' 6. Optionally applies heritage renaming to all parts
' 7. Generates STEP_1_MAPPING.txt for reference tracking
'
' CHANGELOG:
' - Jan 20, 2026: Fixed recursive IDW processing - now updates IPT references
'                 in ALL sub-assembly IDWs (Bottom, Top, Middle, etc.)
'                 Added ScanIDWFilesForUpdate() for recursive IDW collection
' - Jan 14, 2026: Added sub-assembly support, copies entire hierarchy
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_DestLogFileNum
Dim g_DestLogPath
Dim g_CopiedFiles       ' Dictionary: originalPath -> newPath
Dim g_PlantSection      ' User-defined plant section prefix (optional)
Dim g_DoRename          ' Boolean: whether to rename parts
Dim g_ComponentGroups   ' Dictionary for grouping (if renaming)
Dim g_NamingSchemes     ' Dictionary for naming schemes (if renaming)
Dim g_SourceRoot        ' Source root folder for preserving folder structure

Call ASSEMBLY_CLONER_MAIN()

Sub ASSEMBLY_CLONER_MAIN()
    Call StartLogging
    LogMessage "=== ASSEMBLY CLONER ==="
    LogMessage "Copy assembly with all sub-assemblies and parts to new isolated location"
    
    Dim result
    result = MsgBox("ASSEMBLY CLONER" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Detect your currently open assembly" & vbCrLf & _
                    "2. Copy assembly + ALL sub-assemblies + parts to a NEW folder" & vbCrLf & _
                    "3. Update references to use local copies" & vbCrLf & _
                    "4. Copy associated IDW drawings" & vbCrLf & _
                    "5. Optionally rename parts with heritage naming" & vbCrLf & vbCrLf & _
                    "This creates a FULLY ISOLATED copy with no" & vbCrLf & _
                    "cross-references to original files!" & vbCrLf & vbCrLf & _
                    "Make sure your source assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Assembly Cloner")
    
    If result = vbNo Then
        LogMessage "User cancelled workflow"
        Exit Sub
    End If
    
    ' Initialize collections
    Set g_CopiedFiles = CreateObject("Scripting.Dictionary")
    Set g_ComponentGroups = CreateObject("Scripting.Dictionary")
    Set g_NamingSchemes = CreateObject("Scripting.Dictionary")
    
    ' Connect to existing Inventor application
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open your source assembly first.", vbCritical
        Exit Sub
    End If
    
    LogMessage "SUCCESS: Connected to existing Inventor instance"
    Err.Clear
    On Error GoTo 0
    
    ' Step 1: Detect open assembly
    LogMessage "STEP 1: Detecting open assembly"
    Dim sourceDoc
    Set sourceDoc = DetectOpenAssembly(invApp)
    If sourceDoc Is Nothing Then
        MsgBox "ERROR: No assembly is currently open in Inventor!" & vbCrLf & _
               "Please open your source assembly first.", vbCritical
        Exit Sub
    End If
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourceDir
    sourceDir = fso.GetParentFolderName(sourceDoc.FullFileName)
    Dim sourceFileName
    sourceFileName = fso.GetFileName(sourceDoc.FullFileName)
    
    ' Store source root for preserving folder structure
    g_SourceRoot = sourceDir
    
    ' Step 2: Get destination folder
    LogMessage "STEP 2: Getting destination folder from user"
    Dim destFolder
    destFolder = GetDestinationFolder(sourceDir)
    If destFolder = "" Then
        LogMessage "User cancelled - no destination folder selected"
        Exit Sub
    End If

    ' Start destination logging in the selected folder
    Call StartDestinationLogging(destFolder)
    
    ' Get new assembly name based on destination folder
    Dim destFolderName
    destFolderName = fso.GetFileName(destFolder)  ' Gets folder name like "Access Walkway-3"
    Dim newAsmFileName
    newAsmFileName = destFolderName & ".iam"
    LogMessage "New assembly name will be: " & newAsmFileName
    
    ' Step 3: Ask about renaming
    LogMessage "STEP 3: Asking about heritage renaming"
    Dim renameResult
    renameResult = MsgBox("HERITAGE RENAMING" & vbCrLf & vbCrLf & _
                          "Do you want to rename parts with heritage naming?" & vbCrLf & vbCrLf & _
                          "YES = Rename parts (e.g., PLANT-001-PL1, PLANT-001-CH1)" & vbCrLf & _
                          "NO = Keep original part names" & vbCrLf & vbCrLf & _
                          "Either way, parts will be copied to new folder.", _
                          vbYesNo + vbQuestion, "Rename Parts?")
    
    g_DoRename = (renameResult = vbYes)
    
    If g_DoRename Then
        ' Get naming prefix
        Call GetPlantSectionNaming()
    End If
    
    ' Step 4: Analyze and collect all referenced files
    LogMessage "STEP 4: Analyzing assembly structure"
    Dim allParts
    Set allParts = CreateObject("Scripting.Dictionary")
    Call CollectAllReferencedParts(sourceDoc, allParts)
    
    ' Also collect IDW drawing files from the source directory
    Call CollectIDWFiles(sourceDir, allParts)
    
    LogMessage "ANALYZE: Found " & allParts.Count & " unique parts to copy"
    
    ' Step 5: If renaming, group components and get naming schemes
    If g_DoRename Then
        LogMessage "STEP 5: Grouping components for heritage naming"
        Call GroupPartsForRenaming(invApp, allParts)
        Call GetUserNamingSchemes()
    End If
    
    ' Step 6: Copy assembly file first (with new name based on destination folder)
    LogMessage "STEP 6: Copying assembly to destination"
    Dim newAsmPath
    newAsmPath = destFolder & "\" & newAsmFileName
    
    ' Close the source document before copying
    LogMessage "Closing source assembly for safe copy..."
    Dim sourceFullPath
    sourceFullPath = sourceDoc.FullFileName
    sourceDoc.Close
    Set sourceDoc = Nothing
    
    ' Copy assembly file with new name
    On Error Resume Next
    fso.CopyFile sourceFullPath, newAsmPath, True
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not copy assembly: " & Err.Description
        MsgBox "ERROR: Could not copy assembly file!" & vbCrLf & Err.Description, vbCritical
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    LogMessage "COPIED: Assembly " & sourceFileName & " -> " & newAsmFileName
    
    ' Also store assembly mapping for IDW reference updates
    g_CopiedFiles.Add sourceFullPath, newAsmPath
    
    ' Step 7: Copy all parts and sub-assemblies to destination (with optional renaming for parts)
    LogMessage "STEP 7: Copying parts and sub-assemblies to destination"
    Call CopyAllFiles(invApp, allParts, destFolder)
    
    ' Step 8: Update assembly references
    ' CRITICAL: Open MAIN assembly FIRST to load all parts into memory
    ' This prevents "multiple files found" dialogs for sub-assemblies
    LogMessage "STEP 8: Updating references in ALL copied assemblies"
    
    ' ========================================================================
    ' CRITICAL: Open files in DEPENDENCY ORDER to avoid project file dialogs
    ' Parts have no references so they open cleanly.
    ' Once parts are in Inventor's memory, assemblies find them there first.
    ' ========================================================================
    
    Dim invAppForUpdate
    Set invAppForUpdate = GetObject(, "Inventor.Application")
    
    ' Store original settings
    Dim origSilent
    origSilent = invAppForUpdate.SilentOperation
    
    ' Enable silent mode 
    invAppForUpdate.SilentOperation = True
    
    LogMessage "STEP 8a: Pre-loading all copied PARTS into memory..."
    
    ' First, open ALL copied parts (they have no references, so no dialogs)
    Dim copiedKey, copiedPath
    For Each copiedKey In g_CopiedFiles.Keys
        copiedPath = g_CopiedFiles.Item(copiedKey)
        
        ' Only open .ipt files (parts)
        If LCase(Right(copiedPath, 4)) = ".ipt" Then
            On Error Resume Next
            Dim partDoc
            Set partDoc = invAppForUpdate.Documents.Open(copiedPath, False)
            If Not partDoc Is Nothing Then
                LogMessage "PRELOAD: " & GetFileNameFromPath(copiedPath)
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    
    LogMessage "STEP 8b: Pre-loading all copied SUB-ASSEMBLIES..."
    
    ' Next, open ALL copied assemblies EXCEPT the main one (smaller first)
    For Each copiedKey In g_CopiedFiles.Keys
        copiedPath = g_CopiedFiles.Item(copiedKey)
        
        ' Only open .iam files (assemblies), skip main assembly
        If LCase(Right(copiedPath, 4)) = ".iam" Then
            If LCase(copiedPath) <> LCase(newAsmPath) Then
                On Error Resume Next
                Dim subAsmDoc
                Set subAsmDoc = invAppForUpdate.Documents.Open(copiedPath, False)
                If Not subAsmDoc Is Nothing Then
                    LogMessage "PRELOAD: " & GetFileNameFromPath(copiedPath)
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next
    
    LogMessage "STEP 8c: Opening main assembly with all components in memory..."
    
    ' Finally, open the main assembly - all its references should already be loaded
    Dim mainAsmDoc
    On Error Resume Next
    Set mainAsmDoc = invAppForUpdate.Documents.Open(newAsmPath, False)
    
    If Not mainAsmDoc Is Nothing Then
        LogMessage "STEP 8c: Main assembly opened successfully"
        
        ' Now update references in all assemblies using in-memory documents
        LogMessage "STEP 8d: Updating file references in all assemblies..."
        
        ' Count assemblies to process
        Dim asmCount
        asmCount = 0
        For Each copiedKey In g_CopiedFiles.Keys
            copiedPath = g_CopiedFiles.Item(copiedKey)
            If LCase(Right(copiedPath, 4)) = ".iam" Then
                asmCount = asmCount + 1
            End If
        Next
        LogMessage "STEP 8d: Found " & asmCount & " assemblies to update"
        
        ' Process all assemblies including the main one
        Dim processedCount
        processedCount = 0
        For Each copiedKey In g_CopiedFiles.Keys
            copiedPath = g_CopiedFiles.Item(copiedKey)
            
            ' Only update .iam files (assemblies)
            If LCase(Right(copiedPath, 4)) = ".iam" Then
                processedCount = processedCount + 1
                LogMessage "STEP 8d: [" & processedCount & "/" & asmCount & "] Updating references in " & GetFileNameFromPath(copiedPath)
                Call UpdateInMemoryAssemblyReferences(invAppForUpdate, copiedPath)
            End If
        Next
        
        ' CRITICAL: Explicitly update the main assembly (in case it was missed in the loop)
        LogMessage "STEP 8d: Explicitly updating MAIN assembly: " & GetFileNameFromPath(newAsmPath)
        Call UpdateInMemoryAssemblyReferences(invAppForUpdate, newAsmPath)
        
        ' Save all open documents
        LogMessage "STEP 8e: Saving all documents..."
        Dim doc
        For Each doc In invAppForUpdate.Documents
            If doc.Dirty Then
                doc.Save
                LogMessage "SAVED: " & GetFileNameFromPath(doc.FullFileName)
            End If
        Next
        
        LogMessage "STEP 8: Complete"
    Else
        LogMessage "STEP 8: ERROR - Could not open main assembly: " & Err.Description
    End If
    
    Err.Clear
    On Error GoTo 0
    
    ' Restore original settings
    invAppForUpdate.SilentOperation = origSilent
    
    ' Step 9: Update IDW drawing references to point to new assemblies and parts
    LogMessage "STEP 9: Updating IDW drawing references"
    Call UpdateIDWReferences(invApp, destFolder)
    
    ' Step 10: Write STEP_1_MAPPING.txt for future IDW updates
    LogMessage "STEP 10: Writing STEP_1_MAPPING.txt"
    Call WriteMappingFile(destFolder)
    
    ' Step 11: Update iProperties for copied documents (parts and assemblies)
    ' CRITICAL: This updates Part Number iProperty to match new filename
    LogMessage "STEP 11: Updating iProperties for copied documents"
    Call UpdateIPropertiesForCopiedDocuments(invApp)
    
    ' Step 12: Validation scan and destination log
    LogMessage "STEP 12: Validation scan and destination logging"
    Call ValidateCloneAndLog(sourceDir, destFolder)

    LogMessage "=== ASSEMBLY CLONER COMPLETED ==="
    Call StopLogging
    
    Dim summaryMsg
    summaryMsg = "ASSEMBLY CLONE COMPLETED!" & vbCrLf & vbCrLf & _
                 "✅ Assembly copied to: " & destFolder & vbCrLf & _
                 "✅ " & g_CopiedFiles.Count & " files copied (parts + sub-assemblies)"
    
    If g_DoRename Then
        summaryMsg = summaryMsg & " with part renaming"
    End If
    
    summaryMsg = summaryMsg & vbCrLf & _
                 "✅ References updated to local copies" & vbCrLf & _
                 "✅ iProperties updated in copied parts" & vbCrLf & _
                 "✅ IDW files copied and updated" & vbCrLf & vbCrLf & _
                 "The new assembly is now completely isolated!" & vbCrLf & vbCrLf & _
                 "Log: " & g_LogPath
    
    MsgBox summaryMsg, vbInformation, "Success!"
End Sub

Function DetectOpenAssembly(invApp)
    On Error Resume Next
    
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument
    
    If Err.Number <> 0 Or activeDoc Is Nothing Then
        LogMessage "No active document found"
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If
    
    ' Check by file extension
    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        LogMessage "File extension is not .iam: " & activeDoc.FullFileName
        MsgBox "Current file is not an assembly (.iam)!" & vbCrLf & _
               "Please open an assembly file.", vbExclamation
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If
    
    LogMessage "DETECTED: " & activeDoc.DisplayName
    LogMessage "DETECTED: Full path - " & activeDoc.FullFileName
    
    ' Count occurrences
    Dim occCount
    occCount = activeDoc.ComponentDefinition.Occurrences.Count
    
    ' Get folder path
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath
    folderPath = fso.GetParentFolderName(activeDoc.FullFileName)
    
    ' Confirm with user
    Dim confirmMsg
    confirmMsg = "SOURCE ASSEMBLY DETECTED" & vbCrLf & vbCrLf & _
                 "Assembly: " & activeDoc.DisplayName & vbCrLf & _
                 "Parts Count: " & occCount & " occurrences" & vbCrLf & _
                 "Location: " & folderPath & vbCrLf & vbCrLf & _
                 "Clone this assembly to a new location?"
    
    Dim confirmResult
    confirmResult = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Source Assembly")
    
    If confirmResult = vbNo Then
        LogMessage "User cancelled assembly confirmation"
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If
    
    Set DetectOpenAssembly = activeDoc
    Err.Clear
End Function

Function GetDestinationFolder(sourceDir)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Use folder browser dialog only (no manual paste)
    Dim shell
    Set shell = CreateObject("Shell.Application")

    Dim startFolder
    startFolder = fso.GetParentFolderName(sourceDir)

    Dim folder
    Set folder = shell.BrowseForFolder(0, "Select DESTINATION folder for the cloned assembly:" & vbCrLf & _
                                          "Source: " & sourceDir & vbCrLf & vbCrLf & _
                                          "TIP: Click 'Make New Folder' to create a new destination", _
                                          &H0011, startFolder)  ' &H0011 = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE

    If folder Is Nothing Then
        GetDestinationFolder = ""
        Exit Function
    End If

    Dim destPath
    destPath = folder.Self.Path

    ' Validate destination is different from source
    If LCase(destPath) = LCase(sourceDir) Then
        MsgBox "Destination cannot be the same as source folder!" & vbCrLf & _
               "Please select a different folder.", vbExclamation
        GetDestinationFolder = ""
        Exit Function
    End If

    LogMessage "DESTINATION: " & destPath
    GetDestinationFolder = destPath
End Function

Sub GetPlantSectionNaming()
    LogMessage "PLANT: Getting plant section naming convention from user"
    
    Dim plantInput
    plantInput = InputBox("DEFINE PROJECT PREFIX" & vbCrLf & vbCrLf & _
                         "Enter the project prefix for heritage naming:" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "  WALKWAY-3-    (for Walkway 3)" & vbCrLf & _
                         "  PLANT1-000-   (for Plant 1)" & vbCrLf & _
                         "  SSCR05-001-   (for Section 1)" & vbCrLf & vbCrLf & _
                         "This will create part numbers like:" & vbCrLf & _
                         "  PREFIX-PL1, PREFIX-CH1, PREFIX-B1, etc.", _
                         "Define Project Prefix", "CLONE-001-")
    
    If plantInput = "" Then
        g_PlantSection = "CLONE-001-"
    Else
        plantInput = Trim(plantInput)
        If Right(plantInput, 1) <> "-" Then
            plantInput = plantInput & "-"
        End If
        g_PlantSection = UCase(plantInput)
    End If
    
    LogMessage "PLANT: Using prefix: " & g_PlantSection
End Sub

Sub CollectAllReferencedParts(asmDoc, allParts)
    ' Recursively collect all unique part and sub-assembly files referenced by the assembly
    ' Uses ComponentDefinition.Occurrences for reliable traversal of assembly hierarchy
    LogMessage "COLLECT: Scanning assembly for all referenced parts and sub-assemblies..."
    
    Call CollectPartsRecursively(asmDoc, allParts, "ROOT")
    
    LogMessage "COLLECT: Found " & allParts.Count & " unique files"
End Sub

Sub CollectIDWFiles(sourceDir, allParts)
    ' Collect all IDW drawing files from the source directory tree
    LogMessage "COLLECT: Scanning for IDW drawing files..."
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Call CollectIDWFilesRecursive(sourceDir, allParts, fso)
End Sub

Sub CollectIDWFilesRecursive(folderPath, allParts, fso)
    On Error Resume Next
    
    Dim folder
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    ' Check files in current folder
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            Dim fullPath
            fullPath = file.Path
            If Not allParts.Exists(fullPath) Then
                allParts.Add fullPath, "DRAWING"
                LogMessage "COLLECT: DRAWING " & GetFileNameFromPath(fullPath)
            End If
        End If
    Next
    
    ' Recurse into subfolders (skip OldVersions)
    Dim subFolder
    For Each subFolder In folder.SubFolders
        ' Skip OldVersions folders completely
        If LCase(subFolder.Name) <> "oldversions" Then
            Call CollectIDWFilesRecursive(subFolder.Path, allParts, fso)
        Else
            LogMessage "COLLECT: Skipping OldVersions folder: " & subFolder.Path
        End If
    Next
    
    Err.Clear
End Sub

Sub CollectPartsRecursively(asmDoc, allParts, level)
    ' Recursively traverse assembly hierarchy using Occurrences
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences
    
    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)
        
        If Not occ.Suppressed Then
            On Error Resume Next
            Dim doc
            Set doc = occ.Definition.Document
            
            If Err.Number = 0 And Not doc Is Nothing Then
                Dim fullPath
                fullPath = doc.FullFileName
                
                ' Skip files from OldVersions folders
                If InStr(1, LCase(fullPath), "\oldversions\", vbTextCompare) > 0 Then
                    LogMessage "COLLECT: Skipping OldVersions file: " & GetFileNameFromPath(fullPath)
                Else
                
                Dim fileName
                fileName = GetFileNameFromPath(fullPath)
                
                If LCase(Right(fileName, 4)) = ".ipt" Then
                    ' It's a part file
                    If Not allParts.Exists(fullPath) Then
                        ' Get description for grouping
                        Dim description
                        description = GetDescriptionFromIProperty(doc)
                        allParts.Add fullPath, description
                        LogMessage "COLLECT: PART " & fileName & " (" & description & ") at " & level
                    End If
                    
                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    ' It's a sub-assembly - add it and recurse
                    If Not allParts.Exists(fullPath) Then
                        allParts.Add fullPath, "SUB-ASSEMBLY"
                        LogMessage "COLLECT: SUB-ASSEMBLY " & fileName & " at " & level
                    End If
                    
                    ' Recurse into sub-assembly regardless of name
                    Call CollectPartsRecursively(doc, allParts, level & ">" & fileName)
                End If
                
                End If  ' End OldVersions check
            Else
                LogMessage "COLLECT: Could not access document for occurrence " & occ.Name & " at " & level
            End If
            Err.Clear
        Else
            LogMessage "COLLECT: Skipping suppressed occurrence " & occ.Name & " at " & level
        End If
    Next
End Sub


Sub GroupPartsForRenaming(invApp, allParts)
    ' Group parts by description for heritage naming (skip sub-assemblies)
    LogMessage "GROUP: Grouping parts for heritage naming..."
    
    Dim partKeys
    partKeys = allParts.Keys
    
    Dim i
    For i = 0 To UBound(partKeys)
        Dim partPath
        partPath = partKeys(i)
        Dim description
        description = allParts.Item(partPath)
        
        ' Skip sub-assemblies and drawings - only process parts
        Dim fileName
        fileName = GetFileNameFromPath(partPath)
        If LCase(Right(fileName, 4)) = ".iam" Or LCase(Right(fileName, 4)) = ".idw" Then
            Dim fileType
            If LCase(Right(fileName, 4)) = ".iam" Then
                fileType = "sub-assembly"
            Else
                fileType = "drawing"
            End If
            LogMessage "GROUP: Skipping " & fileType & " " & fileName
        Else
            ' Classify by description
            Dim groupCode
            groupCode = ClassifyByDescription(description)
            
            If groupCode <> "SKIP" Then
                If Not g_ComponentGroups.Exists(groupCode) Then
                    g_ComponentGroups.Add groupCode, CreateObject("Scripting.Dictionary")
                End If
                
                Dim groupDict
                Set groupDict = g_ComponentGroups.Item(groupCode)
                
                If Not groupDict.Exists(partPath) Then
                    groupDict.Add partPath, partPath & "|" & description & "|" & fileName
                End If
            End If
        End If
    Next
    
    LogMessage "GROUP: Created " & g_ComponentGroups.Count & " groups"
End Sub

Sub GetUserNamingSchemes()
    ' Get naming schemes for each group
    Dim groupKeys
    groupKeys = g_ComponentGroups.Keys
    
    Dim i
    For i = 0 To UBound(groupKeys)
        Dim groupName
        groupName = groupKeys(i)
        
        Dim groupDict
        Set groupDict = g_ComponentGroups.Item(groupName)
        
        ' Generate default scheme
        Dim defaultScheme
        defaultScheme = g_PlantSection & groupName & "{N}"
        
        ' For simplicity, just use default schemes (can add InputBox for customization)
        g_NamingSchemes.Add groupName, defaultScheme
        LogMessage "SCHEME: " & groupName & " -> " & defaultScheme
    Next
End Sub

Sub CopyAllFiles(invApp, allParts, destFolder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim partKeys
    partKeys = allParts.Keys
    
    ' If renaming, use group-based counters
    Dim groupCounters
    Set groupCounters = CreateObject("Scripting.Dictionary")

    ' Initialize counters from Registry if renaming
    If g_DoRename Then
        LogMessage "REGISTRY: Loading existing counters for prefix: " & g_PlantSection

        Dim existingCounters
        Set existingCounters = CreateObject("Scripting.Dictionary")
        Call ScanRegistryForCounters(existingCounters, g_PlantSection)

        Dim groupKeys
        groupKeys = g_ComponentGroups.Keys
        Dim g
        For g = 0 To UBound(groupKeys)
            Dim groupName
            groupName = groupKeys(g)
            Dim prefixGroupKey
            prefixGroupKey = g_PlantSection & groupName

            Dim startingCounter
            If existingCounters.Exists(prefixGroupKey) Then
                Dim highestExisting
                highestExisting = existingCounters.Item(prefixGroupKey)
                startingCounter = highestExisting + 1
                LogMessage "REGISTRY: Group '" & groupName & "' continuing from number " & startingCounter & " (found existing highest = " & highestExisting & ")"
            Else
                startingCounter = 1
                LogMessage "REGISTRY: Group '" & groupName & "' starting from number 1 (new prefix or group - key '" & prefixGroupKey & "' not found)"
            End If

            groupCounters.Add groupName, startingCounter
        Next
    End If

    Dim i
    For i = 0 To UBound(partKeys)
        Dim originalPath
        originalPath = partKeys(i)
        
        Dim originalFileName
        originalFileName = fso.GetFileName(originalPath)
        
        Dim newFileName
        Dim newPath
        
        ' Check file type and determine new name
        If LCase(Right(originalFileName, 4)) = ".iam" Then
            ' Sub-assembly - keep original name (don't apply part renaming logic)
            newFileName = originalFileName
            LogMessage "COPY: SUB-ASSEMBLY " & originalFileName & " (keeping original name)"
        ElseIf LCase(Right(originalFileName, 4)) = ".idw" Then
            ' Drawing file - keep original name
            newFileName = originalFileName
            LogMessage "COPY: DRAWING " & originalFileName & " (keeping original name)"
        ElseIf g_DoRename Then
            ' Part file with renaming enabled
            Dim description
            description = allParts.Item(originalPath)
            Dim groupCode
            groupCode = ClassifyByDescription(description)
            
            If groupCode = "SKIP" Then
                ' Hardware - just copy with original name
                newFileName = originalFileName
            Else
                ' Generate new name using scheme
                Dim scheme
                scheme = g_NamingSchemes.Item(groupCode)
                Dim counter
                counter = groupCounters.Item(groupCode)
                
                newFileName = Replace(scheme, "{N}", CStr(counter))
                If LCase(Right(newFileName, 4)) <> ".ipt" Then
                    newFileName = newFileName & ".ipt"
                End If
                
                ' Increment counter
                groupCounters.Item(groupCode) = counter + 1
            End If
        Else
            ' Part file without renaming - keep original name
            newFileName = originalFileName
        End If
        
        ' Compute relative path from source root to preserve folder structure
        Dim relativePath
        Dim originalDir
        originalDir = fso.GetParentFolderName(originalPath)
        
        If Len(originalDir) > Len(g_SourceRoot) And LCase(Left(originalDir, Len(g_SourceRoot))) = LCase(g_SourceRoot) Then
            ' File is in a subfolder of source root - preserve the subfolder structure
            relativePath = Mid(originalDir, Len(g_SourceRoot) + 2)  ' +2 to skip trailing backslash
            newPath = destFolder & "\" & relativePath & "\" & newFileName
            
            ' Create the subfolder if it doesn't exist
            Dim targetDir
            targetDir = destFolder & "\" & relativePath
            If Not fso.FolderExists(targetDir) Then
                Call CreateFolderRecursive(targetDir, fso)
                LogMessage "FOLDER: Created " & relativePath
            End If
        Else
            ' File is at source root level or from different location
            newPath = destFolder & "\" & newFileName
        End If
        
        ' Copy the file
        On Error Resume Next
        fso.CopyFile originalPath, newPath, True
        
        If Err.Number = 0 Then
            LogMessage "COPIED: " & originalFileName & " -> " & newFileName
            g_CopiedFiles.Add originalPath, newPath
        Else
            LogMessage "ERROR: Could not copy " & originalFileName & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    Next

    ' Save final counters to Registry if renaming was enabled
    If g_DoRename Then
        LogMessage "REGISTRY: Saving final counters to Registry"

        Dim groupKeysSave
        groupKeysSave = groupCounters.Keys
        Dim gs
        For gs = 0 To UBound(groupKeysSave)
            Dim groupNameSave
            groupNameSave = groupKeysSave(gs)
            Dim prefixGroupKeySave
            prefixGroupKeySave = g_PlantSection & groupNameSave
            Dim finalCounter
            finalCounter = groupCounters.Item(groupNameSave) - 1  ' Last used number (counter was incremented after use)

            Call SaveCounterToRegistry(prefixGroupKeySave, finalCounter)
        Next

        LogMessage "REGISTRY: All counters saved successfully"
    End If
End Sub

Sub ProcessIDWFilesWithReferenceUpdate(invApp, sourceDir, destFolder, newAsmPath)
    ' Process IDW files: Open from source, update references, save to destination
    ' This avoids the "Non-Unique Project File Names" dialog by updating references
    ' BEFORE moving the file to the new location
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    Dim folder
    Set folder = fso.GetFolder(sourceDir)
    
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            LogMessage "IDW PROCESS: Found " & file.Name
            
            ' Ask user for new IDW name with auto-increment suggestion
            Dim suggestedName
            suggestedName = IncrementFileName(file.Name)
            
            Dim newIdwName
            newIdwName = InputBox("NEW IDW NAME" & vbCrLf & vbCrLf & _
                                  "Original IDW: " & file.Name & vbCrLf & vbCrLf & _
                                  "Enter new name for the IDW file:" & vbCrLf & _
                                  "(Include .idw extension)", _
                                  "Rename IDW", suggestedName)
            
            If newIdwName = "" Then
                newIdwName = suggestedName  ' Use suggested if cancelled
            End If
            
            ' Ensure .idw extension
            If LCase(Right(newIdwName, 4)) <> ".idw" Then
                newIdwName = newIdwName & ".idw"
            End If
            
            Dim destIdwPath
            destIdwPath = destFolder & "\" & newIdwName
            
            ' Close all documents first
            invApp.Documents.CloseAll
            
            ' Suppress dialogs during IDW processing
            invApp.SilentOperation = True
            
            ' Open the ORIGINAL IDW (it has valid references to original parts)
            LogMessage "IDW PROCESS: Opening original IDW from source location..."
            Dim idwDoc
            Set idwDoc = invApp.Documents.Open(file.Path, False)
            
            If Err.Number <> 0 Or idwDoc Is Nothing Then
                LogMessage "IDW PROCESS: ERROR - Could not open " & file.Name & ": " & Err.Description
                Err.Clear
            Else
                LogMessage "IDW PROCESS: Opened successfully, now updating references..."
                
                ' Update references to point to NEW paths in destination folder
                Dim fileDescriptors
                Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
                
                LogMessage "IDW PROCESS: Found " & fileDescriptors.Count & " references"
                
                Dim updatedCount
                updatedCount = 0
                
                Dim i
                For i = 1 To fileDescriptors.Count
                    Dim fd
                    Set fd = fileDescriptors.Item(i)
                    
                    Dim refPath
                    refPath = fd.FullFileName
                    Dim refFileName
                    refFileName = GetFileNameFromPath(refPath)
                    
                    ' Check if we have a mapping for this file (original -> new)
                    If g_CopiedFiles.Exists(refPath) Then
                        Dim newRefPath
                        newRefPath = g_CopiedFiles.Item(refPath)
                        
                        LogMessage "IDW PROCESS: Updating " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
                        
                        fd.ReplaceReference newRefPath
                        
                        If Err.Number = 0 Then
                            LogMessage "IDW PROCESS: SUCCESS"
                            updatedCount = updatedCount + 1
                        Else
                            LogMessage "IDW PROCESS: ERROR - " & Err.Description
                        End If
                        Err.Clear
                    Else
                        LogMessage "IDW PROCESS: No mapping for " & refFileName
                    End If
                Next
                
                ' Now save to the NEW location with NEW name
                ' This is key: the IDW now has updated references pointing to new parts
                LogMessage "IDW PROCESS: Saving to destination: " & destIdwPath
                idwDoc.SaveAs destIdwPath, False
                
                If Err.Number = 0 Then
                    LogMessage "IDW PROCESS: Successfully saved " & newIdwName & " (" & updatedCount & " references updated)"
                Else
                    LogMessage "IDW PROCESS: ERROR saving - " & Err.Description
                End If
                
                idwDoc.Close
                Err.Clear
            End If
        End If
    Next
    
    ' Restore normal operation
    invApp.SilentOperation = False
End Sub

Function IncrementFileName(fileName)
    ' Try to increment a number in the filename
    ' e.g., "MGY-200-DRD-01-11.idw" -> "MGY-200-DRD-01-12.idw"
    Dim baseName, ext
    ext = Right(fileName, 4)  ' .idw
    baseName = Left(fileName, Len(fileName) - 4)
    
    ' Find the last number in the filename
    Dim i, lastNumStart, lastNumEnd
    lastNumStart = 0
    lastNumEnd = 0
    
    Dim inNumber
    inNumber = False
    
    For i = Len(baseName) To 1 Step -1
        Dim c
        c = Mid(baseName, i, 1)
        If IsNumeric(c) Then
            If Not inNumber Then
                lastNumEnd = i
                inNumber = True
            End If
            lastNumStart = i
        Else
            If inNumber Then
                Exit For
            End If
        End If
    Next
    
    If lastNumStart > 0 And lastNumEnd > 0 Then
        ' Extract and increment the number
        Dim numStr, numVal, numLen
        numStr = Mid(baseName, lastNumStart, lastNumEnd - lastNumStart + 1)
        numLen = Len(numStr)
        numVal = CInt(numStr) + 1
        
        ' Pad with zeros to maintain length
        Dim newNumStr
        newNumStr = CStr(numVal)
        Do While Len(newNumStr) < numLen
            newNumStr = "0" & newNumStr
        Loop
        
        ' Rebuild filename
        IncrementFileName = Left(baseName, lastNumStart - 1) & newNumStr & Mid(baseName, lastNumEnd + 1) & ext
    Else
        ' No number found, just add -2
        IncrementFileName = baseName & "-2" & ext
    End If
End Function

Sub UpdateAssemblyReferencesWithApprentice(asmPath)
    ' Update assembly references - try ApprenticeServer first, fall back to Inventor
    On Error Resume Next
    
    ' Try ApprenticeServer first (silent, no dialogs)
    Dim apprentice
    Set apprentice = CreateObject("Inventor.ApprenticeServerComponent")
    
    If Err.Number <> 0 Or apprentice Is Nothing Then
        LogMessage "ASSEMBLY UPDATE: ApprenticeServer not available, using Inventor method"
        Err.Clear
        
        ' Fall back to Inventor method - update ALL copied assemblies
        Dim fallbackKey
        For Each fallbackKey In g_CopiedFiles.Keys
            copiedPath = g_CopiedFiles.Item(fallbackKey)
            
            ' Only update .iam files (assemblies)
            If LCase(Right(copiedPath, 4)) = ".iam" Then
                LogMessage "ASSEMBLY UPDATE (Inventor): Updating references in " & GetFileNameFromPath(copiedPath)
                Call UpdateAssemblyReferencesWithInventor(copiedPath)
            End If
        Next
        
        Exit Sub
    End If
    
    LogMessage "ASSEMBLY UPDATE: Using ApprenticeServer for silent reference updates"
    
    ' Update main assembly copy first
    Call UpdateAssemblyFileReferencesWithApprentice(apprentice, asmPath)
    
    ' Update all copied sub-assemblies (do NOT touch originals)
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim copiedPath
        copiedPath = g_CopiedFiles.Item(key)
        If LCase(Right(copiedPath, 4)) = ".iam" Then
            If LCase(copiedPath) <> LCase(asmPath) Then
                Call UpdateAssemblyFileReferencesWithApprentice(apprentice, copiedPath)
            End If
        End If
    Next
    
    Set apprentice = Nothing
    Err.Clear
End Sub

Sub UpdateAssemblyFileReferencesWithApprentice(apprentice, asmPath)
    ' Update references inside a single assembly file using ApprenticeServer
    On Error Resume Next
    
    Dim appDoc
    Set appDoc = apprentice.Open(asmPath)
    
    If Err.Number <> 0 Or appDoc Is Nothing Then
        LogMessage "ASSEMBLY UPDATE: ApprenticeServer failed to open " & asmPath & " - " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    LogMessage "ASSEMBLY UPDATE: Opened " & GetFileNameFromPath(asmPath)
    
    ' Get file references from the assembly
    Dim fileDescriptors
    Set fileDescriptors = appDoc.ReferencedFileDescriptors
    
    LogMessage "ASSEMBLY UPDATE: Found " & fileDescriptors.Count & " file references"
    
    Dim updatedCount
    updatedCount = 0
    
    Dim i
    For i = 1 To fileDescriptors.Count
        Dim fd
        Set fd = fileDescriptors.Item(i)
        
        Dim refPath
        refPath = fd.FullFileName
        Dim refFileName
        refFileName = GetFileNameFromPath(refPath)
        
        ' Check if we have a mapping for this file
        If g_CopiedFiles.Exists(refPath) Then
            Dim newRefPath
            newRefPath = g_CopiedFiles.Item(refPath)
            
            LogMessage "ASSEMBLY UPDATE: Replacing " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
            
            fd.ReplaceReference newRefPath
            
            If Err.Number = 0 Then
                LogMessage "ASSEMBLY UPDATE: SUCCESS"
                updatedCount = updatedCount + 1
            Else
                LogMessage "ASSEMBLY UPDATE: ERROR - " & Err.Description
            End If
            Err.Clear
        Else
            LogMessage "ASSEMBLY UPDATE: No mapping for " & refFileName
        End If
    Next
    
    ' Save the assembly
    LogMessage "ASSEMBLY UPDATE: Saving assembly..."
    appDoc.SaveAs asmPath, False
    
    If Err.Number = 0 Then
        LogMessage "ASSEMBLY UPDATE: Saved successfully (" & updatedCount & " references updated)"
    Else
        LogMessage "ASSEMBLY UPDATE: ERROR saving - " & Err.Description
    End If
    
    appDoc.Close
    Err.Clear
End Sub

Sub UpdateAssemblyFileReferences(invApp, asmDoc)
    ' Update file references in an open assembly document
    ' Uses File.ReferencedFileDescriptors to access stored (not resolved) paths
    
    On Error Resume Next
    
    LogMessage "REF UPDATE: Processing " & GetFileNameFromPath(asmDoc.FullFileName)
    
    ' Build lookup dictionaries
    Dim fileNameLookup, pathLookup
    Set fileNameLookup = CreateObject("Scripting.Dictionary")
    Set pathLookup = CreateObject("Scripting.Dictionary")
    
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim origFileName
        origFileName = LCase(GetFileNameFromPath(key))
        If Not fileNameLookup.Exists(origFileName) Then
            fileNameLookup.Add origFileName, g_CopiedFiles.Item(key)
        End If
        pathLookup.Add LCase(key), g_CopiedFiles.Item(key)
    Next
    
    ' Get file references
    Dim refDescs
    Set refDescs = asmDoc.File.ReferencedFileDescriptors
    
    If Err.Number <> 0 Or refDescs Is Nothing Then
        LogMessage "REF UPDATE: Cannot get ReferencedFileDescriptors: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    LogMessage "REF UPDATE: Found " & refDescs.Count & " references"
    
    Dim updatedCount
    updatedCount = 0
    
    Dim i
    For i = 1 To refDescs.Count
        Dim fd
        Set fd = refDescs.Item(i)
        
        If Not fd Is Nothing Then
            ' Get the STORED path (this is what we need to update)
            Dim storedPath
            storedPath = fd.FullFileName
            Dim storedFileName
            storedFileName = LCase(GetFileNameFromPath(storedPath))
            Dim storedPathLower
            storedPathLower = LCase(storedPath)
            
            ' Try to find a mapping
            Dim newRefPath
            newRefPath = ""
            
            ' First check exact path match
            If pathLookup.Exists(storedPathLower) Then
                newRefPath = pathLookup.Item(storedPathLower)
            ' Then check filename match
            ElseIf fileNameLookup.Exists(storedFileName) Then
                newRefPath = fileNameLookup.Item(storedFileName)
            End If
            
            If newRefPath <> "" Then
                ' Only update if paths differ
                If storedPathLower <> LCase(newRefPath) Then
                    LogMessage "REF UPDATE: " & storedFileName & " -> " & GetFileNameFromPath(newRefPath)
                    
                    fd.ReplaceReference newRefPath
                    
                    If Err.Number = 0 Then
                        updatedCount = updatedCount + 1
                    Else
                        LogMessage "REF UPDATE: ReplaceReference FAILED: " & Err.Description
                        Err.Clear
                    End If
                Else
                    LogMessage "REF UPDATE: SKIP (already correct): " & storedFileName
                End If
            Else
                LogMessage "REF UPDATE: NO MAPPING for: " & storedFileName
            End If
        End If
    Next
    
    LogMessage "REF UPDATE: Updated " & updatedCount & " of " & refDescs.Count & " references"
    
    Set fileNameLookup = Nothing
    Set pathLookup = Nothing
    Err.Clear
End Sub

Sub UpdateInMemoryAssemblyReferences(invApp, asmPath)
    ' Update references in an ALREADY-OPEN assembly document
    ' This function expects the assembly to already be loaded in Inventor
    
    On Error Resume Next
    
    ' Find the document in Inventor's open documents
    Dim asmDoc
    Set asmDoc = Nothing
    
    Dim doc
    For Each doc In invApp.Documents
        If LCase(doc.FullFileName) = LCase(asmPath) Then
            Set asmDoc = doc
            Exit For
        End If
    Next
    
    If asmDoc Is Nothing Then
        LogMessage "IN-MEMORY UPDATE: Document not found in memory: " & GetFileNameFromPath(asmPath)
        Exit Sub
    End If
    
    LogMessage "IN-MEMORY UPDATE: Found " & GetFileNameFromPath(asmPath) & " in memory"
    
    ' Build TWO lookups:
    ' 1. fileNameLookup: originalFilename -> newFullPath (for reference updates)
    ' 2. pathLookup: originalFullPath -> newFullPath (for exact path matching)
    Dim fileNameLookup, pathLookup, newPathSet
    Set fileNameLookup = CreateObject("Scripting.Dictionary")
    Set pathLookup = CreateObject("Scripting.Dictionary")
    Set newPathSet = CreateObject("Scripting.Dictionary")
    
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim origFileName
        origFileName = LCase(GetFileNameFromPath(key))
        If Not fileNameLookup.Exists(origFileName) Then
            fileNameLookup.Add origFileName, g_CopiedFiles.Item(key)
        End If
        ' Also add full path lookup (case-insensitive)
        pathLookup.Add LCase(key), g_CopiedFiles.Item(key)
        ' Track all NEW paths so we can skip them if already correct
        newPathSet.Add LCase(g_CopiedFiles.Item(key)), True
    Next
    
    LogMessage "IN-MEMORY UPDATE: Built lookup with " & fileNameLookup.Count & " filenames, " & pathLookup.Count & " paths"
    
    ' Get file references from the assembly
    Dim refDescs
    Set refDescs = asmDoc.File.ReferencedFileDescriptors
    
    If Err.Number <> 0 Or refDescs Is Nothing Then
        LogMessage "IN-MEMORY UPDATE: Cannot get ReferencedFileDescriptors: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    LogMessage "IN-MEMORY UPDATE: Found " & refDescs.Count & " references"
    
    Dim updatedCount
    updatedCount = 0
    
    Dim i
    For i = 1 To refDescs.Count
        Dim fd
        Set fd = refDescs.Item(i)
        
        If Not fd Is Nothing Then
            Dim refPath
            refPath = fd.FullFileName
            Dim refFileName
            refFileName = LCase(GetFileNameFromPath(refPath))
            Dim refPathLower
            refPathLower = LCase(refPath)
            
            ' Debug: Log what we're checking
            LogMessage "IN-MEMORY UPDATE: Checking ref: " & refFileName & " (path: " & refPath & ")"
            
            ' Check if reference is ALREADY pointing to a new path (skip if so)
            If newPathSet.Exists(refPathLower) Then
                LogMessage "IN-MEMORY UPDATE: SKIP (already correct path)"
            Else
                ' First try exact path match (handles case where reference is to original location)
                Dim newRefPath
                newRefPath = ""
                
                If pathLookup.Exists(refPathLower) Then
                    newRefPath = pathLookup.Item(refPathLower)
                    LogMessage "IN-MEMORY UPDATE: Found by EXACT PATH match"
                ElseIf fileNameLookup.Exists(refFileName) Then
                    newRefPath = fileNameLookup.Item(refFileName)
                    LogMessage "IN-MEMORY UPDATE: Found by FILENAME match"
                End If
                
                If newRefPath <> "" Then
                    ' Only replace if the paths are different
                    If LCase(refPath) <> LCase(newRefPath) Then
                        LogMessage "IN-MEMORY UPDATE: REPLACING " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
                    
                    fd.ReplaceReference newRefPath
                    
                    If Err.Number = 0 Then
                        updatedCount = updatedCount + 1
                    Else
                        LogMessage "IN-MEMORY UPDATE: ReplaceReference failed: " & Err.Description
                        Err.Clear
                    End If
                Else
                    LogMessage "IN-MEMORY UPDATE: SKIP (already correct path)"
                End If
            Else
                LogMessage "IN-MEMORY UPDATE: NO MAPPING for " & refFileName
            End If
        End If
    Next
    
    LogMessage "IN-MEMORY UPDATE: Updated " & updatedCount & " references in " & GetFileNameFromPath(asmPath)
    
    Set fileNameLookup = Nothing
    Set pathLookup = Nothing
    Set newPathSet = Nothing
    Err.Clear
End Sub

Sub UpdateAssemblyReferencesWithInventor(asmPath)
    ' Update assembly references using Inventor (fallback method)
    ' Opens the assembly, updates references via occurrences, saves
    
    On Error Resume Next
    
    ' Get Inventor application
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ASSEMBLY UPDATE (Inventor): ERROR - Cannot connect to Inventor"
        Exit Sub
    End If
    Err.Clear
    
    LogMessage "ASSEMBLY UPDATE (Inventor): Opening assembly..."
    
    ' Store and set options to suppress dialogs AND auto-resolve references
    Dim originalSilent, originalResolve
    originalSilent = invApp.SilentOperation
    originalResolve = invApp.FileOptions.ResolveFileOption
    
    invApp.SilentOperation = True
    invApp.FileOptions.ResolveFileOption = 54275 ' kSkipUnresolvedFiles
    
    ' Open the assembly
    Dim asmDoc
    Set asmDoc = invApp.Documents.Open(asmPath, False)
    
    If Err.Number <> 0 Or asmDoc Is Nothing Then
        LogMessage "ASSEMBLY UPDATE (Inventor): ERROR - Could not open assembly: " & Err.Description
        ' Restore settings before exit
        invApp.SilentOperation = originalSilent
        invApp.FileOptions.ResolveFileOption = originalResolve
        Exit Sub
    End If
    
    LogMessage "ASSEMBLY UPDATE (Inventor): Opened " & GetFileNameFromPath(asmPath)
    
    ' Update references using the recursive function
    Call UpdateAssemblyReferences(asmDoc)
    
    ' Save the assembly
    asmDoc.Save
    
    If Err.Number = 0 Then
        LogMessage "ASSEMBLY UPDATE (Inventor): Saved successfully"
    Else
        LogMessage "ASSEMBLY UPDATE (Inventor): ERROR saving - " & Err.Description
    End If
    
    ' Keep assembly open for user to verify
    LogMessage "ASSEMBLY UPDATE (Inventor): Complete - assembly left open for verification"
    
    ' Restore original settings
    invApp.SilentOperation = originalSilent
    invApp.FileOptions.ResolveFileOption = originalResolve
    
    Err.Clear
End Sub

Sub UpdateAssemblyReferences(asmDoc)
    ' Update all references in the assembly to point to local copies
    LogMessage "UPDATE: Updating assembly references to local copies..."
    
    Call UpdateReferencesRecursively(asmDoc, "ROOT")
End Sub

Sub UpdateReferencesRecursively(asmDoc, level)
    ' CRITICAL FIX: The copied assembly has references that may return relative paths
    ' or paths that don't match g_CopiedFiles keys exactly.
    ' Solution: Build a lookup by FILENAME (not full path) for matching
    
    On Error Resume Next
    
    ' Build a filename-based lookup from g_CopiedFiles
    Dim fileNameLookup
    Set fileNameLookup = CreateObject("Scripting.Dictionary")
    
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim origFileName
        origFileName = LCase(GetFileNameFromPath(key))
        If Not fileNameLookup.Exists(origFileName) Then
            fileNameLookup.Add origFileName, g_CopiedFiles.Item(key)
        End If
    Next
    
    LogMessage "UPDATE: Built filename lookup with " & fileNameLookup.Count & " entries"
    
    ' For Inventor documents, use asmDoc.File.ReferencedFileDescriptors
    Dim refDescs
    Set refDescs = asmDoc.File.ReferencedFileDescriptors
    
    If Err.Number <> 0 Or refDescs Is Nothing Then
        LogMessage "UPDATE: Cannot get File.ReferencedFileDescriptors (" & Err.Description & ")"
        Err.Clear
        Set fileNameLookup = Nothing
        Exit Sub
    End If
    
    LogMessage "UPDATE: Using File.ReferencedFileDescriptors method - " & refDescs.Count & " references at level " & level
    
    Dim updatedCount
    updatedCount = 0
    
    Dim i
    For i = 1 To refDescs.Count
        Dim fd
        Set fd = refDescs.Item(i)
        
        If Err.Number <> 0 Then
            LogMessage "UPDATE: Error accessing reference " & i & ": " & Err.Description
            Err.Clear
        Else
            Dim refPath
            refPath = fd.FullFileName
            Dim refFileName
            refFileName = GetFileNameFromPath(refPath)
            Dim refFileNameLower
            refFileNameLower = LCase(refFileName)
            
            ' Check if we have a mapping for this file BY FILENAME
            If fileNameLookup.Exists(refFileNameLower) Then
                Dim newRefPath
                newRefPath = fileNameLookup.Item(refFileNameLower)
                Dim newFileName
                newFileName = GetFileNameFromPath(newRefPath)
                
                LogMessage "UPDATE: " & refFileName & " -> " & newFileName
                
                fd.ReplaceReference newRefPath
                
                If Err.Number = 0 Then
                    updatedCount = updatedCount + 1
                Else
                    LogMessage "UPDATE ERROR: ReplaceReference failed - " & Err.Description
                End If
                Err.Clear
            Else
                ' Only log if it's a part/assembly we should care about
                If LCase(Right(refFileName, 4)) = ".ipt" Or LCase(Right(refFileName, 4)) = ".iam" Then
                    LogMessage "UPDATE: No mapping for " & refFileName
                End If
            End If
        End If
    Next
    
    LogMessage "UPDATE: Updated " & updatedCount & " of " & refDescs.Count & " references at level " & level
    
    Set fileNameLookup = Nothing
    Err.Clear
End Sub

Sub UpdateIDWReferences(invApp, destFolder)
    ' Update references in all copied IDW files using ApprenticeServer
    ' This avoids the "Non-Unique Project File Names" dialog
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    ' Create ApprenticeServer for silent file manipulation
    Dim apprentice
    Set apprentice = CreateObject("Inventor.ApprenticeServerComponent")
    
    If Err.Number <> 0 Or apprentice Is Nothing Then
        LogMessage "IDW UPDATE: Could not create ApprenticeServer, falling back to Inventor"
        LogMessage "IDW UPDATE: Error: " & Err.Description
        Err.Clear
        ' Fall back to regular Inventor method
        LogMessage "IDW UPDATE: Calling Inventor fallback method"
        Call UpdateIDWReferencesWithInventor(invApp, destFolder)
        Exit Sub
    End If
    
    LogMessage "IDW UPDATE: Using ApprenticeServer for silent reference updates"
    
    ' CRITICAL: Scan recursively for all IDW files (same as Inventor fallback)
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    
    Dim folder
    Set folder = fso.GetFolder(destFolder)
    Call ScanIDWFilesForUpdate(folder, idwFiles, fso)
    
    LogMessage "IDW UPDATE: Found " & idwFiles.Count & " IDW files (recursive scan)"
    
    Dim idwPath
    For Each idwPath In idwFiles.Keys
        LogMessage "IDW UPDATE: Processing " & idwFiles.Item(idwPath)
        
        Dim appDoc
        Set appDoc = apprentice.Open(idwPath)
            
            If Err.Number = 0 And Not appDoc Is Nothing Then
                ' Get file references
                Dim fileDescriptors
                Set fileDescriptors = appDoc.ReferencedFileDescriptors
                
                LogMessage "IDW UPDATE: Found " & fileDescriptors.Count & " references to check"
                
                Dim updatedCount
                updatedCount = 0
                
                Dim i
                For i = 1 To fileDescriptors.Count
                    Dim fd
                    Set fd = fileDescriptors.Item(i)

                    Dim refPath
                    refPath = fd.FullFileName
                    Dim refFileName
                    refFileName = GetFileNameFromPath(refPath)
                    Dim refFileNameLower
                    refFileNameLower = LCase(refFileName)

                    ' Try to find a mapping - first exact path, then filename
                    Dim newRefPath
                    newRefPath = ""

                    ' First check exact path match
                    If g_CopiedFiles.Exists(refPath) Then
                        newRefPath = g_CopiedFiles.Item(refPath)
                        LogMessage "IDW UPDATE: Found by EXACT PATH match"
                    Else
                        ' Fallback: lookup by filename (handles multiple folders with same part names)
                        Dim key
                        For Each key In g_CopiedFiles.Keys
                            If LCase(GetFileNameFromPath(key)) = refFileNameLower Then
                                newRefPath = g_CopiedFiles.Item(key)
                                LogMessage "IDW UPDATE: Found by FILENAME match"
                                Exit For
                            End If
                        Next
                    End If

                    If newRefPath <> "" Then
                        LogMessage "IDW UPDATE: Replacing " & refFileName & " -> " & GetFileNameFromPath(newRefPath)

                        fd.ReplaceReference newRefPath

                        If Err.Number = 0 Then
                            LogMessage "IDW UPDATE: SUCCESS - Updated reference"
                            updatedCount = updatedCount + 1
                        Else
                            LogMessage "IDW UPDATE: ERROR - " & Err.Description
                        End If
                        Err.Clear
                    Else
                        LogMessage "IDW UPDATE: No mapping for " & refFileName & " (path: " & refPath & ")"
                    End If
                Next
                
                ' Save changes
                appDoc.SaveAs idwPath, False
                appDoc.Close
                LogMessage "IDW UPDATE: Saved " & idwFiles.Item(idwPath) & " (" & updatedCount & " references updated)"
            Else
                LogMessage "IDW UPDATE: Could not open " & idwFiles.Item(idwPath) & " - " & Err.Description
            End If
            Err.Clear
    Next
    
    ' Clean up
    Set apprentice = Nothing
End Sub

Sub UpdateIDWReferencesWithInventor(invApp, destFolder)
    ' Uses SilentOperation to suppress all dialogs including multi-file resolution
    LogMessage "IDW UPDATE (Inventor): Starting IDW update for destFolder: " & destFolder
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    ' Store original settings
    Dim originalResolveMode
    originalResolveMode = invApp.FileOptions.ResolveFileOption
    
    Dim originalSilentMode
    originalSilentMode = invApp.SilentOperation
    
    ' CRITICAL: Enable silent operation to suppress ALL dialogs (including multi-file resolution)
    invApp.SilentOperation = True
    LogMessage "IDW UPDATE (Inventor): Enabled SilentOperation mode"
    
    ' Set resolve options to skip missing files globally
    invApp.FileOptions.ResolveFileOption = 54275 ' kSkipUnresolvedFiles
    LogMessage "IDW UPDATE (Inventor): Set resolve mode to skip unresolved files"
    
    ' Close all documents first
    invApp.Documents.CloseAll
    
    ' CRITICAL: Open the main cloned assembly FIRST so all parts are loaded in Inventor's memory
    ' This prevents "multiple files found" dialogs when opening IDW files
    Dim mainAsmPath
    mainAsmPath = ""
    
    Dim folder
    Set folder = fso.GetFolder(destFolder)
    
    ' Find the main assembly (should match folder name)
    Dim destFolderName
    destFolderName = fso.GetFileName(destFolder)
    Dim expectedAsmName
    expectedAsmName = destFolderName & ".iam"
    
    Dim file
    For Each file In folder.Files
        If LCase(file.Name) = LCase(expectedAsmName) Then
            mainAsmPath = file.Path
            Exit For
        End If
    Next
    
    ' If not found by folder name, find any .iam file
    If mainAsmPath = "" Then
        For Each file In folder.Files
            If LCase(Right(file.Name, 4)) = ".iam" Then
                mainAsmPath = file.Path
                Exit For
            End If
        Next
    End If
    
    ' Open the main assembly to load all parts into memory
    If mainAsmPath <> "" Then
        LogMessage "IDW UPDATE: Opening main assembly to load all parts: " & GetFileNameFromPath(mainAsmPath)
        Dim mainAsmDoc
        Set mainAsmDoc = invApp.Documents.Open(mainAsmPath, False)
        
        If Err.Number = 0 And Not mainAsmDoc Is Nothing Then
            LogMessage "IDW UPDATE: Main assembly opened successfully - all parts now in memory"
        Else
            LogMessage "IDW UPDATE: WARNING - Could not open main assembly: " & Err.Description
        End If
        Err.Clear
    Else
        LogMessage "IDW UPDATE: WARNING - No main assembly found in destination folder"
    End If
    
    ' RECURSIVE: Collect ALL IDW files from destination folder and all subfolders
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    Call ScanIDWFilesForUpdate(folder, idwFiles, fso)
    
    LogMessage "IDW UPDATE (Inventor): Found " & idwFiles.Count & " IDW files (recursive scan)"
    
    ' Now process ALL IDW files - references will resolve to already-loaded documents
    Dim idwPath
    For Each idwPath In idwFiles.Keys
        LogMessage "IDW UPDATE (Inventor): Processing " & idwFiles.Item(idwPath)
        
        Dim idwDoc
        Set idwDoc = invApp.Documents.Open(idwPath, False)
        
        If Err.Number = 0 And Not idwDoc Is Nothing Then
            Dim fileDescriptors
            Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
            
            LogMessage "IDW UPDATE: Found " & fileDescriptors.Count & " references"
            
            Dim updatedCount
            updatedCount = 0
            
            Dim i
            For i = 1 To fileDescriptors.Count
                Dim fd
                Set fd = fileDescriptors.Item(i)
                
                Dim refPath
                refPath = fd.FullFileName
                Dim refFileName
                refFileName = GetFileNameFromPath(refPath)
                
                LogMessage "IDW UPDATE: Checking reference: " & refFileName & " (path: " & refPath & ")"
                
                ' Check if we have a mapping for this file - first try full path, then try filename
                Dim newRefPath
                newRefPath = ""
                
                If g_CopiedFiles.Exists(refPath) Then
                    newRefPath = g_CopiedFiles.Item(refPath)
                Else
                    ' Fallback: lookup by filename
                    Dim origKey
                    For Each origKey In g_CopiedFiles.Keys
                        If LCase(GetFileNameFromPath(origKey)) = LCase(refFileName) Then
                            newRefPath = g_CopiedFiles.Item(origKey)
                            Exit For
                        End If
                    Next
                End If
                
                If newRefPath <> "" Then
                    LogMessage "IDW UPDATE: FOUND mapping - " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
                    
                    fd.ReplaceReference newRefPath
                    
                    If Err.Number = 0 Then
                        LogMessage "IDW UPDATE: SUCCESS - Updated reference"
                        updatedCount = updatedCount + 1
                    Else
                        LogMessage "IDW UPDATE: ERROR - " & Err.Description
                    End If
                    Err.Clear
                End If
            Next
            
            idwDoc.Save
            idwDoc.Close
            LogMessage "IDW UPDATE: Saved " & idwFiles.Item(idwPath) & " (" & updatedCount & " updated)"
        Else
            LogMessage "IDW UPDATE: Could not open " & idwFiles.Item(idwPath)
        End If
        Err.Clear
    Next
    
    ' Final cleanup - restore original settings
    On Error Resume Next
    invApp.FileOptions.ResolveFileOption = originalResolveMode
    invApp.SilentOperation = originalSilentMode
    LogMessage "IDW UPDATE (Inventor): Restored original settings"
    Err.Clear
End Sub

' === HELPER FUNCTIONS ===

Sub ScanIDWFilesForUpdate(folderObj, idwDict, fso)
    ' Recursively collect all IDW files from folder and subfolders for reference update
    ' Stores in dictionary: fullPath -> fileName (for logging)
    
    Dim file
    For Each file In folderObj.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            ' Skip OldVersions folder
            If InStr(LCase(folderObj.Path), "oldversions") = 0 Then
                idwDict.Add file.Path, file.Name
            End If
        End If
    Next
    
    ' Recurse into subfolders
    Dim subFolder
    For Each subFolder In folderObj.SubFolders
        ' Skip OldVersions folders
        If LCase(subFolder.Name) <> "oldversions" Then
            Call ScanIDWFilesForUpdate(subFolder, idwDict, fso)
        End If
    Next
End Sub

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Sub CreateFolderRecursive(folderPath, fso)
    ' Recursively create folder structure if it doesn't exist
    If fso.FolderExists(folderPath) Then Exit Sub
    
    Dim parentPath
    parentPath = fso.GetParentFolderName(folderPath)
    
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then
        Call CreateFolderRecursive(parentPath, fso)
    End If
    
    On Error Resume Next
    fso.CreateFolder folderPath
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not create folder: " & folderPath & " - " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Function GetDescriptionFromIProperty(doc)
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
        ClassifyByDescription = "B"  ' European I-beams (now in B group with UB/UC)
    Else
        ' Default - unclassified part
        ClassifyByDescription = "OTHER"
    End If
End Function

' === REGISTRY FUNCTIONS ===

Sub ScanRegistryForCounters(existingCounters, userPrefix)
    ' Scan Windows Registry for existing counters to continue numbering
    ' Much more reliable than file-based approaches
    ' Dynamically generates counter keys based on user's prefix

    LogMessage "REGISTRY SCAN: Scanning for existing counters with prefix: " & userPrefix

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
            LogMessage "REGISTRY SCAN: Found existing counter: " & keyName & " = " & currentValue
            foundCount = foundCount + 1
        Else
            LogMessage "REGISTRY SCAN: No existing counter for: " & keyName & " (will start from 1)"
        End If

        On Error GoTo 0
    Next

    If foundCount > 0 Then
        LogMessage "REGISTRY SCAN: Loaded " & foundCount & " existing counters from Registry"
    Else
        LogMessage "REGISTRY SCAN: No existing counters found in Registry - starting fresh"
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
    Next

    Err.Clear
End Function

' === LOGGING FUNCTIONS ===

Sub StartLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    Dim rootDir
    rootDir = fso.GetParentFolderName(scriptDir)
    Dim logsDir
    logsDir = rootDir & "\Logs"
    
    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder(logsDir)
    End If
    
    g_LogPath = logsDir & "\AssemblyCloner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFileNum = fso.CreateTextFile(g_LogPath, True)
End Sub

Sub StartDestinationLogging(destFolder)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(destFolder) Then Exit Sub
    
    g_DestLogPath = destFolder & "\AssemblyCloner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_DestLogFileNum = fso.CreateTextFile(g_DestLogPath, True)

    If Err.Number <> 0 Then
        LogMessage "DEST LOG: ERROR creating destination log - " & Err.Description
        Err.Clear
    Else
        LogMessage "DEST LOG: Writing destination log to " & g_DestLogPath
    End If
End Sub

Sub LogMessage(message)
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
    End If
    If Not IsEmpty(g_DestLogFileNum) Then
        g_DestLogFileNum.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
    End If
    WScript.Echo message
End Sub

Sub StopLogging()
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.Close
    End If
    If Not IsEmpty(g_DestLogFileNum) Then
        g_DestLogFileNum.Close
    End If
End Sub

Sub WriteMappingFile(destFolder)
    ' Write STEP_1_MAPPING.txt with all file mappings for future IDW updates
    ' Format: OLD_FILENAME|NEW_FILENAME (filename-only for cloner compatibility)
    ' Also writes STEP_1_MAPPING_FULLPATH.txt with full paths for Assembly Renamer compatibility
    
    On Error Resume Next
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' File 1: Filename-based mapping (for IDW updater filename matching)
    Dim mappingPath
    mappingPath = destFolder & "\STEP_1_MAPPING.txt"
    
    Dim mappingFile
    Set mappingFile = fso.CreateTextFile(mappingPath, True)
    
    mappingFile.WriteLine "# Filename-based mapping file for cloned folder: " & destFolder
    mappingFile.WriteLine "# Generated by Assembly Cloner on " & Now()
    mappingFile.WriteLine "# Format: OLD_FILENAME|NEW_FILENAME"
    mappingFile.WriteLine "# For filename-only matching"
    mappingFile.WriteLine ""
    
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim oldPath, newPath
        oldPath = key
        newPath = g_CopiedFiles.Item(key)
        
        Dim oldFileName, newFileName
        oldFileName = fso.GetFileName(oldPath)
        newFileName = fso.GetFileName(newPath)
        
        mappingFile.WriteLine LCase(oldFileName) & "|" & newFileName
    Next
    
    mappingFile.Close
    LogMessage "MAPPING: Wrote " & g_CopiedFiles.Count & " entries to STEP_1_MAPPING.txt"
    
    ' File 2: Full path mapping (for Assembly Renamer compatibility)
    Dim fullPathMappingPath
    fullPathMappingPath = destFolder & "\STEP_1_MAPPING_FULLPATH.txt"
    
    Dim fullPathFile
    Set fullPathFile = fso.CreateTextFile(fullPathMappingPath, True)
    
    fullPathFile.WriteLine "# Full path mapping file for cloned folder: " & destFolder
    fullPathFile.WriteLine "# Generated by Assembly Cloner on " & Now()
    fullPathFile.WriteLine "# Format: OLD_FULLPATH|NEW_FULLPATH"
    fullPathFile.WriteLine "# For exact path matching"
    fullPathFile.WriteLine ""
    
    For Each key In g_CopiedFiles.Keys
        oldPath = key
        newPath = g_CopiedFiles.Item(key)
        fullPathFile.WriteLine oldPath & "|" & newPath
    Next
    
    fullPathFile.Close
    LogMessage "MAPPING: Wrote " & g_CopiedFiles.Count & " entries to STEP_1_MAPPING_FULLPATH.txt"
    
    Err.Clear
End Sub

Sub ValidateCloneAndLog(sourceDir, destFolder)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    LogMessage "VALIDATE: Writing source and destination inventories..."
    
    Dim sourceFiles
    Set sourceFiles = CreateObject("Scripting.Dictionary")
    Dim destFiles
    Set destFiles = CreateObject("Scripting.Dictionary")
    
    Call BuildFileInventory(sourceDir, sourceFiles)
    Call BuildFileInventory(destFolder, destFiles)
    
    LogMessage "VALIDATE: Source file count = " & sourceFiles.Count
    LogMessage "VALIDATE: Destination file count = " & destFiles.Count
    
    Dim key
    For Each key In sourceFiles.Keys
        LogMessage "SOURCE FILE: " & key
    Next
    For Each key In destFiles.Keys
        LogMessage "DEST FILE: " & key
    Next
    
    LogMessage "VALIDATE: Checking copied mappings..."
    Dim missingCount
    missingCount = 0
    For Each key In g_CopiedFiles.Keys
        Dim copiedPath
        copiedPath = g_CopiedFiles.Item(key)
        If fso.FileExists(copiedPath) Then
            LogMessage "MAP OK: " & key & " -> " & copiedPath
        Else
            LogMessage "MAP MISSING: " & key & " -> " & copiedPath
            missingCount = missingCount + 1
        End If
    Next
    LogMessage "VALIDATE: Missing mapped files = " & missingCount
    
    Err.Clear
End Sub

Sub BuildFileInventory(folderPath, dict)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Dim file
    For Each file In folder.Files
        dict.Add file.Path, True
    Next
    
    Dim subFolder
    For Each subFolder In folder.SubFolders
        Call BuildFileInventory(subFolder.Path, dict)
    Next
    Err.Clear
End Sub

Sub UpdateIPropertiesForCopiedDocuments(invApp)
    ' Update iProperties for all copied documents (parts and assemblies)
    ' Replaces old names/suffixes with new names/suffixes in string properties
    ' CRITICAL: Also explicitly sets Part Number to match new filename
    
    On Error Resume Next
    
    LogMessage "IPROP: Starting iProperties update for " & g_CopiedFiles.Count & " copied files"
    
    Dim updatedCount
    updatedCount = 0
    
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim oldPath
        oldPath = key
        Dim newPath
        newPath = g_CopiedFiles.Item(key)
        
        Dim ext
        ext = LCase(Right(newPath, 4))
        
        If ext = ".ipt" Or ext = ".iam" Then
            ' Open the copied document
            Dim doc
            Set doc = invApp.Documents.Open(newPath, False)
            
            If Err.Number <> 0 Or doc Is Nothing Then
                LogMessage "IPROP: ERROR opening " & GetFileNameFromPath(newPath) & ": " & Err.Description
                Err.Clear
            Else
                ' Get old and new filenames without extension
                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")
                Dim oldFileName
                oldFileName = fso.GetBaseName(oldPath)
                Dim newFileName
                newFileName = fso.GetBaseName(newPath)
                
                ' Find differing suffix
                Dim minLen, diffIndex, oldSuffix, newSuffix
                minLen = IIf(Len(oldFileName) < Len(newFileName), Len(oldFileName), Len(newFileName))
                diffIndex = 0
                While diffIndex < minLen And Mid(oldFileName, diffIndex + 1, 1) = Mid(newFileName, diffIndex + 1, 1)
                    diffIndex = diffIndex + 1
                Wend
                oldSuffix = Mid(oldFileName, diffIndex + 1)
                newSuffix = Mid(newFileName, diffIndex + 1)
                
                Dim replaced
                replaced = False
                
                ' ===== CRITICAL FIX: Explicitly update Part Number iProperty =====
                Dim designProps
                Set designProps = doc.PropertySets.Item("Design Tracking Properties")
                If Err.Number = 0 And Not designProps Is Nothing Then
                    Dim partNumProp
                    Set partNumProp = designProps.Item("Part Number")
                    If Err.Number = 0 And Not partNumProp Is Nothing Then
                        Dim oldPartNum
                        oldPartNum = partNumProp.Value
                        partNumProp.Value = newFileName
                        LogMessage "IPROP: Part Number updated: " & oldPartNum & " -> " & newFileName
                        replaced = True
                    End If
                    Err.Clear
                End If
                Err.Clear
                ' ===== END Part Number Fix =====
                
                ' Iterate through all property sets
                Dim propSet
                For Each propSet In doc.PropertySets
                    Dim prop
                    For Each prop In propSet
                        If Not prop.Value Is Nothing And VarType(prop.Value) = vbString Then
                            Dim valueStr
                            valueStr = prop.Value
                            
                            ' For parts, skip Description (Comments property) and Part Number (already updated)
                            If ext = ".ipt" And (prop.Name = "Comments" Or prop.Name = "Part Number") Then
                                ' Skip - Description should not change, Part Number already set
                            Else
                                Dim newValue
                                newValue = valueStr
                                
                                ' Replace old suffix with new suffix
                                If oldSuffix <> "" Then
                                    newValue = Replace(newValue, oldSuffix, newSuffix, 1, -1, vbTextCompare)
                                End If
                                
                                ' Replace old filename with new filename
                                If oldFileName <> newFileName Then
                                    newValue = Replace(newValue, oldFileName, newFileName, 1, -1, vbTextCompare)
                                End If
                                
                                If newValue <> valueStr Then
                                    prop.Value = newValue
                                    replaced = True
                                End If
                            End If
                        End If
                    Next
                Next
                
                If replaced Then
                    LogMessage "IPROP: Updated iProperties for " & GetFileNameFromPath(newPath)
                    updatedCount = updatedCount + 1
                End If
                
                ' Save the document
                doc.Save
                
                If Err.Number <> 0 Then
                    LogMessage "IPROP: ERROR saving " & GetFileNameFromPath(newPath) & ": " & Err.Description
                    Err.Clear
                End If
                
                ' Close the document
                doc.Close False
                Err.Clear
            End If
        End If
    Next
    
    LogMessage "IPROP: Completed - updated " & updatedCount & " documents"
    Err.Clear
End Sub
