Option Explicit

' ==============================================================================
' STEP 2: IDW REFERENCE UPDATES - DESIGN ASSISTANT METHOD
' ==============================================================================
' This script:
' 1. USER MUST SELECT MAPPING FILE - NO AUTO-DETECTION
' 2. Finds IDW files ONLY in directories from the mapping file
' 3. Updates all IDW references using Design Assistant's exact method
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_ComprehensiveMapping ' Master mapping loaded from STEP 1
Dim g_MappingFilePath      ' Path to mapping file selected by user
Dim g_ProjectDir           ' Directory derived from mapping file location

Call STEP_2_IDW_UPDATES()

Sub STEP_2_IDW_UPDATES()
    Call StartLogging
    LogMessage "=== STEP 2: IDW REFERENCE UPDATES - DESIGN ASSISTANT METHOD ==="
    LogMessage "NO AUTO-DETECTION - USER MUST SELECT MAPPING FILE"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' =========================================================================
    ' STEP 0: GET MAPPING FILE FROM USER - THIS IS MANDATORY
    ' =========================================================================
    LogMessage "STEP 0: ASKING USER FOR MAPPING FILE (MANDATORY)"
    
    ' Use InputBox - simple and ALWAYS works
    g_MappingFilePath = InputBox("ENTER THE FULL PATH TO YOUR STEP_1_MAPPING.txt FILE:" & vbCrLf & vbCrLf & _
                                  "Example:" & vbCrLf & _
                                  "C:\Users\Quintin\Documents\Spectiv\3. Working\TEMP - RENAME TEST\000 Structure & Walkway before renaming\000 Structure & Walkway\STEP_1_MAPPING.txt" & vbCrLf & vbCrLf & _
                                  "Copy and paste the full path here:", _
                                  "SELECT MAPPING FILE - REQUIRED", _
                                  "")
    
    ' Check if user cancelled or entered empty
    If g_MappingFilePath = "" Then
        LogMessage "ERROR: User did not provide mapping file path"
        MsgBox "ERROR: YOU MUST PROVIDE A MAPPING FILE PATH!" & vbCrLf & vbCrLf & _
               "Without STEP_1_MAPPING.txt, this script cannot run." & vbCrLf & _
               "Exiting.", vbCritical, "No Mapping File"
        Call StopLogging
        Exit Sub
    End If
    
    ' Clean up the path (remove quotes if user copied with them)
    g_MappingFilePath = Replace(g_MappingFilePath, """", "")
    g_MappingFilePath = Trim(g_MappingFilePath)
    
    LogMessage "USER PROVIDED PATH: " & g_MappingFilePath
    
    ' Verify the file exists
    If Not fso.FileExists(g_MappingFilePath) Then
        LogMessage "ERROR: File does not exist: " & g_MappingFilePath
        MsgBox "ERROR: FILE NOT FOUND!" & vbCrLf & vbCrLf & _
               "Path: " & g_MappingFilePath & vbCrLf & vbCrLf & _
               "Please check the path and try again.", vbCritical, "File Not Found"
        Call StopLogging
        Exit Sub
    End If
    
    LogMessage "VERIFIED: Mapping file exists"
    
    ' Get the project directory from mapping file location
    g_ProjectDir = fso.GetParentFolderName(g_MappingFilePath)
    LogMessage "PROJECT DIRECTORY: " & g_ProjectDir
    
    ' =========================================================================
    ' CONFIRM WITH USER BEFORE PROCEEDING
    ' =========================================================================
    Dim confirmResult
    confirmResult = MsgBox("CONFIRM MAPPING FILE" & vbCrLf & vbCrLf & _
                           "File: " & g_MappingFilePath & vbCrLf & vbCrLf & _
                           "Project Directory: " & g_ProjectDir & vbCrLf & vbCrLf & _
                           "IDW files will ONLY be searched in directories" & vbCrLf & _
                           "that appear in this mapping file." & vbCrLf & vbCrLf & _
                           "Is this correct?", vbYesNo + vbQuestion, "Confirm Before Processing")
    
    If confirmResult = vbNo Then
        LogMessage "User cancelled after seeing confirmation"
        MsgBox "Cancelled. Please restart and enter the correct path.", vbInformation
        Call StopLogging
        Exit Sub
    End If
    
    ' =========================================================================
    ' CONNECT TO INVENTOR
    ' =========================================================================
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor first.", vbCritical
        Call StopLogging
        Exit Sub
    End If
    
    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear
    On Error GoTo 0
    
    ' =========================================================================
    ' LOAD MAPPING FILE
    ' =========================================================================
    Set g_ComprehensiveMapping = CreateObject("Scripting.Dictionary")
    
    LogMessage "LOADING MAPPING FILE: " & g_MappingFilePath
    
    If Not LoadMappingFromFile(g_MappingFilePath) Then
        MsgBox "ERROR: Could not load mapping file!" & vbCrLf & vbCrLf & _
               "File: " & g_MappingFilePath, vbCritical, "Load Failed"
        Call StopLogging
        Exit Sub
    End If
    
    LogMessage "LOADED: " & g_ComprehensiveMapping.Count & " mappings"
    
    If g_ComprehensiveMapping.Count = 0 Then
        MsgBox "ERROR: Mapping file is empty or invalid!" & vbCrLf & vbCrLf & _
               "File: " & g_MappingFilePath, vbCritical, "Empty Mapping"
        Call StopLogging
        Exit Sub
    End If
    
    ' =========================================================================
    ' FIND IDW FILES - ONLY FROM DIRECTORIES IN MAPPING FILE
    ' =========================================================================
    LogMessage "FINDING IDW FILES - ONLY FROM MAPPING DIRECTORIES"
    
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    
    ' Get unique directories from the mapping file
    Dim searchDirs
    Set searchDirs = GetDirectoriesFromMapping()
    
    LogMessage "WILL SEARCH " & searchDirs.Count & " DIRECTORIES FROM MAPPING"
    
    ' Search each directory for IDW files
    Dim dirKey
    For Each dirKey In searchDirs.Keys
        If fso.FolderExists(dirKey) Then
            LogMessage "SEARCHING: " & dirKey
            Call FindIDWsInDirectory(dirKey, idwFiles)
        Else
            LogMessage "SKIP (not found): " & dirKey
        End If
    Next
    
    LogMessage "FOUND: " & idwFiles.Count & " IDW files"
    
    If idwFiles.Count = 0 Then
        MsgBox "No IDW files found in mapping directories!" & vbCrLf & vbCrLf & _
               "Searched " & searchDirs.Count & " directories.", vbExclamation
        Call StopLogging
        Exit Sub
    End If
    
    ' =========================================================================
    ' UPDATE IDW FILES
    ' =========================================================================
    LogMessage "UPDATING IDW REFERENCES"
    
    Dim totalUpdates, totalErrors
    totalUpdates = 0
    totalErrors = 0
    
    Call UpdateAllIDWFilesWithDesignAssistantMethod(invApp, idwFiles, totalUpdates, totalErrors)
    
    ' =========================================================================
    ' COMPLETE
    ' =========================================================================
    LogMessage "=== STEP 2 COMPLETE ==="
    LogMessage "Mappings: " & g_ComprehensiveMapping.Count
    LogMessage "IDW Files: " & idwFiles.Count
    LogMessage "Updates: " & totalUpdates
    LogMessage "Errors: " & totalErrors
    
    Call StopLogging
    
    MsgBox "STEP 2 COMPLETE" & vbCrLf & vbCrLf & _
           "Mappings: " & g_ComprehensiveMapping.Count & vbCrLf & _
           "IDW Files: " & idwFiles.Count & vbCrLf & _
           "Updates: " & totalUpdates & vbCrLf & _
           "Errors: " & totalErrors & vbCrLf & vbCrLf & _
           "Log: " & g_LogPath, vbInformation, "Complete"
End Sub

' Load mapping from specified file path
Function LoadMappingFromFile(filePath)
    On Error Resume Next
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim mappingFile
    Set mappingFile = fso.OpenTextFile(filePath, 1)
    
    If Err.Number <> 0 Then
        LogMessage "ERROR opening file: " & Err.Description
        LoadMappingFromFile = False
        Exit Function
    End If
    
    Dim lineCount
    lineCount = 0
    
    Do While Not mappingFile.AtEndOfStream
        Dim line
        line = mappingFile.ReadLine
        
        ' Skip comments and empty lines
        If Left(Trim(line), 1) <> "#" And Trim(line) <> "" Then
            ' Parse format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description
            Dim parts
            parts = Split(line, "|")
            If UBound(parts) >= 3 Then
                Dim originalPath, newPath
                originalPath = Trim(parts(0))
                newPath = Trim(parts(1))
                
                If Not g_ComprehensiveMapping.Exists(originalPath) Then
                    g_ComprehensiveMapping.Add originalPath, newPath
                    lineCount = lineCount + 1
                End If
            End If
        End If
    Loop
    
    mappingFile.Close
    LogMessage "Loaded " & lineCount & " mappings from file"
    LoadMappingFromFile = True
End Function

' Get unique directories from mapping file - ONLY THE EXACT DIRECTORIES, NO PARENTS
' FIXED: Now gets directories from DESTINATION (new) paths, not source paths
Function GetDirectoriesFromMapping()
    Dim dirs
    Set dirs = CreateObject("Scripting.Dictionary")
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim keys
    keys = g_ComprehensiveMapping.Keys
    
    Dim i
    For i = 0 To UBound(keys)
        ' Get the NEW path (destination) from mapping value
        Dim newFilePath
        newFilePath = g_ComprehensiveMapping.Item(keys(i))
        
        Dim dirPath
        dirPath = fso.GetParentFolderName(newFilePath)
        
        If dirPath <> "" And Not dirs.Exists(dirPath) Then
            dirs.Add dirPath, True
        End If
        
        ' DO NOT ADD PARENT DIRECTORIES - THIS CAUSES THE BUG!
        ' Parent dirs can be "3. Working" which contains all other projects
    Next
    
    Set GetDirectoriesFromMapping = dirs
End Function

' Find IDW files in a single directory - NOT RECURSIVE (only direct files)
Sub FindIDWsInDirectory(dirPath, idwFiles)
    On Error Resume Next
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(dirPath) Then Exit Sub
    
    Dim folder
    Set folder = fso.GetFolder(dirPath)
    
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            If Not idwFiles.Exists(file.Path) Then
                idwFiles.Add file.Path, file.Name
                LogMessage "FOUND IDW: " & file.Name
            End If
        End If
    Next
    
    ' DO NOT RECURSE INTO SUBFOLDERS
    ' Each directory from the mapping is searched individually
    ' This prevents accidentally searching into other project folders
    
    Err.Clear
End Sub

Sub UpdateAllIDWFilesWithDesignAssistantMethod(invApp, idwFiles, ByRef totalUpdates, ByRef totalErrors)
    LogMessage "IDW: Processing " & idwFiles.Count & " IDW files with Design Assistant method"

    Dim idwKeys
    idwKeys = idwFiles.Keys
    totalUpdates = 0
    totalErrors = 0

    Dim i
    For i = 0 To UBound(idwKeys)
        Dim idwPath
        idwPath = idwKeys(i)
        Dim idwName
        idwName = idwFiles.Item(idwPath)

        LogMessage "IDW: Processing [" & (i + 1) & "/" & (UBound(idwKeys) + 1) & "] " & idwName

        Dim updates, errors
        Call UpdateSingleIDWWithDesignAssistantMethod(invApp, idwPath, updates, errors)
        totalUpdates = totalUpdates + updates
        totalErrors = totalErrors + errors
    Next

    LogMessage ""
    LogMessage "=== DESIGN ASSISTANT METHOD SUMMARY ==="
    LogMessage "Total IDW files processed: " & idwFiles.Count
    LogMessage "Total reference updates: " & totalUpdates
    LogMessage "Total errors: " & totalErrors
End Sub

Sub UpdateSingleIDWWithDesignAssistantMethod(invApp, idwPath, ByRef updateCount, ByRef errorCount)
    On Error Resume Next
    updateCount = 0
    errorCount = 0

    LogMessage "IDW: Opening " & GetFileNameFromPath(idwPath)

    ' Set resolve options to skip missing files
    Dim originalResolveMode
    On Error Resume Next
    originalResolveMode = invApp.FileOptions.ResolveFileOption
    invApp.FileOptions.ResolveFileOption = 54275 ' kSkipUnresolvedFiles
    
    ' CRITICAL: Enable silent operation to suppress ALL dialogs
    Dim originalSilentMode
    originalSilentMode = invApp.SilentOperation
    invApp.SilentOperation = True
    LogMessage "IDW: Enabled SilentOperation mode"
    Err.Clear

    Dim idwDoc
    Set idwDoc = invApp.Documents.Open(idwPath, False)

    If Err.Number <> 0 Then
        LogMessage "IDW: ERROR - Could not open: " & Err.Description

        ' Restore original resolve mode
        On Error Resume Next
        invApp.FileOptions.ResolveFileOption = originalResolveMode
        Err.Clear

        errorCount = 1
        Exit Sub
    End If

    LogMessage "IDW: Successfully opened IDW document"

    ' Restore original resolve mode
    On Error Resume Next
    invApp.FileOptions.ResolveFileOption = originalResolveMode
    Err.Clear

    ' === DESIGN ASSISTANT METHOD ===
    LogMessage "IDW: Using Design Assistant method - accessing ReferencedFileDescriptors"

    ' Access file descriptors - EXACTLY like Design Assistant
    Dim fileDescriptors
    Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
    LogMessage "IDW: Found " & fileDescriptors.Count & " referenced files"

    ' === CRITICAL FIX: Build live assembly reference map ===
    Dim assemblyDoc
    Set assemblyDoc = GetAssemblyForIDW(invApp, idwDoc, idwPath)
    Dim liveReferenceMap
    Set liveReferenceMap = CreateObject("Scripting.Dictionary")
    If Not (assemblyDoc Is Nothing) Then
        Call BuildLiveReferenceMap(assemblyDoc, liveReferenceMap)
        LogMessage "IDW: Built live reference map with " & liveReferenceMap.Count & " entries from assembly"
    Else
        LogMessage "IDW: WARNING - Could not load parent assembly, will use mapping file only"
    End If

    ' Update each reference using Design Assistant method
    Dim i
    For i = 1 To fileDescriptors.Count
        Dim fd
        Set fd = fileDescriptors.Item(i)

        Dim currentFullPath
        currentFullPath = fd.FullFileName
        Dim currentFileName
        currentFileName = GetFileNameFromPath(currentFullPath)
        LogMessage "IDW:   Processing reference: " & currentFileName
        LogMessage "IDW:   Full path from IDW: " & currentFullPath

        ' Check if we have a full path mapping for this file
        Dim mappingFound
        Dim newPath
        mappingFound = False

        ' STRATEGY: Try multiple methods to find the new path
        ' Method 1: Exact path match in mapping (fastest)
        If g_ComprehensiveMapping.Exists(currentFullPath) Then
            newPath = g_ComprehensiveMapping.Item(currentFullPath)
            mappingFound = True
            LogMessage "IDW:   Method 1: Found exact path match in mapping"
        End If

        ' Method 2: Resolve the actual current file path (handles renaming)
        If Not mappingFound Then
            LogMessage "IDW:   Method 2: Resolving actual file path"
            Dim resolvedPath
            resolvedPath = ResolveActualFilePath(currentFullPath)

            If resolvedPath <> "" Then
                LogMessage "IDW:   Method 2: Resolved to: " & GetFileNameFromPath(resolvedPath)

                ' Now check if this resolved path has a mapping
                If g_ComprehensiveMapping.Exists(resolvedPath) Then
                    newPath = g_ComprehensiveMapping.Item(resolvedPath)
                    mappingFound = True
                    LogMessage "IDW:   Method 2: Found mapping for resolved path"
                End If
            Else
                LogMessage "IDW:   Method 2: Could not resolve actual file path"
            End If
        End If

        ' Method 3: Try assembly live reference tracing
        If Not mappingFound And liveReferenceMap.Count > 0 Then
            Dim fso3
            Set fso3 = CreateObject("Scripting.FileSystemObject")
            LogMessage "IDW:   Method 3: Searching live assembly references"

            ' Search for this file in the live assembly by filename
            Dim key
            For Each key In liveReferenceMap.Keys
                Dim keyFileName
                keyFileName = fso3.GetFileName(key)
                If LCase(keyFileName) = LCase(currentFileName) Then
                    Dim currentAssemblyPath
                    currentAssemblyPath = liveReferenceMap.Item(key)
                    LogMessage "IDW:   Method 3: Found in assembly at: " & currentAssemblyPath

                    ' Now find this current assembly path in the mapping
                    If g_ComprehensiveMapping.Exists(currentAssemblyPath) Then
                        newPath = g_ComprehensiveMapping.Item(currentAssemblyPath)
                        mappingFound = True
                        LogMessage "IDW:   Method 3: Traced through assembly: " & currentFileName & " -> " & fso3.GetFileName(currentAssemblyPath) & " -> " & fso3.GetFileName(newPath)
                        Exit For
                    End If
                End If
            Next
        End If

        ' Method 4: Smart filename + path structure matching
        If Not mappingFound Then
            LogMessage "IDW:   Method 4: Trying smart path structure matching"
            newPath = FindSmartMapping(currentFullPath, currentFileName)
            If newPath <> "" Then
                mappingFound = True
                LogMessage "IDW:   Method 4: Found smart match"
            End If
        End If

        If mappingFound Then
            Dim newFileName
            newFileName = GetFileNameFromPath(newPath)

            ' Check if this is an identity mapping (old = new)
            If currentFullPath = newPath Then
                LogMessage "IDW:     (Already correct: " & currentFileName & " - no update needed)"
            Else
                LogMessage "IDW:     UPDATING: " & currentFileName & " -> " & newFileName
                LogMessage "IDW:     NEW PATH: " & newPath

                ' Check if new file exists
                Dim fso2
                Set fso2 = CreateObject("Scripting.FileSystemObject")
                If fso2.FileExists(newPath) Then
                    ' THIS IS THE MAGIC METHOD DESIGN ASSISTANT USES!
                    Err.Clear
                    fd.ReplaceReference newPath

                    If Err.Number = 0 Then
                        LogMessage "IDW:     ✓ SUCCESS - Reference updated using Design Assistant method"
                        updateCount = updateCount + 1
                    Else
                        LogMessage "IDW:     ✗ ERROR - ReplaceReference failed: " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                Else
                    LogMessage "IDW:     ✗ ERROR - New file doesn't exist: " & newPath
                    errorCount = errorCount + 1
                End If
            End If
        Else
            LogMessage "IDW:     (No mapping found - keeping current reference)"
        End If
    Next

    ' Save IDW if updates were made
    If updateCount > 0 Then
        LogMessage "IDW: Saving with " & updateCount & " updates using Design Assistant method..."
        idwDoc.Save2(True)
        LogMessage "IDW: SUCCESS - Saved with " & updateCount & " updates"
    Else
        LogMessage "IDW: No updates made"
    End If

    ' Close assembly if we opened it
    If Not (assemblyDoc Is Nothing) Then
        assemblyDoc.Close
    End If

    ' Close document
    idwDoc.Close
    
    ' Restore original settings
    invApp.FileOptions.ResolveFileOption = originalResolveMode
    invApp.SilentOperation = originalSilentMode
    Err.Clear
End Sub

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function ResolveActualFilePath(idwReferencePath)
    ' Given a path that an IDW thinks a file is at,
    ' find where the file ACTUALLY is now (after renaming)
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' If the file exists at the path the IDW says, return it
    If fso.FileExists(idwReferencePath) Then
        ResolveActualFilePath = idwReferencePath
        Exit Function
    End If

    ' File doesn't exist - it's been renamed
    ' Search the same directory for renamed files in our mapping
    Dim currentDir
    currentDir = fso.GetParentFolderName(idwReferencePath)
    Dim originalFileName
    originalFileName = fso.GetFileName(idwReferencePath)

    LogMessage "RESOLVE: Looking for renamed version of: " & originalFileName
    LogMessage "RESOLVE: Searching directory: " & currentDir

    If Not fso.FolderExists(currentDir) Then
        LogMessage "RESOLVE: ERROR - Directory doesn't exist: " & currentDir
        ResolveActualFilePath = ""
        Exit Function
    End If

    ' Look through all .ipt files in the directory
    Dim folder
    Set folder = fso.GetFolder(currentDir)
    Dim file
    Dim candidateFiles
    Set candidateFiles = CreateObject("Scripting.Dictionary")

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".ipt" Then
            ' Check if this file appears in our mapping as an "original" (key)
            ' If it does, it means this file was renamed and might be our target
            If g_ComprehensiveMapping.Exists(file.Path) Then
                candidateFiles.Add file.Path, file.Name
                LogMessage "RESOLVE: Candidate: " & file.Name & " (found in mapping as original)"
            End If
        End If
    Next

    ' If we found exactly one renamed file in this directory, use it
    If candidateFiles.Count = 1 Then
        Dim keys
        keys = candidateFiles.Keys
        ResolveActualFilePath = keys(0)
        LogMessage "RESOLVE: Found unique renamed file: " & candidateFiles.Item(keys(0))
        Exit Function
    ElseIf candidateFiles.Count > 1 Then
        ' Multiple candidates - try to match by looking at file size, mod date, etc.
        ' For now, just log and return the first one
        LogMessage "RESOLVE: WARNING - Multiple candidates found (" & candidateFiles.Count & "), using first"
        keys = candidateFiles.Keys
        ResolveActualFilePath = keys(0)
        Exit Function
    Else
        LogMessage "RESOLVE: No renamed file found in mapping for: " & originalFileName
        ResolveActualFilePath = ""
    End If
End Function

Function GetAssemblyForIDW(invApp, idwDoc, idwPath)
    ' Try to find and open the parent assembly for this IDW
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the directory containing the IDW
    Dim idwDir
    idwDir = fso.GetParentFolderName(idwPath)

    ' Look for .iam files in the same directory
    Dim folder
    Set folder = fso.GetFolder(idwDir)

    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".iam" Then
            LogMessage "IDW: Found potential assembly: " & file.Name
            ' Try to open it
            Err.Clear
            Dim asmDoc
            Set asmDoc = invApp.Documents.Open(file.Path, False)
            If Err.Number = 0 And Not (asmDoc Is Nothing) Then
                LogMessage "IDW: Successfully opened assembly: " & file.Name
                Set GetAssemblyForIDW = asmDoc
                Exit Function
            End If
        End If
    Next

    ' If no assembly found in same directory, try looking for already open assemblies
    Dim doc
    For Each doc In invApp.Documents
        If doc.DocumentType = 12290 Then ' kAssemblyDocumentObject
            Dim docDir
            docDir = fso.GetParentFolderName(doc.FullFileName)
            If LCase(docDir) = LCase(idwDir) Then
                LogMessage "IDW: Using already open assembly: " & doc.DisplayName
                Set GetAssemblyForIDW = doc
                Exit Function
            End If
        End If
    Next

    Set GetAssemblyForIDW = Nothing
    Err.Clear
End Function

Sub BuildLiveReferenceMap(assemblyDoc, referenceMap)
    ' Build a map of original names -> current paths from the live assembly
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Walk through all occurrences in the assembly
    Dim occurrences
    Set occurrences = assemblyDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Get the current file descriptor
        Dim fd
        Set fd = occ.ReferencedFileDescriptor

        Dim currentPath
        currentPath = fd.FullFileName

        Dim fileName
        fileName = fso.GetFileName(currentPath)

        ' Extract original name from current name if possible
        ' Examples:
        '   NCRH01-000-PL174.ipt might have been Part1.ipt
        '   SCRH01-752-PL1.ipt might have been NCRH01-000-PL174.ipt
        ' We need to create multiple lookups

        ' Store current path under its current filename
        If Not referenceMap.Exists(fileName) Then
            referenceMap.Add fileName, currentPath
            LogMessage "LIVE-REF: " & fileName & " -> " & currentPath
        End If

        ' Also try to extract and store any "embedded" original names
        ' by looking for patterns in iProperties or part number
        Dim partDoc
        On Error Resume Next
        Set partDoc = occ.Definition.Document
        If Not (partDoc Is Nothing) Then
            Dim propSet
            Set propSet = partDoc.PropertySets.Item("Design Tracking Properties")
            If Not (propSet Is Nothing) Then
                Dim prop
                ' Try to get Part Number which might contain original name
                Set prop = propSet.Item("Part Number")
                If Not (prop Is Nothing) And prop.Value <> "" Then
                    Dim partNumber
                    partNumber = Trim(prop.Value)
                    ' If part number is different from filename, add it too
                    If partNumber <> fileName And partNumber <> "" Then
                        If Not referenceMap.Exists(partNumber) Then
                            referenceMap.Add partNumber, currentPath
                            LogMessage "LIVE-REF: " & partNumber & " [from iProperty] -> " & currentPath
                        End If
                    End If
                End If
            End If
        End If
        Err.Clear
    Next

    LogMessage "LIVE-REF: Built reference map with " & referenceMap.Count & " entries"
End Sub

Function FindSmartMapping(idwPath, fileName)
    ' Smart mapping lookup for when directory paths don't match exactly
    ' Compares filename + relative path structure to find matches
    LogMessage "IDW:   Smart matching for: " & fileName

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Extract relative path structure from IDW path (last 3 folder levels)
    Dim idwRelativeStructure
    idwRelativeStructure = ExtractRelativeStructure(idwPath)
    LogMessage "IDW:   IDW relative structure: " & idwRelativeStructure

    ' Search through all mappings for filename + structure matches
    Dim mappingKeys
    mappingKeys = g_ComprehensiveMapping.Keys

    Dim i
    For i = 0 To UBound(mappingKeys)
        Dim mappingOriginalPath
        mappingOriginalPath = mappingKeys(i)
        Dim mappingFileName
        mappingFileName = fso.GetFileName(mappingOriginalPath)

        ' Check if filename matches
        If LCase(mappingFileName) = LCase(fileName) Then
            ' Check if relative structure also matches
            Dim mappingRelativeStructure
            mappingRelativeStructure = ExtractRelativeStructure(mappingOriginalPath)
            LogMessage "IDW:   Checking mapping structure: " & mappingRelativeStructure

            If LCase(mappingRelativeStructure) = LCase(idwRelativeStructure) Then
                LogMessage "IDW:   SMART MATCH FOUND: " & mappingFileName & " with matching structure"
                FindSmartMapping = g_ComprehensiveMapping.Item(mappingOriginalPath)
                Exit Function
            End If
        End If
    Next

    LogMessage "IDW:   No smart match found for: " & fileName
    FindSmartMapping = ""
End Function

Function ExtractRelativeStructure(fullPath)
    ' Extract last 3 directory levels + filename for comparison
    ' Example: "D:\Path\To\Project\Head Chute\Launder-1\Liners\Liner1.ipt"
    '       -> "Head Chute\Launder-1\Liners\Liner1.ipt"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim pathParts
    pathParts = Split(fullPath, "\")

    If UBound(pathParts) >= 3 Then
        ' Take last 4 parts (3 folders + filename)
        Dim result
        result = pathParts(UBound(pathParts) - 3) & "\" & _
                pathParts(UBound(pathParts) - 2) & "\" & _
                pathParts(UBound(pathParts) - 1) & "\" & _
                pathParts(UBound(pathParts))
        ExtractRelativeStructure = result
    Else
        ' If path too short, use as-is
        ExtractRelativeStructure = fullPath
    End If
End Function

Function GetScriptDirectory()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetScriptDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
End Function

Sub BrowseForIDWFiles(invApp, idwFiles)
    ' Allow user to manually browse for IDW files when automatic detection fails
    LogMessage "BROWSE: Starting manual IDW file selection"

    Dim continueAdding
    continueAdding = True

    Do While continueAdding
        ' Create file dialog for IDW selection
        Dim shell
        Set shell = CreateObject("Shell.Application")

        ' Show message about what to do
        Dim instruction
        instruction = MsgBox("Manual IDW File Selection" & vbCrLf & vbCrLf & _
                           "Click OK to browse for an IDW file." & vbCrLf & _
                           "Select IDW files that need reference updates." & vbCrLf & vbCrLf & _
                           "Currently selected: " & idwFiles.Count & " files", _
                           vbOKCancel + vbInformation, "Browse for IDW Files")

        If instruction = vbCancel Then
            LogMessage "BROWSE: User cancelled file selection"
            Exit Do
        End If

        ' Use Inventor's file dialog (most reliable)
        Dim selectedPath
        selectedPath = ShowInventorFileDialog(invApp, "Select IDW File to Update")

        If selectedPath = "" Then
            LogMessage "BROWSE: User finished file selection"
            Exit Do
        End If

        ' Validate file exists and is IDW
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        If Not fso.FileExists(selectedPath) Then
            MsgBox "File not found: " & selectedPath, vbExclamation
            LogMessage "BROWSE: File not found - " & selectedPath
        ElseIf LCase(Right(selectedPath, 4)) <> ".idw" Then
            MsgBox "File is not an IDW file: " & selectedPath, vbExclamation
            LogMessage "BROWSE: Not an IDW file - " & selectedPath
        ElseIf idwFiles.Exists(selectedPath) Then
            MsgBox "File already selected: " & fso.GetFileName(selectedPath), vbInformation
            LogMessage "BROWSE: File already selected - " & selectedPath
        Else
            ' Add to collection
            idwFiles.Add selectedPath, fso.GetFileName(selectedPath)
            LogMessage "BROWSE: Added IDW file - " & selectedPath

            ' Ask if user wants to add more files
            Dim addMore
            addMore = MsgBox("Added: " & fso.GetFileName(selectedPath) & vbCrLf & vbCrLf & _
                           "Total selected: " & idwFiles.Count & " files" & vbCrLf & vbCrLf & _
                           "Add another IDW file?" & vbCrLf & vbCrLf & _
                           "YES = Add another file" & vbCrLf & _
                           "NO = Proceed with selected files", _
                           vbYesNo + vbQuestion, "Add More Files?")

            If addMore = vbNo Then
                LogMessage "BROWSE: User finished with " & idwFiles.Count & " files selected"
                Exit Do
            End If
        End If
    Loop

    LogMessage "BROWSE: Manual file selection completed - " & idwFiles.Count & " files selected"
End Sub

Function ShowInventorFileDialog(invApp, title)
    ' Simple file browser using PowerShell - DIRECT FILE SELECTION
    On Error Resume Next

    LogMessage "BROWSE: Creating PowerShell file browser dialog"

    ' Use PowerShell to show a proper Windows file dialog
    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")

    cmd = "powershell -WindowStyle Hidden -Command """ & _
          "Add-Type -AssemblyName System.Windows.Forms;" & _
          "$dialog = New-Object System.Windows.Forms.OpenFileDialog;" & _
          "$dialog.Title = 'Select IDW File';" & _
          "$dialog.Filter = 'Inventor Drawing Files (*.idw)|*.idw|All Files (*.*)|*.*';" & _
          "$dialog.FilterIndex = 1;" & _
          "$dialog.Multiselect = $false;" & _
          "if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {" & _
          "    Write-Output $dialog.FileName" & _
          "} else {" & _
          "    Write-Output 'CANCELLED'" & _
          "}" & _
          """"

    result = shell.Exec(cmd).StdOut.ReadAll()
    result = Trim(result)

    If result <> "" And result <> "CANCELLED" Then
        ShowInventorFileDialog = result
        LogMessage "BROWSE: User selected file - " & result
    Else
        ShowInventorFileDialog = ""
        LogMessage "BROWSE: User cancelled file selection"
    End If

    Err.Clear
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
    g_LogPath = logsDir & "\Step2_IDW_Updates_DesignAssistant_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
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