Option Explicit

' ==============================================================================
' RECURSIVE IDW UPDATER - AGGREGATES ALL MAPPINGS
' ==============================================================================
' This script:
' 1. Recursively scans ENTIRE folder structure for ALL STEP_1_MAPPING.txt files
' 2. Aggregates ALL mappings into one comprehensive dictionary
' 3. Recursively finds ALL .idw files
' 4. For each IDW, updates references using aggregated mappings
' 5. Generates detailed report showing which IDWs used which mapping file(s)
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_ReportFileNum
Dim g_ReportPath
Dim g_ComprehensiveMapping ' Master aggregated mapping from ALL mapping files
Dim g_FilenameMapping ' Filename-based mapping for cloner compatibility
Dim g_IDWToMappingReport ' Maps IDW path -> List of mapping files that were used

Call RECURSIVE_IDW_UPDATER()

Sub RECURSIVE_IDW_UPDATER()
    ' Call StartLogging
    LogMessage "=== RECURSIVE IDW UPDATER - AGGREGATED MAPPINGS ==="
    LogMessage "Purpose: Fix scattered mapping issue by aggregating ALL mappings"

    Dim result
    result = MsgBox("RECURSIVE IDW UPDATER" & vbCrLf & vbCrLf & _
                    "NEW: Aggregates ALL STEP_1_MAPPING.txt files!" & vbCrLf & _
                    "1. Recursively scan ALL folders for mapping files" & vbCrLf & _
                    "2. Aggregate ALL mappings into comprehensive dictionary" & vbCrLf & _
                    "3. Recursively find ALL .idw files" & vbCrLf & _
                    "4. Update all IDWs using aggregated mappings" & vbCrLf & _
                    "5. Generate detailed report" & vbCrLf & vbCrLf & _
                    "Make sure Inventor is running!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Recursive IDW Updater")

    If result = vbNo Then
        LogMessage "User cancelled"
        Exit Sub
    End If

    ' Initialize collections
    Set g_ComprehensiveMapping = CreateObject("Scripting.Dictionary")
    Set g_FilenameMapping = CreateObject("Scripting.Dictionary")
    Set g_IDWToMappingReport = CreateObject("Scripting.Dictionary")

    ' Connect to Inventor (invisible mode for silent operation)
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "Inventor not running, starting in invisible mode..."
        Set invApp = CreateObject("Inventor.Application")
        If Not invApp Is Nothing Then
            invApp.Visible = False
            LogMessage "Started Inventor in invisible mode"
        Else
            LogMessage "ERROR: Could not start Inventor!"
            MsgBox "ERROR: Could not start Inventor!", vbCritical
            Exit Sub
        End If
    Else
        LogMessage "Connected to running Inventor instance"
    End If
    
    ' Store and set options to suppress dialogs (like Assembly Cloner)
    Dim originalSilent, originalResolve
    originalSilent = invApp.SilentOperation
    originalResolve = invApp.FileOptions.ResolveFileOption
    
    invApp.SilentOperation = True
    invApp.FileOptions.ResolveFileOption = 54275 ' kSkipUnresolvedFiles

    Err.Clear

    ' Get root directory to scan
    Dim rootDir
    rootDir = GetRootDirectory()

    If rootDir = "" Then
        MsgBox "ERROR: Could not determine root directory!" & vbCrLf & vbCrLf & _
               "Please make sure you have an Inventor document open.", vbCritical
        Exit Sub
    End If

    LogMessage "ROOT: " & rootDir

    ' Step 1: Recursively scan for ALL mapping files
    LogMessage ""
    LogMessage "STEP 1: SCANNING FOR ALL MAPPING FILES"
    Dim mappingFiles
    Set mappingFiles = CreateObject("Scripting.Dictionary")
    Call FindAllMappingFiles(rootDir, mappingFiles)

    If mappingFiles.Count = 0 Then
        LogMessage "ERROR: No mapping files found!"
        MsgBox "ERROR: No STEP_1_MAPPING.txt files found!" & vbCrLf & _
               "Searched in: " & rootDir, vbCritical, "No Mappings Found"
        Exit Sub
    End If

    LogMessage "FOUND " & mappingFiles.Count & "_MAPPING FILES"

    ' Step 2: Aggregate ALL mappings
    LogMessage ""
    LogMessage "STEP 2: AGGREGATING ALL MAPPINGS"
    Call AggregateAllMappings(mappingFiles)

    LogMessage "AGGREGATED " & g_ComprehensiveMapping.Count & " PATH MAPPINGS AND " & g_FilenameMapping.Count & " FILENAME MAPPINGS FROM ALL FILES"

    ' Step 3: Recursively find ALL IDW files
    LogMessage ""
    LogMessage "STEP 3: FINDING ALL IDW FILES"
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    Call FindAllIDWFiles(rootDir, idwFiles)

    If idwFiles.Count = 0 Then
        LogMessage "ERROR: No IDW files found!"
        MsgBox "ERROR: No .idw files found!" & vbCrLf & _
               "Searched in: " & rootDir, vbCritical, "No IDWs Found"
        Exit Sub
    End If

    LogMessage "FOUND " & idwFiles.Count & " IDW FILES TO PROCESS"

    ' Step 4: Update all IDWs using aggregated mappings
    LogMessage ""
    LogMessage "STEP 4: UPDATING ALL IDWs WITH AGGREGATED MAPPINGS"
    LogMessage "About to call UpdateAllIDWFiles with " & idwFiles.Count & " files"
    Dim totalUpdates, totalErrors
    Call UpdateAllIDWFiles(invApp, idwFiles, totalUpdates, totalErrors)
    LogMessage "UpdateAllIDWFiles completed - updates: " & totalUpdates & ", errors: " & totalErrors

    ' Step 5: Generate detailed report
    LogMessage ""
    LogMessage "STEP 5: GENERATING DETAILED REPORT"
    Call GenerateMappingReport(mappingFiles, idwFiles)

    LogMessage ""
    LogMessage "=== RECURSIVE IDW UPDATER COMPLETED ==="
    Call StopLogging

    ' Show completion message
    Dim resultMsg
    resultMsg = "RECURSIVE IDW UPDATER COMPLETED!" & vbCrLf & vbCrLf & _
           "Mapping files found: " & mappingFiles.Count & vbCrLf & _
           "Path mappings aggregated: " & g_ComprehensiveMapping.Count & vbCrLf & _
           "Filename mappings aggregated: " & g_FilenameMapping.Count & vbCrLf & _
           "IDW files processed: " & idwFiles.Count & vbCrLf & _
           "Reference updates: " & totalUpdates & vbCrLf & _
           "Errors: " & totalErrors & vbCrLf & vbCrLf & _
           "Log: " & g_LogPath & vbCrLf & _
           "Report: " & g_ReportPath

    MsgBox resultMsg, vbInformation, "Recursive IDW Updater - Complete"
End Sub

Function GetRootDirectory()
    ' Get root directory from active document or ask user
    On Error Resume Next

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If Not (activeDoc Is Nothing) Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        GetRootDirectory = fso.GetParentFolderName(activeDoc.FullFileName)
        LogMessage "ROOT: Found from active document: " & activeDoc.DisplayName
        Exit Function
    End If

    ' If no active document, ask user to browse
    GetRootDirectory = BrowseForFolder("Select root folder to scan for mapping files and IDWs")
End Function

Sub FindAllMappingFiles(rootDir, mappingFiles)
    ' Recursively find ALL STEP_1_MAPPING.txt files
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder
    Set folder = fso.GetFolder(rootDir)

    LogMessage "SCANNING: " & rootDir

    ' Check for mapping file in current directory
    Dim mappingFilePath
    mappingFilePath = rootDir & "\STEP_1_MAPPING.txt"

    If fso.FileExists(mappingFilePath) Then
        mappingFiles.Add mappingFilePath, True
        LogMessage "FOUND: " & mappingFilePath
    End If

    ' Also check for filename mapping file
    Dim filenameMappingPath
    filenameMappingPath = rootDir & "\STEP_1_MAPPING_FILENAME.txt"

    If fso.FileExists(filenameMappingPath) Then
        mappingFiles.Add filenameMappingPath, True
        LogMessage "FOUND: " & filenameMappingPath
    End If
    
    ' Also check for full path mapping file (generated by Assembly Cloner)
    Dim fullPathMappingPath
    fullPathMappingPath = rootDir & "\STEP_1_MAPPING_FULLPATH.txt"

    If fso.FileExists(fullPathMappingPath) Then
        mappingFiles.Add fullPathMappingPath, True
        LogMessage "FOUND: " & fullPathMappingPath
    End If

    ' Recursively process subdirectories
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" And _
           LCase(subFolder.Name) <> "temp" And _
           LCase(subFolder.Name) <> "$recycle.bin" Then
            Call FindAllMappingFiles(subFolder.Path, mappingFiles)
        End If
    Next
End Sub

Sub AggregateAllMappings(mappingFiles)
    ' Load all mapping files into comprehensive dictionary
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim totalMappings
    totalMappings = 0
    Dim filenameMappings
    filenameMappings = 0

    Dim mappingKeys
    mappingKeys = mappingFiles.Keys

    Dim i
    For i = 0 To UBound(mappingKeys)
        Dim mappingPath
        mappingPath = mappingKeys(i)

        LogMessage "LOADING: " & mappingPath

        Dim mappingFile
        Set mappingFile = fso.OpenTextFile(mappingPath, 1)
        Dim fileMappings
        fileMappings = 0
        Dim fileFilenameMappings
        fileFilenameMappings = 0

        Do While Not mappingFile.AtEndOfStream
            Dim line
            line = mappingFile.ReadLine

            ' Skip comments and empty lines
            If Left(Trim(line), 1) <> "#" And Trim(line) <> "" Then
                ' Parse format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description
                Dim parts
                parts = Split(line, "|")

                If UBound(parts) >= 3 Then
                    ' Full format: OriginalPath|NewPath|OriginalFile|NewFile|...
                    Dim originalPath, newPath, originalFile, newFile
                    originalPath = Trim(parts(0))
                    newPath = Trim(parts(1))
                    originalFile = Trim(parts(2))
                    newFile = Trim(parts(3))

                    ' Add to comprehensive mapping
                    If Not g_ComprehensiveMapping.Exists(originalPath) Then
                        g_ComprehensiveMapping.Add originalPath, newPath
                        fileMappings = fileMappings + 1
                        totalMappings = totalMappings + 1
                    Else
                        LogMessage "DUPLICATE: " & originalFile & " (already exists in other mapping file)"
                    End If
                ElseIf UBound(parts) = 1 Then
                    ' Two-part format - could be filename-only or full-path
                    Dim part0, part1
                    part0 = Trim(parts(0))
                    part1 = Trim(parts(1))
                    
                    ' Check if it looks like a full path (contains backslash or drive letter)
                    If InStr(part0, "\") > 0 Or (Len(part0) > 2 And Mid(part0, 2, 1) = ":") Then
                        ' Full path format: OLD_FULLPATH|NEW_FULLPATH
                        If Not g_ComprehensiveMapping.Exists(part0) Then
                            g_ComprehensiveMapping.Add part0, part1
                            fileMappings = fileMappings + 1
                            totalMappings = totalMappings + 1
                        End If
                    Else
                        ' Filename-only format: OLD_FILENAME|NEW_FILENAME
                        If Not g_FilenameMapping.Exists(LCase(part0)) Then
                            g_FilenameMapping.Add LCase(part0), part1
                            fileFilenameMappings = fileFilenameMappings + 1
                            filenameMappings = filenameMappings + 1
                        End If
                    End If
                End If
            End If
        Loop

        mappingFile.Close
        LogMessage "LOADED " & fileMappings & " path mappings and " & fileFilenameMappings & " filename mappings from this file"
    Next

    LogMessage "TOTAL PATH MAPPINGS AGGREGATED: " & totalMappings
    LogMessage "TOTAL FILENAME MAPPINGS AGGREGATED: " & filenameMappings
End Sub

Sub FindAllIDWFiles(dirPath, idwFiles)
    ' Recursively find ALL .idw files
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dirPath) Then
        Exit Sub
    End If

    Dim folder
    Set folder = fso.GetFolder(dirPath)

    ' Process files in current directory
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            If Not idwFiles.Exists(file.Path) Then
                idwFiles.Add file.Path, file.Name
                LogMessage "FOUND IDW: " & file.Path
            End If
        End If
    Next

    ' Process subdirectories recursively
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" And _
           LCase(subFolder.Name) <> "temp" And _
           LCase(subFolder.Name) <> "$recycle.bin" Then
            Call FindAllIDWFiles(subFolder.Path, idwFiles)
        End If
    Next
End Sub

Function GetFreshInventorConnection()
    ' Get a fresh connection to Inventor for each file to avoid COM issues
    On Error Resume Next
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")
    
    If Err.Number <> 0 Or invApp Is Nothing Then
        Set invApp = CreateObject("Inventor.Application")
        If Not invApp Is Nothing Then
            invApp.Visible = False
        End If
    Else
        invApp.Visible = False
    End If
    
    Err.Clear
    On Error GoTo 0
    Set GetFreshInventorConnection = invApp
End Function

Sub UpdateAllIDWFiles(invApp, idwFiles, ByRef totalUpdates, ByRef totalErrors)
    totalUpdates = 0
    totalErrors = 0

    Dim idwKeys
    idwKeys = idwFiles.Keys

    Dim i
    For i = 0 To UBound(idwKeys)
        Dim idwPath
        idwPath = idwKeys(i)
        Dim idwName
        idwName = idwFiles.Item(idwPath)

        LogMessage ""
        LogMessage "IDW [" & (i + 1) & "/" & (UBound(idwKeys) + 1) & "]: " & idwName

        On Error Resume Next
        Dim updates, errors, mappingsUsed
        Call UpdateSingleIDW(invApp, idwPath, updates, errors, mappingsUsed)
        
        If Err.Number <> 0 Then
            LogMessage "IDW: ERROR - UpdateSingleIDW crashed: " & Err.Description
            errors = 1
            updates = 0
            mappingsUsed = ""
            Err.Clear
        End If
        On Error GoTo 0

        totalUpdates = totalUpdates + updates
        totalErrors = totalErrors + errors

        ' Track which mapping files were used for this IDW
        If mappingsUsed <> "" Then
            g_IDWToMappingReport.Add idwPath, mappingsUsed
        End If
    Next

    LogMessage ""
    LogMessage "=== UPDATE SUMMARY ==="
    LogMessage "Total IDW files processed: " & idwFiles.Count
    LogMessage "Total reference updates: " & totalUpdates
    LogMessage "Total errors: " & totalErrors
End Sub

Sub UpdateSingleIDW(invApp, idwPath, ByRef updateCount, ByRef errorCount, ByRef mappingsUsed)
    ' Update a single IDW file using aggregated mapping with Inventor
    ' Uses same approach as Assembly Cloner for stability
    On Error Resume Next
    updateCount = 0
    errorCount = 0
    mappingsUsed = ""

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    LogMessage "IDW: Opening " & GetFileNameFromPath(idwPath)

    ' Check if file exists
    If Not fso.FileExists(idwPath) Then
        LogMessage "IDW: ERROR - File does not exist: " & idwPath
        errorCount = 1
        Exit Sub
    End If

    ' Close all documents first (like Assembly Cloner does)
    invApp.Documents.CloseAll
    Err.Clear

    ' Open the IDW
    Dim idwDoc
    Set idwDoc = invApp.Documents.Open(idwPath, False)

    If Err.Number <> 0 Or idwDoc Is Nothing Then
        LogMessage "IDW: ERROR - Could not open: " & Err.Description
        Err.Clear
        errorCount = 1
        Exit Sub
    End If

    LogMessage "IDW: Opened successfully"

    ' Access file references
    Dim fileDescriptors
    Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
    LogMessage "IDW: Found " & fileDescriptors.Count & " referenced files"

    Dim mappingUsage ' Dictionary to track which mapping files were used
    Set mappingUsage = CreateObject("Scripting.Dictionary")

    ' Variables for the loop
    ' Dim i
    ' Dim fd
    ' Dim currentFullPath
    ' Dim currentFileName
    ' Dim mappingFound
    ' Dim newPath
    ' Dim mappingFileUsed
    ' Dim mappingSource
    ' Dim newFileName
    ' Dim newFilename

    For i = 1 To fileDescriptors.Count
        Set fd = fileDescriptors.Item(i)

        currentFullPath = fd.FullFileName
        currentFileName = fso.GetFileName(currentFullPath)

        ' Only log first 3 references in detail to avoid spam
        If i <= 3 Then
            LogMessage "  [" & i & "] Reference: " & currentFileName
            LogMessage "      Full path: " & currentFullPath
        ElseIf i = 4 Then
            LogMessage "  [4+] ... (remaining references not logged individually)"
        End If

        ' Check if this file is in our aggregated mapping
        mappingFound = False

        ' Method 1: Exact path match
        If g_ComprehensiveMapping.Exists(currentFullPath) Then
            newPath = g_ComprehensiveMapping.Item(currentFullPath)
            mappingFound = True
            mappingSource = "exact path"

            ' Track which mapping file provided this mapping
            mappingFileUsed = GetMappingFileForPath(currentFullPath)
            If mappingFileUsed <> "" And Not mappingUsage.Exists(mappingFileUsed) Then
                mappingUsage.Add mappingFileUsed, 1
            ElseIf mappingFileUsed <> "" Then
                mappingUsage.Item(mappingFileUsed) = mappingUsage.Item(mappingFileUsed) + 1
            End If
        End If

        ' Method 3: Filename-only match (for cloner compatibility)
        If Not mappingFound Then
            If i <= 3 Then LogMessage "      Checking filename mapping for: " & LCase(currentFileName)
            If g_FilenameMapping.Exists(LCase(currentFileName)) Then
                If i <= 3 Then LogMessage "      FOUND mapping: " & currentFileName & " -> " & g_FilenameMapping.Item(LCase(currentFileName))
                newFilename = g_FilenameMapping.Item(LCase(currentFileName))

                ' Find the actual file path for this filename in the same directory or subdirectories
                newPath = FindFileByNameInDirectory(fso.GetParentFolderName(idwPath), newFilename)
                If i <= 3 Then LogMessage "      Searched in: " & fso.GetParentFolderName(idwPath) & ", found: " & newPath

                If newPath <> "" Then
                    mappingFound = True
                    mappingSource = "filename match"
                    mappingFileUsed = "FILENAME_MAPPING"

                    ' Track filename mapping usage
                    If Not mappingUsage.Exists("FILENAME_MAPPING") Then
                        mappingUsage.Add "FILENAME_MAPPING", 1
                    Else
                        mappingUsage.Item("FILENAME_MAPPING") = mappingUsage.Item("FILENAME_MAPPING") + 1
                    End If
                End If
            Else
                If i <= 3 Then LogMessage "      No filename mapping found for " & LCase(currentFileName)
            End If
        End If

        If mappingFound Then
            newFileName = fso.GetFileName(newPath)

            If currentFullPath = newPath Then
                If i <= 3 Then LogMessage "      Already correct - no update needed"
            Else
                LogMessage "  UPDATING [" & i & "]: " & currentFileName & " -> " & newFileName
                LogMessage "      OLD: " & currentFullPath
                LogMessage "      NEW: " & newPath

                If fso.FileExists(newPath) Then
                    Err.Clear
                    fd.ReplaceReference newPath

                    If Err.Number = 0 Then
                        LogMessage "  SUCCESS"
                        updateCount = updateCount + 1
                    Else
                        LogMessage "  ERROR: " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                Else
                    LogMessage "  ERROR: New file doesn't exist: " & newPath
                    errorCount = errorCount + 1
                End If
            End If
        Else
            If i <= 3 Then LogMessage "      No mapping found - keeping current reference"
        End If
    Next

    ' Build mappings used string
    Dim usageKeys
    usageKeys = mappingUsage.Keys
    Dim j
    For j = 0 To UBound(usageKeys)
        Dim source
        source = usageKeys(j)
        mappingsUsed = mappingsUsed & source & " (" & mappingUsage.Item(source) & " refs)"
        If j < UBound(usageKeys) Then
            mappingsUsed = mappingsUsed & " | "
        End If
    Next

    If mappingsUsed = "" Then
        mappingsUsed = "No mappings used"
    End If

    ' Save IDW if updates were made
    If updateCount > 0 Then
        LogMessage "IDW: Saving with " & updateCount & " updates..."
        idwDoc.Save
        LogMessage "IDW: SUCCESS - Saved with " & updateCount & " updates"
        mappingsUsed = mappingsUsed & " (UPDATED)"
    Else
        LogMessage "IDW: No updates made"
    End If

    idwDoc.Close
    Err.Clear
End Sub

Function GetMappingFileForPath(originalPath)
    ' Get the mapping file that provided the mapping for this original path
    ' This is tracked when we load mappings
    On Error Resume Next

    If g_ComprehensiveMapping.Exists(originalPath) Then
        Dim mappingValue
        mappingValue = g_ComprehensiveMapping.Item(originalPath)

        ' The mapping value might be stored with metadata
        ' For now, just trace back to find which mapping file has this original path
        ' This is a simplified approach - we'd need to track metadata during aggregation

        ' Get the script directory to search for mapping files
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim rootDir
        rootDir = GetScriptDirectory()

        ' Try to find mapping files and check which one has this path
        Dim mappingFiles
        Set mappingFiles = CreateObject("Scripting.Dictionary")
        Call FindAllMappingFiles(rootDir, mappingFiles)

        Dim mappingKeys
        mappingKeys = mappingFiles.Keys

        Dim i
        For i = 0 To UBound(mappingKeys)
            Dim mappingPath
            mappingPath = mappingKeys(i)

            Dim mappingFile
            Set mappingFile = fso.OpenTextFile(mappingPath, 1)

            Do While Not mappingFile.AtEndOfStream
                Dim line
                line = mappingFile.ReadLine

                If Left(Trim(line), 1) <> "#" And Trim(line) <> "" Then
                    Dim parts
                    parts = Split(line, "|")

                    If UBound(parts) >= 3 Then
                        Dim original
                        original = Trim(parts(0))

                        If original = originalPath Then
                            mappingFile.Close
                            GetMappingFileForPath = mappingPath
                            Exit Function
                        End If
                    End If
                End If
            Loop

            mappingFile.Close
        Next
    End If

    GetMappingFileForPath = ""
End Function

Function ResolveActualFilePath(idwReferencePath)
    ' Resolve where a file actually is now (after renaming)
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(idwReferencePath) Then
        ResolveActualFilePath = idwReferencePath
        Exit Function
    End If

    Dim currentDir
    currentDir = fso.GetParentFolderName(idwReferencePath)

    If Not fso.FolderExists(currentDir) Then
        ResolveActualFilePath = ""
        Exit Function
    End If

    Dim folder
    Set folder = fso.GetFolder(currentDir)
    Dim file
    Dim candidateFiles
    Set candidateFiles = CreateObject("Scripting.Dictionary")

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".ipt" Then
            If g_ComprehensiveMapping.Exists(file.Path) Then
                candidateFiles.Add file.Path, file.Name
            End If
        End If
    Next

    If candidateFiles.Count = 1 Then
        Dim keys
        keys = candidateFiles.Keys
        ResolveActualFilePath = keys(0)
    Else
        ResolveActualFilePath = ""
    End If
End Function

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function GetScriptDirectory()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetScriptDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
End Function

Function BrowseForFolder(prompt)
    On Error Resume Next

    Dim shell
    Set shell = CreateObject("Shell.Application")

    Dim folder
    Set folder = shell.BrowseForFolder(0, prompt, 0, 0)

    If folder Is Nothing Then
        BrowseForFolder = ""
        Exit Function
    End If

    BrowseForFolder = folder.Self.Path
    Err.Clear
End Function

Sub GenerateMappingReport(mappingFiles, idwFiles)
    ' Generate detailed report showing which IDWs used which mapping files
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim scriptDir
    scriptDir = GetScriptDirectory()
    Dim rootDir
    rootDir = fso.GetParentFolderName(scriptDir)
    Dim logsDir
    logsDir = rootDir & "\Logs"

    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder(logsDir)
    End If

    g_ReportPath = logsDir & "\Recursive_IDW_Updater_Report_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".txt"

    Set g_ReportFileNum = fso.CreateTextFile(g_ReportPath, True)

    g_ReportFileNum.WriteLine "=================================================================="
    g_ReportFileNum.WriteLine "RECURSIVE IDW UPDATER - DETAILED REPORT"
    g_ReportFileNum.WriteLine "Generated: " & Now
    g_ReportFileNum.WriteLine "=================================================================="
    g_ReportFileNum.WriteLine ""

    g_ReportFileNum.WriteLine "MAPPING FILES FOUND: " & mappingFiles.Count
    g_ReportFileNum.WriteLine "---------------------------"

    Dim mappingKeys
    mappingKeys = mappingFiles.Keys
    Dim i
    For i = 0 To UBound(mappingKeys)
        g_ReportFileNum.WriteLine "  [" & (i + 1) & "] " & mappingKeys(i)
    Next

    g_ReportFileNum.WriteLine ""
    g_ReportFileNum.WriteLine "MAPPINGS BY SOURCE FILE:"
    g_ReportFileNum.WriteLine "---------------------------"

    ' Count mappings per file
    Dim mappingsPerFile
    Set mappingsPerFile = CreateObject("Scripting.Dictionary")

    Dim comprehensiveKeys
    comprehensiveKeys = g_ComprehensiveMapping.Keys
    For i = 0 To UBound(comprehensiveKeys)
        Dim originalPath
        originalPath = comprehensiveKeys(i)

        Dim mappingFile
        mappingFile = GetMappingFileForPath(originalPath)

        If mappingFile <> "" Then
            If Not mappingsPerFile.Exists(mappingFile) Then
                mappingsPerFile.Add mappingFile, 1
            Else
                mappingsPerFile.Item(mappingFile) = mappingsPerFile.Item(mappingFile) + 1
            End If
        End If
    Next

    Dim mappingCountsKeys
    mappingCountsKeys = mappingsPerFile.Keys
    For i = 0 To UBound(mappingCountsKeys)
        g_ReportFileNum.WriteLine "  " & mappingCountsKeys(i) & " : " & mappingsPerFile.Item(mappingCountsKeys(i)) & " mappings"
    Next

    g_ReportFileNum.WriteLine ""
    g_ReportFileNum.WriteLine "IDW FILES PROCESSED: " & idwFiles.Count
    g_ReportFileNum.WriteLine "---------------------------"

    Dim idwKeys
    idwKeys = idwFiles.Keys
    For i = 0 To UBound(idwKeys)
        Dim idwPath
        idwPath = idwKeys(i)

        g_ReportFileNum.WriteLine "  [" & (i + 1) & "] " & idwFiles.Item(idwPath)

        Dim mappingsUsedStr
        mappingsUsedStr = ""

        If g_IDWToMappingReport.Exists(idwPath) Then
            g_ReportFileNum.WriteLine "    Mappings: " & g_IDWToMappingReport.Item(idwPath)
        Else
            g_ReportFileNum.WriteLine "    Mappings: None (no mappings used)"
        End If
    Next

    ' Count IDWs by mapping usage
    g_ReportFileNum.WriteLine ""
    g_ReportFileNum.WriteLine "IDWs BY MAPPING FILE USAGE:"
    g_ReportFileNum.WriteLine "---------------------------"

    Dim idwMappingStats
    Set idwMappingStats = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(idwKeys)
        idwPath = idwKeys(i)

        If g_IDWToMappingReport.Exists(idwPath) Then
            Dim mapping
            mapping = g_IDWToMappingReport.Item(idwPath)

            If Not idwMappingStats.Exists(mapping) Then
                idwMappingStats.Add mapping, 1
            Else
                idwMappingStats.Item(mapping) = idwMappingStats.Item(mapping) + 1
            End If
        End If
    Next

    Dim statsKeys
    statsKeys = idwMappingStats.Keys
    For i = 0 To UBound(statsKeys)
        g_ReportFileNum.WriteLine "  " & statsKeys(i) & " : " & idwMappingStats.Item(statsKeys(i)) & " IDWs"
    Next

    g_ReportFileNum.WriteLine ""
    g_ReportFileNum.WriteLine "IDWs WITH NO MAPPINGS USED:"
    g_ReportFileNum.WriteLine "---------------------------"

    Dim noMappingCount
    noMappingCount = 0

    For i = 0 To UBound(idwKeys)
        idwPath = idwKeys(i)

        If Not g_IDWToMappingReport.Exists(idwPath) Then
            g_ReportFileNum.WriteLine "  " & idwFiles.Item(idwPath)
            noMappingCount = noMappingCount + 1
        End If
    Next

    If noMappingCount = 0 Then
        g_ReportFileNum.WriteLine "  None - all IDWs had mapping data"
    End If

    g_ReportFileNum.WriteLine ""
    g_ReportFileNum.WriteLine "=================================================================="
    g_ReportFileNum.WriteLine "END OF REPORT"
    g_ReportFileNum.WriteLine "=================================================================="

    g_ReportFileNum.Close

    LogMessage "REPORT GENERATED: " & g_ReportPath
End Sub

Sub StartLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = GetScriptDirectory()
    Dim rootDir
    rootDir = fso.GetParentFolderName(scriptDir)
    Dim logsDir
    logsDir = rootDir & "\Logs"

    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder(logsDir)
    End If

    g_LogPath = logsDir & "\Recursive_IDW_Updater_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"

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

Function FindFileByNameInDirectory(dirPath, filename)
    ' Recursively search for a file by name in a directory
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(dirPath) Then
        FindFileByNameInDirectory = ""
        Exit Function
    End If

    Dim folder
    Set folder = fso.GetFolder(dirPath)

    ' Check files in current directory
    Dim file
    For Each file In folder.Files
        If LCase(file.Name) = LCase(filename) Then
            FindFileByNameInDirectory = file.Path
            Exit Function
        End If
    Next

    ' Recursively check subdirectories
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" And _
           LCase(subFolder.Name) <> "temp" And _
           LCase(subFolder.Name) <> "$recycle.bin" Then
            Dim result
            result = FindFileByNameInDirectory(subFolder.Path, filename)
            If result <> "" Then
                FindFileByNameInDirectory = result
                Exit Function
            End If
        End If
    Next

    FindFileByNameInDirectory = ""
End Function
