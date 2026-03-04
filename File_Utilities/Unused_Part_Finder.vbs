' ============================================================================
' UNUSED PART FINDER - DETAILING WORKFLOW STEP 13: Clean up unused parts
' ============================================================================
' DETAILING WORKFLOW - STEP 13: Clean up unused parts into a folder
' Description: Finds unused IPT files in project folder and moves to backup
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
' ============================================================================
' This tool:
' 1. Scans the currently open assembly and ALL sub-assemblies
' 2. Builds a list of ALL parts referenced in the assembly
' 3. Scans the assembly's folder and all subfolders for .ipt files
' 4. Identifies .ipt files that are NOT referenced in the assembly
' 5. Moves unused .ipt files to a backup folder
'
' WHEN TO USE:
' - After renaming parts to clean up old unreferenced files
' - Before archiving a project to remove unused parts
' - When project folder has accumulated orphaned part files
' ============================================================================

Option Explicit

' Global variables
Dim g_LogFile
Dim g_LogPath
Dim g_ScriptPath
Dim g_ScannedParts ' Dictionary: fullPath -> True (parts in assembly)
Dim g_AllIPTFiles ' Dictionary: fullPath -> True (all IPT files in folders)
Dim g_UnusedFiles ' Array of unused file paths

Call MAIN()

Sub MAIN()
    Call InitializeLogging
    LogMessage "=========================================="
    LogMessage "UNUSED PART FINDER - STARTING"
    LogMessage "=========================================="

    ' Confirm with user
    If Not GetUserConfirmation() Then
        LogMessage "Operation cancelled by user"
        Exit Sub
    End If

    ' Step 1: Connect to Inventor and get assembly
    Dim invApp, asmDoc, asmFolder
    Set invApp = ConnectToInventor()
    If invApp Is Nothing Then
        MsgBox "Failed to connect to Inventor!" & vbCrLf & _
               "Please make sure Inventor is running with an assembly open.", vbCritical, "Error"
        Exit Sub
    End If

    Set asmDoc = GetActiveAssembly(invApp)
    If asmDoc Is Nothing Then
        MsgBox "No assembly is currently open!" & vbCrLf & _
               "Please open your main assembly first.", vbCritical, "Error"
        Exit Sub
    End If

    asmFolder = GetFolderFromPath(asmDoc.FullFileName)
    LogMessage "ASSEMBLY: " & asmDoc.DisplayName
    LogMessage "FOLDER: " & asmFolder

    ' Step 2: Scan assembly for all referenced parts
    LogMessage "=========================================="
    LogMessage "STEP 1: SCANNING ASSEMBLY FOR PARTS"
    LogMessage "=========================================="
    Call ScanAssemblyForParts(asmDoc)
    LogMessage "PARTS IN ASSEMBLY: " & g_ScannedParts.Count

    ' Step 3: Scan folders for all IPT files
    LogMessage "=========================================="
    LogMessage "STEP 2: SCANNING FOLDERS FOR IPT FILES"
    LogMessage "=========================================="
    Call ScanFolderForIPTFiles(asmFolder)
    LogMessage "TOTAL IPT FILES FOUND: " & g_AllIPTFiles.Count

    ' Step 4: Find unused files
    LogMessage "=========================================="
    LogMessage "STEP 3: FINDING UNUSED FILES"
    LogMessage "=========================================="
    Call FindUnusedFiles()

    ' Step 5: Show results and confirm
    If Not ConfirmMoveOperation() Then
        LogMessage "Operation cancelled by user"
        Exit Sub
    End If

    ' Step 6: Move unused files to backup
    LogMessage "=========================================="
    LogMessage "STEP 4: MOVING UNUSED FILES TO BACKUP"
    LogMessage "=========================================="
    Call MoveUnusedFilesToBackup(asmFolder)

    LogMessage "=========================================="
    LogMessage "UNUSED PART FINDER - COMPLETED"
    LogMessage "=========================================="

    MsgBox "Operation completed successfully!" & vbCrLf & vbCrLf & _
           "Unused files moved: " & UBound(g_UnusedFiles) + 1 & vbCrLf & vbCrLf & _
           "Backup folder created in assembly directory" & vbCrLf & _
           "Log file: " & g_LogPath, vbInformation, "Complete"

    Call CleanupLogging
End Sub

' ============================================================================
' INITIALIZATION AND LOGGING
' ============================================================================
Sub InitializeLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    g_ScriptPath = fso.GetParentFolderName(WScript.ScriptFullName)

    ' Create Logs folder if not exists
    Dim logsDir
    logsDir = g_ScriptPath & "\..\Logs"
    logsDir = fso.GetAbsolutePathName(logsDir)

    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder logsDir
    End If

    ' Create log file with timestamp
    Dim timestamp
    timestamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

    g_LogPath = logsDir & "\Unused_Part_Finder_" & timestamp & ".log"
    Set g_LogFile = fso.CreateTextFile(g_LogPath, True)

    ' Initialize collections
    Set g_ScannedParts = CreateObject("Scripting.Dictionary")
    Set g_AllIPTFiles = CreateObject("Scripting.Dictionary")
End Sub

Sub LogMessage(message)
    Dim timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)

    If Not g_LogFile Is Nothing Then
        g_LogFile.WriteLine timestamp & " | " & message
    End If
    WScript.Echo timestamp & " - " & message
End Sub

Sub CleanupLogging()
    If Not g_LogFile Is Nothing Then
        g_LogFile.Close
        Set g_LogFile = Nothing
    End If
    Set g_ScannedParts = Nothing
    Set g_AllIPTFiles = Nothing
End Sub

' ============================================================================
' USER INTERACTION
' ============================================================================
Function GetUserConfirmation()
    Dim result
    result = MsgBox("UNUSED PART FINDER" & vbCrLf & vbCrLf & _
                    "This tool will:" & vbCrLf & vbCrLf & _
                    "1. Scan your open assembly for all referenced parts" & vbCrLf & _
                    "2. Scan the assembly folder for ALL .ipt files" & vbCrLf & _
                    "3. Find .ipt files NOT used in the assembly" & vbCrLf & _
                    "4. Move unused files to a backup folder" & vbCrLf & vbCrLf & _
                    "Make sure your main assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Unused Part Finder")

    GetUserConfirmation = (result = vbYes)
End Function

Function ConfirmMoveOperation()
    Dim unusedCount
    unusedCount = UBound(g_UnusedFiles) + 1

    If unusedCount = 0 Then
        MsgBox "No unused files found!" & vbCrLf & vbCrLf & _
               "All .ipt files in the folder are referenced by the assembly.", vbInformation, "No Unused Files"
        ConfirmMoveOperation = False
        Exit Function
    End If

    Dim msg
    msg = "UNUSED FILES FOUND: " & unusedCount & vbCrLf & vbCrLf
    msg = msg & "These files will be moved to a backup folder:" & vbCrLf & vbCrLf

    ' Show first 10 files
    Dim i, maxShow
    maxShow = 10
    If unusedCount < maxShow Then maxShow = unusedCount

    For i = 0 To maxShow - 1
        msg = msg & "  - " & GetFileNameFromPath(g_UnusedFiles(i)) & vbCrLf
    Next

    If unusedCount > maxShow Then
        msg = msg & "  ... and " & (unusedCount - maxShow) & " more files" & vbCrLf
    End If

    msg = msg & vbCrLf & "Create backup folder and move these files?"

    Dim result
    result = MsgBox(msg, vbYesNo + vbQuestion, "Confirm Move Operation")

    ConfirmMoveOperation = (result = vbYes)
End Function

' ============================================================================
' INVENTOR CONNECTION
' ============================================================================
Function ConnectToInventor()
    On Error Resume Next
    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to connect to Inventor - " & Err.Description
        Set ConnectToInventor = Nothing
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0
    LogMessage "SUCCESS: Connected to Inventor"
    Set ConnectToInventor = invApp
End Function

Function GetActiveAssembly(invApp)
    On Error Resume Next
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If Err.Number <> 0 Or activeDoc Is Nothing Then
        LogMessage "ERROR: No active document"
        Set GetActiveAssembly = Nothing
        Exit Function
    End If

    ' Check if it's an assembly
    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        LogMessage "ERROR: Active document is not an assembly: " & activeDoc.DisplayName
        Set GetActiveAssembly = Nothing
        Exit Function
    End If

    Set GetActiveAssembly = activeDoc
    Err.Clear
End Function

' ============================================================================
' SCAN ASSEMBLY FOR PARTS
' ============================================================================
Sub ScanAssemblyForParts(asmDoc)
    LogMessage "Scanning: " & asmDoc.DisplayName

    ' Process this assembly's occurrences
    Call ProcessAssemblyOccurrences(asmDoc)

    LogMessage "Assembly scan complete"
End Sub

Sub ProcessAssemblyOccurrences(asmDoc)
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "Processing " & occurrences.Count & " occurrences in " & asmDoc.DisplayName

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Skip suppressed occurrences
        If occ.Suppressed Then
            LogMessage "SKIP: Suppressed - " & occ.Name
        Else
            ' Get referenced document
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If Not refDoc Is Nothing Then
                Dim fullPath
                fullPath = refDoc.FullFileName
                Dim fileName
                fileName = GetFileNameFromPath(fullPath)

                ' Check file type
                If LCase(Right(fileName, 4)) = ".ipt" Then
                    ' Part file - add to scanned parts if not already there
                    If Not g_ScannedParts.Exists(fullPath) Then
                        g_ScannedParts.Add fullPath, True
                        LogMessage "PART: Found - " & fileName
                    End If

                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    ' Sub-assembly - recurse into it
                    ' Skip bolted connections and content center files
                    If InStr(LCase(fileName), "bolted connection") = 0 And _
                       InStr(LCase(fileName), "content center") = 0 Then
                        LogMessage "SUB-ASM: Recursing into - " & fileName
                        Call ProcessAssemblyOccurrences(refDoc)
                    Else
                        LogMessage "SKIP: Content Center/Bolted Connection - " & fileName
                    End If
                End If
            End If
        End If
    Next

    ' Also check any occurrence patterns (mirror, rectangular pattern, etc.)
    Call ProcessOccurrencePatterns(asmDoc)
End Sub

Sub ProcessOccurrencePatterns(asmDoc)
    On Error Resume Next

    Dim Patterns
    Set Patterns = asmDoc.ComponentDefinition.OccurrencePatterns

    If Patterns Is Nothing Or Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    ' Process each pattern type
    Dim patternType, pattern
    For Each patternType In Array("RectangularPatternFeatures", "CircularPatternFeatures", "MirrorFeatures")
        On Error Resume Next
        Dim patternCollection
        Set patternCollection = asmDoc.ComponentDefinition.Features.Item(patternType)

        If Not patternCollection Is Nothing And Err.Number = 0 Then
            ' Patterns are handled through the main occurrences collection
            ' This is just a placeholder for future enhancement
        End If
        Err.Clear
    Next

    On Error GoTo 0
End Sub

' ============================================================================
' SCAN FOLDER FOR IPT FILES
' ============================================================================
Sub ScanFolderForIPTFiles(folderPath)
    LogMessage "Scanning folder: " & folderPath

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        LogMessage "ERROR: Folder not found - " & folderPath
        Exit Sub
    End If

    ' Get folder object
    Dim folder
    Set folder = fso.GetFolder(folderPath)

    ' Scan for IPT files in this folder
    Call ScanFolderForIPTsRecursive(folder)

    LogMessage "Folder scan complete"
End Sub

Sub ScanFolderForIPTsRecursive(folder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    LogMessage "Scanning: " & folder.Path

    ' Scan files in this folder
    Dim file
    For Each file in folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "ipt" Then
            If Not g_AllIPTFiles.Exists(file.Path) Then
                g_AllIPTFiles.Add file.Path, True
                LogMessage "FILE: Found - " & file.Name
            End If
        End If
    Next

    ' Recursively scan subfolders (SKIP backup folders and OldVersions)
    Dim subFolder
    For Each subFolder In folder.SubFolders
        ' Skip backup folders created by this tool and OldVersions
        If Left(LCase(subFolder.Name), 19) <> "unused_parts_backup" And _
           LCase(subFolder.Name) <> "oldversions" Then
            Call ScanFolderForIPTsRecursive(subFolder)
        Else
            LogMessage "SKIP: Backup/OldVersions folder - " & subFolder.Name
        End If
    Next
End Sub

' ============================================================================
' FIND UNUSED FILES
' ============================================================================
Sub FindUnusedFiles()
    LogMessage "Comparing assembly parts vs folder files..."

    Dim unusedList
    Set unusedList = CreateObject("Scripting.Dictionary")

    ' Find all IPT files that are NOT in the scanned parts
    Dim filePath
    For Each filePath In g_AllIPTFiles.Keys
        If Not g_ScannedParts.Exists(filePath) Then
            unusedList.Add filePath, True
            LogMessage "UNUSED: " & GetFileNameFromPath(filePath)
        End If
    Next

    ' Convert to array
    ReDim g_UnusedFiles(unusedList.Count - 1)
    Dim i
    i = 0
    For Each filePath In unusedList.Keys
        g_UnusedFiles(i) = filePath
        i = i + 1
    Next

    LogMessage "Unused files found: " & (UBound(g_UnusedFiles) + 1)
End Sub

' ============================================================================
' MOVE UNUSED FILES TO BACKUP
' ============================================================================
Sub MoveUnusedFilesToBackup(asmFolder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create backup folder
    Dim backupFolder
    backupFolder = asmFolder & "Unused_Parts_Backup_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)

    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
        LogMessage "Created backup folder: " & backupFolder
    End If

    ' Move files
    Dim i, successCount, errorCount
    successCount = 0
    errorCount = 0

    For i = 0 To UBound(g_UnusedFiles)
        Dim filePath
        filePath = g_UnusedFiles(i)
        Dim relativePath
        relativePath = GetRelativePath(filePath, asmFolder)
        Dim destPath
        destPath = backupFolder & "\" & relativePath

        ' Ensure destination directory exists
        Dim destDir
        destDir = fso.GetParentFolderName(destPath)
        If destDir <> "" And Not fso.FolderExists(destDir) Then
            Call CreateFolderRecursive(fso, destDir)
        End If

        On Error Resume Next
        fso.MoveFile filePath, destPath

        If Err.Number = 0 Then
            LogMessage "MOVED: " & relativePath & " -> Backup"
            successCount = successCount + 1
        Else
            LogMessage "ERROR: Failed to move " & relativePath & " - " & Err.Description
            errorCount = errorCount + 1
            Err.Clear
        End If

        On Error GoTo 0
    Next

    LogMessage "Move operation complete"
    LogMessage "Success: " & successCount & ", Errors: " & errorCount

    ' Create summary file in backup folder
    Call CreateBackupSummary(backupFolder, asmFolder, successCount, errorCount)
End Sub

Sub CreateBackupSummary(backupFolder, asmFolder, successCount, errorCount)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim summaryPath
    summaryPath = backupFolder & "\Backup_Summary.txt"

    Dim summaryFile
    Set summaryFile = fso.CreateTextFile(summaryPath, True)

    summaryFile.WriteLine "=========================================="
    summaryFile.WriteLine "UNUSED PARTS BACKUP SUMMARY"
    summaryFile.WriteLine "=========================================="
    summaryFile.WriteLine ""
    summaryFile.WriteLine "Backup Date: " & Now()
    summaryFile.WriteLine ""
    summaryFile.WriteLine "Files Moved: " & successCount
    summaryFile.WriteLine "Errors: " & errorCount
    summaryFile.WriteLine ""
    summaryFile.WriteLine "=========================================="
    summaryFile.WriteLine "MOVED FILES (relative paths):"
    summaryFile.WriteLine "=========================================="
    summaryFile.WriteLine ""

    Dim i
    For i = 0 To UBound(g_UnusedFiles)
        summaryFile.WriteLine GetRelativePath(g_UnusedFiles(i), asmFolder)
    Next

    summaryFile.WriteLine ""
    summaryFile.WriteLine "=========================================="
    summaryFile.WriteLine "Log file: " & g_LogPath
    summaryFile.WriteLine "=========================================="

    summaryFile.Close
    LogMessage "Created summary: " & summaryPath
End Sub

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================
Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function GetFolderFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderFromPath = fso.GetParentFolderName(fullPath) & "\"
End Function

Function GetRelativePath(fullPath, baseFolder)
    ' Get the relative path from baseFolder to fullPath
    If LCase(Left(fullPath, Len(baseFolder))) = LCase(baseFolder) Then
        GetRelativePath = Mid(fullPath, Len(baseFolder) + 1)
    Else
        ' Fallback to filename if not under baseFolder
        GetRelativePath = GetFileNameFromPath(fullPath)
    End If
End Function

Sub CreateFolderRecursive(fso, folderPath)
    ' Recursively create folders
    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentFolder
    parentFolder = fso.GetParentFolderName(folderPath)

    If parentFolder <> "" And Not fso.FolderExists(parentFolder) Then
        Call CreateFolderRecursive(fso, parentFolder)
    End If

    fso.CreateFolder folderPath
End Sub
