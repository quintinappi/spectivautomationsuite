Option Explicit

' ==============================================================================
' EMERGENCY IDW FIXER - DETAILING WORKFLOW STEP 2a: Fix IDW References
' ==============================================================================
' DETAILING WORKFLOW - STEP 2a: Run idw fixer
'
' This standalone rescue tool fixes IDW references after renaming
' Use this when the main STEP 1 misses specific IDW files in problematic folders
'
' This tool:
' 1. Asks user for a specific folder to scan
' 2. Finds all IDW files in that folder
' 3. Finds all .ipt files (original and heritage names)
' 4. Intelligently builds mapping from file analysis
' 5. Updates IDW references using proven Design Assistant method
'
' WHEN TO USE:
' - After completing PART RENAMING (STEP 1)
' - As part of DETAILING WORKFLOW (STEP 2)
' - When specific folders like "Launder 2" or "Extension" weren't updated
' - When you need to fix just ONE problematic assembly folder
' - When the main STEP 1 missed some IDWs
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_LocalMapping ' Built from folder analysis

Call EMERGENCY_IDW_FIXER()

Sub EMERGENCY_IDW_FIXER()
    Call StartLogging
    LogMessage "=== EMERGENCY IDW FIXER ==="
    LogMessage "Standalone rescue tool for problematic folders"

    Dim result
    result = MsgBox("EMERGENCY IDW FIXER - DETAILING STEP 2a" & vbCrLf & vbCrLf & _
                    "DETAILING WORKFLOW - STEP 2a: Run idw fixer" & vbCrLf & vbCrLf & _
                    "Use this tool when STEP 1 misses specific folders!" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Ask you to select a folder to fix" & vbCrLf & _
                    "2. Scan for IDW files in that folder" & vbCrLf & _
                    "3. Build mapping from heritage files found" & vbCrLf & _
                    "4. Update IDW references automatically" & vbCrLf & vbCrLf & _
                    "Perfect for folders like:" & vbCrLf & _
                    "  - Launder 2" & vbCrLf & _
                    "  - Extension" & vbCrLf & _
                    "  - Any subfolder with missed IDWs" & vbCrLf & vbCrLf & _
                    "⚠️  Run this AFTER completing PART RENAMING (STEP 1)!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Step 2a: IDW Fixer")

    If result = vbNo Then
        LogMessage "User cancelled emergency fixer"
        Exit Sub
    End If

    ' Initialize local mapping
    Set g_LocalMapping = CreateObject("Scripting.Dictionary")

    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor first.", vbCritical
        Exit Sub
    End If

    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear

    ' Step 1: Get folder to fix
    LogMessage "STEP 1: Getting folder path from user"
    Dim folderPath
    folderPath = BrowseForFolder()

    If folderPath = "" Then
        LogMessage "User cancelled folder selection"
        MsgBox "No folder selected. Exiting.", vbInformation
        Exit Sub
    End If

    LogMessage "FOLDER: Selected folder: " & folderPath

    ' Step 2: Scan folder for IDW files
    LogMessage "STEP 2: Scanning folder for IDW files"
    Dim idwFiles
    Set idwFiles = CreateObject("Scripting.Dictionary")
    Call FindIDWsInFolder(folderPath, idwFiles)

    If idwFiles.Count = 0 Then
        LogMessage "ERROR: No IDW files found in selected folder"
        MsgBox "No IDW files found in:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
               "Make sure you selected the correct folder.", vbExclamation
        Exit Sub
    End If

    LogMessage "IDW: Found " & idwFiles.Count & " IDW files to fix"

    ' Step 3: Build intelligent mapping from folder analysis
    LogMessage "STEP 3: Building intelligent mapping from folder analysis"
    Call BuildIntelligentMappingFromFolder(folderPath)

    If g_LocalMapping.Count = 0 Then
        LogMessage "WARNING: No heritage files detected in folder"
        Dim continueAnyway
        continueAnyway = MsgBox("WARNING: No heritage files detected!" & vbCrLf & vbCrLf & _
                               "Found " & idwFiles.Count & " IDW files" & vbCrLf & _
                               "But found 0 heritage file mappings" & vbCrLf & vbCrLf & _
                               "This might mean:" & vbCrLf & _
                               "  - Parts haven't been renamed yet (run STEP 1)" & vbCrLf & _
                               "  - Wrong folder selected" & vbCrLf & vbCrLf & _
                               "Continue anyway? (Will process IDWs but may not update references)", _
                               vbYesNo + vbQuestion, "No Heritage Files Found")

        If continueAnyway = vbNo Then
            LogMessage "User cancelled - no mappings found"
            Exit Sub
        End If
    Else
        LogMessage "MAPPING: Built " & g_LocalMapping.Count & " mappings from folder"
    End If

    ' Step 4: Update IDW files
    LogMessage "STEP 4: Updating IDW references with emergency fixer"
    Dim totalUpdates, totalErrors
    Call UpdateIDWFilesWithLocalMapping(invApp, idwFiles, totalUpdates, totalErrors)

    LogMessage "=== EMERGENCY IDW FIXER COMPLETED ==="
    Call StopLogging

    ' Show results
    Dim resultMsg
    If totalErrors > 0 Then
        resultMsg = "EMERGENCY FIX COMPLETED WITH ERRORS!" & vbCrLf & vbCrLf
    ElseIf totalUpdates = 0 Then
        resultMsg = "EMERGENCY FIX COMPLETED - NO UPDATES NEEDED" & vbCrLf & vbCrLf
    Else
        resultMsg = "EMERGENCY FIX COMPLETED SUCCESSFULLY!" & vbCrLf & vbCrLf
    End If

    resultMsg = resultMsg & _
               "Folder: " & folderPath & vbCrLf & _
               "IDW files found: " & idwFiles.Count & vbCrLf & _
               "Mappings built: " & g_LocalMapping.Count & vbCrLf & _
               "References updated: " & totalUpdates & vbCrLf & _
               "Errors: " & totalErrors & vbCrLf & vbCrLf & _
               "Log: " & g_LogPath

    MsgBox resultMsg, vbInformation, "Emergency Fixer Complete"
End Sub

Function BrowseForFolder()
    ' Show folder browser using PowerShell
    LogMessage "BROWSE: Opening folder browser dialog"

    Dim shell, cmd, result
    Set shell = CreateObject("WScript.Shell")

    cmd = "powershell -WindowStyle Hidden -Command """ & _
          "Add-Type -AssemblyName System.Windows.Forms;" & _
          "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog;" & _
          "$dialog.Description = 'Select folder containing IDW files to fix';" & _
          "$dialog.ShowNewFolderButton = $false;" & _
          "if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {" & _
          "    Write-Output $dialog.SelectedPath" & _
          "} else {" & _
          "    Write-Output 'CANCELLED'" & _
          "}" & _
          """"

    On Error Resume Next
    result = shell.Exec(cmd).StdOut.ReadAll()

    ' Remove ALL whitespace including newlines
    result = Replace(result, vbCrLf, "")
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Trim(result)

    If result <> "" And result <> "CANCELLED" Then
        BrowseForFolder = result
        LogMessage "BROWSE: User selected: " & result
    Else
        BrowseForFolder = ""
        LogMessage "BROWSE: User cancelled"
    End If

    Err.Clear
End Function

Sub FindIDWsInFolder(folderPath, idwFiles)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        LogMessage "ERROR: Folder not found: " & folderPath
        Exit Sub
    End If

    Dim folder
    Set folder = fso.GetFolder(folderPath)

    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            idwFiles.Add file.Path, file.Name
            LogMessage "IDW: Found - " & file.Name
        End If
    Next

    ' Also check subdirectories
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" Then
            Call FindIDWsInFolder(subFolder.Path, idwFiles)
        End If
    Next
End Sub

Sub BuildIntelligentMappingFromFolder(folderPath)
    LogMessage "MAPPING: Analyzing folder for heritage files"

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        LogMessage "ERROR: Folder not found: " & folderPath
        Exit Sub
    End If

    ' Strategy: Find heritage files (with prefix patterns) and original files
    ' Heritage files typically have patterns like: PREFIX-000-CODE123.ipt
    ' Original files typically have simple names: Part1.ipt, Beam 203x203.ipt

    Dim folder
    Set folder = fso.GetFolder(folderPath)

    ' First pass: Find all .ipt files and categorize them
    Dim heritageFiles, possibleOriginalFiles
    Set heritageFiles = CreateObject("Scripting.Dictionary")
    Set possibleOriginalFiles = CreateObject("Scripting.Dictionary")

    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".ipt" Then
            ' Detect if this is a heritage file (has prefix pattern)
            If IsHeritageFileName(file.Name) Then
                heritageFiles.Add file.Path, file.Name
                LogMessage "HERITAGE: Detected heritage file - " & file.Name
            Else
                possibleOriginalFiles.Add file.Path, file.Name
                LogMessage "ORIGINAL: Possible original file - " & file.Name
            End If
        End If
    Next

    ' Also check subdirectories
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" Then
            For Each file In subFolder.Files
                If LCase(Right(file.Name, 4)) = ".ipt" Then
                    If IsHeritageFileName(file.Name) Then
                        heritageFiles.Add file.Path, file.Name
                        LogMessage "HERITAGE: Detected heritage file in subfolder - " & file.Name
                    Else
                        possibleOriginalFiles.Add file.Path, file.Name
                        LogMessage "ORIGINAL: Possible original file in subfolder - " & file.Name
                    End If
                End If
            Next
        End If
    Next

    LogMessage "MAPPING: Found " & heritageFiles.Count & " heritage files and " & possibleOriginalFiles.Count & " possible original files"

    ' Second pass: For each original file, try to find matching heritage file by location
    ' The heritage file will be in the same directory as the original
    Dim origKeys
    If possibleOriginalFiles.Count > 0 Then
        origKeys = possibleOriginalFiles.Keys

        Dim i
        For i = 0 To UBound(origKeys)
            Dim origPath
            origPath = origKeys(i)
            Dim origName
            origName = possibleOriginalFiles.Item(origPath)
            Dim origDir
            origDir = fso.GetParentFolderName(origPath)

            ' Look for heritage file in same directory
            Dim heritageKeys
            heritageKeys = heritageFiles.Keys

            Dim j
            For j = 0 To UBound(heritageKeys)
                Dim heritagePath
                heritagePath = heritageKeys(j)
                Dim heritageName
                heritageName = heritageFiles.Item(heritagePath)
                Dim heritageDir
                heritageDir = fso.GetParentFolderName(heritagePath)

                ' If in same directory, this might be a mapping pair
                If LCase(origDir) = LCase(heritageDir) Then
                    ' Add mapping: original -> heritage
                    If Not g_LocalMapping.Exists(origPath) Then
                        g_LocalMapping.Add origPath, heritagePath
                        LogMessage "MAPPING: " & origName & " -> " & heritageName
                    End If
                End If
            Next
        Next
    End If

    LogMessage "MAPPING: Built " & g_LocalMapping.Count & " intelligent mappings"
End Sub

Function IsHeritageFileName(fileName)
    ' Detect if filename matches heritage pattern
    ' Patterns: PREFIX-###-CODEXXX.ipt
    ' Examples: NCRH01-000-PL123.ipt, PLANT1-000-B45.ipt

    Dim baseName
    baseName = fileName
    If LCase(Right(baseName, 4)) = ".ipt" Then
        baseName = Left(baseName, Len(baseName) - 4)
    End If

    ' Count dashes - heritage files typically have at least 2 dashes
    Dim dashCount
    dashCount = 0
    Dim i
    For i = 1 To Len(baseName)
        If Mid(baseName, i, 1) = "-" Then
            dashCount = dashCount + 1
        End If
    Next

    ' Heritage files have format: XXX-###-CODEXXX
    ' Must have at least 2 dashes and end with letters+numbers
    If dashCount >= 2 Then
        ' Check if ends with letter+number pattern (e.g., PL123, B45)
        Dim lastChar
        lastChar = Right(baseName, 1)
        If IsNumeric(lastChar) Then
            IsHeritageFileName = True
            Exit Function
        End If
    End If

    IsHeritageFileName = False
End Function

Sub UpdateIDWFilesWithLocalMapping(invApp, idwFiles, ByRef totalUpdates, ByRef totalErrors)
    LogMessage "IDW: Processing " & idwFiles.Count & " IDW files with local mapping"

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

        LogMessage "IDW: Processing [" & (i + 1) & "/" & (UBound(idwKeys) + 1) & "] " & idwName

        Dim updates, errors
        Call UpdateSingleIDWWithLocalMapping(invApp, idwPath, updates, errors)
        totalUpdates = totalUpdates + updates
        totalErrors = totalErrors + errors
    Next

    LogMessage "IDW: Total updates: " & totalUpdates & ", Total errors: " & totalErrors
End Sub

Sub UpdateSingleIDWWithLocalMapping(invApp, idwPath, ByRef updateCount, ByRef errorCount)
    On Error Resume Next
    updateCount = 0
    errorCount = 0

    LogMessage "IDW: Opening " & GetFileNameFromPath(idwPath)

    ' Set resolve options
    Dim originalResolveMode
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
        invApp.FileOptions.ResolveFileOption = originalResolveMode
        invApp.SilentOperation = originalSilentMode
        errorCount = 1
        Exit Sub
    End If

    ' Restore resolve mode
    invApp.FileOptions.ResolveFileOption = originalResolveMode
    Err.Clear

    ' Access file descriptors
    Dim fileDescriptors
    Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
    LogMessage "IDW: Found " & fileDescriptors.Count & " referenced files"

    ' Update each reference
    Dim i
    For i = 1 To fileDescriptors.Count
        Dim fd
        Set fd = fileDescriptors.Item(i)

        Dim currentPath
        currentPath = fd.FullFileName
        Dim currentName
        currentName = GetFileNameFromPath(currentPath)

        LogMessage "IDW:   Checking reference: " & currentName

        ' Check if we have a mapping for this file
        If g_LocalMapping.Exists(currentPath) Then
            Dim newPath
            newPath = g_LocalMapping.Item(currentPath)
            Dim newName
            newName = GetFileNameFromPath(newPath)

            If currentPath <> newPath Then
                LogMessage "IDW:     UPDATING: " & currentName & " -> " & newName

                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")
                If fso.FileExists(newPath) Then
                    Err.Clear
                    fd.ReplaceReference newPath

                    If Err.Number = 0 Then
                        LogMessage "IDW:     ✓ SUCCESS"
                        updateCount = updateCount + 1
                    Else
                        LogMessage "IDW:     ✗ ERROR: " & Err.Description
                        errorCount = errorCount + 1
                        Err.Clear
                    End If
                Else
                    LogMessage "IDW:     ✗ ERROR: New file not found: " & newPath
                    errorCount = errorCount + 1
                End If
            Else
                LogMessage "IDW:     (Already correct)"
            End If
        Else
            LogMessage "IDW:     (No mapping - keeping current)"
        End If
    Next

    ' Save if updates made
    If updateCount > 0 Then
        LogMessage "IDW: Saving with " & updateCount & " updates..."
        idwDoc.Save2(True)
        LogMessage "IDW: Saved successfully"
    Else
        LogMessage "IDW: No updates made"
    End If

    idwDoc.Close
    
    ' Restore original settings
    invApp.SilentOperation = originalSilentMode
    Err.Clear
End Sub

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
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
    g_LogPath = logsDir & "\Emergency_IDW_Fixer_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
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