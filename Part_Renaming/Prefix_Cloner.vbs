Option Explicit

' ==============================================================================
' PREFIX CLONER - Clone Assembly with Prefix Replacement Only
' ==============================================================================
' Author: Quintin de Bruin © 2025
' 
' This script:
' 1. Detects currently open assembly in Inventor
' 2. Scans all files to detect the common prefix
' 3. Asks user to confirm old prefix and enter new prefix
' 4. Copies assembly, ALL sub-assemblies, parts, and IDW files to destination
' 5. Replaces ONLY the prefix portion of filenames (keeps suffixes intact)
' 6. Updates all references in copied assemblies and IDW files
' 7. Generates STEP_1_MAPPING.txt for traceability
'
' Example: N1SCR04-780-B1.IPT with prefix N1SCR04-780- becomes N2SCR04-780-B1.IPT
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_DestLogFileNum
Dim g_DestLogPath
Dim g_CopiedFiles       ' Dictionary: originalPath -> newPath
Dim g_OldPrefix         ' The prefix to replace
Dim g_NewPrefix         ' The new prefix
Dim g_MappingData       ' Array to store mapping data for STEP_1_MAPPING.txt
Dim g_SourceRoot        ' Source root folder for preserving folder structure

Call PREFIX_CLONER_MAIN()

Sub PREFIX_CLONER_MAIN()
    Call StartLogging
    LogMessage "=== PREFIX CLONER ==="
    LogMessage "Clone assembly with prefix replacement only (keep part suffixes intact)"
    
    Dim result
    result = MsgBox("PREFIX CLONER (Prefix Changer Only)" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Detect your currently open assembly" & vbCrLf & _
                    "2. Scan files to detect the common PREFIX" & vbCrLf & _
                    "3. Copy assembly + ALL parts to a NEW folder" & vbCrLf & _
                    "4. Replace ONLY the prefix (keep part suffixes same)" & vbCrLf & _
                    "5. Update all assembly and IDW references" & vbCrLf & _
                    "6. Generate STEP_1_MAPPING.txt for traceability" & vbCrLf & vbCrLf & _
                    "Example:" & vbCrLf & _
                    "  N1SCR04-780-B1.IPT  ->  N2SCR04-780-B1.IPT" & vbCrLf & _
                    "  (Prefix N1SCR04-780- replaced with N2SCR04-780-)" & vbCrLf & vbCrLf & _
                    "Make sure your source assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Prefix Cloner")
    
    If result = vbNo Then
        LogMessage "User cancelled workflow"
        Exit Sub
    End If
    
    ' Initialize collections
    Set g_CopiedFiles = CreateObject("Scripting.Dictionary")
    g_MappingData = Array()
    
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
    
    ' Step 2: Collect all referenced files
    LogMessage "STEP 2: Analyzing assembly structure"
    Dim allParts
    Set allParts = CreateObject("Scripting.Dictionary")
    Call CollectAllReferencedParts(sourceDoc, allParts)
    Call CollectIDWFiles(sourceDir, allParts)
    
    LogMessage "ANALYZE: Found " & allParts.Count & " unique files to process"
    
    ' Step 3: Detect common prefix from filenames
    LogMessage "STEP 3: Detecting common prefix from filenames"
    Dim detectedPrefix
    detectedPrefix = DetectCommonPrefix(allParts, sourceFileName)
    
    ' Step 4: Get prefix from user (confirm detected or enter manually)
    LogMessage "STEP 4: Getting prefix information from user"
    If Not GetPrefixFromUser(detectedPrefix) Then
        LogMessage "User cancelled prefix input"
        Exit Sub
    End If
    
    ' Step 5: Get destination folder
    LogMessage "STEP 5: Getting destination folder from user"
    Dim destFolder
    destFolder = GetDestinationFolder(sourceDir)
    If destFolder = "" Then
        LogMessage "User cancelled - no destination folder selected"
        Exit Sub
    End If

    ' Start destination logging in the selected folder
    Call StartDestinationLogging(destFolder)
    
    ' Get new assembly name (replace prefix in original name)
    Dim newAsmFileName
    newAsmFileName = ReplacePrefix(sourceFileName)
    LogMessage "New assembly name will be: " & newAsmFileName
    
    ' Step 6: Close source and copy assembly file first
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
    
    ' Store assembly mapping
    g_CopiedFiles.Add sourceFullPath, newAsmPath
    Call AddMappingData(sourceFullPath, newAsmPath, "ASSEMBLY")
    
    ' Step 7: Copy all parts and sub-assemblies with prefix replacement
    LogMessage "STEP 7: Copying parts and sub-assemblies with prefix replacement"
    Call CopyAllFilesWithPrefixReplacement(allParts, destFolder)
    
    ' Step 8: Update assembly references
    LogMessage "STEP 8: Updating references in ALL copied assemblies"
    
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
    
    ' Next, open ALL copied assemblies EXCEPT the main one
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
    
    ' Finally, open the main assembly
    Dim mainAsmDoc
    On Error Resume Next
    Set mainAsmDoc = invAppForUpdate.Documents.Open(newAsmPath, False)
    
    If Not mainAsmDoc Is Nothing Then
        LogMessage "STEP 8c: Main assembly opened successfully"
        
        ' Now update references in all assemblies
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
        
        ' CRITICAL: Explicitly update the main assembly
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
    
    ' Step 9: Update IDW drawing references
    LogMessage "STEP 9: Updating IDW drawing references"
    Call UpdateIDWReferences(invApp, destFolder)
    
    ' Step 10: Update iProperties for copied documents
    ' CRITICAL: This updates Part Number iProperty to match new filename
    LogMessage "STEP 10: Updating iProperties for copied documents"
    Call UpdateIPropertiesForCopiedDocuments(invApp)
    
    ' Step 11: Write mapping file
    LogMessage "STEP 11: Writing STEP_1_MAPPING.txt"
    Call WriteMappingFile(destFolder)
    
    ' Step 12: Validation scan
    LogMessage "STEP 12: Validation scan and destination logging"
    Call ValidateCloneAndLog(sourceDir, destFolder)

    LogMessage "=== PREFIX CLONER COMPLETED ==="
    Call StopLogging
    
    Dim summaryMsg
    summaryMsg = "PREFIX CLONE COMPLETED!" & vbCrLf & vbCrLf & _
                 "✅ Assembly copied to: " & destFolder & vbCrLf & _
                 "✅ " & g_CopiedFiles.Count & " files copied" & vbCrLf & _
                 "✅ Prefix replaced: " & g_OldPrefix & " -> " & g_NewPrefix & vbCrLf & _
                 "✅ References updated to local copies" & vbCrLf & _
                 "✅ iProperties updated" & vbCrLf & _
                 "✅ IDW files copied and updated" & vbCrLf & _
                 "✅ STEP_1_MAPPING.txt generated" & vbCrLf & vbCrLf & _
                 "The new assembly is now completely isolated!" & vbCrLf & vbCrLf & _
                 "Log: " & g_LogPath
    
    MsgBox summaryMsg, vbInformation, "Success!"
End Sub

' ==============================================================================
' PREFIX DETECTION AND REPLACEMENT
' ==============================================================================

Function DetectCommonPrefix(allParts, mainAsmName)
    ' Analyze filenames to detect the most common prefix
    ' Strategy: Find the longest common prefix among .ipt and .iam files
    
    LogMessage "PREFIX DETECT: Analyzing " & allParts.Count & " files for common prefix..."
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Collect all filenames (without extension)
    Dim fileNames
    Set fileNames = CreateObject("Scripting.Dictionary")
    
    ' Add main assembly name first
    Dim mainBaseName
    mainBaseName = fso.GetBaseName(mainAsmName)
    fileNames.Add mainBaseName, 1
    
    Dim key
    For Each key In allParts.Keys
        Dim fileName
        fileName = fso.GetBaseName(fso.GetFileName(key))
        If Not fileNames.Exists(fileName) Then
            fileNames.Add fileName, 1
        End If
    Next
    
    LogMessage "PREFIX DETECT: Collected " & fileNames.Count & " unique filenames"
    
    ' Try to find common prefix by looking at the main assembly name
    ' Usually the prefix is everything up to and including the last hyphen before the part code
    Dim prefix
    prefix = ""
    
    ' Strategy 1: Find prefix from main assembly name
    ' Look for pattern like "N1SCR04-780-" or "PLANT1-001-"
    Dim lastHyphen
    lastHyphen = InStrRev(mainBaseName, "-")
    
    If lastHyphen > 0 Then
        ' Check if there's a second-to-last hyphen
        Dim beforeLastHyphen
        beforeLastHyphen = Left(mainBaseName, lastHyphen - 1)
        Dim secondLastHyphen
        secondLastHyphen = InStrRev(beforeLastHyphen, "-")
        
        If secondLastHyphen > 0 Then
            ' Found pattern like "N1SCR04-780-B1" - prefix is "N1SCR04-780-"
            prefix = Left(mainBaseName, lastHyphen)
        Else
            ' Only one hyphen, prefix is everything up to and including it
            prefix = Left(mainBaseName, lastHyphen)
        End If
    End If
    
    ' Verify how many files match this prefix
    If prefix <> "" Then
        Dim matchCount
        matchCount = 0
        Dim names
        names = fileNames.Keys
        Dim i
        For i = 0 To UBound(names)
            If Left(UCase(names(i)), Len(prefix)) = UCase(prefix) Then
                matchCount = matchCount + 1
            End If
        Next
        
        LogMessage "PREFIX DETECT: Prefix '" & prefix & "' matches " & matchCount & " of " & fileNames.Count & " files"
    End If
    
    DetectCommonPrefix = prefix
End Function

Function GetPrefixFromUser(detectedPrefix)
    ' Ask user to confirm the old prefix and enter the new prefix
    
    Dim oldPrefixInput
    oldPrefixInput = InputBox("CURRENT PREFIX DETECTED" & vbCrLf & vbCrLf & _
                              "The following prefix was detected:" & vbCrLf & _
                              "  " & detectedPrefix & vbCrLf & vbCrLf & _
                              "Please confirm or modify the CURRENT prefix:" & vbCrLf & vbCrLf & _
                              "This is the prefix that will be REPLACED in all filenames." & vbCrLf & _
                              "(Include the trailing hyphen if applicable)", _
                              "Confirm Current Prefix", detectedPrefix)
    
    If oldPrefixInput = "" Then
        GetPrefixFromUser = False
        Exit Function
    End If
    
    g_OldPrefix = oldPrefixInput
    LogMessage "PREFIX: Old prefix confirmed: " & g_OldPrefix
    
    ' Now ask for new prefix
    Dim suggestedNew
    suggestedNew = SuggestNewPrefix(g_OldPrefix)
    
    Dim newPrefixInput
    newPrefixInput = InputBox("ENTER NEW PREFIX" & vbCrLf & vbCrLf & _
                              "Current prefix: " & g_OldPrefix & vbCrLf & vbCrLf & _
                              "Enter the NEW prefix to replace it:" & vbCrLf & vbCrLf & _
                              "Example:" & vbCrLf & _
                              "  " & g_OldPrefix & "B1.IPT  ->  " & suggestedNew & "B1.IPT" & vbCrLf & vbCrLf & _
                              "(Include the trailing hyphen if applicable)", _
                              "Enter New Prefix", suggestedNew)
    
    If newPrefixInput = "" Then
        GetPrefixFromUser = False
        Exit Function
    End If
    
    g_NewPrefix = newPrefixInput
    LogMessage "PREFIX: New prefix: " & g_NewPrefix
    
    ' Confirm with user
    Dim confirmResult
    confirmResult = MsgBox("CONFIRM PREFIX REPLACEMENT" & vbCrLf & vbCrLf & _
                           "Old Prefix: " & g_OldPrefix & vbCrLf & _
                           "New Prefix: " & g_NewPrefix & vbCrLf & vbCrLf & _
                           "Example transformation:" & vbCrLf & _
                           "  " & g_OldPrefix & "B1.IPT" & vbCrLf & _
                           "  becomes:" & vbCrLf & _
                           "  " & g_NewPrefix & "B1.IPT" & vbCrLf & vbCrLf & _
                           "Proceed with this prefix replacement?", _
                           vbYesNo + vbQuestion, "Confirm Prefix Change")
    
    If confirmResult = vbNo Then
        GetPrefixFromUser = False
        Exit Function
    End If
    
    GetPrefixFromUser = True
End Function

Function SuggestNewPrefix(oldPrefix)
    ' Try to suggest a new prefix by incrementing numbers
    ' e.g., "N1SCR04-780-" -> "N2SCR04-780-"
    
    Dim i
    Dim result
    result = ""
    
    ' Find the first number in the prefix and try to increment it
    For i = 1 To Len(oldPrefix)
        Dim c
        c = Mid(oldPrefix, i, 1)
        If IsNumeric(c) Then
            ' Found a number - try to increment it
            Dim numStart
            numStart = i
            Dim numEnd
            numEnd = i
            
            ' Find the end of the number
            Do While numEnd < Len(oldPrefix)
                If IsNumeric(Mid(oldPrefix, numEnd + 1, 1)) Then
                    numEnd = numEnd + 1
                Else
                    Exit Do
                End If
            Loop
            
            Dim numStr
            numStr = Mid(oldPrefix, numStart, numEnd - numStart + 1)
            Dim numVal
            numVal = CLng(numStr) + 1
            
            ' Pad with zeros if needed
            Dim newNumStr
            newNumStr = CStr(numVal)
            Do While Len(newNumStr) < Len(numStr)
                newNumStr = "0" & newNumStr
            Loop
            
            result = Left(oldPrefix, numStart - 1) & newNumStr & Mid(oldPrefix, numEnd + 1)
            Exit For
        End If
    Next
    
    If result = "" Then
        ' No number found, just append "2"
        If Right(oldPrefix, 1) = "-" Then
            result = Left(oldPrefix, Len(oldPrefix) - 1) & "2-"
        Else
            result = oldPrefix & "2"
        End If
    End If
    
    SuggestNewPrefix = result
End Function

Function ReplacePrefix(fileName)
    ' Replace the old prefix with new prefix in the filename
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim baseName
    baseName = fso.GetBaseName(fileName)
    Dim ext
    ext = fso.GetExtensionName(fileName)
    
    ' Check if the filename starts with the old prefix (case-insensitive)
    If Left(UCase(baseName), Len(g_OldPrefix)) = UCase(g_OldPrefix) Then
        ' Replace prefix
        Dim suffix
        suffix = Mid(baseName, Len(g_OldPrefix) + 1)
        ReplacePrefix = g_NewPrefix & suffix & "." & ext
    Else
        ' No match - keep original name
        ReplacePrefix = fileName
    End If
End Function

' ==============================================================================
' ASSEMBLY AND FILE DETECTION
' ==============================================================================

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
                 "Clone this assembly with prefix replacement?"
    
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

    Dim shell
    Set shell = CreateObject("Shell.Application")

    Dim startFolder
    startFolder = fso.GetParentFolderName(sourceDir)

    Dim folder
    Set folder = shell.BrowseForFolder(0, "Select DESTINATION folder for the cloned assembly:" & vbCrLf & _
                                          "Source: " & sourceDir & vbCrLf & vbCrLf & _
                                          "TIP: Click 'Make New Folder' to create a new destination", _
                                          &H0011, startFolder)

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

' ==============================================================================
' FILE COLLECTION
' ==============================================================================

Sub CollectAllReferencedParts(asmDoc, allParts)
    LogMessage "COLLECT: Scanning assembly for all referenced parts and sub-assemblies..."
    Call CollectPartsRecursively(asmDoc, allParts, "ROOT")
    LogMessage "COLLECT: Found " & allParts.Count & " unique files"
End Sub

Sub CollectPartsRecursively(asmDoc, allParts, level)
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
                    If Not allParts.Exists(fullPath) Then
                        Dim description
                        description = GetDescriptionFromIProperty(doc)
                        allParts.Add fullPath, description
                        LogMessage "COLLECT: PART " & fileName & " at " & level
                    End If
                    
                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    If Not allParts.Exists(fullPath) Then
                        allParts.Add fullPath, "SUB-ASSEMBLY"
                        LogMessage "COLLECT: SUB-ASSEMBLY " & fileName & " at " & level
                    End If
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

Sub CollectIDWFiles(sourceDir, allParts)
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

' ==============================================================================
' FILE COPYING WITH PREFIX REPLACEMENT
' ==============================================================================

Sub CopyAllFilesWithPrefixReplacement(allParts, destFolder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim partKeys
    partKeys = allParts.Keys
    
    Dim i
    For i = 0 To UBound(partKeys)
        Dim originalPath
        originalPath = partKeys(i)
        
        Dim originalFileName
        originalFileName = fso.GetFileName(originalPath)
        
        ' Replace prefix in filename
        Dim newFileName
        newFileName = ReplacePrefix(originalFileName)
        
        ' Compute relative path from source root to preserve folder structure
        Dim newPath
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
        
        ' Determine file type for logging and mapping
        Dim fileType
        If LCase(Right(originalFileName, 4)) = ".iam" Then
            fileType = "SUB-ASSEMBLY"
        ElseIf LCase(Right(originalFileName, 4)) = ".idw" Then
            fileType = "DRAWING"
        Else
            fileType = allParts.Item(originalPath)
            If fileType = "" Then fileType = "PART"
        End If
        
        ' Copy the file
        On Error Resume Next
        fso.CopyFile originalPath, newPath, True
        
        If Err.Number = 0 Then
            If originalFileName <> newFileName Then
                LogMessage "COPIED: " & originalFileName & " -> " & newFileName
            Else
                LogMessage "COPIED: " & originalFileName & " (no prefix match, kept original name)"
            End If
            g_CopiedFiles.Add originalPath, newPath
            Call AddMappingData(originalPath, newPath, fileType)
        Else
            LogMessage "ERROR: Could not copy " & originalFileName & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Sub

' ==============================================================================
' REFERENCE UPDATES
' ==============================================================================

Sub UpdateInMemoryAssemblyReferences(invApp, asmPath)
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
    
    ' Build lookup dictionaries
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
        pathLookup.Add LCase(key), g_CopiedFiles.Item(key)
        ' Track all NEW paths so we can skip them if already correct
        newPathSet.Add LCase(g_CopiedFiles.Item(key)), True
    Next
    
    LogMessage "IN-MEMORY UPDATE: Built lookup with " & fileNameLookup.Count & " filenames"
    
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
            
            ' Check if reference is ALREADY pointing to a new path (skip if so)
            If newPathSet.Exists(refPathLower) Then
                LogMessage "IN-MEMORY UPDATE: SKIP (already correct path): " & refFileName
            Else
                ' First try exact path match
                Dim newRefPath
                newRefPath = ""
                
                If pathLookup.Exists(refPathLower) Then
                    newRefPath = pathLookup.Item(refPathLower)
                ElseIf fileNameLookup.Exists(refFileName) Then
                    newRefPath = fileNameLookup.Item(refFileName)
                End If
                
                If newRefPath <> "" Then
                    If LCase(refPath) <> LCase(newRefPath) Then
                        LogMessage "IN-MEMORY UPDATE: REPLACING " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
                        
                        fd.ReplaceReference newRefPath
                        
                        If Err.Number = 0 Then
                            updatedCount = updatedCount + 1
                        Else
                            LogMessage "IN-MEMORY UPDATE: ReplaceReference failed: " & Err.Description
                            Err.Clear
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    LogMessage "IN-MEMORY UPDATE: Updated " & updatedCount & " references in " & GetFileNameFromPath(asmPath)
    
    Set fileNameLookup = Nothing
    Set pathLookup = Nothing
    Set newPathSet = Nothing
    Err.Clear
End Sub

Sub UpdateIDWReferences(invApp, destFolder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    ' Create ApprenticeServer for silent file manipulation
    Dim apprentice
    Set apprentice = CreateObject("Inventor.ApprenticeServerComponent")
    
    If Err.Number <> 0 Or apprentice Is Nothing Then
        LogMessage "IDW UPDATE: Could not create ApprenticeServer, falling back to Inventor"
        Err.Clear
        Call UpdateIDWReferencesWithInventor(invApp, destFolder)
        Exit Sub
    End If
    
    LogMessage "IDW UPDATE: Using ApprenticeServer for silent reference updates"
    
    Dim folder
    Set folder = fso.GetFolder(destFolder)
    
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            LogMessage "IDW UPDATE: Processing " & file.Name
            
            Dim appDoc
            Set appDoc = apprentice.Open(file.Path)
            
            If Err.Number = 0 And Not appDoc Is Nothing Then
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
                    
                    ' Check if we have a mapping for this file
                    If g_CopiedFiles.Exists(refPath) Then
                        Dim newRefPath
                        newRefPath = g_CopiedFiles.Item(refPath)
                        
                        LogMessage "IDW UPDATE: Replacing " & refFileName & " -> " & GetFileNameFromPath(newRefPath)
                        
                        fd.ReplaceReference newRefPath
                        
                        If Err.Number = 0 Then
                            updatedCount = updatedCount + 1
                        Else
                            LogMessage "IDW UPDATE: ERROR - " & Err.Description
                        End If
                        Err.Clear
                    Else
                        ' Try filename lookup
                        Dim origKey
                        For Each origKey In g_CopiedFiles.Keys
                            If LCase(GetFileNameFromPath(origKey)) = LCase(refFileName) Then
                                fd.ReplaceReference g_CopiedFiles.Item(origKey)
                                If Err.Number = 0 Then
                                    updatedCount = updatedCount + 1
                                End If
                                Err.Clear
                                Exit For
                            End If
                        Next
                    End If
                Next
                
                appDoc.SaveAs file.Path, False
                appDoc.Close
                LogMessage "IDW UPDATE: Saved " & file.Name & " (" & updatedCount & " references updated)"
            Else
                LogMessage "IDW UPDATE: Could not open " & file.Name & " - " & Err.Description
            End If
            Err.Clear
        End If
    Next
    
    Set apprentice = Nothing
End Sub

Sub UpdateIDWReferencesWithInventor(invApp, destFolder)
    LogMessage "IDW UPDATE (Inventor): Starting IDW update for destFolder: " & destFolder
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    ' Store original settings
    Dim originalResolveMode
    originalResolveMode = invApp.FileOptions.ResolveFileOption
    
    Dim originalSilentMode
    originalSilentMode = invApp.SilentOperation
    
    invApp.SilentOperation = True
    invApp.FileOptions.ResolveFileOption = 54275
    
    ' Close all documents first
    invApp.Documents.CloseAll
    
    ' Find and open the main assembly first
    Dim folder
    Set folder = fso.GetFolder(destFolder)
    
    ' Find any .iam file to open first
    Dim file
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".iam" Then
            Dim mainAsmDoc
            Set mainAsmDoc = invApp.Documents.Open(file.Path, False)
            If Not mainAsmDoc Is Nothing Then
                LogMessage "IDW UPDATE: Opened main assembly to load all parts"
            End If
            Exit For
        End If
    Next
    Err.Clear
    
    ' Now process IDW files
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".idw" Then
            LogMessage "IDW UPDATE (Inventor): Processing " & file.Name
            
            Dim idwDoc
            Set idwDoc = invApp.Documents.Open(file.Path, False)
            
            If Err.Number = 0 And Not idwDoc Is Nothing Then
                Dim fileDescriptors
                Set fileDescriptors = idwDoc.File.ReferencedFileDescriptors
                
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
                    
                    Dim newRefPath
                    newRefPath = ""
                    
                    If g_CopiedFiles.Exists(refPath) Then
                        newRefPath = g_CopiedFiles.Item(refPath)
                    Else
                        Dim origKey
                        For Each origKey In g_CopiedFiles.Keys
                            If LCase(GetFileNameFromPath(origKey)) = LCase(refFileName) Then
                                newRefPath = g_CopiedFiles.Item(origKey)
                                Exit For
                            End If
                        Next
                    End If
                    
                    If newRefPath <> "" Then
                        fd.ReplaceReference newRefPath
                        If Err.Number = 0 Then
                            updatedCount = updatedCount + 1
                        End If
                        Err.Clear
                    End If
                Next
                
                idwDoc.Save
                idwDoc.Close
                LogMessage "IDW UPDATE: Saved " & file.Name & " (" & updatedCount & " updated)"
            Else
                LogMessage "IDW UPDATE: Could not open " & file.Name
            End If
            Err.Clear
        End If
    Next
    
    invApp.FileOptions.ResolveFileOption = originalResolveMode
    invApp.SilentOperation = originalSilentMode
    Err.Clear
End Sub

' ==============================================================================
' MAPPING FILE GENERATION
' ==============================================================================

Sub AddMappingData(originalPath, newPath, fileType)
    ' Add mapping entry to the array for later writing
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim originalFile
    originalFile = fso.GetFileName(originalPath)
    Dim newFile
    newFile = fso.GetFileName(newPath)
    
    ' Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description
    Dim entry
    entry = originalPath & "|" & newPath & "|" & originalFile & "|" & newFile & "|" & fileType & "|PREFIX_CLONE"
    
    ' Resize array and add entry
    Dim arrSize
    On Error Resume Next
    arrSize = UBound(g_MappingData)
    If Err.Number <> 0 Then
        arrSize = -1
    End If
    Err.Clear
    On Error GoTo 0
    
    ReDim Preserve g_MappingData(arrSize + 1)
    g_MappingData(arrSize + 1) = entry
End Sub

Sub WriteMappingFile(destFolder)
    On Error Resume Next
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim mappingPath
    mappingPath = destFolder & "\STEP_1_MAPPING.txt"
    
    Dim mappingFile
    Set mappingFile = fso.CreateTextFile(mappingPath, True)
    
    If Err.Number <> 0 Then
        LogMessage "MAPPING: ERROR creating mapping file - " & Err.Description
        Exit Sub
    End If
    
    ' Write header
    mappingFile.WriteLine "# STEP 1 MAPPING FILE - Generated by PREFIX CLONER: " & Now()
    mappingFile.WriteLine "# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description"
    mappingFile.WriteLine "# Prefix Replacement: " & g_OldPrefix & " -> " & g_NewPrefix
    mappingFile.WriteLine "# Total mappings: " & (UBound(g_MappingData) + 1)
    mappingFile.WriteLine ""
    
    ' Write all mappings
    Dim i
    For i = 0 To UBound(g_MappingData)
        mappingFile.WriteLine g_MappingData(i)
    Next
    
    mappingFile.WriteLine ""
    mappingFile.WriteLine "# End of mapping file - " & (UBound(g_MappingData) + 1) & " mappings written"
    
    mappingFile.Close
    
    LogMessage "MAPPING: Written " & (UBound(g_MappingData) + 1) & " mappings to " & mappingPath
    Err.Clear
End Sub

' ==============================================================================
' iPROPERTY UPDATES
' ==============================================================================

Sub UpdateIPropertiesForCopiedDocuments(invApp)
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
            Dim doc
            Set doc = invApp.Documents.Open(newPath, False)
            
            If Err.Number <> 0 Or doc Is Nothing Then
                LogMessage "IPROP: ERROR opening " & GetFileNameFromPath(newPath) & ": " & Err.Description
                Err.Clear
            Else
                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")
                Dim oldFileName
                oldFileName = fso.GetBaseName(oldPath)
                Dim newFileName
                newFileName = fso.GetBaseName(newPath)
                
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
                
                ' Iterate through all property sets and replace old prefix/filename with new
                Dim propSet
                For Each propSet In doc.PropertySets
                    Dim prop
                    For Each prop In propSet
                        If Not prop.Value Is Nothing And VarType(prop.Value) = vbString Then
                            Dim valueStr
                            valueStr = prop.Value
                            
                            Dim newValue
                            newValue = valueStr
                            
                            ' Replace old prefix with new prefix
                            If g_OldPrefix <> "" Then
                                newValue = Replace(newValue, g_OldPrefix, g_NewPrefix, 1, -1, vbTextCompare)
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
                    Next
                Next
                
                If replaced Then
                    LogMessage "IPROP: Updated iProperties for " & GetFileNameFromPath(newPath)
                    updatedCount = updatedCount + 1
                End If
                
                doc.Save
                
                If Err.Number <> 0 Then
                    LogMessage "IPROP: ERROR saving " & GetFileNameFromPath(newPath) & ": " & Err.Description
                    Err.Clear
                End If
                
                doc.Close False
                Err.Clear
            End If
        End If
    Next
    
    LogMessage "IPROP: Completed - updated " & updatedCount & " documents"
    Err.Clear
End Sub

' ==============================================================================
' HELPER FUNCTIONS
' ==============================================================================

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

' ==============================================================================
' VALIDATION
' ==============================================================================

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
    
    LogMessage "VALIDATE: Checking copied mappings..."
    Dim missingCount
    missingCount = 0
    Dim key
    For Each key In g_CopiedFiles.Keys
        Dim copiedPath
        copiedPath = g_CopiedFiles.Item(key)
        If fso.FileExists(copiedPath) Then
            LogMessage "MAP OK: " & GetFileNameFromPath(key) & " -> " & GetFileNameFromPath(copiedPath)
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

' ==============================================================================
' LOGGING FUNCTIONS
' ==============================================================================

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
    
    g_LogPath = logsDir & "\PrefixCloner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFileNum = fso.CreateTextFile(g_LogPath, True)
End Sub

Sub StartDestinationLogging(destFolder)
    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(destFolder) Then Exit Sub
    
    g_DestLogPath = destFolder & "\PrefixCloner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
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
