' Fix_Derived_Parts.vbs
' =============================================================================
' EXPERIMENTAL: Fix derived parts in an already-cloned assembly
' 
' This script will:
' 1. Scan the active assembly for derived parts with EXTERNAL base components
' 2. Copy those base files into the assembly folder
' 3. Rename them with a user-specified prefix
' 4. Update the derived parts to point to the new local copies
'
' Run with Inventor open and your CLONED assembly (.iam) active
' =============================================================================
Option Explicit

Const kPartDocumentObject = 12290
Const kAssemblyDocumentObject = 12291

Dim invApp, activeDoc, asmFolder
Dim fso, logFile, logPath
Dim baseFilesToCopy  ' Dictionary: originalPath -> newPath
Dim derivedPartsToFix ' Dictionary: derivedPartPath -> Array of {derivedPartDoc, originalBase, newBase}
Dim userPrefix
Dim fixCount, copyCount

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
Set baseFilesToCopy = CreateObject("Scripting.Dictionary")
baseFilesToCopy.CompareMode = vbTextCompare

Set derivedPartsToFix = CreateObject("Scripting.Dictionary")
derivedPartsToFix.CompareMode = vbTextCompare

logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Fix_Derived_Log.txt"

' Connect to Inventor
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    MsgBox "Inventor is not running. Please open Inventor with a cloned assembly first.", vbCritical, "Error"
    WScript.Quit
End If
On Error GoTo 0

Set activeDoc = invApp.ActiveDocument
If activeDoc Is Nothing Then
    MsgBox "No document is open in Inventor.", vbCritical, "Error"
    WScript.Quit
End If

If activeDoc.DocumentType <> kAssemblyDocumentObject Then
    MsgBox "Active document is not an assembly (.iam). Please open your cloned assembly.", vbCritical, "Error"
    WScript.Quit
End If

asmFolder = fso.GetParentFolderName(activeDoc.FullFileName)

' Try to detect prefix from existing files in the folder
Dim detectedPrefix
detectedPrefix = DetectPrefixFromFolder(asmFolder)

' Get prefix from user
Dim promptMsg
promptMsg = "Enter the PREFIX for the base component copies:" & vbCrLf & vbCrLf & _
            "Assembly folder: " & asmFolder & vbCrLf & vbCrLf

If detectedPrefix <> "" Then
    promptMsg = promptMsg & "Detected prefix from existing files: " & detectedPrefix & vbCrLf & vbCrLf
End If

userPrefix = InputBox(promptMsg, "Derived Parts Fixer", detectedPrefix)

If Trim(userPrefix) = "" Then
    MsgBox "No prefix entered. Operation cancelled.", vbInformation, "Cancelled"
    WScript.Quit
End If

userPrefix = Trim(userPrefix)

' Confirm with user
Dim confirmMsg
confirmMsg = "DERIVED PARTS FIXER" & vbCrLf & vbCrLf & _
             "Assembly: " & activeDoc.DisplayName & vbCrLf & _
             "Folder: " & asmFolder & vbCrLf & _
             "Prefix: " & userPrefix & vbCrLf & vbCrLf & _
             "This will:" & vbCrLf & _
             "1. Find derived parts with EXTERNAL base files" & vbCrLf & _
             "2. Copy base files into assembly folder" & vbCrLf & _
             "3. Rename with prefix: " & userPrefix & "-filename.ipt" & vbCrLf & _
             "4. Update derived parts to use local copies" & vbCrLf & vbCrLf & _
             "Continue?"

If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Operation") <> vbYes Then
    MsgBox "Operation cancelled.", vbInformation, "Cancelled"
    WScript.Quit
End If

' Create log file
Set logFile = fso.CreateTextFile(logPath, True)
WriteLog "=========================================="
WriteLog " DERIVED PARTS FIXER - EXECUTION LOG"
WriteLog "=========================================="
WriteLog ""
WriteLog "Assembly: " & activeDoc.FullFileName
WriteLog "Folder: " & asmFolder
WriteLog "Prefix: " & userPrefix
WriteLog "Start Time: " & Now()
WriteLog ""

' Enable silent operation
invApp.SilentOperation = True

fixCount = 0
copyCount = 0

' PHASE 1: Scan and collect all external base files
WriteLog "=========================================="
WriteLog " PHASE 1: SCANNING FOR EXTERNAL BASE FILES"
WriteLog "=========================================="
WriteLog ""

ScanForExternalBaseFiles

WriteLog ""
WriteLog "External base files found: " & baseFilesToCopy.Count
WriteLog "Same-folder derived parts to fix: " & derivedPartsToFix.Count
WriteLog ""

If baseFilesToCopy.Count = 0 And derivedPartsToFix.Count = 0 Then
    WriteLog "No external base files or same-folder derived parts found - nothing to fix!"
    CleanupAndExit "No derived parts need fixing.", False
End If

' Show what we found
Dim key
If baseFilesToCopy.Count > 0 Then
    WriteLog "External files to copy:"
    For Each key In baseFilesToCopy.Keys
        WriteLog "  FROM: " & key
        WriteLog "  TO:   " & baseFilesToCopy(key)
        WriteLog ""
    Next
End If

If derivedPartsToFix.Count > 0 Then
    WriteLog "Same-folder derived parts to update:"
    For Each key In derivedPartsToFix.Keys
        Dim fixInfo
        fixInfo = derivedPartsToFix(key)
        WriteLog "  Derived Part: " & fso.GetFileName(key)
        WriteLog "  FROM: " & fso.GetFileName(fixInfo(1))
        WriteLog "  TO:   " & fso.GetFileName(fixInfo(2))
        WriteLog ""
    Next
End If

' PHASE 2: Copy base files
WriteLog "=========================================="
WriteLog " PHASE 2: COPYING BASE FILES"
WriteLog "=========================================="
WriteLog ""

CopyBaseFiles

WriteLog ""
WriteLog "Files copied: " & copyCount
WriteLog ""

' PHASE 3: Update derived part references
WriteLog "=========================================="
WriteLog " PHASE 3: UPDATING DERIVED PART REFERENCES"
WriteLog "=========================================="
WriteLog ""

UpdateDerivedReferences

WriteLog ""
WriteLog "Derived parts updated: " & fixCount
WriteLog ""

' PHASE 4: Save all modified documents
WriteLog "=========================================="
WriteLog " PHASE 4: SAVING MODIFIED DOCUMENTS"
WriteLog "=========================================="
WriteLog ""

SaveAllModified

' Done
WriteLog ""
WriteLog "=========================================="
WriteLog " COMPLETE"
WriteLog "=========================================="
WriteLog "End Time: " & Now()
WriteLog "Base files copied: " & copyCount
WriteLog "Derived parts fixed: " & fixCount

CleanupAndExit "Operation complete!" & vbCrLf & vbCrLf & _
               "Base files copied: " & copyCount & vbCrLf & _
               "Derived parts fixed: " & fixCount & vbCrLf & vbCrLf & _
               "Please check the log for details.", True

' =============================================================================
' SUBROUTINES
' =============================================================================

Sub ScanForExternalBaseFiles()
    Dim doc, partDoc
    
    For Each doc In activeDoc.AllReferencedDocuments
        If doc.DocumentType = kPartDocumentObject Then
            CheckPartForExternalBase doc
        End If
    Next
End Sub

Sub CheckPartForExternalBase(partDoc)
    On Error Resume Next
    
    Dim partDef, refComps, derivedParts
    Dim derivedComp, docDesc, basePath, baseFolder
    Dim partPath, partName
    
    partPath = partDoc.FullFileName
    partName = fso.GetFileName(partPath)
    
    Set partDef = partDoc.ComponentDefinition
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Set refComps = partDef.ReferenceComponents
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Set derivedParts = refComps.DerivedPartComponents
    If Err.Number <> 0 Or derivedParts Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    If derivedParts.Count = 0 Then
        Exit Sub
    End If
    
    Dim i
    For i = 1 To derivedParts.Count
        Set derivedComp = derivedParts.Item(i)
        
        If derivedComp.LinkedToFile Then
            Set docDesc = derivedComp.ReferencedDocumentDescriptor
            If Not docDesc Is Nothing Then
                basePath = docDesc.FullDocumentName
                baseFolder = fso.GetParentFolderName(basePath)

                ' Check if base file is EXTERNAL (not in assembly folder)
                If LCase(baseFolder) <> LCase(asmFolder) Then
                    WriteLog "EXTERNAL BASE FOUND:"
                    WriteLog "  Derived Part: " & partName
                    WriteLog "  Base Component: " & fso.GetFileName(basePath)
                    WriteLog "  Base Location: " & basePath

                    ' Add to copy list if not already there
                    If Not baseFilesToCopy.Exists(basePath) Then
                        Dim newBaseName, newBasePath
                        newBaseName = userPrefix & "-" & fso.GetFileName(basePath)
                        newBasePath = asmFolder & "\" & newBaseName
                        baseFilesToCopy.Add basePath, newBasePath
                        WriteLog "  Will copy as: " & newBaseName
                    Else
                        WriteLog "  (Already in copy list)"
                    End If
                    WriteLog ""
                Else
                    ' SAME-FOLDER base file - check if it needs to be updated to new renamed file
                    WriteLog "SAME-FOLDER DERIVED PART FOUND:"
                    WriteLog "  Derived Part: " & partName
                    WriteLog "  Current Base Component: " & fso.GetFileName(basePath)

                    ' Check if mapping file exists for this assembly
                    Dim mappingPath
                    mappingPath = asmFolder & "\STEP_1_MAPPING.txt"

                    If fso.FileExists(mappingPath) Then
                        ' Read mapping file to find new base file name
                        Dim mappingFile
                        Set mappingFile = fso.OpenTextFile(mappingPath, 1)

                        Dim baseFileName
                        baseFileName = fso.GetBaseName(basePath)

                        Dim newBaseFileForDerived
                        newBaseFileForDerived = ""

                        Dim mapLine, mapParts, originalFile, newFile

                        While Not mappingFile.AtEndOfStream
                            mapLine = mappingFile.ReadLine

                            ' Skip comments and empty lines
                            If Trim(mapLine) <> "" And Left(Trim(mapLine), 1) <> "#" Then
                                ' Parse: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description
                                mapParts = Split(mapLine, "|")

                                If UBound(mapParts) >= 3 Then
                                    originalFile = fso.GetBaseName(mapParts(2))

                                    ' Check if this mapping matches our current base file
                                    If originalFile = baseFileName Then
                                        newFile = fso.GetBaseName(mapParts(3))
                                        newBaseFileForDerived = asmFolder & "\" & newFile & ".ipt"
                                        WriteLog "  MAPPING FOUND: " & originalFile & " -> " & newFile
                                        Exit While
                                    End If
                                End If
                            End If
                        Wend

                        mappingFile.Close

                        ' Add to derived parts to fix list if mapping found
                        If newBaseFileForDerived <> "" Then
                            If Not derivedPartsToFix.Exists(partPath) Then
                                derivedPartsToFix.Add partPath, Array(partDoc, basePath, newBaseFileForDerived)
                                WriteLog "  WILL UPDATE TO: " & fso.GetFileName(newBaseFileForDerived)
                            Else
                                WriteLog "  (Already in fix list)"
                            End If
                            WriteLog ""
                        Else
                            WriteLog "  WARNING: No mapping found for base file - skipping"
                            WriteLog ""
                        End If
                    Else
                        WriteLog "  WARNING: No STEP_1_MAPPING.txt found - cannot fix same-folder derived parts"
                        WriteLog ""
                    End If
                End If
            End If
        End If
    Next
    
    On Error GoTo 0
End Sub

Sub CopyBaseFiles()
    Dim srcPath, destPath
    
    For Each srcPath In baseFilesToCopy.Keys
        destPath = baseFilesToCopy(srcPath)
        
        WriteLog "Copying: " & fso.GetFileName(srcPath)
        WriteLog "  From: " & srcPath
        WriteLog "  To: " & destPath
        
        On Error Resume Next
        
        If fso.FileExists(srcPath) Then
            If fso.FileExists(destPath) Then
                WriteLog "  WARNING: Destination already exists - skipping copy"
            Else
                fso.CopyFile srcPath, destPath, False
                If Err.Number = 0 Then
                    WriteLog "  SUCCESS: File copied"
                    copyCount = copyCount + 1
                Else
                    WriteLog "  ERROR: " & Err.Description
                    Err.Clear
                End If
            End If
        Else
            WriteLog "  ERROR: Source file not found!"
        End If
        
        On Error GoTo 0
        WriteLog ""
    Next
End Sub

Sub UpdateDerivedReferences()
    Dim doc, partDoc
    
    For Each doc In activeDoc.AllReferencedDocuments
        If doc.DocumentType = kPartDocumentObject Then
            UpdatePartDerivedRefs doc
        End If
    Next
End Sub

Sub UpdatePartDerivedRefs(partDoc)
    On Error Resume Next
    
    Dim partDef, refComps, derivedParts
    Dim derivedComp, docDesc, basePath, baseFolder
    Dim partPath, partName
    Dim newBasePath
    
    partPath = partDoc.FullFileName
    partName = fso.GetFileName(partPath)
    
    Set partDef = partDoc.ComponentDefinition
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Set refComps = partDef.ReferenceComponents
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Set derivedParts = refComps.DerivedPartComponents
    If Err.Number <> 0 Or derivedParts Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    If derivedParts.Count = 0 Then
        Exit Sub
    End If
    
    Dim i
    For i = 1 To derivedParts.Count
        Set derivedComp = derivedParts.Item(i)
        
        If derivedComp.LinkedToFile Then
            Set docDesc = derivedComp.ReferencedDocumentDescriptor
            If Not docDesc Is Nothing Then
                basePath = docDesc.FullDocumentName
                
                ' Check if this base file was copied
                If baseFilesToCopy.Exists(basePath) Then
                    newBasePath = baseFilesToCopy(basePath)

                    WriteLog "Updating derived reference in: " & partName
                    WriteLog "  Derived Component: " & derivedComp.Name
                    WriteLog "  Old Base: " & basePath
                    WriteLog "  New Base: " & newBasePath

                    ' Use Replace method to update the reference
                    derivedComp.Replace newBasePath, Nothing

                    If Err.Number = 0 Then
                        WriteLog "  SUCCESS: Reference updated"
                        fixCount = fixCount + 1
                    Else
                        WriteLog "  ERROR: " & Err.Description
                        Err.Clear
                    End If
                    WriteLog ""
                ' Check if this is a same-folder derived part that needs updating
                ElseIf derivedPartsToFix.Exists(partPath) Then
                    Dim fixInfo
                    fixInfo = derivedPartsToFix(partPath)

                    ' Verify that this derived component is the one that needs fixing
                    If basePath = fixInfo(1) Then
                        newBasePath = fixInfo(2)

                        WriteLog "Updating same-folder derived reference in: " & partName
                        WriteLog "  Derived Component: " & derivedComp.Name
                        WriteLog "  Old Base: " & basePath
                        WriteLog "  New Base: " & newBasePath

                        ' Use Replace method to update the reference
                        derivedComp.Replace newBasePath, Nothing

                        If Err.Number = 0 Then
                            WriteLog "  SUCCESS: Reference updated"
                            fixCount = fixCount + 1
                        Else
                            WriteLog "  ERROR: " & Err.Description
                            Err.Clear
                        End If
                        WriteLog ""
                    End If
                End If
            End If
        End If
    Next
    
    On Error GoTo 0
End Sub

Sub SaveAllModified()
    Dim doc
    
    ' Save all referenced documents that are dirty
    For Each doc In activeDoc.AllReferencedDocuments
        If doc.Dirty Then
            WriteLog "Saving: " & fso.GetFileName(doc.FullFileName)
            On Error Resume Next
            doc.Save
            If Err.Number <> 0 Then
                WriteLog "  ERROR saving: " & Err.Description
                Err.Clear
            Else
                WriteLog "  Saved successfully"
            End If
            On Error GoTo 0
        End If
    Next
    
    ' Save main assembly if dirty
    If activeDoc.Dirty Then
        WriteLog "Saving assembly: " & fso.GetFileName(activeDoc.FullFileName)
        On Error Resume Next
        activeDoc.Save
        If Err.Number <> 0 Then
            WriteLog "  ERROR saving assembly: " & Err.Description
            Err.Clear
        Else
            WriteLog "  Assembly saved successfully"
        End If
        On Error GoTo 0
    End If
End Sub

Sub WriteLog(msg)
    logFile.WriteLine msg
    WScript.Echo msg
End Sub

Sub CleanupAndExit(msg, openLog)
    invApp.SilentOperation = False
    logFile.Close
    
    MsgBox msg, vbInformation, "Derived Parts Fixer"
    
    If openLog Then
        Dim shell
        Set shell = CreateObject("WScript.Shell")
        shell.Run """" & logPath & """", 1, False
    End If
    
    WScript.Quit
End Sub

' Detect prefix from existing files in folder (looks for pattern PREFIX-filename.ipt/iam)
Function DetectPrefixFromFolder(folderPath)
    On Error Resume Next
    DetectPrefixFromFolder = ""
    
    Dim folder, file, fileName, dashPos
    Dim prefixCounts, prefix, maxCount, bestPrefix
    
    Set prefixCounts = CreateObject("Scripting.Dictionary")
    prefixCounts.CompareMode = vbTextCompare
    
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files
        fileName = file.Name
        
        ' Only check Inventor files
        If LCase(fso.GetExtensionName(fileName)) = "ipt" Or _
           LCase(fso.GetExtensionName(fileName)) = "iam" Then
            
            ' Look for dash in filename (PREFIX-rest.ext)
            dashPos = InStr(fileName, "-")
            If dashPos > 1 Then
                prefix = Left(fileName, dashPos - 1)
                
                ' Only consider if prefix looks like a code (has letters/numbers, not too long)
                If Len(prefix) <= 20 And Len(prefix) >= 2 Then
                    If prefixCounts.Exists(prefix) Then
                        prefixCounts(prefix) = prefixCounts(prefix) + 1
                    Else
                        prefixCounts.Add prefix, 1
                    End If
                End If
            End If
        End If
    Next
    
    ' Find most common prefix
    maxCount = 0
    bestPrefix = ""
    
    Dim key
    For Each key In prefixCounts.Keys
        If prefixCounts(key) > maxCount Then
            maxCount = prefixCounts(key)
            bestPrefix = key
        End If
    Next
    
    ' Only return if found in multiple files
    If maxCount >= 2 Then
        DetectPrefixFromFolder = bestPrefix
    End If
    
    On Error GoTo 0
End Function
