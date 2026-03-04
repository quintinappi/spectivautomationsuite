Option Explicit

' =============================================================================
' iLOGIC PATCHER - DETAILING WORKFLOW STEP 2b: Fix iLogic Rules
' =============================================================================
' Author: Quintin de Bruin © 2025
' DETAILING WORKFLOW - STEP 2b: Run iLogic patcher
'
' This script fixes component name references in iLogic rules after renaming
' Run this as part of STEP 2 (Detailing) after running the main renamer (STEP 1)
'
' This script:
' 1. Reads STEP_0_AUDIT.txt to see what iLogic rules exist
' 2. Reads STEP_1_MAPPING.txt to get old → new component name mappings
' 3. Updates iLogic rules with new component names
' 4. Saves STEP_2_ILOGIC_PATCHES.txt with changes made
'
' WHEN TO USE:
' - After completing PART RENAMING (STEP 1)
' - As part of DETAILING WORKFLOW (STEP 2)
' - When iLogic rules reference old component names
' =============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_OldToNewMapping   ' Dictionary: old filename -> new filename (without extension)

' Inventor document type constants
Const kPartDocumentObject = 12290
Const kAssemblyDocumentObject = 12291

Sub MAIN()
    Call StartLogging
    LogMessage "=== ILOGIC PATCHER ==="
    LogMessage "Fixing component name references in iLogic rules after renaming"

    Dim result
    result = MsgBox("ILOGIC PATCHER - DETAILING STEP 2b" & vbCrLf & vbCrLf & _
                    "DETAILING WORKFLOW - STEP 2b: Run iLogic patcher" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Read STEP_0_AUDIT.txt to find iLogic rules" & vbCrLf & _
                    "2. Read STEP_1_MAPPING.txt for name mappings" & vbCrLf & _
                    "3. Update iLogic rules with new component names" & vbCrLf & _
                    "4. Save patch log to STEP_2_ILOGIC_PATCHES.txt" & vbCrLf & vbCrLf & _
                    "⚠️  Run this AFTER completing PART RENAMING (STEP 1)!" & vbCrLf & vbCrLf & _
                    "Make sure your assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Step 2b: iLogic Patcher")

    If result = vbNo Then
        LogMessage "User cancelled workflow"
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
    On Error GoTo 0

    ' Detect open assembly
    LogMessage "STEP 1: Detecting open assembly"
    Dim activeDoc
    Set activeDoc = DetectOpenAssembly(invApp)
    If activeDoc Is Nothing Then
        MsgBox "ERROR: No assembly is currently open in Inventor!", vbCritical
        Exit Sub
    End If

    ' Read mapping file
    LogMessage "STEP 2: Reading STEP_1_MAPPING.txt for name mappings"
    If Not ReadMappingFile(activeDoc) Then
        MsgBox "ERROR: Could not read STEP_1_MAPPING.txt file!", vbCritical
        Exit Sub
    End If

    ' Create patch log file
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim asmDir
    asmDir = fso.GetParentFolderName(activeDoc.FullFileName)
    Dim patchLogPath
    patchLogPath = asmDir & "\STEP_2_ILOGIC_PATCHES.txt"

    Dim patchLogFile
    Set patchLogFile = fso.CreateTextFile(patchLogPath, True)

    patchLogFile.WriteLine "================================================================"
    patchLogFile.WriteLine "STEP 2: ILOGIC RULE PATCHING"
    patchLogFile.WriteLine "================================================================"
    patchLogFile.WriteLine "Created: " & Now()
    patchLogFile.WriteLine "Assembly: " & activeDoc.DisplayName
    patchLogFile.WriteLine "================================================================"
    patchLogFile.WriteLine ""

    ' Patch iLogic rules in assembly
    LogMessage "STEP 3: Patching iLogic rules in assembly"
    Dim totalPatches
    totalPatches = 0
    Call PatchILogicRules(invApp, activeDoc, patchLogFile, "ASSEMBLY", totalPatches)

    ' Patch iLogic rules in referenced parts
    LogMessage "STEP 4: Patching iLogic rules in referenced parts"
    patchLogFile.WriteLine ""
    patchLogFile.WriteLine "REFERENCED PARTS:"
    patchLogFile.WriteLine "================================================================"

    Dim asmCompDef
    Set asmCompDef = activeDoc.ComponentDefinition

    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    Dim occ
    For Each occ In asmCompDef.Occurrences
        On Error Resume Next
        Set doc = occ.Definition.Document
        If Err.Number = 0 And Not doc Is Nothing Then
            If doc.DocumentType = kPartDocumentObject Then
                Dim uniqueKey
                uniqueKey = doc.FullFileName
                If Not uniqueParts.Exists(uniqueKey) Then
                    uniqueParts.Add uniqueKey, doc.DisplayName
                End If
            End If
        End If
        Err.Clear
    Next

    Dim partKey
    For Each partKey In uniqueParts.Keys
        Dim partPath
        partPath = CStr(partKey)
        Dim partDoc
        On Error Resume Next
        Set partDoc = invApp.Documents.Open(partPath, False)
        If Err.Number = 0 And Not partDoc Is Nothing Then
            Call PatchILogicRules(invApp, partDoc, patchLogFile, "PART - " & uniqueParts(partKey), totalPatches)
            partDoc.Close True
        End If
        Err.Clear
    Next

    patchLogFile.WriteLine ""
    patchLogFile.WriteLine "================================================================"
    patchLogFile.WriteLine "PATCHING COMPLETED"
    patchLogFile.WriteLine "Total patches made: " & totalPatches
    patchLogFile.WriteLine "================================================================"

    patchLogFile.Close

    LogMessage "PATCHING COMPLETED"
    LogMessage "Total patches made: " & totalPatches
    LogMessage "Patch log saved to: " & patchLogPath

    MsgBox "ILOGIC PATCHING COMPLETED!" & vbCrLf & vbCrLf & _
           "Total patches made: " & totalPatches & vbCrLf & vbCrLf & _
           "Patch log: " & patchLogPath & vbCrLf & vbCrLf & _
           "Check the log file for details.", vbInformation, "Success!"

    Call StopLogging
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

    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        LogMessage "File extension is not .iam: " & activeDoc.FullFileName
        MsgBox "Not an assembly file!" & vbCrLf & vbCrLf & _
               "Current file: " & activeDoc.DisplayName & vbCrLf & vbCrLf & _
               "Please open an assembly (.iam) file.", vbExclamation
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If

    LogMessage "DETECTED: Assembly - " & activeDoc.DisplayName
    Set DetectOpenAssembly = activeDoc
    Err.Clear
End Function

Function ReadMappingFile(asmDoc)
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim asmDir
    asmDir = fso.GetParentFolderName(asmDoc.FullFileName)
    Dim mappingPath
    mappingPath = asmDir & "\STEP_1_MAPPING.txt"

    If Not fso.FileExists(mappingPath) Then
        LogMessage "ERROR: STEP_1_MAPPING.txt not found: " & mappingPath
        ReadMappingFile = False
        Exit Function
    End If

    LogMessage "Reading mapping file: " & mappingPath

    Set g_OldToNewMapping = CreateObject("Scripting.Dictionary")

    Dim mappingFile
    Set mappingFile = fso.OpenTextFile(mappingPath, 1) ' 1 = ForReading

    Dim line
    Dim lineCount
    lineCount = 0

    On Error Resume Next
    While Not mappingFile.AtEndOfStream
        line = mappingFile.ReadLine
        lineCount = lineCount + 1

        ' Skip comments and empty lines
        If Trim(line) = "" Or Left(Trim(line), 1) = "#" Then
            ' Skip this line
        Else
            ' Parse: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description
            Dim parts
            parts = Split(line, "|")

            If UBound(parts) >= 3 Then
                Dim oldFile
                Dim newFile
                oldFile = parts(2) ' OriginalFile (with extension)
                newFile = parts(3) ' NewFile (with extension)

                ' Strip extensions for iLogic matching
                Dim oldNameNoExt
                Dim newNameNoExt
                oldNameNoExt = fso.GetBaseName(oldFile)
                newNameNoExt = fso.GetBaseName(newFile)

                ' Add to mapping dictionary
                If Not g_OldToNewMapping.Exists(oldNameNoExt) Then
                    g_OldToNewMapping.Add oldNameNoExt, newNameNoExt
                    LogMessage "MAPPING: " & oldNameNoExt & " -> " & newNameNoExt
                End If
            End If
        End If

        Err.Clear
    Wend

    mappingFile.Close

    LogMessage "Read " & g_OldToNewMapping.Count & " file name mappings from " & lineCount & " lines"

    ReadMappingFile = True
    Err.Clear
    On Error GoTo 0
End Function

Sub PatchILogicRules(invApp, doc, patchLogFile, docTypeLabel, ByRef totalPatches)
    ' Patch iLogic rules in a document using AttributeSet method
    On Error Resume Next

    Dim attrSets
    Set attrSets = doc.AttributeSets

    Dim ruleCount
    ruleCount = 0

    Dim attrSet
    For Each attrSet In attrSets
        Dim setName
        setName = attrSet.Name

        ' Check for iLogic rule patterns
        If Left(setName, 10) = "iLogicRule" Or _
           Left(setName, 6) = "iLogic" Or _
           InStr(1, setName, "Rule", vbTextCompare) > 0 Then

            Dim ruleText
            ruleText = ""
            Dim ruleName
            ruleName = setName

            ' Get rule text from attributes
            Dim attr
            Dim textAttrName
            For Each attr In attrSet
                On Error Resume Next
                Dim attrName
                attrName = attr.Name

                ' Look for the iLogicRuleText attribute (actual iLogic storage)
                If attrName = "iLogicRuleText" Then
                    ruleText = CStr(attr.Value)
                    textAttrName = attrName
                    LogMessage "  Found text in attribute: " & attrName & " (length: " & Len(ruleText) & ")"
                End If

                If attrName = "name" Or attrName = "iLogicRuleName" Then
                    ruleName = CStr(attr.Value)
                    LogMessage "  Found name: " & attrName & " = " & ruleName
                End If
                Err.Clear
            Next

            If ruleText <> "" Then
                ruleCount = ruleCount + 1

                Dim originalRuleText
                originalRuleText = ruleText

                ' Apply patches based on mapping
                Dim oldNameNoExt
                For Each oldNameNoExt In g_OldToNewMapping.Keys
                    Dim newNameNoExt
                    newNameNoExt = g_OldToNewMapping(oldNameNoExt)

                    ' Replace old component names with new ones
                    ' Pattern matches: "OldName:1", "OldName:2", etc.
                    Dim oldPattern
                    oldPattern = """" & oldNameNoExt & ":"

                    Dim newPattern
                    newPattern = """" & newNameNoExt & ":"

                    ruleText = Replace(ruleText, oldPattern, newPattern)

                    Err.Clear
                Next

                ' If rule was modified, save the changes
                If ruleText <> originalRuleText Then
                    totalPatches = totalPatches + 1

                    patchLogFile.WriteLine ""
                    patchLogFile.WriteLine "RULE PATCHED: " & ruleName
                    patchLogFile.WriteLine "LOCATION: " & docTypeLabel
                    patchLogFile.WriteLine "----------------------------------------------------------------"

                    ' Find what was changed
                    Dim oldNamesFound
                    Set oldNamesFound = CreateObject("Scripting.Dictionary")

                    Dim oldName
                    For Each oldName In g_OldToNewMapping.Keys
                        Dim checkPattern
                        checkPattern = """" & oldName & ":"

                        If InStr(originalRuleText, checkPattern) > 0 Then
                            oldNamesFound.Add oldName, g_OldToNewMapping(oldName)
                        End If
                    Next

                    If oldNamesFound.Count > 0 Then
                        patchLogFile.WriteLine "CHANGES MADE:"
                        Dim oldKey
                        For Each oldKey In oldNamesFound.Keys
                            patchLogFile.WriteLine "  " & oldKey & " -> " & oldNamesFound(oldKey)
                        Next
                    End If

                    patchLogFile.WriteLine "----------------------------------------------------------------"

                    ' Save the patched rule text back to the attribute
                    Dim saveAttr
                    For Each saveAttr In attrSet
                        On Error Resume Next
                        Dim saveAttrName
                        saveAttrName = saveAttr.Name

                        If saveAttrName = textAttrName Then
                            saveAttr.Value = ruleText
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Err.Clear
    Next

    If ruleCount = 0 Then
        patchLogFile.WriteLine "NO ILOGIC RULES in " & docTypeLabel
    Else
        LogMessage "PATCHING: Checked " & ruleCount & " iLogic rules in " & docTypeLabel
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub StartLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempPath
    tempPath = fso.GetSpecialFolder(2) ' Temporary folder

    Dim logName
    logName = "iLogic_Patcher_" & Year(Now) & "-" & _
              Right("0" & Month(Now), 2) & "-" & _
              Right("0" & Day(Now), 2) & "_" & _
              Right("0" & Hour(Now), 2) & "-" & _
              Right("0" & Minute(Now), 2) & "-" & _
              Right("0" & Second(Now), 2) & ".txt"

    g_LogPath = tempPath & "\" & logName

    Set g_LogFileNum = fso.CreateTextFile(g_LogPath, True)

    LogMessage "=== ILOGIC PATCHER LOG STARTED ==="
End Sub

Sub LogMessage(message)
    On Error Resume Next
    Dim timestamp
    timestamp = Now()
    g_LogFileNum.WriteLine timestamp & " | " & message
    g_LogFileNum.Flush
    Err.Clear
End Sub

Sub StopLogging()
    On Error Resume Next
    LogMessage "=== ILOGIC PATCHER LOG ENDED ==="
    g_LogFileNum.Close
    Err.Clear
End Sub

Call MAIN()
