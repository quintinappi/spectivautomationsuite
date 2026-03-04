Option Explicit

' ==============================================================================
' iLOGIC SCANNER - Scan Assembly for iLogic Rules and Properties
' ==============================================================================
' Author: Quintin de Bruin © 2025
' 
' This script:
' 1. Detects currently open assembly/part in Inventor
' 2. Connects to the iLogic Add-In API
' 3. Lists all iLogic rules embedded in the document
' 4. Shows rule properties and source code
' 5. Optionally exports rules to external text files
' ==============================================================================

Const ILOGIC_ADDIN_GUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

Dim g_LogFile       ' TextStream object for logging
Dim g_LogPath
Dim g_OutputFolder
Dim g_fso           ' Global FileSystemObject

Call ILOGIC_SCANNER_MAIN()

Sub ILOGIC_SCANNER_MAIN()
    Call StartLogging
    LogMessage "=== iLOGIC SCANNER ==="
    LogMessage "Scan document for iLogic rules and export source code"
    
    Dim result
    result = MsgBox("iLOGIC SCANNER" & vbCrLf & vbCrLf & _
                    "This tool will:" & vbCrLf & _
                    "1. Detect your currently open document" & vbCrLf & _
                    "2. Scan for embedded iLogic rules" & vbCrLf & _
                    "3. Display rule names and properties" & vbCrLf & _
                    "4. Show rule source code" & vbCrLf & _
                    "5. Optionally export rules to text files" & vbCrLf & vbCrLf & _
                    "Make sure your document is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "iLogic Scanner")
    
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
               "Please start Inventor and open your document first.", vbCritical
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    LogMessage "SUCCESS: Connected to Inventor instance"
    
    ' Step 1: Detect open document
    LogMessage "STEP 1: Detecting open document"
    Dim sourceDoc
    Set sourceDoc = DetectOpenDocument(invApp)
    If sourceDoc Is Nothing Then
        MsgBox "ERROR: No document is currently open in Inventor!" & vbCrLf & _
               "Please open a part (.ipt) or assembly (.iam) first.", vbCritical
        Exit Sub
    End If
    
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourceDir
    sourceDir = fso.GetParentFolderName(sourceDoc.FullFileName)
    Dim sourceFileName
    sourceFileName = fso.GetBaseName(sourceDoc.FullFileName)
    
    ' Step 2: Connect to iLogic Add-In
    LogMessage "STEP 2: Connecting to iLogic Add-In"
    Dim iLogicAddin, iLogicAuto
    Set iLogicAddin = GetILogicAddin(invApp)
    
    If iLogicAddin Is Nothing Then
        MsgBox "ERROR: Could not find iLogic Add-In!" & vbCrLf & vbCrLf & _
               "The iLogic Add-In may not be installed or enabled.", vbCritical
        Exit Sub
    End If
    
    ' Activate the add-in if needed and get automation interface
    On Error Resume Next
    iLogicAddin.Activate
    Set iLogicAuto = iLogicAddin.Automation
    
    If Err.Number <> 0 Or iLogicAuto Is Nothing Then
        LogMessage "ERROR: Could not get iLogic Automation interface: " & Err.Description
        MsgBox "ERROR: Could not access iLogic Automation!" & vbCrLf & _
               Err.Description, vbCritical
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    LogMessage "SUCCESS: Connected to iLogic Automation interface"
    
    ' Step 3: Get rules from document using AttributeSets (where iLogic stores rules)
    LogMessage "STEP 3: Scanning for iLogic rules via AttributeSets"
    
    ' iLogic stores rules in AttributeSets with specific naming
    ' The attribute set name is "iLogicRule_" + ruleName
    
    Dim attrSets
    Set attrSets = sourceDoc.AttributeSets
    
    Dim ruleCount
    ruleCount = 0
    Dim ruleInfo
    ruleInfo = ""
    Dim ruleNames
    Set ruleNames = CreateObject("Scripting.Dictionary")
    Dim ruleTexts
    Set ruleTexts = CreateObject("Scripting.Dictionary")
    
    Dim attrSet
    On Error Resume Next
    For Each attrSet In attrSets
        ' iLogic rules have attribute sets starting with specific patterns
        Dim setName
        setName = attrSet.Name
        
        ' Check for iLogic rule patterns
        If Left(setName, 10) = "iLogicRule" Or _
           Left(setName, 6) = "iLogic" Or _
           InStr(1, setName, "Rule", vbTextCompare) > 0 Then
            
            LogMessage "  Found AttributeSet: " & setName
            
            ' Try to get rule text from attributes
            Dim attr
            Dim ruleText
            ruleText = ""
            Dim ruleName
            ruleName = setName
            
            For Each attr In attrSet
                LogMessage "    Attribute: " & attr.Name & " = " & Left(CStr(attr.Value), 100)
                
                ' Look for the text/code attribute
                If LCase(attr.Name) = "text" Or LCase(attr.Name) = "code" Or LCase(attr.Name) = "rule" Then
                    ruleText = CStr(attr.Value)
                End If
                If LCase(attr.Name) = "name" Or LCase(attr.Name) = "rulename" Then
                    ruleName = CStr(attr.Value)
                End If
            Next
            
            If ruleText <> "" Or InStr(1, setName, "iLogic", vbTextCompare) > 0 Then
                ruleCount = ruleCount + 1
                ruleNames.Add ruleCount, ruleName
                ruleTexts.Add ruleCount, ruleText
                ruleInfo = ruleInfo & ruleCount & ". " & ruleName & vbCrLf
            End If
        End If
    Next
    Err.Clear
    
    ' If no rules found via AttributeSets, try using RunRule to detect rules
    If ruleCount = 0 Then
        LogMessage "No rules found via AttributeSets, trying alternate detection..."
        
        ' Try to call iLogicAuto.RunRule with a test name to see if it errors differently
        On Error Resume Next
        Dim testResult
        testResult = iLogicAuto.RunRule(sourceDoc, "Configuration")
        
        If Err.Number = 0 Then
            ' The rule exists and ran!
            LogMessage "Rule 'Configuration' exists and ran successfully"
            ruleCount = 1
            ruleNames.Add 1, "Configuration"
            ruleTexts.Add 1, "(Rule exists but source code not accessible via VBScript)"
            ruleInfo = "1. Configuration" & vbCrLf
        ElseIf InStr(1, Err.Description, "not found", vbTextCompare) > 0 Then
            LogMessage "RunRule test: Rule not found - " & Err.Description
        Else
            LogMessage "RunRule test returned: " & Err.Description
            ' The error might indicate the rule exists but couldn't run
            If Err.Number <> 0 And InStr(1, Err.Description, "Invalid", vbTextCompare) = 0 Then
                ruleCount = 1
                ruleNames.Add 1, "Configuration (detected)"
                ruleTexts.Add 1, "(Rule detected but source code not accessible)"
                ruleInfo = "1. Configuration (detected)" & vbCrLf
            End If
        End If
        Err.Clear
    End If
    
    On Error GoTo 0
    
    LogMessage "FOUND: " & ruleCount & " iLogic rule(s)"
    
    ' Step 4: Display summary
    If ruleCount = 0 Then
        MsgBox "NO iLOGIC RULES FOUND" & vbCrLf & vbCrLf & _
               "Document: " & sourceDoc.DisplayName & vbCrLf & vbCrLf & _
               "This document does not contain any embedded iLogic rules." & vbCrLf & vbCrLf & _
               "Note: Rules may exist but may not be accessible via external VBScript.", _
               vbInformation, "iLogic Scanner"
        LogMessage "No rules found in document"
        Call StopLogging
        Exit Sub
    End If
    
    Dim summaryMsg
    summaryMsg = "iLOGIC RULES FOUND: " & ruleCount & vbCrLf & vbCrLf & _
                 "Document: " & sourceDoc.DisplayName & vbCrLf & vbCrLf & _
                 "Rules:" & vbCrLf & ruleInfo & vbCrLf & _
                 "Would you like to view the rule details?"
    
    Dim viewCode
    viewCode = MsgBox(summaryMsg, vbYesNo + vbQuestion, "iLogic Scanner - Rules Found")
    
    If viewCode = vbYes Then
        ' Step 5: Show each rule's info
        Call DisplayRuleInfo(ruleNames, ruleTexts)
    End If
    
    ' Step 6: Ask about export
    Dim exportChoice
    exportChoice = MsgBox("EXPORT RULES TO FILES?" & vbCrLf & vbCrLf & _
                          "Would you like to export all " & ruleCount & " rule(s) to text files?" & vbCrLf & vbCrLf & _
                          "Files will be saved to:" & vbCrLf & _
                          sourceDir & "\iLogic_Export\" & vbCrLf & vbCrLf & _
                          "This creates a backup of your iLogic code.", _
                          vbYesNo + vbQuestion, "Export Rules?")
    
    If exportChoice = vbYes Then
        g_OutputFolder = sourceDir & "\iLogic_Export"
        Call ExportRulesFromDict(fso, ruleNames, ruleTexts, sourceFileName, sourceDoc.DisplayName)
    End If
    
    ' Also scan referenced parts in assembly
    If LCase(Right(sourceDoc.FullFileName, 4)) = ".iam" Then
        Dim scanParts
        scanParts = MsgBox("SCAN REFERENCED PARTS?" & vbCrLf & vbCrLf & _
                           "This is an assembly. Would you like to scan" & vbCrLf & _
                           "all referenced parts for iLogic rules too?" & vbCrLf & vbCrLf & _
                           "(This may take a few minutes for large assemblies)", _
                           vbYesNo + vbQuestion, "Scan Parts?")
        
        If scanParts = vbYes Then
            Call ScanReferencedParts(invApp, iLogicAuto, sourceDoc, fso)
        End If
    End If
    
    LogMessage "=== iLOGIC SCANNER COMPLETED ==="
    Call StopLogging
    
    MsgBox "iLogic scan completed!" & vbCrLf & vbCrLf & _
           "Log saved to:" & vbCrLf & g_LogPath, vbInformation, "Complete"
End Sub

' ==============================================================================
' HELPER FUNCTIONS
' ==============================================================================

Function DetectOpenDocument(invApp)
    On Error Resume Next
    
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument
    
    If Err.Number <> 0 Or activeDoc Is Nothing Then
        LogMessage "No active document found"
        Set DetectOpenDocument = Nothing
        Exit Function
    End If
    
    ' Check by file extension - accept .iam, .ipt, .ipn
    Dim ext
    ext = LCase(Right(activeDoc.FullFileName, 4))
    
    If ext <> ".iam" And ext <> ".ipt" And ext <> ".ipn" Then
        LogMessage "File extension not supported: " & activeDoc.FullFileName
        MsgBox "Current file type not supported!" & vbCrLf & _
               "Please open a part (.ipt), assembly (.iam), or presentation (.ipn).", vbExclamation
        Set DetectOpenDocument = Nothing
        Exit Function
    End If
    
    Dim docType
    Select Case ext
        Case ".iam"
            docType = "Assembly"
        Case ".ipt"
            docType = "Part"
        Case ".ipn"
            docType = "Presentation"
    End Select
    
    LogMessage "DETECTED: " & docType & " - " & activeDoc.DisplayName
    LogMessage "DETECTED: Full path - " & activeDoc.FullFileName
    
    ' Confirm with user
    Dim confirmMsg
    confirmMsg = "DOCUMENT DETECTED" & vbCrLf & vbCrLf & _
                 "Type: " & docType & vbCrLf & _
                 "Name: " & activeDoc.DisplayName & vbCrLf & vbCrLf & _
                 "Scan this document for iLogic rules?"
    
    Dim confirmResult
    confirmResult = MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Document")
    
    If confirmResult = vbNo Then
        LogMessage "User cancelled document confirmation"
        Set DetectOpenDocument = Nothing
        Exit Function
    End If
    
    Set DetectOpenDocument = activeDoc
    Err.Clear
End Function

Function GetILogicAddin(invApp)
    On Error Resume Next
    
    Dim addins
    Set addins = invApp.ApplicationAddIns
    
    Dim iLogicAddin
    Set iLogicAddin = addins.ItemById(ILOGIC_ADDIN_GUID)
    
    If Err.Number <> 0 Or iLogicAddin Is Nothing Then
        LogMessage "ERROR: iLogic Add-In not found by GUID"
        
        ' Try to find by name as fallback
        Dim addin
        For Each addin In addins
            If InStr(1, addin.DisplayName, "iLogic", vbTextCompare) > 0 Then
                LogMessage "Found iLogic by name: " & addin.DisplayName
                Set iLogicAddin = addin
                Exit For
            End If
        Next
    End If
    
    If iLogicAddin Is Nothing Then
        LogMessage "iLogic Add-In not found"
        Set GetILogicAddin = Nothing
    Else
        LogMessage "SUCCESS: Found iLogic Add-In: " & iLogicAddin.DisplayName
        Set GetILogicAddin = iLogicAddin
    End If
    
    Err.Clear
End Function

Sub DisplayRuleInfo(ruleNames, ruleTexts)
    On Error Resume Next
    
    Dim i
    Dim totalRules
    totalRules = ruleNames.Count
    
    For i = 1 To totalRules
        Dim ruleName, ruleText
        ruleName = ruleNames(i)
        ruleText = ruleTexts(i)
        
        ' Truncate if too long for msgbox
        Dim displayText
        If Len(ruleText) > 1500 Then
            displayText = Left(ruleText, 1500) & vbCrLf & vbCrLf & _
                         "... [TRUNCATED - " & Len(ruleText) & " characters total]"
        Else
            displayText = ruleText
        End If
        
        If displayText = "" Then
            displayText = "(No source code available - rule detected via AttributeSets)"
        End If
        
        Dim continueView
        continueView = MsgBox("RULE " & i & " of " & totalRules & ": " & ruleName & vbCrLf & _
                              String(50, "-") & vbCrLf & vbCrLf & _
                              displayText & vbCrLf & vbCrLf & _
                              String(50, "-") & vbCrLf & _
                              "View next rule?", _
                              vbYesNo + vbInformation, "Rule Details")
        
        LogMessage "  Rule '" & ruleName & "' - " & Len(ruleText) & " characters"
        
        If continueView = vbNo Then Exit For
    Next
    
    Err.Clear
End Sub

Sub ExportRulesFromDict(fso, ruleNames, ruleTexts, baseName, docDisplayName)
    On Error Resume Next
    
    ' Create export folder
    If Not fso.FolderExists(g_OutputFolder) Then
        fso.CreateFolder g_OutputFolder
        If Err.Number <> 0 Then
            LogMessage "ERROR: Could not create export folder: " & Err.Description
            MsgBox "ERROR: Could not create export folder!", vbCritical
            Exit Sub
        End If
    End If
    
    LogMessage "EXPORT: Saving rules to " & g_OutputFolder
    
    Dim i
    Dim exportCount
    exportCount = 0
    
    For i = 1 To ruleNames.Count
        Dim ruleName, ruleText
        ruleName = ruleNames(i)
        ruleText = ruleTexts(i)
        
        ' Clean filename (remove invalid characters)
        Dim safeRuleName
        safeRuleName = CleanFileName(ruleName)
        
        Dim exportPath
        exportPath = g_OutputFolder & "\" & baseName & "_" & safeRuleName & ".vb"
        
        ' Write to file
        Dim outFile
        Set outFile = fso.CreateTextFile(exportPath, True, False)
        
        ' Add header comment
        outFile.WriteLine "' =============================================="
        outFile.WriteLine "' iLogic Rule Export"
        outFile.WriteLine "' =============================================="
        outFile.WriteLine "' Source Document: " & docDisplayName
        outFile.WriteLine "' Rule Name: " & ruleName
        outFile.WriteLine "' Exported: " & Now()
        outFile.WriteLine "' =============================================="
        outFile.WriteLine ""
        outFile.WriteLine ruleText
        outFile.Close
        
        exportCount = exportCount + 1
        LogMessage "  EXPORTED: " & safeRuleName & ".vb"
    Next
    
    LogMessage "EXPORT COMPLETE: " & exportCount & " rules exported"
    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           exportCount & " rules exported to:" & vbCrLf & _
           g_OutputFolder, vbInformation, "Export Complete"
    
    Err.Clear
End Sub

Function CleanFileName(name)
    Dim result
    result = name
    
    ' Replace invalid filename characters
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    result = Replace(result, " ", "_")
    
    CleanFileName = result
End Function

Sub ScanReferencedParts(invApp, iLogicAuto, asmDoc, fso)
    LogMessage "SCANNING: Referenced parts for iLogic rules"
    
    Dim compDef
    Set compDef = asmDoc.ComponentDefinition
    
    Dim allParts
    Set allParts = CreateObject("Scripting.Dictionary")
    
    ' Collect all unique part paths
    Call CollectPartsRecursive(compDef.Occurrences, allParts)
    
    LogMessage "FOUND: " & allParts.Count & " unique referenced documents"
    
    Dim partsWithRules
    partsWithRules = 0
    
    Dim totalRules
    totalRules = 0
    
    Dim partReport
    partReport = ""
    
    Dim partPath
    For Each partPath In allParts.Keys
        On Error Resume Next
        
        ' Get the document (may already be open)
        Dim partDoc
        Set partDoc = invApp.Documents.Open(partPath, False) ' Open invisible if not already open
        
        If Err.Number = 0 And Not partDoc Is Nothing Then
            ' Get rules for this part
            Dim partRules
            Set partRules = iLogicAuto.Rules(partDoc)
            
            Dim partRuleCount
            partRuleCount = 0
            
            Dim r
            For Each r In partRules
                partRuleCount = partRuleCount + 1
            Next
            
            If partRuleCount > 0 Then
                partsWithRules = partsWithRules + 1
                totalRules = totalRules + partRuleCount
                
                partReport = partReport & fso.GetFileName(partPath) & ": " & partRuleCount & " rule(s)" & vbCrLf
                LogMessage "  " & fso.GetFileName(partPath) & ": " & partRuleCount & " rules"
            End If
        End If
        
        Err.Clear
    Next
    On Error GoTo 0
    
    ' Show summary
    Dim summaryMsg
    If partsWithRules = 0 Then
        summaryMsg = "REFERENCED PARTS SCAN COMPLETE" & vbCrLf & vbCrLf & _
                     "Scanned: " & allParts.Count & " documents" & vbCrLf & vbCrLf & _
                     "No iLogic rules found in referenced parts."
    Else
        summaryMsg = "REFERENCED PARTS SCAN COMPLETE" & vbCrLf & vbCrLf & _
                     "Scanned: " & allParts.Count & " documents" & vbCrLf & _
                     "Parts with rules: " & partsWithRules & vbCrLf & _
                     "Total rules found: " & totalRules & vbCrLf & vbCrLf & _
                     "Details:" & vbCrLf & partReport
    End If
    
    MsgBox summaryMsg, vbInformation, "Referenced Parts Scan"
    LogMessage "SCAN COMPLETE: " & partsWithRules & " parts with " & totalRules & " total rules"
End Sub

Sub CollectPartsRecursive(occurrences, partsDict)
    On Error Resume Next
    
    Dim occ
    For Each occ In occurrences
        ' Get the referenced document path
        Dim refDoc
        Set refDoc = occ.Definition.Document
        
        If Not refDoc Is Nothing Then
            Dim docPath
            docPath = refDoc.FullFileName
            
            ' Add to dictionary if not already there
            If Not partsDict.Exists(docPath) Then
                partsDict.Add docPath, True
            End If
            
            ' If this is a sub-assembly, recurse into it
            Dim ext
            ext = LCase(Right(docPath, 4))
            If ext = ".iam" Then
                Call CollectPartsRecursive(occ.SubOccurrences, partsDict)
            End If
        End If
    Next
    
    Err.Clear
End Sub

' ==============================================================================
' LOGGING FUNCTIONS
' ==============================================================================

Sub StartLogging()
    Set g_fso = CreateObject("Scripting.FileSystemObject")
    
    Dim scriptPath
    scriptPath = g_fso.GetParentFolderName(WScript.ScriptFullName)
    
    ' Go up one level to main folder, then into Logs
    Dim mainFolder
    mainFolder = g_fso.GetParentFolderName(scriptPath)
    
    Dim logFolder
    logFolder = mainFolder & "\Logs"
    
    If Not g_fso.FolderExists(logFolder) Then
        g_fso.CreateFolder logFolder
    End If
    
    Dim timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & "_" & _
                Right("0" & Hour(Now), 2) & "-" & Right("0" & Minute(Now), 2) & "-" & Right("0" & Second(Now), 2)
    
    g_LogPath = logFolder & "\iLogic_Scanner_" & timestamp & ".txt"
    
    Set g_LogFile = g_fso.CreateTextFile(g_LogPath, True, False)
    
    LogMessage "=== iLogic Scanner Log ==="
    LogMessage "Started: " & Now()
    LogMessage "=========================================="
End Sub

Sub StopLogging()
    LogMessage "=========================================="
    LogMessage "Completed: " & Now()
    g_LogFile.Close
    Set g_LogFile = Nothing
End Sub

Sub LogMessage(msg)
    On Error Resume Next
    g_LogFile.WriteLine Now() & " - " & msg
    Err.Clear
End Sub
