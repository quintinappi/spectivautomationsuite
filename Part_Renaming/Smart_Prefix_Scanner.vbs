Option Explicit

' ==============================================================================
' SMART PREFIX SCANNER - PART RENAMING STEP 1a: ENSURE PREFIX READINESS
' ==============================================================================
' PART RENAMING WORKFLOW - STEP 1a: Ensure prefix is ready
'
' This tool prevents duplicate numbering by scanning existing files
' Use BEFORE running the main renamer (STEP 1b) on new assemblies!
'
' This tool:
' 1. Scans the currently open assembly (or entire Structure.iam)
' 2. Detects the prefix used (e.g., NCRH01-000-)
' 3. Finds the highest number used for each group (PL, B, CH, A, etc.)
' 4. Saves these numbers to Registry for continuation
' 5. Main renamer (STEP 1b) will then continue from these numbers
'
' WHEN TO USE:
' - Before adding new assemblies (like Access Walkway) to existing project
' - When Registry has been cleared but files still exist with numbers
' - To ensure numbering continues without duplicates
' - To "catch up" Registry with actual file state
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath
Dim g_DetectedPrefix
Dim g_GroupCounters ' Dictionary: groupCode -> highest number found

Call SMART_PREFIX_SCANNER()

Sub SMART_PREFIX_SCANNER()
    Call StartLogging
    LogMessage "=== SMART PREFIX SCANNER ==="
    LogMessage "Prevent duplicate numbering by scanning existing model"

    Dim result
    result = MsgBox("SMART PREFIX SCANNER - PART RENAMING STEP 1a" & vbCrLf & vbCrLf & _
                    "PART RENAMING WORKFLOW - STEP 1a: Ensure prefix is ready" & vbCrLf & vbCrLf & _
                    "Use this BEFORE running the main renamer (STEP 1b)!" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Scan your currently open assembly" & vbCrLf & _
                    "2. Detect the prefix used (e.g., NCRH01-000-)" & vbCrLf & _
                    "3. Find highest numbers for each group:" & vbCrLf & _
                    "   - PL (plates), B (beams), CH (channels)" & vbCrLf & _
                    "   - A (angles), P (pipes), SQ (square tubes), etc." & vbCrLf & _
                    "4. Save to Registry so main renamer continues correctly" & vbCrLf & vbCrLf & _
                    "⚠️  Run this FIRST if adding assemblies to existing projects!" & vbCrLf & vbCrLf & _
                    "Make sure your main assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Step 1a: Prefix Scanner")

    If result = vbNo Then
        LogMessage "User cancelled scanner"
        Exit Sub
    End If

    ' Initialize collections
    Set g_GroupCounters = CreateObject("Scripting.Dictionary")

    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open your main assembly first.", vbCritical
        Exit Sub
    End If

    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear

    ' Detect open assembly
    Dim activeDoc
    Set activeDoc = DetectOpenAssembly(invApp)
    If activeDoc Is Nothing Then
        MsgBox "ERROR: No assembly is currently open!" & vbCrLf & _
               "Please open your main assembly (e.g., Structure.iam) first.", vbCritical
        Exit Sub
    End If

    LogMessage "ASSEMBLY: Detected - " & activeDoc.DisplayName

    ' Step 1: Recursively scan assembly for heritage files
    LogMessage "STEP 1: Recursively scanning assembly for heritage part numbers"
    Call ScanAssemblyForHeritageFiles(activeDoc)

    If g_GroupCounters.Count = 0 Then
        LogMessage "WARNING: No heritage files found in assembly"
        MsgBox "WARNING: No heritage files detected!" & vbCrLf & vbCrLf & _
               "This might mean:" & vbCrLf & _
               "  - Parts haven't been renamed yet (run STEP 1 first)" & vbCrLf & _
               "  - Wrong assembly opened" & vbCrLf & _
               "  - Looking for wrong prefix pattern" & vbCrLf & vbCrLf & _
               "No counters will be saved to Registry.", vbExclamation
        Exit Sub
    End If

    ' Step 2: Show results and ask for confirmation
    LogMessage "STEP 2: Showing detected prefix and highest numbers"
    Dim confirmation
    confirmation = ShowScanResults()

    If Not confirmation Then
        LogMessage "User cancelled - no Registry updates"
        Exit Sub
    End If

    ' Step 3: Save to Registry
    LogMessage "STEP 3: Saving counters to Registry"
    Call SaveCountersToRegistry()

    LogMessage "=== SMART PREFIX SCANNER COMPLETED ==="
    Call StopLogging

    MsgBox "SMART PREFIX SCANNER COMPLETED!" & vbCrLf & vbCrLf & _
           "Detected prefix: " & g_DetectedPrefix & vbCrLf & _
           "Groups scanned: " & g_GroupCounters.Count & vbCrLf & vbCrLf & _
           "✓ Registry updated successfully" & vbCrLf & _
           "✓ STEP 1 will now continue numbering from these values" & vbCrLf & _
           "✓ No duplicate part numbers will be created" & vbCrLf & vbCrLf & _
           "You can now safely run STEP 1 on new assemblies!" & vbCrLf & vbCrLf & _
           "Log: " & g_LogPath, vbInformation, "Scanner Complete"
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

    ' Check if it's an assembly by extension
    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        LogMessage "Active document is not an assembly: " & activeDoc.FullFileName
        Set DetectOpenAssembly = Nothing
        Exit Function
    End If

    Set DetectOpenAssembly = activeDoc
    Err.Clear
End Function

Sub ScanAssemblyForHeritageFiles(asmDoc)
    LogMessage "SCAN: Recursively scanning - " & asmDoc.DisplayName

    ' Create unique parts tracker
    Dim uniqueParts
    Set uniqueParts = CreateObject("Scripting.Dictionary")

    ' Start recursive scan
    Call ProcessAssemblyRecursively(asmDoc, uniqueParts)

    LogMessage "SCAN: Completed - Processed " & uniqueParts.Count & " unique parts"
End Sub

Sub ProcessAssemblyRecursively(asmDoc, uniqueParts)
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    LogMessage "SCAN: Processing " & occurrences.Count & " occurrences in " & asmDoc.DisplayName

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Skip suppressed
        If Not occ.Suppressed Then
            Dim doc
            Set doc = occ.Definition.Document

            Dim fileName
            fileName = GetFileNameFromPath(doc.FullFileName)
            Dim fullPath
            fullPath = doc.FullFileName

            ' Process part files
            If LCase(Right(fileName, 4)) = ".ipt" Then
                If Not uniqueParts.Exists(fullPath) Then
                    uniqueParts.Add fullPath, True

                    ' Check if this is a heritage file
                    If IsHeritageFile(fileName) Then
                        LogMessage "HERITAGE: Found - " & fileName
                        Call ExtractAndStoreCounter(fileName)
                    Else
                        LogMessage "ORIGINAL: Skipping - " & fileName
                    End If
                End If

            ' Recurse into sub-assemblies
            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                If InStr(LCase(fileName), "bolted connection") = 0 Then
                    LogMessage "SCAN: Recursing into sub-assembly - " & fileName
                    Call ProcessAssemblyRecursively(doc, uniqueParts)
                End If
            End If
        End If
    Next
End Sub

Function IsHeritageFile(fileName)
    ' Detect heritage filename pattern: PREFIX-###-CODEXXX.ipt
    ' Examples: NCRH01-000-PL123.ipt, PLANT1-000-B45.ipt

    Dim baseName
    baseName = fileName
    If LCase(Right(baseName, 4)) = ".ipt" Then
        baseName = Left(baseName, Len(baseName) - 4)
    End If

    ' Count dashes
    Dim dashCount
    dashCount = 0
    Dim i
    For i = 1 To Len(baseName)
        If Mid(baseName, i, 1) = "-" Then
            dashCount = dashCount + 1
        End If
    Next

    ' Heritage files need at least 2 dashes
    If dashCount >= 2 Then
        ' Must end with a number
        Dim lastChar
        lastChar = Right(baseName, 1)
        If IsNumeric(lastChar) Then
            IsHeritageFile = True
            Exit Function
        End If
    End If

    IsHeritageFile = False
End Function

Sub ExtractAndStoreCounter(fileName)
    ' Extract prefix, group code, and number from heritage filename
    ' Format: PREFIX-###-CODEXXX.ipt
    ' Example: NCRH01-000-PL173.ipt -> Prefix="NCRH01-000-", Group="PL", Number=173

    LogMessage "EXTRACT: Analyzing - " & fileName

    Dim baseName
    baseName = fileName
    If LCase(Right(baseName, 4)) = ".ipt" Then
        baseName = Left(baseName, Len(baseName) - 4)
    End If

    ' Find the last dash (separates prefix from code+number)
    Dim lastDashPos
    lastDashPos = InStrRev(baseName, "-")

    If lastDashPos = 0 Then
        LogMessage "EXTRACT: ERROR - No dash found in: " & fileName
        Exit Sub
    End If

    ' Extract prefix (everything up to and including last dash)
    Dim prefixPart
    prefixPart = Left(baseName, lastDashPos)

    ' Extract code+number part (everything after last dash)
    Dim codeNumberPart
    codeNumberPart = Mid(baseName, lastDashPos + 1)

    LogMessage "EXTRACT: Prefix='" & prefixPart & "', CodeNumber='" & codeNumberPart & "'"

    ' Separate letters from numbers in code+number part
    ' Scan backwards to find where numbers start
    Dim i, numberStart
    numberStart = 0

    For i = Len(codeNumberPart) To 1 Step -1
        Dim char
        char = Mid(codeNumberPart, i, 1)
        If IsNumeric(char) Then
            numberStart = i
        Else
            Exit For
        End If
    Next

    If numberStart = 0 Then
        LogMessage "EXTRACT: ERROR - No number found in: " & codeNumberPart
        Exit Sub
    End If

    ' Extract group code and number
    Dim groupCode
    groupCode = Left(codeNumberPart, numberStart - 1)
    Dim numberPart
    numberPart = Mid(codeNumberPart, numberStart)

    If Not IsNumeric(numberPart) Then
        LogMessage "EXTRACT: ERROR - Invalid number: " & numberPart
        Exit Sub
    End If

    Dim currentNumber
    currentNumber = CInt(numberPart)

    LogMessage "EXTRACT: Group='" & groupCode & "', Number=" & currentNumber

    ' Set detected prefix (first time)
    If g_DetectedPrefix = "" Then
        g_DetectedPrefix = prefixPart
        LogMessage "PREFIX: Detected prefix: " & g_DetectedPrefix
    ElseIf g_DetectedPrefix <> prefixPart Then
        LogMessage "PREFIX: WARNING - Multiple prefixes detected! Previous: " & g_DetectedPrefix & ", Current: " & prefixPart
    End If

    ' Store/update highest number for this group
    Dim prefixGroupKey
    prefixGroupKey = prefixPart & groupCode

    If g_GroupCounters.Exists(prefixGroupKey) Then
        Dim existingNumber
        existingNumber = g_GroupCounters.Item(prefixGroupKey)
        If currentNumber > existingNumber Then
            g_GroupCounters.Item(prefixGroupKey) = currentNumber
            LogMessage "COUNTER: Updated " & prefixGroupKey & " from " & existingNumber & " to " & currentNumber
        Else
            LogMessage "COUNTER: Kept " & prefixGroupKey & " at " & existingNumber & " (current " & currentNumber & " is lower)"
        End If
    Else
        g_GroupCounters.Add prefixGroupKey, currentNumber
        LogMessage "COUNTER: Added " & prefixGroupKey & " = " & currentNumber
    End If
End Sub

Function ShowScanResults()
    ' Build results message
    Dim msg
    msg = "SCAN RESULTS" & vbCrLf & vbCrLf

    If g_DetectedPrefix = "" Then
        msg = msg & "No prefix detected" & vbCrLf
        ShowScanResults = False
        MsgBox msg, vbExclamation, "No Results"
        Exit Function
    End If

    msg = msg & "Detected Prefix: " & g_DetectedPrefix & vbCrLf & vbCrLf
    msg = msg & "Highest numbers found for each group:" & vbCrLf & vbCrLf

    ' Show each group's highest number
    Dim keys
    keys = g_GroupCounters.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim key
        key = keys(i)
        Dim value
        value = g_GroupCounters.Item(key)

        ' Extract just the group code for display
        Dim groupCode
        groupCode = Replace(key, g_DetectedPrefix, "")

        ' Add friendly description
        Dim description
        description = GetGroupDescription(groupCode)

        msg = msg & "  " & key & " = " & value & vbCrLf
        msg = msg & "    (" & description & ")" & vbCrLf & vbCrLf
    Next

    msg = msg & vbCrLf & "STEP 1 will continue from these numbers." & vbCrLf & vbCrLf
    msg = msg & "For example:" & vbCrLf

    ' Show example of next numbers
    For i = 0 To UBound(keys)
        If i >= 3 Then Exit For ' Show max 3 examples
        key = keys(i)
        value = g_GroupCounters.Item(key)
        groupCode = Replace(key, g_DetectedPrefix, "")
        Dim nextNumber
        nextNumber = value + 1
        msg = msg & "  Next " & groupCode & " part will be: " & key & nextNumber & ".ipt" & vbCrLf
    Next

    msg = msg & vbCrLf & "Save these counters to Registry?"

    Dim result
    result = MsgBox(msg, vbYesNo + vbQuestion, "Scan Results")

    ShowScanResults = (result = vbYes)
End Function

Function GetGroupDescription(groupCode)
    ' Return friendly description for group code
    Select Case UCase(groupCode)
        Case "B"
            GetGroupDescription = "I and H sections (beams/columns)"
        Case "PL"
            GetGroupDescription = "Platework"
        Case "LPL"
            GetGroupDescription = "Liners"
        Case "A"
            GetGroupDescription = "Angles"
        Case "CH"
            GetGroupDescription = "Channels"
        Case "P"
            GetGroupDescription = "Circular hollow sections"
        Case "SQ"
            GetGroupDescription = "Square/rectangular hollow"
        Case "FL"
            GetGroupDescription = "Flatbar"
        Case "IPE"
            GetGroupDescription = "European I-beams"
        Case Else
            GetGroupDescription = "Other components"
    End Select
End Function

Sub SaveCountersToRegistry()
    LogMessage "REGISTRY: Saving counters to Windows Registry"

    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    Dim keys
    keys = g_GroupCounters.Keys

    Dim i, successCount, errorCount
    successCount = 0
    errorCount = 0

    For i = 0 To UBound(keys)
        Dim key
        key = keys(i)
        Dim value
        value = g_GroupCounters.Item(key)

        On Error Resume Next
        shell.RegWrite regPath & key, value, "REG_DWORD"

        If Err.Number = 0 Then
            LogMessage "REGISTRY: Saved " & key & " = " & value
            successCount = successCount + 1
        Else
            LogMessage "REGISTRY: ERROR - Could not save " & key & ": " & Err.Description
            errorCount = errorCount + 1
        End If

        Err.Clear
    Next

    LogMessage "REGISTRY: Save complete - Success: " & successCount & ", Errors: " & errorCount

    If errorCount > 0 Then
        MsgBox "WARNING: Some counters could not be saved!" & vbCrLf & vbCrLf & _
               "Success: " & successCount & vbCrLf & _
               "Errors: " & errorCount & vbCrLf & vbCrLf & _
               "Check log for details.", vbExclamation
    End If
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
    g_LogPath = logsDir & "\Smart_Prefix_Scanner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
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