' ===================================================================
' INVENTOR RENAMER - REGISTRY DATABASE MANAGER - PART RENAMING STEP 1: Familiarize with current registry
' ===================================================================
' PART RENAMING WORKFLOW - STEP 1: Familiarize with current registry for prefix
' View counters, scan project to update registry, or clear the database

Option Explicit

Dim shell
Set shell = CreateObject("WScript.Shell")

' Show GUI with 3 options
Dim actionChoice
actionChoice = InputBox("INVENTOR RENAMER - Registry Database Manager" & vbCrLf & vbCrLf & _
                        "Choose an action:" & vbCrLf & vbCrLf & _
                        "1 = SCAN Registry (view current counters)" & vbCrLf & _
                        "2 = SCAN PROJECT & Update Registry (scan open assembly)" & vbCrLf & _
                        "3 = CLEAR Registry (delete counters for prefix)" & vbCrLf & vbCrLf & _
                        "Enter 1, 2, or 3 (or leave blank to cancel):" & vbCrLf & vbCrLf & _
                        "⚠️ Option 2 requires Inventor to be running with an assembly open!" & vbCrLf & _
                        "⚠️ Option 3 will delete numbering history!", _
                        "Registry Database Manager", "1")

Select Case Trim(actionChoice)
    Case "1"
        Call ScanRegistry()
    Case "2"
        Call ScanProjectAndUpdateRegistry()
    Case "3"
        Call ClearRegistry()
    Case Else
        WScript.Quit
End Select

' ===================================================================
Sub ScanRegistry()
    ' Ask user for prefix to scan
    Dim userPrefix
    userPrefix = InputBox("Enter the prefix to scan for:" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "  NCRH01-000-    (original default)" & vbCrLf & _
                         "  PLANT1-000-    (Plant 1)" & vbCrLf & _
                         "  TEST-000-      (test prefix)" & vbCrLf & _
                         "  AREA2-000-     (Area 2)" & vbCrLf & vbCrLf & _
                         "Or leave blank to show ALL entries", _
                         "Scan Registry - Choose Prefix", "NCRH01-000-")

    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    Dim report
    If userPrefix = "" Then
        report = "ALL REGISTRY ENTRIES:" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    Else
        ' Ensure prefix ends with dash
        If Right(userPrefix, 1) <> "-" Then
            userPrefix = userPrefix & "-"
        End If
        report = "REGISTRY ENTRIES FOR PREFIX: " & userPrefix & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    End If

    ' Get counter keys based on user choice
    Dim counterKeys
    Set counterKeys = CreateObject("Scripting.Dictionary")

    If userPrefix = "" Then
        ' Show all entries
        Call EnumerateRegistryKeys(regPath, counterKeys)
    Else
        ' Show only entries for specific prefix
        Call ScanSpecificPrefix(regPath, userPrefix, counterKeys)
    End If

    If counterKeys.Count = 0 Then
        report = report & "No counters found - database is empty." & vbCrLf
        report = report & "Next run will start from: CH1, PL1, B1, etc."
    Else
        report = report & "Found " & counterKeys.Count & " counter(s):" & vbCrLf & vbCrLf

        ' Show all existing counters
        Dim keyNames
        keyNames = counterKeys.Keys
        Dim i
        For i = 0 To UBound(keyNames)
            Dim keyName
            keyName = keyNames(i)
            Dim currentValue
            currentValue = counterKeys.Item(keyName)
            report = report & keyName & " = " & currentValue & vbCrLf
        Next

        report = report & vbCrLf & "NEXT NUMBERS WILL BE:" & vbCrLf & String(30, "-") & vbCrLf

        ' Show what the next numbers will be
        For i = 0 To UBound(keyNames)
            Dim keyName2
            keyName2 = keyNames(i)
            Dim nextValue
            nextValue = counterKeys.Item(keyName2)

            ' Extract group name (everything after the last dash)
            Dim groupName
            groupName = Right(keyName2, Len(keyName2) - InStrRev(keyName2, "-"))
            If InStrRev(keyName2, "-") = 0 Then
                groupName = keyName2  ' No dash found, use whole key
            End If
            report = report & groupName & " will continue from " & (nextValue + 1) & vbCrLf
        Next
    End If

    report = report & vbCrLf & String(50, "=")

    MsgBox report, vbInformation, "Registry Scan Results"
End Sub

' ===================================================================
Sub ClearRegistry()
    ' SAFE CLEAR - Asks for prefix and shows what will be deleted

    ' Ask user for prefix to clear
    Dim userPrefix
    userPrefix = InputBox("⚠️ CLEAR REGISTRY DATABASE" & vbCrLf & vbCrLf & _
                         "Enter the PREFIX to clear:" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "  NCRH01-000-    (original default)" & vbCrLf & _
                         "  TEST-000-      (test prefix)" & vbCrLf & vbCrLf & _
                         "⚠️ This will DELETE all counters for this prefix!" & vbCrLf & _
                         "Leave blank to cancel.", _
                         "Clear Registry - Choose Prefix", "NCRH01-000-")

    ' Exit if user cancelled
    If userPrefix = "" Then
        MsgBox "Clear operation cancelled.", vbInformation
        WScript.Quit
    End If

    ' Ensure prefix ends with dash
    If Right(userPrefix, 1) <> "-" Then
        userPrefix = userPrefix & "-"
    End If

    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    ' First, scan what exists for this prefix
    Dim counterKeys
    Set counterKeys = CreateObject("Scripting.Dictionary")
    Call ScanSpecificPrefix(regPath, userPrefix, counterKeys)

    If counterKeys.Count = 0 Then
        MsgBox "No registry entries found for prefix: " & userPrefix & vbCrLf & vbCrLf & _
               "Nothing to clear.", vbInformation, "Registry Clear"
        WScript.Quit
    End If

    ' Show what will be deleted and ask for confirmation
    Dim confirmMsg
    confirmMsg = "⚠️ CONFIRM DELETION" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "The following entries will be DELETED:" & vbCrLf & vbCrLf

    Dim keyNames
    keyNames = counterKeys.Keys
    Dim i
    For i = 0 To UBound(keyNames)
        confirmMsg = confirmMsg & keyNames(i) & " = " & counterKeys.Item(keyNames(i)) & vbCrLf
    Next

    confirmMsg = confirmMsg & vbCrLf & String(50, "=") & vbCrLf
    confirmMsg = confirmMsg & "Found " & counterKeys.Count & " entries to delete." & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "⚠️ This cannot be undone!" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "Are you SURE you want to clear these counters?"

    Dim confirm
    confirm = MsgBox(confirmMsg, vbYesNo + vbExclamation, "Confirm Registry Clear")

    If confirm = vbNo Then
        MsgBox "Clear operation cancelled. No changes made.", vbInformation
        WScript.Quit
    End If

    ' Delete the entries
    Dim deletedCount
    deletedCount = 0

    For i = 0 To UBound(keyNames)
        On Error Resume Next
        shell.RegDelete regPath & keyNames(i)

        If Err.Number = 0 Then
            deletedCount = deletedCount + 1
        End If

        Err.Clear
    Next

    ' Show results
    MsgBox "✅ Registry Clear Complete!" & vbCrLf & vbCrLf & _
           "Deleted: " & deletedCount & " / " & counterKeys.Count & " entries" & vbCrLf & vbCrLf & _
           "Next STEP 1 run will start from:" & vbCrLf & _
           "CH1, PL1, B1, A1, FL1, etc.", _
           vbInformation, "Registry Cleared"
End Sub

' ===================================================================
Sub EnumerateRegistryKeys(regPath, counterKeys)
    ' Dynamically discover all existing counter keys in the registry
    ' This replaces the hardcoded approach and works with any prefix

    On Error Resume Next

    ' Try to enumerate keys using WMI
    Dim objWMIService
    Set objWMIService = GetObject("winmgmts:\\.\root\default:StdRegProv")

    Dim arrSubKeys
    Dim strKeyPath
    strKeyPath = "Software\InventorRenamer"

    ' Enumerate all value names under the registry key
    Dim arrValueNames, arrValueTypes
    objWMIService.EnumValues &H80000001, strKeyPath, arrValueNames, arrValueTypes

    If IsArray(arrValueNames) Then
        Dim i
        For i = 0 To UBound(arrValueNames)
            Dim valueName
            valueName = arrValueNames(i)

            ' Read the value
            Dim valueData
            objWMIService.GetDWORDValue &H80000001, strKeyPath, valueName, valueData

            If Err.Number = 0 Then
                counterKeys.Add valueName, valueData
            End If
        Next
    End If

    ' Fallback method using shell.RegRead if WMI fails
    If counterKeys.Count = 0 Then
        ' Try common prefixes and group combinations
        Dim commonPrefixes
        commonPrefixes = Array("NCRH01-000-", "TEST-000-", "PLANT1-000-", "AREA2-000-", "SEC-A-000-", "BLOCK3-000-")

        Dim commonGroups
        commonGroups = Array("CH", "PL", "B", "A", "P", "SQ", "FL", "LPL", "IPE", "R", "OTHER", "PART")

        Dim j, k
        For j = 0 To UBound(commonPrefixes)
            For k = 0 To UBound(commonGroups)
                Dim testKey
                testKey = commonPrefixes(j) & commonGroups(k)

                On Error Resume Next
                Dim testValue
                testValue = shell.RegRead(regPath & testKey)

                If Err.Number = 0 Then
                    counterKeys.Add testKey, testValue
                End If

                Err.Clear
            Next
        Next
    End If

    On Error GoTo 0
End Sub

' ===================================================================
Sub ScanSpecificPrefix(regPath, userPrefix, counterKeys)
    ' Scan for counter keys with specific prefix only

    Dim commonGroups
    commonGroups = Array("CH", "PL", "B", "A", "P", "SQ", "FL", "LPL", "IPE", "R", "OTHER", "PART")

    Dim i
    For i = 0 To UBound(commonGroups)
        Dim testKey
        testKey = userPrefix & commonGroups(i)

        On Error Resume Next
        Dim testValue
        testValue = shell.RegRead(regPath & testKey)

        If Err.Number = 0 Then
            counterKeys.Add testKey, testValue
        End If

        Err.Clear
    Next

    On Error GoTo 0
End Sub

' ===================================================================
Sub ScanProjectAndUpdateRegistry()
    ' Scan currently open assembly in Inventor
    ' Detect prefix from heritage file names
    ' Update Registry with highest numbers found

    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        MsgBox "ERROR: Inventor is not running!" & vbCrLf & vbCrLf & _
               "Please start Inventor and open your assembly first.", vbCritical
        Exit Sub
    End If
    Err.Clear

    ' Get active document
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If activeDoc Is Nothing Then
        MsgBox "ERROR: No document is open in Inventor!" & vbCrLf & vbCrLf & _
               "Please open an assembly first.", vbCritical
        Exit Sub
    End If

    ' Check if it's an assembly
    If LCase(Right(activeDoc.FullFileName, 4)) <> ".iam" Then
        MsgBox "ERROR: Active document is not an assembly!" & vbCrLf & vbCrLf & _
               "Please open an assembly (.iam) file.", vbCritical
        Exit Sub
    End If

    ' Scan assembly for heritage files and detect prefix
    Dim groupCounters
    Set groupCounters = CreateObject("Scripting.Dictionary")
    Dim detectedPrefix
    detectedPrefix = ""

    Call ScanAssemblyForPrefix(activeDoc, groupCounters, detectedPrefix)

    If detectedPrefix = "" Then
        MsgBox "WARNING: No heritage files found!" & vbCrLf & vbCrLf & _
               "This could mean:" & vbCrLf & _
               "  - Parts haven't been renamed yet" & vbCrLf & _
               "  - Wrong assembly is open" & vbCrLf & vbCrLf & _
               "Registry will NOT be updated.", vbExclamation
        Exit Sub
    End If

    ' Show results and confirm
    Dim report
    report = "DETECTED PREFIX: " & detectedPrefix & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    report = report & "Found highest numbers:" & vbCrLf & vbCrLf

    Dim keys
    keys = groupCounters.Keys
    Dim i
    For i = 0 To UBound(keys)
        Dim groupName
        groupName = keys(i)
        Dim highestNum
        highestNum = groupCounters.Item(groupName)
        report = report & "  " & groupName & " = " & highestNum & vbCrLf
    Next

    report = report & vbCrLf & String(50, "=") & vbCrLf
    report = report & "This will UPDATE the Registry with these values." & vbCrLf & vbCrLf
    report = report & "Continue?"

    Dim confirm
    confirm = MsgBox(report, vbYesNo + vbQuestion, "Confirm Registry Update")

    If confirm <> vbYes Then
        MsgBox "Registry update cancelled.", vbInformation
        Exit Sub
    End If

    ' Save to Registry
    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    Dim savedCount
    savedCount = 0

    For i = 0 To UBound(keys)
        Dim prefixGroupKey
        prefixGroupKey = detectedPrefix & keys(i)
        Dim finalCounter
        finalCounter = groupCounters.Item(keys(i))

        On Error Resume Next
        shell.RegWrite regPath & prefixGroupKey, finalCounter, "REG_DWORD"

        If Err.Number = 0 Then
            savedCount = savedCount + 1
        End If

        Err.Clear
    Next

    MsgBox "✅ Registry Updated Successfully!" & vbCrLf & vbCrLf & _
           "Prefix: " & detectedPrefix & vbCrLf & _
           "Groups updated: " & savedCount & vbCrLf & vbCrLf & _
           "Assembly Cloner and Part Renaming will now" & vbCrLf & _
           "continue numbering from these values.", _
           vbInformation, "Registry Updated"
End Sub

' ===================================================================
Sub ScanAssemblyForPrefix(asmDoc, groupCounters, ByRef detectedPrefix)
    ' Recursively scan assembly for heritage file names
    ' Extract prefix and find highest number for each group

    On Error Resume Next

    Call ProcessAssemblyOccurrences(asmDoc, groupCounters, detectedPrefix, "ROOT")

    On Error GoTo 0
End Sub

' ===================================================================
Sub ProcessAssemblyOccurrences(asmDoc, groupCounters, ByRef detectedPrefix, level)
    ' Recursively process assembly occurrences to find heritage files

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

                ' Skip OldVersions
                If InStr(1, LCase(fullPath), "\oldversions\", vbTextCompare) = 0 Then

                    Dim fileName
                    fileName = GetFileNameFromPath(fullPath)

                    ' Only process parts (.ipt)
                    If LCase(Right(fileName, 4)) = ".ipt" Then
                        Call ParseHeritageFileName(fileName, groupCounters, detectedPrefix)
                    End If

                    ' Recurse into sub-assemblies
                    If LCase(Right(fileName, 4)) = ".iam" Then
                        Call ProcessAssemblyOccurrences(doc, groupCounters, detectedPrefix, level & ">" & fileName)
                    End If

                End If
            End If

            Err.Clear
        End If
    Next
End Sub

' ===================================================================
Sub ParseHeritageFileName(fileName, groupCounters, ByRef detectedPrefix)
    ' Parse heritage file name: PREFIX-###-GROUP###.ipt
    ' Example: NCRH01-000-PL173.ipt

    ' Remove .ipt extension
    Dim baseName
    baseName = Left(fileName, Len(fileName) - 4)

    ' Pattern: PREFIX-GROUPNUMBER where PREFIX ends with dash and contains at least one dash
    ' Examples:
    '   NCRH01-000-PL173 -> Prefix: NCRH01-000-, Group: PL, Number: 173
    '   WALKWAY-3-B25    -> Prefix: WALKWAY-3-, Group: B, Number: 25

    ' Find the LAST dash (before the group+number)
    Dim lastDashPos
    lastDashPos = InStrRev(baseName, "-")

    If lastDashPos > 0 Then
        Dim groupAndNumber
        groupAndNumber = Mid(baseName, lastDashPos + 1)

        ' Parse group and number
        ' Group is letters, Number is digits
        Dim groupPart
        Dim numberPart
        groupPart = ""
        numberPart = ""

        Dim j
        For j = 1 To Len(groupAndNumber)
            Dim char
            char = Mid(groupAndNumber, j, 1)

            If IsNumeric(char) And numberPart = "" And groupPart <> "" Then
                ' Start of number - everything before is group
                numberPart = Mid(groupAndNumber, j)
                Exit For
            ElseIf Not IsNumeric(char) Then
                groupPart = groupPart & char
            End If
        Next

        If groupPart <> "" And IsNumeric(numberPart) Then
            ' Valid heritage file name detected
            Dim prefix
            prefix = Left(baseName, lastDashPos)

            ' Set detected prefix (first time only)
            If detectedPrefix = "" Then
                detectedPrefix = prefix
            ElseIf detectedPrefix <> prefix Then
                ' Different prefix found - skip
                Exit Sub
            End If

            ' Update highest number for this group
            Dim currentNumber
            currentNumber = CInt(numberPart)

            If groupCounters.Exists(groupPart) Then
                Dim existingNumber
                existingNumber = groupCounters.Item(groupPart)
                If currentNumber > existingNumber Then
                    groupCounters.Item(groupPart) = currentNumber
                End If
            Else
                groupCounters.Add groupPart, currentNumber
            End If
        End If
    End If
End Sub

' ===================================================================
Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function