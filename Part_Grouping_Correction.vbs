' ******************************************************************************
' PART GROUPING CORRECTION TOOL
' ******************************************************************************
' Purpose: Rename parts from one grouping to another (e.g., IPE → B)
'          Updates all assembly references and IDW drawings
'
' Workflow:
'   1. Scan assembly for all parts and groupings
'   2. Display unique groupings found
'   3. Select source grouping to change
'   4. Select target grouping
'   5. Ask for prefix
'   6. Check registry for next available numbers
'   7. Rename parts sequentially using heritage method
'   8. Update all assembly references
'   9. Ask for IDW folder location
'  10. Update IDW references
'
' Date: January 21, 2026
' ******************************************************************************

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim registryPath
Dim partList()  ' Array to store part information
Dim partCount
Dim groupings  ' Dictionary to store unique groupings
Dim currentPrefix

' Registry path
registryPath = "HKEY_CURRENT_USER\Software\InventorRenamer\"

' ******************************************************************************
' MAIN ENTRY POINT
' ******************************************************************************
Sub Main
    On Error Resume Next

    ' Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    partCount = 0
    Set groupings = CreateObject("Scripting.Dictionary")

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "Inventor is not running!" & vbCrLf & _
               "Please open Inventor and your assembly first.", vbCritical, "Error"
        WScript.Quit
    End If
    Err.Clear

    ' Get active document
    Set invDoc = invApp.ActiveDocument
    If invDoc.DocumentType <> 12291 Then  ' kAssemblyDocumentObject
        MsgBox "Please open an assembly file (.iam) in Inventor first.", vbExclamation, "Wrong Document Type"
        WScript.Quit
    End If

    ' Create log file
    Dim logPath
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Grouping_Correction_Log.txt"
    Set logFile = fso.CreateTextFile(logPath, True)
    LogMessage "=== PART GROUPING CORRECTION TOOL ==="
    LogMessage "Date: " & Now()
    LogMessage "Assembly: " & invDoc.FullFileName
    LogMessage ""

    ' Step 1: Scan assembly for parts and groupings
    LogMessage "STEP 1: Scanning assembly for parts..."
    Call ScanAssembly(invDoc.ComponentDefinition.Occurrences, "")

    LogMessage "Total parts found: " & partCount
    LogMessage "Unique groupings: " & groupings.Count
    LogMessage ""

    ' Display groupings to user
    Dim groupingKeys
    groupingKeys = groupings.Keys
    Dim groupingList
    groupingList = "=== GROUPINGS FOUND IN ASSEMBLY ===" & vbCrLf & vbCrLf
    Dim i
    For i = 0 To groupings.Count - 1
        groupingList = groupingList & "[" & (i + 1) & "] " & groupingKeys(i) & _
                      " (" & groupings(groupingKeys(i)) & " parts)" & vbCrLf
    Next
    groupingList = groupingList & vbCrLf & "Total: " & groupings.Count & " unique groupings"

    LogMessage groupingList
    MsgBox groupingList, vbInformation, "Groupings Found"

    ' Step 2: Select source grouping
    Dim sourceGrouping
    sourceGrouping = InputBox("Enter the grouping you want to CHANGE FROM:" & vbCrLf & vbCrLf & _
                             Join(groupingKeys, vbCrLf), "Select Source Grouping")

    If sourceGrouping = "" Then
        LogMessage "CANCELLED: No source grouping selected"
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        logFile.Close
        WScript.Quit
    End If

    ' Verify source grouping exists
    If Not groupings.Exists(sourceGrouping) Then
        LogMessage "ERROR: Source grouping '" & sourceGrouping & "' not found!"
        MsgBox "Grouping '" & sourceGrouping & "' not found in assembly!", vbExclamation, "Error"
        logFile.Close
        WScript.Quit
    End If

    LogMessage "Source grouping selected: " & sourceGrouping

    ' Step 3: Select target grouping
    Dim targetGrouping
    targetGrouping = InputBox("Enter the grouping you want to CHANGE TO:" & vbCrLf & vbCrLf & _
                             "Available options: PL, B, CH, A, FL, LPL, P, SQ, IPE, R, FLG, etc." & vbCrLf & vbCrLf & _
                             "Example: Enter 'B' to change to beam grouping", "Select Target Grouping")

    If targetGrouping = "" Then
        LogMessage "CANCELLED: No target grouping selected"
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        logFile.Close
        WScript.Quit
    End If

    If sourceGrouping = targetGrouping Then
        LogMessage "ERROR: Source and target groupings are the same!"
        MsgBox "Source and target groupings cannot be the same!", vbExclamation, "Error"
        logFile.Close
        WScript.Quit
    End If

    LogMessage "Target grouping selected: " & targetGrouping

    ' Step 4: Ask for prefix
    currentPrefix = InputBox("Enter the prefix for part numbering:" & vbCrLf & vbCrLf & _
                            "Example: NCRH01-000-", "Enter Prefix", "NCRH01-000-")

    If currentPrefix = "" Then
        LogMessage "CANCELLED: No prefix entered"
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        logFile.Close
        WScript.Quit
    End If

    ' Ensure prefix ends with dash
    If Right(currentPrefix, 1) <> "-" Then
        currentPrefix = currentPrefix & "-"
    End If

    LogMessage "Prefix: " & currentPrefix

    ' Step 5: Check registry for next available number
    Dim nextNumber
    nextNumber = GetNextRegistryNumber(currentPrefix, targetGrouping)

    LogMessage "Registry check for " & currentPrefix & targetGrouping & ": Starting from " & nextNumber

    ' Step 6: Show summary and confirm
    Dim confirmMsg
    confirmMsg = "=== GROUPING CORRECTION SUMMARY ===" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "Source Grouping: " & sourceGrouping & vbCrLf
    confirmMsg = confirmMsg & "Target Grouping: " & targetGrouping & vbCrLf
    confirmMsg = confirmMsg & "Prefix: " & currentPrefix & vbCrLf
    confirmMsg = confirmMsg & "Parts to rename: " & groupings(sourceGrouping) & vbCrLf
    confirmMsg = confirmMsg & "Starting Number: " & nextNumber & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "NEW NAMES WILL BE:" & vbCrLf

    Dim partIdx
    For partIdx = 0 To partCount - 1
        If partList(partIdx)("Group") = sourceGrouping Then
            confirmMsg = confirmMsg & "  " & partList(partIdx)("OldName") & " -> " & _
                        currentPrefix & targetGrouping & (nextNumber + partList(partIdx)("NewIndex")) & ".ipt" & vbCrLf
        End If
    Next

    confirmMsg = confirmMsg & vbCrLf & "Do you want to proceed?"

    LogMessage "Waiting for user confirmation..."
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Changes") <> vbYes Then
        LogMessage "CANCELLED: User declined changes"
        MsgBox "Operation cancelled.", vbInformation, "Cancelled"
        logFile.Close
        WScript.Quit
    End If

    LogMessage "User confirmed. Proceeding with renaming..."
    LogMessage ""

    ' Step 7: Rename parts
    LogMessage "STEP 7: Renaming parts..."
    Dim renameSuccess
    renameSuccess = RenameParts(sourceGrouping, targetGrouping, nextNumber)

    If Not renameSuccess Then
        LogMessage "ERROR: Renaming failed!"
        MsgBox "Part renaming failed! Check log file for details.", vbCritical, "Error"
        logFile.Close
        WScript.Quit
    End If

    LogMessage "Parts renamed successfully!"
    LogMessage ""

    ' Step 8: Update assembly references
    LogMessage "STEP 8: Updating assembly references..."
    Call UpdateAssemblyReferences

    LogMessage "Assembly references updated!"
    LogMessage ""

    ' Step 9: Update IDW references
    LogMessage "STEP 9: Updating IDW references..."
    Call UpdateIDWReferences

    LogMessage "IDW references updated!"
    LogMessage ""

    ' Success!
    LogMessage "=== GROUPING CORRECTION COMPLETE ==="
    logFile.Close

    Dim successMsg
    successMsg = "Grouping correction completed successfully!" & vbCrLf & vbCrLf
    successMsg = successMsg & "Parts renamed: " & groupings(sourceGrouping) & vbCrLf
    successMsg = successMsg & "From: " & sourceGrouping & " -> To: " & targetGrouping & vbCrLf
    successMsg = successMsg & vbCrLf & "Log file saved to:" & vbCrLf & logPath

    MsgBox successMsg, vbInformation, "Success"
End Sub

' ******************************************************************************
' SCAN ASSEMBLY RECURSIVELY
' ******************************************************************************
Sub ScanAssembly(occurrences, path)
    Dim occ
    For Each occ In occurrences
        Dim fullPath
        fullPath = path & occ.Name

        If occ.DefinitionDocumentType <> 12291 Then  ' kAssemblyDocumentObject
            ' This is a part - process it
            ProcessPart occ, fullPath
        Else
            ' This is a sub-assembly - recurse
            Call ScanAssembly(occ.SubOccurrences, fullPath & " -> ")
        End If
    Next
End Sub

' ******************************************************************************
' PROCESS INDIVIDUAL PART
' ******************************************************************************
Sub ProcessPart(occurrence, path)
    On Error Resume Next

    Dim partDoc, partPath, partName, description, groupCode

    partPath = occurrence.ReferencedFileDescriptor.FullFileName
    partName = fso.GetFileName(partPath)

    ' Open part to read description
    Set partDoc = invApp.Documents.Open(partPath, False)

    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not open part: " & partName
        Err.Clear
        Exit Sub
    End If

    ' Get description from Design Tracking Properties
    Dim propertySet
    Set propertySet = partDoc.PropertySets.Item("Design Tracking Properties")
    Dim descriptionProp
    Set descriptionProp = propertySet.Item("Description")
    description = Trim(descriptionProp.Value)

    partDoc.Close

    ' Determine grouping
    groupCode = GetPartGrouping(description, partName)

    ' Store part information
    ReDim Preserve partList(partCount)
    Set partList(partCount) = CreateObject("Scripting.Dictionary")
    partList(partCount)("Path") = partPath
    partList(partCount)("OldName") = partName
    partList(partCount)("Description") = description
    partList(partCount)("Group") = groupCode
    partList(partCount)("NewIndex") = -1  ' Will be assigned when renaming

    ' Track groupings
    If groupings.Exists(groupCode) Then
        groupings(groupCode) = groupings(groupCode) + 1
    Else
        groupings.Add groupCode, 1
    End If

    partCount = partCount + 1
End Sub

' ******************************************************************************
' GET PART GROUPING FROM DESCRIPTION
' ******************************************************************************
Function GetPartGrouping(description, fileName)
    Dim desc
    desc = UCase(Trim(description))

    ' Check for IPE sections (special case - will be moved to B)
    If InStr(desc, "IPE") > 0 Then
        GetPartGrouping = "IPE"
    ElseIf InStr(desc, "HEA") > 0 Or InStr(desc, "HEB") > 0 Or InStr(desc, "HEM") > 0 Then
        GetPartGrouping = "IPE"
    ' UB or UC sections -> B
    ElseIf InStr(desc, "UB") > 0 Or InStr(desc, "UC") > 0 Then
        GetPartGrouping = "B"
    ' Platework: PL + S355JR
    ElseIf InStr(desc, "PL") > 0 And InStr(desc, "S355JR") > 0 Then
        GetPartGrouping = "PL"
    ' Liners: PL + anything else (NOT S355JR)
    ElseIf InStr(desc, "PL") > 0 Then
        GetPartGrouping = "LPL"
    ' Angles: L prefix (but not PL, FL, etc.)
    ElseIf (InStr(desc, " L") > 0 Or Left(desc, 2) = "L ") And _
           InStr(desc, "LPL") = 0 And InStr(desc, "PL") = 0 And _
           InStr(desc, "FL") = 0 Then
        GetPartGrouping = "A"
    ' Channels: PFC or TFC prefix
    ElseIf InStr(desc, "PFC") > 0 Or InStr(desc, "TFC") > 0 Then
        GetPartGrouping = "CH"
    ' Circular hollow sections
    ElseIf InStr(desc, "CHS") > 0 Then
        GetPartGrouping = "P"
    ' Square/rectangular hollow
    ElseIf InStr(desc, "SHS") > 0 Or InStr(desc, "RHS") > 0 Then
        GetPartGrouping = "SQ"
    ' Flatbar
    ElseIf InStr(desc, "FL") > 0 And InStr(desc, "FLG") = 0 Then
        GetPartGrouping = "FL"
    ' Flanges
    ElseIf InStr(desc, "FLG") > 0 Then
        GetPartGrouping = "FLG"
    ' Rings
    ElseIf InStr(desc, "R ") > 0 Or InStr(desc, " RING") > 0 Then
        GetPartGrouping = "R"
    ' Check filename for existing heritage format
    ElseIf InStr(fileName, "-") > 2 Then
        Dim parts, groupCode
        parts = Split(fileName, "-")
        If UBound(parts) >= 2 Then
            groupCode = ExtractGroupCode(parts(2))
            If groupCode <> "" Then
                GetPartGrouping = groupCode
                Exit Function
            End If
        End If
    ' Default: use description
    Else
        GetPartGrouping = "MISC"
    End If
End Function

' ******************************************************************************
' EXTRACT GROUP CODE FROM HERITAGE FILENAME
' ******************************************************************************
Function ExtractGroupCode(groupPart)
    Dim i, code
    code = ""
    For i = 1 To Len(groupPart)
        Dim c
        c = Mid(groupPart, i, 1)
        If c >= "A" And c <= "Z" Then
            code = code & c
        ElseIf c >= "0" And c <= "9" Then
            If code <> "" Then
                Exit For
            End If
        End If
    Next
    ExtractGroupCode = code
End Function

' ******************************************************************************
' GET NEXT REGISTRY NUMBER
' ******************************************************************************
Function GetNextRegistryNumber(prefix, groupCode)
    On Error Resume Next

    Dim regKey, currentValue
    regKey = registryPath & prefix & groupCode

    ' Create shell object for registry access
    Dim shell
    Set shell = CreateObject("WScript.Shell")

    ' Try to read existing value
    currentValue = shell.RegRead(regKey)

    If Err.Number <> 0 Then
        ' Key doesn't exist - start from 1
        GetNextRegistryNumber = 1
        Err.Clear
    Else
        ' Key exists - continue from next number
        GetNextRegistryNumber = CInt(currentValue) + 1
    End If

    LogMessage "Registry key: " & regKey & " = " & (GetNextRegistryNumber - 1)
End Function

' ******************************************************************************
' RENAME PARTS USING HERITAGE METHOD
' ******************************************************************************
Function RenameParts(sourceGrouping, targetGrouping, startNumber)
    On Error Resume Next

    Dim partIdx, currentNumber
    currentNumber = startNumber

    For partIdx = 0 To partCount - 1
        If partList(partIdx)("Group") = sourceGrouping Then
            Dim oldPath, oldName, newFileName, newPath

            oldPath = partList(partIdx)("Path")
            oldName = partList(partIdx)("OldName")
            newFileName = currentPrefix & targetGrouping & currentNumber & ".ipt"
            newPath = fso.GetParentFolderName(oldPath) & "\" & newFileName

            LogMessage "Renaming: " & oldName & " -> " & newFileName

            ' Open part document
            Dim partDoc
            Set partDoc = invApp.Documents.Open(oldPath, False)

            If Err.Number <> 0 Then
                LogMessage "ERROR: Could not open " & oldName & " - " & Err.Description
                Err.Clear
                RenameParts = False
                Exit Function
            End If

            ' Save as heritage file
            partDoc.SaveAs newPath, True

            If Err.Number <> 0 Then
                LogMessage "ERROR: Could not save " & newFileName & " - " & Err.Description
                Err.Clear
                partDoc.Close
                RenameParts = False
                Exit Function
            End If

            ' Update part list with new information
            partList(partIdx)("NewPath") = newPath
            partList(partIdx)("NewName") = newFileName
            partList(partIdx)("NewIndex") = currentNumber - startNumber

            partDoc.Close

            ' Update registry
            Call UpdateRegistry(currentPrefix, targetGrouping, currentNumber)

            currentNumber = currentNumber + 1
        End If
    Next

    RenameParts = True
End Function

' ******************************************************************************
' UPDATE REGISTRY COUNTER
' ******************************************************************************
Sub UpdateRegistry(prefix, groupCode, number)
    On Error Resume Next

    Dim regKey, shell
    regKey = registryPath & prefix & groupCode
    Set shell = CreateObject("WScript.Shell")

    shell.RegWrite regKey, number, "REG_DWORD"

    If Err.Number = 0 Then
        LogMessage "Registry updated: " & regKey & " = " & number
    Else
        LogMessage "WARNING: Could not update registry: " & Err.Description
        Err.Clear
    End If
End Sub

' ******************************************************************************
' UPDATE ASSEMBLY REFERENCES
' ******************************************************************************
Sub UpdateAssemblyReferences
    On Error Resume Next

    ' Re-scan assembly to update all references
    Call UpdateAssemblyReferencesRecursive(invDoc.ComponentDefinition.Occurrences)

    ' Save main assembly
    invDoc.Save2 True
    LogMessage "Main assembly saved: " & invDoc.FullFileName
End Sub

' ******************************************************************************
' UPDATE ASSEMBLY REFERENCES RECURSIVE
' ******************************************************************************
Sub UpdateAssemblyReferencesRecursive(occurrences)
    Dim occ
    For Each occ In occurrences
        If occ.DefinitionDocumentType = 12291 Then  ' kAssemblyDocumentObject
            ' Update sub-assembly references
            Call UpdateAssemblyReferencesInSubAssembly(occ)

            ' Recurse
            Call UpdateAssemblyReferencesRecursive(occ.SubOccurrences)
        End If
    Next
End Sub

' ******************************************************************************
' UPDATE REFERENCES IN SUB-ASSEMBLY
' ******************************************************************************
Sub UpdateAssemblyReferencesInSubAssembly(occurrence)
    On Error Resume Next

    Dim subAsmDoc
    Set subAsmDoc = invApp.Documents.Open(occurrence.ReferencedFileDescriptor.FullFileName, False)

    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not open sub-assembly: " & occurrence.Name
        Err.Clear
        Exit Sub
    End If

    LogMessage "Updating references in: " & fso.GetFileName(subAsmDoc.FullFileName)

    ' Find and replace occurrences
    Dim occ, partIdx
    For Each occ In subAsmDoc.ComponentDefinition.Occurrences
        If occ.DefinitionDocumentType <> 12291 Then  ' Not assembly
            Dim occPath
            occPath = occ.ReferencedFileDescriptor.FullFileName

            ' Check if this part was renamed
            For partIdx = 0 To partCount - 1
                If partList(partIdx).Exists("NewPath") Then
                    If occPath = partList(partIdx)("Path") Then
                        LogMessage "  Replacing: " & partList(partIdx)("OldName") & " -> " & partList(partIdx)("NewName")
                        occ.Replace partList(partIdx)("NewPath"), True
                        Exit For
                    End If
                End If
            Next
        End If
    Next

    ' Save and close
    subAsmDoc.Save2 True
    subAsmDoc.Close
End Sub

' ******************************************************************************
' UPDATE IDW REFERENCES
' ******************************************************************************
Sub UpdateIDWReferences
    ' Ask user for IDW folder location
    Dim folderPath, shell, folderDialog
    Set shell = CreateObject("Shell.Application")

    ' Use PowerShell to show folder browser
    Dim psCommand
    psCommand = "powershell.exe -NoProfile -Command ""Add-Type -AssemblyName System.Windows.Forms; $folder = New-Object System.Windows.Forms.FolderBrowserDialog; $folder.Description = 'Select folder containing IDW drawings to update'; $folder.ShowNewFolderButton = $false; if($folder.ShowDialog() -eq 'OK'){$folder.SelectedPath}"""

    Dim exec
    Set exec = shell.Exec(psCommand)

    Do While exec.Status = 0
        WScript.Sleep 100
    Loop

    folderPath = exec.StdOut.ReadAll
    ' Clean up the path (remove newlines)
    folderPath = Replace(folderPath, vbCrLf, "")
    folderPath = Replace(folderPath, vbCr, "")
    folderPath = Replace(folderPath, vbLf, "")
    folderPath = Trim(folderPath)

    If folderPath = "" Or Not fso.FolderExists(folderPath) Then
        LogMessage "CANCELLED: No IDW folder selected"
        MsgBox "IDW update skipped.", vbInformation, "Skipped"
        Exit Sub
    End If

    LogMessage "IDW folder: " & folderPath

    ' Find all IDW files
    Dim idwFiles
    Set idwFiles = FindIDWFiles(folderPath)

    LogMessage "IDW files found: " & idwFiles.Count

    If idwFiles.Count = 0 Then
        LogMessage "No IDW files found in folder."
        MsgBox "No IDW files found in selected folder.", vbExclamation, "No Files"
        Exit Sub
    End If

    ' Process each IDW
    Dim idwPath
    For Each idwPath In idwFiles
        Call UpdateIDWFile(idwPath)
    Next
End Sub

' ******************************************************************************
' FIND ALL IDW FILES IN FOLDER
' ******************************************************************************
Function FindIDWFiles(folderPath)
    Dim folder, file, subFolder
    Set folder = fso.GetFolder(folderPath)

    Dim result
    Set result = CreateObject("Scripting.Dictionary")

    ' Check all files in this folder
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "idw" Then
            result.Add file.Path, file.Path
        End If
    Next

    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        Dim subResults
        Set subResults = FindIDWFiles(subFolder.Path)

        Dim subPath
        For Each subPath In subResults.Keys
            result.Add subPath, subPath
        Next
    Next

    Set FindIDWFiles = result
End Function

' ******************************************************************************
' UPDATE SINGLE IDW FILE
' ******************************************************************************
Sub UpdateIDWFile(idwPath)
    On Error Resume Next

    LogMessage "Processing IDW: " & fso.GetFileName(idwPath)

    ' Open drawing document
    Dim drawDoc
    Set drawDoc = invApp.Documents.Open(idwPath, False)

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not open " & idwPath & " - " & Err.Description
        Err.Clear
        Exit Sub
    End If

    Dim updateCount
    updateCount = 0

    ' Get FileDescriptor collection for references
    Dim fdColl
    Set fdColl = drawDoc.FileReferences

    ' Update each renamed part reference
    Dim partIdx, fd
    For Each fd In fdColl
        Dim refPath
        refPath = fd.FullFileName

        ' Check if this reference was renamed
        For partIdx = 0 To partCount - 1
            If partList(partIdx).Exists("NewPath") Then
                ' Match by old path
                If refPath = partList(partIdx)("Path") Then
                    LogMessage "  Updating: " & partList(partIdx)("OldName") & " -> " & partList(partIdx)("NewName")

                    ' Replace reference using Design Assistant method
                    fd.ReplaceReference partList(partIdx)("NewPath")

                    updateCount = updateCount + 1
                    Exit For
                End If
            End If
        Next
    Next

    ' Save and close
    drawDoc.Save2 True
    drawDoc.Close

    LogMessage "  Updated " & updateCount & " references"
End Sub

' ******************************************************************************
' LOG MESSAGE
' ******************************************************************************
Sub LogMessage(msg)
    WScript.Echo msg
    logFile.WriteLine msg
End Sub

' ******************************************************************************
' RUN MAIN
' ******************************************************************************
Call Main
