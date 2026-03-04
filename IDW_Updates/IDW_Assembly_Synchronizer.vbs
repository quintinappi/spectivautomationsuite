Option Explicit

' ==============================================================================
' IDW-ASSEMBLY SYNCHRONIZER - DYNAMIC REFERENCE FIXER
' ==============================================================================
' Syncs IDW references to match whatever the parent assembly currently uses
' Handles edge cases where folder structure breaks STEP 1's reference updating
'
' HOW IT WORKS:
' 1. Find IDW's parent assembly
' 2. Read what files the assembly currently references
' 3. Match IDW references to assembly occurrences
' 4. Update IDW to use same files as assembly
'
' WHEN TO USE:
' - When STEP 1 worked but IDWs still reference old files
' - When main assembly is in one folder, parts in another
' - When folder structure prevents normal reference updating
' ==============================================================================

Dim fso, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Dim g_LogFile
Dim g_ProcessedCount, g_UpdatedCount, g_ErrorCount

' ==============================================================================
' MAIN EXECUTION
' ==============================================================================

WScript.Echo "=========================================="
WScript.Echo "  IDW-ASSEMBLY SYNCHRONIZER"
WScript.Echo "=========================================="
WScript.Echo ""
WScript.Echo "This tool syncs IDW references to match"
WScript.Echo "whatever the parent assembly uses."
WScript.Echo ""
WScript.Echo "Handles edge cases where folder structure"
WScript.Echo "breaks normal reference updating."
WScript.Echo ""

' Ask user for folder to process
Dim folderPath
folderPath = BrowseForFolder("Select folder containing IDW files to sync:")

If folderPath = "" Then
    WScript.Echo "No folder selected. Exiting."
    WScript.Quit
End If

WScript.Echo "Selected folder: " & folderPath
WScript.Echo ""

' Initialize log file
Dim logPath
logPath = folderPath & "\IDW_Sync_Log.txt"
Set g_LogFile = fso.CreateTextFile(logPath, True)

LogMessage "=========================================="
LogMessage "IDW-ASSEMBLY SYNCHRONIZER"
LogMessage "=========================================="
LogMessage "Started: " & Now()
LogMessage "Folder: " & folderPath
LogMessage ""

' Initialize counters
g_ProcessedCount = 0
g_UpdatedCount = 0
g_ErrorCount = 0

' Get Inventor application
Dim invApp
On Error Resume Next
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Inventor is not running!"
    WScript.Echo "Please start Inventor and try again."
    LogMessage "ERROR: Inventor not running"
    g_LogFile.Close
    WScript.Quit
End If
On Error GoTo 0

WScript.Echo "Connected to Inventor"
WScript.Echo ""

' Process all IDW files in folder
Call ProcessFolder(folderPath, invApp)

' Show results
WScript.Echo ""
WScript.Echo "=========================================="
WScript.Echo "  SYNC COMPLETE"
WScript.Echo "=========================================="
WScript.Echo "IDW files processed: " & g_ProcessedCount
WScript.Echo "References updated: " & g_UpdatedCount
WScript.Echo "Errors: " & g_ErrorCount
WScript.Echo ""
WScript.Echo "Log saved to: " & logPath
WScript.Echo ""

LogMessage ""
LogMessage "=========================================="
LogMessage "SUMMARY"
LogMessage "=========================================="
LogMessage "IDW files processed: " & g_ProcessedCount
LogMessage "References updated: " & g_UpdatedCount
LogMessage "Errors: " & g_ErrorCount
LogMessage "Completed: " & Now()

g_LogFile.Close

WScript.Echo "Press any key to exit..."
WScript.StdIn.ReadLine()

' ==============================================================================
' PROCESS FOLDER - FIND AND SYNC ALL IDW FILES
' ==============================================================================
Sub ProcessFolder(folderPath, invApp)
    Dim folder
    Set folder = fso.GetFolder(folderPath)

    ' Process IDW files in current folder
    Dim file
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "idw" Then
            Call ProcessIDW(file.Path, invApp)
        End If
    Next

    ' Process subfolders recursively
    Dim subfolder
    For Each subfolder In folder.SubFolders
        Call ProcessFolder(subfolder.Path, invApp)
    Next
End Sub

' ==============================================================================
' PROCESS IDW - SYNC SINGLE IDW FILE TO ITS ASSEMBLY
' ==============================================================================
Sub ProcessIDW(idwPath, invApp)
    g_ProcessedCount = g_ProcessedCount + 1

    Dim idwName
    idwName = fso.GetFileName(idwPath)

    WScript.Echo "Processing: " & idwName
    LogMessage ""
    LogMessage "=========================================="
    LogMessage "IDW: " & idwName
    LogMessage "=========================================="

    On Error Resume Next

    ' Find parent assembly
    Dim assemblyPath
    assemblyPath = FindParentAssembly(idwPath)

    If assemblyPath = "" Then
        WScript.Echo "  ERROR: Could not find parent assembly"
        LogMessage "ERROR: Parent assembly not found"
        g_ErrorCount = g_ErrorCount + 1
        Exit Sub
    End If

    Dim assemblyName
    assemblyName = fso.GetFileName(assemblyPath)
    WScript.Echo "  Assembly: " & assemblyName
    LogMessage "Assembly: " & assemblyPath

    ' Open assembly and build occurrence map
    Dim assemblyDoc
    Set assemblyDoc = invApp.Documents.Open(assemblyPath, False)

    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not open assembly - " & Err.Description
        LogMessage "ERROR: Could not open assembly - " & Err.Description
        g_ErrorCount = g_ErrorCount + 1
        Err.Clear
        Exit Sub
    End If

    WScript.Echo "  Building occurrence map..."
    LogMessage "Building occurrence map from assembly..."

    Dim occurrenceMap
    Set occurrenceMap = BuildOccurrenceMap(assemblyDoc)

    LogMessage "Found " & occurrenceMap.Count & " occurrences in assembly"

    ' Open IDW
    Dim idwDoc
    Set idwDoc = invApp.Documents.Open(idwPath, False)

    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not open IDW - " & Err.Description
        LogMessage "ERROR: Could not open IDW - " & Err.Description
        assemblyDoc.Close(True)
        g_ErrorCount = g_ErrorCount + 1
        Err.Clear
        Exit Sub
    End If

    WScript.Echo "  Syncing references..."
    LogMessage "Syncing IDW references to assembly..."

    ' Process all sheets
    Dim updateCount
    updateCount = 0

    Dim sheet
    For Each sheet In idwDoc.Sheets
        Dim view
        For Each view In sheet.DrawingViews
            If Not view Is Nothing Then
                updateCount = updateCount + SyncViewReferences(view, occurrenceMap, assemblyName)
            End If
        Next
    Next

    ' Save IDW if updates were made
    If updateCount > 0 Then
        idwDoc.Save2(True)
        WScript.Echo "  SUCCESS: Updated " & updateCount & " reference(s)"
        LogMessage "SUCCESS: Updated " & updateCount & " reference(s)"
        g_UpdatedCount = g_UpdatedCount + updateCount
    Else
        WScript.Echo "  No updates needed"
        LogMessage "No updates needed - all references already correct"
    End If

    ' Close documents
    idwDoc.Close(True)
    assemblyDoc.Close(True)

    On Error GoTo 0
End Sub

' ==============================================================================
' FIND PARENT ASSEMBLY - DETERMINE WHICH ASSEMBLY THIS IDW BELONGS TO
' ==============================================================================
Function FindParentAssembly(idwPath)
    FindParentAssembly = ""

    ' Strategy 1: Same name (MGY-100-SCR-01-50.idw → MGY-100-SCR-01-50.iam)
    Dim baseName
    baseName = fso.GetBaseName(idwPath)
    Dim idwFolder
    idwFolder = fso.GetParentFolderName(idwPath)

    Dim testPath
    testPath = idwFolder & "\" & baseName & ".iam"

    If fso.FileExists(testPath) Then
        FindParentAssembly = testPath
        Exit Function
    End If

    ' Strategy 2: Search parent folders for matching assembly
    Dim searchFolder
    searchFolder = idwFolder

    Dim i
    For i = 1 To 3  ' Search up to 3 levels
        Dim folder
        Set folder = fso.GetFolder(searchFolder)

        Dim file
        For Each file In folder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "iam" Then
                ' Check if assembly name matches IDW pattern
                If InStr(baseName, fso.GetBaseName(file.Name)) > 0 Or _
                   InStr(fso.GetBaseName(file.Name), baseName) > 0 Then
                    FindParentAssembly = file.Path
                    Exit Function
                End If
            End If
        Next

        ' Move up one level
        If fso.GetParentFolderName(searchFolder) <> "" Then
            searchFolder = fso.GetParentFolderName(searchFolder)
        Else
            Exit For
        End If
    Next

    ' Strategy 3: Look for any .iam in same folder
    Set folder = fso.GetFolder(idwFolder)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "iam" Then
            FindParentAssembly = file.Path
            Exit Function
        End If
    Next
End Function

' ==============================================================================
' BUILD OCCURRENCE MAP - CREATE MAPPING OF PART NAMES TO CURRENT FILE PATHS
' ==============================================================================
Function BuildOccurrenceMap(assemblyDoc)
    Dim occMap
    Set occMap = CreateObject("Scripting.Dictionary")
    occMap.CompareMode = 1  ' Case-insensitive

    ' Recursively process all occurrences
    Call ProcessOccurrences(assemblyDoc.ComponentDefinition.Occurrences, occMap)

    Set BuildOccurrenceMap = occMap
End Function

Sub ProcessOccurrences(occurrences, occMap)
    If occurrences Is Nothing Then Exit Sub

    Dim occ
    For Each occ In occurrences
        ' Get base name (strip occurrence index)
        Dim occName
        occName = occ.Name

        ' Strip :1, :2, etc.
        If InStr(occName, ":") > 0 Then
            occName = Left(occName, InStr(occName, ":") - 1)
        End If

        ' Get current file path
        Dim currentPath
        currentPath = occ.ReferencedFileDescriptor.FullFileName

        ' Get just filename
        Dim currentFile
        currentFile = fso.GetFileName(currentPath)

        ' Add to map (use occurrence name as key)
        If Not occMap.Exists(occName) Then
            occMap.Add occName, currentPath
        End If

        ' Process sub-assembly occurrences recursively
        On Error Resume Next
        Dim subDef
        Set subDef = occ.Definition
        If Err.Number = 0 Then
            If Not subDef Is Nothing Then
                Call ProcessOccurrences(subDef.Occurrences, occMap)
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Sub

' ==============================================================================
' SYNC VIEW REFERENCES - UPDATE IDW VIEW TO MATCH ASSEMBLY OCCURRENCES
' ==============================================================================
Function SyncViewReferences(view, occurrenceMap, assemblyName)
    SyncViewReferences = 0

    On Error Resume Next

    ' Get model reference
    Dim modelRef
    Set modelRef = view.ReferencedDocumentDescriptor

    If Err.Number <> 0 Or modelRef Is Nothing Then
        Err.Clear
        Exit Function
    End If

    ' Get current file
    Dim currentFile
    currentFile = modelRef.FullFileName

    If currentFile = "" Then
        Exit Function
    End If

    ' Extract base filename (no path, no extension)
    Dim baseName
    baseName = fso.GetBaseName(currentFile)

    ' Remove occurrence index if present
    If InStr(baseName, ":") > 0 Then
        baseName = Left(baseName, InStr(baseName, ":") - 1)
    End If

    ' Look up in occurrence map
    If occurrenceMap.Exists(baseName) Then
        Dim newPath
        newPath = occurrenceMap.Item(baseName)

        ' Only update if different
        If currentFile <> newPath Then
            LogMessage "  Updating: " & fso.GetFileName(currentFile) & " → " & fso.GetFileName(newPath)

            ' Use ReplaceReference method
            Dim fd
            Set fd = view.ReferencedDocumentDescriptor
            fd.ReplaceReference currentFile, newPath

            If Err.Number = 0 Then
                SyncViewReferences = SyncViewReferences + 1
            Else
                LogMessage "  ERROR: " & Err.Description
                g_ErrorCount = g_ErrorCount + 1
                Err.Clear
            End If
        End If
    End If

    On Error GoTo 0
End Function

' ==============================================================================
' BROWSE FOR FOLDER
' ==============================================================================
Function BrowseForFolder(prompt)
    BrowseForFolder = ""

    ' Use PowerShell folder browser
    Dim psCommand
    psCommand = "powershell -Command """ & _
                "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog; " & _
                "$dialog.Description = '" & prompt & "'; " & _
                "$dialog.ShowNewFolderButton = $false; " & _
                "if ($dialog.ShowDialog() -eq 'OK') { Write-Output $dialog.SelectedPath }"""

    Dim result
    result = shell.Exec(psCommand).StdOut.ReadAll()

    ' Clean up result (remove newlines)
    result = Replace(result, vbCrLf, "")
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "")
    result = Trim(result)

    BrowseForFolder = result
End Function

' ==============================================================================
' LOGGING
' ==============================================================================
Sub LogMessage(message)
    WScript.Echo message
    g_LogFile.WriteLine message
End Sub