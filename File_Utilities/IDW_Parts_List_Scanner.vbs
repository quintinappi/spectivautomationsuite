' ==============================================================================
' IDW PARTS LIST SCANNER - MOVE UNRENAMED PARTS
' ==============================================================================
' This tool:
' 1. Scans the currently open IDW drawing
' 2. Extracts all parts from the parts list (BOM)
' 3. RECURSIVELY scans ALL sub-assemblies for parts (handles nested assemblies!)
' 4. Lists all referenced parts
' 5. Finds ALL .ipt files in the same folder as the IDW
' 6. Moves parts NOT in the parts list to "Unrenamed Parts" folder
'
' USE CASE:
' - After running STEP 1 (Part Renaming) on an assembly
' - You have both renamed (heritage) and unrenamed parts in the folder
' - The IDW only references the renamed parts
' - This tool moves the old unreferenced parts to "Unrenamed Parts" for cleanup
' - NOW HANDLES SUB-ASSEMBLIES: Parts inside sub-assemblies are correctly detected
' ==============================================================================

Option Explicit

Const kDrawingDocumentObject = 12292

Dim g_LogFile
Dim g_LogPath
Dim g_PartsList ' Dictionary of part names in the IDW's parts list

Call IDW_PARTS_LIST_SCANNER()

Sub IDW_PARTS_LIST_SCANNER()
    WScript.Echo "=========================================="
    WScript.Echo "  IDW PARTS LIST SCANNER"
    WScript.Echo "=========================================="
    WScript.Echo ""

    ' Initialize logging
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    g_LogPath = scriptDir & "\IDW_Parts_List_Scanner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFile = fso.CreateTextFile(g_LogPath, True)

    ' Initialize parts list dictionary
    Set g_PartsList = CreateObject("Scripting.Dictionary")

    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: No running Inventor instance found!"
        LogMessage "Please start Inventor first."
        g_LogFile.Close
        WScript.Quit 1
    End If

    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear

    ' Check for active document
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If Err.Number <> 0 Or activeDoc Is Nothing Then
        LogMessage "ERROR: No active document found!"
        LogMessage "Please open an IDW drawing first."
        g_LogFile.Close
        WScript.Quit 1
    End If
    Err.Clear

    ' Check if active document is a drawing
    If activeDoc.DocumentType <> kDrawingDocumentObject Then
        LogMessage "ERROR: Active document is not a drawing (IDW)"
        LogMessage "Current: " & activeDoc.DisplayName
        LogMessage "Please open an IDW drawing file."
        g_LogFile.Close
        WScript.Quit 1
    End If

    Dim drawingDoc
    Set drawingDoc = activeDoc
    LogMessage "SUCCESS: Drawing document found: " & drawingDoc.DisplayName
    WScript.Echo ""

    ' Step 1: Extract parts from the drawing's parts list/BOM
    WScript.Echo "STEP 1: Extracting parts from drawing references..."
    LogMessage "STEP 1: Extracting parts from drawing references"
    Call ExtractPartsFromDrawing(drawingDoc)

    If g_PartsList.Count = 0 Then
        LogMessage "ERROR: No parts found in drawing!"
        LogMessage "The drawing may not have any model references."
        g_LogFile.Close
        WScript.Quit 1
    End If

    LogMessage "SUCCESS: Found " & g_PartsList.Count & " parts in drawing"
    WScript.Echo "Found " & g_PartsList.Count & " parts in drawing"
    WScript.Echo ""

    ' Display parts list to user
    Call DisplayPartsList

    ' Step 2: Find all IPT files in the same folder as the IDW
    WScript.Echo "STEP 2: Scanning folder for all IPT files..."
    LogMessage "STEP 2: Scanning folder for all IPT files"

    Dim idwPath, idwFolder
    idwPath = drawingDoc.FullFileName
    idwFolder = GetParentFolder(idwPath)

    LogMessage "FOLDER: " & idwFolder

    Dim allIPTFiles
    Set allIPTFiles = CreateObject("Scripting.Dictionary")
    Call FindAllIPTFiles(idwFolder, allIPTFiles)

    LogMessage "FOUND: " & allIPTFiles.Count & " IPT files in folder"
    WScript.Echo "Found " & allIPTFiles.Count & " IPT files in folder"
    WScript.Echo ""

    ' Step 3: Identify unreferenced files
    WScript.Echo "STEP 3: Identifying unreferenced IPT files..."
    LogMessage "STEP 3: Identifying unreferenced IPT files"

    Dim unreferencedFiles
    Set unreferencedFiles = CreateObject("Scripting.Dictionary")

    Dim keys
    keys = allIPTFiles.Keys
    Dim i
    For i = 0 To UBound(keys)
        Dim filePath
        filePath = keys(i)
        Dim fileName
        fileName = allIPTFiles.Item(filePath)

        ' Check if this file is in the parts list
        If Not g_PartsList.Exists(fileName) Then
            unreferencedFiles.Add filePath, fileName
            LogMessage "UNREFERENCED: " & fileName
        Else
            LogMessage "REFERENCED:   " & fileName
        End If
    Next

    LogMessage "TOTAL: " & unreferencedFiles.Count & " unreferenced files found"

    ' Step 4: Move unreferenced files to "Unrenamed Parts" folder
    If unreferencedFiles.Count > 0 Then
        WScript.Echo ""
        WScript.Echo "STEP 4: Moving " & unreferencedFiles.Count & " unreferenced files to 'Unrenamed Parts' folder..."
        LogMessage "STEP 4: Moving unreferenced files to 'Unrenamed Parts' folder"

        Call DisplayUnreferencedFiles(unreferencedFiles)
        WScript.Echo ""
        WScript.Echo "Moving files now..."

        Call MoveUnreferencedFiles(unreferencedFiles, idwFolder)
    Else
        WScript.Echo ""
        WScript.Echo "SUCCESS: All IPT files are referenced in the drawing!"
        LogMessage "SUCCESS: All IPT files are referenced in the drawing!"
    End If

    WScript.Echo ""
    WScript.Echo "=========================================="
    WScript.Echo "  SCAN COMPLETE"
    WScript.Echo "=========================================="
    WScript.Echo "Log: " & g_LogPath
    WScript.Echo ""

    LogMessage "=== IDW PARTS LIST SCANNER COMPLETED ==="
    g_LogFile.Close
End Sub

Sub ExtractPartsFromDrawing(drawingDoc)
    ' Extract all referenced part files from the drawing

    ' Method 1: Check Sheet BOM (Parts Lists)
    LogMessage "BOM: Checking sheet parts lists..."

    Dim sheets
    Set sheets = drawingDoc.Sheets

    Dim sheetNum
    For sheetNum = 1 To sheets.Count
        Dim sheet
        Set sheet = sheets.Item(sheetNum)

        LogMessage "BOM: Sheet " & sheetNum & " - " & sheet.Name

        ' Check for parts lists on this sheet
        Dim partsLists
        Set partsLists = sheet.PartsLists

        If partsLists.Count > 0 Then
            LogMessage "BOM:   Found " & partsLists.Count & " parts list(s)"

            Dim plNum
            For plNum = 1 To partsLists.Count
                Dim partsList
                Set partsList = partsLists.Item(plNum)

                LogMessage "BOM:     Processing parts list " & plNum

                ' Iterate through parts list rows
                Dim row
                For Each row In partsList.PartsListRows
                    Dim partsListRow
                    Set partsListRow = row

                    ' Get the referenced files for this row (plural!)
                    On Error Resume Next
                    Dim refFiles
                    Set refFiles = partsListRow.ReferencedFiles

                    If Not refFiles Is Nothing Then
                        Dim refFile
                        For Each refFile In refFiles
                            Dim fileName
                            fileName = GetFileNameFromPath(refFile.FullFileName)

                            ' Check file type
                            If LCase(Right(fileName, 4)) = ".ipt" Then
                                ' Part file - add to parts list
                                If Not g_PartsList.Exists(fileName) Then
                                    g_PartsList.Add fileName, refFile.FullFileName
                                    LogMessage "BOM:       Added: " & fileName
                                End If
                            ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                                ' Sub-assembly - recursively scan it
                                LogMessage "BOM:       Found sub-assembly: " & fileName & " - Scanning contents..."
                                Call ScanAssemblyForParts(refFile)
                            End If
                        Next
                    End If
                    Err.Clear
                Next
            Next
        End If
    Next

    ' Method 2: Check all referenced file descriptors (catches parts not in BOM)
    LogMessage "REFS: Checking all referenced file descriptors..."

    Dim fileDescriptors
    Set fileDescriptors = drawingDoc.File.ReferencedFileDescriptors

    Dim fd
    For Each fd In fileDescriptors
        fileName = GetFileNameFromPath(fd.FullFileName)

        ' Check file type
        If LCase(Right(fileName, 4)) = ".ipt" Then
            ' Part file - add to parts list
            If Not g_PartsList.Exists(fileName) Then
                g_PartsList.Add fileName, fd.FullFileName
                LogMessage "REFS:   Added: " & fileName
            Else
                LogMessage "REFS:   Already exists: " & fileName
            End If
        ElseIf LCase(Right(fileName, 4)) = ".iam" Then
            ' Assembly - recursively scan for parts inside
            LogMessage "REFS:   Found assembly: " & fileName & " - Scanning contents..."
            Call ScanAssemblyForParts(fd)
        End If
    Next

    LogMessage "EXTRACT: Total unique parts found: " & g_PartsList.Count
End Sub

' ============================================================================
' RECURSIVE ASSEMBLY SCANNING - Handles sub-assemblies
' ============================================================================
Sub ScanAssemblyForParts(asmFileDesc)
    ' Recursively scan an assembly file descriptor for all parts
    On Error Resume Next
    
    Dim asmDoc
    Set asmDoc = Nothing
    Set asmDoc = asmFileDesc.ReferencedFile
    
    If asmDoc Is Nothing Or Err.Number <> 0 Then
        LogMessage "  ERROR: Could not open assembly: " & GetFileNameFromPath(asmFileDesc.FullFileName)
        Err.Clear
        Exit Sub
    End If
    
    ' Check if this is actually an assembly
    If LCase(Right(asmDoc.FullFileName, 4)) <> ".iam" Then
        Exit Sub
    End If
    
    LogMessage "  Scanning assembly: " & asmDoc.DisplayName
    
    ' Process all occurrences in this assembly
    Call ProcessAssemblyOccurrences(asmDoc)
    
    Err.Clear
End Sub

Sub ProcessAssemblyOccurrences(asmDoc)
    ' Process all occurrences in an assembly document
    On Error Resume Next
    
    Dim occs
    Set occs = asmDoc.ComponentDefinition.Occurrences
    
    If occs Is Nothing Or Err.Number <> 0 Then
        LogMessage "  ERROR: Could not get occurrences from: " & asmDoc.DisplayName
        Err.Clear
        Exit Sub
    End If
    
    LogMessage "  Found " & occs.Count & " occurrences in " & asmDoc.DisplayName
    
    Dim occ
    For Each occ In occs
        ' Skip suppressed occurrences
        If Not occ.Suppressed Then
            Dim defDoc
            Set defDoc = Nothing
            Set defDoc = occ.Definition.Document
            
            If Not defDoc Is Nothing And Err.Number = 0 Then
                Dim fullPath
                fullPath = defDoc.FullFileName
                Dim fileName
                fileName = GetFileNameFromPath(fullPath)
                
                ' Check file type
                If LCase(Right(fileName, 4)) = ".ipt" Then
                    ' Part file - add to parts list
                    If Not g_PartsList.Exists(fileName) Then
                        g_PartsList.Add fileName, fullPath
                        LogMessage "    ADDED PART: " & fileName
                    End If
                ElseIf LCase(Right(fileName, 4)) = ".iam" Then
                    ' Sub-assembly - recurse (skip Content Center and bolted connections)
                    If InStr(LCase(fileName), "bolted connection") = 0 And _
                       InStr(LCase(fileName), "content center") = 0 And _
                       InStr(LCase(fullPath), "\content center\") = 0 Then
                        LogMessage "    FOUND SUB-ASM: " & fileName & " - Recursing..."
                        Call ProcessAssemblyOccurrences(defDoc)
                    Else
                        LogMessage "    SKIP: Content Center/Bolted Connection - " & fileName
                    End If
                End If
            Else
                If Err.Number <> 0 Then
                    LogMessage "  WARNING: Could not resolve occurrence: " & occ.Name & " - " & Err.Description
                    Err.Clear
                End If
            End If
        Else
            LogMessage "  SKIP: Suppressed occurrence - " & occ.Name
        End If
    Next
    
    Err.Clear
End Sub

Sub FindAllIPTFiles(folderPath, allFiles)
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
        If LCase(Right(file.Name, 4)) = ".ipt" Then
            ' Skip files containing "development" in the name
            If InStr(LCase(file.Name), "development") > 0 Then
                LogMessage "SKIP: Development file - " & file.Name
            Else
                allFiles.Add file.Path, file.Name
                LogMessage "IPT: Found " & file.Name
            End If
        End If
    Next

    ' Also check subdirectories (excluding OldVersions)
    Dim subFolder
    For Each subFolder In folder.SubFolders
        If LCase(subFolder.Name) <> "oldversions" Then
            Call FindAllIPTFiles(subFolder.Path, allFiles)
        End If
    Next
End Sub

Sub DisplayPartsList()
    LogMessage ""
    LogMessage "=== PARTS LIST SUMMARY ==="
    LogMessage "Total parts referenced in drawing: " & g_PartsList.Count
    LogMessage ""

    WScript.Echo "Parts referenced in drawing (" & g_PartsList.Count & "):"
    WScript.Echo ""

    Dim keys
    keys = g_PartsList.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim fileName
        fileName = g_PartsList.Item(keys(i))
        LogMessage "  " & (i + 1) & ". " & fileName
        WScript.Echo "  " & (i + 1) & ". " & fileName
    Next

    WScript.Echo ""
    LogMessage ""
End Sub

Sub DisplayUnreferencedFiles(unreferencedFiles)
    WScript.Echo "Unreferenced files to be moved:"

    Dim keys
    keys = unreferencedFiles.Keys

    Dim i
    Dim maxDisplay
    maxDisplay = 20 ' Show up to 20 files

    For i = 0 To UBound(keys)
        If i < maxDisplay Then
            WScript.Echo "  - " & unreferencedFiles.Item(keys(i))
        ElseIf i = maxDisplay Then
            WScript.Echo "  ... and " & (unreferencedFiles.Count - maxDisplay) & " more"
            Exit For
        End If
    Next
End Sub

Sub MoveUnreferencedFiles(unreferencedFiles, sourceFolder)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create "Unrenamed Parts" folder
    Dim unrenamedFolder
    unrenamedFolder = sourceFolder & "\Unrenamed Parts"

    If Not fso.FolderExists(unrenamedFolder) Then
        LogMessage "Creating folder: " & unrenamedFolder
        fso.CreateFolder unrenamedFolder
    Else
        LogMessage "Folder already exists: " & unrenamedFolder
    End If

    ' Move files
    Dim keys
    keys = unreferencedFiles.Keys

    Dim movedCount
    movedCount = 0
    Dim errorCount
    errorCount = 0

    Dim i
    For i = 0 To UBound(keys)
        Dim sourcePath
        sourcePath = keys(i)
        Dim fileName
        fileName = unreferencedFiles.Item(sourcePath)
        Dim destPath
        destPath = unrenamedFolder & "\" & fileName

        ' Check if destination file already exists
        If fso.FileExists(destPath) Then
            LogMessage "WARNING: File already exists in destination: " & fileName
            LogMessage "WARNING: Skipping move to avoid overwrite"
            errorCount = errorCount + 1
        Else
            On Error Resume Next
            fso.MoveFile sourcePath, destPath

            If Err.Number = 0 Then
                LogMessage "MOVED: " & fileName & " -> Unrenamed Parts\"
                movedCount = movedCount + 1
            Else
                LogMessage "ERROR: Failed to move " & fileName & ": " & Err.Description
                errorCount = errorCount + 1
                Err.Clear
            End If
        End If
    Next

    LogMessage "MOVE COMPLETE: " & movedCount & " files moved, " & errorCount & " errors"

    WScript.Echo "=========================================="
    WScript.Echo "Files moved: " & movedCount
    WScript.Echo "Errors: " & errorCount
    WScript.Echo "Destination: " & unrenamedFolder
    WScript.Echo "=========================================="
End Sub

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function GetParentFolder(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(fullPath)
End Function

Sub LogMessage(message)
    WScript.Echo message
    g_LogFile.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
End Sub
