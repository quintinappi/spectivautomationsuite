Option Explicit

' ==============================================================================
' MAPPING RECOVERY TOOL - Recover mapping from existing renamed assembly
' ==============================================================================
' This script:
' 1. Scans the assembly to find all components
' 2. Attempts to recover original names by checking OldVersions folder
' 3. Creates STEP_1_MAPPING.txt with recovered mappings
' ==============================================================================

Call RECOVER_MAPPING()

Sub RECOVER_MAPPING()
    Dim result
    result = MsgBox("MAPPING RECOVERY TOOL" & vbCrLf & vbCrLf & _
                    "This will attempt to recover mappings from a renamed assembly." & vbCrLf & _
                    "1. Scan assembly for all components" & vbCrLf & _
                    "2. Search OldVersions folder for original files" & vbCrLf & _
                    "3. Create STEP_1_MAPPING.txt" & vbCrLf & vbCrLf & _
                    "Make sure your renamed assembly is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Mapping Recovery")

    If result = vbNo Then
        Exit Sub
    End If

    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        MsgBox "ERROR: Inventor is not running!", vbCritical
        Exit Sub
    End If
    Err.Clear

    ' Get active document
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If activeDoc Is Nothing Or Not (activeDoc.DocumentType = 12290) Then
        MsgBox "ERROR: No assembly is open in Inventor!", vbCritical
        Exit Sub
    End If

    WScript.Echo "Assembly: " & activeDoc.DisplayName
    WScript.Echo "Path: " & activeDoc.FullFileName

    ' Build mapping from OldVersions files
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim asmDir
    asmDir = fso.GetParentFolderName(activeDoc.FullFileName)

    Dim oldVersionsDir
    oldVersionsDir = asmDir & "\OldVersions"

    Dim hasOldVersions
    hasOldVersions = fso.FolderExists(oldVersionsDir)

    WScript.Echo ""
    WScript.Echo "SCANNING COMPONENTS..."
    WScript.Echo "Assembly directory: " & asmDir
    WScript.Echo "OldVersions folder: " & IIf(hasOldVersions, "FOUND", "NOT FOUND")

    ' Get all components in assembly
    Dim components
    Set components = CreateObject("Scripting.Dictionary")

    Call ScanAssemblyComponents(activeDoc, components)

    WScript.Echo ""
    WScript.Echo "FOUND " & components.Count & " COMPONENTS"

    ' Build mapping dictionary
    Dim mappingDict
    Set mappingDict = CreateObject("Scripting.Dictionary")

    ' Strategy 1: Match by checking if original exists in OldVersions
    If hasOldVersions Then
        WScript.Echo "STRATEGY 1: Matching with OldVersions folder..."
        Call MatchWithOldVersions(components, oldVersionsDir, mappingDict)
    End If

    ' Strategy 2: Match by pattern (if OldVersions not available or incomplete)
    If mappingDict.Count < components.Count Then
        WScript.Echo "STRATEGY 2: Pattern matching..."
        Call MatchByPattern(components, asmDir, mappingDict)
    End If

    ' Save mapping file
    WScript.Echo ""
    WScript.Echo "SAVING MAPPING FILE..."

    Dim mappingFilePath
    mappingFilePath = asmDir & "\STEP_1_MAPPING.txt"

    Dim mappingFile
    Set mappingFile = fso.CreateTextFile(mappingFilePath, True)

    mappingFile.WriteLine "# STEP 1 MAPPING FILE - Recovered: " & Now
    mappingFile.WriteLine "# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description"
    mappingFile.WriteLine ""

    Dim keys
    keys = mappingDict.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim newPath
        newPath = keys(i)
        Dim originalPath
        originalPath = mappingDict.Item(newPath)

        ' Parse file names
        Dim newFile
        newFile = fso.GetFileName(newPath)
        Dim originalFile
        originalFile = fso.GetFileName(originalPath)

        ' Extract group from filename (first 2-3 chars before number)
        Dim group
        group = ExtractGroup(newFile)

        ' Try to get description
        Dim description
        description = "N/A"

        mappingFile.WriteLine originalPath & "|" & newPath & "|" & originalFile & "|" & newFile & "|" & group & "|" & description
    Next

    mappingFile.WriteLine ""
    mappingFile.WriteLine "# End of mapping file"
    mappingFile.Close

    WScript.Echo "MAPPING FILE CREATED: " & mappingFilePath
    WScript.Echo "TOTAL MAPPINGS: " & mappingDict.Count
    WScript.Echo ""
    WScript.Echo "You can now run the IDW Reference Updater!"

    MsgBox "Mapping Recovery Complete!" & vbCrLf & vbCrLf & _
           "Total components: " & components.Count & vbCrLf & _
           "Mappings recovered: " & mappingDict.Count & vbCrLf & vbCrLf & _
           "Mapping file: " & mappingFilePath & vbCrLf & vbCrLf & _
           "Next: Run IDW Reference Updater", vbInformation, "Recovery Complete"
End Sub

Sub ScanAssemblyComponents(assemblyDoc, components)
    ' Recursively scan all components in assembly
    On Error Resume Next

    Dim occurrences
    Set occurrences = assemblyDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        Dim fd
        Set fd = occ.ReferencedFileDescriptor

        If fd.DocumentType = 12292 Then ' kPartDocumentObject
            Dim fullPath
            fullPath = fd.FullFileName

            ' Only add if not already in list (avoid duplicates)
            If Not components.Exists(fullPath) Then
                components.Add fullPath, fullPath
            End If
        End If

        ' Recursively process sub-assembly components
        If fd.DocumentType = 12290 Then ' kAssemblyDocumentObject
            Dim subAssemblyDoc
            Set subAssemblyDoc = occ.Definition.Document
            If Not (subAssemblyDoc Is Nothing) Then
                Call ScanAssemblyComponents(subAssemblyDoc, components)
            End If
        End If
    Next

    Err.Clear
End Sub

Sub MatchWithOldVersions(components, oldVersionsDir, mappingDict)
    ' Try to find original files in OldVersions folder
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim oldVersionsFolder
    Set oldVersionsFolder = fso.GetFolder(oldVersionsDir)

    Dim oldFiles
    Set oldFiles = CreateObject("Scripting.Dictionary")

    ' Build list of files in OldVersions
    Dim file
    For Each file In oldVersionsFolder.Files
        Dim fileName
        fileName = file.Name

        ' Store full path
        oldFiles.Add file.Name, file.Path
    Next

    Dim keys
    keys = components.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim newPath
        newPath = keys(i)

        Dim fso2
        Set fso2 = CreateObject("Scripting.FileSystemObject")
        Dim newFileName
        newFileName = fso2.GetFileName(newPath)

        ' Try to find original in OldVersions
        Dim originalPath
        originalPath = ""

        If oldFiles.Exists(newFileName) Then
            ' Found exact match in OldVersions
            originalPath = oldFiles.Item(newFileName)
        Else
            ' Try removing numbering suffix
            Dim baseName
            baseName = StripNumberSuffix(newFileName)

            If oldFiles.Exists(baseName) Then
                originalPath = oldFiles.Item(baseName)
            End If
        End If

        If originalPath <> "" Then
            If Not mappingDict.Exists(newPath) Then
                mappingDict.Add newPath, originalPath
                WScript.Echo "MAPPED: " & newFileName & " -> " & fso2.GetFileName(originalPath)
            End If
        End If
    Next

    Err.Clear
End Sub

Sub MatchByPattern(components, asmDir, mappingDict)
    ' Fallback: Try to guess original names by removing prefix
    Dim keys
    keys = components.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim newPath
        newPath = keys(i)

        ' If already mapped, skip
        If mappingDict.Exists(newPath) Then
        Else
            Dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            Dim newFileName
            newFileName = fso.GetFileName(newPath)

            ' Try to guess original by removing structured prefix
            ' Example: "SSCR05-001-CH1.ipt" -> "Part1.ipt"
            ' Example: "SSCR05-001-PL1.ipt" -> "Part1.ipt"

            Dim originalFileName
            originalFileName = GuessOriginalName(newFileName)

            Dim originalPath
            originalPath = asmDir & "\" & originalFileName

            mappingDict.Add newPath, originalPath
            WScript.Echo "GUESSED: " & newFileName & " -> " & originalFileName
        End If
    Next
End Sub

Function ExtractGroup(fileName)
    ' Extract group from filename (first chars before number)
    Dim group
    
    ' Try patterns like "SSCR05-001-CH1" -> "CH"
    If InStr(fileName, "-") > 0 Then
        Dim parts
        parts = Split(fileName, "-")
        
        If UBound(parts) >= 2 Then
            Dim lastPart
            lastPart = parts(2)
            
            ' Extract alpha prefix (e.g., "CH1" -> "CH")
            If Len(lastPart) > 1 Then
                group = Left(lastPart, 2)
                
                ' Remove digits if present
                Dim i
                For i = Len(group) To 1 Step -1
                    If Not IsNumeric(Mid(group, i, 1)) Then
                        group = Left(group, i)
                        Exit For
                    End If
                Next
            End If
        End If
    End If
    
    If group = "" Then
        group = "UNKNOWN"
    End If
    
    ExtractGroup = group
End Function

Function StripNumberSuffix(fileName)
    ' Strip numbering suffix from filename
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim baseName
    baseName = fso.GetBaseName(fileName)
    Dim ext
    ext = "." & fso.GetExtensionName(fileName)
    
    ' Remove trailing numbers
    Dim i
    For i = Len(baseName) To 1 Step -1
        If Not IsNumeric(Mid(baseName, i, 1)) Then
            baseName = Left(baseName, i)
            Exit For
        End If
    Next
    
    StripNumberSuffix = baseName & ext
End Function

Function GuessOriginalName(newFileName)
    ' Guess original name by removing structured prefix
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim baseName
    baseName = fso.GetBaseName(newFileName)

    ' Pattern 1: Structured naming like "SSCR05-001-CH1"
    ' Try to convert to "Part1"
    
    ' Count unique positions in baseName
    Dim parts
    parts = Split(baseName, "-")
    
    If UBound(parts) >= 2 Then
        Dim lastPart
        lastPart = parts(2)
        
        ' Extract alpha group (CH, PL, A, etc.)
        Dim groupPrefix
        groupPrefix = ""
        Dim i
        For i = 1 To Len(lastPart)
            Dim char
            char = Mid(lastPart, i, 1)
            If Not IsNumeric(char) Then
                groupPrefix = groupPrefix & char
            Else
                Exit For
            End If
        Next
        
        ' Extract sequence number
        Dim seqNum
        seqNum = ""
        For i = Len(groupPrefix) + 1 To Len(lastPart)
            seqNum = seqNum & Mid(lastPart, i, 1)
        Next
        
        GuessOriginalName = "Part" & seqNum & ".ipt"
    Else
        ' Fallback: use simple counter
        GuessOriginalName = "Part1.ipt"
    End If
End Function
