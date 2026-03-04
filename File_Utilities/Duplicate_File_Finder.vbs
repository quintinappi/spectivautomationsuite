' ===================================================================
' INVENTOR RENAMER - DUPLICATE PART NUMBER SCANNER
' ===================================================================
' Scans project for duplicate heritage part numbers

Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Get project root directory
Dim projectRoot
projectRoot = InputBox("Enter project root directory to scan:" & vbCrLf & vbCrLf & _
                      "Example: D:\Pentalin\Backup\Backup 2\Workspaces\Workspace\05. Model", _
                      "Project Root Directory", _
                      "D:\Pentalin\Backup\Backup 2\Workspaces\Workspace\05. Model")

If projectRoot = "" Then
    MsgBox "Scan cancelled.", vbInformation
    WScript.Quit
End If

If Not fso.FolderExists(projectRoot) Then
    MsgBox "Directory not found: " & projectRoot, vbCritical
    WScript.Quit
End If

' Dictionary to store part numbers: partNumber -> array of file paths
Dim partNumbers
Set partNumbers = CreateObject("Scripting.Dictionary")

' Scan project recursively
Call ScanDirectoryRecursively(projectRoot, partNumbers)

' Analyze for duplicates
Call ShowDuplicateReport(partNumbers)

' ===================================================================
Sub ScanDirectoryRecursively(folderPath, partNumbers)
    Dim folder, file, subfolder

    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)

    If Err.Number <> 0 Then
        Exit Sub
    End If

    ' Process files in current folder
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".ipt" Then
            ' Check if it's a heritage file (NCRH01-000-*)
            If Left(UCase(file.Name), 10) = "NCRH01-000" Then
                Call ExtractPartNumber(file.Path, file.Name, partNumbers)
            End If
        End If
    Next

    ' Process subfolders recursively
    For Each subfolder In folder.SubFolders
        Call ScanDirectoryRecursively(subfolder.Path, partNumbers)
    Next

    On Error GoTo 0
End Sub

' ===================================================================
Sub ExtractPartNumber(filePath, fileName, partNumbers)
    ' Extract part number from filename like NCRH01-000-CH13.ipt

    Dim baseName
    If LCase(Right(fileName, 4)) = ".ipt" Then
        baseName = Left(fileName, Len(fileName) - 4)
    Else
        baseName = fileName
    End If

    ' Find the last dash and extract group+number
    Dim lastDashPos
    lastDashPos = InStrRev(baseName, "-")

    If lastDashPos > 0 Then
        Dim partNumber
        partNumber = Mid(baseName, lastDashPos + 1)  ' e.g., "CH13"

        ' Store in dictionary
        If partNumbers.Exists(partNumber) Then
            ' Add to existing array
            Dim existingPaths
            existingPaths = partNumbers.Item(partNumber)
            ReDim Preserve existingPaths(UBound(existingPaths) + 1)
            existingPaths(UBound(existingPaths)) = filePath
            partNumbers.Item(partNumber) = existingPaths
        Else
            ' Create new array
            Dim newPaths(0)
            newPaths(0) = filePath
            partNumbers.Add partNumber, newPaths
        End If
    End If
End Sub

' ===================================================================
Sub ShowDuplicateReport(partNumbers)
    Dim report
    report = "DUPLICATE PART NUMBER SCAN RESULTS" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf

    Dim totalFiles, duplicateCount, duplicatePartCount
    totalFiles = 0
    duplicateCount = 0
    duplicatePartCount = 0

    ' Count total files and find duplicates
    Dim partKeys, i
    partKeys = partNumbers.Keys

    For i = 0 To UBound(partKeys)
        Dim partNumber
        partNumber = partKeys(i)

        Dim filePaths
        filePaths = partNumbers.Item(partNumber)

        totalFiles = totalFiles + UBound(filePaths) + 1

        If UBound(filePaths) > 0 Then
            ' Duplicate found!
            duplicatePartCount = duplicatePartCount + 1
            duplicateCount = duplicateCount + UBound(filePaths) + 1

            report = report & "🚨 DUPLICATE: " & partNumber & " (" & (UBound(filePaths) + 1) & " files)" & vbCrLf

            Dim j
            For j = 0 To UBound(filePaths)
                Dim shortPath
                shortPath = Right(filePaths(j), 60)
                If Len(filePaths(j)) > 60 Then shortPath = "..." & shortPath
                report = report & "   " & (j + 1) & ". " & shortPath & vbCrLf
            Next

            report = report & vbCrLf
        End If
    Next

    ' Add summary
    If duplicatePartCount = 0 Then
        report = report & "✅ NO DUPLICATES FOUND!" & vbCrLf & vbCrLf
        report = report & "Scanned " & totalFiles & " heritage files." & vbCrLf
        report = report & "All part numbers are unique." & vbCrLf
    Else
        report = report & "SUMMARY:" & vbCrLf & String(20, "-") & vbCrLf
        report = report & "Total files scanned: " & totalFiles & vbCrLf
        report = report & "Unique part numbers: " & (UBound(partKeys) + 1 - duplicatePartCount) & vbCrLf
        report = report & "Duplicate part numbers: " & duplicatePartCount & vbCrLf
        report = report & "Total duplicate files: " & (duplicateCount - duplicatePartCount) & vbCrLf
    End If

    report = report & vbCrLf & String(60, "=")

    ' Show report
    MsgBox report, vbInformation, "Duplicate Scanner Results"

    ' Option to save report to file
    If duplicatePartCount > 0 Then
        Dim saveReport
        saveReport = MsgBox("Duplicates found! Save detailed report to file?", vbYesNo + vbQuestion, "Save Report")

        If saveReport = vbYes Then
            Call SaveReportToFile(report)
        End If
    End If
End Sub

' ===================================================================
Sub SaveReportToFile(report)
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

    Dim reportPath
    reportPath = scriptDir & "\DuplicateReport_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & ".txt"

    Dim reportFile
    Set reportFile = fso.CreateTextFile(reportPath, True)
    reportFile.Write report
    reportFile.Close

    MsgBox "Report saved to:" & vbCrLf & reportPath, vbInformation, "Report Saved"
End Sub