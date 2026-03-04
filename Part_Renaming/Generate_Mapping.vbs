' Generate_Mapping.vbs - Creates STEP_1_MAPPING.txt for a given folder
' Usage: cscript Generate_Mapping.vbs "path\to\folder"

Dim fso, folder, mappingFile
Dim sourceFolder, destFolder
Dim fileCount

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: cscript Generate_Mapping.vbs ""path\to\folder"""
    WScript.Quit 1
End If

sourceFolder = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")

' Create mapping file in the folder
mappingFile = fso.BuildPath(sourceFolder, "STEP_1_MAPPING.txt")
Set outFile = fso.CreateTextFile(mappingFile, True)

outFile.WriteLine("# Generated mapping file for folder: " & sourceFolder)
outFile.WriteLine("# Format: OLD_FILENAME|OLD_FULLPATH|NEW_FILENAME|NEW_FULLPATH")
outFile.WriteLine("")

fileCount = 0

' Recursively scan for .ipt and .iam files
ScanFolder sourceFolder, sourceFolder

outFile.WriteLine("")
outFile.WriteLine("# Total files mapped: " & fileCount)
outFile.Close

WScript.Echo "Mapping file created: " & mappingFile
WScript.Echo "Total files mapped: " & fileCount

Sub ScanFolder(basePath, currentPath)
    Dim folder, files, subfolders, file, relPath

    Set folder = fso.GetFolder(currentPath)

    ' Process files
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Path)) = "ipt" Or LCase(fso.GetExtensionName(file.Path)) = "iam" Then
            relPath = Mid(file.Path, Len(basePath) + 2) ' Remove base path + \
            ' For this simple mapping, assume old name is same as new name for inspection
            ' In reality, the cloner renames them, but this shows current state
            outFile.WriteLine(file.Name & "|" & file.Path & "|" & file.Name & "|" & file.Path)
            fileCount = fileCount + 1
        End If
    Next

    ' Process subfolders
    For Each subfolders In folder.SubFolders
        ScanFolder basePath, subfolders.Path
    Next
End Sub