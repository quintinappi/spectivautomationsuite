Option Explicit

' ==============================================================================
' REVERSE MAPPING FILE - Swap source and destination in mapping file
' ==============================================================================
' This creates a reversed mapping file to undo incorrect IDW updates
' ==============================================================================

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Ask for original mapping file
Dim originalPath
originalPath = InputBox("Enter path to ORIGINAL mapping file (STEP_1_MAPPING.txt):", "Select Original Mapping")
If originalPath = "" Then
    MsgBox "Cancelled.", vbInformation
    WScript.Quit
End If

If Not fso.FileExists(originalPath) Then
    MsgBox "File not found: " & originalPath, vbCritical
    WScript.Quit
End If

' Create reversed mapping file path
Dim reversedPath
reversedPath = fso.GetParentFolderName(originalPath) & "\STEP_1_MAPPING_REVERSED.txt"

' Read original file
Dim originalFile
Set originalFile = fso.OpenTextFile(originalPath, 1)

' Create reversed file
Dim reversedFile
Set reversedFile = fso.CreateTextFile(reversedPath, True)

reversedFile.WriteLine "# REVERSED MAPPING FILE - Generated: " & Now
reversedFile.WriteLine "# This file has SOURCE and DESTINATION swapped to UNDO incorrect updates"
reversedFile.WriteLine "# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description"
reversedFile.WriteLine ""

Dim lineCount
lineCount = 0

Do While Not originalFile.AtEndOfStream
    Dim line
    line = originalFile.ReadLine
    
    ' Skip comments and empty lines
    If Left(Trim(line), 1) = "#" Or Trim(line) = "" Then
        reversedFile.WriteLine line
    Else
        ' Parse and swap fields
        Dim parts
        parts = Split(line, "|")
        
        If UBound(parts) >= 5 Then
            ' Original format: SourcePath|DestPath|SourceFile|DestFile|Type|Operation
            ' Reversed format: DestPath|SourcePath|DestFile|SourceFile|Type|REVERSED
            Dim sourcePath, destPath, sourceFile, destFile, fileType, operation
            sourcePath = parts(0)
            destPath = parts(1)
            sourceFile = parts(2)
            destFile = parts(3)
            fileType = parts(4)
            operation = parts(5)
            
            ' Write reversed entry (swap positions 0↔1 and 2↔3)
            reversedFile.WriteLine destPath & "|" & sourcePath & "|" & destFile & "|" & sourceFile & "|" & fileType & "|REVERSED"
            lineCount = lineCount + 1
        End If
    End If
Loop

originalFile.Close
reversedFile.Close

MsgBox "REVERSED MAPPING CREATED!" & vbCrLf & vbCrLf & _
       "Original: " & originalPath & vbCrLf & _
       "Reversed: " & reversedPath & vbCrLf & vbCrLf & _
       "Reversed " & lineCount & " entries" & vbCrLf & vbCrLf & _
       "Now run IDW_Reference_Updater with the REVERSED file to fix source IDWs", vbInformation

WScript.Echo "Done!"
