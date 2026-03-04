Option Explicit

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim basePath
basePath = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\21. SSCR04 - Primary Floats D&R Screen Station\N1SCR06-000\000 Structure & Walkway\"

Dim originalFile, backupFile, fixedFile
originalFile = basePath & "STEP_1_MAPPING.txt"
backupFile = basePath & "STEP_1_MAPPING_BACKUP.txt"
fixedFile = basePath & "STEP_1_MAPPING_FIXED.txt"

' Backup original
If fso.FileExists(originalFile) Then
    fso.CopyFile originalFile, backupFile, True
    WScript.Echo "Backed up original to: " & backupFile
End If

' Replace with fixed version
fso.CopyFile fixedFile, originalFile, True
WScript.Echo "Replaced original with fixed version"
WScript.Echo "You can now run the IDW Fixer again!"
