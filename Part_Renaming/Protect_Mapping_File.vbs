Option Explicit

' ==============================================================================
' PROTECT MAPPING FILE - SET HIDDEN ATTRIBUTE
' ==============================================================================
' Makes STEP_1_MAPPING.txt hidden (but still readable by scripts)
' Run this once to protect the mapping file from accidental deletion
' ==============================================================================

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Get mapping file location
Dim scriptDir
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Dim rootDir
rootDir = fso.GetParentFolderName(scriptDir)
Dim mappingFilePath
mappingFilePath = rootDir & "\STEP_1_MAPPING.txt"

WScript.Echo "Looking for mapping file at: " & mappingFilePath

If Not fso.FileExists(mappingFilePath) Then
    MsgBox "ERROR: STEP_1_MAPPING.txt not found!" & vbCrLf & vbCrLf & _
           "Expected location:" & vbCrLf & _
           mappingFilePath & vbCrLf & vbCrLf & _
           "Run STEP 1 first to create the mapping file.", vbExclamation, "File Not Found"
    WScript.Quit
End If

' Get the file object
Dim mappingFile
Set mappingFile = fso.GetFile(mappingFilePath)

' Check current attributes
Dim isHidden
isHidden = (mappingFile.Attributes And 2) = 2

If isHidden Then
    MsgBox "Mapping file is already hidden!" & vbCrLf & vbCrLf & _
           "File: " & mappingFilePath & vbCrLf & vbCrLf & _
           "The file is protected from accidental deletion." & vbCrLf & _
           "Scripts can still read and update it normally.", vbInformation, "Already Protected"
    WScript.Quit
End If

' Set hidden attribute
On Error Resume Next
mappingFile.Attributes = mappingFile.Attributes Or 2  ' Add hidden flag

If Err.Number <> 0 Then
    MsgBox "ERROR: Could not set hidden attribute!" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Try running as administrator.", vbCritical, "Protection Failed"
    WScript.Quit
End If

On Error GoTo 0

MsgBox "SUCCESS: Mapping file is now HIDDEN!" & vbCrLf & vbCrLf & _
       "File: " & mappingFilePath & vbCrLf & vbCrLf & _
       "✓ Protected from accidental deletion" & vbCrLf & _
       "✓ Still readable by all scripts" & vbCrLf & _
       "✓ Still writable by all scripts" & vbCrLf & vbCrLf & _
       "The mapping file won't show in Windows Explorer" & vbCrLf & _
       "unless you enable 'Show hidden files'.", vbInformation, "Mapping File Protected"

WScript.Echo "Mapping file protected successfully!"