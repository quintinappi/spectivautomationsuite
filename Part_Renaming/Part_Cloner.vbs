Option Explicit

' ==============================================================================
' PART CLONER - Copy Single Part to New Location
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' This script:
' 1. Detects currently open part in Inventor
' 2. Asks for destination folder
' 3. Copies part file to destination
' 4. Optionally renames the part
' 5. Reads and displays iProperties
' ==============================================================================

Dim g_LogFileNum
Dim g_LogPath

Call PART_CLONER_MAIN()

Sub PART_CLONER_MAIN()
    Call StartLogging
    LogMessage "=== PART CLONER ==="
    LogMessage "Copy single part to new isolated location"

    Dim result
    result = MsgBox("PART CLONER" & vbCrLf & vbCrLf & _
                    "This will:" & vbCrLf & _
                    "1. Detect your currently open part" & vbCrLf & _
                    "2. Copy part to a NEW folder" & vbCrLf & _
                    "3. Optionally rename the part" & vbCrLf & _
                    "4. Display part iProperties" & vbCrLf & vbCrLf & _
                    "Make sure your source part is open in Inventor!" & vbCrLf & vbCrLf & _
                    "Continue?", vbYesNo + vbQuestion, "Part Cloner")

    If result = vbNo Then
        LogMessage "User cancelled workflow"
        Exit Sub
    End If

    ' Connect to Inventor
    Dim invApp
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: Inventor not running"
        MsgBox "ERROR: Inventor is not running!", vbCritical
        Exit Sub
    End If

    LogMessage "SUCCESS: Connected to Inventor"
    Err.Clear

    ' Detect open part
    Dim activeDoc
    Set activeDoc = invApp.ActiveDocument

    If activeDoc Is Nothing Then
        LogMessage "ERROR: No active document"
        MsgBox "ERROR: No document is open in Inventor!", vbCritical
        Exit Sub
    End If

    If LCase(Right(activeDoc.FullFileName, 4)) <> ".ipt" Then
        LogMessage "ERROR: Active document is not a part: " & activeDoc.FullFileName
        MsgBox "ERROR: Please open a part (.ipt) file!", vbCritical
        Exit Sub
    End If

    LogMessage "DETECTED: " & activeDoc.FullFileName

    ' Display iProperties
    Call DisplayPartProperties(activeDoc)

    ' Ask for destination folder
    Dim destFolder
    destFolder = BrowseForFolder("Select destination folder for part copy", activeDoc.FullFileName)

    If destFolder = "" Then
        LogMessage "User cancelled folder selection"
        Exit Sub
    End If

    LogMessage "DESTINATION: " & destFolder

    ' Get new part name with registry-based numbering
    Dim originalName
    originalName = GetFileNameFromPath(activeDoc.FullFileName)

    ' Get prefix from user
    Dim userPrefix
    userPrefix = InputBox("ENTER PREFIX" & vbCrLf & vbCrLf & _
                         "Example: NCRH01-000-" & vbCrLf & vbCrLf & _
                         "Enter your project prefix:", _
                         "Prefix", "NCRH01-000-")

    If userPrefix = "" Then
        LogMessage "User cancelled prefix entry"
        Exit Sub
    End If

    ' Ensure prefix ends with dash
    If Right(userPrefix, 1) <> "-" Then
        userPrefix = userPrefix & "-"
    End If

    LogMessage "PREFIX: " & userPrefix

    ' Get part group from user
    Dim partGroup
    partGroup = InputBox("ENTER PART GROUP" & vbCrLf & vbCrLf & _
                        "Common groups:" & vbCrLf & _
                        "  PL  - Platework (S355JR)" & vbCrLf & _
                        "  B   - Beams/Columns (UB/UC)" & vbCrLf & _
                        "  CH  - Channels (PFC/TFC)" & vbCrLf & _
                        "  A   - Angles" & vbCrLf & _
                        "  FL  - Flatbar" & vbCrLf & _
                        "  LPL - Liners (non-S355JR)" & vbCrLf & _
                        "  SQ  - Square/Rect Hollow (SHS)" & vbCrLf & _
                        "  P   - Circular Hollow (CHS)" & vbCrLf & vbCrLf & _
                        "Enter part group code:", _
                        "Part Group", "")

    If partGroup = "" Then
        LogMessage "User cancelled part group entry"
        Exit Sub
    End If

    ' Uppercase the group
    partGroup = UCase(Trim(partGroup))
    LogMessage "PART GROUP: " & partGroup

    ' Get suggested name from registry
    Dim suggestedName
    suggestedName = GetSuggestedNameFromRegistry(userPrefix, partGroup)

    ' Get new part name with suggested value pre-filled
    Dim newName
    newName = InputBox("NEW PART NAME" & vbCrLf & vbCrLf & _
                      "Original:  " & originalName & vbCrLf & _
                      "Suggested: " & suggestedName & vbCrLf & vbCrLf & _
                      "Registry scan found next available number for " & userPrefix & partGroup & vbCrLf & vbCrLf & _
                      "Enter new name (or accept suggestion):", _
                      "Rename Part", suggestedName)

    If newName = "" Then
        newName = suggestedName
    End If

    If LCase(Right(newName, 4)) <> ".ipt" Then
        newName = newName & ".ipt"
    End If

    ' Copy the file
    Dim destPath
    destPath = destFolder & "\" & newName

    On Error Resume Next
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile activeDoc.FullFileName, destPath, True

    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not copy part: " & Err.Description
        MsgBox "ERROR: Could not copy part file!" & vbCrLf & Err.Description, vbCritical
        Exit Sub
    End If

    Err.Clear
    LogMessage "COPIED: " & originalName & " -> " & newName

    ' Success message
    Dim summaryMsg
    summaryMsg = "PART CLONE COMPLETED!" & vbCrLf & vbCrLf & _
                 "✅ Part copied to: " & destPath & vbCrLf & vbCrLf & _
                 "The part is now isolated and ready for modification." & vbCrLf & vbCrLf & _
                 "Log: " & g_LogPath

    MsgBox summaryMsg, vbInformation, "Success!"

    Call StopLogging
End Sub

Sub DisplayPartProperties(partDoc)
    LogMessage "DISPLAYING PART PROPERTIES"

    On Error Resume Next

    Dim propSets
    Set propSets = partDoc.PropertySets

    Dim summary
    summary = "PART PROPERTIES:" & vbCrLf & vbCrLf

    ' Summary properties
    summary = summary & "Summary:" & vbCrLf
    summary = summary & "- Title: " & GetPropertyValue(propSets, "Inventor Summary Information", 2) & vbCrLf
    summary = summary & "- Subject: " & GetPropertyValue(propSets, "Inventor Summary Information", 3) & vbCrLf
    summary = summary & "- Author: " & GetPropertyValue(propSets, "Inventor Summary Information", 4) & vbCrLf
    summary = summary & "- Keywords: " & GetPropertyValue(propSets, "Inventor Summary Information", 5) & vbCrLf
    summary = summary & "- Comments: " & GetPropertyValue(propSets, "Inventor Summary Information", 6) & vbCrLf
    summary = summary & "- Last Saved By: " & GetPropertyValue(propSets, "Inventor Summary Information", 8) & vbCrLf

    summary = summary & vbCrLf & "Project:" & vbCrLf
    summary = summary & "- Part Number: " & GetPropertyValue(propSets, "Design Tracking Properties", 5) & vbCrLf
    summary = summary & "- Stock Number: " & GetPropertyValue(propSets, "Design Tracking Properties", 55) & vbCrLf
    summary = summary & "- Description: " & GetPropertyValue(propSets, "Design Tracking Properties", 29) & vbCrLf

    MsgBox summary, vbInformation, "Part iProperties"

    LogMessage "Properties displayed to user"
End Sub

Function GetPropertyValue(propSets, setName, propId)
    On Error Resume Next
    Dim propSet
    Set propSet = propSets.Item(setName)
    If Not propSet Is Nothing Then
        Dim prop
        Set prop = propSet.ItemByPropId(propId)
        If Not prop Is Nothing Then
            GetPropertyValue = prop.Value
        Else
            GetPropertyValue = "(not set)"
        End If
    Else
        GetPropertyValue = "(set not found)"
    End If
    If Err.Number <> 0 Then
        GetPropertyValue = "(error)"
        Err.Clear
    End If
End Function

Function BrowseForFolder(prompt, sourcePath)
    ' Traditional folder picker with classic tree view
    On Error Resume Next

    Dim shellApp
    Set shellApp = CreateObject("Shell.Application")

    ' Get source part directory
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourceDir
    sourceDir = fso.GetParentFolderName(sourcePath)

    ' Show folder browser dialog
    ' &H0041 = BIF_NEWDIALOGSTYLE + BIF_RETURNONLYFSDIRS
    Dim folder
    Set folder = shellApp.BrowseForFolder(0, "Select destination folder", &H0041, sourceDir)

    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If
End Function

Function GetFileNameFromPath(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(fullPath)
End Function

Function GetSuggestedNameFromRegistry(prefix, partGroup)
    ' Scan registry for the counter and suggest next available number
    On Error Resume Next

    Dim regPath
    regPath = "HKEY_CURRENT_USER\Software\InventorRenamer\" & prefix & partGroup

    Dim shell
    Set shell = CreateObject("WScript.Shell")

    Dim currentCounter
    currentCounter = 0

    ' Try to read from registry
    currentCounter = shell.RegRead(regPath)

    If Err.Number <> 0 Then
        ' Key doesn't exist, start from 0
        currentCounter = 0
        LogMessage "Registry key not found: " & regPath & " - starting from 1"
        Err.Clear
    Else
        LogMessage "Registry found: " & regPath & " = " & currentCounter
    End If

    ' Next number is current + 1
    Dim nextNumber
    nextNumber = currentCounter + 1

    ' Build suggested name: PREFIX-GROUP###.ipt
    GetSuggestedNameFromRegistry = prefix & partGroup & nextNumber & ".ipt"

    LogMessage "SUGGESTED NAME: " & GetSuggestedNameFromRegistry & " (next after " & currentCounter & ")"
End Function

Sub StartLogging()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    Dim rootDir
    rootDir = fso.GetParentFolderName(scriptDir)
    Dim logsDir
    logsDir = rootDir & "\Logs"

    If Not fso.FolderExists(logsDir) Then
        fso.CreateFolder(logsDir)
    End If

    g_LogPath = logsDir & "\PartCloner_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFileNum = fso.CreateTextFile(g_LogPath, True)
End Sub

Sub LogMessage(message)
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
    End If
    WScript.Echo message
End Sub

Sub StopLogging()
    If Not IsEmpty(g_LogFileNum) Then
        g_LogFileNum.Close
    End If
End Sub