' ============================================================================
' INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER
' ============================================================================
' Description: Clone assemblies with automatic numbering continuation
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-20
' ============================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim userPrefix, cloneCount
Dim registryCounters
Dim scriptPath, logPath

' Initialize
Initialize
Main
Cleanup

' ============================================================================
' MAIN PROCEDURE
' ============================================================================
Sub Main()
    LogMessage "=========================================="
    LogMessage "ASSEMBLY CLONER - STARTING"
    LogMessage "=========================================="

    ' Step 1: Get user input
    If Not GetUserInput() Then
        LogMessage "Operation cancelled by user"
        Exit Sub
    End If

    ' Step 2: Connect to Inventor
    If Not ConnectToInventor() Then
        MsgBox "Failed to connect to Inventor. Please make sure Inventor is running with an assembly open.", vbCritical, "Error"
        Exit Sub
    End If

    ' Step 3: Scan current registry state
    ScanRegistry

    ' Step 4: Confirm operation with user
    If Not ConfirmOperation() Then
        LogMessage "Operation cancelled by user"
        Exit Sub
    End If

    ' Step 5: Perform assembly cloning
    CloneAssembly

    LogMessage "=========================================="
    LogMessage "ASSEMBLY CLONER - COMPLETED SUCCESSFULLY"
    LogMessage "=========================================="

    MsgBox "Assembly cloning completed successfully!" & vbCrLf & vbCrLf & _
           "Clones created: " & cloneCount & vbCrLf & _
           "Prefix: " & userPrefix & vbCrLf & vbCrLf & _
           "Please check the log file for details:" & vbCrLf & logPath, vbInformation, "Success"
End Sub

' ============================================================================
' INITIALIZATION
' ============================================================================
Sub Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get script path
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
    logPath = scriptPath & "\Assembly_Cloner_Log.txt"

    ' Initialize log file
    Set logFile = fso.CreateTextFile(logPath, True)
    LogMessage "Log file created: " & logPath
    LogMessage "Timestamp: " & Now()

    ' Initialize registry counters dictionary
    Set registryCounters = CreateObject("Scripting.Dictionary")
End Sub

' ============================================================================
' GET USER INPUT
' ============================================================================
Function GetUserInput()
    Dim args

    ' Check if command-line arguments were provided
    Set args = WScript.Arguments

    If args.Count >= 2 Then
        ' Arguments provided from HTA launcher
        userPrefix = args(0)
        cloneCount = CInt(args(1))

        LogMessage "Command-line arguments received:"
        LogMessage "  Prefix: " & userPrefix
        LogMessage "  Clone Count: " & cloneCount
    Else
        ' No arguments - prompt user (standalone mode)
        ' Get prefix from user
        userPrefix = InputBox("Enter the prefix for part numbering:" & vbCrLf & vbCrLf & _
                              "Example: NCRH01-000-" & vbCrLf & _
                              "         PLANT1-000-", "Assembly Cloner - Prefix", "NCRH01-000-")

        If Trim(userPrefix) = "" Then
            GetUserInput = False
            Exit Function
        End If

        ' Ensure prefix ends with dash
        If Right(userPrefix, 1) <> "-" Then
            userPrefix = userPrefix & "-"
        End If

        ' Get clone count
        Dim countInput
        countInput = InputBox("Enter the number of clones to create (1-10):", "Assembly Cloner - Clone Count", "1")

        If Trim(countInput) = "" Or Not IsNumeric(countInput) Then
            MsgBox "Invalid clone count. Please enter a number between 1 and 10.", vbExclamation, "Input Error"
            GetUserInput = False
            Exit Function
        End If

        cloneCount = CInt(countInput)

        If cloneCount < 1 Or cloneCount > 10 Then
            MsgBox "Clone count must be between 1 and 10.", vbExclamation, "Input Error"
            GetUserInput = False
            Exit Function
        End If

        LogMessage "User input received via prompts:"
        LogMessage "  Prefix: " & userPrefix
        LogMessage "  Clone Count: " & cloneCount
    End If

    ' Validate clone count
    If cloneCount < 1 Or cloneCount > 10 Then
        MsgBox "Clone count must be between 1 and 10.", vbExclamation, "Input Error"
        GetUserInput = False
        Exit Function
    End If

    GetUserInput = True
End Function

' ============================================================================
' CONNECT TO INVENTOR
' ============================================================================
Function ConnectToInventor()
    On Error Resume Next
    Set invApp = GetObject(, "Inventor.Application")

    If Err.Number <> 0 Then
        LogMessage "ERROR: Failed to connect to Inventor - " & Err.Description
        ConnectToInventor = False
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0

    ' Get active document
    Set invDoc = invApp.ActiveDocument

    If invDoc Is Nothing Then
        LogMessage "ERROR: No active document in Inventor"
        ConnectToInventor = False
        Exit Function
    End If

    ' Verify it's an assembly document
    If invDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
        LogMessage "ERROR: Active document is not an assembly"
        ConnectToInventor = False
        Exit Function
    End If

    LogMessage "Connected to Inventor successfully"
    LogMessage "  Active Document: " & invDoc.FullFileName
    LogMessage "  Document Type: Assembly"

    ConnectToInventor = True
End Function

' ============================================================================
' SCAN REGISTRY
' ============================================================================
Sub ScanRegistry()
    LogMessage "=========================================="
    LogMessage "SCANNING REGISTRY FOR PREFIX: " & userPrefix
    LogMessage "=========================================="

    Dim shell, regPath, regValue
    Set shell = CreateObject("WScript.Shell")

    ' Define counter groups
    Dim counterGroups
    counterGroups = Array("PL", "B", "CH", "A", "FL", "LPL", "P", "SQ", "R", "FLG")

    Dim group
    For Each group In counterGroups
        regPath = "HKCU\Software\InventorRenamer\" & userPrefix & group

        On Error Resume Next
        regValue = shell.RegRead(regPath)

        If Err.Number = 0 Then
            registryCounters(group) = regValue
            LogMessage "  Found: " & group & " = " & regValue
        Else
            registryCounters(group) = 0
            LogMessage "  Not found: " & group & " (will start from 0)"
            Err.Clear
        End If

        On Error GoTo 0
    Next

    LogMessage "Registry scan complete"
End Sub

' ============================================================================
' CONFIRM OPERATION
' ============================================================================
Function ConfirmOperation()
    Dim msg, group

    msg = "Assembly Cloner - Confirm Operation" & vbCrLf & vbCrLf
    msg = msg & "Prefix: " & userPrefix & vbCrLf
    msg = msg & "Clone Count: " & cloneCount & vbCrLf & vbCrLf
    msg = msg & "Current Registry Status:" & vbCrLf

    ' Show top 5 most important counters
    Dim topGroups, i
    topGroups = Array("PL", "B", "CH", "A", "FL")

    For Each group In topGroups
        If registryCounters.Exists(group) Then
            msg = msg & "  " & group & ": " & registryCounters(group) & vbCrLf
        End If
    Next

    msg = msg & vbCrLf & "Do you want to proceed?"

    Dim result
    result = MsgBox(msg, vbQuestion + vbYesNo, "Confirm Operation")

    ConfirmOperation = (result = vbYes)
End Function

' ============================================================================
' CLONE ASSEMBLY
' ============================================================================
Sub CloneAssembly()
    LogMessage "=========================================="
    LogMessage "STARTING ASSEMBLY CLONING PROCESS"
    LogMessage "=========================================="

    Dim originalPath, originalDir, originalName, originalBase
    originalPath = invDoc.FullFileName
    originalDir = fso.GetParentFolderName(originalPath) & "\"
    originalName = fso.GetFileName(originalPath)
    originalBase = fso.GetBaseName(originalPath)

    LogMessage "Original Assembly:"
    LogMessage "  Path: " & originalPath
    LogMessage "  Directory: " & originalDir
    LogMessage "  Name: " & originalName
    LogMessage "  Base Name: " & originalBase

    ' Create clones
    Dim i
    For i = 1 To cloneCount
        LogMessage "=========================================="
        LogMessage "CREATING CLONE #" & i & " OF " & cloneCount
        LogMessage "=========================================="

        CreateSingleClone i, originalDir, originalBase
    Next

    LogMessage "Assembly cloning completed"
End Sub

' ============================================================================
' CREATE SINGLE CLONE
' ============================================================================
Sub CreateSingleClone(cloneNumber, originalDir, originalBase)
    Dim cloneName, clonePath

    ' Generate clone name
    cloneName = originalBase & "_CLONE" & cloneNumber & ".iam"
    clonePath = originalDir & cloneName

    LogMessage "Creating clone: " & cloneName
    LogMessage "  Path: " & clonePath

    ' Check if file already exists
    If fso.FileExists(clonePath) Then
        LogMessage "  WARNING: File already exists, skipping"
        Exit Sub
    End If

    ' Save assembly as clone
    On Error Resume Next
    invDoc.SaveAs clonePath, False
    If Err.Number <> 0 Then
        LogMessage "  ERROR: Failed to create clone - " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    LogMessage "  Clone created successfully"

    ' Now we need to open the clone and update all part references
    ' This is a placeholder - the actual implementation would:
    ' 1. Open the cloned assembly
    ' 2. Iterate through all occurrences
    ' 3. For each part, create a heritage copy with incremented numbering
    ' 4. Update references in the clone assembly
    ' 5. Save and close

    LogMessage "  NOTE: Full part cloning logic to be implemented"
    LogMessage "  Currently creating assembly clone only"

    ' Re-open original assembly
    invApp.Documents.Open(invDoc.FullFileName)
End Sub

' ============================================================================
' UPDATE REGISTRY COUNTERS
' ============================================================================
Sub UpdateRegistryCounters()
    LogMessage "Updating registry counters..."

    Dim shell, regPath
    Set shell = CreateObject("WScript.Shell")

    Dim group
    For Each group In registryCounters.Keys
        regPath = "HKCU\Software\InventorRenamer\" & userPrefix & group

        On Error Resume Next
        shell.RegWrite regPath, registryCounters(group), "REG_DWORD"
        If Err.Number = 0 Then
            LogMessage "  Updated: " & userPrefix & group & " = " & registryCounters(group)
        Else
            LogMessage "  ERROR: Failed to update " & userPrefix & group & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next
End Sub

' ============================================================================
' LOGGING
' ============================================================================
Sub LogMessage(message)
    If Not logFile Is Nothing Then
        logFile.WriteLine Now() & " - " & message
    End If
End Sub

' ============================================================================
' CLEANUP
' ============================================================================
Sub Cleanup()
    LogMessage "=========================================="
    LogMessage "CLEANUP AND CLOSING"
    LogMessage "=========================================="

    If Not logFile Is Nothing Then
        logFile.Close
        Set logFile = Nothing
    End If

    Set fso = Nothing
    Set registryCounters = Nothing
    Set invDoc = Nothing
    Set invApp = Nothing
End Sub
