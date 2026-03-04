' ============================================================================
' CLEAR RENAMER REGISTRY COUNTERS
' ============================================================================
' Description: Clear registry counters for a specific prefix or all prefixes
'              Used when you want to restart numbering from 1
' Author: Spectiv Solutions
' Date: 2026-01-23
' ============================================================================

Option Explicit

Dim shell
Set shell = CreateObject("WScript.Shell")

Dim regBasePath
regBasePath = "HKCU\Software\InventorRenamer\"

' Get the prefix to clear
Dim userChoice
userChoice = MsgBox("CLEAR REGISTRY COUNTERS" & vbCrLf & vbCrLf & _
                    "This tool clears heritage numbering counters" & vbCrLf & _
                    "stored in the Windows Registry." & vbCrLf & vbCrLf & _
                    "YES = Clear counters for a SPECIFIC prefix" & vbCrLf & _
                    "NO = Clear ALL heritage counters" & vbCrLf & _
                    "CANCEL = Exit without changes", _
                    vbYesNoCancel + vbQuestion, "Clear Registry Counters")

If userChoice = vbCancel Then
    WScript.Quit
End If

If userChoice = vbYes Then
    ' Clear specific prefix
    Dim prefixInput
    prefixInput = InputBox("Enter the PREFIX to clear counters for:" & vbCrLf & vbCrLf & _
                           "Example: N1SCR05-000-" & vbCrLf & _
                           "         PLANT1-001-" & vbCrLf & vbCrLf & _
                           "Include the trailing hyphen if applicable.", _
                           "Enter Prefix", "")
    
    If prefixInput = "" Then
        MsgBox "No prefix entered. Operation cancelled.", vbInformation
        WScript.Quit
    End If
    
    ' Clear counters for this prefix
    Dim groupCodes
    groupCodes = Array("CH", "PL", "B", "A", "P", "SQ", "FL", "LPL", "IPE", "OTHER", "FLG", "R")
    
    Dim clearedCount
    clearedCount = 0
    Dim report
    report = "Clearing counters for prefix: " & prefixInput & vbCrLf & vbCrLf
    
    Dim g
    For g = 0 To UBound(groupCodes)
        Dim keyName
        keyName = prefixInput & groupCodes(g)
        
        On Error Resume Next
        Dim currentValue
        currentValue = shell.RegRead(regBasePath & keyName)
        
        If Err.Number = 0 Then
            ' Key exists - delete it
            shell.RegDelete regBasePath & keyName
            If Err.Number = 0 Then
                report = report & "DELETED: " & keyName & " (was " & currentValue & ")" & vbCrLf
                clearedCount = clearedCount + 1
            Else
                report = report & "ERROR deleting: " & keyName & " - " & Err.Description & vbCrLf
            End If
        End If
        Err.Clear
    Next
    
    report = report & vbCrLf & "Cleared " & clearedCount & " registry entries."
    MsgBox report, vbInformation, "Registry Cleared"
    
Else
    ' Clear ALL counters
    Dim confirmAll
    confirmAll = MsgBox("WARNING: This will delete ALL heritage counter entries!" & vbCrLf & vbCrLf & _
                        "This affects ALL projects and prefixes." & vbCrLf & _
                        "Next run will start all groups from 1." & vbCrLf & vbCrLf & _
                        "Are you sure?", vbYesNo + vbExclamation, "Confirm Delete All")
    
    If confirmAll = vbNo Then
        MsgBox "Operation cancelled.", vbInformation
        WScript.Quit
    End If
    
    ' Delete the entire InventorRenamer key
    On Error Resume Next
    shell.RegDelete "HKCU\Software\InventorRenamer\"
    
    If Err.Number = 0 Then
        MsgBox "All registry counters have been cleared." & vbCrLf & vbCrLf & _
               "Next renaming operation will start all groups from 1.", _
               vbInformation, "Registry Cleared"
    Else
        MsgBox "Could not delete registry key: " & Err.Description & vbCrLf & vbCrLf & _
               "The key may not exist or you may not have permission.", _
               vbExclamation, "Error"
    End If
End If

WScript.Echo "Done."
