' Test to find the correct Document Settings command
Option Explicit

Dim inventorApp
Set inventorApp = GetObject(, "Inventor.Application")

If inventorApp Is Nothing Then
    MsgBox "Inventor not running"
    WScript.Quit 1
End If

' Try to find Document Settings command
Dim cmdMgr, ctrlDefs
Set cmdMgr = inventorApp.CommandManager
Set ctrlDefs = cmdMgr.ControlDefinitions

Dim commandsToTry
commandsToTry = Array( _
    "AppDocSettingsCmd", _
    "PartDocSettingsCmd", _
    "AppDocumentSettingsCmd", _
    "PartDocumentSettingsCmd", _
    "DocSettingsCmd", _
    "AppOptionsCmd", _
    "ToolsOptionsCmd" _
)

Dim result
result = "Available Document Settings Commands:" & vbCrLf & vbCrLf

Dim i, cmd, cmdObj
For i = 0 To UBound(commandsToTry)
    cmd = commandsToTry(i)
    On Error Resume Next
    Set cmdObj = ctrlDefs.Item(cmd)
    If Err.Number = 0 And Not cmdObj Is Nothing Then
        result = result & "✓ " & cmd & " - FOUND" & vbCrLf
    Else
        result = result & "✗ " & cmd & " - NOT FOUND" & vbCrLf
    End If
    Err.Clear
Next

MsgBox result, vbInformation, "Command Test"

' Now try to execute the first one found
For i = 0 To UBound(commandsToTry)
    cmd = commandsToTry(i)
    On Error Resume Next
    Set cmdObj = ctrlDefs.Item(cmd)
    If Err.Number = 0 And Not cmdObj Is Nothing Then
        MsgBox "Executing: " & cmd, vbInformation
        cmdObj.Execute
        WScript.Sleep 2000
        Exit For
    End If
    Err.Clear
Next
