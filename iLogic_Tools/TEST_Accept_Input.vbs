' TEST_Accept_Input.vbs
' Try various methods to accept the current input in Inventor dialog
Option Explicit
On Error Resume Next

Dim invApp, cmdMgr, WshShell

WScript.Echo "=== ACCEPT INPUT ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set cmdMgr = invApp.CommandManager
Set WshShell = CreateObject("WScript.Shell")

' Method 1: Try to stop the current command (accept input)
WScript.Echo "Method 1: StopActiveCommand..."
cmdMgr.StopActiveCommand
If Err.Number = 0 Then
    WScript.Echo "StopActiveCommand executed"
Else
    WScript.Echo "Failed: " & Err.Description
    Err.Clear
End If

WScript.Sleep 500

' Method 2: Try accepting via built-in Accept command
WScript.Echo ""
WScript.Echo "Method 2: Looking for Accept/OK commands..."

Dim acceptCmds, cmdName, cmd
acceptCmds = Array("AcceptCmd", "OKCmd", "ApplyCmd", "FinishCmd", _
                   "DialogAcceptCmd", "DialogOKCmd", "CommandAcceptCmd")

For Each cmdName In acceptCmds
    Set cmd = cmdMgr.ControlDefinitions.Item(cmdName)
    If Not cmd Is Nothing And cmd.Enabled Then
        WScript.Echo "Found: " & cmdName & " - Executing..."
        cmd.Execute
        If Err.Number = 0 Then
            WScript.Echo "Executed!"
            Exit For
        End If
    End If
    Err.Clear
Next

' Method 3: Use Inventor window activation and keyboard
WScript.Echo ""
WScript.Echo "Method 3: SendKeys with specific timing..."

' Get the Inventor window title
Dim inventorTitle
inventorTitle = invApp.Caption
WScript.Echo "Inventor caption: " & inventorTitle

' Activate and send keys
WshShell.AppActivate inventorTitle
WScript.Sleep 200

' Try spacebar (often selects in 3D)
WScript.Echo "Sending Space..."
WshShell.SendKeys " "
WScript.Sleep 200

' Try Enter
WScript.Echo "Sending Enter..."
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200

' Try clicking in dialog OK button
WScript.Echo "Sending Tab+Enter (to reach OK button)..."
WshShell.SendKeys "{TAB}{ENTER}"

WScript.Echo ""
WScript.Echo "Check Inventor to see if the command completed."
WScript.Echo ""
WScript.Echo "=== DONE ==="
