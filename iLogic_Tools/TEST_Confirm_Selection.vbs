' TEST_Confirm_Selection.vbs
' Confirm the selection by sending Enter key
Option Explicit
On Error Resume Next

Dim invApp, WshShell

WScript.Echo "=== CONFIRM SELECTION ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set WshShell = CreateObject("WScript.Shell")

' Make sure Inventor is in focus
WScript.Echo "Activating Inventor..."
WshShell.AppActivate "Autodesk Inventor"
WScript.Sleep 500

' Send Enter key to confirm
WScript.Echo "Sending Enter key..."
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500

' Also try right-click to accept (sometimes needed)
WScript.Echo "Sending Right-click (context menu)..."
WshShell.SendKeys "+{F10}"
WScript.Sleep 300
WshShell.SendKeys "{ENTER}"

WScript.Echo ""
WScript.Echo "Keys sent! Check Inventor."
WScript.Echo ""
WScript.Echo "=== DONE ==="
