' TEST_Click_Face.vbs
' Simulate mouse click to confirm face selection
Option Explicit
On Error Resume Next

Dim invApp, partDoc, compDef, selectSet, WshShell

WScript.Echo "=== CLICK TO CONFIRM FACE ==="
WScript.Echo ""

Set invApp = GetObject(, "Inventor.Application")
Set WshShell = CreateObject("WScript.Shell")
Set partDoc = invApp.ActiveDocument
Set compDef = partDoc.ComponentDefinition
Set selectSet = partDoc.SelectSet

' Find and select largest face
Dim body, faces, face, largestFace, largestArea
largestArea = 0

Set body = compDef.SurfaceBodies.Item(1)
Set faces = body.Faces

For Each face In faces
    Dim area
    area = face.Evaluator.Area * 100
    If Err.Number <> 0 Then
        area = 0
        Err.Clear
    End If
    
    If area > largestArea Then
        largestArea = area
        Set largestFace = face
    End If
Next

WScript.Echo "Largest face: " & FormatNumber(largestArea, 0) & " mm²"

' Select the face
selectSet.Clear
selectSet.Select largestFace
WScript.Echo "Face selected: " & selectSet.Count

' Activate Inventor
WshShell.AppActivate "Autodesk Inventor"
WScript.Sleep 300

' Try different key combinations to accept
WScript.Echo "Trying Right-Click to accept..."
WshShell.SendKeys "+{F10}"  ' Shift+F10 = Right-click context menu
WScript.Sleep 200
WshShell.SendKeys "{ENTER}" ' Select first option (usually Accept/OK)
WScript.Sleep 300

' Also try double-click simulation via Space
WScript.Echo "Trying Space (like clicking selected item)..."
WshShell.SendKeys " "
WScript.Sleep 300

' Try Escape then re-approach
WScript.Echo "Trying mouse click simulation..."

' Use mouse_event via PowerShell to click
Dim psCmd
psCmd = "powershell -Command ""Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Mouse { [DllImport(""""user32.dll"""")] public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo); public static void Click() { mouse_event(2, 0, 0, 0, 0); mouse_event(4, 0, 0, 0, 0); } }'; [Mouse]::Click()"""

WshShell.Run psCmd, 0, True
WScript.Sleep 500

WScript.Echo ""
WScript.Echo "Check Inventor - the face should be confirmed now."
WScript.Echo ""
WScript.Echo "=== DONE ==="
