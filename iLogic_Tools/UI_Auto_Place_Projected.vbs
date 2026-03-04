Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If
Err.Clear

If invApp.ActiveDocument Is Nothing Or invApp.ActiveDocument.DocumentType <> 12292 Then
    WScript.Echo "ERROR: Active document is not a drawing. Open the target drawing and re-run."
    WScript.Quit 1
End If

Dim drawDoc, sheet, baseView, v
Set drawDoc = invApp.ActiveDocument
Set sheet = drawDoc.Sheets.Item(1)

' Find base view (ISO1 preferred)
Set baseView = Nothing
For Each v In sheet.DrawingViews
    On Error Resume Next
    If LCase(Trim(v.Name)) = "iso1" Then
        Set baseView = v
        Exit For
    End If
Next
If baseView Is Nothing Then
    If sheet.DrawingViews.Count >= 1 Then
        Set baseView = sheet.DrawingViews.Item(1)
        WScript.Echo "Info: ISO1 not found; using first view: " & baseView.Name
    Else
        WScript.Echo "ERROR: No drawing views present on the sheet to project from."
        WScript.Quit 1
    End If
Else
    WScript.Echo "Found base view: " & baseView.Name
End If

Dim initialCount
initialCount = sheet.DrawingViews.Count
WScript.Echo "Initial drawing view count: " & initialCount

' Select the base view so UI command picks it
On Error Resume Next
drawDoc.SelectSet.Clear
drawDoc.SelectSet.Select baseView
WScript.Echo "Selected base view for UI command"

' Gather candidate ControlDefinitions containing "project" or "projected"
Dim cm, ctrl, candidates
Set cm = invApp.CommandManager
Set candidates = CreateObject("System.Collections.ArrayList")
For Each ctrl In cm.ControlDefinitions
    On Error Resume Next
    Dim dname
    dname = ""
    dname = ctrl.DisplayName
    If Err.Number <> 0 Then
        Err.Clear
        dname = CStr(ctrl)
    End If
    If InStr(1, LCase(dname), "project", vbTextCompare) > 0 Then
        candidates.Add ctrl
        WScript.Echo "Candidate command found: " & dname
    End If
Next

If candidates.Count = 0 Then
    WScript.Echo "No candidate Projected view commands found; aborting UI automation"
    WScript.Quit 1
End If

Dim sh
Set sh = CreateObject("WScript.Shell")

Dim i, executed, beforeCount
executed = False
beforeCount = sheet.DrawingViews.Count

For i = 0 To candidates.Count - 1
    On Error Resume Next
    Dim c
    Set c = candidates.Item(i)
    WScript.Echo "Trying command: " & c.DisplayName
    c.Execute
    WScript.Sleep 600

    ' Bring Inventor to foreground
    On Error Resume Next
    sh.AppActivate "Autodesk Inventor"
    WScript.Sleep 300

    ' Send Enter to accept placement (if command is in place mode)
    On Error Resume Next
    sh.SendKeys "{ENTER}"
    WScript.Sleep 400
    sh.SendKeys "{ENTER}"
    WScript.Sleep 700

    ' Check if a new view was added
    If sheet.DrawingViews.Count > beforeCount Then
        WScript.Echo "Success: command created a new view (" & sheet.DrawingViews.Count - beforeCount & ")"
        executed = True
        Exit For
    Else
        WScript.Echo "No view created by this command"
    End If
Next

If Not executed Then
    WScript.Echo "UI automation could not create projected views automatically."
Else
    WScript.Echo "UI automation succeeded; updating drawing..."
    drawDoc.Update
End If

WScript.Echo "Finished UI automation test."