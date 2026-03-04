' TEST SCRIPT - Simple view placement
' This tests if AddBaseView works at all

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

' Main
On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_simple.log"
Set logFile = fso.CreateTextFile(logPath, True)

LogMsg "=== TEST SIMPLE VIEW PLACEMENT ==="

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    LogMsg "ERROR: Cannot connect to Inventor"
    WScript.Quit
End If

Set invDoc = invApp.ActiveDocument
If invDoc Is Nothing Then
    LogMsg "ERROR: No active document"
    WScript.Quit
End If

LogMsg "Drawing: " & invDoc.DisplayName
LogMsg "Sheet count: " & invDoc.Sheets.Count

' Get sheet 2 (or create it)
Dim targetSheet
If invDoc.Sheets.Count < 2 Then
    Set targetSheet = invDoc.Sheets.Add()
    targetSheet.Name = "Sheet:2"
    LogMsg "Created Sheet 2"
Else
    Set targetSheet = invDoc.Sheets.Item(2)
    LogMsg "Using Sheet 2"
End If

' Get a part document - try to find one already open
Dim partDoc, i, foundPart
foundPart = False
For i = 1 To invApp.Documents.Count
    Dim doc
    Set doc = invApp.Documents.Item(i)
    LogMsg "Checking doc: " & doc.DisplayName & " (Type: " & doc.DocumentType & ")"
    If doc.DocumentType = 12290 Then ' kPartDocumentObject
        Set partDoc = doc
        foundPart = True
        LogMsg "Found part: " & partDoc.DisplayName
        Exit For
    End If
Next

If Not foundPart Then
    LogMsg "ERROR: No part document found open"
    WScript.Quit
End If

' Activate the sheet
targetSheet.Activate
LogMsg "Activated sheet: " & invDoc.ActiveSheet.Name

' Create position
Dim x, y
x = 10
y = 20
Dim position
Set position = invApp.TransientGeometry.CreatePoint2d(x, y)
LogMsg "Position created: (" & x & ", " & y & ")"

' Count before
Dim viewsBefore
viewsBefore = targetSheet.DrawingViews.Count
LogMsg "Views before: " & viewsBefore

' Try AddBaseView with literal 1
Err.Clear
Dim baseView
LogMsg "Calling AddBaseView with scale = 1..."
Set baseView = targetSheet.DrawingViews.AddBaseView(partDoc, position, 1)

' Check result
Dim viewsAfter
viewsAfter = targetSheet.DrawingViews.Count
LogMsg "Views after: " & viewsAfter

If Err.Number <> 0 Then
    LogMsg "ERROR: " & Err.Description
    Err.Clear
ElseIf baseView Is Nothing Then
    LogMsg "ERROR: baseView is Nothing"
ElseIf viewsAfter <= viewsBefore Then
    LogMsg "ERROR: View count didn't increase (before=" & viewsBefore & ", after=" & viewsAfter & ")"
Else
    LogMsg "SUCCESS: View created! Name=" & baseView.Name
End If

' Save
invDoc.Save2 True
LogMsg "Saved"

logFile.Close
MsgBox "Test complete. Check log: " & logPath, vbInformation, "Done"

Sub LogMsg(msg)
    Dim timestamp
    timestamp = Now
    If Not logFile Is Nothing Then
        logFile.WriteLine timestamp & " - " & msg
    End If
End Sub
