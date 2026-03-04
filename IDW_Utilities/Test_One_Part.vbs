' TEST: Place just ONE part, no scale worries
' Uses scale 1:1 which we know works

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

' Setup log
Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_one_part.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: PLACE ONE PART ==="

' Connect to Inventor
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Log "ERROR: Cannot connect to Inventor"
    WScript.Quit
End If
Log "Connected to Inventor"

' Get active document
Set invDoc = invApp.ActiveDocument
If invDoc Is Nothing Then
    Log "ERROR: No active document"
    WScript.Quit
End If
Log "Document: " & invDoc.DisplayName

' Get Sheet 2
Dim targetSheet
If invDoc.Sheets.Count < 2 Then
    Set targetSheet = invDoc.Sheets.Add()
    targetSheet.Name = "Sheet:2"
    Log "Created Sheet 2"
Else
    Set targetSheet = invDoc.Sheets.Item(2)
    Log "Using Sheet 2"
End If

' Find first PART (not assembly)
Dim partDoc, i, doc, found
found = False
For i = 1 To invApp.Documents.Count
    Set doc = invApp.Documents.Item(i)
    If doc.DocumentType = 12290 Then ' kPartDocumentObject
        Set partDoc = doc
        found = True
        Log "Found part: " & doc.DisplayName
        Exit For
    End If
Next

If Not found Then
    Log "ERROR: No part file found open"
    WScript.Quit
End If

' Activate sheet
targetSheet.Activate
Log "Activated Sheet 2"

' Count views before
Dim viewsBefore
viewsBefore = targetSheet.DrawingViews.Count
Log "Views before: " & viewsBefore

' Create position at center of sheet
Dim position
Set position = invApp.TransientGeometry.CreatePoint2d(10, 15)
Log "Position: (10, 15)"

' Create view with scale 1 (KNOWN TO WORK)
Dim baseView
Log "Calling AddBaseView(partDoc, position, 1)..."
Set baseView = targetSheet.DrawingViews.AddBaseView(partDoc, position, 1)

' Check result
Dim viewsAfter
viewsAfter = targetSheet.DrawingViews.Count
Log "Views after: " & viewsAfter

If Err.Number <> 0 Then
    Log "ERROR: " & Err.Description
    Err.Clear
ElseIf baseView Is Nothing Then
    Log "ERROR: baseView is Nothing"
ElseIf viewsAfter <= viewsBefore Then
    Log "ERROR: View count didn't increase"
Else
    Log "SUCCESS! View created: " & baseView.Name
    Log "Scale: " & baseView.Scale
End If

' Save
invDoc.Save2 True
Log "Saved"

logFile.Close
MsgBox "Test complete. Check: " & logPath, vbInformation, "Done"
