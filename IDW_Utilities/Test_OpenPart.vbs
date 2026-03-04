' TEST: Open part fresh and place view

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_openpart.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: OPEN PART FRESH ==="

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Log "ERROR: " & Err.Description
    WScript.Quit
End If

Set invDoc = invApp.ActiveDocument
Log "Drawing: " & invDoc.DisplayName

' Create new sheet
Dim newSheet
Set newSheet = invDoc.Sheets.Add()
newSheet.Name = "TestOpen"
Log "Created sheet: " & newSheet.Name
newSheet.Activate

' Find a part file path
Dim partPath, d, i
partPath = ""
For i = 1 To invApp.Documents.Count
    Set d = invApp.Documents.Item(i)
    If d.DocumentType = 12290 Then
        partPath = d.FullFileName
        Log "Found part at: " & partPath
        Exit For
    End If
Next

If partPath = "" Then
    Log "ERROR: No part found"
    WScript.Quit
End If

' Close the part if it's already open
On Error Resume Next
invApp.Documents.Item(partPath).Close
If Err.Number <> 0 Then
    Err.Clear
End If

' Now open the part FRESH (hidden)
Dim freshPart
Log "Opening part fresh..."
Set freshPart = invApp.Documents.Open(partPath, False) ' False = hidden

If Err.Number <> 0 Then
    Log "ERROR opening part: " & Err.Description
    Err.Clear
    WScript.Quit
End If

Log "Part opened: " & freshPart.DisplayName

' Make sure drawing is active
invDoc.Activate
newSheet.Activate

' Try AddBaseView
Dim pt, bv, vb4, va4
Set pt = invApp.TransientGeometry.CreatePoint2d(5, 5)

vb4 = newSheet.DrawingViews.Count
Log "Views before: " & vb4

Set bv = newSheet.DrawingViews.AddBaseView(freshPart, pt, 1)

va4 = newSheet.DrawingViews.Count
Log "Views after: " & va4

If Err.Number <> 0 Then
    Log "ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear
ElseIf bv Is Nothing Then
    Log "bv is Nothing"
Else
    Log "SUCCESS! View: " & bv.Name
End If

' Clean up
freshPart.Close
newSheet.Delete
invDoc.Save2 True

logFile.Close
MsgBox "Done!", vbInformation, "Test"
