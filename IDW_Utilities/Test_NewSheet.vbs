' TEST: Create a BRAND NEW sheet and place a view there

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_newsheet.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: NEW SHEET ==="

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Log "ERROR: " & Err.Description
    WScript.Quit
End If

Set invDoc = invApp.ActiveDocument
If invDoc Is Nothing Then
    Log "ERROR: No doc"
    WScript.Quit
End If

Log "Drawing: " & invDoc.DisplayName
Log "Current sheets: " & invDoc.Sheets.Count

' Create a brand new sheet
Dim newSheet, sheetNum
sheetNum = invDoc.Sheets.Count + 1
Set newSheet = invDoc.Sheets.Add()
newSheet.Name = "TestSheet" & sheetNum
Log "Created new sheet: " & newSheet.Name

' Activate it
newSheet.Activate
Log "Activated. Views: " & newSheet.DrawingViews.Count

' Get a part
Dim pd, d, i
For i = 1 To invApp.Documents.Count
    Set d = invApp.Documents.Item(i)
    If d.DocumentType = 12290 Then
        Set pd = d
        Log "Part: " & pd.DisplayName
        Exit For
    End If
Next

If pd Is Nothing Then
    Log "ERROR: No part"
    WScript.Quit
End If

' Make sure we're on the drawing
invDoc.Activate

' Create point
Dim pt
Set pt = invApp.TransientGeometry.CreatePoint2d(5, 5)

' Try AddBaseView
Dim bv, vb4, va4
vb4 = newSheet.DrawingViews.Count
Log "Views before: " & vb4

Set bv = newSheet.DrawingViews.AddBaseView(pd, pt, 1)

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

' Clean up - delete test sheet
On Error Resume Next
newSheet.Delete
If Err.Number <> 0 Then
    Log "Could not delete test sheet: " & Err.Description
    Err.Clear
Else
    Log "Test sheet deleted"
End If

invDoc.Save2 True
Log "Saved"

logFile.Close
MsgBox "Done! Check: " & logPath, vbInformation, "Test"
