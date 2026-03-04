' TEST: Direct AddBaseView with minimal code
' This mimics exactly what the user does manually

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_direct.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== DIRECT AddBaseView TEST ==="

' Connect
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Log "ERROR: Cannot connect: " & Err.Description
    WScript.Quit
End If
Log "Connected"

' Get document
Set invDoc = invApp.ActiveDocument
If invDoc Is Nothing Then
    Log "ERROR: No doc"
    WScript.Quit
End If
Log "Doc: " & invDoc.DisplayName & " (Type: " & invDoc.DocumentType & ")"

' Check it's a drawing
If invDoc.DocumentType <> 12294 And invDoc.DocumentType <> 12292 Then
    Log "ERROR: Not a drawing"
    WScript.Quit
End If

' Get or create Sheet 2
Dim sh
If invDoc.Sheets.Count < 2 Then
    Set sh = invDoc.Sheets.Add()
    sh.Name = "Sheet:2"
    Log "Created Sheet 2"
Else
    Set sh = invDoc.Sheets.Item(2)
    Log "Got Sheet 2: " & sh.Name
End If

' Activate it
sh.Activate
Log "Activated. Views on sheet: " & sh.DrawingViews.Count

' Find ANY part document
Dim pd, found, d, idx
found = False
For idx = 1 To invApp.Documents.Count
    Set d = invApp.Documents.Item(idx)
    Log "Doc " & idx & ": " & d.DisplayName & " (Type: " & d.DocumentType & ")"
    If d.DocumentType = 12290 Then ' Part
        Set pd = d
        found = True
        Log "Using part: " & pd.DisplayName
        Exit For
    End If
Next

If Not found Then
    Log "ERROR: No part found"
    WScript.Quit
End If

' Make sure drawing is active document
invDoc.Activate
Log "Drawing activated"

' Wait a moment
WScript.Sleep 100

' Create point
Dim pt
Set pt = invApp.TransientGeometry.CreatePoint2d(5, 5)
Log "Point created: (5, 5)"

' Count before
Dim vb4, va4
vb4 = sh.DrawingViews.Count
Log "Views before AddBaseView: " & vb4

' THE CALL - simplest possible
Dim bv
Set bv = Nothing
Log "About to call AddBaseView..."
Set bv = sh.DrawingViews.AddBaseView(pd, pt, 1)

' Check immediately
If Err.Number <> 0 Then
    Log "ERROR from AddBaseView: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    Log "No error from AddBaseView"
End If

' Count after
va4 = sh.DrawingViews.Count
Log "Views after: " & va4

' Check result
If bv Is Nothing Then
    Log "bv is Nothing"
Else
    Log "bv.Name: " & bv.Name
    Log "bv.Scale: " & bv.Scale
End If

' Save
On Error Resume Next
invDoc.Save2 True
If Err.Number <> 0 Then
    Log "Save error: " & Err.Description
    Err.Clear
Else
    Log "Saved OK"
End If

logFile.Close
MsgBox "Done! Check: " & logPath, vbInformation, "Test Complete"
