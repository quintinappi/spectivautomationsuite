' TEST: Place ONE part on Sheet 2 (which already exists and has views)

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_one_sheet2.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: ONE PART ON SHEET 2 ==="

Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then
    Log "ERROR: Cannot connect"
    WScript.Quit
End If

Set invDoc = invApp.ActiveDocument
Log "Drawing: " & invDoc.DisplayName

' Get Sheet 2 (existing)
Dim sh
Set sh = invDoc.Sheets.Item(2)
Log "Using: " & sh.Name & " (has " & sh.DrawingViews.Count & " views)"

' Get first available part
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

' CRITICAL: Activate the drawing document first
invDoc.Activate
Log "Drawing activated"

' Activate the sheet
sh.Activate
Log "Sheet activated"

' Wait
WScript.Sleep 200

' Position - try center of sheet area
Dim x, y
x = 10
y = 20
Dim pt
Set pt = invApp.TransientGeometry.CreatePoint2d(x, y)
Log "Position: (" & x & ", " & y & ")"

' Count before
Dim vb4, va4
vb4 = sh.DrawingViews.Count
Log "Views before: " & vb4

' THE CALL - exactly as in original working script
Dim bv
Set bv = sh.DrawingViews.AddBaseView(pd, pt, 1)

va4 = sh.DrawingViews.Count
Log "Views after: " & va4

If Err.Number <> 0 Then
    Log "ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear
ElseIf bv Is Nothing Then
    Log "ERROR: bv is Nothing"
Else
    Log "SUCCESS! View: " & bv.Name
    Log "Scale: " & bv.Scale
End If

invDoc.Save2 True
Log "Saved"

logFile.Close
MsgBox "Done! Check: " & logPath, vbInformation, "Test"
