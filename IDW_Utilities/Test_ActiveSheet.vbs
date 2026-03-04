' TEST: Use ActiveSheet property

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_activesheet.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: ACTIVE SHEET ==="

Set invApp = GetObject(, "Inventor.Application")
Set invDoc = invApp.ActiveDocument
Log "Drawing: " & invDoc.DisplayName

' Get part first
Dim pd, found, d, i
found = False
For i = 1 To invApp.Documents.Count
    Set d = invApp.Documents.Item(i)
    If d.DocumentType = 12290 Then
        Set pd = d
        found = True
        Log "Part: " & pd.DisplayName
        Exit For
    End If
Next

If Not found Then
    Log "ERROR: No part"
    WScript.Quit
End If

' Switch to Sheet 2 and make it active
invDoc.Sheets.Item(2).Activate
Log "Activated Sheet 2"

' Now use ActiveSheet
Dim activeSh
Set activeSh = invDoc.ActiveSheet
Log "ActiveSheet: " & activeSh.Name

' Create point
Dim pt
Set pt = invApp.TransientGeometry.CreatePoint2d(10, 20)

' Count before
Dim vb4
vb4 = activeSh.DrawingViews.Count
Log "Views before: " & vb4

' Try AddBaseView on ActiveSheet
Dim bv
Log "Calling AddBaseView on ActiveSheet..."
Set bv = activeSh.DrawingViews.AddBaseView(pd, pt, 1)

Dim va4
va4 = activeSh.DrawingViews.Count
Log "Views after: " & va4

If Err.Number <> 0 Then
    Log "ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear
ElseIf bv Is Nothing Then
    Log "ERROR: bv is Nothing"
Else
    Log "SUCCESS: " & bv.Name
End If

invDoc.Save2 True
logFile.Close
MsgBox "Done", vbInformation, "Test"
