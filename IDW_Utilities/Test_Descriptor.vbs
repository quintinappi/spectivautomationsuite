' TEST: Use DocumentDescriptor instead of Document

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_descriptor.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: DOCUMENT DESCRIPTOR ==="

Set invApp = GetObject(, "Inventor.Application")
Set invDoc = invApp.ActiveDocument
Log "Drawing: " & invDoc.DisplayName

' Get Sheet 2
Dim sh
Set sh = invDoc.Sheets.Item(2)
Log "Sheet 2 has " & sh.DrawingViews.Count & " views"

' Get a part and its descriptor
Dim pd, pdDesc, found, d, i
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

' Try using the descriptor
On Error Resume Next
Set pdDesc = pd.FullDocumentName
If Err.Number <> 0 Then
    Log "ERROR getting descriptor: " & Err.Description
    Err.Clear
Else
    Log "Descriptor: " & pdDesc
End If

' Activate everything
invDoc.Activate
sh.Activate

' Create position
Dim pt
Set pt = invApp.TransientGeometry.CreatePoint2d(10, 20)

' Count before
Dim vb4
vb4 = sh.DrawingViews.Count
Log "Views before: " & vb4

' Try AddBaseView with descriptor
Dim bv
Log "Trying AddBaseView..."
Set bv = sh.DrawingViews.AddBaseView(pd, pt, 1)

Dim va4
va4 = sh.DrawingViews.Count
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
