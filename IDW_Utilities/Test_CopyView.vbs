' TEST: Copy an existing view instead of AddBaseView
' CopyTo might work if AddBaseView is broken

Option Explicit

Dim invApp, invDoc, fso, logFile, logPath

Set fso = CreateObject("Scripting.FileSystemObject")
logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\test_copyview.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    WScript.Echo msg
    If Not logFile Is Nothing Then logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== TEST: COPY VIEW (not AddBaseView) ==="

Set invApp = GetObject(, "Inventor.Application")
Set invDoc = invApp.ActiveDocument
Log "Drawing: " & invDoc.DisplayName

' Get Sheet 1 (source with views)
Dim sourceSheet
Set sourceSheet = invDoc.Sheets.Item(1)
Log "Sheet 1 has " & sourceSheet.DrawingViews.Count & " views"

If sourceSheet.DrawingViews.Count = 0 Then
    Log "ERROR: Sheet 1 has no views to copy"
    WScript.Quit
End If

' Get Sheet 2 (target)
Dim targetSheet
Set targetSheet = invDoc.Sheets.Item(2)
Log "Sheet 2 has " & targetSheet.DrawingViews.Count & " views"

' Get first view from Sheet 1
Dim sourceView
Set sourceView = sourceSheet.DrawingViews.Item(1)
Log "Source view: " & sourceView.Name

' Activate target
targetSheet.Activate

' Count before
Dim vb4, va4
vb4 = targetSheet.DrawingViews.Count
Log "Views on Sheet 2 before: " & vb4

' Try CopyTo
Dim newView
Log "Calling CopyTo..."
Set newView = sourceView.CopyTo(targetSheet)

va4 = targetSheet.DrawingViews.Count
Log "Views on Sheet 2 after: " & va4

If Err.Number <> 0 Then
    Log "ERROR: " & Err.Number & " - " & Err.Description
    Err.Clear
ElseIf newView Is Nothing Then
    Log "ERROR: newView is Nothing"
Else
    Log "SUCCESS! Copied view: " & newView.Name
    Log "Original scale: " & sourceView.Scale
    Log "New scale: " & newView.Scale
    
    ' Try to change scale
    On Error Resume Next
    newView.Scale = 0.1
    If Err.Number <> 0 Then
        Log "ERROR setting scale: " & Err.Description
        Err.Clear
    Else
        Log "Scale changed to: " & newView.Scale
    End If
End If

invDoc.Save2 True
logFile.Close
MsgBox "Done! Check: " & logPath, vbInformation, "Test"
