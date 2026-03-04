Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If
Err.Clear

Dim drawDoc, sheet
If Not invApp.ActiveDocument Is Nothing And invApp.ActiveDocument.DocumentType = 12292 Then
    Set drawDoc = invApp.ActiveDocument
    Set sheet = drawDoc.Sheets.Item(1)
Else
    WScript.Echo "ERROR: No active drawing"
    WScript.Quit 1
End If

Dim partDoc
Set partDoc = Nothing
Dim d
For Each d In invApp.Documents
    If d.DocumentType = 12290 Then
        Set partDoc = d
        Exit For
    End If
Next

If partDoc Is Nothing Then
    WScript.Echo "ERROR: No open part document found to test"
    WScript.Quit 1
End If

Dim tg
Set tg = invApp.TransientGeometry
Dim pt
Set pt = tg.CreatePoint2d(sheet.Width/2, sheet.Height/2)

On Error Resume Next
Dim v
Set v = sheet.DrawingViews.AddBaseView(partDoc, pt, 1/5, 13768, True, Nothing, "TESTPARTISO", Nothing)
If Err.Number <> 0 Or v Is Nothing Then
    WScript.Echo "Failed to add base view from part: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Successfully added base view from part: " & v.Name
End If

sheet.DrawingViews.Item(sheet.DrawingViews.Count).Scale = 1/5

drawDoc.Update
WScript.Echo "Test complete"