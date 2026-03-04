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

Dim asmDoc
Set asmDoc = Nothing
Dim d
For Each d In invApp.Documents
    If d.DocumentType = 12291 Then
        Set asmDoc = d
        Exit For
    End If
Next
If asmDoc Is Nothing Then
    WScript.Echo "ERROR: No open assembly"
    WScript.Quit 1
End If

Dim tg
Set tg = invApp.TransientGeometry
Dim pt
Set pt = tg.CreatePoint2d(sheet.Width/2, sheet.Height/2)

Dim scales
scales = Array(1/20, 1/50, 1/100, 1/200, 1/500, 1/1000)
Dim i, s
For i = 0 To UBound(scales)
    s = scales(i)
    On Error Resume Next
    WScript.Echo "Trying scale " & s
    Dim v
    Set v = sheet.DrawingViews.AddBaseView(asmDoc, pt, s, 13768, True, Nothing, "TESTASM" & i, Nothing)
    If Err.Number <> 0 Or v Is Nothing Then
        WScript.Echo "Failed at scale " & s & " : " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Success: created view at scale " & s
        Exit For
    End If
Next

If Err.Number = 0 Then drawDoc.Update
WScript.Echo "Test complete"