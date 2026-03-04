Option Explicit
On Error Resume Next
Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor not running"
    WScript.Quit 1
End If
Err.Clear

If invApp.ActiveDocument Is Nothing Then
    WScript.Echo "No active document"
    WScript.Quit 0
End If

Dim doc
Set doc = invApp.ActiveDocument
WScript.Echo "DisplayName: " & doc.DisplayName
WScript.Echo "FullFileName: '" & doc.FullFileName & "'"
WScript.Echo "DocumentType: " & doc.DocumentType
On Error Resume Next
WScript.Echo "SubType: " & doc.SubType
If Err.Number <> 0 Then
    Err.Clear
    WScript.Echo "SubType: <not available>"
End If

If doc.DocumentType = 12292 Then
    Dim sheet
    Set sheet = doc.Sheets.Item(1)
    On Error Resume Next
    WScript.Echo "Sheet count: " & doc.Sheets.Count
    WScript.Echo "Sheet Width x Height: " & sheet.Width & " x " & sheet.Height
    WScript.Echo "DrawingViews collection type: " & TypeName(sheet.DrawingViews)
    On Error Resume Next
    WScript.Echo "DrawingViews Count: " & sheet.DrawingViews.Count
    If Err.Number <> 0 Then
        Err.Clear
        WScript.Echo "DrawingViews.Count not accessible"
    End If
End If
