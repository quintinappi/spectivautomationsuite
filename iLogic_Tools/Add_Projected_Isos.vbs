Option Explicit
On Error Resume Next

Const kDrawingDocumentObject = 12292
Const kIsoTopLeftViewOrientation = 13767
Const kIsoTopRightViewOrientation = 13768
Const kIsoBottomLeftViewOrientation = 13769
Const kIsoBottomRightViewOrientation = 13770

Dim invApp
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running."
    WScript.Quit 1
End If
Err.Clear

If invApp.ActiveDocument Is Nothing Or invApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
    WScript.Echo "ERROR: Active document is not a drawing. Open the target drawing and re-run."
    WScript.Quit 1
End If

Dim drawDoc
Set drawDoc = invApp.ActiveDocument
Dim sheet
Set sheet = drawDoc.Sheets.Item(1)
Dim tg
Set tg = invApp.TransientGeometry

' Find base view (prefer name ISO1)
Dim baseView, v
Set baseView = Nothing
For Each v In sheet.DrawingViews
    On Error Resume Next
    If LCase(Trim(v.Name)) = "iso1" Then
        Set baseView = v
        Exit For
    End If
Next

If baseView Is Nothing Then
    If sheet.DrawingViews.Count >= 1 Then
        Set baseView = sheet.DrawingViews.Item(1)
        WScript.Echo "Info: ISO1 not found, using first drawing view: " & baseView.Name
    Else
        WScript.Echo "ERROR: No drawing views present on the sheet to project from."
        WScript.Quit 1
    End If
Else
    WScript.Echo "Found base view: " & baseView.Name
End If

' Compute positions for projected views - corners of a 2x2 grid
Dim pos(3)
Set pos(0) = GridPoint(sheet, tg, 1, 0, 2, 2) ' top-right
Set pos(1) = GridPoint(sheet, tg, 0, 1, 2, 2) ' bottom-left
Set pos(2) = GridPoint(sheet, tg, 1, 1, 2, 2) ' bottom-right

Dim pv
Dim names
names = Array("ISO2","ISO3","ISO4")
Dim i
For i = 0 To 2
    Set pv = Nothing
    On Error Resume Next
    ' Preferred method: AddProjectedView(baseView, Point2d, Orientation, True, Name)
    Set pv = sheet.DrawingViews.AddProjectedView(baseView, pos(i), 0, True, names(i))
    If Err.Number <> 0 Or pv Is Nothing Then
        Err.Clear
        ' Try alternate signature: without name
        On Error Resume Next
        sheet.DrawingViews.AddProjectedView baseView, pos(i)
        If Err.Number = 0 Then
            Set pv = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
        Else
            Err.Clear
            ' Try passing orientation explicitly as constant
            On Error Resume Next
            sheet.DrawingViews.AddProjectedView baseView, pos(i), kIsoTopLeftViewOrientation
            If Err.Number = 0 Then
                Set pv = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
            End If
        End If
    End If

    If Not pv Is Nothing Then
        WScript.Echo "Added projected view: " & names(i) & " (object: " & TypeName(pv) & ")"
        ' Attempt to rename if possible
        On Error Resume Next
        pv.Name = names(i)
    Else
        WScript.Echo "Warning: failed to add projected view: " & names(i)
    End If
Next

' Add parts list anchored to base view
On Error Resume Next
Dim plPt
Set plPt = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)
sheet.PartsLists.Add baseView, plPt
If Err.Number = 0 Then
    WScript.Echo "Parts list added and anchored to base view"
Else
    WScript.Echo "Warning: could not add parts list - " & Err.Description
    Err.Clear
End If

' Update drawing
drawDoc.Update
WScript.Echo "Done. Check drawing for added projected isos and parts list." 

Function GridPoint(sheet, tg, col, row, cols, rows)
    Dim margin, cellW, cellH, x, y
    margin = 10
    cellW = (sheet.Width - (2 * margin)) / cols
    cellH = (sheet.Height - (2 * margin)) / rows
    x = margin + (col + 0.5) * cellW
    y = margin + (row + 0.5) * cellH
    Set GridPoint = tg.CreatePoint2d(x, y)
End Function