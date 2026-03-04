Option Explicit
On Error Resume Next

Const kAssemblyDocumentObject = 12291
Const kDrawingDocumentObject = 12292
Const kFrontViewOrientation = 13761
Const kLeftViewOrientation = 13763
Const kRightViewOrientation = 13764
Const kTopViewOrientation = 13765
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

Dim drawDoc
Set drawDoc = Nothing
If Not invApp.ActiveDocument Is Nothing Then
    If invApp.ActiveDocument.DocumentType = kDrawingDocumentObject Then
        Set drawDoc = invApp.ActiveDocument
    End If
End If

If drawDoc Is Nothing Then
    WScript.Echo "ERROR: No active drawing. Open the target drawing and re-run."
    WScript.Quit 1
End If

' Find an open assembly to reference for views
Dim docs, d
Set docs = invApp.Documents
Dim asmDoc
Set asmDoc = Nothing
For Each d In docs
    If d.DocumentType = kAssemblyDocumentObject Then
        Set asmDoc = d
        Exit For
    End If
Next

If asmDoc Is Nothing Then
    WScript.Echo "ERROR: No open assembly found to create views from. Open the assembly and re-run."
    WScript.Quit 1
End If

Dim sheet
Set sheet = drawDoc.Sheets.Item(1)
If sheet Is Nothing Then
    WScript.Echo "ERROR: Drawing has no sheets."
    WScript.Quit 1
End If

Dim tg
Set tg = invApp.TransientGeometry

Dim scale
scale = 1 / 20 ' 1:20

' Build 2x2 grid on the sheet
Dim positions(3)
Set positions(0) = GridPoint(sheet, tg, 0, 0, 2, 2)
Set positions(1) = GridPoint(sheet, tg, 1, 0, 2, 2)
Set positions(2) = GridPoint(sheet, tg, 0, 1, 2, 2)
Set positions(3) = GridPoint(sheet, tg, 1, 1, 2, 2)

WScript.Echo "DEBUG: Sheet Width x Height = " & sheet.Width & " x " & sheet.Height
WScript.Echo "DEBUG: Position(0) type: " & TypeName(positions(0))

Dim views(3)
Dim i
For i = 0 To 3
    ' Choose orientations for a varied look
    Dim orient
    Select Case i
        Case 0
            orient = kIsoTopRightViewOrientation
        Case 1
            orient = kIsoTopLeftViewOrientation
        Case 2
            orient = kIsoBottomRightViewOrientation
        Case 3
            orient = kIsoBottomLeftViewOrientation
    End Select

    On Error Resume Next
    Dim v
    Set v = Nothing

    ' Try primary form (returns object)
    Set v = sheet.DrawingViews.AddBaseView(asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing)
    If Err.Number <> 0 Or v Is Nothing Then
        Err.Clear
        ' Try calling as sub-style and grab last-added view
        On Error Resume Next
        sheet.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing
        If Err.Number = 0 Then
            Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
        Else
            Err.Clear
            ' Try simplified call (fewer arguments)
            sheet.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient, True
            If Err.Number = 0 Then
                Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
            Else
                Err.Clear
                ' Try with just four arguments
                sheet.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient
                If Err.Number = 0 Then
                    Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
                Else
                    Err.Clear
                    ' Try passing filename string instead of Document object
                    On Error Resume Next
                    sheet.DrawingViews.AddBaseView asmDoc.FullFileName, positions(i), scale, orient, True
                    If Err.Number = 0 Then
                        Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
                    Else
                        Err.Clear
                        ' Try a part document from the assembly occurrences
                        Dim occDoc
                        Set occDoc = Nothing
                        On Error Resume Next
                        If asmDoc.ComponentDefinition.Occurrences.Count > 0 Then
                            occDoc = asmDoc.ComponentDefinition.Occurrences.Item(1).Definition.Document
                        End If
                        If Not occDoc Is Nothing Then
                            sheet.DrawingViews.AddBaseView occDoc, positions(i), scale, orient, True
                            If Err.Number = 0 Then
                                Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
                            Else
                                WScript.Echo "Warning: failed to add ISO" & (i+1) & " (all attempts) - " & Err.Description
                                Err.Clear
                            End If
                        Else
                            WScript.Echo "Warning: failed to add ISO" & (i+1) & " (all attempts) - " & Err.Description
                            Err.Clear
                        End If
                    End If
                End If
            End If
        End If
    Else
        WScript.Echo "Added ISO" & (i+1) & " at scale 1:20"
    End If

    Set views(i) = v
    If Not v Is Nothing Then
        WScript.Echo "INFO: ISO" & (i+1) & " object obtained: " & TypeName(v)
    End If
Next

' Add a parts list anchored to the first iso view if we have it
If Not IsNull(views(0)) And Not views(0) Is Nothing Then
    On Error Resume Next
    Dim plPoint
    Set plPoint = tg.CreatePoint2d(CDbl(sheet.Width) - 40, CDbl(sheet.Height) - 40)
    sheet.PartsLists.Add views(0), plPoint
    If Err.Number = 0 Then
        WScript.Echo "Parts list added and anchored to ISO1"
    Else
        WScript.Echo "Warning: could not add parts list - " & Err.Description
        Err.Clear
    End If
Else
    WScript.Echo "INFO: No valid ISO view to anchor parts list"
End If

' Update drawing
On Error Resume Next
drawDoc.Update
If Err.Number = 0 Then
    WScript.Echo "Drawing updated. Done."
Else
    WScript.Echo "Warning: drawing update returned error: " & Err.Description
    Err.Clear
End If

Function GridPoint(sheet, tg, col, row, cols, rows)
    Dim margin
    margin = 10 ' units consistent with sheet.Width/Height (tweak if needed)
    Dim cellW, cellH
    cellW = (sheet.Width - (2 * margin)) / cols
    cellH = (sheet.Height - (2 * margin)) / rows

    Dim x, y
    x = margin + (col + 0.5) * cellW
    y = margin + (row + 0.5) * cellH
    Set GridPoint = tg.CreatePoint2d(x, y)
End Function
