Option Explicit
On Error Resume Next

Const kAssemblyDocumentObject = 12291
Const kDrawingDocumentObject = 12292
Const kIsoTopLeftViewOrientation = 13767
Const kIsoTopRightViewOrientation = 13768
Const kIsoBottomLeftViewOrientation = 13769
Const kIsoBottomRightViewOrientation = 13770

Dim invApp, fso
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running or couldn't be reached."
    WScript.Quit 1
End If
Err.Clear
Set fso = CreateObject("Scripting.FileSystemObject")

' Find the first open assembly
Dim docs, d, asmDoc
Set docs = invApp.Documents
Set asmDoc = Nothing
For Each d In docs
    If d.DocumentType = kAssemblyDocumentObject Then
        Set asmDoc = d
        Exit For
    End If
Next

If asmDoc Is Nothing Then
    WScript.Echo "ERROR: No open assembly found. Open the assembly and re-run."
    WScript.Quit 1
End If

WScript.Echo "Using assembly: " & asmDoc.DisplayName

' Resolve template
Dim templatePath
templatePath = ResolveTemplatePath()
If templatePath = "" Then
    WScript.Echo "ERROR: Could not find a suitable template."
    WScript.Quit 1
End If
WScript.Echo "Template: " & templatePath

' Create drawing from template
Dim drawDoc
Set drawDoc = invApp.Documents.Add(kDrawingDocumentObject, templatePath, True)
If drawDoc Is Nothing Then
    WScript.Echo "ERROR: Could not create drawing from template."
    WScript.Quit 1
End If

' Determine save path next to assembly
Dim asmFolder, asmBase, savePath, counter
asmFolder = Left(asmDoc.FullFileName, InStrRev(asmDoc.FullFileName, "\"))
asmBase = Left(asmDoc.DisplayName, InStrRev(asmDoc.DisplayName, ".") - 1)
savePath = asmFolder & asmBase & "-DETAIL.idw"
counter = 1
Do While fso.FileExists(savePath)
    savePath = asmFolder & asmBase & "-DETAIL(" & counter & ").idw"
    counter = counter + 1
Loop

On Error Resume Next
drawDoc.SaveAs savePath, True
If Err.Number <> 0 Then
    WScript.Echo "Warning: Could not save drawing to " & savePath & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "Drawing saved to: " & savePath
End If

' Place four isometric base views at 1:20
Dim sheet
Set sheet = drawDoc.Sheets.Item(1)
sheet.Activate
' Give Inventor a moment to initialize the new drawing
WScript.Sleep 500
Dim tg
Set tg = invApp.TransientGeometry
Dim scale
scale = 1 / 20

Dim positions(3)
Set positions(0) = GridPoint(sheet, tg, 0, 0, 2, 2)
Set positions(1) = GridPoint(sheet, tg, 1, 0, 2, 2)
Set positions(2) = GridPoint(sheet, tg, 0, 1, 2, 2)
Set positions(3) = GridPoint(sheet, tg, 1, 1, 2, 2)

Dim views(3)
Dim i, orient
For i = 0 To 3
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

    ' Attempt 1: Try sheet.DrawingViews first
    On Error Resume Next
    Set v = sheet.DrawingViews.AddBaseView(asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing)
    If Err.Number = 0 And Not v Is Nothing Then
        WScript.Echo "Method Succeeded (sheet.DrawingViews) for ISO" & (i+1)
    Else
        Err.Clear
        ' Attempt 2: drawDoc.DrawingViews full form
        On Error Resume Next
        Set v = drawDoc.DrawingViews.AddBaseView(asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing)
        If Err.Number = 0 And Not v Is Nothing Then
            WScript.Echo "Method Succeeded (drawDoc.DrawingViews full) for ISO" & (i+1)
        Else
            Err.Clear
            ' Attempt 3: Sub-style call with all args on sheet
            On Error Resume Next
            sheet.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing
            If Err.Number = 0 Then
                Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
                WScript.Echo "Method Succeeded (sheet sub-style) for ISO" & (i+1)
            Else
                Err.Clear
                ' Attempt 4: drawDoc sub-style
                On Error Resume Next
                drawDoc.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i+1), Nothing
                If Err.Number = 0 Then
                    Set v = drawDoc.DrawingViews.Item(drawDoc.DrawingViews.Count)
                    WScript.Echo "Method Succeeded (drawDoc sub-style) for ISO" & (i+1)
                Else
                    Err.Clear
                    ' Attempt 5: Minimal args
                    On Error Resume Next
                    sheet.DrawingViews.AddBaseView asmDoc, positions(i), scale, orient
                    If Err.Number = 0 Then
                        Set v = sheet.DrawingViews.Item(sheet.DrawingViews.Count)
                        WScript.Echo "Method Succeeded (minimal) for ISO" & (i+1)
                    Else
                        Err.Clear
                        WScript.Echo "All attempts failed for ISO" & (i+1) & " - " & Err.Description
                        Err.Clear
                    End If
                End If
            End If
        End If
    End If

    Set views(i) = v
    If Not v Is Nothing Then
        WScript.Echo "INFO: ISO" & (i+1) & " object obtained: " & TypeName(v)
    End If
Next

' Add parts list anchored to the first ISO if available
If Not IsNull(views(0)) And Not views(0) Is Nothing Then
    On Error Resume Next
    Dim plPt
    Set plPt = tg.CreatePoint2d(CDbl(sheet.Width) - 40, CDbl(sheet.Height) - 40)
    sheet.PartsLists.Add views(0), plPt
    If Err.Number = 0 Then
        WScript.Echo "Parts list added and anchored to ISO1"
    Else
        WScript.Echo "Warning: could not add parts list - " & Err.Description
        Err.Clear
    End If
Else
    WScript.Echo "No ISO view available to anchor parts list"
End If

' Final update and save
On Error Resume Next
drawDoc.Update
If Err.Number = 0 Then
    drawDoc.Save
    WScript.Echo "Drawing updated and saved."
Else
    WScript.Echo "Warning: Update error - " & Err.Description
    Err.Clear
End If

WScript.Echo "Done. Open the created drawing to review the layout."

' ---------------------- helper functions ----------------------
Function GridPoint(sheet, tg, col, row, cols, rows)
    Dim margin, cellW, cellH, x, y
    margin = 10
    cellW = (sheet.Width - (2 * margin)) / cols
    cellH = (sheet.Height - (2 * margin)) / rows
    x = margin + (col + 0.5) * cellW
    y = margin + (row + 0.5) * cellH
    Set GridPoint = tg.CreatePoint2d(x, y)
End Function

Function ResolveTemplatePath()
    Dim baseDir
    baseDir = "C:\Users\Public\Documents\Autodesk\Inventor 2026\Templates\"
    If Not fso.FolderExists(baseDir) Then
        ResolveTemplatePath = ""
        Exit Function
    End If
    Dim rootFolder
    Set rootFolder = fso.GetFolder(baseDir)
    ResolveTemplatePath = RecursiveSearchForPentalin(rootFolder)
End Function

Function RecursiveSearchForPentalin(folder)
    Dim f, sf
    RecursiveSearchForPentalin = ""
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "idw" Then
            If InStr(1, LCase(f.Name), "pentalin", vbTextCompare) > 0 Then
                RecursiveSearchForPentalin = f.Path
                Exit Function
            End If
        End If
    Next
    For Each sf In folder.SubFolders
        Dim result
        result = RecursiveSearchForPentalin(sf)
        If result <> "" Then
            RecursiveSearchForPentalin = result
            Exit Function
        End If
    Next
End Function
