Option Explicit

' Test Detailer - non-destructive prototype
' - Requires an active assembly in Inventor
' - Creates a new drawing from the Pentalin template
' - Sheet1: four isometric views at 1:20 + parts list
' - Sheet2: elevations + plan at 1:20
' - Subsequent sheets: one isometric view per unique part (simple layout for testing)
' No existing project files are modified.

' Inventor constants (late-bound)
Const kAssemblyDocumentObject = 12291
Const kDrawingDocumentObject = 12292
Const kArbitraryViewOrientation = 13760
Const kFrontViewOrientation = 13761
Const kBackViewOrientation = 13762
Const kLeftViewOrientation = 13763
Const kRightViewOrientation = 13764
Const kTopViewOrientation = 13765
Const kBottomViewOrientation = 13766
Const kIsoTopLeftViewOrientation = 13767
Const kIsoTopRightViewOrientation = 13768
Const kIsoBottomLeftViewOrientation = 13769
Const kIsoBottomRightViewOrientation = 13770
Const kCurrentViewOrientation = 13771

Dim invApp
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Main

Sub Main()
    On Error Resume Next

    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Or invApp Is Nothing Then
        WScript.Echo "ERROR: Inventor is not running. Start Inventor with an assembly open first."
        Exit Sub
    End If
    Err.Clear

    If invApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document. Open an assembly and re-run."
        Exit Sub
    End If

    If invApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Active document is not an assembly (.iam)."
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = invApp.ActiveDocument

    Dim templatePath
    templatePath = ResolveTemplatePath()
    If templatePath = "" Then
        Exit Sub
    End If

    Dim drawDoc
    Set drawDoc = invApp.Documents.Add(kDrawingDocumentObject, templatePath, True)
    If drawDoc Is Nothing Then
        WScript.Echo "ERROR: Could not create drawing from template."
        Exit Sub
    End If

    Dim tg
    Set tg = invApp.TransientGeometry

    ' Sheet 1: four iso views and parts list
    Dim sheet1
    Set sheet1 = drawDoc.Sheets.Item(1)
    sheet1.Activate

    Dim isoScale
    isoScale = 1 / 20 ' 1:20

    Dim isoViews
    isoViews = PlaceIsoGrid(drawDoc, sheet1, asmDoc, tg, isoScale)

    If Not isoViews Is Nothing Then
        AddPartsList sheet1, isoViews(0), tg
    End If

    ' Sheet 2: elevations + plan
    Dim sheet2
    Set sheet2 = drawDoc.Sheets.Add(sheet1.Size, sheet1.Name)
    sheet2.Activate
    PlaceElevationSet drawDoc, sheet2, asmDoc, tg, isoScale
    If Not isoViews Is Nothing Then
        ' reuse first iso view for BOM; if not available, skip
        AddPartsList sheet2, isoViews(0), tg
    End If

    ' Part sheets: one per unique part for testing
    Dim partDocs
    Set partDocs = CollectUniqueParts(asmDoc)

    Dim i
    For i = 0 To partDocs.Count - 1
        Dim pDoc
        Set pDoc = partDocs.Item(i)
        Dim partSheet
        Set partSheet = drawDoc.Sheets.Add(sheet1.Size, "PART-" & (i + 1))
        partSheet.Activate

        Dim pt
        Set pt = GetCenteredPoint(partSheet, tg)
        Dim partView
        Set partView = drawDoc.DrawingViews.AddBaseView(pDoc, pt, isoScale, kIsoTopRightViewOrientation, True, Nothing, "PARTISO", Nothing)
        AddPartsList partSheet, partView, tg
    Next

    drawDoc.Sheets.Item(1).Activate
    WScript.Echo "Test drawing created from template. Check Inventor for the new IDW."
End Sub

Function PlaceIsoGrid(drawDoc, sheet, modelDoc, tg, scale)
    On Error Resume Next

    Dim positions(3)
    Set positions(0) = GridPoint(sheet, tg, 0, 0, 2, 2)
    Set positions(1) = GridPoint(sheet, tg, 1, 0, 2, 2)
    Set positions(2) = GridPoint(sheet, tg, 0, 1, 2, 2)
    Set positions(3) = GridPoint(sheet, tg, 1, 1, 2, 2)

    Dim views()
    ReDim views(3)

    views(0) = sheet.DrawingViews.AddBaseView(modelDoc, positions(0), scale, kIsoTopRightViewOrientation, True, Nothing, "ISO1", Nothing)
    views(1) = sheet.DrawingViews.AddBaseView(modelDoc, positions(1), scale, kIsoTopLeftViewOrientation, True, Nothing, "ISO2", Nothing)
    views(2) = sheet.DrawingViews.AddBaseView(modelDoc, positions(2), scale, kIsoBottomRightViewOrientation, True, Nothing, "ISO3", Nothing)
    views(3) = sheet.DrawingViews.AddBaseView(modelDoc, positions(3), scale, kIsoBottomLeftViewOrientation, True, Nothing, "ISO4", Nothing)

    PlaceIsoGrid = views
End Function

Sub PlaceElevationSet(drawDoc, sheet, modelDoc, tg, scale)
    On Error Resume Next

    Dim pFront, pRight, pLeft, pTop
    Set pFront = GridPoint(sheet, tg, 0, 0, 2, 2)
    Set pRight = GridPoint(sheet, tg, 1, 0, 2, 2)
    Set pLeft = GridPoint(sheet, tg, 0, 1, 2, 2)
    Set pTop = GridPoint(sheet, tg, 1, 1, 2, 2)

    drawDoc.DrawingViews.AddBaseView modelDoc, pFront, scale, kFrontViewOrientation, True, Nothing, "FRONT", Nothing
    drawDoc.DrawingViews.AddBaseView modelDoc, pRight, scale, kRightViewOrientation, True, Nothing, "RIGHT", Nothing
    drawDoc.DrawingViews.AddBaseView modelDoc, pLeft, scale, kLeftViewOrientation, True, Nothing, "LEFT", Nothing
    drawDoc.DrawingViews.AddBaseView modelDoc, pTop, scale, kTopViewOrientation, True, Nothing, "PLAN", Nothing
End Sub

Sub AddPartsList(sheet, anchorView, tg)
    On Error Resume Next
    If anchorView Is Nothing Then Exit Sub

    Dim plPoint
    Set plPoint = tg.CreatePoint2d(sheet.Width - 5, sheet.Height - 5) ' near upper-right corner
    sheet.PartsLists.Add anchorView, plPoint
End Sub

Function GridPoint(sheet, tg, col, row, cols, rows)
    Dim margin
    margin = 3 ' cm
    Dim cellW, cellH
    cellW = (sheet.Width - (2 * margin)) / cols
    cellH = (sheet.Height - (2 * margin)) / rows

    Dim x, y
    x = margin + (col + 0.5) * cellW
    y = margin + (row + 0.5) * cellH
    Set GridPoint = tg.CreatePoint2d(x, y)
End Function

Function GetCenteredPoint(sheet, tg)
    Set GetCenteredPoint = tg.CreatePoint2d(sheet.Width / 2, sheet.Height / 2)
End Function

Function ResolveTemplatePath()
    Dim baseDir, preferredName, fullPath
    baseDir = "C:\Users\Public\Documents\Autodesk\Inventor 2026\Templates\"
    preferredName = "Pentalin 1.idw"
    fullPath = baseDir & preferredName

    If fso.FileExists(fullPath) Then
        ResolveTemplatePath = fullPath
        Exit Function
    End If

    ' Search recursively through Templates and subfolders for a filename that contains "Pentalin"
    If fso.FolderExists(baseDir) Then
        Dim rootFolder
        Set rootFolder = fso.GetFolder(baseDir)
        Dim foundPath
        foundPath = RecursiveSearchForPentalin(rootFolder)
        If foundPath <> "" Then
            ResolveTemplatePath = foundPath
            WScript.Echo "INFO: Using template found by search: " & foundPath
            Exit Function
        End If
    End If

    WScript.Echo "ERROR: Template not found. Expected (or similar): " & fullPath
    ResolveTemplatePath = ""
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

Function CollectUniqueParts(asmDoc)
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    RecurseOccurrences asmDoc.ComponentDefinition.Occurrences, dict

    Dim arr
    arr = dict.Items

    Dim list
    Set list = CreateObject("System.Collections.ArrayList")
    Dim i
    For i = 0 To UBound(arr)
        list.Add arr(i)
    Next

    Set CollectUniqueParts = list
End Function

Sub RecurseOccurrences(occurrences, dict)
    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)
        If occ.Suppressed Then
            ' skip
        ElseIf LCase(Right(occ.Definition.Document.FullFileName, 4)) = ".ipt" Then
            If Not dict.Exists(occ.Definition.Document.FullFileName) Then
                dict.Add occ.Definition.Document.FullFileName, occ.Definition.Document
            End If
        ElseIf LCase(Right(occ.Definition.Document.FullFileName, 4)) = ".iam" Then
            RecurseOccurrences occ.Definition.Occurrences, dict
        End If
    Next
End Sub
