Option Explicit
On Error Resume Next

Const kAssemblyDocumentObject = 12291
Const kDrawingDocumentObject = 12292
Const kIsoTopLeftViewOrientation = 13767
Const kIsoTopRightViewOrientation = 13768
Const kIsoBottomLeftViewOrientation = 13769
Const kIsoBottomRightViewOrientation = 13770

Dim invApp, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Set invApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Or invApp Is Nothing Then
    WScript.Echo "ERROR: Inventor is not running."
    WScript.Quit 1
End If
Err.Clear

' Find active assembly
Dim asmDoc
Set asmDoc = Nothing
If Not invApp.ActiveDocument Is Nothing Then
    If invApp.ActiveDocument.DocumentType = kAssemblyDocumentObject Then
        Set asmDoc = invApp.ActiveDocument
    End If
End If

If asmDoc Is Nothing Then
    ' fallback: pick first open assembly
    Dim d
    For Each d In invApp.Documents
        If d.DocumentType = kAssemblyDocumentObject Then
            Set asmDoc = d
            Exit For
        End If
    Next
End If

If asmDoc Is Nothing Then
    WScript.Echo "ERROR: No open assembly found. Open an assembly and re-run."
    WScript.Quit 1
End If

' Resolve template path (search recursively for Pentalin)
Dim templatePath
templatePath = ResolveTemplatePath()
If templatePath = "" Then
    WScript.Echo "ERROR: Cannot locate Pentalin template. Aborting."
    WScript.Quit 1
End If
WScript.Echo "Using template: " & templatePath

' Create drawing from template
Dim drawDoc
Set drawDoc = invApp.Documents.Add(kDrawingDocumentObject, templatePath, True)
If drawDoc Is Nothing Then
    WScript.Echo "ERROR: Could not create drawing from template."
    WScript.Quit 1
End If

' Save drawing next to assembly
Dim asmFolder, asmName, baseName, targetPath, index
asmFolder = GetDirectoryFromPath(asmDoc.FullFileName)
asmName = GetFileNameFromPath(asmDoc.FullFileName)
baseName = Left(asmName, InStrRev(asmName, ".") - 1)
index = 0
Do
    If index = 0 Then
        targetPath = asmFolder & "\" & baseName & "-DETAIL.idw"
    Else
        targetPath = asmFolder & "\" & baseName & "-DETAIL(" & index & ").idw"
    End If
    index = index + 1
Loop While fso.FileExists(targetPath)

On Error Resume Next
drawDoc.SaveAs targetPath, False
If Err.Number <> 0 Then
    WScript.Echo "Warning: SaveAs returned error: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Saved new drawing: " & targetPath
End If

' Place 4 isometric views at scale 1:20
Dim sheet, tg, scale
Set sheet = drawDoc.Sheets.Item(1)
Set tg = invApp.TransientGeometry
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
    Set views(i) = sheet.DrawingViews.AddBaseView(asmDoc, positions(i), scale, orient, True, Nothing, "ISO" & (i + 1), Nothing)
    If Err.Number <> 0 Or views(i) Is Nothing Then
        WScript.Echo "Warning: failed to add ISO" & (i + 1) & " - " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Added ISO" & (i + 1) & " at scale 1:20"
    End If
Next

' Add parts list anchored to first iso view
If Not IsNull(views(0)) And Not views(0) Is Nothing Then
    On Error Resume Next
    Dim plPoint
    Set plPoint = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)
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

' Update and save
drawDoc.Update
drawDoc.Save
WScript.Echo "Detail drawing complete and saved: " & targetPath

' Helper functions
Function ResolveTemplatePath()
    Dim baseDir, rootFolder
    baseDir = "C:\\Users\\Public\\Documents\\Autodesk\\Inventor 2026\\Templates\\"
    If Not fso.FolderExists(baseDir) Then
        ResolveTemplatePath = ""
        Exit Function
    End If
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

Function GridPoint(sheet, tg, col, row, cols, rows)
    Dim margin, cellW, cellH, x, y
    margin = 10
    cellW = (sheet.Width - (2 * margin)) / cols
    cellH = (sheet.Height - (2 * margin)) / rows
    x = margin + (col + 0.5) * cellW
    y = margin + (row + 0.5) * cellH
    Set GridPoint = tg.CreatePoint2d(x, y)
End Function

Function GetDirectoryFromPath(fullPath)
    GetDirectoryFromPath = Left(fullPath, InStrRev(fullPath, "\") - 1)
End Function

Function GetFileNameFromPath(fullPath)
    GetFileNameFromPath = Mid(fullPath, InStrRev(fullPath, "\") + 1)
End Function
