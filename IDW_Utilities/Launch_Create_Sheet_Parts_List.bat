@echo off
echo === CREATE SHEET-SPECIFIC PARTS LIST ===
echo.
echo This tool creates a parts list containing only the components
echo that are visible in the drawing views on the current sheet.
echo.
echo REQUIREMENTS:
echo - Inventor must be running
echo - A DRAWING document (.idw) must be open
echo - The sheet must contain drawing views
echo.
echo PROCESS:
echo 1. Scans all drawing views on the current sheet
echo 2. Collects all components visible in those views
echo 3. Creates a parts list with only those components
echo 4. Places the parts list in the bottom-right corner
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

echo.
echo Starting Sheet-Specific Parts List Creator...

REM Check if Inventor is running
tasklist /FI "IMAGENAME eq Inventor.exe" 2>NUL | find /I /N "Inventor.exe">NUL
if "%ERRORLEVEL%"=="1" (
    echo ERROR: Inventor is not running. Please start Inventor first.
    echo.
    pause
    exit /b 1
)

echo Inventor is running. Running VBScript...

REM Run the VBScript
cscript //nologo "Create_Sheet_Parts_List.vbs"

echo.
echo Sheet-specific parts list creation completed.
echo.
pause
    End If

    WScript.Echo "Found " & visibleComponents.Count & " visible components on this sheet."
    WScript.Echo ""

    ' Create filtered parts list
    CreateFilteredPartsList activeSheet, visibleComponents

    WScript.Echo ""
    WScript.Echo "Parts list created successfully!"

End Sub

Function GetVisibleComponentsOnSheet(sheet)
    Dim components
    Set components = CreateObject("Scripting.Dictionary")

    Dim view
    For Each view In sheet.DrawingViews
        If Not view Is Nothing Then
            Dim refDocDesc
            Set refDocDesc = view.ReferencedDocumentDescriptor

            If Not refDocDesc Is Nothing Then
                Dim doc
                Set doc = refDocDesc.ReferencedDocument

                If Not doc Is Nothing Then
                    If LCase(Right(doc.FullFileName, 4)) = ".iam" Then
                        ' Assembly - get all leaf components
                        CollectLeafComponentsFromAssembly doc, components
                    ElseIf LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                        ' Single part
                        components.Add refDocDesc.FullFileName, refDocDesc
                    End If
                End If
            End If
        End If
    Next

    Set GetVisibleComponentsOnSheet = components
End Function

Sub CollectLeafComponentsFromAssembly(asmDoc, components)
    On Error Resume Next

    Dim compDef
    Set compDef = asmDoc.ComponentDefinition

    If compDef Is Nothing Then Exit Sub

    Dim bom
    Set bom = compDef.BOM

    If bom Is Nothing Then Exit Sub

    Dim bomViews
    Set bomViews = bom.BOMViews

    If bomViews Is Nothing Then Exit Sub

    Dim structuredView
    Set structuredView = bomViews.Item("Structured")

    If structuredView Is Nothing Then Exit Sub

    Dim bomRows
    Set bomRows = structuredView.BOMRows

    Dim i
    For i = 1 To bomRows.Count
        Dim bomRow
        Set bomRow = bomRows.Item(i)
        CollectLeafComponentsFromBOMRow bomRow, components
    Next
End Sub

Sub CollectLeafComponentsFromBOMRow(bomRow, components)
    On Error Resume Next

    If bomRow.ChildRows Is Nothing Or bomRow.ChildRows.Count = 0 Then
        ' Leaf component
        Dim compDefs
        Set compDefs = bomRow.ComponentDefinitions

        If Not compDefs Is Nothing And compDefs.Count > 0 Then
            Dim compDef
            Set compDef = compDefs.Item(1)

            If Not compDef Is Nothing Then
                Dim doc
                Set doc = compDef.Document

                If Not doc Is Nothing Then
                    Dim refDescs
                    Set refDescs = doc.ReferencedFileDescriptors

                    If Not refDescs Is Nothing And refDescs.Count > 0 Then
                        Dim refDesc
                        Set refDesc = refDescs.Item(1)

                        If Not refDesc Is Nothing Then
                            If Not components.Exists(refDesc.FullFileName) Then
                                components.Add refDesc.FullFileName, refDesc
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        ' Recurse into sub-assemblies
        Dim childRows
        Set childRows = bomRow.ChildRows

        Dim j
        For j = 1 To childRows.Count
            Dim childRow
            Set childRow = childRows.Item(j)
            CollectLeafComponentsFromBOMRow childRow, components
        Next
    End If
End Sub

Sub CreateFilteredPartsList(sheet, visibleComponents)
    On Error Resume Next

    If sheet.DrawingViews.Count = 0 Then
        WScript.Echo "No drawing views found on this sheet."
        Exit Sub
    End If

    ' Use the first view as anchor
    Dim anchorView
    Set anchorView = sheet.DrawingViews.Item(1)

    ' Create placement point (bottom right corner)
    Dim tg
    Set tg = sheet.Application.TransientGeometry

    Dim plPoint
    Set plPoint = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)

    ' Create parts list
    Dim partsList
    Set partsList = sheet.PartsLists.Add(anchorView, plPoint)

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create parts list - " & Err.Description
        Err.Clear
        Exit Sub
    End If

    WScript.Echo "Parts list created. Filtering to sheet-specific components..."

    ' Filter rows - remove components not visible on this sheet
    Dim rowsToDelete
    Set rowsToDelete = CreateObject("Scripting.Dictionary")

    Dim row
    For Each row In partsList.PartsListRows
        Dim rowRefDesc
        Set rowRefDesc = row.ReferencedDocumentDescriptor

        If Not rowRefDesc Is Nothing Then
            If Not visibleComponents.Exists(rowRefDesc.FullFileName) Then
                rowsToDelete.Add row.ItemNumber, row
            End If
        End If
    Next

    ' Delete rows in reverse order to avoid index issues
    Dim itemNumbers
    itemNumbers = rowsToDelete.Keys

    Dim k
    For k = UBound(itemNumbers) To LBound(itemNumbers) Step -1
        Dim rowToDelete
        Set rowToDelete = rowsToDelete.Item(itemNumbers(k))
        rowToDelete.Delete
    Next

    WScript.Echo "Filtered parts list: removed " & rowsToDelete.Count & " non-visible components."

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Create_Sheet_Parts_List.vbs