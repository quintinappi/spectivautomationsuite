' Sheet-Specific Parts List Creator
' Creates a parts list containing only components visible on the current sheet
' Author: Quintin de Bruin © 2026

Option Explicit

Const kDrawingDocumentObject = 12292

Dim m_InventorApp
Dim m_LogPath

Sub Main()
    On Error Resume Next

    WScript.Echo "=== SHEET-SPECIFIC PARTS LIST CREATOR ==="
    WScript.Echo ""

    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    WScript.Echo "Inventor application object type: " & TypeName(m_InventorApp)
    WScript.Echo "Inventor version: " & m_InventorApp.Version

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
        WScript.Echo "ERROR: Not a drawing document"
        Exit Sub
    End If

    Dim drawDoc
    Set drawDoc = m_InventorApp.ActiveDocument

    Dim activeSheet
    Set activeSheet = drawDoc.ActiveSheet

    WScript.Echo "Processing sheet: " & activeSheet.Name
    WScript.Echo ""

    ' Get all visible components on this sheet
    Dim visibleComponents
    Set visibleComponents = GetVisibleComponentsOnSheet(activeSheet)

    If visibleComponents.Count = 0 Then
        WScript.Echo "No components found on this sheet."
        Exit Sub
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

    WScript.Echo "Scanning " & sheet.DrawingViews.Count & " drawing views on sheet..."

    Dim view
    For Each view In sheet.DrawingViews
        On Error Resume Next
        If Not view Is Nothing Then
            Dim refDocDesc
            Set refDocDesc = view.ReferencedDocumentDescriptor

            If Err.Number <> 0 Then
                ' Skip views with errors
                Err.Clear
            ElseIf Not refDocDesc Is Nothing Then
                Dim doc
                Set doc = refDocDesc.ReferencedDocument

                If Err.Number <> 0 Then
                    Err.Clear
                ElseIf Not doc Is Nothing Then
                    If LCase(Right(doc.FullFileName, 4)) = ".iam" Then
                        ' Assembly - get all leaf components
                        CollectLeafComponentsFromAssembly doc, components
                    ElseIf LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                        ' Single part
                        If Not components.Exists(doc.FullFileName) Then
                            components.Add doc.FullFileName, refDocDesc
                        End If
                    End If
                End If
            End If
        End If
        On Error GoTo 0
    Next

    WScript.Echo "Found " & components.Count & " unique components across all views"
    WScript.Echo "Components found:"
    Dim key
    For Each key In components.Keys
        WScript.Echo "  - " & key
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

    WScript.Echo "Creating filtered parts list for " & visibleComponents.Count & " components..."

    ' Find an assembly view on any sheet to use as the basis
    Dim assemblyView
    Set assemblyView = Nothing
    
    Dim sh
    For Each sh In sheet.Parent.Sheets
        Dim v
        For Each v In sh.DrawingViews
            On Error Resume Next
            Dim refDesc
            Set refDesc = v.ReferencedDocumentDescriptor
            If Not refDesc Is Nothing Then
                Dim refDoc
                Set refDoc = refDesc.ReferencedDocument
                If Not refDoc Is Nothing Then
                    If LCase(Right(refDoc.FullFileName, 4)) = ".iam" Then
                        Set assemblyView = v
                        WScript.Echo "Found assembly view on sheet: " & sh.Name
                        Exit For
                    End If
                End If
            End If
            Err.Clear
        Next
        If Not assemblyView Is Nothing Then Exit For
    Next

    If assemblyView Is Nothing Then
        WScript.Echo "ERROR: No assembly view found in drawing to create parts list from"
        Exit Sub
    End If

    ' Create placement point (bottom right corner)
    Dim tg
    Set tg = sheet.Application.TransientGeometry

    Dim plPoint
    Set plPoint = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)

    ' Create parts list from the assembly
    WScript.Echo "Creating parts list from assembly..."
    Dim partsList
    Set partsList = sheet.PartsLists.Add(assemblyView, plPoint)

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Failed to create parts list - " & Err.Description
        Exit Sub
    End If

    WScript.Echo "Parts list created with " & partsList.PartsListRows.Count & " total rows"

    ' Hide rows for components NOT on this sheet
    WScript.Echo "Filtering parts list to show only components on this sheet..."
    
    ' Build a set of part numbers from visible components
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim visiblePartNumbers
    Set visiblePartNumbers = CreateObject("Scripting.Dictionary")
    
    Dim compPath
    For Each compPath In visibleComponents.Keys
        Dim partNum
        partNum = LCase(fso.GetBaseName(compPath))
        visiblePartNumbers.Add partNum, True
    Next
    
    WScript.Echo "Looking for these " & visiblePartNumbers.Count & " part numbers in the parts list"
    
    Dim hiddenCount
    hiddenCount = 0
    Dim visibleCount
    visibleCount = 0
    Dim rowNum
    rowNum = 0
    
    Dim row
    For Each row In partsList.PartsListRows
        rowNum = rowNum + 1
        On Error Resume Next
        
        ' Get the part number from the row - try different columns
        Dim rowPartNum
        rowPartNum = ""
        
        ' Try column 2 (usually part number column)
        rowPartNum = LCase(Trim(row.Item(2).Value))
        
        If rowPartNum = "" Then
            ' Try column 1
            rowPartNum = LCase(Trim(row.Item(1).Value))
        End If
        
        If rowNum <= 3 Then
            WScript.Echo "Row " & rowNum & " part number: [" & rowPartNum & "]"
        End If
        
        ' Check if this part number is on the current sheet
        If visiblePartNumbers.Exists(rowPartNum) Then
            visibleCount = visibleCount + 1
            If rowNum <= 3 Then
                WScript.Echo "  -> Keeping this row"
            End If
        Else
            row.Visible = False
            hiddenCount = hiddenCount + 1
        End If
        
        Err.Clear
    Next

    WScript.Echo "Hidden " & hiddenCount & " rows not on this sheet"
    WScript.Echo "Parts list shows " & visibleCount & " components from this sheet"

End Sub

Main