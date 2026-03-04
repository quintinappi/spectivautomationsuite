Imports Inventor
Imports System.Text
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class PopulateDwgRefTool
        Private ReadOnly m_InventorApp As Inventor.Application
        Private ReadOnly m_AutoPlaceMissing As Boolean

        Public Sub New(ByVal inventorApp As Inventor.Application, ByVal autoPlaceMissing As Boolean)
            m_InventorApp = inventorApp
            m_AutoPlaceMissing = autoPlaceMissing
        End Sub

        Public Sub Execute()
            Using progress As New ToolProgressForm("Populate DWG REF")
                progress.Show()
                progress.UpdateProgress(5, "Validating drawing context...")

            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running this tool.")
            End If

            progress.UpdateProgress(20, "Scanning parts lists...")
            Dim modelPaths As HashSet(Of String) = CollectModelPathsFromPartsLists(drawingDoc)
            If modelPaths.Count = 0 Then
                progress.UpdateProgress(100, "No referenced models found.")
                MessageBox.Show("No referenced models were found in non-DXF sheet parts lists.", "Populate DWG REF", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            If m_AutoPlaceMissing Then
                progress.UpdateProgress(30, "Auto-placing missing parts...")
                AutoPlaceUndetailedParts(drawingDoc, modelPaths)
            End If

            progress.UpdateProgress(45, "Collecting model placements...")
            Dim placements As Dictionary(Of String, SortedSet(Of Integer)) = CollectPlacementsByModel(drawingDoc)
            Dim drawingBaseName As String = GetDrawingBaseName(drawingDoc)
            progress.UpdateProgress(55, "Updating parts list DWG REF cells...")
            Dim partsListCellsUpdated As Integer = UpdatePartsListDwgRefCells(drawingDoc, drawingBaseName, placements)
            Dim updated As Integer = 0
            Dim errors As Integer = 0
            Dim processed As Integer = 0

            For Each modelPath As String In modelPaths
                Dim refs As String = BuildRefsString(drawingBaseName, modelPath, placements)
                If String.IsNullOrWhiteSpace(refs) Then
                    processed += 1
                    Continue For
                End If

                If DrawingAutomationHelpers.TrySetModelDwgRef(m_InventorApp, modelPath, refs, errors) Then
                    updated += 1
                End If

                processed += 1
                Dim pct As Integer = 45 + CInt((processed / Math.Max(1, modelPaths.Count)) * 45)
                progress.UpdateProgress(Math.Min(90, pct), "Updating model DWG REF values... " & processed.ToString() & "/" & modelPaths.Count.ToString())
            Next

            progress.UpdateProgress(95, "Saving drawing...")
            DrawingAutomationHelpers.SetDrawingUserProperty(drawingDoc, "DWG REF", BuildDrawingSummary(drawingBaseName, placements))
            drawingDoc.Update()
            drawingDoc.Save2(True)

                progress.CompleteSuccess("Populate DWG REF complete.")

            MessageBox.Show("Populate DWG REF complete." & vbCrLf & vbCrLf &
                            "Models found: " & modelPaths.Count.ToString() & vbCrLf &
                            "Parts list cells updated: " & partsListCellsUpdated.ToString() & vbCrLf &
                            "Models updated: " & updated.ToString() & vbCrLf &
                            "Update errors: " & errors.ToString(),
                            "Populate DWG REF", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        End Sub

        Private Function CollectModelPathsFromPartsLists(ByVal drawingDoc As DrawingDocument) As HashSet(Of String)
            Dim paths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each sheet As Sheet In drawingDoc.Sheets
                If DrawingAutomationHelpers.IsDxfSheet(sheet.Name) Then
                    Continue For
                End If

                For Each partsList As PartsList In sheet.PartsLists
                    For Each row As PartsListRow In partsList.PartsListRows
                        Dim rowPaths As List(Of String) = DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                        For Each pathValue As String In rowPaths
                            If Not String.IsNullOrWhiteSpace(pathValue) Then
                                paths.Add(pathValue)
                            End If
                        Next
                    Next
                Next
            Next

            Return paths
        End Function

        Private Function CollectPlacementsByModel(ByVal drawingDoc As DrawingDocument) As Dictionary(Of String, SortedSet(Of Integer))
            Dim map As New Dictionary(Of String, SortedSet(Of Integer))(StringComparer.OrdinalIgnoreCase)

            For Each sheet As Sheet In drawingDoc.Sheets
                If DrawingAutomationHelpers.IsDxfSheet(sheet.Name) Then
                    Continue For
                End If

                Dim sheetNumber As Integer = DrawingAutomationHelpers.GetSheetNumber(sheet)
                For Each view As DrawingView In sheet.DrawingViews
                    Dim pathValue As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                    If String.IsNullOrWhiteSpace(pathValue) Then
                        Continue For
                    End If

                    If Not map.ContainsKey(pathValue) Then
                        map(pathValue) = New SortedSet(Of Integer)()
                    End If
                    map(pathValue).Add(sheetNumber)
                Next
            Next

            Return map
        End Function

        Private Function BuildRefsString(ByVal drawingBase As String,
                                         ByVal modelPath As String,
                                         ByVal placements As Dictionary(Of String, SortedSet(Of Integer))) As String
            If Not placements.ContainsKey(modelPath) Then
                Return String.Empty
            End If

            Dim suffixes As New List(Of String)()
            For Each value As Integer In placements(modelPath)
                If value > 0 Then
                    suffixes.Add(value.ToString("00"))
                End If
            Next

            If suffixes.Count = 0 Then
                Return String.Empty
            End If

            Return drawingBase & "-" & String.Join("/", suffixes)
        End Function

        Private Function BuildRefsString(ByVal drawingBase As String,
                                         ByVal sheetNumbers As IEnumerable(Of Integer)) As String
            Dim suffixes As New List(Of String)()
            For Each value As Integer In sheetNumbers.OrderBy(Function(x) x)
                If value > 0 Then
                    suffixes.Add(value.ToString("00"))
                End If
            Next

            If suffixes.Count = 0 Then
                Return String.Empty
            End If

            Return drawingBase & "-" & String.Join("/", suffixes)
        End Function

        Private Function GetDrawingBaseName(ByVal drawingDoc As DrawingDocument) As String
            If drawingDoc Is Nothing Then
                Return "DRAWING"
            End If

            Dim drawingBase As String = String.Empty

            Try
                drawingBase = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.FullFileName)
            Catch
                drawingBase = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(drawingBase) Then
                Try
                    drawingBase = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName)
                Catch
                    drawingBase = String.Empty
                End Try
            End If

            If String.IsNullOrWhiteSpace(drawingBase) Then
                Return "DRAWING"
            End If

            Return drawingBase.Trim()
        End Function

        Private Function UpdatePartsListDwgRefCells(ByVal drawingDoc As DrawingDocument,
                                                    ByVal drawingBase As String,
                                                    ByVal placements As Dictionary(Of String, SortedSet(Of Integer))) As Integer
            Dim updatedCount As Integer = 0

            For Each sheet As Sheet In drawingDoc.Sheets
                If DrawingAutomationHelpers.IsDxfSheet(sheet.Name) Then
                    Continue For
                End If

                For Each partsList As PartsList In sheet.PartsLists
                    Dim dwgRefColumnIndex As Integer = FindDwgRefColumnIndex(partsList)
                    If dwgRefColumnIndex <= 0 Then
                        Continue For
                    End If

                    For Each row As PartsListRow In partsList.PartsListRows
                        Dim rowPaths As List(Of String) = GetPreferredPathsFromRow(row)
                        Dim rowSheets As New SortedSet(Of Integer)()

                        For Each rowPath As String In rowPaths
                            If placements.ContainsKey(rowPath) Then
                                For Each sheetNo As Integer In placements(rowPath)
                                    rowSheets.Add(sheetNo)
                                Next
                            End If
                        Next

                        Dim rowRefValue As String = BuildRefsString(drawingBase, rowSheets)

                        Try
                            Dim cell As PartsListCell = row.Item(dwgRefColumnIndex)
                            If cell IsNot Nothing Then
                                cell.Value = rowRefValue
                                updatedCount += 1
                            End If
                        Catch
                        End Try
                    Next
                Next
            Next

            Return updatedCount
        End Function

        Private Function FindDwgRefColumnIndex(ByVal partsList As PartsList) As Integer
            For idx As Integer = 1 To partsList.PartsListColumns.Count
                Try
                    Dim col As PartsListColumn = partsList.PartsListColumns.Item(idx)
                    If IsDwgRefTitle(col.Title) Then
                        Return idx
                    End If
                Catch
                End Try
            Next

            Return 0
        End Function

        Private Function IsDwgRefTitle(ByVal value As String) As Boolean
            Dim normalized As String = value.ToUpperInvariant()
            normalized = normalized.Replace(" ", String.Empty)
            normalized = normalized.Replace(".", String.Empty)
            normalized = normalized.Replace("_", String.Empty)
            normalized = normalized.Replace("-", String.Empty)
            Return normalized.Contains("DWGREF")
        End Function

        Private Function GetPreferredPathsFromRow(ByVal row As PartsListRow) As List(Of String)
            Dim iptPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim iamPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each pathValue As String In DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                If String.IsNullOrWhiteSpace(pathValue) Then
                    Continue For
                End If

                If pathValue.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    iptPaths.Add(pathValue)
                ElseIf pathValue.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    iamPaths.Add(pathValue)
                End If
            Next

            If iptPaths.Count > 0 Then
                Return iptPaths.OrderBy(Function(x) x).ToList()
            End If

            If iamPaths.Count > 0 Then
                Return iamPaths.OrderBy(Function(x) x).ToList()
            End If

            Return New List(Of String)()
        End Function

        Private Function BuildDrawingSummary(ByVal drawingBase As String,
                                             ByVal placements As Dictionary(Of String, SortedSet(Of Integer))) As String
            Dim builder As New StringBuilder()
            builder.AppendLine("Drawing: " & drawingBase)

            Dim orderedKeys As List(Of String) = placements.Keys.OrderBy(Function(x) x).ToList()
            For Each modelPath As String In orderedKeys
                Dim refs As String = BuildRefsString(drawingBase, modelPath, placements)
                If Not String.IsNullOrWhiteSpace(refs) Then
                    builder.AppendLine(System.IO.Path.GetFileNameWithoutExtension(modelPath) & " = " & refs)
                End If
            Next

            Return builder.ToString().Trim()
        End Function

        Private Sub AutoPlaceUndetailedParts(ByVal drawingDoc As DrawingDocument, ByVal modelPaths As HashSet(Of String))
            Dim placedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each sheet As Sheet In drawingDoc.Sheets
                For Each view As DrawingView In sheet.DrawingViews
                    Dim pathValue As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                    If Not String.IsNullOrWhiteSpace(pathValue) Then
                        placedPaths.Add(pathValue)
                    End If
                Next
            Next

            Dim missingParts As List(Of String) = modelPaths.
                Where(Function(pathValue) pathValue.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) AndAlso Not placedPaths.Contains(pathValue)).
                OrderBy(Function(pathValue) pathValue).
                ToList()

            If missingParts.Count = 0 Then
                Return
            End If

            Dim targetSheet As Sheet = DrawingAutomationHelpers.SelectNonDxfSheet(drawingDoc, "Populate DWG REF + Auto-place Missing Parts", "Select target sheet for placing missing parts:")
            If targetSheet Is Nothing Then
                Return
            End If

            Dim slot As Integer = 0
            Dim perSheet As Integer = 9
            Dim extraSheetIndex As Integer = 0
            Dim currentSheet As Sheet = targetSheet

            For Each partPath As String In missingParts
                Dim partDoc As PartDocument = DrawingAutomationHelpers.OpenPartDocument(m_InventorApp, partPath)
                If partDoc Is Nothing Then
                    Continue For
                End If

                Dim isPlate As Boolean = DrawingAutomationHelpers.IsPlatePart(partDoc)
                Dim neededSlots As Integer = If(isPlate, 2, 1)
                If slot + neededSlots > perSheet Then
                    extraSheetIndex += 1
                    currentSheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, targetSheet, "AUTO PARTS " & extraSheetIndex.ToString())
                    slot = 0
                End If

                Dim point1 As Point2d = DrawingAutomationHelpers.GetGridPoint(m_InventorApp, currentSheet, slot, 3, 3)
                DrawingAutomationHelpers.AddPartView(currentSheet, partDoc, point1, 0.1, True)
                slot += 1

                If isPlate Then
                    Dim point2 As Point2d = DrawingAutomationHelpers.GetGridPoint(m_InventorApp, currentSheet, slot, 3, 3)
                    DrawingAutomationHelpers.AddPartView(currentSheet, partDoc, point2, 0.1, False)
                    slot += 1
                End If
            Next
        End Sub
    End Class

    Public Class CreateDxfForModelPlatesTool
        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Using progress As New ToolProgressForm("CREATE DXF FOR MODEL PLATES")
                progress.Show()
                progress.UpdateProgress(5, "Validating drawing context...")

            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running this tool.")
            End If

            progress.UpdateProgress(15, "Selecting source sheet...")
            Dim sourceSheet As Sheet = DrawingAutomationHelpers.SelectNonDxfSheet(drawingDoc, "CREATE DXF FOR MODEL PLATES", "Select source sheet (must include assembly view):")
            If sourceSheet Is Nothing Then
                Return
            End If

            progress.UpdateProgress(30, "Finding source assembly...")
            Dim assemblyDoc As AssemblyDocument = FindAssemblyOnSheet(sourceSheet)
            If assemblyDoc Is Nothing Then
                MessageBox.Show("No assembly view found on the selected sheet.", "CREATE DXF FOR MODEL PLATES", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            progress.UpdateProgress(40, "Collecting plate parts...")
            Dim plateParts As List(Of PartDocument) = CollectPlateParts(assemblyDoc)
            If plateParts.Count = 0 Then
                MessageBox.Show("No plate parts found in selected assembly.", "CREATE DXF FOR MODEL PLATES", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim modelName As String = System.IO.Path.GetFileNameWithoutExtension(assemblyDoc.FullFileName)
            Dim firstSheet As Sheet = drawingDoc.Sheets.Add(sourceSheet.Size)
            firstSheet.Name = "DXF FOR " & modelName

            Dim currentSheet As Sheet = firstSheet
            Dim slot As Integer = 0
            Dim perSheet As Integer = 12
            Dim overflowIndex As Integer = 0
            Dim firstView As DrawingView = Nothing
            Dim placed As Integer = 0

            For Each partDoc As PartDocument In plateParts
                If slot >= perSheet Then
                    overflowIndex += 1
                    currentSheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, firstSheet, "DXF FOR " & modelName & "-" & (overflowIndex + 1).ToString())
                    slot = 0
                End If

                Dim pointValue As Point2d = DrawingAutomationHelpers.GetGridPoint(m_InventorApp, currentSheet, slot, 4, 3)
                Dim viewValue As DrawingView = DrawingAutomationHelpers.AddPartView(currentSheet, partDoc, pointValue, 1.0, False)
                If firstView Is Nothing AndAlso viewValue IsNot Nothing Then
                    firstView = viewValue
                End If

                slot += 1
                placed += 1
                Dim pct As Integer = 45 + CInt((placed / Math.Max(1, plateParts.Count)) * 45)
                progress.UpdateProgress(Math.Min(90, pct), "Placing plate views... " & placed.ToString() & "/" & plateParts.Count.ToString())
            Next

            If firstView IsNot Nothing Then
                Dim partsList As PartsList = DrawingAutomationHelpers.TryAddAssemblyPartsList(firstSheet, assemblyDoc, "Parts List", m_InventorApp)
                If partsList IsNot Nothing Then
                    FilterPartsListToPlates(partsList)
                End If
            End If

            progress.UpdateProgress(95, "Saving drawing...")
            drawingDoc.Update()
            drawingDoc.Save2(True)

                progress.CompleteSuccess("DXF creation complete.")

            MessageBox.Show("CREATE DXF FOR MODEL PLATES complete." & vbCrLf & vbCrLf &
                            "Assembly: " & modelName & vbCrLf &
                            "Plate parts placed: " & plateParts.Count.ToString(),
                            "CREATE DXF FOR MODEL PLATES", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        End Sub

        Private Function FindAssemblyOnSheet(ByVal sourceSheet As Sheet) As AssemblyDocument
            For Each view As DrawingView In sourceSheet.DrawingViews
                Dim viewPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                If viewPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    Dim referenced As Document = Nothing
                    Try
                        referenced = view.ReferencedDocumentDescriptor.ReferencedDocument
                    Catch
                        referenced = Nothing
                    End Try

                    If referenced IsNot Nothing AndAlso referenced.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Return CType(referenced, AssemblyDocument)
                    End If
                End If
            Next

            Return Nothing
        End Function

        Private Function CollectPlateParts(ByVal assemblyDoc As AssemblyDocument) As List(Of PartDocument)
            Dim map As New Dictionary(Of String, PartDocument)(StringComparer.OrdinalIgnoreCase)

            For Each occurrence As ComponentOccurrence In assemblyDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
                If occurrence.Suppressed Then
                    Continue For
                End If

                Dim doc As Document = Nothing
                Try
                    doc = occurrence.Definition.Document
                Catch
                    doc = Nothing
                End Try

                Dim partDoc As PartDocument = TryCast(doc, PartDocument)
                If partDoc Is Nothing Then
                    Continue For
                End If

                If DrawingAutomationHelpers.IsPlatePart(partDoc) Then
                    Dim key As String = partDoc.FullFileName
                    If Not map.ContainsKey(key) Then
                        map.Add(key, partDoc)
                    End If
                End If
            Next

            Return map.Values.OrderBy(Function(x) x.DisplayName).ToList()
        End Function

        Private Sub FilterPartsListToPlates(ByVal partsList As PartsList)
            For Each row As PartsListRow In partsList.PartsListRows
                Dim keepRow As Boolean = False
                Dim rowPaths As List(Of String) = DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                For Each pathValue As String In rowPaths
                    If pathValue.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                        Dim doc As PartDocument = DrawingAutomationHelpers.OpenPartDocument(m_InventorApp, pathValue)
                        If doc IsNot Nothing AndAlso DrawingAutomationHelpers.IsPlatePart(doc) Then
                            keepRow = True
                            Exit For
                        End If
                    End If
                Next

                If Not keepRow Then
                    Try
                        row.Visible = False
                    Catch
                    End Try
                End If
            Next
        End Sub
    End Class

    Public Class PlacePartsFromOpenAssemblyTool
        Private ReadOnly m_InventorApp As Inventor.Application

        Private Class AssemblyNode
            Public Property AssemblyDoc As AssemblyDocument
            Public Property Name As String
            Public Property Parts As List(Of PartDocument)
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Using progress As New ToolProgressForm("Place parts from open Assembly")
                progress.Show()
                progress.UpdateProgress(5, "Validating assembly and drawing...")

            Dim assemblyDoc As AssemblyDocument = DrawingAutomationHelpers.GetOpenAssembly(m_InventorApp)
            If assemblyDoc Is Nothing Then
                Throw New InvalidOperationException("Open an assembly (.iam) before running Place parts from open Assembly.")
            End If

            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running Place parts from open Assembly.")
            End If

            Dim nodes As List(Of AssemblyNode) = BuildAssemblySequence(assemblyDoc)
            If nodes.Count = 0 Then
                MessageBox.Show("No assemblies were detected for placement.", "Place parts from open Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim referenceSheet As Sheet = drawingDoc.ActiveSheet
            Dim totalNodes As Integer = nodes.Count
            Dim nodeIndex As Integer = 0

            For Each node As AssemblyNode In nodes
                nodeIndex += 1
                progress.UpdateProgress(10 + CInt((nodeIndex / Math.Max(1, totalNodes)) * 80), "Creating sheets for " & node.Name & " (" & nodeIndex.ToString() & "/" & totalNodes.ToString() & ")")
                Dim isoSheet As Sheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, referenceSheet, node.Name & " - ISO")
                CreateIsoSheet(node.AssemblyDoc, isoSheet)

                Dim baseSheet As Sheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, referenceSheet, node.Name & " - BASE")
                CreateBaseSheet(node.AssemblyDoc, baseSheet)

                If node.Parts.Count > 0 Then
                    CreatePartSheets(node, drawingDoc, referenceSheet)
                End If
            Next

            progress.UpdateProgress(95, "Saving drawing...")
            drawingDoc.Update()
            drawingDoc.Save2(True)

                progress.CompleteSuccess("Place parts complete.")

            MessageBox.Show("Place parts from open Assembly complete." & vbCrLf & vbCrLf &
                            "Assembly groups processed: " & nodes.Count.ToString(),
                            "Place parts from open Assembly", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        End Sub

        Private Function BuildAssemblySequence(ByVal rootAssembly As AssemblyDocument) As List(Of AssemblyNode)
            Dim list As New List(Of AssemblyNode)()

            list.Add(New AssemblyNode With {
                .AssemblyDoc = rootAssembly,
                .Name = System.IO.Path.GetFileNameWithoutExtension(rootAssembly.FullFileName),
                .Parts = CollectPartsForAssembly(rootAssembly)
            })

            Dim subAssemblies As New Dictionary(Of String, AssemblyDocument)(StringComparer.OrdinalIgnoreCase)
            For Each occurrence As ComponentOccurrence In rootAssembly.ComponentDefinition.Occurrences
                If occurrence.Suppressed Then
                    Continue For
                End If

                Dim subAsm As AssemblyDocument = Nothing
                Try
                    subAsm = TryCast(occurrence.Definition.Document, AssemblyDocument)
                Catch
                    subAsm = Nothing
                End Try

                If subAsm IsNot Nothing AndAlso Not subAssemblies.ContainsKey(subAsm.FullFileName) Then
                    subAssemblies.Add(subAsm.FullFileName, subAsm)
                End If
            Next

            For Each subAsm As AssemblyDocument In subAssemblies.Values.OrderBy(Function(x) x.DisplayName)
                list.Add(New AssemblyNode With {
                    .AssemblyDoc = subAsm,
                    .Name = System.IO.Path.GetFileNameWithoutExtension(subAsm.FullFileName),
                    .Parts = CollectPartsForAssembly(subAsm)
                })
            Next

            Return list
        End Function

        Private Function CollectPartsForAssembly(ByVal assemblyDoc As AssemblyDocument) As List(Of PartDocument)
            Dim map As New Dictionary(Of String, PartDocument)(StringComparer.OrdinalIgnoreCase)

            For Each occurrence As ComponentOccurrence In assemblyDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
                If occurrence.Suppressed Then
                    Continue For
                End If

                Dim partDoc As PartDocument = Nothing
                Try
                    partDoc = TryCast(occurrence.Definition.Document, PartDocument)
                Catch
                    partDoc = Nothing
                End Try

                If partDoc IsNot Nothing AndAlso Not map.ContainsKey(partDoc.FullFileName) Then
                    map.Add(partDoc.FullFileName, partDoc)
                End If
            Next

            Return map.Values.OrderBy(Function(x) x.DisplayName).ToList()
        End Function

        Private Sub CreateIsoSheet(ByVal assemblyDoc As AssemblyDocument, ByVal sheet As Sheet)
            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim points As New List(Of Point2d) From {
                tg.CreatePoint2d(sheet.Width * 0.25, sheet.Height * 0.70),
                tg.CreatePoint2d(sheet.Width * 0.75, sheet.Height * 0.70),
                tg.CreatePoint2d(sheet.Width * 0.25, sheet.Height * 0.30),
                tg.CreatePoint2d(sheet.Width * 0.75, sheet.Height * 0.30)
            }

            Dim orientations As New List(Of ViewOrientationTypeEnum) From {
                ViewOrientationTypeEnum.kIsoTopRightViewOrientation,
                ViewOrientationTypeEnum.kIsoTopLeftViewOrientation,
                ViewOrientationTypeEnum.kIsoBottomRightViewOrientation,
                ViewOrientationTypeEnum.kIsoBottomLeftViewOrientation
            }

            Dim firstView As DrawingView = Nothing
            For index As Integer = 0 To points.Count - 1
                Dim viewValue As DrawingView = sheet.DrawingViews.AddBaseView(
                    assemblyDoc,
                    points(index),
                    0.05,
                    orientations(index),
                    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)

                If firstView Is Nothing Then
                    firstView = viewValue
                End If
            Next

            If firstView IsNot Nothing Then
                DrawingAutomationHelpers.ConfigureAssemblyBomForList(assemblyDoc, True)
                DrawingAutomationHelpers.TryAddPartsList(sheet, firstView, "Parts List", m_InventorApp)
            End If
        End Sub

        Private Sub CreateBaseSheet(ByVal assemblyDoc As AssemblyDocument, ByVal sheet As Sheet)
            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim viewPoint As Point2d = tg.CreatePoint2d(sheet.Width * 0.5, sheet.Height * 0.55)
            Dim baseView As DrawingView = sheet.DrawingViews.AddBaseView(
                assemblyDoc,
                viewPoint,
                0.1,
                ViewOrientationTypeEnum.kFrontViewOrientation,
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)

            DrawingAutomationHelpers.ConfigureAssemblyBomForList(assemblyDoc, True)
            DrawingAutomationHelpers.TryAddPartsList(sheet, baseView, "Parts List", m_InventorApp)
        End Sub

        Private Sub CreatePartSheets(ByVal node As AssemblyNode, ByVal drawingDoc As DrawingDocument, ByVal referenceSheet As Sheet)
            Dim currentSheet As Sheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, referenceSheet, node.Name & " - PARTS")
            Dim slot As Integer = 0
            Dim perSheet As Integer = 9
            Dim pageIndex As Integer = 1

            For Each partDoc As PartDocument In node.Parts
                If slot >= perSheet Then
                    pageIndex += 1
                    currentSheet = DrawingAutomationHelpers.CreateSiblingSheet(drawingDoc, referenceSheet, node.Name & " - PARTS " & pageIndex.ToString())
                    slot = 0
                End If

                Dim pointValue As Point2d = DrawingAutomationHelpers.GetGridPoint(m_InventorApp, currentSheet, slot, 3, 3)
                DrawingAutomationHelpers.AddPartView(currentSheet, partDoc, pointValue, 0.1, True)
                slot += 1
            Next
        End Sub

        Private Sub CreateAndFilterSheetPartsList(ByVal sheet As Sheet, ByVal assemblyDoc As AssemblyDocument)
            AddInDiagnostics.Log("CreateAndFilterSheetPartsList", "Start | Sheet='" & SafeSheetName(sheet) & "' | Assembly='" & SafeDocName(assemblyDoc) & "'")
            Dim partsList As PartsList = DrawingAutomationHelpers.TryAddAssemblyPartsList(sheet, assemblyDoc, "Parts List", m_InventorApp)
            If partsList Is Nothing Then
                AddInDiagnostics.Log("CreateAndFilterSheetPartsList", "FAILED | Parts list was not created | Sheet='" & SafeSheetName(sheet) & "' | Assembly='" & SafeDocName(assemblyDoc) & "'")
                Return
            End If

            FilterPartsListToVisibleSheetModels(sheet, partsList)
            RenumberVisiblePartsListRows(partsList)

            Dim totalRows As Integer = 0
            Dim visibleRows As Integer = 0
            Try
                totalRows = partsList.PartsListRows.Count
                For Each row As PartsListRow In partsList.PartsListRows
                    Try
                        If row.Visible Then
                            visibleRows += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            AddInDiagnostics.Log("CreateAndFilterSheetPartsList", "Success | TotalRows=" & totalRows.ToString() & " | VisibleRows=" & visibleRows.ToString() & " | Sheet='" & SafeSheetName(sheet) & "'")
        End Sub

        Private Sub FilterPartsListToVisibleSheetModels(ByVal sheet As Sheet, ByVal partsList As PartsList)
            Dim visiblePartPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim visiblePartNumbers As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each view As DrawingView In sheet.DrawingViews
                Dim modelPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                If String.IsNullOrWhiteSpace(modelPath) Then
                    Continue For
                End If

                If Not modelPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                visiblePartPaths.Add(modelPath)

                Try
                    Dim partNumber As String = System.IO.Path.GetFileNameWithoutExtension(modelPath)
                    If Not String.IsNullOrWhiteSpace(partNumber) Then
                        visiblePartNumbers.Add(partNumber.Trim())
                    End If
                Catch
                End Try

                Try
                    Dim partDoc As PartDocument = DrawingAutomationHelpers.OpenPartDocument(m_InventorApp, modelPath)
                    If partDoc IsNot Nothing Then
                        Dim designSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                        Dim propPartNo As String = Convert.ToString(designSet.Item("Part Number").Value)
                        If Not String.IsNullOrWhiteSpace(propPartNo) Then
                            visiblePartNumbers.Add(propPartNo.Trim())
                        End If

                        Dim propDesc As String = Convert.ToString(designSet.Item("Description").Value)
                        If Not String.IsNullOrWhiteSpace(propDesc) Then
                            visiblePartNumbers.Add(propDesc.Trim())
                        End If
                    End If
                Catch
                End Try
            Next

            If visiblePartNumbers.Count = 0 AndAlso visiblePartPaths.Count = 0 Then
                AddInDiagnostics.Log("FilterPartsListToVisibleSheetModels", "No visible .ipt models found on sheet; skipping filter | Sheet='" & SafeSheetName(sheet) & "'")
                Return
            End If

            AddInDiagnostics.Log("FilterPartsListToVisibleSheetModels", "Visible sets built | Paths=" & visiblePartPaths.Count.ToString() & " | Keys=" & visiblePartNumbers.Count.ToString() & " | Sheet='" & SafeSheetName(sheet) & "'")

            Dim kept As Integer = 0
            Dim hidden As Integer = 0

            For Each row As PartsListRow In partsList.PartsListRows
                Dim keepRow As Boolean = False

                For Each rowPath As String In DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                    If String.IsNullOrWhiteSpace(rowPath) Then
                        Continue For
                    End If

                    If visiblePartPaths.Contains(rowPath) Then
                        keepRow = True
                        Exit For
                    End If
                Next

                If Not keepRow Then
                    Dim rowPartNumber As String = String.Empty

                    Try
                        rowPartNumber = Convert.ToString(row.Item(2).Value)
                    Catch
                        rowPartNumber = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(rowPartNumber) Then
                        Try
                            rowPartNumber = Convert.ToString(row.Item(1).Value)
                        Catch
                            rowPartNumber = String.Empty
                        End Try
                    End If

                    rowPartNumber = If(rowPartNumber, String.Empty).Trim()
                    If Not String.IsNullOrWhiteSpace(rowPartNumber) AndAlso visiblePartNumbers.Contains(rowPartNumber) Then
                        keepRow = True
                    End If
                End If

                If Not keepRow Then
                    Try
                        row.Visible = False
                        hidden += 1
                    Catch
                    End Try
                Else
                    kept += 1
                End If
            Next

            Try
                Dim sampleIndex As Integer = 0
                For Each row As PartsListRow In partsList.PartsListRows
                    sampleIndex += 1
                    If sampleIndex > 5 Then
                        Exit For
                    End If

                    Dim c1 As String = String.Empty
                    Dim c2 As String = String.Empty
                    Dim pathSample As String = String.Empty

                    Try
                        c1 = Convert.ToString(row.Item(1).Value)
                    Catch
                    End Try

                    Try
                        c2 = Convert.ToString(row.Item(2).Value)
                    Catch
                    End Try

                    Try
                        Dim paths As List(Of String) = DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                        If paths IsNot Nothing AndAlso paths.Count > 0 Then
                            pathSample = paths(0)
                        End If
                    Catch
                    End Try

                    AddInDiagnostics.Log("FilterPartsListToVisibleSheetModels", "RowSample " & sampleIndex.ToString() & " | C1='" & c1 & "' | C2='" & c2 & "' | Path='" & pathSample & "'")
                Next
            Catch
            End Try

            AddInDiagnostics.Log("FilterPartsListToVisibleSheetModels", "Filter complete | Kept=" & kept.ToString() & " | Hidden=" & hidden.ToString() & " | Sheet='" & SafeSheetName(sheet) & "'")
        End Sub

        Private Sub RenumberVisiblePartsListRows(ByVal partsList As PartsList)
            Dim itemNumber As Integer = 1

            For Each row As PartsListRow In partsList.PartsListRows
                Dim rowVisible As Boolean = True
                Try
                    rowVisible = row.Visible
                Catch
                End Try

                If Not rowVisible Then
                    Continue For
                End If

                Try
                    row.Item(1).Value = itemNumber.ToString("00")
                Catch
                End Try

                itemNumber += 1
            Next

            AddInDiagnostics.Log("RenumberVisiblePartsListRows", "Renumbered visible rows | Count=" & (itemNumber - 1).ToString())
        End Sub

        Private Function SafeSheetName(ByVal sheet As Sheet) As String
            Try
                If sheet IsNot Nothing Then
                    Return sheet.Name
                End If
            Catch
            End Try

            Return "<null>"
        End Function

        Private Function SafeDocName(ByVal doc As Document) As String
            Try
                If doc IsNot Nothing Then
                    If Not String.IsNullOrWhiteSpace(doc.FullFileName) Then
                        Return doc.FullFileName
                    End If
                    Return doc.DisplayName
                End If
            Catch
            End Try

            Return "<null>"
        End Function
    End Class

    Public Class CleanUpUnusedFilesTool
        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running Clean Up Unused Files.")
            End If

            Dim drawingPath As String = String.Empty
            Try
                drawingPath = drawingDoc.FullFileName
            Catch
                drawingPath = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(drawingPath) Then
                Throw New InvalidOperationException("Active IDW must be saved before running Clean Up Unused Files.")
            End If

            Dim drawingFolder As String = System.IO.Path.GetDirectoryName(drawingPath)
            If String.IsNullOrWhiteSpace(drawingFolder) OrElse Not System.IO.Directory.Exists(drawingFolder) Then
                Throw New InvalidOperationException("Could not locate the active drawing folder.")
            End If

            Dim referencedPartFileNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim visitedAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim openedByTool As New List(Of Document)()

            Try
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Start | Drawing='" & drawingPath & "'")

                CollectReferencedPartFileNames(drawingDoc, referencedPartFileNames, visitedAssemblies, openedByTool)
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Referenced parts collected | Count=" & referencedPartFileNames.Count.ToString())

                If referencedPartFileNames.Count = 0 Then
                    MessageBox.Show("No referenced part files were found in the active drawing context.",
                                    "Clean Up Unused Files",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                    Return
                End If

                Dim allIptFiles As New List(Of String)()
                CollectAllIptFiles(drawingFolder, allIptFiles)

                Dim unreferenced As New List(Of String)()
                For Each iptPath As String In allIptFiles
                    Dim fileName As String = String.Empty
                    Try
                        fileName = System.IO.Path.GetFileName(iptPath)
                    Catch
                        fileName = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(fileName) Then
                        Continue For
                    End If

                    If Not referencedPartFileNames.Contains(fileName) Then
                        unreferenced.Add(iptPath)
                    End If
                Next

                Dim moved As Integer = 0
                Dim moveErrors As Integer = 0
                Dim skippedExisting As Integer = 0

                If unreferenced.Count > 0 Then
                    Dim destinationFolder As String = System.IO.Path.Combine(drawingFolder, "Unrenamed Parts")
                    If Not System.IO.Directory.Exists(destinationFolder) Then
                        System.IO.Directory.CreateDirectory(destinationFolder)
                    End If

                    For Each sourcePath As String In unreferenced
                        Dim fileName As String = String.Empty
                        Try
                            fileName = System.IO.Path.GetFileName(sourcePath)
                        Catch
                            fileName = String.Empty
                        End Try

                        If String.IsNullOrWhiteSpace(fileName) Then
                            Continue For
                        End If

                        Dim destinationPath As String = System.IO.Path.Combine(destinationFolder, fileName)
                        If System.IO.File.Exists(destinationPath) Then
                            skippedExisting += 1
                            AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Skip existing destination | File='" & destinationPath & "'")
                            Continue For
                        End If

                        Try
                            System.IO.File.Move(sourcePath, destinationPath)
                            moved += 1
                        Catch ex As Exception
                            moveErrors += 1
                            AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Move failed | Source='" & sourcePath & "' | Destination='" & destinationPath & "' | " & ex.Message)
                        End Try
                    Next
                End If

                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Complete | Referenced=" & referencedPartFileNames.Count.ToString() & " | FolderIpt=" & allIptFiles.Count.ToString() & " | Unreferenced=" & unreferenced.Count.ToString() & " | Moved=" & moved.ToString() & " | SkippedExisting=" & skippedExisting.ToString() & " | MoveErrors=" & moveErrors.ToString())

                MessageBox.Show("Clean Up Unused Files complete." & vbCrLf & vbCrLf &
                                "Drawing folder IPT files: " & allIptFiles.Count.ToString() & vbCrLf &
                                "Referenced part files: " & referencedPartFileNames.Count.ToString() & vbCrLf &
                                "Unreferenced files found: " & unreferenced.Count.ToString() & vbCrLf &
                                "Moved to 'Unrenamed Parts': " & moved.ToString() & vbCrLf &
                                "Skipped (already existed): " & skippedExisting.ToString() & vbCrLf &
                                "Move errors: " & moveErrors.ToString(),
                                "Clean Up Unused Files",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
            Finally
                For Each doc As Document In openedByTool
                    Try
                        doc.Close(False)
                    Catch
                    End Try
                Next
            End Try
        End Sub

        Private Sub CollectReferencedPartFileNames(ByVal drawingDoc As DrawingDocument,
                                                   ByVal partFileNames As HashSet(Of String),
                                                   ByVal visitedAssemblies As HashSet(Of String),
                                                   ByVal openedByTool As List(Of Document))
            For Each sheet As Sheet In drawingDoc.Sheets
                For Each partsList As PartsList In sheet.PartsLists
                    For Each row As PartsListRow In partsList.PartsListRows
                        Dim rowPaths As List(Of String) = DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                        For Each pathValue As String In rowPaths
                            CollectFromModelPath(pathValue, partFileNames, visitedAssemblies, openedByTool)
                        Next
                    Next
                Next
            Next

            Try
                For Each descriptor As FileDescriptor In drawingDoc.File.ReferencedFileDescriptors
                    Dim fullPath As String = String.Empty
                    Try
                        fullPath = Convert.ToString(descriptor.FullFileName)
                    Catch
                        fullPath = String.Empty
                    End Try

                    CollectFromModelPath(fullPath, partFileNames, visitedAssemblies, openedByTool)
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "ReferencedFileDescriptors scan failed | " & ex.Message)
            End Try
        End Sub

        Private Sub CollectFromModelPath(ByVal modelPath As String,
                                         ByVal partFileNames As HashSet(Of String),
                                         ByVal visitedAssemblies As HashSet(Of String),
                                         ByVal openedByTool As List(Of Document))
            If String.IsNullOrWhiteSpace(modelPath) Then
                Return
            End If

            Dim extensionValue As String = String.Empty
            Try
                extensionValue = System.IO.Path.GetExtension(modelPath)
            Catch
                extensionValue = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(extensionValue) Then
                Return
            End If

            If extensionValue.Equals(".ipt", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim fileName As String = System.IO.Path.GetFileName(modelPath)
                    If Not String.IsNullOrWhiteSpace(fileName) Then
                        partFileNames.Add(fileName)
                    End If
                Catch
                End Try
                Return
            End If

            If extensionValue.Equals(".iam", StringComparison.OrdinalIgnoreCase) Then
                If IsSkippedAssemblyPath(modelPath) Then
                    Return
                End If
                ScanAssemblyForParts(modelPath, partFileNames, visitedAssemblies, openedByTool)
            End If
        End Sub

        Private Sub ScanAssemblyForParts(ByVal assemblyPath As String,
                                         ByVal partFileNames As HashSet(Of String),
                                         ByVal visitedAssemblies As HashSet(Of String),
                                         ByVal openedByTool As List(Of Document))
            If String.IsNullOrWhiteSpace(assemblyPath) Then
                Return
            End If

            If visitedAssemblies.Contains(assemblyPath) Then
                Return
            End If

            visitedAssemblies.Add(assemblyPath)

            Dim asmDoc As AssemblyDocument = Nothing
            Dim openedThisDoc As Boolean = False

            For Each openDoc As Document In m_InventorApp.Documents
                If String.Equals(openDoc.FullFileName, assemblyPath, StringComparison.OrdinalIgnoreCase) Then
                    asmDoc = TryCast(openDoc, AssemblyDocument)
                    Exit For
                End If
            Next

            If asmDoc Is Nothing Then
                Try
                    asmDoc = TryCast(m_InventorApp.Documents.Open(assemblyPath, False), AssemblyDocument)
                    openedThisDoc = asmDoc IsNot Nothing
                Catch ex As Exception
                    AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Could not open assembly | Path='" & assemblyPath & "' | " & ex.Message)
                    Return
                End Try
            End If

            If asmDoc Is Nothing OrElse asmDoc.ComponentDefinition Is Nothing Then
                Return
            End If

            If openedThisDoc Then
                openedByTool.Add(asmDoc)
            End If

            Try
                For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                    Dim isSuppressed As Boolean = False
                    Try
                        isSuppressed = occ.Suppressed
                    Catch
                        isSuppressed = False
                    End Try

                    If isSuppressed Then
                        Continue For
                    End If

                    Dim occPath As String = String.Empty
                    Try
                        occPath = occ.Definition.Document.FullFileName
                    Catch
                        occPath = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(occPath) Then
                        Continue For
                    End If

                    Dim ext As String = String.Empty
                    Try
                        ext = System.IO.Path.GetExtension(occPath)
                    Catch
                        ext = String.Empty
                    End Try

                    If ext.Equals(".ipt", StringComparison.OrdinalIgnoreCase) Then
                        Try
                            Dim fileName As String = System.IO.Path.GetFileName(occPath)
                            If Not String.IsNullOrWhiteSpace(fileName) Then
                                partFileNames.Add(fileName)
                            End If
                        Catch
                        End Try
                    ElseIf ext.Equals(".iam", StringComparison.OrdinalIgnoreCase) Then
                        If Not IsSkippedAssemblyPath(occPath) Then
                            ScanAssemblyForParts(occPath, partFileNames, visitedAssemblies, openedByTool)
                        End If
                    End If
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Assembly occurrence scan failed | Assembly='" & assemblyPath & "' | " & ex.Message)
            End Try
        End Sub

        Private Function IsSkippedAssemblyPath(ByVal assemblyPath As String) As Boolean
            If String.IsNullOrWhiteSpace(assemblyPath) Then
                Return True
            End If

            Dim lowerPath As String = assemblyPath.ToLowerInvariant()
            Dim lowerName As String = String.Empty
            Try
                lowerName = System.IO.Path.GetFileNameWithoutExtension(assemblyPath).ToLowerInvariant()
            Catch
                lowerName = lowerPath
            End Try

            If lowerName.Contains("bolted connection") Then
                Return True
            End If

            If lowerName.Contains("content center") Then
                Return True
            End If

            If lowerPath.Contains("\\content center\\") Then
                Return True
            End If

            Return False
        End Function

        Private Sub CollectAllIptFiles(ByVal rootFolder As String, ByVal result As List(Of String))
            If String.IsNullOrWhiteSpace(rootFolder) OrElse Not System.IO.Directory.Exists(rootFolder) Then
                Return
            End If

            Try
                For Each filePath As String In System.IO.Directory.GetFiles(rootFolder, "*.ipt")
                    Dim fileName As String = String.Empty
                    Try
                        fileName = System.IO.Path.GetFileName(filePath)
                    Catch
                        fileName = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(fileName) Then
                        Continue For
                    End If

                    If fileName.IndexOf("development", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Continue For
                    End If

                    result.Add(filePath)
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "File scan failed | Folder='" & rootFolder & "' | " & ex.Message)
            End Try

            Try
                For Each childFolder As String In System.IO.Directory.GetDirectories(rootFolder)
                    Dim folderName As String = String.Empty
                    Try
                        folderName = System.IO.Path.GetFileName(childFolder)
                    Catch
                        folderName = String.Empty
                    End Try

                    If String.Equals(folderName, "OldVersions", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    CollectAllIptFiles(childFolder, result)
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CleanUpUnusedFilesTool", "Directory scan failed | Folder='" & rootFolder & "' | " & ex.Message)
            End Try
        End Sub
    End Class

    Friend NotInheritable Class AddInDiagnostics
        Private Sub New()
        End Sub

        Public Shared Sub Log(ByVal context As String, ByVal message As String)
            Try
                Dim baseDir As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Spectiv", "InventorAutomationSuite", "Logs")
                If Not System.IO.Directory.Exists(baseDir) Then
                    System.IO.Directory.CreateDirectory(baseDir)
                End If

                Dim logPath As String = System.IO.Path.Combine(baseDir, "DrawingAutomationTools.log")
                Dim line As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & " | " & context & " | " & message
                System.IO.File.AppendAllText(logPath, line & System.Environment.NewLine)
            Catch
            End Try
        End Sub
    End Class

    Public Class CreateSheetPartsListTool
        Private ReadOnly m_InventorApp As Inventor.Application

        Private Class AssemblyCandidate
            Public Property SheetName As String
            Public Property SheetNumber As String
            Public Property AssemblyView As DrawingView
            Public Property AssemblyName As String

            Public ReadOnly Property Label As String
                Get
                    Return "Sheet: " & SheetName & " (" & SheetNumber & ") — " & AssemblyName
                End Get
            End Property
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running Create Sheet Parts List.")
            End If

            Dim targetSheet As Sheet = drawingDoc.ActiveSheet
            If targetSheet Is Nothing Then
                Throw New InvalidOperationException("No active sheet found.")
            End If

            Dim visiblePartNumbers As HashSet(Of String) = CollectVisiblePartNumbersOnSheet(targetSheet)
            AddInDiagnostics.Log("CreateSheetPartsListTool", "Start | Sheet='" & targetSheet.Name & "' | VisiblePartNumbers=" & visiblePartNumbers.Count.ToString())

            If visiblePartNumbers.Count = 0 Then
                MessageBox.Show("No components found on this sheet.", "Create Sheet Parts List", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim sourceAssemblyView As DrawingView = SelectAssemblySourceView(drawingDoc, targetSheet)
            If sourceAssemblyView Is Nothing Then
                AddInDiagnostics.Log("CreateSheetPartsListTool", "No assembly view selected/found | Sheet='" & targetSheet.Name & "'")
                MessageBox.Show("No assembly view found in drawing to create parts list from.", "Create Sheet Parts List", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            RemoveExistingPartsLists(targetSheet)
            Dim partsList As PartsList = TryAddPartsListFromAssemblyView(targetSheet, sourceAssemblyView)
            If partsList Is Nothing Then
                AddInDiagnostics.Log("CreateSheetPartsListTool", "FAILED | Could not create parts list from selected assembly view | Sheet='" & targetSheet.Name & "'")
                MessageBox.Show("Failed to create parts list on active sheet.", "Create Sheet Parts List", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                partsList.Title = "Parts List"
            Catch
            End Try

            FilterPartsListToVisiblePartNumbers(partsList, visiblePartNumbers)
            RenumberVisibleRows(partsList)

            Dim totalRows As Integer = 0
            Dim visibleRows As Integer = 0
            Try
                totalRows = partsList.PartsListRows.Count
                For Each row As PartsListRow In partsList.PartsListRows
                    Try
                        If row.Visible Then
                            visibleRows += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            AddInDiagnostics.Log("CreateSheetPartsListTool", "Success | Sheet='" & targetSheet.Name & "' | TotalRows=" & totalRows.ToString() & " | VisibleRows=" & visibleRows.ToString())

            drawingDoc.Update()
            drawingDoc.Save2(True)

            MessageBox.Show("Create Sheet Parts List complete." & vbCrLf & "Visible rows: " & visibleRows.ToString() & " / " & totalRows.ToString(), "Create Sheet Parts List", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function CollectVisiblePartNumbersOnSheet(ByVal sheet As Sheet) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each view As DrawingView In sheet.DrawingViews
                Dim modelPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                If String.IsNullOrWhiteSpace(modelPath) Then
                    Continue For
                End If

                If modelPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    Dim baseName As String = String.Empty
                    Try
                        baseName = System.IO.Path.GetFileNameWithoutExtension(modelPath)
                    Catch
                        baseName = String.Empty
                    End Try

                    If Not String.IsNullOrWhiteSpace(baseName) Then
                        result.Add(baseName.Trim().ToLowerInvariant())
                    End If

                ElseIf modelPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    Dim asmDoc As AssemblyDocument = Nothing
                    Try
                        asmDoc = TryCast(view.ReferencedDocumentDescriptor.ReferencedDocument, AssemblyDocument)
                    Catch
                        asmDoc = Nothing
                    End Try

                    If asmDoc IsNot Nothing Then
                        CollectLeafPartNumbersFromAssembly(asmDoc, result)
                    End If
                End If
            Next

            AddInDiagnostics.Log("CreateSheetPartsListTool", "CollectVisiblePartNumbersOnSheet | Sheet='" & sheet.Name & "' | Count=" & result.Count.ToString())
            Return result
        End Function

        Private Sub CollectLeafPartNumbersFromAssembly(ByVal assemblyDoc As AssemblyDocument,
                                                       ByVal numbers As HashSet(Of String))
            If assemblyDoc Is Nothing Then
                Return
            End If

            Try
                Dim compDef As AssemblyComponentDefinition = assemblyDoc.ComponentDefinition
                If compDef Is Nothing Then
                    Return
                End If

                Dim bom As BOM = compDef.BOM
                If bom Is Nothing Then
                    Return
                End If

                bom.StructuredViewEnabled = True
                bom.StructuredViewFirstLevelOnly = False

                Dim structuredView As BOMView = bom.BOMViews.Item("Structured")
                If structuredView Is Nothing OrElse structuredView.BOMRows Is Nothing Then
                    Return
                End If

                CollectLeafPartNumbersFromBomRows(structuredView.BOMRows, numbers)
            Catch ex As Exception
                AddInDiagnostics.Log("CreateSheetPartsListTool", "CollectLeafPartNumbersFromAssembly failed | Assembly='" & assemblyDoc.DisplayName & "' | " & ex.Message)
            End Try
        End Sub

        Private Sub CollectLeafPartNumbersFromBomRows(ByVal rows As BOMRowsEnumerator,
                                                      ByVal numbers As HashSet(Of String))
            If rows Is Nothing Then
                Return
            End If

            For i As Integer = 1 To rows.Count
                Try
                    Dim bomRow As BOMRow = rows.Item(i)
                    Dim hasChildren As Boolean = False

                    Try
                        hasChildren = bomRow.ChildRows IsNot Nothing AndAlso bomRow.ChildRows.Count > 0
                    Catch
                        hasChildren = False
                    End Try

                    If hasChildren Then
                        CollectLeafPartNumbersFromBomRows(bomRow.ChildRows, numbers)
                        Continue For
                    End If

                    Try
                        Dim compDefs As ComponentDefinitionsEnumerator = bomRow.ComponentDefinitions
                        If compDefs Is Nothing OrElse compDefs.Count = 0 Then
                            Continue For
                        End If

                        Dim compDef As ComponentDefinition = compDefs.Item(1)
                        If compDef Is Nothing OrElse compDef.Document Is Nothing Then
                            Continue For
                        End If

                        Dim fullPath As String = String.Empty
                        Try
                            fullPath = compDef.Document.FullFileName
                        Catch
                            fullPath = String.Empty
                        End Try

                        If String.IsNullOrWhiteSpace(fullPath) OrElse Not fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If

                        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(fullPath)
                        If Not String.IsNullOrWhiteSpace(baseName) Then
                            numbers.Add(baseName.Trim().ToLowerInvariant())
                        End If
                    Catch
                    End Try
                Catch
                End Try
            Next
        End Sub

        Private Function SelectAssemblySourceView(ByVal drawingDoc As DrawingDocument,
                                                  ByVal targetSheet As Sheet) As DrawingView
            Dim candidates As New List(Of AssemblyCandidate)()
            Dim seenSheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each sheet As Sheet In drawingDoc.Sheets
                Dim sheetNumber As String = "?"
                Try
                    sheetNumber = sheet.Index.ToString()
                Catch
                End Try

                For Each view As DrawingView In sheet.DrawingViews
                    Dim modelPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                    If String.IsNullOrWhiteSpace(modelPath) OrElse Not modelPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If seenSheets.Contains(sheet.Name) Then
                        Continue For
                    End If

                    Dim assemblyName As String = String.Empty
                    Try
                        assemblyName = view.ReferencedDocumentDescriptor.ReferencedDocument.DisplayName
                    Catch
                        Try
                            assemblyName = System.IO.Path.GetFileName(modelPath)
                        Catch
                            assemblyName = "Assembly"
                        End Try
                    End Try

                    seenSheets.Add(sheet.Name)
                    candidates.Add(New AssemblyCandidate() With {
                        .SheetName = sheet.Name,
                        .SheetNumber = sheetNumber,
                        .AssemblyView = view,
                        .AssemblyName = assemblyName
                    })
                    Exit For
                Next
            Next

            If candidates.Count = 0 Then
                Return Nothing
            End If

            If candidates.Count = 1 Then
                AddInDiagnostics.Log("CreateSheetPartsListTool", "Using only assembly view candidate | Sheet='" & candidates(0).SheetName & "' | Assembly='" & candidates(0).AssemblyName & "'")
                Return candidates(0).AssemblyView
            End If

            Dim labels As New List(Of String)()
            For Each candidate As AssemblyCandidate In candidates
                labels.Add(candidate.Label)
            Next

            Dim defaultLabel As String = labels(0)
            For Each candidate As AssemblyCandidate In candidates
                If String.Equals(candidate.SheetName, targetSheet.Name, StringComparison.OrdinalIgnoreCase) Then
                    defaultLabel = candidate.Label
                    Exit For
                End If
            Next

            Dim selected As String = DropdownSelectionForm.ShowDialogResult(
                "Create Sheet Parts List",
                "Select assembly source for parts list:",
                labels,
                defaultLabel)

            If String.IsNullOrWhiteSpace(selected) Then
                Return Nothing
            End If

            For Each candidate As AssemblyCandidate In candidates
                If String.Equals(candidate.Label, selected, StringComparison.OrdinalIgnoreCase) Then
                    AddInDiagnostics.Log("CreateSheetPartsListTool", "User selected assembly source | Sheet='" & candidate.SheetName & "' | Assembly='" & candidate.AssemblyName & "'")
                    Return candidate.AssemblyView
                End If
            Next

            Return Nothing
        End Function

        Private Function TryAddPartsListFromAssemblyView(ByVal sheet As Sheet,
                                                         ByVal assemblyView As DrawingView) As PartsList
            If sheet Is Nothing OrElse assemblyView Is Nothing Then
                Return Nothing
            End If

            Try
                Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
                Dim plPoint As Point2d = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)
                Dim listValue As PartsList = sheet.PartsLists.Add(assemblyView, plPoint)
                If listValue IsNot Nothing Then
                    AddInDiagnostics.Log("CreateSheetPartsListTool", "Parts list created from assembly view | Rows=" & listValue.PartsListRows.Count.ToString())
                End If
                Return listValue
            Catch ex As Exception
                AddInDiagnostics.Log("CreateSheetPartsListTool", "Parts list creation failed from assembly view | " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Sub RemoveExistingPartsLists(ByVal sheet As Sheet)
            For i As Integer = sheet.PartsLists.Count To 1 Step -1
                Try
                    sheet.PartsLists.Item(i).Delete()
                Catch
                End Try
            Next
        End Sub

        Private Function GetRowPartNumber(ByVal row As PartsListRow) As String
            Dim rowPartNum As String = String.Empty

            Try
                rowPartNum = Convert.ToString(row.Item(2).Value)
            Catch
                rowPartNum = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(rowPartNum) Then
                Try
                    rowPartNum = Convert.ToString(row.Item(1).Value)
                Catch
                    rowPartNum = String.Empty
                End Try
            End If

            Return If(rowPartNum, String.Empty).Trim().ToLowerInvariant()
        End Function

        Private Sub FilterPartsListToVisiblePartNumbers(ByVal partsList As PartsList,
                                                        ByVal visiblePartNumbers As HashSet(Of String))
            Dim hidden As Integer = 0
            Dim kept As Integer = 0

            For Each row As PartsListRow In partsList.PartsListRows
                Dim rowPartNum As String = GetRowPartNumber(row)
                If visiblePartNumbers.Contains(rowPartNum) Then
                    kept += 1
                Else
                    Try
                        row.Visible = False
                        hidden += 1
                    Catch
                    End Try
                End If
            Next

            AddInDiagnostics.Log("CreateSheetPartsListTool", "Filter complete | Kept=" & kept.ToString() & " | Hidden=" & hidden.ToString())
        End Sub

        Private Sub RenumberVisibleRows(ByVal partsList As PartsList)
            Dim itemNo As Integer = 1
            For Each row As PartsListRow In partsList.PartsListRows
                Dim isVisible As Boolean = True
                Try
                    isVisible = row.Visible
                Catch
                End Try

                If Not isVisible Then
                    Continue For
                End If

                Try
                    row.Item(1).Value = itemNo.ToString("00")
                Catch
                End Try

                itemNo += 1
            Next
        End Sub
    End Class

    Public Class CreateGAPartsListTopLevelTool
        Private ReadOnly m_InventorApp As Inventor.Application

        Private Class AssemblyCandidate
            Public Property SheetName As String
            Public Property SheetNumber As String
            Public Property AssemblyView As DrawingView
            Public Property AssemblyDoc As AssemblyDocument
            Public Property AssemblyName As String

            Public ReadOnly Property Label As String
                Get
                    Return "Sheet: " & SheetName & " (" & SheetNumber & ") — " & AssemblyName
                End Get
            End Property
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Dim drawingDoc As DrawingDocument = DrawingAutomationHelpers.GetActiveDrawing(m_InventorApp)
            If drawingDoc Is Nothing Then
                Throw New InvalidOperationException("Open an IDW drawing before running Create GA Parts List (Top Level).")
            End If

            Dim targetSheet As Sheet = drawingDoc.ActiveSheet
            If targetSheet Is Nothing Then
                Throw New InvalidOperationException("No active sheet found.")
            End If

            AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Start | Sheet='" & targetSheet.Name & "'")

            Dim selectedAssembly As AssemblyCandidate = SelectAssemblySourceCandidate(drawingDoc, targetSheet)
            If selectedAssembly Is Nothing OrElse selectedAssembly.AssemblyView Is Nothing OrElse selectedAssembly.AssemblyDoc Is Nothing Then
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "No assembly source selected/found | Sheet='" & targetSheet.Name & "'")
                MessageBox.Show("No assembly source selected or found.", "Create GA Parts List (Top Level)", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim allowedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim allowedKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            CollectTopLevelOccurrenceSets(selectedAssembly.AssemblyDoc, allowedPaths, allowedKeys)
            AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Assembly top-level context | Sheet='" & targetSheet.Name & "' | Paths=" & allowedPaths.Count.ToString() & " | Keys=" & allowedKeys.Count.ToString())

            If allowedPaths.Count = 0 AndAlso allowedKeys.Count = 0 Then
                MessageBox.Show("No components found in the selected assembly.", "Create GA Parts List (Top Level)", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            RemoveExistingPartsLists(targetSheet)
            Dim partsList As PartsList = TryAddStructuredFirstLevelPartsList(targetSheet, selectedAssembly.AssemblyView, selectedAssembly.AssemblyDoc)
            If partsList Is Nothing Then
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "FAILED | Could not create GA parts list | Sheet='" & targetSheet.Name & "' | Assembly='" & selectedAssembly.AssemblyName & "'")
                MessageBox.Show("Failed to create GA parts list on active sheet.", "Create GA Parts List (Top Level)", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                partsList.Title = "GA Parts List"
            Catch
            End Try

            ' Show all first-level components - no filtering needed
            RenumberVisibleRows(partsList)

            Dim totalRows As Integer = 0
            Dim visibleRows As Integer = 0
            Try
                totalRows = partsList.PartsListRows.Count
                For Each row As PartsListRow In partsList.PartsListRows
                    Try
                        If row.Visible Then
                            visibleRows += 1
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try

            AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Success | Sheet='" & targetSheet.Name & "' | Assembly='" & selectedAssembly.AssemblyName & "' | TotalRows=" & totalRows.ToString() & " | VisibleRows=" & visibleRows.ToString())

            drawingDoc.Update()
            drawingDoc.Save2(True)

            MessageBox.Show("Create GA Parts List (Top Level) complete." & vbCrLf & "Components shown: " & visibleRows.ToString() & " / " & totalRows.ToString(), "Create GA Parts List (Top Level)", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Sub CollectVisibleTopLevelContext(ByVal sheet As Sheet,
                                                  ByVal allowedPaths As HashSet(Of String),
                                                  ByVal allowedKeys As HashSet(Of String))
            For Each view As DrawingView In sheet.DrawingViews
                Dim modelPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                If String.IsNullOrWhiteSpace(modelPath) Then
                    Continue For
                End If

                allowedPaths.Add(modelPath)
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Collected visible model path | Path='" & modelPath & "'")

                Try
                    Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(modelPath)
                    If Not String.IsNullOrWhiteSpace(baseName) Then
                        allowedKeys.Add(baseName.Trim().ToLowerInvariant())
                        AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Collected visible model key | Key='" & baseName.Trim().ToLowerInvariant() & "'")
                    End If
                Catch
                End Try

                If modelPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        Dim partDoc As PartDocument = DrawingAutomationHelpers.OpenPartDocument(m_InventorApp, modelPath)
                        If partDoc IsNot Nothing Then
                            Dim designSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                            Dim partNo As String = Convert.ToString(designSet.Item("Part Number").Value)
                            If Not String.IsNullOrWhiteSpace(partNo) Then
                                allowedKeys.Add(partNo.Trim().ToLowerInvariant())
                                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Collected part number key | Key='" & partNo.Trim().ToLowerInvariant() & "'")
                            End If

                            Dim desc As String = Convert.ToString(designSet.Item("Description").Value)
                            If Not String.IsNullOrWhiteSpace(desc) Then
                                allowedKeys.Add(desc.Trim().ToLowerInvariant())
                                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Collected description key | Key='" & desc.Trim().ToLowerInvariant() & "'")
                            End If
                        End If
                    Catch
                    End Try
                End If
            Next
        End Sub

        Private Sub CollectDirectOccurrenceContext(ByVal assemblyDoc As AssemblyDocument,
                                                   ByVal allowedPaths As HashSet(Of String),
                                                   ByVal allowedKeys As HashSet(Of String))
            If assemblyDoc Is Nothing Then
                Return
            End If

            Try
                For Each occurrence As ComponentOccurrence In assemblyDoc.ComponentDefinition.Occurrences
                    Dim isSuppressed As Boolean = False
                    Try
                        isSuppressed = occurrence.Suppressed
                    Catch
                        isSuppressed = False
                    End Try

                    If isSuppressed Then
                        Continue For
                    End If

                    Dim pathValue As String = String.Empty
                    Try
                        pathValue = occurrence.Definition.Document.FullFileName
                    Catch
                        pathValue = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(pathValue) Then
                        Continue For
                    End If

                    allowedPaths.Add(pathValue)

                    Try
                        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(pathValue)
                        If Not String.IsNullOrWhiteSpace(baseName) Then
                            allowedKeys.Add(baseName.Trim().ToLowerInvariant())
                        End If
                    Catch
                    End Try

                    Try
                        Dim designSet As PropertySet = occurrence.Definition.Document.PropertySets.Item("Design Tracking Properties")
                        Dim partNo As String = Convert.ToString(designSet.Item("Part Number").Value)
                        If Not String.IsNullOrWhiteSpace(partNo) Then
                            allowedKeys.Add(partNo.Trim().ToLowerInvariant())
                        End If
                    Catch
                    End Try
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "CollectDirectOccurrenceContext failed | Assembly='" & assemblyDoc.DisplayName & "' | " & ex.Message)
            End Try
        End Sub

        Private Function SelectAssemblySourceCandidate(ByVal drawingDoc As DrawingDocument,
                                                       ByVal targetSheet As Sheet) As AssemblyCandidate
            Dim candidates As New List(Of AssemblyCandidate)()
            Dim seenPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each sheet As Sheet In drawingDoc.Sheets
                Dim sheetNumber As String = "?"
                Try
                    sheetNumber = sheet.Index.ToString()
                Catch
                End Try

                For Each view As DrawingView In sheet.DrawingViews
                    Dim modelPath As String = DrawingAutomationHelpers.ResolveViewModelPath(view)
                    If String.IsNullOrWhiteSpace(modelPath) OrElse Not modelPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If seenPaths.Contains(modelPath) Then
                        Continue For
                    End If

                    Dim asmDoc As AssemblyDocument = Nothing
                    Try
                        asmDoc = TryCast(view.ReferencedDocumentDescriptor.ReferencedDocument, AssemblyDocument)
                    Catch
                        asmDoc = Nothing
                    End Try

                    If asmDoc Is Nothing Then
                        Continue For
                    End If

                    seenPaths.Add(modelPath)
                    candidates.Add(New AssemblyCandidate() With {
                        .SheetName = sheet.Name,
                        .SheetNumber = sheetNumber,
                        .AssemblyView = view,
                        .AssemblyDoc = asmDoc,
                        .AssemblyName = asmDoc.DisplayName
                    })
                Next
            Next

            If candidates.Count = 0 Then
                Return Nothing
            End If

            If candidates.Count = 1 Then
                Return candidates(0)
            End If

            Dim labels As List(Of String) = candidates.Select(Function(c) c.Label).ToList()
            Dim defaultLabel As String = labels(0)
            For Each candidate As AssemblyCandidate In candidates
                If String.Equals(candidate.SheetName, targetSheet.Name, StringComparison.OrdinalIgnoreCase) Then
                    defaultLabel = candidate.Label
                    Exit For
                End If
            Next

            Dim selected As String = DropdownSelectionForm.ShowDialogResult(
                "Create GA Parts List (Top Level)",
                "Select assembly source for GA top-level parts list:",
                labels,
                defaultLabel)

            If String.IsNullOrWhiteSpace(selected) Then
                Return Nothing
            End If

            For Each candidate As AssemblyCandidate In candidates
                If String.Equals(candidate.Label, selected, StringComparison.OrdinalIgnoreCase) Then
                    Return candidate
                End If
            Next

            Return Nothing
        End Function

        Private Sub CollectTopLevelOccurrenceSets(ByVal assemblyDoc As AssemblyDocument,
                                                   ByVal topLevelPaths As HashSet(Of String),
                                                   ByVal topLevelKeys As HashSet(Of String))
            If assemblyDoc Is Nothing Then
                Return
            End If

            Try
                For Each occurrence As ComponentOccurrence In assemblyDoc.ComponentDefinition.Occurrences
                    Dim isSuppressed As Boolean = False
                    Try
                        isSuppressed = occurrence.Suppressed
                    Catch
                        isSuppressed = False
                    End Try

                    If isSuppressed Then
                        Continue For
                    End If

                    Dim pathValue As String = String.Empty
                    Try
                        pathValue = occurrence.Definition.Document.FullFileName
                    Catch
                        pathValue = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(pathValue) Then
                        Continue For
                    End If

                    topLevelPaths.Add(pathValue)

                    Try
                        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(pathValue)
                        If Not String.IsNullOrWhiteSpace(baseName) Then
                            topLevelKeys.Add(baseName.Trim().ToLowerInvariant())
                        End If
                    Catch
                    End Try

                    Try
                        Dim designSet As PropertySet = occurrence.Definition.Document.PropertySets.Item("Design Tracking Properties")
                        Dim partNo As String = Convert.ToString(designSet.Item("Part Number").Value)
                        If Not String.IsNullOrWhiteSpace(partNo) Then
                            topLevelKeys.Add(partNo.Trim().ToLowerInvariant())
                        End If
                    Catch
                    End Try
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "CollectTopLevelOccurrenceSets failed | Assembly='" & assemblyDoc.DisplayName & "' | " & ex.Message)
            End Try
        End Sub

        Private Function TryAddStructuredFirstLevelPartsList(ByVal sheet As Sheet,
                                                              ByVal assemblyView As DrawingView,
                                                              ByVal assemblyDoc As AssemblyDocument) As PartsList
            If sheet Is Nothing OrElse assemblyView Is Nothing OrElse assemblyDoc Is Nothing Then
                Return Nothing
            End If

            Try
                Dim bom As BOM = assemblyDoc.ComponentDefinition.BOM
                bom.StructuredViewEnabled = True
                bom.StructuredViewFirstLevelOnly = True
                Try
                    bom.PartsOnlyViewEnabled = True
                Catch
                End Try
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Configured assembly BOM for structured first level | Assembly='" & assemblyDoc.DisplayName & "'")
            Catch ex As Exception
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "BOM configuration failed | " & ex.Message)
            End Try

            Try
                Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
                Dim plPoint As Point2d = tg.CreatePoint2d(sheet.Width - 40, sheet.Height - 40)
                Dim listValue As PartsList = Nothing

                Try
                    listValue = sheet.PartsLists.Add(assemblyView,
                                                     plPoint,
                                                     PartsListLevelEnum.kFirstLevelComponents)
                    AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Created parts list with level kFirstLevelComponents")
                Catch ex As Exception
                    AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "kFirstLevelComponents failed, retrying default level | " & ex.Message)
                    listValue = sheet.PartsLists.Add(assemblyView, plPoint)
                End Try

                Return listValue
            Catch ex As Exception
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Parts list creation failed | " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Sub FilterPartsListByVisibleTopLevelContext(ByVal partsList As PartsList,
                                                            ByVal allowedPaths As HashSet(Of String),
                                                            ByVal allowedKeys As HashSet(Of String))
            Dim kept As Integer = 0
            Dim hidden As Integer = 0

            For Each row As PartsListRow In partsList.PartsListRows
                Dim keepRow As Boolean = False

                Dim rowPaths As List(Of String) = DrawingAutomationHelpers.GetPathsFromPartsListRow(row)
                AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Checking row | RowPaths=" & String.Join(", ", rowPaths) & " | C1='" & Convert.ToString(row.Item(1).Value) & "' | C2='" & Convert.ToString(row.Item(2).Value) & "'")

                For Each rowPath As String In rowPaths
                    If String.IsNullOrWhiteSpace(rowPath) Then
                        Continue For
                    End If

                    If allowedPaths.Contains(rowPath) Then
                        keepRow = True
                        AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Row kept by path match | RowPath='" & rowPath & "'")
                        Exit For
                    End If
                Next

                If Not keepRow Then
                    Dim rowName As String = String.Empty
                    Try
                        rowName = Convert.ToString(row.Item(2).Value)
                    Catch
                        rowName = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(rowName) Then
                        Try
                            rowName = Convert.ToString(row.Item(1).Value)
                        Catch
                            rowName = String.Empty
                        End Try
                    End If

                    rowName = If(rowName, String.Empty).Trim().ToLowerInvariant()
                    If Not String.IsNullOrWhiteSpace(rowName) AndAlso allowedKeys.Contains(rowName) Then
                        keepRow = True
                        AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Row kept by key match | RowName='" & rowName & "'")
                    End If
                End If

                If keepRow Then
                    kept += 1
                Else
                    Try
                        row.Visible = False
                        hidden += 1
                        AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Row hidden | No match found")
                    Catch
                    End Try
                End If
            Next

            AddInDiagnostics.Log("CreateGAPartsListTopLevelTool", "Filter complete | Kept=" & kept.ToString() & " | Hidden=" & hidden.ToString())
        End Sub

        Private Sub RemoveExistingPartsLists(ByVal sheet As Sheet)
            For i As Integer = sheet.PartsLists.Count To 1 Step -1
                Try
                    sheet.PartsLists.Item(i).Delete()
                Catch
                End Try
            Next
        End Sub

        Private Sub RenumberVisibleRows(ByVal partsList As PartsList)
            Dim itemNo As Integer = 1
            For Each row As PartsListRow In partsList.PartsListRows
                Dim isVisible As Boolean = True
                Try
                    isVisible = row.Visible
                Catch
                End Try

                If Not isVisible Then
                    Continue For
                End If

                Try
                    row.Item(1).Value = itemNo.ToString("00")
                Catch
                End Try

                itemNo += 1
            Next
        End Sub
    End Class

    Friend NotInheritable Class DrawingAutomationHelpers
        Private Const SheetMetalSubType As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

        Public Shared Function GetActiveDrawing(ByVal inventorApp As Inventor.Application) As DrawingDocument
            If inventorApp Is Nothing OrElse inventorApp.ActiveDocument Is Nothing Then
                Return Nothing
            End If

            If inventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Return Nothing
            End If

            Return TryCast(inventorApp.ActiveDocument, DrawingDocument)
        End Function

        Public Shared Function GetOpenAssembly(ByVal inventorApp As Inventor.Application) As AssemblyDocument
            If inventorApp Is Nothing Then
                Return Nothing
            End If

            If inventorApp.ActiveDocument IsNot Nothing AndAlso inventorApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Return CType(inventorApp.ActiveDocument, AssemblyDocument)
            End If

            For Each documentValue As Document In inventorApp.Documents
                If documentValue.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return CType(documentValue, AssemblyDocument)
                End If
            Next

            Return Nothing
        End Function

        Public Shared Function IsDxfSheet(ByVal sheetName As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(sheetName) AndAlso sheetName.IndexOf("DXF", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Public Shared Function SelectNonDxfSheet(ByVal drawingDoc As DrawingDocument,
                                                 ByVal titleText As String,
                                                 ByVal promptText As String) As Sheet
            Dim names As New List(Of String)()
            For Each sheet As Sheet In drawingDoc.Sheets
                If Not IsDxfSheet(sheet.Name) Then
                    names.Add(sheet.Name)
                End If
            Next

            If names.Count = 0 Then
                Return Nothing
            End If

            Dim selectedName As String = DropdownSelectionForm.ShowDialogResult(titleText, promptText, names, names(0))
            If String.IsNullOrWhiteSpace(selectedName) Then
                Return Nothing
            End If

            For Each sheet As Sheet In drawingDoc.Sheets
                If String.Equals(sheet.Name, selectedName, StringComparison.OrdinalIgnoreCase) Then
                    Return sheet
                End If
            Next

            Return Nothing
        End Function

        Public Shared Function GetSheetNumber(ByVal sheet As Sheet) As Integer
            If sheet Is Nothing Then
                Return 0
            End If

            ' First try to parse from sheet name (e.g., "Sheet:1" -> 1, "AUTO PARTS 2" -> 2)
            Try
                Dim sheetName As String = sheet.Name.Trim()
                If sheetName.Contains(":") Then
                    ' Handle "Sheet:1" format
                    Dim parts As String() = sheetName.Split(":"c)
                    If parts.Length >= 2 Then
                        Dim numberPart As String = parts(1).Trim()
                        Dim parsed As Integer = 0
                        If Integer.TryParse(numberPart, parsed) AndAlso parsed > 0 Then
                            Return parsed
                        End If
                    End If
                End If

                ' Handle other formats by extracting the last number
                Dim digits As String = New String(sheetName.Where(Function(ch) Char.IsDigit(ch)).ToArray())
                If Not String.IsNullOrWhiteSpace(digits) Then
                    Dim parsed As Integer = 0
                    If Integer.TryParse(digits, parsed) AndAlso parsed > 0 Then
                        Return parsed
                    End If
                End If
            Catch
            End Try

            ' Fallback to sheet index (1-based position in collection)
            Try
                Dim indexValue As Integer = CInt(sheet.Index)
                If indexValue > 0 Then
                    Return indexValue
                End If
            Catch
            End Try

            Return 1
        End Function

        Public Shared Function ResolveViewModelPath(ByVal view As DrawingView) As String
            Try
                Dim descriptor As DocumentDescriptor = view.ReferencedDocumentDescriptor
                If descriptor IsNot Nothing Then
                    If descriptor.ReferencedDocument IsNot Nothing Then
                        Return descriptor.ReferencedDocument.FullFileName
                    End If

                    Try
                        Dim fullName As String = Convert.ToString(descriptor.FullDocumentName)
                        If Not String.IsNullOrWhiteSpace(fullName) Then
                            Return fullName
                        End If
                    Catch
                    End Try

                    Try
                        Dim fileDesc As FileDescriptor = descriptor.ReferencedFileDescriptor
                        If fileDesc IsNot Nothing Then
                            Dim fileName As String = Convert.ToString(fileDesc.FullFileName)
                            If Not String.IsNullOrWhiteSpace(fileName) Then
                                Return fileName
                            End If
                        End If
                    Catch
                    End Try
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Public Shared Function GetPathsFromPartsListRow(ByVal row As PartsListRow) As List(Of String)
            Dim values As New List(Of String)()

            Try
                For Each ref As Object In row.ReferencedFiles
                    Dim pathValue As String = String.Empty

                    Try
                        pathValue = CStr(ref.FullFileName)
                    Catch
                        pathValue = String.Empty
                    End Try

                    If String.IsNullOrWhiteSpace(pathValue) Then
                        Try
                            pathValue = CStr(ref.ReferencedDocument.FullFileName)
                        Catch
                            pathValue = String.Empty
                        End Try
                    End If

                    If Not String.IsNullOrWhiteSpace(pathValue) Then
                        values.Add(pathValue)
                    End If
                Next
            Catch
            End Try

            Return values
        End Function

        Public Shared Function TrySetModelDwgRef(ByVal inventorApp As Inventor.Application,
                                                 ByVal modelPath As String,
                                                 ByVal refs As String,
                                                 ByRef errorCount As Integer) As Boolean
            Dim modelDoc As Document = Nothing
            Dim openedByTool As Boolean = False

            For Each openDoc As Document In inventorApp.Documents
                If String.Equals(openDoc.FullFileName, modelPath, StringComparison.OrdinalIgnoreCase) Then
                    modelDoc = openDoc
                    Exit For
                End If
            Next

            Try
                If modelDoc Is Nothing Then
                    modelDoc = inventorApp.Documents.Open(modelPath, False)
                    openedByTool = True
                End If
            Catch ex As Exception
                AddInDiagnostics.Log("TrySetModelDwgRef", "Failed to open model | Path='" & modelPath & "' | " & ex.Message)
                modelDoc = Nothing
            End Try

            If modelDoc Is Nothing Then
                errorCount += 1
                Return False
            End If

            Try
                SetUserProperty(modelDoc, "DWG REF", refs)
                SetUserProperty(modelDoc, "DWG. REF.", refs)
                SetUserProperty(modelDoc, "DWG_REF", refs)
                SetUserProperty(modelDoc, "DWGREF", refs)
                modelDoc.Save2(True)

                If openedByTool Then
                    Try
                        modelDoc.Close(True)
                    Catch ex As Exception
                        AddInDiagnostics.Log("TrySetModelDwgRef", "Failed to close model | Path='" & modelPath & "' | " & ex.Message)
                    End Try
                End If

                Return True
            Catch ex As Exception
                AddInDiagnostics.Log("TrySetModelDwgRef", "Failed to set DWG REF | Path='" & modelPath & "' | Refs='" & refs & "' | " & ex.Message)
                If openedByTool AndAlso modelDoc IsNot Nothing Then
                    Try
                        modelDoc.Close(False)
                    Catch ex2 As Exception
                        AddInDiagnostics.Log("TrySetModelDwgRef", "Failed to close model after error | Path='" & modelPath & "' | " & ex2.Message)
                    End Try
                End If
                errorCount += 1
                Return False
            End Try
        End Function

        Public Shared Sub SetDrawingUserProperty(ByVal drawingDoc As DrawingDocument, ByVal name As String, ByVal value As String)
            SetUserProperty(drawingDoc, name, value)
        End Sub

        Private Shared Sub SetUserProperty(ByVal doc As Document, ByVal propertyName As String, ByVal propertyValue As String)
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then
                Return
            End If

            Try
                Dim userSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                If userSet Is Nothing Then
                    AddInDiagnostics.Log("SetUserProperty", "User Defined Properties set not found | Doc='" & doc.FullFileName & "'")
                    Return
                End If

                ' Clean the property value to avoid issues
                Dim cleanValue As String = If(propertyValue, String.Empty).Trim()
                If cleanValue.Length > 255 Then
                    cleanValue = cleanValue.Substring(0, 255) ' Limit property value length
                    AddInDiagnostics.Log("SetUserProperty", "Property value truncated to 255 chars | Property='" & propertyName & "' | OriginalLength=" & propertyValue.Length.ToString())
                End If

                Try
                    Dim [property] As Inventor.Property = userSet.Item(propertyName)
                    [property].Value = cleanValue
                    AddInDiagnostics.Log("SetUserProperty", "Updated existing property | Property='" & propertyName & "' | Value='" & cleanValue & "'")
                Catch propEx As Exception
                    ' Property doesn't exist, try to add it
                    Try
                        userSet.Add(cleanValue, propertyName)
                        AddInDiagnostics.Log("SetUserProperty", "Added new property | Property='" & propertyName & "' | Value='" & cleanValue & "'")
                    Catch addEx As Exception
                        AddInDiagnostics.Log("SetUserProperty", "Failed to add property | Property='" & propertyName & "' | Value='" & cleanValue & "' | " & addEx.Message)
                    End Try
                End Try
            Catch ex As Exception
                AddInDiagnostics.Log("SetUserProperty", "Failed to access user properties | Doc='" & doc.FullFileName & "' | Property='" & propertyName & "' | " & ex.Message)
            End Try
        End Sub

        Public Shared Function OpenPartDocument(ByVal inventorApp As Inventor.Application, ByVal partPath As String) As PartDocument
            If String.IsNullOrWhiteSpace(partPath) OrElse Not partPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                Return Nothing
            End If

            For Each openDoc As Document In inventorApp.Documents
                If String.Equals(openDoc.FullFileName, partPath, StringComparison.OrdinalIgnoreCase) Then
                    Return TryCast(openDoc, PartDocument)
                End If
            Next

            Try
                Return CType(inventorApp.Documents.Open(partPath, False), PartDocument)
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Function TryAddPartsList(ByVal sheet As Sheet,
                                               ByVal anchorView As DrawingView,
                                               ByVal titleText As String,
                                               ByVal inventorApp As Inventor.Application,
                                               Optional ByVal preferPartsOnly As Boolean = False) As PartsList
            If sheet Is Nothing OrElse anchorView Is Nothing Then
                Return Nothing
            End If

            Dim tg As TransientGeometry = inventorApp.TransientGeometry
            Dim points As New List(Of Point2d) From {
                tg.CreatePoint2d(sheet.Width - 10, sheet.Height - 10),
                tg.CreatePoint2d(sheet.Width - 20, sheet.Height - 20),
                tg.CreatePoint2d(sheet.Width * 0.8, sheet.Height * 0.8)
            }

            For Each pointValue As Point2d In points
                Try
                    Dim partsList As PartsList = Nothing
                    Dim createdPartsOnly As Boolean = False

                    If preferPartsOnly Then
                        Try
                            partsList = sheet.PartsLists.Add(anchorView,
                                                             pointValue,
                                                             PartsListLevelEnum.kPartsOnly)
                            createdPartsOnly = partsList IsNot Nothing
                        Catch ex As Exception
                            AddInDiagnostics.Log("TryAddPartsList", "kPartsOnly failed, falling back to default level | " & ex.Message)
                            partsList = sheet.PartsLists.Add(anchorView, pointValue)
                        End Try
                    Else
                        partsList = sheet.PartsLists.Add(anchorView, pointValue)
                    End If

                    If createdPartsOnly AndAlso partsList IsNot Nothing Then
                        Dim rowCount As Integer = 0
                        Try
                            rowCount = partsList.PartsListRows.Count
                        Catch
                            rowCount = 0
                        End Try

                        If rowCount = 0 Then
                            AddInDiagnostics.Log("TryAddPartsList", "kPartsOnly returned 0 rows; recreating list at default level for sub-assembly-only case")
                            Try
                                partsList.Delete()
                            Catch
                            End Try

                            Try
                                partsList = sheet.PartsLists.Add(anchorView, pointValue)
                            Catch ex As Exception
                                AddInDiagnostics.Log("TryAddPartsList", "Default-level recreate failed after empty kPartsOnly | " & ex.Message)
                                partsList = Nothing
                            End Try
                        End If
                    End If

                    If partsList IsNot Nothing Then
                        Try
                            partsList.Title = titleText
                        Catch
                        End Try
                        Return partsList
                    End If
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Public Shared Function IsPlatePart(ByVal partDoc As PartDocument) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            Dim fileBase As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName).ToUpperInvariant()
            Dim partNo As String = String.Empty
            Dim description As String = String.Empty

            Try
                Dim designSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                partNo = Convert.ToString(designSet.Item("Part Number").Value).ToUpperInvariant()
                description = Convert.ToString(designSet.Item("Description").Value).ToUpperInvariant()
            Catch
            End Try

            If String.Equals(partDoc.SubType, SheetMetalSubType, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            If fileBase.Contains("-PL") OrElse fileBase.Contains("-LPL") OrElse fileBase.StartsWith("PL") Then
                Return True
            End If

            If partNo.Contains("-PL") OrElse partNo.Contains("-LPL") OrElse partNo.StartsWith("PL") Then
                Return True
            End If

            Return description.Contains("PLATE")
        End Function

        Public Shared Function CreateSiblingSheet(ByVal drawingDoc As DrawingDocument, ByVal referenceSheet As Sheet, ByVal requestedName As String) As Sheet
            Dim sheetValue As Sheet = Nothing

            Try
                sheetValue = referenceSheet.CopyTo(drawingDoc)
            Catch
                sheetValue = Nothing
            End Try

            If sheetValue Is Nothing Then
                sheetValue = drawingDoc.Sheets.Add(referenceSheet.Size)
                CopySheetFrame(sheetValue, referenceSheet)
            Else
                ClearCopiedDrawingContent(sheetValue)
            End If

            sheetValue.Name = MakeUniqueSheetName(drawingDoc, requestedName)
            Return sheetValue
        End Function

        Private Shared Sub ClearCopiedDrawingContent(ByVal sheet As Sheet)
            If sheet Is Nothing Then
                Return
            End If

            Try
                For i As Integer = sheet.PartsLists.Count To 1 Step -1
                    sheet.PartsLists.Item(i).Delete()
                Next
            Catch
            End Try

            Try
                For i As Integer = sheet.DrawingViews.Count To 1 Step -1
                    sheet.DrawingViews.Item(i).Delete()
                Next
            Catch
            End Try
        End Sub

        Private Shared Sub CopySheetFrame(ByVal targetSheet As Sheet, ByVal referenceSheet As Sheet)
            If targetSheet Is Nothing OrElse referenceSheet Is Nothing Then
                Return
            End If

            Try
                If targetSheet.Border IsNot Nothing Then
                    targetSheet.Border.Delete()
                End If
            Catch
            End Try

            Try
                If referenceSheet.Border IsNot Nothing Then
                    targetSheet.AddBorder(referenceSheet.Border.Definition)
                End If
            Catch
            End Try

            Try
                If targetSheet.TitleBlock IsNot Nothing Then
                    targetSheet.TitleBlock.Delete()
                End If
            Catch
            End Try

            Try
                If referenceSheet.TitleBlock IsNot Nothing Then
                    targetSheet.AddTitleBlock(referenceSheet.TitleBlock.Definition)
                End If
            Catch
            End Try
        End Sub

        Private Shared Function MakeUniqueSheetName(ByVal drawingDoc As DrawingDocument, ByVal requestedName As String) As String
            Dim candidate As String = requestedName
            Dim index As Integer = 2

            While SheetExists(drawingDoc, candidate)
                candidate = requestedName & " (" & index.ToString() & ")"
                index += 1
            End While

            Return candidate
        End Function

        Private Shared Function SheetExists(ByVal drawingDoc As DrawingDocument, ByVal name As String) As Boolean
            For Each sheet As Sheet In drawingDoc.Sheets
                If String.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Shared Function GetGridPoint(ByVal inventorApp As Inventor.Application,
                                            ByVal sheet As Sheet,
                                            ByVal slot As Integer,
                                            ByVal columns As Integer,
                                            ByVal rows As Integer) As Point2d
            Dim margin As Double = 10.0
            Dim cellWidth As Double = (sheet.Width - (2 * margin)) / columns
            Dim cellHeight As Double = (sheet.Height - (2 * margin)) / rows

            Dim col As Integer = slot Mod columns
            Dim row As Integer = slot \ columns

            Dim x As Double = margin + ((col + 0.5) * cellWidth)
            Dim y As Double = sheet.Height - (margin + ((row + 0.5) * cellHeight))

            Return inventorApp.TransientGeometry.CreatePoint2d(x, y)
        End Function

        Public Shared Function AddPartView(ByVal sheet As Sheet,
                                           ByVal partDoc As PartDocument,
                                           ByVal pointValue As Point2d,
                                           ByVal scale As Double,
                                           ByVal folded As Boolean) As DrawingView
            Try
                Dim options As NameValueMap = Nothing
                If String.Equals(partDoc.SubType, SheetMetalSubType, StringComparison.OrdinalIgnoreCase) Then
                    options = CType(sheet.Parent.Parent.TransientObjects.CreateNameValueMap(), NameValueMap)
                    options.Add("SheetMetalFoldedModel", folded)
                End If

                If options Is Nothing Then
                    Return sheet.DrawingViews.AddBaseView(partDoc,
                                                          pointValue,
                                                          scale,
                                                          ViewOrientationTypeEnum.kFrontViewOrientation,
                                                          DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                End If

                Return sheet.DrawingViews.AddBaseView(partDoc,
                                                      pointValue,
                                                      scale,
                                                      ViewOrientationTypeEnum.kFrontViewOrientation,
                                                      DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle,
                                                      ,
                                                      ,
                                                      options)
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Function TryAddAssemblyPartsList(ByVal sheet As Sheet,
                                                       ByVal assemblyDoc As AssemblyDocument,
                                                       ByVal titleText As String,
                                                       ByVal inventorApp As Inventor.Application) As PartsList
            If sheet Is Nothing OrElse assemblyDoc Is Nothing Then
                AddInDiagnostics.Log("TryAddAssemblyPartsList", "Invalid input | SheetIsNothing=" & (sheet Is Nothing).ToString() & " | AssemblyIsNothing=" & (assemblyDoc Is Nothing).ToString())
                Return Nothing
            End If

            Try
                ConfigureAssemblyBomForList(assemblyDoc, False)
                AddInDiagnostics.Log("TryAddAssemblyPartsList", "BOM configured for structured all-levels | Assembly='" & assemblyDoc.DisplayName & "'")
            Catch ex As Exception
                AddInDiagnostics.Log("TryAddAssemblyPartsList", "BOM configuration failed | " & ex.Message)
            End Try

            Dim tg As TransientGeometry = inventorApp.TransientGeometry
            Dim tempPoints As New List(Of Point2d) From {
                tg.CreatePoint2d(sheet.Width - 6, 6),
                tg.CreatePoint2d(sheet.Width - 12, 12)
            }

            Dim requiresPartRows As Boolean = SheetHasPartViews(sheet)
            AddInDiagnostics.Log("TryAddAssemblyPartsList", "Row expectations | RequiresPartRows=" & requiresPartRows.ToString() & " | Sheet='" & sheet.Name & "'")

            ' First try creating the list directly from model document (more reliable for some assemblies)
            For Each modelPoint As Point2d In tempPoints
                Dim modelList As PartsList = Nothing
                Try
                    AddInDiagnostics.Log("TryAddAssemblyPartsList", "Attempting direct model list at X=" & modelPoint.X.ToString("0.###") & " Y=" & modelPoint.Y.ToString("0.###") & " | Sheet='" & sheet.Name & "'")
                    modelList = TryAddPartsListFromModel(sheet, assemblyDoc, modelPoint, titleText, True)
                    If modelList Is Nothing Then
                        Continue For
                    End If

                    If requiresPartRows AndAlso Not PartsListHasIptRows(modelList) Then
                        AddInDiagnostics.Log("TryAddAssemblyPartsList", "Rejecting direct model .iam-only list on PARTS sheet; retrying | Rows=" & modelList.PartsListRows.Count.ToString() & " | Sheet='" & sheet.Name & "'")
                        Try
                            modelList.Delete()
                        Catch
                        End Try
                        modelList = Nothing
                        Continue For
                    End If

                    AddInDiagnostics.Log("TryAddAssemblyPartsList", "Parts list created successfully (direct model) | Rows=" & modelList.PartsListRows.Count.ToString() & " | Sheet='" & sheet.Name & "'")
                    Return modelList
                Catch ex As Exception
                    AddInDiagnostics.Log("TryAddAssemblyPartsList", "Direct model attempt failed | " & ex.Message)
                    If modelList IsNot Nothing Then
                        Try
                            modelList.Delete()
                        Catch
                        End Try
                    End If
                End Try
            Next

            For Each tempPoint As Point2d In tempPoints
                Try
                    AddInDiagnostics.Log("TryAddAssemblyPartsList", "Attempting temp view at X=" & tempPoint.X.ToString("0.###") & " Y=" & tempPoint.Y.ToString("0.###") & " | Sheet='" & sheet.Name & "'")
                    Dim tempView As DrawingView = sheet.DrawingViews.AddBaseView(
                        assemblyDoc,
                        tempPoint,
                        0.001,
                        ViewOrientationTypeEnum.kFrontViewOrientation,
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)

                    If tempView Is Nothing Then
                        AddInDiagnostics.Log("TryAddAssemblyPartsList", "Temp view creation returned Nothing")
                        Continue For
                    End If

                    Try
                        tempView.ShowLabel = False
                    Catch
                    End Try

                    Dim listValue As PartsList = TryAddPartsList(sheet, tempView, titleText, inventorApp, True)
                    If listValue IsNot Nothing Then
                        If requiresPartRows AndAlso Not PartsListHasIptRows(listValue) Then
                            AddInDiagnostics.Log("TryAddAssemblyPartsList", "Rejecting .iam-only list on PARTS sheet; retrying | Rows=" & listValue.PartsListRows.Count.ToString() & " | Sheet='" & sheet.Name & "'")
                            Try
                                listValue.Delete()
                            Catch
                            End Try

                            ConfigureAssemblyBomForList(assemblyDoc, False)
                            Continue For
                        End If

                        AddInDiagnostics.Log("TryAddAssemblyPartsList", "Parts list created successfully | Rows=" & listValue.PartsListRows.Count.ToString() & " | Sheet='" & sheet.Name & "'")
                        Return listValue
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("TryAddAssemblyPartsList", "Attempt failed | " & ex.Message)
                End Try
            Next

            AddInDiagnostics.Log("TryAddAssemblyPartsList", "All attempts failed to create parts list | Sheet='" & sheet.Name & "'")
            Return Nothing
        End Function

        Private Shared Function TryAddPartsListFromModel(ByVal sheet As Sheet,
                                                         ByVal assemblyDoc As AssemblyDocument,
                                                         ByVal pointValue As Point2d,
                                                         ByVal titleText As String,
                                                         ByVal preferPartsOnly As Boolean) As PartsList
            If sheet Is Nothing OrElse assemblyDoc Is Nothing OrElse pointValue Is Nothing Then
                Return Nothing
            End If

            Dim partsList As PartsList = Nothing
            Dim createdPartsOnly As Boolean = False

            Try
                If preferPartsOnly Then
                    Try
                        partsList = sheet.PartsLists.Add(assemblyDoc,
                                                         pointValue,
                                                         PartsListLevelEnum.kPartsOnly)
                        createdPartsOnly = partsList IsNot Nothing
                    Catch ex As Exception
                        AddInDiagnostics.Log("TryAddPartsListFromModel", "kPartsOnly failed, falling back to default level | " & ex.Message)
                        partsList = sheet.PartsLists.Add(assemblyDoc, pointValue)
                    End Try
                Else
                    partsList = sheet.PartsLists.Add(assemblyDoc, pointValue)
                End If

                If createdPartsOnly AndAlso partsList IsNot Nothing Then
                    Dim rowCount As Integer = 0
                    Try
                        rowCount = partsList.PartsListRows.Count
                    Catch
                        rowCount = 0
                    End Try

                    If rowCount = 0 Then
                        AddInDiagnostics.Log("TryAddPartsListFromModel", "kPartsOnly returned 0 rows; recreating list at default level for sub-assembly-only case")
                        Try
                            partsList.Delete()
                        Catch
                        End Try

                        partsList = sheet.PartsLists.Add(assemblyDoc, pointValue)
                    End If
                End If

                If partsList IsNot Nothing Then
                    Try
                        partsList.Title = titleText
                    Catch
                    End Try
                End If

                Return partsList
            Catch ex As Exception
                AddInDiagnostics.Log("TryAddPartsListFromModel", "Model-based list creation failed | " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Function SheetHasPartViews(ByVal sheet As Sheet) As Boolean
            If sheet Is Nothing Then
                Return False
            End If

            Try
                For Each view As DrawingView In sheet.DrawingViews
                    Dim pathValue As String = ResolveViewModelPath(view)
                    If Not String.IsNullOrWhiteSpace(pathValue) AndAlso pathValue.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function PartsListHasIptRows(ByVal partsList As PartsList) As Boolean
            If partsList Is Nothing Then
                Return False
            End If

            Try
                For Each row As PartsListRow In partsList.PartsListRows
                    For Each pathValue As String In GetPathsFromPartsListRow(row)
                        If Not String.IsNullOrWhiteSpace(pathValue) AndAlso pathValue.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    Next
                Next
            Catch
            End Try

            Return False
        End Function

        Public Shared Sub ConfigureAssemblyBomForList(ByVal assemblyDoc As AssemblyDocument,
                                                      ByVal firstLevelOnly As Boolean)
            If assemblyDoc Is Nothing Then
                Return
            End If

            Try
                Dim bom As BOM = assemblyDoc.ComponentDefinition.BOM
                If bom IsNot Nothing Then
                    bom.StructuredViewEnabled = True
                    bom.StructuredViewFirstLevelOnly = firstLevelOnly
                    Try
                        bom.PartsOnlyViewEnabled = True
                    Catch
                    End Try
                    AddInDiagnostics.Log("ConfigureAssemblyBomForList", "Configured BOM | FirstLevelOnly=" & firstLevelOnly.ToString() & " | PartsOnlyEnabled=True | Assembly='" & assemblyDoc.DisplayName & "'")
                End If
            Catch
            End Try
        End Sub
    End Class

    Public Class DropdownSelectionForm
        Inherits Form

        Private ReadOnly m_Combo As ComboBox
        Private ReadOnly m_OkButton As Button
        Private ReadOnly m_CancelButton As Button

        Public Shared Function ShowDialogResult(ByVal titleText As String,
                                                ByVal promptText As String,
                                                ByVal items As List(Of String),
                                                ByVal defaultValue As String) As String
            Using form As New DropdownSelectionForm(titleText, promptText, items, defaultValue)
                If form.ShowDialog() = DialogResult.OK Then
                    Return form.SelectedValue()
                End If
            End Using

            Return String.Empty
        End Function

        Public Sub New(ByVal titleText As String,
                       ByVal promptText As String,
                       ByVal items As List(Of String),
                       ByVal defaultValue As String)
            Me.Text = titleText
            Me.Width = 560
            Me.Height = 180
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False

            Dim prompt As New Label() With {
                .Left = 14,
                .Top = 14,
                .Width = 520,
                .Text = promptText
            }

            m_Combo = New ComboBox() With {
                .Left = 14,
                .Top = 42,
                .Width = 520,
                .DropDownStyle = ComboBoxStyle.DropDownList
            }

            For Each item As String In items
                m_Combo.Items.Add(item)
            Next

            If m_Combo.Items.Count > 0 Then
                Dim index As Integer = Math.Max(0, m_Combo.FindStringExact(defaultValue))
                m_Combo.SelectedIndex = index
            End If

            m_OkButton = New Button() With {
                .Text = "OK",
                .Left = 368,
                .Top = 80,
                .Width = 80
            }
            AddHandler m_OkButton.Click, AddressOf OnOkClick

            m_CancelButton = New Button() With {
                .Text = "Cancel",
                .Left = 454,
                .Top = 80,
                .Width = 80
            }
            AddHandler m_CancelButton.Click, AddressOf OnCancelClick

            Me.Controls.Add(prompt)
            Me.Controls.Add(m_Combo)
            Me.Controls.Add(m_OkButton)
            Me.Controls.Add(m_CancelButton)
        End Sub

        Private Sub OnOkClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Private Sub OnCancelClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

        Public Function SelectedValue() As String
            If m_Combo.SelectedItem Is Nothing Then
                Return String.Empty
            End If

            Return m_Combo.SelectedItem.ToString()
        End Function
    End Class

    Public Class ToolProgressForm
        Inherits Form

        Private ReadOnly m_StatusLabel As System.Windows.Forms.Label
        Private ReadOnly m_ProgressBar As System.Windows.Forms.ProgressBar

        Public Sub New(ByVal titleText As String)
            Me.Text = titleText
            Me.Width = 520
            Me.Height = 140
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ControlBox = False
            Me.TopMost = True

            m_StatusLabel = New System.Windows.Forms.Label() With {
                .Left = 14,
                .Top = 14,
                .Width = 480,
                .Text = "Starting..."
            }

            m_ProgressBar = New System.Windows.Forms.ProgressBar() With {
                .Left = 14,
                .Top = 44,
                .Width = 480,
                .Height = 24,
                .Style = System.Windows.Forms.ProgressBarStyle.Continuous,
                .Minimum = 0,
                .Maximum = 100,
                .Value = 0
            }

            Me.Controls.Add(m_StatusLabel)
            Me.Controls.Add(m_ProgressBar)
        End Sub

        Public Sub UpdateProgress(ByVal percent As Integer, ByVal statusText As String)
            Dim safePercent As Integer = Math.Max(0, Math.Min(100, percent))
            m_StatusLabel.Text = statusText
            m_ProgressBar.Value = safePercent
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()
        End Sub

        Public Sub CompleteSuccess(ByVal statusText As String)
            UpdateProgress(100, statusText & " ✅")
        End Sub
    End Class

End Namespace
