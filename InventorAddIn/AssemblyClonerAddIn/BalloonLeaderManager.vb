Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic

Namespace AssemblyClonerAddIn

    Public Class BalloonLeaderManager
        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Try
                If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Open an IDW drawing, select one template balloon, then run this command.",
                                    "Auto Balloon Leaders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
                Dim sheet As Sheet = drawingDoc.ActiveSheet
                Dim templateBalloon As Balloon = GetSelectedTemplateBalloon(drawingDoc)

                If templateBalloon Is Nothing Then
                    MessageBox.Show("Select exactly one balloon on the active sheet, then run the command again.",
                                    "Auto Balloon Leaders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim targetView As DrawingView = ResolveViewForBalloon(sheet, templateBalloon)
                If targetView Is Nothing Then
                    MessageBox.Show("Could not resolve the target drawing view from the selected balloon.",
                                    "Auto Balloon Leaders", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim modeResult As DialogResult = MessageBox.Show(
                    "YES: Try to add leaders to selected balloon first (vertex-leader style), then fallback to new balloons if needed." & vbCrLf &
                    "NO: Create separate balloons only (recommended for reliability).",
                    "Leader Mode",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)

                If modeResult = DialogResult.Cancel Then
                    Return
                End If

                Dim tryVertexMode As Boolean = (modeResult = DialogResult.Yes)

                Dim candidates As List(Of BalloonCandidate) = BuildCandidates(sheet, targetView)
                If candidates.Count = 0 Then
                    MessageBox.Show("No additional visible occurrences found to balloon in the selected view.",
                                    "Auto Balloon Leaders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                PlaceCandidatesWithRouting(sheet, targetView, templateBalloon, candidates, tryVertexMode)

            Catch ex As Exception
                MessageBox.Show("Auto Balloon Leaders failed: " & ex.Message,
                                "Auto Balloon Leaders", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function GetSelectedTemplateBalloon(ByVal drawingDoc As DrawingDocument) As Balloon
            Dim selectedBalloon As Balloon = Nothing

            For Each selectedObj As Object In drawingDoc.SelectSet
                If TypeOf selectedObj Is Balloon Then
                    If selectedBalloon IsNot Nothing Then
                        Return Nothing
                    End If
                    selectedBalloon = CType(selectedObj, Balloon)
                End If
            Next

            Return selectedBalloon
        End Function

        Private Function ResolveViewForBalloon(ByVal sheet As Sheet, ByVal balloon As Balloon) As DrawingView
            Dim intentView As DrawingView = TryGetBalloonAttachedView(balloon)
            If intentView IsNot Nothing Then
                Return intentView
            End If

            Dim bestView As DrawingView = Nothing
            Dim bestDistSq As Double = Double.MaxValue

            Dim bPos As Point2d = Nothing
            Try
                bPos = balloon.Position
            Catch
                bPos = Nothing
            End Try

            If bPos Is Nothing Then
                Return Nothing
            End If

            For Each view As DrawingView In sheet.DrawingViews
                Try
                    Dim rb As Box2d = view.RangeBox
                    If rb Is Nothing Then Continue For

                    If bPos.X >= rb.MinPoint.X AndAlso bPos.X <= rb.MaxPoint.X AndAlso
                       bPos.Y >= rb.MinPoint.Y AndAlso bPos.Y <= rb.MaxPoint.Y Then
                        Return view
                    End If

                    Dim cx As Double = (rb.MinPoint.X + rb.MaxPoint.X) / 2.0
                    Dim cy As Double = (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0
                    Dim dx As Double = bPos.X - cx
                    Dim dy As Double = bPos.Y - cy
                    Dim d2 As Double = (dx * dx) + (dy * dy)
                    If d2 < bestDistSq Then
                        bestDistSq = d2
                        bestView = view
                    End If
                Catch
                End Try
            Next

            Return bestView
        End Function

        Private Function TryGetBalloonAttachedView(ByVal balloon As Balloon) As DrawingView
            Try
                Dim leaderObj As Object = balloon.Leader
                If leaderObj Is Nothing Then Return Nothing

                Dim rootNode As Object = leaderObj.RootNode
                If rootNode Is Nothing Then Return Nothing

                Dim childNodes As Object = rootNode.ChildNodes
                If childNodes Is Nothing OrElse childNodes.Count = 0 Then Return Nothing

                For i As Integer = 1 To childNodes.Count
                    Dim node As Object = childNodes.Item(i)
                    If node Is Nothing Then Continue For

                    Dim attached As Object = Nothing
                    Try
                        attached = node.AttachedEntity
                    Catch
                        attached = Nothing
                    End Try

                    Dim resolved As DrawingView = ResolveViewFromAttachedEntity(attached)
                    If resolved IsNot Nothing Then Return resolved
                Next
            Catch
            End Try

            Return Nothing
        End Function

        Private Function ResolveViewFromAttachedEntity(ByVal attached As Object) As DrawingView
            If attached Is Nothing Then Return Nothing

            Try
                Dim t As String = TypeName(attached)

                If t = "GeometryIntent" Then
                    Dim geomObj As Object = Nothing
                    Try
                        geomObj = attached.Geometry
                    Catch
                        geomObj = Nothing
                    End Try
                    Return ResolveViewFromAttachedEntity(geomObj)
                End If

                If t = "DrawingCurve" Then
                    Return CType(attached.Parent, DrawingView)
                End If

                If t = "DrawingCurveSegment" Then
                    Return CType(attached.Parent.Parent, DrawingView)
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function BuildCandidates(ByVal sheet As Sheet, ByVal targetView As DrawingView) As List(Of BalloonCandidate)
            Dim candidates As New List(Of BalloonCandidate)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim alreadyBallooned As HashSet(Of String) = CollectAlreadyBalloonedOccurrenceKeys(sheet, targetView)

            Dim occs As Object = GetViewOccurrences(targetView)
            If occs Is Nothing Then
                Return candidates
            End If

            For i As Integer = 1 To occs.Count
                Dim occ As Object = occs.Item(i)
                If occ Is Nothing Then Continue For

                Dim occKey As String = GetOccurrenceKey(occ)
                If String.IsNullOrWhiteSpace(occKey) Then Continue For
                If seen.Contains(occKey) Then Continue For
                seen.Add(occKey)

                If alreadyBallooned.Contains(occKey) Then Continue For
                If Not IsOccurrenceVisibleInView(occ, targetView) Then Continue For

                Dim curve As DrawingCurve = GetFirstDrawingCurveForOccurrence(targetView, occ)
                If curve Is Nothing Then Continue For

                Dim intent As GeometryIntent = Nothing
                Try
                    intent = sheet.CreateGeometryIntent(curve)
                Catch
                    intent = Nothing
                End Try
                If intent Is Nothing Then Continue For

                Dim attachPoint As Point2d = GetCurveMidPoint2d(curve)
                If attachPoint Is Nothing Then Continue For

                Dim candidate As New BalloonCandidate() With {
                    .Occurrence = occ,
                    .OccurrenceKey = occKey,
                    .AttachIntent = intent,
                    .AttachPoint = attachPoint,
                    .LabelText = GetOccurrenceLabel(occ)
                }
                candidates.Add(candidate)
            Next

            Return candidates
        End Function

        Private Function CollectAlreadyBalloonedOccurrenceKeys(ByVal sheet As Sheet, ByVal targetView As DrawingView) As HashSet(Of String)
            Dim setKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each balloon As Balloon In sheet.Balloons
                Try
                    Dim leaderObj As Object = balloon.Leader
                    If leaderObj Is Nothing Then Continue For
                    Dim rootNode As Object = leaderObj.RootNode
                    If rootNode Is Nothing Then Continue For
                    Dim childNodes As Object = rootNode.ChildNodes
                    If childNodes Is Nothing Then Continue For

                    For i As Integer = 1 To childNodes.Count
                        Dim node As Object = childNodes.Item(i)
                        If node Is Nothing Then Continue For
                        Dim attached As Object = Nothing
                        Try
                            attached = node.AttachedEntity
                        Catch
                            attached = Nothing
                        End Try

                        If attached Is Nothing Then Continue For
                        Dim viewForThis As DrawingView = ResolveViewFromAttachedEntity(attached)
                        If viewForThis Is Nothing OrElse Not Object.ReferenceEquals(viewForThis, targetView) Then
                            Continue For
                        End If

                        Dim occKey As String = GetOccurrenceKeyFromAttached(attached)
                        If Not String.IsNullOrWhiteSpace(occKey) Then
                            setKeys.Add(occKey)
                        End If
                    Next
                Catch
                End Try
            Next

            Return setKeys
        End Function

        Private Function GetOccurrenceKeyFromAttached(ByVal attached As Object) As String
            If attached Is Nothing Then Return String.Empty

            Try
                If TypeName(attached) = "GeometryIntent" Then
                    Return GetOccurrenceKeyFromAttached(attached.Geometry)
                End If
            Catch
            End Try

            Try
                If TypeName(attached) = "DrawingCurve" Then
                    Return GetOccurrenceKey(attached.ModelGeometry.Parent)
                End If
            Catch
            End Try

            Try
                If TypeName(attached) = "DrawingCurveSegment" Then
                    Return GetOccurrenceKey(attached.Parent.ModelGeometry.Parent)
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function GetOccurrenceKey(ByVal occ As Object) As String
            If occ Is Nothing Then Return String.Empty

            Try
                Return CStr(occ.InternalName)
            Catch
            End Try

            Try
                Return CStr(occ.Name)
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function GetViewOccurrences(ByVal view As DrawingView) As Object
            Try
                If view.ReferencedDocumentDescriptor IsNot Nothing Then
                    Dim refDoc As Document = view.ReferencedDocumentDescriptor.ReferencedDocument
                    If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Return refDoc.ComponentDefinition.Occurrences
                    End If
                End If
            Catch
            End Try

            Try
                Dim refDoc As Document = view.ReferencedDocument
                If refDoc IsNot Nothing AndAlso refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return refDoc.ComponentDefinition.Occurrences
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function IsOccurrenceVisibleInView(ByVal occ As Object, ByVal view As DrawingView) As Boolean
            Try
                Return CBool(occ.VisibleInDrawingView(view))
            Catch
                Return True
            End Try
        End Function

        Private Function GetFirstDrawingCurveForOccurrence(ByVal view As DrawingView, ByVal occ As Object) As DrawingCurve
            Try
                Dim curves As DrawingCurvesEnumerator = view.DrawingCurves(occ)
                If curves IsNot Nothing AndAlso curves.Count > 0 Then
                    Return curves.Item(1)
                End If
            Catch
            End Try

            Return Nothing
        End Function

        Private Function GetCurveMidPoint2d(ByVal curve As DrawingCurve) As Point2d
            Try
                Dim rb As Box2d = curve.RangeBox
                If rb Is Nothing Then Return Nothing

                Dim x As Double = (rb.MinPoint.X + rb.MaxPoint.X) / 2.0
                Dim y As Double = (rb.MinPoint.Y + rb.MaxPoint.Y) / 2.0
                Return m_InventorApp.TransientGeometry.CreatePoint2d(x, y)
            Catch
                Return Nothing
            End Try
        End Function

        Private Function GetOccurrenceLabel(ByVal occ As Object) As String
            Try
                Dim itemNumber As String = ""
                Try
                    itemNumber = CStr(occ.ItemNumber)
                Catch
                    itemNumber = ""
                End Try
                If Not String.IsNullOrWhiteSpace(itemNumber) Then
                    Return itemNumber
                End If
            Catch
            End Try

            Try
                Dim refDoc As Document = occ.ReferencedDocument
                If refDoc IsNot Nothing Then
                    Dim pn As String = CStr(refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value)
                    If Not String.IsNullOrWhiteSpace(pn) Then
                        Return pn
                    End If
                End If
            Catch
            End Try

            Try
                Return CStr(occ.Name)
            Catch
                Return ""
            End Try
        End Function

        Private Sub PlaceCandidatesWithRouting(
            ByVal sheet As Sheet,
            ByVal view As DrawingView,
            ByVal templateBalloon As Balloon,
            ByVal candidates As List(Of BalloonCandidate),
            ByVal tryVertexMode As Boolean)

            Dim right As New List(Of BalloonCandidate)()
            Dim left As New List(Of BalloonCandidate)()

            Dim viewCenterX As Double = (view.RangeBox.MinPoint.X + view.RangeBox.MaxPoint.X) / 2.0

            For Each candidate As BalloonCandidate In candidates
                If candidate.AttachPoint.X >= viewCenterX Then
                    right.Add(candidate)
                Else
                    left.Add(candidate)
                End If
            Next

            right.Sort(Function(a, b) b.AttachPoint.Y.CompareTo(a.AttachPoint.Y))
            left.Sort(Function(a, b) b.AttachPoint.Y.CompareTo(a.AttachPoint.Y))

            Dim createdCount As Integer = 0
            Dim vertexCount As Integer = 0

            createdCount += PlaceSide(sheet, view, templateBalloon, right, True, tryVertexMode, vertexCount)
            createdCount += PlaceSide(sheet, view, templateBalloon, left, False, tryVertexMode, vertexCount)

            MessageBox.Show(
                "Processed: " & candidates.Count.ToString() & vbCrLf &
                "Vertex leaders added: " & vertexCount.ToString() & vbCrLf &
                "New balloons added: " & createdCount.ToString(),
                "Auto Balloon Leaders",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        End Sub

        Private Function PlaceSide(
            ByVal sheet As Sheet,
            ByVal view As DrawingView,
            ByVal templateBalloon As Balloon,
            ByVal sideCandidates As List(Of BalloonCandidate),
            ByVal isRight As Boolean,
            ByVal tryVertexMode As Boolean,
            ByRef vertexCount As Integer) As Integer

            If sideCandidates.Count = 0 Then Return 0

            Dim added As Integer = 0

            Dim rb As Box2d = view.RangeBox
            Dim topY As Double = rb.MaxPoint.Y
            Dim bottomY As Double = rb.MinPoint.Y
            Dim span As Double = Math.Max(0.001, topY - bottomY)

            Dim gap As Double = Math.Max(0.6, span / Math.Max(8.0, CDbl(sideCandidates.Count + 2)))
            Dim outerOffset As Double = Math.Max(1.2, span * 0.03)
            Dim innerOffset As Double = Math.Max(0.6, outerOffset * 0.55)

            Dim endpointX As Double = If(isRight, rb.MaxPoint.X + outerOffset, rb.MinPoint.X - outerOffset)
            Dim elbowX As Double = If(isRight, rb.MaxPoint.X + innerOffset, rb.MinPoint.X - innerOffset)

            Dim nextY As Double = topY - gap
            For Each candidate As BalloonCandidate In sideCandidates
                Dim slotY As Double = Math.Min(nextY, candidate.AttachPoint.Y)
                slotY = Math.Max(slotY, bottomY + gap)

                Dim endPt As Point2d = m_InventorApp.TransientGeometry.CreatePoint2d(endpointX, slotY)
                Dim elbowPt As Point2d = m_InventorApp.TransientGeometry.CreatePoint2d(elbowX, candidate.AttachPoint.Y)

                Dim handledByVertex As Boolean = False
                If tryVertexMode Then
                    handledByVertex = TryAddVertexLeader(templateBalloon, candidate.AttachIntent, endPt, elbowPt)
                    If handledByVertex Then
                        vertexCount += 1
                    End If
                End If

                If Not handledByVertex Then
                    If AddBalloonUsingTemplateStyle(sheet, templateBalloon, candidate, endPt, elbowPt) Then
                        added += 1
                    End If
                End If

                nextY = slotY - gap
            Next

            Return added
        End Function

        Private Function TryAddVertexLeader(
            ByVal templateBalloon As Balloon,
            ByVal attachIntent As GeometryIntent,
            ByVal endPt As Point2d,
            ByVal elbowPt As Point2d) As Boolean

            TryAddVertexLeader = False

            Try
                Dim leaderObj As Object = templateBalloon.Leader
                If leaderObj Is Nothing Then Return False

                Dim rootNode As Object = leaderObj.RootNode
                If rootNode Is Nothing Then Return False

                Dim childNodes As Object = rootNode.ChildNodes
                If childNodes Is Nothing Then Return False

                Try
                    CallByName(childNodes, "Add", CallType.Method, endPt, elbowPt, attachIntent)
                    Return True
                Catch
                End Try

                Try
                    CallByName(childNodes, "Add", CallType.Method, endPt, attachIntent)
                    Return True
                Catch
                End Try
            Catch
            End Try

            Return False
        End Function

        Private Function AddBalloonUsingTemplateStyle(
            ByVal sheet As Sheet,
            ByVal templateBalloon As Balloon,
            ByVal candidate As BalloonCandidate,
            ByVal endPt As Point2d,
            ByVal elbowPt As Point2d) As Boolean

            Try
                Dim leaderPoints As ObjectCollection = m_InventorApp.TransientObjects.CreateObjectCollection()
                leaderPoints.Add(endPt)
                leaderPoints.Add(elbowPt)
                leaderPoints.Add(candidate.AttachIntent)

                Dim newBalloon As Balloon = Nothing
                Try
                    newBalloon = sheet.Balloons.Add(leaderPoints, , , , templateBalloon.Style)
                Catch
                    newBalloon = Nothing
                End Try

                If newBalloon Is Nothing Then
                    newBalloon = sheet.Balloons.Add(leaderPoints)
                    If newBalloon IsNot Nothing Then
                        Try
                            newBalloon.Style = templateBalloon.Style
                        Catch
                        End Try
                    End If
                End If

                If newBalloon Is Nothing Then
                    Return False
                End If

                If Not String.IsNullOrWhiteSpace(candidate.LabelText) Then
                    Try
                        newBalloon.Text = candidate.LabelText
                    Catch
                    End Try
                End If

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Class BalloonCandidate
            Public Property Occurrence As Object
            Public Property OccurrenceKey As String
            Public Property AttachIntent As GeometryIntent
            Public Property AttachPoint As Point2d
            Public Property LabelText As String
        End Class

    End Class

End Namespace
