Imports Inventor
Imports System.Collections.Generic
Imports System.Globalization
Imports Microsoft.Win32
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class AutoBalloonerTool
        Private ReadOnly m_InventorApp As Inventor.Application
        Private Const DefaultLeaderDistanceMultiplier As Double = 10.0
        Private Const DefaultElbowDistanceRatio As Double = 0.45
        Private Const DefaultMinSpacingRatio As Double = 0.85

        Private Shared s_LastLeaderDistanceMultiplier As Decimal = 10D
        Private Shared s_LastElbowDistanceRatio As Decimal = 0.45D
        Private Shared s_LastMinSpacingRatio As Decimal = 0.85D
        Private Shared s_LastPlacementMode As BalloonPlacementMode = BalloonPlacementMode.AutoQuadrant
        Private Shared s_LastProcessAllSheets As Boolean = False
        Private Shared s_LastReplaceExistingBalloons As Boolean = False
        Private Shared s_LastMinimizeLeaderCrossing As Boolean = True
        Private Shared s_LastCombineSamePartMultiLeader As Boolean = False
        Private Shared s_HasLoadedPersistedSettings As Boolean = False

        Private Const AutoBalloonerRegistryPath As String = "Software\Spectiv\InventorAutomationSuite\AutoBallooner"

        Private Enum BalloonSide
            LeftSide
            RightSide
            TopSide
            BottomSide
        End Enum

        Private Enum BalloonPlacementMode
            AutoQuadrant = 0
            HorizontalOnly = 1
            VerticalOnly = 2
            ForceRight = 3
            ForceLeft = 4
            ForceTop = 5
            ForceBottom = 6
        End Enum

        Private Class AutoBalloonerConfig
            Public Property LeaderDistanceMultiplier As Double
            Public Property ElbowDistanceRatio As Double
            Public Property MinSpacingRatio As Double
            Public Property PlacementMode As BalloonPlacementMode
            Public Property ProcessAllSheets As Boolean
            Public Property ReplaceExistingBalloons As Boolean
            Public Property MinimizeLeaderCrossing As Boolean
            Public Property CombineSamePartMultiLeader As Boolean
        End Class

        Private Class BalloonPlacementCandidate
            Public Property ViewOccurrenceKey As String
            Public Property GroupingKey As String
            Public Property AttachIntent As GeometryIntent
            Public Property AttachPoint As Point2d
            Public Property EndPoint As Point2d
            Public Property Side As BalloonSide
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Try
                If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Open an IDW drawing before running Auto Ballooner.",
                                    "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
                Dim sheet As Sheet = drawingDoc.ActiveSheet
                If sheet Is Nothing Then
                    MessageBox.Show("No active sheet is selected.",
                                    "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim config As AutoBalloonerConfig = ShowSettingsDialog()
                If config Is Nothing Then
                    Return
                End If

                Dim sheetsToProcess As New List(Of Sheet)()
                If config.ProcessAllSheets Then
                    For Each candidateSheet As Sheet In drawingDoc.Sheets
                        If candidateSheet IsNot Nothing AndAlso candidateSheet.DrawingViews.Count > 0 Then
                            sheetsToProcess.Add(candidateSheet)
                        End If
                    Next
                Else
                    If sheet.DrawingViews.Count = 0 Then
                        MessageBox.Show("The active sheet has no drawing views to balloon.",
                                        "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If
                    sheetsToProcess.Add(sheet)
                End If

                If sheetsToProcess.Count = 0 Then
                    MessageBox.Show("No drawing views found in the selected scope.",
                                    "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim preferredStyle As BalloonStyle = ResolvePreferredBalloonStyle(drawingDoc)

                Dim addedCount As Integer = 0
                Dim leaderAddedCount As Integer = 0
                Dim skippedCount As Integer = 0
                Dim candidateCount As Integer = 0
                Dim viewCount As Integer = 0
                Dim removedBalloonCount As Integer = 0

                For Each targetSheet As Sheet In sheetsToProcess
                    If config.ReplaceExistingBalloons Then
                        removedBalloonCount += RemoveAllBalloonsFromSheet(targetSheet)
                    End If

                    Dim alreadyBallooned As HashSet(Of String)
                    If config.ReplaceExistingBalloons Then
                        alreadyBallooned = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    Else
                        alreadyBallooned = CollectAlreadyBalloonedViewOccurrenceKeys(targetSheet)
                    End If

                    For Each view As DrawingView In targetSheet.DrawingViews
                        viewCount += 1
                        Dim occupiedEndPoints As New List(Of Point2d)()
                        ProcessView(targetSheet, view, preferredStyle, alreadyBallooned, occupiedEndPoints, config, addedCount, leaderAddedCount, skippedCount, candidateCount)
                    Next
                Next

                Dim scopeText As String = If(config.ProcessAllSheets,
                                             "all sheets",
                                             "active sheet: " & sheet.Name)

                MessageBox.Show("Auto Ballooner complete for " & scopeText & vbCrLf & vbCrLf &
                                "Sheets processed: " & sheetsToProcess.Count.ToString() & vbCrLf &
                                "Views scanned: " & viewCount.ToString() & vbCrLf &
                                "Candidates found: " & candidateCount.ToString() & vbCrLf &
                                "Balloons added: " & addedCount.ToString() & vbCrLf &
                                "Extra leaders added: " & leaderAddedCount.ToString() & vbCrLf &
                                "Balloons replaced: " & removedBalloonCount.ToString() & vbCrLf &
                                "Skipped: " & skippedCount.ToString(),
                                "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Auto Ballooner failed: " & ex.Message,
                                "Auto Ballooner", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub ProcessView(
            ByVal sheet As Sheet,
            ByVal view As DrawingView,
            ByVal preferredStyle As BalloonStyle,
            ByVal alreadyBallooned As HashSet(Of String),
            ByVal occupiedEndPoints As List(Of Point2d),
            ByVal config As AutoBalloonerConfig,
            ByRef addedCount As Integer,
            ByRef leaderAddedCount As Integer,
            ByRef skippedCount As Integer,
            ByRef candidateCount As Integer)

            If view Is Nothing OrElse view.Suppressed Then
                Return
            End If

            Dim occurrences As Object = GetViewOccurrences(view)
            If occurrences Is Nothing Then
                Return
            End If

            Dim viewLeft As Double = view.Left
            Dim viewTop As Double = view.Top
            Dim viewRight As Double = viewLeft + view.Width
            Dim viewBottom As Double = viewTop - view.Height

            If viewBottom > viewTop Then
                Dim tmp As Double = viewTop
                viewTop = viewBottom
                viewBottom = tmp
            End If

            Dim viewCenterX As Double = (viewLeft + viewRight) / 2.0
            Dim viewCenterY As Double = (viewBottom + viewTop) / 2.0
            Dim viewSpan As Double = Math.Max(Math.Abs(view.Width), Math.Abs(view.Height))
            Dim leaderOffset As Double = Math.Max(0.5, viewSpan * 0.04) * config.LeaderDistanceMultiplier
            Dim elbowOffset As Double = Math.Max(0.25, leaderOffset * config.ElbowDistanceRatio)
            Dim minSpacing As Double = Math.Max(0.45, leaderOffset * config.MinSpacingRatio)

            Dim candidatesBySide As New Dictionary(Of BalloonSide, List(Of BalloonPlacementCandidate)) From {
                {BalloonSide.LeftSide, New List(Of BalloonPlacementCandidate)()},
                {BalloonSide.RightSide, New List(Of BalloonPlacementCandidate)()},
                {BalloonSide.TopSide, New List(Of BalloonPlacementCandidate)()},
                {BalloonSide.BottomSide, New List(Of BalloonPlacementCandidate)()}
            }

            For i As Integer = 1 To occurrences.Count
                Dim occ As Object = Nothing
                Try
                    occ = occurrences.Item(i)
                Catch
                    occ = Nothing
                End Try

                If occ Is Nothing Then
                    Continue For
                End If

                Dim occurrenceKey As String = GetOccurrenceKey(occ)
                If String.IsNullOrWhiteSpace(occurrenceKey) Then
                    skippedCount += 1
                    Continue For
                End If

                Dim viewOccurrenceKey As String = BuildViewOccurrenceKey(view, occurrenceKey)
                If alreadyBallooned.Contains(viewOccurrenceKey) Then
                    Continue For
                End If

                If Not IsOccurrenceVisibleInView(occ, view) Then
                    Continue For
                End If

                Dim curves As DrawingCurvesEnumerator = Nothing
                Try
                    curves = view.DrawingCurves(occ)
                Catch
                    curves = Nothing
                End Try

                If curves Is Nothing OrElse curves.Count = 0 Then
                    Continue For
                End If

                Dim pointSamples As List(Of Point2d) = CollectCurveSamplePoints(curves)
                If pointSamples.Count = 0 Then
                    skippedCount += 1
                    Continue For
                End If

                candidateCount += 1

                Dim minX As Double = Double.MaxValue
                Dim minY As Double = Double.MaxValue
                Dim maxX As Double = Double.MinValue
                Dim maxY As Double = Double.MinValue

                For Each pt As Point2d In pointSamples
                    If pt Is Nothing Then Continue For
                    If pt.X < minX Then minX = pt.X
                    If pt.X > maxX Then maxX = pt.X
                    If pt.Y < minY Then minY = pt.Y
                    If pt.Y > maxY Then maxY = pt.Y
                Next

                If minX = Double.MaxValue OrElse minY = Double.MaxValue Then
                    skippedCount += 1
                    Continue For
                End If

                Dim partCenterX As Double = (minX + maxX) / 2.0
                Dim partCenterY As Double = (minY + maxY) / 2.0

                Dim side As BalloonSide = ResolveSideFromPlacementMode(partCenterX, partCenterY, viewCenterX, viewCenterY, config.PlacementMode)

                Dim desiredAttachPoint As Point2d = CreateAttachPointForSide(minX, minY, maxX, maxY, side)
                Dim bestCurve As DrawingCurve = Nothing
                Dim bestCurvePoint As Point2d = Nothing
                FindBestCurveForPoint(curves, desiredAttachPoint, bestCurve, bestCurvePoint)

                If bestCurve Is Nothing Then
                    skippedCount += 1
                    Continue For
                End If

                Dim attachIntent As GeometryIntent = CreateIntent(sheet, bestCurve, bestCurvePoint)
                If attachIntent Is Nothing Then
                    skippedCount += 1
                    Continue For
                End If

                Dim endPoint As Point2d = OffsetPoint(bestCurvePoint, side, leaderOffset)
                Dim groupingKey As String = ResolveOccurrenceGroupingKey(occ)
                If String.IsNullOrWhiteSpace(groupingKey) Then
                    groupingKey = viewOccurrenceKey
                End If

                candidatesBySide(side).Add(New BalloonPlacementCandidate() With {
                    .ViewOccurrenceKey = viewOccurrenceKey,
                    .GroupingKey = groupingKey,
                    .AttachIntent = attachIntent,
                    .AttachPoint = bestCurvePoint,
                    .EndPoint = endPoint,
                    .Side = side
                })
            Next

            Dim groupBalloonMap As Dictionary(Of String, Balloon) = Nothing
            If config.CombineSamePartMultiLeader Then
                groupBalloonMap = New Dictionary(Of String, Balloon)(StringComparer.OrdinalIgnoreCase)
            End If

            For Each sideKey As BalloonSide In candidatesBySide.Keys
                Dim sideCandidates As List(Of BalloonPlacementCandidate) = candidatesBySide(sideKey)
                If sideCandidates.Count = 0 Then
                    Continue For
                End If

                If config.MinimizeLeaderCrossing Then
                    ApplyNonCrossingLayout(sideCandidates, sideKey, minSpacing)
                End If

                For Each candidate As BalloonPlacementCandidate In sideCandidates
                    Dim finalEndPoint As Point2d = candidate.EndPoint
                    If Not config.MinimizeLeaderCrossing Then
                        finalEndPoint = ApplyEndpointSpacing(finalEndPoint, sideKey, occupiedEndPoints, minSpacing)
                    End If

                    Dim clearance As Double = Math.Max(0.35, Math.Min(leaderOffset, minSpacing))
                    finalEndPoint = EnsurePointOutsideView(finalEndPoint, sideKey, viewLeft, viewRight, viewTop, viewBottom, clearance)
                    finalEndPoint = PushPointOutsideAllViews(sheet, finalEndPoint, sideKey, Math.Max(0.25, minSpacing * 0.4), 120)

                    Dim elbowPoint As Point2d = OffsetPoint(candidate.AttachPoint, sideKey, elbowOffset)
                    elbowPoint = AlignElbowToEndpoint(elbowPoint, finalEndPoint, sideKey)

                    If config.CombineSamePartMultiLeader AndAlso groupBalloonMap IsNot Nothing Then
                        Dim existingBalloon As Balloon = Nothing
                        If groupBalloonMap.TryGetValue(candidate.GroupingKey, existingBalloon) AndAlso existingBalloon IsNot Nothing Then
                            If TryAddVertexLeader(existingBalloon, candidate.AttachIntent, finalEndPoint, elbowPoint) Then
                                occupiedEndPoints.Add(finalEndPoint)
                                alreadyBallooned.Add(candidate.ViewOccurrenceKey)
                                leaderAddedCount += 1
                            Else
                                skippedCount += 1
                            End If

                            Continue For
                        End If
                    End If

                    Dim createdBalloon As Balloon = Nothing
                    If AddBalloon(sheet, preferredStyle, candidate.AttachIntent, finalEndPoint, elbowPoint, createdBalloon) Then
                        occupiedEndPoints.Add(finalEndPoint)
                        alreadyBallooned.Add(candidate.ViewOccurrenceKey)
                        addedCount += 1

                        If config.CombineSamePartMultiLeader AndAlso groupBalloonMap IsNot Nothing AndAlso
                           Not String.IsNullOrWhiteSpace(candidate.GroupingKey) AndAlso createdBalloon IsNot Nothing Then
                            If Not groupBalloonMap.ContainsKey(candidate.GroupingKey) Then
                                groupBalloonMap.Add(candidate.GroupingKey, createdBalloon)
                            End If
                        End If
                    Else
                        skippedCount += 1
                    End If
                Next
            Next
        End Sub

        Private Function ResolveOccurrenceGroupingKey(ByVal occ As Object) As String
            If occ Is Nothing Then
                Return String.Empty
            End If

            Dim refDoc As Document = Nothing

            Try
                refDoc = TryCast(occ.Definition.Document, Document)
            Catch
                refDoc = Nothing
            End Try

            If refDoc Is Nothing Then
                Try
                    refDoc = TryCast(occ.ReferencedDocument, Document)
                Catch
                    refDoc = Nothing
                End Try
            End If

            If refDoc IsNot Nothing Then
                Try
                    Dim designSet As PropertySet = refDoc.PropertySets.Item("Design Tracking Properties")
                    Dim partNumber As String = Convert.ToString(designSet.Item("Part Number").Value)
                    If Not String.IsNullOrWhiteSpace(partNumber) Then
                        Return "PN|" & partNumber.Trim().ToLowerInvariant()
                    End If
                Catch
                End Try

                Dim fullPath As String = String.Empty
                Try
                    fullPath = refDoc.FullFileName
                Catch
                    fullPath = String.Empty
                End Try

                If Not String.IsNullOrWhiteSpace(fullPath) Then
                    Return "PATH|" & fullPath.Trim().ToLowerInvariant()
                End If
            End If

            Dim occKey As String = GetOccurrenceKey(occ)
            If Not String.IsNullOrWhiteSpace(occKey) Then
                Return "OCC|" & occKey.Trim().ToLowerInvariant()
            End If

            Return String.Empty
        End Function

        Private Sub ApplyNonCrossingLayout(ByVal candidates As List(Of BalloonPlacementCandidate),
                                           ByVal side As BalloonSide,
                                           ByVal minSpacing As Double)
            If candidates Is Nothing OrElse candidates.Count <= 1 Then
                Return
            End If

            If side = BalloonSide.LeftSide OrElse side = BalloonSide.RightSide Then
                candidates.Sort(Function(a, b) b.AttachPoint.Y.CompareTo(a.AttachPoint.Y))

                Dim alignedX As Double = If(side = BalloonSide.LeftSide, Double.MaxValue, Double.MinValue)
                For Each candidate As BalloonPlacementCandidate In candidates
                    If side = BalloonSide.LeftSide Then
                        If candidate.EndPoint.X < alignedX Then
                            alignedX = candidate.EndPoint.X
                        End If
                    Else
                        If candidate.EndPoint.X > alignedX Then
                            alignedX = candidate.EndPoint.X
                        End If
                    End If
                Next

                Dim previousY As Double = Double.MaxValue
                For index As Integer = 0 To candidates.Count - 1
                    Dim desiredY As Double = candidates(index).EndPoint.Y
                    If index > 0 Then
                        Dim maxAllowedY As Double = previousY - minSpacing
                        If desiredY > maxAllowedY Then
                            desiredY = maxAllowedY
                        End If
                    End If

                    candidates(index).EndPoint = m_InventorApp.TransientGeometry.CreatePoint2d(alignedX, desiredY)
                    previousY = desiredY
                Next
            Else
                candidates.Sort(Function(a, b) a.AttachPoint.X.CompareTo(b.AttachPoint.X))

                Dim alignedY As Double = If(side = BalloonSide.TopSide, Double.MinValue, Double.MaxValue)
                For Each candidate As BalloonPlacementCandidate In candidates
                    If side = BalloonSide.TopSide Then
                        If candidate.EndPoint.Y > alignedY Then
                            alignedY = candidate.EndPoint.Y
                        End If
                    Else
                        If candidate.EndPoint.Y < alignedY Then
                            alignedY = candidate.EndPoint.Y
                        End If
                    End If
                Next

                Dim previousX As Double = Double.MinValue
                For index As Integer = 0 To candidates.Count - 1
                    Dim desiredX As Double = candidates(index).EndPoint.X
                    If index > 0 Then
                        Dim minAllowedX As Double = previousX + minSpacing
                        If desiredX < minAllowedX Then
                            desiredX = minAllowedX
                        End If
                    End If

                    candidates(index).EndPoint = m_InventorApp.TransientGeometry.CreatePoint2d(desiredX, alignedY)
                    previousX = desiredX
                Next
            End If
        End Sub

        Private Function EnsurePointOutsideView(ByVal pointValue As Point2d,
                                                ByVal side As BalloonSide,
                                                ByVal viewLeft As Double,
                                                ByVal viewRight As Double,
                                                ByVal viewTop As Double,
                                                ByVal viewBottom As Double,
                                                ByVal clearance As Double) As Point2d
            Dim x As Double = pointValue.X
            Dim y As Double = pointValue.Y

            Select Case side
                Case BalloonSide.LeftSide
                    If x > (viewLeft - clearance) Then
                        x = viewLeft - clearance
                    End If
                Case BalloonSide.RightSide
                    If x < (viewRight + clearance) Then
                        x = viewRight + clearance
                    End If
                Case BalloonSide.TopSide
                    If y < (viewTop + clearance) Then
                        y = viewTop + clearance
                    End If
                Case BalloonSide.BottomSide
                    If y > (viewBottom - clearance) Then
                        y = viewBottom - clearance
                    End If
            End Select

            Return m_InventorApp.TransientGeometry.CreatePoint2d(x, y)
        End Function

        Private Function IsPointInsideAnyView(ByVal sheet As Sheet,
                                              ByVal pointValue As Point2d,
                                              ByVal padding As Double) As Boolean
            If sheet Is Nothing OrElse pointValue Is Nothing Then
                Return False
            End If

            For Each drawingView As DrawingView In sheet.DrawingViews
                If drawingView Is Nothing Then
                    Continue For
                End If

                If drawingView.Suppressed Then
                    Continue For
                End If

                Dim left As Double = drawingView.Left
                Dim top As Double = drawingView.Top
                Dim right As Double = left + drawingView.Width
                Dim bottom As Double = top - drawingView.Height

                If bottom > top Then
                    Dim tmp As Double = top
                    top = bottom
                    bottom = tmp
                End If

                If pointValue.X >= (left - padding) AndAlso pointValue.X <= (right + padding) AndAlso
                   pointValue.Y >= (bottom - padding) AndAlso pointValue.Y <= (top + padding) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Function PushPointOutsideAllViews(ByVal sheet As Sheet,
                                                  ByVal startPoint As Point2d,
                                                  ByVal side As BalloonSide,
                                                  ByVal stepSize As Double,
                                                  ByVal maxSteps As Integer) As Point2d
            Dim adjusted As Point2d = startPoint

            For stepIndex As Integer = 0 To maxSteps - 1
                If Not IsPointInsideAnyView(sheet, adjusted, 0.01) Then
                    Return adjusted
                End If

                Select Case side
                    Case BalloonSide.LeftSide
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(adjusted.X - stepSize, adjusted.Y)
                    Case BalloonSide.RightSide
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(adjusted.X + stepSize, adjusted.Y)
                    Case BalloonSide.TopSide
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(adjusted.X, adjusted.Y + stepSize)
                    Case BalloonSide.BottomSide
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(adjusted.X, adjusted.Y - stepSize)
                End Select
            Next

            Return adjusted
        End Function

        Private Function ShowSettingsDialog() As AutoBalloonerConfig
            LoadPersistedSettingsIfNeeded()

            Using form As New AutoBalloonerSettingsForm()
                If form.ShowDialog() = DialogResult.OK Then
                    Return form.GetConfig()
                End If
            End Using

            Return Nothing
        End Function

        Private Function ResolveSideFromPlacementMode(ByVal partCenterX As Double,
                                                     ByVal partCenterY As Double,
                                                     ByVal viewCenterX As Double,
                                                     ByVal viewCenterY As Double,
                                                     ByVal placementMode As BalloonPlacementMode) As BalloonSide
            Select Case placementMode
                Case BalloonPlacementMode.ForceRight
                    Return BalloonSide.RightSide
                Case BalloonPlacementMode.ForceLeft
                    Return BalloonSide.LeftSide
                Case BalloonPlacementMode.ForceTop
                    Return BalloonSide.TopSide
                Case BalloonPlacementMode.ForceBottom
                    Return BalloonSide.BottomSide
                Case BalloonPlacementMode.HorizontalOnly
                    Return If(partCenterX >= viewCenterX, BalloonSide.RightSide, BalloonSide.LeftSide)
                Case BalloonPlacementMode.VerticalOnly
                    Return If(partCenterY >= viewCenterY, BalloonSide.TopSide, BalloonSide.BottomSide)
                Case Else
                    Return ResolveSideFromViewCenter(partCenterX, partCenterY, viewCenterX, viewCenterY)
            End Select
        End Function

        Private Function ResolvePreferredBalloonStyle(ByVal drawingDoc As DrawingDocument) As BalloonStyle
            For Each selectedObj As Object In drawingDoc.SelectSet
                If TypeOf selectedObj Is Balloon Then
                    Try
                        Return CType(selectedObj, Balloon).Style
                    Catch
                    End Try
                End If
            Next

            Return Nothing
        End Function

        Private Function CollectAlreadyBalloonedViewOccurrenceKeys(ByVal sheet As Sheet) As HashSet(Of String)
            Dim keys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

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

                        Dim occurrenceKey As String = GetOccurrenceKeyFromAttached(attached)
                        Dim attachedView As DrawingView = ResolveViewFromAttachedEntity(attached)
                        If attachedView IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(occurrenceKey) Then
                            keys.Add(BuildViewOccurrenceKey(attachedView, occurrenceKey))
                        End If
                    Next
                Catch
                End Try
            Next

            Return keys
        End Function

        Private Function RemoveAllBalloonsFromSheet(ByVal sheet As Sheet) As Integer
            If sheet Is Nothing Then
                Return 0
            End If

            Dim removed As Integer = 0
            For index As Integer = sheet.Balloons.Count To 1 Step -1
                Try
                    Dim balloon As Balloon = sheet.Balloons.Item(index)
                    If balloon Is Nothing Then
                        Continue For
                    End If

                    balloon.Delete()
                    removed += 1
                Catch
                End Try
            Next

            Return removed
        End Function

        Private Function BuildViewOccurrenceKey(ByVal view As DrawingView, ByVal occurrenceKey As String) As String
            Return GetViewKey(view) & "|" & occurrenceKey
        End Function

        Private Function GetViewKey(ByVal view As DrawingView) As String
            If view Is Nothing Then
                Return "VIEW_UNKNOWN"
            End If

            Try
                If Not String.IsNullOrWhiteSpace(view.Name) Then
                    Return view.Name.Trim()
                End If
            Catch
            End Try

            Return "VIEW_UNKNOWN"
        End Function

        Private Function ResolveViewFromAttachedEntity(ByVal attached As Object) As DrawingView
            If attached Is Nothing Then
                Return Nothing
            End If

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

        Private Function GetOccurrenceKeyFromAttached(ByVal attached As Object) As String
            If attached Is Nothing Then
                Return String.Empty
            End If

            Try
                If TypeName(attached) = "GeometryIntent" Then
                    Return GetOccurrenceKeyFromAttached(attached.Geometry)
                End If
            Catch
            End Try

            Try
                If TypeName(attached) = "DrawingCurve" Then
                    Return GetOccurrenceKeyFromModelGeometry(attached.ModelGeometry)
                End If
            Catch
            End Try

            Try
                If TypeName(attached) = "DrawingCurveSegment" Then
                    Return GetOccurrenceKeyFromModelGeometry(attached.Parent.ModelGeometry)
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function GetOccurrenceKeyFromModelGeometry(ByVal modelGeometry As Object) As String
            If modelGeometry Is Nothing Then
                Return String.Empty
            End If

            Try
                Dim containingOccurrence As Object = modelGeometry.ContainingOccurrence
                Dim key As String = GetOccurrenceKey(containingOccurrence)
                If Not String.IsNullOrWhiteSpace(key) Then
                    Return key
                End If
            Catch
            End Try

            Try
                Return GetOccurrenceKey(modelGeometry.Parent)
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function GetOccurrenceKey(ByVal occ As Object) As String
            If occ Is Nothing Then
                Return String.Empty
            End If

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
                Return CBool(view.GetVisibility(occ))
            Catch
            End Try

            Try
                Return CBool(occ.VisibleInDrawingView(view))
            Catch
            End Try

            Try
                Return CBool(occ.Visible)
            Catch
            End Try

            Return True
        End Function

        Private Function CollectCurveSamplePoints(ByVal curves As DrawingCurvesEnumerator) As List(Of Point2d)
            Dim points As New List(Of Point2d)()

            For i As Integer = 1 To curves.Count
                Dim curve As DrawingCurve = Nothing
                Try
                    curve = curves.Item(i)
                Catch
                    curve = Nothing
                End Try

                If curve Is Nothing Then
                    Continue For
                End If

                AddCurvePoints(curve, points)
            Next

            Return points
        End Function

        Private Sub AddCurvePoints(ByVal curve As DrawingCurve, ByVal points As List(Of Point2d))
            AddPoint(curve.StartPoint, points)
            AddPoint(curve.MidPoint, points)
            AddPoint(curve.EndPoint, points)

            Try
                For Each seg As DrawingCurveSegment In curve.Segments
                    AddPoint(seg.StartPoint, points)
                    AddPoint(seg.EndPoint, points)
                Next
            Catch
            End Try
        End Sub

        Private Sub AddPoint(ByVal pointValue As Point2d, ByVal points As List(Of Point2d))
            If pointValue Is Nothing Then
                Return
            End If

            points.Add(pointValue)
        End Sub

        Private Function ResolveSideFromViewCenter(ByVal partCenterX As Double,
                                                   ByVal partCenterY As Double,
                                                   ByVal viewCenterX As Double,
                                                   ByVal viewCenterY As Double) As BalloonSide
            Dim dx As Double = partCenterX - viewCenterX
            Dim dy As Double = partCenterY - viewCenterY

            If Math.Abs(dx) >= Math.Abs(dy) Then
                Return If(dx >= 0.0, BalloonSide.RightSide, BalloonSide.LeftSide)
            End If

            Return If(dy >= 0.0, BalloonSide.TopSide, BalloonSide.BottomSide)
        End Function

        Private Function CreateAttachPointForSide(ByVal minX As Double,
                                                  ByVal minY As Double,
                                                  ByVal maxX As Double,
                                                  ByVal maxY As Double,
                                                  ByVal side As BalloonSide) As Point2d
            Dim centerX As Double = (minX + maxX) / 2.0
            Dim centerY As Double = (minY + maxY) / 2.0

            Select Case side
                Case BalloonSide.LeftSide
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(minX, centerY)
                Case BalloonSide.RightSide
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(maxX, centerY)
                Case BalloonSide.TopSide
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(centerX, maxY)
                Case Else
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(centerX, minY)
            End Select
        End Function

        Private Sub FindBestCurveForPoint(ByVal curves As DrawingCurvesEnumerator,
                                          ByVal targetPoint As Point2d,
                                          ByRef bestCurve As DrawingCurve,
                                          ByRef bestCurvePoint As Point2d)
            Dim bestDistanceSquared As Double = Double.MaxValue
            bestCurve = Nothing
            bestCurvePoint = Nothing

            For i As Integer = 1 To curves.Count
                Dim curve As DrawingCurve = Nothing
                Try
                    curve = curves.Item(i)
                Catch
                    curve = Nothing
                End Try

                If curve Is Nothing Then
                    Continue For
                End If

                Dim nearestPointOnCurve As Point2d = GetNearestSampledPoint(curve, targetPoint)
                If nearestPointOnCurve Is Nothing Then
                    Continue For
                End If

                Dim d2 As Double = DistanceSquared(nearestPointOnCurve, targetPoint)
                If d2 < bestDistanceSquared Then
                    bestDistanceSquared = d2
                    bestCurve = curve
                    bestCurvePoint = nearestPointOnCurve
                End If
            Next

            If bestCurvePoint Is Nothing Then
                bestCurvePoint = targetPoint
            End If
        End Sub

        Private Function GetNearestSampledPoint(ByVal curve As DrawingCurve, ByVal targetPoint As Point2d) As Point2d
            Dim bestPoint As Point2d = Nothing
            Dim bestDistanceSquared As Double = Double.MaxValue

            Dim samples As New List(Of Point2d)()
            AddCurvePoints(curve, samples)

            For Each sample As Point2d In samples
                If sample Is Nothing Then Continue For
                Dim d2 As Double = DistanceSquared(sample, targetPoint)
                If d2 < bestDistanceSquared Then
                    bestDistanceSquared = d2
                    bestPoint = sample
                End If
            Next

            Return bestPoint
        End Function

        Private Function CreateIntent(ByVal sheet As Sheet,
                                      ByVal curve As DrawingCurve,
                                      ByVal intentPoint As Point2d) As GeometryIntent
            Try
                If intentPoint IsNot Nothing Then
                    Return sheet.CreateGeometryIntent(curve, intentPoint)
                End If
            Catch
            End Try

            Try
                Return sheet.CreateGeometryIntent(curve)
            Catch
            End Try

            Return Nothing
        End Function

        Private Function OffsetPoint(ByVal sourcePoint As Point2d,
                                     ByVal side As BalloonSide,
                                     ByVal offsetValue As Double) As Point2d
            Dim x As Double = sourcePoint.X
            Dim y As Double = sourcePoint.Y

            Select Case side
                Case BalloonSide.LeftSide
                    x -= offsetValue
                Case BalloonSide.RightSide
                    x += offsetValue
                Case BalloonSide.TopSide
                    y += offsetValue
                Case BalloonSide.BottomSide
                    y -= offsetValue
            End Select

            Return m_InventorApp.TransientGeometry.CreatePoint2d(x, y)
        End Function

        Private Function AlignElbowToEndpoint(ByVal elbowPoint As Point2d,
                                              ByVal endPoint As Point2d,
                                              ByVal side As BalloonSide) As Point2d
            Select Case side
                Case BalloonSide.LeftSide, BalloonSide.RightSide
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(elbowPoint.X, endPoint.Y)
                Case Else
                    Return m_InventorApp.TransientGeometry.CreatePoint2d(endPoint.X, elbowPoint.Y)
            End Select
        End Function

        Private Function ApplyEndpointSpacing(ByVal endpoint As Point2d,
                                              ByVal side As BalloonSide,
                                              ByVal occupiedEndPoints As List(Of Point2d),
                                              ByVal minSpacing As Double) As Point2d
            Dim adjusted As Point2d = endpoint
            Dim stepSize As Double = Math.Max(0.25, minSpacing * 0.5)

            For attempt As Integer = 0 To 14
                Dim collision As Boolean = False
                For Each existing As Point2d In occupiedEndPoints
                    If existing Is Nothing Then Continue For
                    If DistanceSquared(adjusted, existing) < (minSpacing * minSpacing) Then
                        collision = True
                        Exit For
                    End If
                Next

                If Not collision Then
                    Return adjusted
                End If

                Dim direction As Double = If((attempt Mod 2) = 0, 1.0, -1.0)
                Dim magnitude As Double = (Math.Floor(attempt / 2.0) + 1.0) * stepSize

                Select Case side
                    Case BalloonSide.LeftSide, BalloonSide.RightSide
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(endpoint.X, endpoint.Y + (direction * magnitude))
                    Case Else
                        adjusted = m_InventorApp.TransientGeometry.CreatePoint2d(endpoint.X + (direction * magnitude), endpoint.Y)
                End Select
            Next

            Return adjusted
        End Function

        Private Function AddBalloon(ByVal sheet As Sheet,
                                    ByVal preferredStyle As BalloonStyle,
                                    ByVal attachIntent As GeometryIntent,
                                    ByVal endPoint As Point2d,
                                    ByVal elbowPoint As Point2d,
                                    ByRef createdBalloon As Balloon) As Boolean
            Try
                Dim leaderPoints As ObjectCollection = m_InventorApp.TransientObjects.CreateObjectCollection()
                leaderPoints.Add(endPoint)
                leaderPoints.Add(elbowPoint)
                leaderPoints.Add(attachIntent)

                Dim balloon As Balloon = Nothing

                If preferredStyle IsNot Nothing Then
                    Try
                        balloon = sheet.Balloons.Add(leaderPoints, , , , preferredStyle)
                    Catch
                        balloon = Nothing
                    End Try
                End If

                If balloon Is Nothing Then
                    balloon = sheet.Balloons.Add(leaderPoints)
                End If

                createdBalloon = balloon
                Return balloon IsNot Nothing
            Catch
                createdBalloon = Nothing
                Return False
            End Try
        End Function

        Private Function TryAddVertexLeader(ByVal targetBalloon As Balloon,
                                            ByVal attachIntent As GeometryIntent,
                                            ByVal endPoint As Point2d,
                                            ByVal elbowPoint As Point2d) As Boolean
            If targetBalloon Is Nothing OrElse attachIntent Is Nothing Then
                Return False
            End If

            Try
                Dim leaderObj As Object = targetBalloon.Leader
                If leaderObj Is Nothing Then Return False

                Dim rootNode As Object = leaderObj.RootNode
                If rootNode Is Nothing Then Return False

                Dim childNodes As Object = rootNode.ChildNodes
                If childNodes Is Nothing Then Return False

                Try
                    CallByName(childNodes, "Add", CallType.Method, endPoint, elbowPoint, attachIntent)
                    Return True
                Catch
                End Try

                Try
                    CallByName(childNodes, "Add", CallType.Method, endPoint, attachIntent)
                    Return True
                Catch
                End Try

                Try
                    CallByName(leaderObj, "AddLeader", CallType.Method, endPoint, elbowPoint, attachIntent)
                    Return True
                Catch
                End Try

                Try
                    CallByName(leaderObj, "AddLeader", CallType.Method, endPoint, attachIntent)
                    Return True
                Catch
                End Try
            Catch
            End Try

            Return False
        End Function

        Private Shared Function PlacementModeLabel(ByVal mode As BalloonPlacementMode) As String
            Select Case mode
                Case BalloonPlacementMode.AutoQuadrant
                    Return "Auto (quadrant from view center)"
                Case BalloonPlacementMode.HorizontalOnly
                    Return "Horizontal only (left/right)"
                Case BalloonPlacementMode.VerticalOnly
                    Return "Vertical only (top/bottom)"
                Case BalloonPlacementMode.ForceRight
                    Return "Force right"
                Case BalloonPlacementMode.ForceLeft
                    Return "Force left"
                Case BalloonPlacementMode.ForceTop
                    Return "Force top"
                Case BalloonPlacementMode.ForceBottom
                    Return "Force bottom"
                Case Else
                    Return "Auto (quadrant from view center)"
            End Select
        End Function

        Private Shared Function PlacementModeFromLabel(ByVal label As String) As BalloonPlacementMode
            For Each mode As BalloonPlacementMode In [Enum].GetValues(GetType(BalloonPlacementMode))
                If String.Equals(PlacementModeLabel(mode), label, StringComparison.OrdinalIgnoreCase) Then
                    Return mode
                End If
            Next

            Return BalloonPlacementMode.AutoQuadrant
        End Function

        Private Shared Sub LoadPersistedSettingsIfNeeded()
            If s_HasLoadedPersistedSettings Then
                Return
            End If

            s_HasLoadedPersistedSettings = True

            Try
                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(AutoBalloonerRegistryPath, False)
                    If key Is Nothing Then
                        Return
                    End If

                    s_LastLeaderDistanceMultiplier = RegistryToDecimal(key.GetValue("LeaderDistanceMultiplier"), CDec(DefaultLeaderDistanceMultiplier))
                    s_LastElbowDistanceRatio = RegistryToDecimal(key.GetValue("ElbowDistanceRatio"), CDec(DefaultElbowDistanceRatio))
                    s_LastMinSpacingRatio = RegistryToDecimal(key.GetValue("MinSpacingRatio"), CDec(DefaultMinSpacingRatio))
                    s_LastProcessAllSheets = RegistryToBoolean(key.GetValue("ProcessAllSheets"), False)
                    s_LastReplaceExistingBalloons = RegistryToBoolean(key.GetValue("ReplaceExistingBalloons"), False)
                    s_LastMinimizeLeaderCrossing = RegistryToBoolean(key.GetValue("MinimizeLeaderCrossing"), True)
                    s_LastCombineSamePartMultiLeader = RegistryToBoolean(key.GetValue("CombineSamePartMultiLeader"), False)

                    Dim placementModeRaw As Object = key.GetValue("PlacementMode")
                    Dim placementModeValue As Integer
                    If placementModeRaw IsNot Nothing AndAlso Integer.TryParse(Convert.ToString(placementModeRaw, CultureInfo.InvariantCulture), placementModeValue) Then
                        If [Enum].IsDefined(GetType(BalloonPlacementMode), placementModeValue) Then
                            s_LastPlacementMode = CType(placementModeValue, BalloonPlacementMode)
                        End If
                    End If
                End Using
            Catch
            End Try
        End Sub

        Private Shared Sub SavePersistedSettings()
            Try
                Using key As RegistryKey = Registry.CurrentUser.CreateSubKey(AutoBalloonerRegistryPath)
                    If key Is Nothing Then
                        Return
                    End If

                    key.SetValue("LeaderDistanceMultiplier", Convert.ToString(s_LastLeaderDistanceMultiplier, CultureInfo.InvariantCulture), RegistryValueKind.String)
                    key.SetValue("ElbowDistanceRatio", Convert.ToString(s_LastElbowDistanceRatio, CultureInfo.InvariantCulture), RegistryValueKind.String)
                    key.SetValue("MinSpacingRatio", Convert.ToString(s_LastMinSpacingRatio, CultureInfo.InvariantCulture), RegistryValueKind.String)
                    key.SetValue("PlacementMode", CInt(s_LastPlacementMode), RegistryValueKind.DWord)
                    key.SetValue("ProcessAllSheets", If(s_LastProcessAllSheets, 1, 0), RegistryValueKind.DWord)
                    key.SetValue("ReplaceExistingBalloons", If(s_LastReplaceExistingBalloons, 1, 0), RegistryValueKind.DWord)
                    key.SetValue("MinimizeLeaderCrossing", If(s_LastMinimizeLeaderCrossing, 1, 0), RegistryValueKind.DWord)
                    key.SetValue("CombineSamePartMultiLeader", If(s_LastCombineSamePartMultiLeader, 1, 0), RegistryValueKind.DWord)
                End Using
            Catch
            End Try
        End Sub

        Private Shared Function RegistryToDecimal(ByVal rawValue As Object, ByVal fallbackValue As Decimal) As Decimal
            If rawValue Is Nothing Then
                Return fallbackValue
            End If

            Dim parsed As Decimal
            If Decimal.TryParse(Convert.ToString(rawValue, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, parsed) Then
                Return parsed
            End If

            Return fallbackValue
        End Function

        Private Shared Function RegistryToBoolean(ByVal rawValue As Object, ByVal fallbackValue As Boolean) As Boolean
            If rawValue Is Nothing Then
                Return fallbackValue
            End If

            Dim numberValue As Integer
            If Integer.TryParse(Convert.ToString(rawValue, CultureInfo.InvariantCulture), numberValue) Then
                Return numberValue <> 0
            End If

            Dim boolValue As Boolean
            If Boolean.TryParse(Convert.ToString(rawValue, CultureInfo.InvariantCulture), boolValue) Then
                Return boolValue
            End If

            Return fallbackValue
        End Function

        Private Class AutoBalloonerSettingsForm
            Inherits Form

            Private ReadOnly m_PlacementCombo As ComboBox
            Private ReadOnly m_LeaderDistanceNumeric As NumericUpDown
            Private ReadOnly m_ElbowRatioNumeric As NumericUpDown
            Private ReadOnly m_MinSpacingRatioNumeric As NumericUpDown
            Private ReadOnly m_ProcessAllSheetsCheck As CheckBox
            Private ReadOnly m_ReplaceExistingCheck As CheckBox
            Private ReadOnly m_MinimizeCrossingCheck As CheckBox
            Private ReadOnly m_CombineSamePartCheck As CheckBox

            Public Sub New()
                LoadPersistedSettingsIfNeeded()

                Me.Text = "Auto Ballooner Settings"
                Me.Width = 560
                Me.Height = 430
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.FormBorderStyle = FormBorderStyle.FixedDialog
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.AutoScaleMode = AutoScaleMode.Font

                Dim introLabel As New Label() With {
                    .Left = 14,
                    .Top = 12,
                    .Width = 528,
                    .Height = 24,
                    .Text = "Configure leader distance and placement behavior for this run."
                }

                Dim placementLabel As New Label() With {
                    .Left = 14,
                    .Top = 44,
                    .Width = 220,
                    .Text = "Placement mode"
                }

                m_PlacementCombo = New ComboBox() With {
                    .Left = 250,
                    .Top = 40,
                    .Width = 292,
                    .DropDownStyle = ComboBoxStyle.DropDownList
                }

                For Each mode As BalloonPlacementMode In [Enum].GetValues(GetType(BalloonPlacementMode))
                    m_PlacementCombo.Items.Add(PlacementModeLabel(mode))
                Next

                Dim selectedLabel As String = PlacementModeLabel(s_LastPlacementMode)
                Dim selectedIndex As Integer = Math.Max(0, m_PlacementCombo.FindStringExact(selectedLabel))
                m_PlacementCombo.SelectedIndex = selectedIndex

                Dim leaderLabel As New Label() With {
                    .Left = 14,
                    .Top = 80,
                    .Width = 220,
                    .Text = "Leader distance multiplier"
                }

                m_LeaderDistanceNumeric = New NumericUpDown() With {
                    .Left = 250,
                    .Top = 76,
                    .Width = 100,
                    .DecimalPlaces = 2,
                    .Increment = 0.25D,
                    .Minimum = 2D,
                    .Maximum = 30D,
                    .Value = Math.Min(30D, Math.Max(2D, s_LastLeaderDistanceMultiplier))
                }

                Dim leaderHint As New Label() With {
                    .Left = 360,
                    .Top = 80,
                    .Width = 142,
                    .Text = "Default: " & DefaultLeaderDistanceMultiplier.ToString("0.##")
                }

                Dim elbowLabel As New Label() With {
                    .Left = 14,
                    .Top = 116,
                    .Width = 220,
                    .Text = "Elbow distance ratio"
                }

                m_ElbowRatioNumeric = New NumericUpDown() With {
                    .Left = 250,
                    .Top = 112,
                    .Width = 100,
                    .DecimalPlaces = 2,
                    .Increment = 0.05D,
                    .Minimum = 0.2D,
                    .Maximum = 1.5D,
                    .Value = Math.Min(1.5D, Math.Max(0.2D, s_LastElbowDistanceRatio))
                }

                Dim elbowHint As New Label() With {
                    .Left = 360,
                    .Top = 116,
                    .Width = 142,
                    .Text = "Default: " & DefaultElbowDistanceRatio.ToString("0.##")
                }

                Dim spacingLabel As New Label() With {
                    .Left = 14,
                    .Top = 152,
                    .Width = 220,
                    .Text = "Minimum spacing ratio"
                }

                m_MinSpacingRatioNumeric = New NumericUpDown() With {
                    .Left = 250,
                    .Top = 148,
                    .Width = 100,
                    .DecimalPlaces = 2,
                    .Increment = 0.05D,
                    .Minimum = 0.3D,
                    .Maximum = 2.0D,
                    .Value = Math.Min(2.0D, Math.Max(0.3D, s_LastMinSpacingRatio))
                }

                Dim spacingHint As New Label() With {
                    .Left = 360,
                    .Top = 152,
                    .Width = 142,
                    .Text = "Default: " & DefaultMinSpacingRatio.ToString("0.##")
                }

                Dim okButton As New Button() With {
                    .Text = "Run",
                    .Left = 376,
                    .Top = 344,
                    .Width = 80
                }
                AddHandler okButton.Click, AddressOf OnRun

                Dim cancelButton As New Button() With {
                    .Text = "Cancel",
                    .Left = 462,
                    .Top = 344,
                    .Width = 80
                }
                AddHandler cancelButton.Click, AddressOf OnCancel

                Dim advancedGroup As New GroupBox() With {
                    .Text = "Advanced",
                    .Left = 14,
                    .Top = 186,
                    .Width = 528,
                    .Height = 144
                }

                m_ProcessAllSheetsCheck = New CheckBox() With {
                    .Left = 12,
                    .Top = 20,
                    .Width = 240,
                    .Text = "Process all sheets",
                    .Checked = s_LastProcessAllSheets
                }

                m_ReplaceExistingCheck = New CheckBox() With {
                    .Left = 270,
                    .Top = 20,
                    .Width = 240,
                    .Text = "Replace existing balloons",
                    .Checked = s_LastReplaceExistingBalloons
                }

                m_MinimizeCrossingCheck = New CheckBox() With {
                    .Left = 12,
                    .Top = 44,
                    .Width = 240,
                    .Text = "Minimize leader crossing",
                    .Checked = s_LastMinimizeLeaderCrossing
                }

                m_CombineSamePartCheck = New CheckBox() With {
                    .Left = 270,
                    .Top = 44,
                    .Width = 240,
                    .Text = "Combine same-part balloons (Part Number)",
                    .Checked = s_LastCombineSamePartMultiLeader
                }

                Dim advancedHint As New Label() With {
                    .Left = 12,
                    .Top = 72,
                    .Width = 504,
                    .Height = 44,
                    .Text = "Replace deletes existing balloons first. Combine adds extra vertex leaders to one balloon per Part Number."
                }

                advancedGroup.Controls.Add(m_ProcessAllSheetsCheck)
                advancedGroup.Controls.Add(m_ReplaceExistingCheck)
                advancedGroup.Controls.Add(m_MinimizeCrossingCheck)
                advancedGroup.Controls.Add(m_CombineSamePartCheck)
                advancedGroup.Controls.Add(advancedHint)

                Me.AcceptButton = okButton
                Me.CancelButton = cancelButton

                Me.Controls.Add(introLabel)
                Me.Controls.Add(placementLabel)
                Me.Controls.Add(m_PlacementCombo)
                Me.Controls.Add(leaderLabel)
                Me.Controls.Add(m_LeaderDistanceNumeric)
                Me.Controls.Add(leaderHint)
                Me.Controls.Add(elbowLabel)
                Me.Controls.Add(m_ElbowRatioNumeric)
                Me.Controls.Add(elbowHint)
                Me.Controls.Add(spacingLabel)
                Me.Controls.Add(m_MinSpacingRatioNumeric)
                Me.Controls.Add(spacingHint)
                Me.Controls.Add(advancedGroup)
                Me.Controls.Add(okButton)
                Me.Controls.Add(cancelButton)
            End Sub

            Private Sub OnRun(ByVal sender As Object, ByVal e As EventArgs)
                s_LastLeaderDistanceMultiplier = m_LeaderDistanceNumeric.Value
                s_LastElbowDistanceRatio = m_ElbowRatioNumeric.Value
                s_LastMinSpacingRatio = m_MinSpacingRatioNumeric.Value
                s_LastPlacementMode = PlacementModeFromLabel(Convert.ToString(m_PlacementCombo.SelectedItem))
                s_LastProcessAllSheets = m_ProcessAllSheetsCheck.Checked
                s_LastReplaceExistingBalloons = m_ReplaceExistingCheck.Checked
                s_LastMinimizeLeaderCrossing = m_MinimizeCrossingCheck.Checked
                s_LastCombineSamePartMultiLeader = m_CombineSamePartCheck.Checked
                SavePersistedSettings()

                Me.DialogResult = DialogResult.OK
                Me.Close()
            End Sub

            Private Sub OnCancel(ByVal sender As Object, ByVal e As EventArgs)
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            End Sub

            Public Function GetConfig() As AutoBalloonerConfig
                Return New AutoBalloonerConfig() With {
                    .LeaderDistanceMultiplier = CDbl(s_LastLeaderDistanceMultiplier),
                    .ElbowDistanceRatio = CDbl(s_LastElbowDistanceRatio),
                    .MinSpacingRatio = CDbl(s_LastMinSpacingRatio),
                    .PlacementMode = s_LastPlacementMode,
                    .ProcessAllSheets = s_LastProcessAllSheets,
                    .ReplaceExistingBalloons = s_LastReplaceExistingBalloons,
                    .MinimizeLeaderCrossing = s_LastMinimizeLeaderCrossing,
                    .CombineSamePartMultiLeader = s_LastCombineSamePartMultiLeader
                }
            End Function
        End Class

        Private Function DistanceSquared(ByVal pt1 As Point2d, ByVal pt2 As Point2d) As Double
            Dim dx As Double = pt1.X - pt2.X
            Dim dy As Double = pt1.Y - pt2.Y
            Return (dx * dx) + (dy * dy)
        End Function

    End Class

End Namespace
