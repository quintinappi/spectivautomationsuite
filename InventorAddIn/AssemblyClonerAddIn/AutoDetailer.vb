Imports Inventor
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.CompilerServices

Namespace AssemblyClonerAddIn

    Public Class AutoDetailer
        Private ReadOnly m_InventorApp As Inventor.Application

        Private Class EndpointIntent
            Public Property SheetPoint As Point2d
            Public Property Intent As GeometryIntent
        End Class

        Private Class HoleCenterIntent
            Public Property SheetPoint As Point2d
            Public Property Intent As GeometryIntent
            Public Property Curve As DrawingCurve
        End Class

        Private Class LinearCurveInfo
            Public Property Curve As DrawingCurve
            Public Property StartPoint As Point2d
            Public Property EndPoint As Point2d
            Public Property MidPoint As Point2d
            Public Property Length As Double
        End Class

        Private Class OverallPairCandidate
            Public Property FirstEndpoint As EndpointIntent
            Public Property SecondEndpoint As EndpointIntent
            Public Property PrimarySpan As Double
            Public Property SecondarySpread As Double
            Public Property ExtremeScore As Double
            Public Property SideBand As Integer
            Public Property SideDistance As Double
        End Class

        Private Class FeatureDimensionSummary
            Public Property TotalCurves As Integer
            Public Property CircularCurves As Integer
            Public Property InteriorCandidates As Integer
            Public Property DimensionsAdded As Integer
            Public Property HoleLeaderNotesAdded As Integer
            Public Property ChamferLeaderNotesAdded As Integer
            Public Property IsAssemblyView As Boolean
            Public Property IsPartView As Boolean
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Open an IDW drawing first.", "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
            Dim sheet As Sheet = drawingDoc.ActiveSheet

            If sheet Is Nothing OrElse sheet.DrawingViews.Count = 0 Then
                MessageBox.Show("No drawing views found on the active sheet.", "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim targetView As DrawingView = SelectTargetView(sheet)
            If targetView Is Nothing Then
                MessageBox.Show("No drawing view selected.", "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim allEndpoints As New List(Of EndpointIntent)()

            For Each curve As DrawingCurve In targetView.DrawingCurves
                For Each segment As DrawingCurveSegment In curve.Segments
                    CaptureEndpoint(sheet, segment, True, allEndpoints)
                    CaptureEndpoint(sheet, segment, False, allEndpoints)
                Next
            Next

            If allEndpoints.Count < 2 Then
                MessageBox.Show("Could not resolve enough geometry points for automatic dimensions.", "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim minX As Double = allEndpoints.Min(Function(pointEntry) pointEntry.SheetPoint.X)
            Dim maxX As Double = allEndpoints.Max(Function(pointEntry) pointEntry.SheetPoint.X)
            Dim minY As Double = allEndpoints.Min(Function(pointEntry) pointEntry.SheetPoint.Y)
            Dim maxY As Double = allEndpoints.Max(Function(pointEntry) pointEntry.SheetPoint.Y)

            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim dims As GeneralDimensions = sheet.DrawingDimensions.GeneralDimensions

            Dim xSpan As Double = Math.Abs(maxX - minX)
            Dim ySpan As Double = Math.Abs(maxY - minY)
            Dim baseOffset As Double = Math.Max(xSpan, ySpan) * 0.08
            If baseOffset <= 0.0001 Then
                baseOffset = 2.0
            End If

            Dim candidatePool As Integer = 24
            Dim minXCandidates As List(Of EndpointIntent) = allEndpoints.OrderBy(Function(pointEntry) pointEntry.SheetPoint.X).Take(candidatePool).ToList()
            Dim maxXCandidates As List(Of EndpointIntent) = allEndpoints.OrderByDescending(Function(pointEntry) pointEntry.SheetPoint.X).Take(candidatePool).ToList()
            Dim minYCandidates As List(Of EndpointIntent) = allEndpoints.OrderBy(Function(pointEntry) pointEntry.SheetPoint.Y).Take(candidatePool).ToList()
            Dim maxYCandidates As List(Of EndpointIntent) = allEndpoints.OrderByDescending(Function(pointEntry) pointEntry.SheetPoint.Y).Take(candidatePool).ToList()

            Dim tx As Transaction = Nothing
            Dim dimensionDiagnostics As New List(Of String)()
            Dim featureDimensionsAdded As Integer = 0
            Dim featureSummary As FeatureDimensionSummary = Nothing
            Try
                tx = m_InventorApp.TransactionManager.StartTransaction(drawingDoc, "Auto Detail IDW")

                Dim horizontalFailureDetails As String = String.Empty
                Dim verticalFailureDetails As String = String.Empty

                Dim horizontalPlaced As Boolean = TryAddOverallDimension(
                    dims,
                    minXCandidates,
                    maxXCandidates,
                    True,
                    minX,
                    maxX,
                    minY,
                    maxY,
                    baseOffset,
                    horizontalFailureDetails,
                    dimensionDiagnostics)

                Dim verticalPlaced As Boolean = TryAddOverallDimension(
                    dims,
                    minYCandidates,
                    maxYCandidates,
                    False,
                    minX,
                    maxX,
                    minY,
                    maxY,
                    baseOffset,
                    verticalFailureDetails,
                    dimensionDiagnostics)

                If Not horizontalPlaced OrElse Not verticalPlaced Then
                    Throw New InvalidOperationException(
                        "Auto-dimension placement failed." & vbCrLf &
                        "Horizontal: " & horizontalFailureDetails & vbCrLf &
                        "Vertical: " & verticalFailureDetails)
                End If

                featureSummary = AddHoleAndCutoutDimensions(
                    sheet,
                    targetView,
                    dims,
                    minXCandidates,
                    maxXCandidates,
                    minYCandidates,
                    maxYCandidates,
                    minX,
                    maxX,
                    minY,
                    maxY,
                    baseOffset,
                    dimensionDiagnostics)
                featureDimensionsAdded = featureSummary.DimensionsAdded

                tx.End()
            Catch ex As Exception
                If tx IsNot Nothing Then
                    tx.Abort()
                End If

                Dim logPath As String = WriteFailureLog(drawingDoc, sheet, ex, dimensionDiagnostics)
                Dim messageText As String = "Failed to place auto dimensions: " & ex.Message
                If Not String.IsNullOrWhiteSpace(logPath) Then
                    messageText &= vbCrLf & vbCrLf & "Diagnostic log: " & logPath
                End If

                MessageBox.Show(messageText,
                                "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            Dim completionMessage As String = "Auto detailing complete."
            If featureSummary IsNot Nothing Then
                completionMessage &= " Overall dimensions placed, " &
                                     featureSummary.DimensionsAdded.ToString() & " linear feature dimensions, " &
                                     featureSummary.HoleLeaderNotesAdded.ToString() & " hole notes, and " &
                                     featureSummary.ChamferLeaderNotesAdded.ToString() & " chamfer notes added for the selected view."
            Else
                completionMessage &= " Overall dimensions and " & featureDimensionsAdded.ToString() & " feature dimensions added for the selected view."
            End If

            If featureSummary IsNot Nothing AndAlso
               featureSummary.DimensionsAdded = 0 AndAlso
               featureSummary.HoleLeaderNotesAdded = 0 AndAlso
               featureSummary.ChamferLeaderNotesAdded = 0 Then
                Dim runLogPath As String = WriteRunLog(drawingDoc, sheet, featureSummary, dimensionDiagnostics)
                If Not String.IsNullOrWhiteSpace(runLogPath) Then
                    completionMessage &= vbCrLf & vbCrLf & "Diagnostic log: " & runLogPath
                End If
            End If

            MessageBox.Show(completionMessage,
                            "Auto Detail IDW", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function SelectTargetView(ByVal sheet As Sheet) As DrawingView
            Dim selectedObj As Object = Nothing

            Try
                selectedObj = m_InventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select drawing view to auto detail")
            Catch
                selectedObj = Nothing
            End Try

            If selectedObj Is Nothing Then
                If sheet IsNot Nothing AndAlso sheet.DrawingViews.Count > 0 Then
                    Return sheet.DrawingViews.Item(1)
                End If
                Return Nothing
            End If

            If TypeOf selectedObj Is DrawingView Then
                Return CType(selectedObj, DrawingView)
            End If

            Return Nothing
        End Function

        Private Function AddHoleAndCutoutDimensions(
            ByVal sheet As Sheet,
            ByVal targetView As DrawingView,
            ByVal dims As GeneralDimensions,
            ByVal minXCandidates As List(Of EndpointIntent),
            ByVal maxXCandidates As List(Of EndpointIntent),
            ByVal minYCandidates As List(Of EndpointIntent),
            ByVal maxYCandidates As List(Of EndpointIntent),
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal baseOffset As Double,
            ByVal diagnostics As List(Of String)) As FeatureDimensionSummary

            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim referencedDoc As Document = GetReferencedDocument(targetView)
            Dim isAssemblyView As Boolean = (referencedDoc IsNot Nothing AndAlso referencedDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject)
            Dim isPartView As Boolean = (referencedDoc IsNot Nothing AndAlso referencedDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject)

            Dim summary As New FeatureDimensionSummary With {
                .TotalCurves = 0,
                .CircularCurves = 0,
                .InteriorCandidates = 0,
                .DimensionsAdded = 0,
                .HoleLeaderNotesAdded = 0,
                .ChamferLeaderNotesAdded = 0,
                .IsAssemblyView = isAssemblyView,
                .IsPartView = isPartView
            }
            Dim viewCenterX As Double = (minX + maxX) / 2.0
            Dim viewCenterY As Double = (minY + maxY) / 2.0
            Dim boundTolerance As Double = Math.Max(baseOffset * 0.08, 0.08)
            Dim textOffset As Double = Math.Max(baseOffset * 0.9, 0.5)
            Dim noteOffset As Double = Math.Max(baseOffset * 1.2, 0.9)
            Dim laneSpacing As Double = Math.Max(baseOffset * 0.35, 0.3)
            Dim holeMergeTolerance As Double = Math.Max(baseOffset * 0.12, 0.08)
            Dim holeCenters As New List(Of HoleCenterIntent)()
            Dim usedDimKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ApplyAutomatedCenterlineSettings(targetView, isAssemblyView, diagnostics)

            For Each curve As DrawingCurve In targetView.DrawingCurves
                summary.TotalCurves += 1
                Dim isCircle As Boolean = (curve.CurveType = CurveTypeEnum.kCircleCurve)
                Dim isArc As Boolean = (curve.CurveType = CurveTypeEnum.kCircularArcCurve)

                If Not isCircle AndAlso Not isArc Then
                    Continue For
                End If

                summary.CircularCurves += 1

                Dim center As Point2d = Nothing
                Try
                    center = curve.CenterPoint
                Catch
                    center = Nothing
                End Try

                If center Is Nothing Then
                    Continue For
                End If

                If Not IsPointInsideBounds(center, minX, maxX, minY, maxY, -boundTolerance) Then
                    Continue For
                End If

                summary.InteriorCandidates += 1

                Dim intent As GeometryIntent = TryCreateCurveCenterIntent(sheet, curve, center)
                AddHoleCenterCandidateUnique(holeCenters, center, intent, curve, holeMergeTolerance)
            Next

            Dim centerMarkCandidates As Integer = CollectHoleCentersFromCentermarks(
                sheet,
                targetView,
                holeCenters,
                minX,
                maxX,
                minY,
                maxY,
                boundTolerance,
                holeMergeTolerance)

            Dim sortedHoles As List(Of HoleCenterIntent) = holeCenters.
                Where(Function(hole) hole IsNot Nothing AndAlso hole.SheetPoint IsNot Nothing).
                OrderBy(Function(hole) hole.SheetPoint.X).
                ThenBy(Function(hole) hole.SheetPoint.Y).
                ToList()

            Dim maxHoleNotes As Integer = If(isAssemblyView, 8, 16)
            Dim leftLaneY As New List(Of Double)()
            Dim rightLaneY As New List(Of Double)()
            Dim seenHoleNoteTypeKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim holeNotePlacementIndex As Integer = 0

            For Each hole As HoleCenterIntent In sortedHoles
                If summary.HoleLeaderNotesAdded >= maxHoleNotes Then
                    Exit For
                End If

                holeNotePlacementIndex += 1

                Dim holeNotePoint As Point2d = BuildOutsideLeaderPoint(
                    hole.SheetPoint,
                    minX,
                    maxX,
                    minY,
                    maxY,
                    noteOffset,
                    laneSpacing,
                    leftLaneY,
                    rightLaneY,
                    True,
                    holeNotePlacementIndex)

                If TryAddHoleThreadLeaderNote(sheet, hole, holeNotePoint, seenHoleNoteTypeKeys, diagnostics) Then
                    summary.HoleLeaderNotesAdded += 1
                End If
            Next

            Dim leftRef As EndpointIntent = If(minXCandidates IsNot Nothing AndAlso minXCandidates.Count > 0, minXCandidates(0), Nothing)
            Dim rightRef As EndpointIntent = If(maxXCandidates IsNot Nothing AndAlso maxXCandidates.Count > 0, maxXCandidates(0), Nothing)
            Dim bottomRef As EndpointIntent = If(minYCandidates IsNot Nothing AndAlso minYCandidates.Count > 0, minYCandidates(0), Nothing)
            Dim topRef As EndpointIntent = If(maxYCandidates IsNot Nothing AndAlso maxYCandidates.Count > 0, maxYCandidates(0), Nothing)

            Dim addLinearHoleDimensions As Boolean = Not isAssemblyView
            If addLinearHoleDimensions Then
                For Each hole As HoleCenterIntent In sortedHoles
                    If hole Is Nothing OrElse hole.Intent Is Nothing OrElse hole.SheetPoint Is Nothing Then
                        Continue For
                    End If

                    Dim useLeft As Boolean = (Math.Abs(hole.SheetPoint.X - minX) <= Math.Abs(maxX - hole.SheetPoint.X))
                    Dim xRef As EndpointIntent = If(useLeft, leftRef, rightRef)
                    If xRef IsNot Nothing AndAlso xRef.Intent IsNot Nothing Then
                        Dim xTextY As Double = If(hole.SheetPoint.Y >= viewCenterY, maxY + textOffset, minY - textOffset)
                        Dim xTextPoint As Point2d = tg.CreatePoint2d((xRef.SheetPoint.X + hole.SheetPoint.X) / 2.0, xTextY)
                        If TryAddLinearDimensionUnique(dims, xTextPoint, xRef.Intent, hole.Intent, True, usedDimKeys, diagnostics) Then
                            summary.DimensionsAdded += 1
                        End If
                    End If

                    Dim useBottom As Boolean = (Math.Abs(hole.SheetPoint.Y - minY) <= Math.Abs(maxY - hole.SheetPoint.Y))
                    Dim yRef As EndpointIntent = If(useBottom, bottomRef, topRef)
                    If yRef IsNot Nothing AndAlso yRef.Intent IsNot Nothing Then
                        Dim yTextX As Double = If(hole.SheetPoint.X >= viewCenterX, maxX + textOffset, minX - textOffset)
                        Dim yTextPoint As Point2d = tg.CreatePoint2d(yTextX, (yRef.SheetPoint.Y + hole.SheetPoint.Y) / 2.0)
                        If TryAddLinearDimensionUnique(dims, yTextPoint, yRef.Intent, hole.Intent, False, usedDimKeys, diagnostics) Then
                            summary.DimensionsAdded += 1
                        End If
                    End If
                Next

                Dim dimensionableHoles As List(Of HoleCenterIntent) = sortedHoles.
                    Where(Function(hole) hole.Intent IsNot Nothing).
                    ToList()

                If dimensionableHoles.Count > 1 Then
                    Dim sortedByX As List(Of HoleCenterIntent) = dimensionableHoles.OrderBy(Function(hole) hole.SheetPoint.X).ToList()
                    Dim sortedByY As List(Of HoleCenterIntent) = dimensionableHoles.OrderBy(Function(hole) hole.SheetPoint.Y).ToList()

                    For i As Integer = 1 To sortedByX.Count - 1
                        Dim h1 As HoleCenterIntent = sortedByX(i - 1)
                        Dim h2 As HoleCenterIntent = sortedByX(i)
                        Dim textPoint As Point2d = tg.CreatePoint2d((h1.SheetPoint.X + h2.SheetPoint.X) / 2.0, maxY + (textOffset * 1.55))
                        If TryAddLinearDimensionUnique(dims, textPoint, h1.Intent, h2.Intent, True, usedDimKeys, diagnostics) Then
                            summary.DimensionsAdded += 1
                        End If
                    Next

                    For i As Integer = 1 To sortedByY.Count - 1
                        Dim h1 As HoleCenterIntent = sortedByY(i - 1)
                        Dim h2 As HoleCenterIntent = sortedByY(i)
                        Dim textPoint As Point2d = tg.CreatePoint2d(maxX + (textOffset * 1.55), (h1.SheetPoint.Y + h2.SheetPoint.Y) / 2.0)
                        If TryAddLinearDimensionUnique(dims, textPoint, h1.Intent, h2.Intent, False, usedDimKeys, diagnostics) Then
                            summary.DimensionsAdded += 1
                        End If
                    Next
                End If
            End If

            Dim maxChamferNotes As Integer = If(isAssemblyView, 3, 8)
            summary.ChamferLeaderNotesAdded = AddChamferLeaderNotes(
                sheet,
                targetView,
                minX,
                maxX,
                minY,
                maxY,
                baseOffset,
                maxChamferNotes,
                diagnostics)

            diagnostics.Add("Feature summary: docType=" & If(isAssemblyView, "Assembly", If(isPartView, "Part", "Other")) &
                            ", totalCurves=" & summary.TotalCurves.ToString() &
                            ", circularCurves=" & summary.CircularCurves.ToString() &
                            ", curveCandidates=" & summary.InteriorCandidates.ToString() &
                            ", centerMarkCandidates=" & centerMarkCandidates.ToString() &
                            ", linearDimensionsAdded=" & summary.DimensionsAdded.ToString() &
                            ", holeLeaderNotesAdded=" & summary.HoleLeaderNotesAdded.ToString() &
                            ", chamferLeaderNotesAdded=" & summary.ChamferLeaderNotesAdded.ToString())

            Return summary
        End Function

        Private Function GetReferencedDocument(ByVal targetView As DrawingView) As Document
            If targetView Is Nothing Then
                Return Nothing
            End If

            Try
                If targetView.ReferencedDocumentDescriptor IsNot Nothing Then
                    Dim descriptorDoc As Document = targetView.ReferencedDocumentDescriptor.ReferencedDocument
                    If descriptorDoc IsNot Nothing Then
                        Return descriptorDoc
                    End If
                End If
            Catch
            End Try

            Try
                Return targetView.ReferencedDocument
            Catch
            End Try

            Return Nothing
        End Function

        Private Sub ApplyAutomatedCenterlineSettings(
            ByVal targetView As DrawingView,
            ByVal isAssemblyView As Boolean,
            ByVal diagnostics As List(Of String))

            If targetView Is Nothing Then
                Return
            End If

            Try
                Dim settings As AutomatedCenterlineSettings = Nothing
                targetView.GetAutomatedCenterlineSettings(settings)
                If settings Is Nothing Then
                    diagnostics.Add("Automated centerline settings not available for this view.")
                    Return
                End If

                settings.ApplyToHoles = True
                settings.ApplyToCircularPatterns = True
                settings.ApplyToRectangularPatterns = True
                settings.ApplyToPunches = True
                settings.ProjectionParallelAxis = True
                settings.ProjectionNormalAxis = True
                settings.ApplyToBends = False
                settings.ApplyToWorkFeatures = False

                If isAssemblyView Then
                    settings.ApplyToFillets = False
                    settings.ApplyToCylinders = False
                    settings.ApplyToRevolutions = False
                    settings.ApplyToSketches = False
                Else
                    settings.ApplyToFillets = True
                    settings.ApplyToCylinders = True
                    settings.ApplyToRevolutions = True
                    settings.ApplyToSketches = True
                End If

                targetView.SetAutomatedCenterlineSettings(settings)
                diagnostics.Add("Applied automated centerline/centermark settings.")
            Catch ex As Exception
                diagnostics.Add("SetAutomatedCenterlineSettings failed: " & ex.Message)
            End Try
        End Sub

        Private Function TryCreateCurveCenterIntent(
            ByVal sheet As Sheet,
            ByVal curve As DrawingCurve,
            ByVal center As Point2d) As GeometryIntent

            If sheet Is Nothing OrElse curve Is Nothing Then
                Return Nothing
            End If

            Try
                Return sheet.CreateGeometryIntent(curve, PointIntentEnum.kCenterPointIntent)
            Catch
            End Try

            If center IsNot Nothing Then
                Try
                    Return sheet.CreateGeometryIntent(curve, center)
                Catch
                End Try
            End If

            Try
                Return sheet.CreateGeometryIntent(curve, Nothing)
            Catch
            End Try

            Return Nothing
        End Function

        Private Function CollectHoleCentersFromCentermarks(
            ByVal sheet As Sheet,
            ByVal targetView As DrawingView,
            ByVal holeCenters As List(Of HoleCenterIntent),
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal boundTolerance As Double,
            ByVal mergeTolerance As Double) As Integer

            If sheet Is Nothing OrElse targetView Is Nothing OrElse holeCenters Is Nothing Then
                Return 0
            End If

            Dim added As Integer = 0

            For Each centerMark As Centermark In sheet.Centermarks
                If centerMark Is Nothing Then
                    Continue For
                End If

                Dim attachedCurve As DrawingCurve = Nothing
                Try
                    attachedCurve = TryCast(centerMark.AttachedEntity, DrawingCurve)
                Catch
                    attachedCurve = Nothing
                End Try

                If attachedCurve Is Nothing OrElse Not IsCurveInView(attachedCurve, targetView) Then
                    Continue For
                End If

                Dim centerPoint As Point2d = Nothing
                Try
                    centerPoint = centerMark.Position
                Catch
                    centerPoint = Nothing
                End Try

                If centerPoint Is Nothing Then
                    Continue For
                End If

                If Not IsPointInsideBounds(centerPoint, minX, maxX, minY, maxY, -boundTolerance) Then
                    Continue For
                End If

                Dim intent As GeometryIntent = TryCreateCurveCenterIntent(sheet, attachedCurve, centerPoint)
                Dim countBefore As Integer = holeCenters.Count
                AddHoleCenterCandidateUnique(holeCenters, centerPoint, intent, attachedCurve, mergeTolerance)
                If holeCenters.Count > countBefore Then
                    added += 1
                End If
            Next

            Return added
        End Function

        Private Function IsCurveInView(ByVal curve As DrawingCurve, ByVal targetView As DrawingView) As Boolean
            If curve Is Nothing OrElse targetView Is Nothing Then
                Return False
            End If

            Try
                Return curve.Parent Is targetView
            Catch
            End Try

            Try
                Dim curveView As DrawingView = curve.Parent
                If curveView IsNot Nothing Then
                    Return String.Equals(curveView.Name, targetView.Name, StringComparison.OrdinalIgnoreCase)
                End If
            Catch
            End Try

            Return False
        End Function

        Private Sub AddHoleCenterCandidateUnique(
            ByVal holeCenters As List(Of HoleCenterIntent),
            ByVal center As Point2d,
            ByVal intent As GeometryIntent,
            ByVal curve As DrawingCurve,
            ByVal mergeTolerance As Double)

            If holeCenters Is Nothing OrElse center Is Nothing Then
                Return
            End If

            If intent Is Nothing AndAlso curve Is Nothing Then
                Return
            End If

            For Each existing As HoleCenterIntent In holeCenters
                If existing Is Nothing OrElse existing.SheetPoint Is Nothing Then
                    Continue For
                End If

                If Distance2d(existing.SheetPoint, center) <= mergeTolerance Then
                    If existing.Intent Is Nothing AndAlso intent IsNot Nothing Then
                        existing.Intent = intent
                    End If

                    If existing.Curve Is Nothing AndAlso curve IsNot Nothing Then
                        existing.Curve = curve
                    End If
                    Return
                End If
            Next

            holeCenters.Add(New HoleCenterIntent With {
                .SheetPoint = center,
                .Intent = intent,
                .Curve = curve
            })
        End Sub

        Private Function BuildOutsideLeaderPoint(
            ByVal centerPoint As Point2d,
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal noteOffset As Double,
            ByVal laneSpacing As Double,
            ByVal leftLane As List(Of Double),
            ByVal rightLane As List(Of Double),
            Optional ByVal forceAngled As Boolean = False,
            Optional ByVal angleSeed As Integer = 0) As Point2d

            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim useLeft As Boolean = (Math.Abs(centerPoint.X - minX) <= Math.Abs(maxX - centerPoint.X))
            Dim laneMinY As Double = minY - (noteOffset * 0.35)
            Dim laneMaxY As Double = maxY + (noteOffset * 0.35)
            Dim preferredY As Double = centerPoint.Y

            If forceAngled Then
                Dim minAngleDelta As Double = Math.Max(noteOffset * 0.35, laneSpacing * 0.9)
                Dim baseDirection As Integer = If((angleSeed Mod 2) = 0, 1, -1)
                Dim tier As Integer = Math.Max(0, (angleSeed - 1) Mod 3)
                Dim tierOffset As Double = tier * (laneSpacing * 0.35)
                preferredY = centerPoint.Y + (baseDirection * (minAngleDelta + tierOffset))
            End If

            If useLeft Then
                Dim laneY As Double = ReserveLaneCoordinate(preferredY, laneMinY, laneMaxY, leftLane, laneSpacing)
                Return tg.CreatePoint2d(minX - noteOffset, laneY)
            End If

            Dim rightY As Double = ReserveLaneCoordinate(preferredY, laneMinY, laneMaxY, rightLane, laneSpacing)
            Return tg.CreatePoint2d(maxX + noteOffset, rightY)
        End Function

        Private Function ReserveLaneCoordinate(
            ByVal preferred As Double,
            ByVal minValue As Double,
            ByVal maxValue As Double,
            ByVal usedValues As List(Of Double),
            ByVal spacing As Double) As Double

            If usedValues Is Nothing Then
                Return preferred
            End If

            Dim clampedPreferred As Double = Math.Max(minValue, Math.Min(maxValue, preferred))

            If usedValues.Count = 0 Then
                usedValues.Add(clampedPreferred)
                Return clampedPreferred
            End If

            For attempt As Integer = 0 To 24
                Dim stepIndex As Integer = attempt \ 2
                Dim direction As Integer = If((attempt Mod 2) = 0, 1, -1)
                Dim candidate As Double = clampedPreferred + (direction * stepIndex * spacing)
                candidate = Math.Max(minValue, Math.Min(maxValue, candidate))

                Dim isAvailable As Boolean = usedValues.All(Function(existingValue) Math.Abs(existingValue - candidate) >= (spacing * 0.85))
                If isAvailable Then
                    usedValues.Add(candidate)
                    Return candidate
                End If
            Next

            usedValues.Add(clampedPreferred)
            Return clampedPreferred
        End Function

        Private Function TryAddHoleThreadLeaderNote(
            ByVal sheet As Sheet,
            ByVal hole As HoleCenterIntent,
            ByVal notePoint As Point2d,
            ByVal seenTypeKeys As HashSet(Of String),
            ByVal diagnostics As List(Of String)) As Boolean

            If sheet Is Nothing OrElse hole Is Nothing OrElse notePoint Is Nothing Then
                Return False
            End If

            Dim notes As HoleThreadNotes = Nothing
            Try
                notes = sheet.DrawingNotes.HoleThreadNotes
            Catch
                notes = Nothing
            End Try

            If notes Is Nothing Then
                Return False
            End If

            If Double.IsNaN(notePoint.X) OrElse Double.IsNaN(notePoint.Y) OrElse
               Double.IsInfinity(notePoint.X) OrElse Double.IsInfinity(notePoint.Y) Then
                Return False
            End If

            Dim createdNote As HoleThreadNote = Nothing
            Dim edgeIntent As GeometryIntent = Nothing

            If hole.Curve IsNot Nothing Then
                Try
                    edgeIntent = sheet.CreateGeometryIntent(hole.Curve, 0.5R)
                Catch
                    edgeIntent = Nothing
                End Try
            End If

            Try
                If hole.Curve IsNot Nothing Then
                    createdNote = notes.Add(notePoint, hole.Curve, False)
                End If

                If createdNote Is Nothing AndAlso edgeIntent IsNot Nothing Then
                    createdNote = notes.Add(notePoint, edgeIntent, False)
                End If

                If createdNote Is Nothing AndAlso hole.Intent IsNot Nothing Then
                    createdNote = notes.Add(notePoint, hole.Intent, False)
                End If
            Catch firstEx As Exception
                Try
                    If hole.Curve IsNot Nothing Then
                        createdNote = notes.Add(notePoint, hole.Curve, False)
                    ElseIf edgeIntent IsNot Nothing Then
                        createdNote = notes.Add(notePoint, edgeIntent, False)
                    ElseIf hole.Intent IsNot Nothing Then
                        createdNote = notes.Add(notePoint, hole.Intent, False)
                    End If
                Catch secondEx As Exception
                    diagnostics.Add("HoleThreadNotes.Add failed: " & firstEx.Message & " | fallback: " & secondEx.Message)
                    Return False
                End Try
            End Try

            If createdNote Is Nothing Then
                Return False
            End If

            Try
                createdNote.LeaderFromCenter = True
            Catch
            End Try

            Try
                createdNote.SingleDimensionLine = False
            Catch
            End Try

            Try
                If createdNote.Text IsNot Nothing Then
                    createdNote.Text.Origin = notePoint
                End If
            Catch ex As Exception
                diagnostics.Add("HoleThreadNote.Text.Origin set failed: " & ex.Message)
            End Try

            If seenTypeKeys IsNot Nothing Then
                Dim holeTypeKey As String = BuildHoleNoteTypeKey(createdNote, hole)
                If Not String.IsNullOrWhiteSpace(holeTypeKey) Then
                    If seenTypeKeys.Contains(holeTypeKey) Then
                        Try
                            createdNote.Delete()
                        Catch
                        End Try

                        diagnostics.Add("Hole note deduped: " & holeTypeKey)
                        Return False
                    End If

                    seenTypeKeys.Add(holeTypeKey)
                End If
            End If

            Return True
        End Function

        Private Function BuildHoleNoteTypeKey(ByVal note As HoleThreadNote, ByVal hole As HoleCenterIntent) As String
            If note Is Nothing Then
                Return String.Empty
            End If

            Dim formattedCallout As String = String.Empty
            Try
                formattedCallout = NormalizeNoteKey(note.FormattedHoleThreadNote)
            Catch
                formattedCallout = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(formattedCallout) Then
                Try
                    If note.Text IsNot Nothing Then
                        formattedCallout = NormalizeNoteKey(note.Text.Text)
                    End If
                Catch
                    formattedCallout = String.Empty
                End Try
            End If

            If Not String.IsNullOrWhiteSpace(formattedCallout) Then
                Return "HOLE|" & formattedCallout
            End If

            Try
                If hole IsNot Nothing AndAlso hole.Curve IsNot Nothing Then
                    Dim edgeTypeText As String = CInt(hole.Curve.EdgeType).ToString()
                    Dim curveTypeText As String = CInt(hole.Curve.CurveType).ToString()
                    Dim radiusText As String = String.Empty

                    Dim center As Point2d = hole.Curve.CenterPoint
                    If center IsNot Nothing AndAlso hole.SheetPoint IsNot Nothing Then
                        radiusText = Math.Round(Distance2d(center, hole.SheetPoint), 4).ToString()
                    End If

                    Return "HOLE|EDGE=" & edgeTypeText & "|CURVE=" & curveTypeText & "|R=" & radiusText
                End If
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function NormalizeNoteKey(ByVal rawText As String) As String
            If String.IsNullOrWhiteSpace(rawText) Then
                Return String.Empty
            End If

            Dim normalized As String = rawText.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ").Trim()
            Do While normalized.Contains("  ")
                normalized = normalized.Replace("  ", " ")
            Loop

            Return normalized.Trim().ToUpperInvariant()
        End Function

        Private Function AddChamferLeaderNotes(
            ByVal sheet As Sheet,
            ByVal targetView As DrawingView,
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal baseOffset As Double,
            ByVal maxNotes As Integer,
            ByVal diagnostics As List(Of String)) As Integer

            If sheet Is Nothing OrElse targetView Is Nothing OrElse maxNotes <= 0 Then
                Return 0
            End If

            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry
            Dim linearCurves As New List(Of LinearCurveInfo)()

            For Each curve As DrawingCurve In targetView.DrawingCurves
                If curve Is Nothing OrElse curve.CurveType <> CurveTypeEnum.kLineSegmentCurve Then
                    Continue For
                End If

                Dim startPoint As Point2d = Nothing
                Dim endPoint As Point2d = Nothing
                Try
                    startPoint = curve.StartPoint
                    endPoint = curve.EndPoint
                Catch
                    Continue For
                End Try

                If startPoint Is Nothing OrElse endPoint Is Nothing Then
                    Continue For
                End If

                Dim length As Double = Distance2d(startPoint, endPoint)
                If length <= 0.0001 Then
                    Continue For
                End If

                linearCurves.Add(New LinearCurveInfo With {
                    .Curve = curve,
                    .StartPoint = startPoint,
                    .EndPoint = endPoint,
                    .MidPoint = tg.CreatePoint2d((startPoint.X + endPoint.X) / 2.0, (startPoint.Y + endPoint.Y) / 2.0),
                    .Length = length
                })
            Next

            If linearCurves.Count = 0 Then
                Return 0
            End If

            Dim viewSpan As Double = Math.Max(Math.Abs(maxX - minX), Math.Abs(maxY - minY))
            Dim shortEdgeThreshold As Double = Math.Max(viewSpan * 0.12, baseOffset * 1.25)
            Dim joinTolerance As Double = Math.Max(baseOffset * 0.12, 0.06)
            Dim noteOffset As Double = Math.Max(baseOffset * 1.2, 0.9)
            Dim laneSpacing As Double = Math.Max(baseOffset * 0.35, 0.3)
            Dim usedPairs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim seenChamferTypeKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim leftLaneY As New List(Of Double)()
            Dim rightLaneY As New List(Of Double)()
            Dim addedCount As Integer = 0

            Dim shortEdges As List(Of LinearCurveInfo) = linearCurves.
                Where(Function(info) info.Length <= shortEdgeThreshold).
                OrderBy(Function(info) info.Length).
                ToList()

            For Each shortEdge As LinearCurveInfo In shortEdges
                For Each candidate As LinearCurveInfo In linearCurves
                    If addedCount >= maxNotes Then
                        Exit For
                    End If

                    If shortEdge Is candidate Then
                        Continue For
                    End If

                    Dim sharedPoint As Point2d = Nothing
                    If Not TryGetSharedEndpoint(shortEdge, candidate, joinTolerance, sharedPoint) Then
                        Continue For
                    End If

                    Dim angle As Double = ComputeSegmentAngleDegrees(shortEdge, candidate, sharedPoint)
                    If angle < 12.0 OrElse angle > 168.0 Then
                        Continue For
                    End If

                    If Math.Abs(angle - 90.0) <= 8.0 Then
                        Continue For
                    End If

                    Dim pairKey As String = BuildCurvePairKey(shortEdge.Curve, candidate.Curve)
                    If usedPairs.Contains(pairKey) Then
                        Continue For
                    End If

                    Dim anchor As Point2d = tg.CreatePoint2d((shortEdge.MidPoint.X + candidate.MidPoint.X) / 2.0, (shortEdge.MidPoint.Y + candidate.MidPoint.Y) / 2.0)
                    Dim notePoint As Point2d = BuildOutsideLeaderPoint(anchor, minX, maxX, minY, maxY, noteOffset, laneSpacing, leftLaneY, rightLaneY)

                    If TryAddChamferLeaderNote(sheet, notePoint, shortEdge.Curve, candidate.Curve, seenChamferTypeKeys, diagnostics) Then
                        usedPairs.Add(pairKey)
                        addedCount += 1
                    End If
                Next

                If addedCount >= maxNotes Then
                    Exit For
                End If
            Next

            Return addedCount
        End Function

        Private Function TryAddChamferLeaderNote(
            ByVal sheet As Sheet,
            ByVal notePoint As Point2d,
            ByVal firstCurve As DrawingCurve,
            ByVal secondCurve As DrawingCurve,
            ByVal seenTypeKeys As HashSet(Of String),
            ByVal diagnostics As List(Of String)) As Boolean

            If sheet Is Nothing OrElse notePoint Is Nothing OrElse firstCurve Is Nothing OrElse secondCurve Is Nothing Then
                Return False
            End If

            Try
                Dim createdNote As ChamferNote = sheet.DrawingNotes.ChamferNotes.Add(notePoint, firstCurve, secondCurve)

                If seenTypeKeys IsNot Nothing Then
                    Dim chamferTypeKey As String = BuildChamferNoteTypeKey(createdNote)
                    If Not String.IsNullOrWhiteSpace(chamferTypeKey) Then
                        If seenTypeKeys.Contains(chamferTypeKey) Then
                            Try
                                createdNote.Delete()
                            Catch
                            End Try

                            diagnostics.Add("Chamfer note deduped: " & chamferTypeKey)
                            Return False
                        End If

                        seenTypeKeys.Add(chamferTypeKey)
                    End If
                End If

                Return True
            Catch ex As Exception
                If diagnostics IsNot Nothing AndAlso diagnostics.Count < 220 Then
                    diagnostics.Add("ChamferNotes.Add failed: " & ex.Message)
                End If
                Return False
            End Try
        End Function

        Private Function BuildChamferNoteTypeKey(ByVal note As ChamferNote) As String
            If note Is Nothing Then
                Return String.Empty
            End If

            Dim formattedCallout As String = String.Empty
            Try
                formattedCallout = NormalizeNoteKey(note.FormattedChamferNote)
            Catch
                formattedCallout = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(formattedCallout) Then
                Try
                    formattedCallout = NormalizeNoteKey(note.Text)
                Catch
                    formattedCallout = String.Empty
                End Try
            End If

            If String.IsNullOrWhiteSpace(formattedCallout) Then
                Return String.Empty
            End If

            Return "CHAMFER|" & formattedCallout
        End Function

        Private Function TryGetSharedEndpoint(
            ByVal firstCurve As LinearCurveInfo,
            ByVal secondCurve As LinearCurveInfo,
            ByVal tolerance As Double,
            ByRef sharedPoint As Point2d) As Boolean

            If firstCurve Is Nothing OrElse secondCurve Is Nothing Then
                Return False
            End If

            If Distance2d(firstCurve.StartPoint, secondCurve.StartPoint) <= tolerance Then
                sharedPoint = firstCurve.StartPoint
                Return True
            End If

            If Distance2d(firstCurve.StartPoint, secondCurve.EndPoint) <= tolerance Then
                sharedPoint = firstCurve.StartPoint
                Return True
            End If

            If Distance2d(firstCurve.EndPoint, secondCurve.StartPoint) <= tolerance Then
                sharedPoint = firstCurve.EndPoint
                Return True
            End If

            If Distance2d(firstCurve.EndPoint, secondCurve.EndPoint) <= tolerance Then
                sharedPoint = firstCurve.EndPoint
                Return True
            End If

            Return False
        End Function

        Private Function ComputeSegmentAngleDegrees(
            ByVal firstCurve As LinearCurveInfo,
            ByVal secondCurve As LinearCurveInfo,
            ByVal sharedPoint As Point2d) As Double

            If firstCurve Is Nothing OrElse secondCurve Is Nothing OrElse sharedPoint Is Nothing Then
                Return 0.0
            End If

            Dim firstOther As Point2d = If(Distance2d(firstCurve.StartPoint, sharedPoint) <= Distance2d(firstCurve.EndPoint, sharedPoint), firstCurve.EndPoint, firstCurve.StartPoint)
            Dim secondOther As Point2d = If(Distance2d(secondCurve.StartPoint, sharedPoint) <= Distance2d(secondCurve.EndPoint, sharedPoint), secondCurve.EndPoint, secondCurve.StartPoint)

            Dim v1x As Double = firstOther.X - sharedPoint.X
            Dim v1y As Double = firstOther.Y - sharedPoint.Y
            Dim v2x As Double = secondOther.X - sharedPoint.X
            Dim v2y As Double = secondOther.Y - sharedPoint.Y

            Dim len1 As Double = Math.Sqrt((v1x * v1x) + (v1y * v1y))
            Dim len2 As Double = Math.Sqrt((v2x * v2x) + (v2y * v2y))
            If len1 <= 0.000001 OrElse len2 <= 0.000001 Then
                Return 0.0
            End If

            Dim dot As Double = ((v1x * v2x) + (v1y * v2y)) / (len1 * len2)
            dot = Math.Max(-1.0, Math.Min(1.0, dot))

            Return Math.Acos(dot) * (180.0 / Math.PI)
        End Function

        Private Function BuildCurvePairKey(ByVal firstCurve As DrawingCurve, ByVal secondCurve As DrawingCurve) As String
            If firstCurve Is Nothing OrElse secondCurve Is Nothing Then
                Return String.Empty
            End If

            Dim firstHash As Integer = RuntimeHelpers.GetHashCode(firstCurve)
            Dim secondHash As Integer = RuntimeHelpers.GetHashCode(secondCurve)
            Dim a As Integer = Math.Min(firstHash, secondHash)
            Dim b As Integer = Math.Max(firstHash, secondHash)
            Return a.ToString() & "|" & b.ToString()
        End Function

        Private Function Distance2d(ByVal firstPoint As Point2d, ByVal secondPoint As Point2d) As Double
            If firstPoint Is Nothing OrElse secondPoint Is Nothing Then
                Return Double.MaxValue
            End If

            Dim dx As Double = firstPoint.X - secondPoint.X
            Dim dy As Double = firstPoint.Y - secondPoint.Y
            Return Math.Sqrt((dx * dx) + (dy * dy))
        End Function

        Private Function TryAddLinearDimensionUnique(
            ByVal dims As GeneralDimensions,
            ByVal textPoint As Point2d,
            ByVal firstIntent As GeometryIntent,
            ByVal secondIntent As GeometryIntent,
            ByVal isHorizontal As Boolean,
            ByVal usedKeys As HashSet(Of String),
            ByVal diagnostics As List(Of String)) As Boolean

            If firstIntent Is Nothing OrElse secondIntent Is Nothing Then
                Return False
            End If

            Dim firstHash As Integer = RuntimeHelpers.GetHashCode(firstIntent)
            Dim secondHash As Integer = RuntimeHelpers.GetHashCode(secondIntent)
            Dim a As Integer = Math.Min(firstHash, secondHash)
            Dim b As Integer = Math.Max(firstHash, secondHash)
            Dim key As String = If(isHorizontal, "H", "V") & "|" & a.ToString() & "|" & b.ToString()

            If usedKeys.Contains(key) Then
                Return False
            End If

            Dim mode As String = String.Empty
            Dim err As String = String.Empty
            If TryAddLinearDimension(dims, textPoint, firstIntent, secondIntent, isHorizontal, mode, err) Then
                usedKeys.Add(key)
                Return True
            End If

            diagnostics.Add("Feature linear add failed: " & err)
            Return False
        End Function

        Private Function TryAddDiameterDimension(
            ByVal dims As GeneralDimensions,
            ByVal textPoint As Point2d,
            ByVal intent As GeometryIntent,
            ByVal diagnostics As List(Of String)) As Boolean

            Try
                dims.AddDiameter(textPoint, intent, True, False, True)
                Return True
            Catch primaryEx As Exception
                Try
                    dims.AddDiameter(textPoint, intent)
                    Return True
                Catch fallbackEx As Exception
                    diagnostics.Add("AddDiameter failed: " & primaryEx.Message & " | fallback: " & fallbackEx.Message)
                    Return False
                End Try
            End Try
        End Function

        Private Function TryAddRadiusDimension(
            ByVal dims As GeneralDimensions,
            ByVal textPoint As Point2d,
            ByVal intent As GeometryIntent,
            ByVal diagnostics As List(Of String)) As Boolean

            Try
                dims.AddRadius(textPoint, intent, True, False, False)
                Return True
            Catch primaryEx As Exception
                Try
                    dims.AddRadius(textPoint, intent)
                    Return True
                Catch fallbackEx As Exception
                    diagnostics.Add("AddRadius failed: " & primaryEx.Message & " | fallback: " & fallbackEx.Message)
                    Return False
                End Try
            End Try
        End Function

        Private Function IsPointInsideBounds(
            ByVal point As Point2d,
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal padding As Double) As Boolean

            Return point.X > (minX + padding) AndAlso
                   point.X < (maxX - padding) AndAlso
                   point.Y > (minY + padding) AndAlso
                   point.Y < (maxY - padding)
        End Function

        Private Sub CaptureEndpoint(
            ByVal sheet As Sheet,
            ByVal segment As DrawingCurveSegment,
            ByVal isStart As Boolean,
            ByVal endpoints As List(Of EndpointIntent))

            Dim pt As Point2d = Nothing
            Dim intent As GeometryIntent = Nothing
            Dim parentCurve As DrawingCurve = Nothing

            Try
                parentCurve = segment.Parent

                If isStart Then
                    pt = segment.StartPoint
                Else
                    pt = segment.EndPoint
                End If

                If parentCurve IsNot Nothing AndAlso pt IsNot Nothing Then
                    intent = sheet.CreateGeometryIntent(parentCurve, pt)
                End If

                If intent Is Nothing Then
                    If isStart Then
                        intent = sheet.CreateGeometryIntent(segment, PointIntentEnum.kStartPointIntent)
                    Else
                        intent = sheet.CreateGeometryIntent(segment, PointIntentEnum.kEndPointIntent)
                    End If
                End If
            Catch
                Return
            End Try

            If pt Is Nothing OrElse intent Is Nothing Then
                Return
            End If

            endpoints.Add(New EndpointIntent With {
                .SheetPoint = pt,
                .Intent = intent
            })
        End Sub

        Private Function TryAddOverallDimension(
            ByVal dims As GeneralDimensions,
            ByVal firstCandidates As List(Of EndpointIntent),
            ByVal secondCandidates As List(Of EndpointIntent),
            ByVal isHorizontal As Boolean,
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal baseOffset As Double,
            ByRef failureDetails As String,
            ByVal diagnostics As List(Of String)) As Boolean

            Const minPrimarySpan As Double = 0.0001
            Dim orientationLabel As String = If(isHorizontal, "horizontal", "vertical")
            Dim overallPrimarySpan As Double = If(isHorizontal, Math.Abs(maxX - minX), Math.Abs(maxY - minY))
            If overallPrimarySpan < minPrimarySpan Then
                overallPrimarySpan = minPrimarySpan
            End If

            Dim secondarySpan As Double = If(isHorizontal, Math.Abs(maxY - minY), Math.Abs(maxX - minX))
            If secondarySpan < minPrimarySpan Then
                secondarySpan = minPrimarySpan
            End If

            Dim secondaryCenter As Double = If(isHorizontal, (minY + maxY) / 2.0, (minX + maxX) / 2.0)
            Dim neutralBandHalfWidth As Double = Math.Max(secondarySpan * 0.06, baseOffset * 0.35)
            Dim maxPreferredSecondarySpread As Double = Math.Max(secondarySpan * 0.22, baseOffset * 1.25)

            Dim minAcceptableSpan As Double = overallPrimarySpan * 0.985
            Dim extremityTolerance As Double = Math.Max(overallPrimarySpan * 0.02, 0.05)
            Dim baseGap As Double = Math.Max(Math.Max(baseOffset, overallPrimarySpan * 0.08), 0.75)

            Dim preferredPairs As New List(Of OverallPairCandidate)()
            Dim fallbackPairs As New List(Of OverallPairCandidate)()

            For Each firstCandidate As EndpointIntent In firstCandidates
                For Each secondCandidate As EndpointIntent In secondCandidates
                    If firstCandidate Is Nothing OrElse secondCandidate Is Nothing Then
                        Continue For
                    End If

                    If firstCandidate.Intent Is Nothing OrElse secondCandidate.Intent Is Nothing Then
                        Continue For
                    End If

                    If firstCandidate.Intent Is secondCandidate.Intent Then
                        Continue For
                    End If

                    Dim primarySpan As Double
                    Dim firstDistanceToExtreme As Double
                    Dim secondDistanceToExtreme As Double
                    Dim firstSecondary As Double
                    Dim secondSecondary As Double

                    If isHorizontal Then
                        primarySpan = Math.Abs(secondCandidate.SheetPoint.X - firstCandidate.SheetPoint.X)
                        firstDistanceToExtreme = Math.Abs(firstCandidate.SheetPoint.X - minX)
                        secondDistanceToExtreme = Math.Abs(secondCandidate.SheetPoint.X - maxX)
                        firstSecondary = firstCandidate.SheetPoint.Y
                        secondSecondary = secondCandidate.SheetPoint.Y
                    Else
                        primarySpan = Math.Abs(secondCandidate.SheetPoint.Y - firstCandidate.SheetPoint.Y)
                        firstDistanceToExtreme = Math.Abs(firstCandidate.SheetPoint.Y - minY)
                        secondDistanceToExtreme = Math.Abs(secondCandidate.SheetPoint.Y - maxY)
                        firstSecondary = firstCandidate.SheetPoint.X
                        secondSecondary = secondCandidate.SheetPoint.X
                    End If

                    If primarySpan < minPrimarySpan Then
                        Continue For
                    End If

                    Dim secondarySpread As Double
                    If isHorizontal Then
                        secondarySpread = Math.Abs(secondCandidate.SheetPoint.Y - firstCandidate.SheetPoint.Y)
                    Else
                        secondarySpread = Math.Abs(secondCandidate.SheetPoint.X - firstCandidate.SheetPoint.X)
                    End If

                    Dim firstBand As Integer = ClassifySideBand(firstSecondary, secondaryCenter, neutralBandHalfWidth)
                    Dim secondBand As Integer = ClassifySideBand(secondSecondary, secondaryCenter, neutralBandHalfWidth)

                    If firstBand <> 0 AndAlso secondBand <> 0 AndAlso firstBand <> secondBand Then
                        Continue For
                    End If

                    Dim averageSecondary As Double = (firstSecondary + secondSecondary) / 2.0
                    Dim resolvedBand As Integer = ResolveSideBand(firstBand, secondBand, averageSecondary, secondaryCenter, neutralBandHalfWidth)
                    Dim sideDistance As Double = Math.Abs(averageSecondary - secondaryCenter)

                    Dim extremeScore As Double = firstDistanceToExtreme + secondDistanceToExtreme
                    Dim pairData As New OverallPairCandidate With {
                        .FirstEndpoint = firstCandidate,
                        .SecondEndpoint = secondCandidate,
                        .PrimarySpan = primarySpan,
                        .SecondarySpread = secondarySpread,
                        .ExtremeScore = extremeScore,
                        .SideBand = resolvedBand,
                        .SideDistance = sideDistance
                    }

                    fallbackPairs.Add(pairData)

                    Dim firstNearExtreme As Boolean = (firstDistanceToExtreme <= extremityTolerance)
                    Dim secondNearExtreme As Boolean = (secondDistanceToExtreme <= extremityTolerance)
                    If firstNearExtreme AndAlso secondNearExtreme AndAlso
                       primarySpan >= minAcceptableSpan AndAlso
                       secondarySpread <= maxPreferredSecondarySpread Then
                        preferredPairs.Add(pairData)
                    End If
                Next
            Next

            If preferredPairs.Count = 0 Then
                preferredPairs = fallbackPairs.Where(Function(pairData) pairData.PrimarySpan >= minAcceptableSpan).ToList()
            End If

            If preferredPairs.Count = 0 Then
                preferredPairs = fallbackPairs
            End If

            preferredPairs = preferredPairs.
                OrderByDescending(Function(pairData) pairData.PrimarySpan).
                ThenByDescending(Function(pairData) pairData.SideDistance).
                ThenBy(Function(pairData) pairData.SecondarySpread).
                ThenBy(Function(pairData) pairData.ExtremeScore).
                ToList()

            Dim lastError As String = "No valid intent pair was found."
            Dim gapFactors As Double() = {1.0, 1.25, 1.55, 1.9, 2.35}
            Dim laneFactors As Double() = {0.0, 0.35, 0.7}

            For Each pairData As OverallPairCandidate In preferredPairs
                Dim firstCandidate As EndpointIntent = pairData.FirstEndpoint
                Dim secondCandidate As EndpointIntent = pairData.SecondEndpoint
                Dim primarySpan As Double = pairData.PrimarySpan
                Dim placePositiveSide As Boolean = (pairData.SideBand >= 0)

                For Each gapFactor As Double In gapFactors
                    For Each laneFactor As Double In laneFactors
                        Dim gap As Double = (baseGap * gapFactor) + (baseOffset * laneFactor)
                        Dim textPoint As Point2d = BuildOverallTextPoint(
                            firstCandidate,
                            secondCandidate,
                            isHorizontal,
                            minX,
                            maxX,
                            minY,
                            maxY,
                            gap,
                            placePositiveSide)

                        Dim successfulMode As String = String.Empty
                        Dim attemptError As String = String.Empty

                        If TryAddLinearDimension(dims, textPoint, firstCandidate.Intent, secondCandidate.Intent, isHorizontal, successfulMode, attemptError) Then
                            Dim sideLabel As String = If(isHorizontal,
                                                         If(placePositiveSide, "top", "bottom"),
                                                         If(placePositiveSide, "right", "left"))

                            diagnostics.Add(orientationLabel & " pair succeeded with mode=" & successfulMode &
                                            ", side=" & sideLabel &
                                            ", span=" & Math.Round(primarySpan, 4).ToString() &
                                            ", gap=" & Math.Round(gap, 4).ToString() & " @ (" &
                                            Math.Round(firstCandidate.SheetPoint.X, 4).ToString() & "," & Math.Round(firstCandidate.SheetPoint.Y, 4).ToString() & ") -> (" &
                                            Math.Round(secondCandidate.SheetPoint.X, 4).ToString() & "," & Math.Round(secondCandidate.SheetPoint.Y, 4).ToString() & ")")
                            Return True
                        End If

                        lastError = orientationLabel & " pair " &
                            "(" & Math.Round(firstCandidate.SheetPoint.X, 4).ToString() & ", " & Math.Round(firstCandidate.SheetPoint.Y, 4).ToString() & ") -> " &
                            "(" & Math.Round(secondCandidate.SheetPoint.X, 4).ToString() & ", " & Math.Round(secondCandidate.SheetPoint.Y, 4).ToString() & ") " &
                            "failed: " & attemptError
                        diagnostics.Add(lastError)
                    Next
                Next
            Next

            failureDetails = lastError
            Return False
        End Function

        Private Function BuildOverallTextPoint(
            ByVal firstCandidate As EndpointIntent,
            ByVal secondCandidate As EndpointIntent,
            ByVal isHorizontal As Boolean,
            ByVal minX As Double,
            ByVal maxX As Double,
            ByVal minY As Double,
            ByVal maxY As Double,
            ByVal gap As Double,
            ByVal placePositiveSide As Boolean) As Point2d

            Dim tg As TransientGeometry = m_InventorApp.TransientGeometry

            If isHorizontal Then
                Dim anchorX As Double = (firstCandidate.SheetPoint.X + secondCandidate.SheetPoint.X) / 2.0
                Dim targetY As Double

                If placePositiveSide Then
                    Dim highestPairY As Double = Math.Max(firstCandidate.SheetPoint.Y, secondCandidate.SheetPoint.Y)
                    targetY = Math.Max(maxY, highestPairY) + gap
                Else
                    Dim lowestPairY As Double = Math.Min(firstCandidate.SheetPoint.Y, secondCandidate.SheetPoint.Y)
                    targetY = Math.Min(minY, lowestPairY) - gap
                End If

                Return tg.CreatePoint2d(anchorX, targetY)
            End If

            Dim anchorY As Double = (firstCandidate.SheetPoint.Y + secondCandidate.SheetPoint.Y) / 2.0
            Dim targetX As Double

            If placePositiveSide Then
                Dim rightmostPairX As Double = Math.Max(firstCandidate.SheetPoint.X, secondCandidate.SheetPoint.X)
                targetX = Math.Max(maxX, rightmostPairX) + gap
            Else
                Dim leftmostPairX As Double = Math.Min(firstCandidate.SheetPoint.X, secondCandidate.SheetPoint.X)
                targetX = Math.Min(minX, leftmostPairX) - gap
            End If

            Return tg.CreatePoint2d(targetX, anchorY)
        End Function

        Private Function ClassifySideBand(
            ByVal value As Double,
            ByVal centerValue As Double,
            ByVal neutralHalfWidth As Double) As Integer

            If value > centerValue + neutralHalfWidth Then
                Return 1
            End If

            If value < centerValue - neutralHalfWidth Then
                Return -1
            End If

            Return 0
        End Function

        Private Function ResolveSideBand(
            ByVal firstBand As Integer,
            ByVal secondBand As Integer,
            ByVal averageValue As Double,
            ByVal centerValue As Double,
            ByVal neutralHalfWidth As Double) As Integer

            If firstBand = secondBand AndAlso firstBand <> 0 Then
                Return firstBand
            End If

            If firstBand <> 0 AndAlso secondBand = 0 Then
                Return firstBand
            End If

            If secondBand <> 0 AndAlso firstBand = 0 Then
                Return secondBand
            End If

            Dim averageBand As Integer = ClassifySideBand(averageValue, centerValue, neutralHalfWidth * 0.4)
            If averageBand <> 0 Then
                Return averageBand
            End If

            Return 1
        End Function

        Private Function TryAddLinearDimension(
            ByVal dims As GeneralDimensions,
            ByVal textPoint As Point2d,
            ByVal firstIntent As GeometryIntent,
            ByVal secondIntent As GeometryIntent,
            ByVal isHorizontal As Boolean,
            ByRef successfulMode As String,
            ByRef errorDetails As String) As Boolean

            Dim requestedType As DimensionTypeEnum = If(isHorizontal, DimensionTypeEnum.kHorizontalDimensionType, DimensionTypeEnum.kVerticalDimensionType)

            Try
                dims.AddLinear(textPoint, firstIntent, secondIntent, requestedType)
                successfulMode = requestedType.ToString()
                Return True
            Catch requestedEx As Exception
                Try
                    dims.AddLinear(textPoint, firstIntent, secondIntent, DimensionTypeEnum.kAlignedDimensionType)
                    successfulMode = DimensionTypeEnum.kAlignedDimensionType.ToString()
                    Return True
                Catch alignedEx As Exception
                    Try
                        dims.AddLinear(textPoint, firstIntent, secondIntent)
                        successfulMode = "Default"
                        Return True
                    Catch defaultEx As Exception
                        errorDetails = requestedType.ToString() & "=" & requestedEx.Message & "; " &
                            "kAlignedDimensionType=" & alignedEx.Message & "; " &
                            "Default=" & defaultEx.Message
                        Return False
                    End Try
                End Try
            End Try
        End Function

        Private Function WriteFailureLog(
            ByVal drawingDoc As DrawingDocument,
            ByVal sheet As Sheet,
            ByVal ex As Exception,
            ByVal diagnostics As List(Of String)) As String

            Try
                Dim baseFolder As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Spectiv", "InventorAutomationSuite", "Logs")
                System.IO.Directory.CreateDirectory(baseFolder)

                Dim fileName As String = "AutoDetailer_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log"
                Dim fullPath As String = System.IO.Path.Combine(baseFolder, fileName)

                Dim lines As New List(Of String)()
                lines.Add("Timestamp: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                lines.Add("Document: " & If(drawingDoc Is Nothing, "(none)", drawingDoc.DisplayName))
                lines.Add("Sheet: " & If(sheet Is Nothing, "(none)", sheet.Name))
                lines.Add("Error: " & ex.Message)
                lines.Add("--- Diagnostics ---")

                If diagnostics Is Nothing OrElse diagnostics.Count = 0 Then
                    lines.Add("(no additional diagnostics)")
                Else
                    Dim tailCount As Integer = Math.Min(40, diagnostics.Count)
                    For Each item As String In diagnostics.Skip(Math.Max(0, diagnostics.Count - tailCount))
                        lines.Add(item)
                    Next
                End If

                System.IO.File.WriteAllLines(fullPath, lines)
                Return fullPath
            Catch
                Return String.Empty
            End Try
        End Function

        Private Function WriteRunLog(
            ByVal drawingDoc As DrawingDocument,
            ByVal sheet As Sheet,
            ByVal summary As FeatureDimensionSummary,
            ByVal diagnostics As List(Of String)) As String

            Try
                Dim baseFolder As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Spectiv", "InventorAutomationSuite", "Logs")
                System.IO.Directory.CreateDirectory(baseFolder)

                Dim fileName As String = "AutoDetailer_Run_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log"
                Dim fullPath As String = System.IO.Path.Combine(baseFolder, fileName)

                Dim lines As New List(Of String)()
                lines.Add("Timestamp: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                lines.Add("Document: " & If(drawingDoc Is Nothing, "(none)", drawingDoc.DisplayName))
                lines.Add("Sheet: " & If(sheet Is Nothing, "(none)", sheet.Name))
                lines.Add("Result: Completed with zero hole/cutout dimensions")

                If summary IsNot Nothing Then
                    lines.Add("TotalCurves: " & summary.TotalCurves.ToString())
                    lines.Add("CircularCurves: " & summary.CircularCurves.ToString())
                    lines.Add("InteriorCandidates: " & summary.InteriorCandidates.ToString())
                    lines.Add("DimensionsAdded: " & summary.DimensionsAdded.ToString())
                End If

                lines.Add("--- Diagnostics ---")
                If diagnostics Is Nothing OrElse diagnostics.Count = 0 Then
                    lines.Add("(no additional diagnostics)")
                Else
                    For Each item As String In diagnostics
                        lines.Add(item)
                    Next
                End If

                System.IO.File.WriteAllLines(fullPath, lines)
                Return fullPath
            Catch
                Return String.Empty
            End Try
        End Function
    End Class

End Namespace
