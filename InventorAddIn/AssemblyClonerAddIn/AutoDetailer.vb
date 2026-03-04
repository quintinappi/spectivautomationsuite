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
        End Class

        Private Class FeatureDimensionSummary
            Public Property TotalCurves As Integer
            Public Property CircularCurves As Integer
            Public Property InteriorCandidates As Integer
            Public Property DimensionsAdded As Integer
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

            Dim horizontalText As Point2d = tg.CreatePoint2d((minX + maxX) / 2.0, maxY + baseOffset)
            Dim verticalText As Point2d = tg.CreatePoint2d(maxX + baseOffset, (minY + maxY) / 2.0)

            Dim minXCandidates As List(Of EndpointIntent) = allEndpoints.OrderBy(Function(pointEntry) pointEntry.SheetPoint.X).Take(16).ToList()
            Dim maxXCandidates As List(Of EndpointIntent) = allEndpoints.OrderByDescending(Function(pointEntry) pointEntry.SheetPoint.X).Take(16).ToList()
            Dim minYCandidates As List(Of EndpointIntent) = allEndpoints.OrderBy(Function(pointEntry) pointEntry.SheetPoint.Y).Take(16).ToList()
            Dim maxYCandidates As List(Of EndpointIntent) = allEndpoints.OrderByDescending(Function(pointEntry) pointEntry.SheetPoint.Y).Take(16).ToList()

            Dim tx As Transaction = Nothing
            Dim dimensionDiagnostics As New List(Of String)()
            Dim featureDimensionsAdded As Integer = 0
            Dim featureSummary As FeatureDimensionSummary = Nothing
            Try
                tx = m_InventorApp.TransactionManager.StartTransaction(drawingDoc, "Auto Detail IDW")

                Dim horizontalFailureDetails As String = String.Empty
                Dim verticalFailureDetails As String = String.Empty

                Dim horizontalPlaced As Boolean = TryAddOverallDimension(dims, horizontalText, minXCandidates, maxXCandidates, True, horizontalFailureDetails, dimensionDiagnostics)
                Dim verticalPlaced As Boolean = TryAddOverallDimension(dims, verticalText, minYCandidates, maxYCandidates, False, verticalFailureDetails, dimensionDiagnostics)

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

            Dim completionMessage As String = "Auto detailing complete: added overall dimensions and " & featureDimensionsAdded.ToString() & " feature dimensions for the selected view."

            If featureDimensionsAdded = 0 Then
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
            Dim summary As New FeatureDimensionSummary With {
                .TotalCurves = 0,
                .CircularCurves = 0,
                .InteriorCandidates = 0,
                .DimensionsAdded = 0
            }
            Dim viewCenterX As Double = (minX + maxX) / 2.0
            Dim viewCenterY As Double = (minY + maxY) / 2.0
            Dim innerPadding As Double = Math.Max(baseOffset * 0.25, 0.2)
            Dim textOffset As Double = Math.Max(baseOffset * 0.75, 0.35)
            Dim holeCenters As New List(Of HoleCenterIntent)()
            Dim usedDimKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

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

                If Not IsPointInsideBounds(center, minX, maxX, minY, maxY, innerPadding) Then
                    Continue For
                End If

                summary.InteriorCandidates += 1

                Dim intent As GeometryIntent = Nothing
                Try
                    intent = sheet.CreateGeometryIntent(curve, center)
                Catch ex As Exception
                    Try
                        intent = sheet.CreateGeometryIntent(curve)
                    Catch
                        intent = Nothing
                    End Try
                End Try

                If intent Is Nothing Then
                    Continue For
                End If

                holeCenters.Add(New HoleCenterIntent With {
                    .SheetPoint = center,
                    .Intent = intent
                })
            Next

            Dim leftRef As EndpointIntent = If(minXCandidates IsNot Nothing AndAlso minXCandidates.Count > 0, minXCandidates(0), Nothing)
            Dim rightRef As EndpointIntent = If(maxXCandidates IsNot Nothing AndAlso maxXCandidates.Count > 0, maxXCandidates(0), Nothing)
            Dim bottomRef As EndpointIntent = If(minYCandidates IsNot Nothing AndAlso minYCandidates.Count > 0, minYCandidates(0), Nothing)
            Dim topRef As EndpointIntent = If(maxYCandidates IsNot Nothing AndAlso maxYCandidates.Count > 0, maxYCandidates(0), Nothing)

            For Each hole As HoleCenterIntent In holeCenters
                If hole Is Nothing OrElse hole.Intent Is Nothing OrElse hole.SheetPoint Is Nothing Then
                    Continue For
                End If

                Dim useLeft As Boolean = (Math.Abs(hole.SheetPoint.X - minX) <= Math.Abs(maxX - hole.SheetPoint.X))
                Dim xRef As EndpointIntent = If(useLeft, leftRef, rightRef)
                If xRef IsNot Nothing AndAlso xRef.Intent IsNot Nothing Then
                    Dim xTextY As Double = If(hole.SheetPoint.Y >= viewCenterY, hole.SheetPoint.Y + textOffset, hole.SheetPoint.Y - textOffset)
                    Dim xTextPoint As Point2d = tg.CreatePoint2d((xRef.SheetPoint.X + hole.SheetPoint.X) / 2.0, xTextY)
                    If TryAddLinearDimensionUnique(dims, xTextPoint, xRef.Intent, hole.Intent, True, usedDimKeys, diagnostics) Then
                        summary.DimensionsAdded += 1
                    End If
                End If

                Dim useBottom As Boolean = (Math.Abs(hole.SheetPoint.Y - minY) <= Math.Abs(maxY - hole.SheetPoint.Y))
                Dim yRef As EndpointIntent = If(useBottom, bottomRef, topRef)
                If yRef IsNot Nothing AndAlso yRef.Intent IsNot Nothing Then
                    Dim yTextX As Double = If(hole.SheetPoint.X >= viewCenterX, hole.SheetPoint.X + textOffset, hole.SheetPoint.X - textOffset)
                    Dim yTextPoint As Point2d = tg.CreatePoint2d(yTextX, (yRef.SheetPoint.Y + hole.SheetPoint.Y) / 2.0)
                    If TryAddLinearDimensionUnique(dims, yTextPoint, yRef.Intent, hole.Intent, False, usedDimKeys, diagnostics) Then
                        summary.DimensionsAdded += 1
                    End If
                End If
            Next

            If holeCenters.Count > 1 Then
                Dim sortedByX As List(Of HoleCenterIntent) = holeCenters.OrderBy(Function(h) h.SheetPoint.X).ToList()
                Dim sortedByY As List(Of HoleCenterIntent) = holeCenters.OrderBy(Function(h) h.SheetPoint.Y).ToList()

                For i As Integer = 1 To sortedByX.Count - 1
                    Dim h1 As HoleCenterIntent = sortedByX(i - 1)
                    Dim h2 As HoleCenterIntent = sortedByX(i)
                    Dim textPoint As Point2d = tg.CreatePoint2d((h1.SheetPoint.X + h2.SheetPoint.X) / 2.0, Math.Max(h1.SheetPoint.Y, h2.SheetPoint.Y) + textOffset)
                    If TryAddLinearDimensionUnique(dims, textPoint, h1.Intent, h2.Intent, True, usedDimKeys, diagnostics) Then
                        summary.DimensionsAdded += 1
                    End If
                Next

                For i As Integer = 1 To sortedByY.Count - 1
                    Dim h1 As HoleCenterIntent = sortedByY(i - 1)
                    Dim h2 As HoleCenterIntent = sortedByY(i)
                    Dim textPoint As Point2d = tg.CreatePoint2d(Math.Max(h1.SheetPoint.X, h2.SheetPoint.X) + textOffset, (h1.SheetPoint.Y + h2.SheetPoint.Y) / 2.0)
                    If TryAddLinearDimensionUnique(dims, textPoint, h1.Intent, h2.Intent, False, usedDimKeys, diagnostics) Then
                        summary.DimensionsAdded += 1
                    End If
                Next
            End If

            diagnostics.Add("Feature summary: totalCurves=" & summary.TotalCurves.ToString() &
                            ", circularCurves=" & summary.CircularCurves.ToString() &
                            ", interiorCandidates=" & summary.InteriorCandidates.ToString() &
                            ", dimensionsAdded=" & summary.DimensionsAdded.ToString())

            Return summary
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
            ByVal textPoint As Point2d,
            ByVal firstCandidates As List(Of EndpointIntent),
            ByVal secondCandidates As List(Of EndpointIntent),
            ByVal isHorizontal As Boolean,
            ByRef failureDetails As String,
            ByVal diagnostics As List(Of String)) As Boolean

            Const minPrimarySpan As Double = 0.0001
            Dim orientationLabel As String = If(isHorizontal, "horizontal", "vertical")
            Dim lastError As String = "No valid intent pair was found."

            For Each firstCandidate As EndpointIntent In firstCandidates
                For Each secondCandidate As EndpointIntent In secondCandidates
                    If firstCandidate Is Nothing OrElse secondCandidate Is Nothing Then
                        Continue For
                    End If

                    If firstCandidate.Intent Is secondCandidate.Intent Then
                        Continue For
                    End If

                    Dim primarySpan As Double
                    If isHorizontal Then
                        primarySpan = Math.Abs(secondCandidate.SheetPoint.X - firstCandidate.SheetPoint.X)
                    Else
                        primarySpan = Math.Abs(secondCandidate.SheetPoint.Y - firstCandidate.SheetPoint.Y)
                    End If

                    If primarySpan < minPrimarySpan Then
                        Continue For
                    End If

                    Dim successfulMode As String = String.Empty
                    Dim attemptError As String = String.Empty

                    If TryAddLinearDimension(dims, textPoint, firstCandidate.Intent, secondCandidate.Intent, isHorizontal, successfulMode, attemptError) Then
                        diagnostics.Add(orientationLabel & " pair succeeded with mode=" & successfulMode & " @ (" &
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

            failureDetails = lastError
            Return False
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
