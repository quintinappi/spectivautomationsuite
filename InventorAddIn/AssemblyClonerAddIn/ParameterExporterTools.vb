' ==============================================================================
' PARAMETER EXPORTER TOOLS - VB.NET Add-In Implementation
' ==============================================================================
' Migration of VBS parameter exporter scripts to Inventor Add-In
' Handles Length, Length2, and Thickness parameter export enabling
' Author: Quintin de Bruin © 2026
' ==============================================================================

Imports Inventor
Imports System.Collections.Generic
Imports System.IO

Namespace AssemblyClonerAddIn

    Public Class ParameterExporterTools

        Private ReadOnly m_InventorApp As Inventor.Application
        Private Const kZeroDecimalPlaceLinearPrecision As Integer = 0
        Private Const kDimensionDisplayAsPreciseValue As Integer = 34817
        Private Const kDimensionDisplayAsValue As Integer = 34821
        Private Const kBomRefreshDummyParameterName As String = "BOM_REFRESH_TEMP"
        Private Const kBomRefreshDummyPropertyName As String = "_PRECISION_UPDATE_"

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub ExecuteLengthParameterExporter()
            ExecuteParameterExporterCore("Length Parameter Exporter", "Length", scanPlateParts:=False)
        End Sub

        Public Sub ExecuteLength2ParameterExporter()
            ExecuteParameterExporterCore("Length2 Parameter Exporter", "Length2", scanPlateParts:=False)
        End Sub

        Public Sub ExecuteThicknessParameterExporter()
            ExecuteParameterExporterCore("Thickness Parameter Exporter", "Thickness", scanPlateParts:=True)
        End Sub

        Public Sub ExecuteUpdateDocumentSettings()
            Try
                If Not ValidateActiveAssemblyDocument() Then
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                Dim plateTargets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                CollectPlatePartsByDescription(asmDoc.ComponentDefinition.Occurrences, plateTargets)

                AddInDiagnostics.Log("ParameterExporterTools", "UpdateDocumentSettings START | Assembly='" & SafeDocumentPath(asmDoc) & "' | Targets=" & plateTargets.Count.ToString())

                If plateTargets.Count = 0 Then
                    MsgBox("No plate parts found in the active assembly.", MsgBoxStyle.Information, "Update Document Settings")
                    Return
                End If

                Dim modePrompt As String = "Found " & plateTargets.Count.ToString() & " plate parts." & vbCrLf & vbCrLf &
                                           "Choose update mode:" & vbCrLf &
                                           "YES = Fast API mode" & vbCrLf &
                                           "NO = Robust UI Apply mode (recommended for stale BOM cache)" & vbCrLf &
                                           "CANCEL = Exit" & vbCrLf & vbCrLf &
                                           "UI mode opens Document Settings per part and triggers the same Apply/OK event path as manual toggling."

                Dim modeChoice As MsgBoxResult = MsgBox(modePrompt,
                                                        MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel Or MsgBoxStyle.DefaultButton2,
                                                        "Update Document Settings")

                If modeChoice = MsgBoxResult.Cancel Then
                    Return
                End If

                If modeChoice = MsgBoxResult.No Then
                    AddInDiagnostics.Log("ParameterExporterTools", "UpdateDocumentSettings user selected UI mode | Assembly='" & SafeDocumentPath(asmDoc) & "' | Targets=" & plateTargets.Count.ToString())
                    Dim uiModeTool As New AssemblyBOMPrecision(m_InventorApp)
                    uiModeTool.Execute(plateTargets.Keys)
                    Return
                End If

                Dim prompt As String = "Found " & plateTargets.Count.ToString() & " plate parts." & vbCrLf & vbCrLf &
                                       "This tool will open each plate and force the precision refresh cycle:" & vbCrLf &
                                       "1) Linear Dim Precision toggle -> 0" & vbCrLf &
                                       "2) Dimension Display toggle -> Display as value" & vbCrLf &
                                       "3) Display Parameter As Expression toggle" & vbCrLf &
                                       "4) Units toggle and restore (dirty/refresh trigger)" & vbCrLf &
                                       "5) Save part" & vbCrLf & vbCrLf &
                                       "Continue?"

                If MsgBox(prompt, MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Update Document Settings") <> MsgBoxResult.Yes Then
                    Return
                End If

                Dim updatedCount As Integer = 0
                Dim skippedCount As Integer = 0
                Dim failedCount As Integer = 0
                Dim dirtyFlagFailureCount As Integer = 0
                Dim noSettingDeltaCount As Integer = 0

                Using progress As New ToolProgressForm("Update Document Settings")
                    progress.Show()

                    Dim index As Integer = 0
                    For Each kvp As KeyValuePair(Of String, String) In plateTargets
                        index += 1
                        Dim pct As Integer = CInt((CDbl(index) / CDbl(plateTargets.Count)) * 100.0)
                        progress.UpdateProgress(Math.Max(5, Math.Min(95, pct)), "Updating " & kvp.Value & "...")

                        Dim status As String = String.Empty
                        Dim dirtyFlagSucceeded As Boolean = False
                        Dim hadAnySettingDelta As Boolean = False

                        If UpdatePlateDocumentSettingsWithDirtyRefresh(kvp.Key,
                                                                      status,
                                                                      dirtyFlagSucceeded,
                                                                      hadAnySettingDelta) Then
                            updatedCount += 1

                            If Not dirtyFlagSucceeded Then
                                dirtyFlagFailureCount += 1
                            End If

                            If Not hadAnySettingDelta Then
                                noSettingDeltaCount += 1
                            End If
                        ElseIf String.Equals(status, "SKIPPED", StringComparison.OrdinalIgnoreCase) Then
                            skippedCount += 1
                        Else
                            failedCount += 1
                        End If

                        Dim normalizedStatus As String = If(String.IsNullOrWhiteSpace(status), "FAILED", status)
                        AddInDiagnostics.Log("ParameterExporterTools", "UpdateDocumentSettings item | Part='" & kvp.Key & "' | Status='" & normalizedStatus & "'")
                    Next

                    progress.UpdateProgress(100, "Refreshing assembly BOM...")
                End Using

                RefreshAssemblyBomPrecision(asmDoc)

                Dim likelyDisplayCacheStillStale As Boolean = updatedCount > 0 AndAlso
                                                         (dirtyFlagFailureCount > 0 OrElse noSettingDeltaCount = updatedCount)

                If likelyDisplayCacheStillStale Then
                    Dim fallbackPrompt As String = "Inventor still reports stale-cache conditions for this run:" & vbCrLf & vbCrLf &
                                                   "- Dirty flag reinforcement failed on " & dirtyFlagFailureCount.ToString() & " part(s)" & vbCrLf &
                                                   "- Settings already at target on " & noSettingDeltaCount.ToString() & " part(s)" & vbCrLf & vbCrLf &
                                                   "Run ROBUST UI fallback now?" & vbCrLf &
                                                   "This uses keyboard automation (legacy SendKeys style)." & vbCrLf &
                                                   "Do not touch keyboard/mouse while it runs."

                    If MsgBox(fallbackPrompt, MsgBoxStyle.Question Or MsgBoxStyle.YesNo, "Update Document Settings") = MsgBoxResult.Yes Then
                        AddInDiagnostics.Log("ParameterExporterTools", "UpdateDocumentSettings launching UI fallback | Assembly='" & SafeDocumentPath(asmDoc) & "' | DirtyFlagFailures=" & dirtyFlagFailureCount.ToString() & " | NoDelta=" & noSettingDeltaCount.ToString())
                        Dim uiFallbackTool As New AssemblyBOMPrecision(m_InventorApp)
                        uiFallbackTool.Execute(plateTargets.Keys)
                        Return
                    End If
                End If

                AddInDiagnostics.Log("ParameterExporterTools", "UpdateDocumentSettings COMPLETE | Assembly='" & SafeDocumentPath(asmDoc) & "' | Targets=" & plateTargets.Count.ToString() & " | Updated=" & updatedCount.ToString() & " | Skipped=" & skippedCount.ToString() & " | Failed=" & failedCount.ToString() & " | DirtyFlagFailures=" & dirtyFlagFailureCount.ToString() & " | NoDelta=" & noSettingDeltaCount.ToString())

                MsgBox("Update Document Settings complete." & vbCrLf & vbCrLf &
                       "Plate parts found: " & plateTargets.Count.ToString() & vbCrLf &
                       "Updated: " & updatedCount.ToString() & vbCrLf &
                       "Skipped: " & skippedCount.ToString() & vbCrLf &
                       "Failed: " & failedCount.ToString() & vbCrLf &
                       "Dirty-flag fallback needed: " & dirtyFlagFailureCount.ToString(),
                       MsgBoxStyle.Information,
                       "Update Document Settings")

            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "ExecuteUpdateDocumentSettings failed | " & ex.Message)
                Throw
            End Try
        End Sub

        Public Sub ExecuteFixNonPlateParts()
            Try
                If Not ValidateActiveAssemblyDocument() Then
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                Dim nonPlateTargets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                CollectNonPlatePartsMissingLength(asmDoc.ComponentDefinition.Occurrences, nonPlateTargets)

                If nonPlateTargets.Count = 0 Then
                    MsgBox("No non-plate parts found that need Length2.", MsgBoxStyle.Information, "Fix Non-Plate Parts")
                    Return
                End If

                Dim successCount As Integer = 0
                Dim failCount As Integer = 0
                Dim skippedCount As Integer = 0

                Using progress As New ToolProgressForm("Fix Non-Plate Parts")
                    progress.Show()

                    Dim index As Integer = 0
                    For Each kvp As KeyValuePair(Of String, String) In nonPlateTargets
                        index += 1
                        Dim pct As Integer = CInt((CDbl(index) / CDbl(nonPlateTargets.Count)) * 100.0)
                        progress.UpdateProgress(Math.Max(5, Math.Min(95, pct)), "Fixing " & kvp.Value & "...")

                        Dim detail As String = String.Empty
                        Dim fixedOk As Boolean = FixLength2ForPartPath(kvp.Key, detail)
                        If fixedOk Then
                            successCount += 1
                        Else
                            If detail = "SKIPPED" Then
                                skippedCount += 1
                            Else
                                failCount += 1
                            End If
                        End If
                    Next

                    progress.UpdateProgress(100, "Complete")
                    progress.CompleteSuccess("Fix Non-Plate Parts complete.")
                End Using

                MsgBox("Fix Non-Plate Parts complete." & vbCrLf & vbCrLf &
                       "Parts targeted: " & nonPlateTargets.Count.ToString() & vbCrLf &
                       "Fixed: " & successCount.ToString() & vbCrLf &
                       "Skipped: " & skippedCount.ToString() & vbCrLf &
                       "Failed: " & failCount.ToString(),
                       MsgBoxStyle.Information,
                       "Fix Non-Plate Parts")

            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "ExecuteFixNonPlateParts failed | " & ex.Message)
                Throw
            End Try
        End Sub

        Public Sub ExecuteFixSinglePartLength2()
            If m_InventorApp Is Nothing OrElse m_InventorApp.ActiveDocument Is Nothing Then
                Throw New InvalidOperationException("No active document. Open a part first.")
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                Throw New InvalidOperationException("Fix Single Part Length2 requires an active part (.ipt).")
            End If

            Dim partDoc As PartDocument = CType(m_InventorApp.ActiveDocument, PartDocument)
            Dim detail As String = String.Empty
            Dim ok As Boolean = EnsureLength2FromLongestModelDimension(partDoc, detail)

            If ok Then
                Try
                    partDoc.Update2(True)
                Catch
                End Try

                Try
                    partDoc.Save2(True)
                Catch
                End Try

                MsgBox("Length2 created/updated successfully on active part.", MsgBoxStyle.Information, "Fix Single Part Length2")
            Else
                Throw New InvalidOperationException("Could not create/update Length2. " & detail)
            End If
        End Sub

        Public Sub ExecuteFixBOMPlateDimensions()
            Try
                If Not ValidateActiveAssemblyDocument() Then
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                Dim plateTargets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                CollectPlatePartsByDescription(asmDoc.ComponentDefinition.Occurrences, plateTargets)

                If plateTargets.Count = 0 Then
                    MsgBox("No plate parts found in the active assembly.", MsgBoxStyle.Information, "Fix BOM Plate Dimensions")
                    Return
                End If

                Dim updated As Integer = 0
                Dim skipped As Integer = 0
                Dim failed As Integer = 0

                Using progress As New ToolProgressForm("Fix BOM Plate Dimensions")
                    progress.Show()

                    Dim index As Integer = 0
                    For Each kvp As KeyValuePair(Of String, String) In plateTargets
                        index += 1
                        Dim pct As Integer = CInt((CDbl(index) / CDbl(plateTargets.Count)) * 100.0)
                        progress.UpdateProgress(Math.Max(5, Math.Min(95, pct)), "Updating " & kvp.Value & "...")

                        Dim status As String = String.Empty
                        If SetPlateBomProperties(kvp.Key, status) Then
                            updated += 1
                        ElseIf status = "SKIPPED" Then
                            skipped += 1
                        Else
                            failed += 1
                        End If
                    Next

                    progress.UpdateProgress(100, "Refreshing assembly...")
                End Using

                Try
                    asmDoc.Update2(True)
                Catch
                End Try

                MsgBox("Fix BOM Plate Dimensions complete." & vbCrLf & vbCrLf &
                       "Plate parts found: " & plateTargets.Count.ToString() & vbCrLf &
                       "Updated: " & updated.ToString() & vbCrLf &
                       "Skipped (not sheet metal): " & skipped.ToString() & vbCrLf &
                       "Failed: " & failed.ToString(),
                       MsgBoxStyle.Information,
                       "Fix BOM Plate Dimensions")

            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "ExecuteFixBOMPlateDimensions failed | " & ex.Message)
                Throw
            End Try
        End Sub

        Public Sub ExecuteApplyPlateDescStockFormula()
            Dim warningText As String = "ONLY RUN THIS ON LAST STEP IN DETAILING PHASE after all parameter tools and plate conversion is run." & vbCrLf & vbCrLf &
                                        "This will inject formula values into Description and Stock Number." & vbCrLf & vbCrLf &
                                        "Are you sure?"

            Dim confirm As MsgBoxResult = MsgBox(warningText,
                                                 MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2,
                                                 "Final Step Warning")
            If confirm <> MsgBoxResult.Yes Then
                Return
            End If

            If m_InventorApp Is Nothing OrElse m_InventorApp.ActiveDocument Is Nothing Then
                Throw New InvalidOperationException("No active model found. Open an IPT or IAM first.")
            End If

            Dim formulaValue As String = "=PL<thickness> S355JR - <sheet metal width> x <sheet metal length>"
            Dim activeDoc As Document = m_InventorApp.ActiveDocument

            If activeDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ApplyFormulaToPart(CType(activeDoc, PartDocument), formulaValue)
                MsgBox("Formula injected into Description and Stock Number for active part.", MsgBoxStyle.Information, "Apply Plate Desc/Stock Formula")
                Return
            End If

            If activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim asmDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
                Dim targets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                CollectPlatePartsByDescription(asmDoc.ComponentDefinition.Occurrences, targets)

                If targets.Count = 0 Then
                    MsgBox("No plate parts found in active assembly.", MsgBoxStyle.Information, "Apply Plate Desc/Stock Formula")
                    Return
                End If

                Dim updated As Integer = 0
                Dim failed As Integer = 0

                Using progress As New ToolProgressForm("Apply Plate Desc/Stock Formula")
                    progress.Show()

                    Dim index As Integer = 0
                    For Each kvp As KeyValuePair(Of String, String) In targets
                        index += 1
                        Dim pct As Integer = CInt((CDbl(index) / CDbl(targets.Count)) * 100.0)
                        progress.UpdateProgress(Math.Max(5, Math.Min(95, pct)), "Applying to " & kvp.Value & "...")

                        Dim partDoc As PartDocument = Nothing
                        Dim openedByTool As Boolean = False
                        Try
                            For Each doc As Document In m_InventorApp.Documents
                                If String.Equals(doc.FullFileName, kvp.Key, StringComparison.OrdinalIgnoreCase) Then
                                    partDoc = TryCast(doc, PartDocument)
                                    Exit For
                                End If
                            Next

                            If partDoc Is Nothing Then
                                partDoc = TryCast(m_InventorApp.Documents.Open(kvp.Key, False), PartDocument)
                                openedByTool = partDoc IsNot Nothing
                            End If

                            If partDoc Is Nothing Then
                                failed += 1
                            Else
                                ApplyFormulaToPart(partDoc, formulaValue)
                                updated += 1
                            End If
                        Catch ex As Exception
                            failed += 1
                            AddInDiagnostics.Log("ParameterExporterTools", "ExecuteApplyPlateDescStockFormula failed for part | Path='" & kvp.Key & "' | " & ex.Message)
                        Finally
                            If openedByTool AndAlso partDoc IsNot Nothing Then
                                Try
                                    partDoc.Close(True)
                                Catch
                                End Try
                            End If
                        End Try
                    Next

                    progress.UpdateProgress(100, "Complete")
                    progress.CompleteSuccess("Apply Plate Desc/Stock Formula complete.")
                End Using

                MsgBox("Apply Plate Desc/Stock Formula complete." & vbCrLf & vbCrLf &
                       "Plate parts found: " & targets.Count.ToString() & vbCrLf &
                       "Updated: " & updated.ToString() & vbCrLf &
                       "Failed: " & failed.ToString(),
                       MsgBoxStyle.Information,
                       "Apply Plate Desc/Stock Formula")
                Return
            End If

            Throw New InvalidOperationException("Active model must be IPT or IAM.")
        End Sub

        Private Sub ExecuteParameterExporterCore(ByVal toolName As String,
                                                ByVal parameterName As String,
                                                ByVal scanPlateParts As Boolean)
            Try
                AddInDiagnostics.Log("ParameterExporterTools", "=== " & toolName & " STARTED ===")

                If Not ValidateActiveAssemblyDocument() Then
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                AddInDiagnostics.Log("ParameterExporterTools", "Processing assembly: " & asmDoc.FullFileName)

                Using progress As New ToolProgressForm(toolName)
                    progress.Show()
                    progress.UpdateProgress(10, "Scanning assembly components...")

                    Dim partMap As Dictionary(Of String, String)
                    If scanPlateParts Then
                        partMap = ScanAssemblyForPlateParts(asmDoc)
                    Else
                        partMap = ScanAssemblyForNonPlateParts(asmDoc)
                    End If

                    If partMap.Count = 0 Then
                        progress.CompleteSuccess("No matching parts found.")

                        If scanPlateParts Then
                            MsgBox("No plate parts found in assembly BOM." & vbCrLf & vbCrLf &
                                   "No parts matched PL/PLATE/VRN/S355JR indicators.",
                                   MsgBoxStyle.Information,
                                   "No Plate Parts Found")
                        Else
                            MsgBox("No non-plate parts found in assembly BOM." & vbCrLf & vbCrLf &
                                   "All parts appear to be plate-type parts.",
                                   MsgBoxStyle.Information,
                                   "No Non-Plate Parts Found")
                        End If
                        Return
                    End If

                    progress.UpdateProgress(20, "Processing " & partMap.Count.ToString() & " parts...")
                    Dim results As ProcessingResults = ProcessParts(partMap, parameterName, progress, 20, 90)

                    progress.UpdateProgress(100, "Complete")
                    progress.CompleteSuccess(toolName & " complete.")
                    ShowResults(toolName, results)
                End Using

            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "ERROR in " & toolName & ": " & ex.Message)
                MsgBox("Error in " & toolName & ": " & ex.Message, MsgBoxStyle.Critical, toolName)
            End Try
        End Sub

        Private Function ValidateActiveAssemblyDocument() As Boolean
            If m_InventorApp Is Nothing OrElse m_InventorApp.ActiveDocument Is Nothing Then
                AddInDiagnostics.Log("ParameterExporterTools", "ERROR: No active document found")
                MsgBox("No active document. Please open an assembly in Inventor.", MsgBoxStyle.Exclamation, "Assembly Required")
                Return False
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                AddInDiagnostics.Log("ParameterExporterTools", "ERROR: Active document is not an assembly")
                MsgBox("Please open an assembly document (.iam) before running this tool.", MsgBoxStyle.Exclamation, "Assembly Required")
                Return False
            End If

            Return True
        End Function

        Private Function UpdatePlateDocumentSettingsWithDirtyRefresh(ByVal partPath As String,
                                                                      ByRef status As String,
                                                                      ByRef dirtyFlagSucceeded As Boolean,
                                                                      ByRef hadAnySettingDelta As Boolean) As Boolean
            status = String.Empty
            dirtyFlagSucceeded = False
            hadAnySettingDelta = False

            If String.IsNullOrWhiteSpace(partPath) Then
                status = "FAILED"
                Return False
            End If

            Dim partDoc As PartDocument = Nothing
            Dim openedByTool As Boolean = False

            Try
                partDoc = TryCast(m_InventorApp.Documents.Open(partPath, True), PartDocument)
                openedByTool = partDoc IsNot Nothing

                If partDoc Is Nothing Then
                    status = "FAILED"
                    Return False
                End If

                AddInDiagnostics.Log("ParameterExporterTools", "Update settings start | Part='" & partPath & "' | OpenedByTool=" & openedByTool.ToString())

                Try
                    partDoc.Activate()
                Catch
                End Try

                If Not partDoc.IsModifiable Then
                    status = "SKIPPED"
                    AddInDiagnostics.Log("ParameterExporterTools", "Update settings skipped (not modifiable) | Part='" & partPath & "'")
                    Return False
                End If

                Dim params As Parameters = partDoc.ComponentDefinition.Parameters
                If params Is Nothing Then
                    status = "FAILED"
                    Return False
                End If

                Dim units As UnitsOfMeasure = Nothing
                Dim unitsPrecisionTrace As String = "LengthDisplayPrecision unavailable"
                Dim hadUomDisplayDelta As Boolean = False
                Try
                    units = partDoc.UnitsOfMeasure
                    If units IsNot Nothing Then
                        Dim originalLengthDisplayPrecision As Integer = CInt(units.LengthDisplayPrecision)
                        hadUomDisplayDelta = originalLengthDisplayPrecision <> kZeroDecimalPlaceLinearPrecision
                        Dim tempLengthDisplayPrecision As Integer = If(originalLengthDisplayPrecision = 3, 2, 3)
                        units.LengthDisplayPrecision = CType(tempLengthDisplayPrecision, LinearPrecisionEnum)
                        units.LengthDisplayPrecision = CType(kZeroDecimalPlaceLinearPrecision, LinearPrecisionEnum)
                        unitsPrecisionTrace = originalLengthDisplayPrecision.ToString() & "->" & CInt(units.LengthDisplayPrecision).ToString()
                    End If
                Catch ex As Exception
                    unitsPrecisionTrace = "LengthDisplayPrecision update failed: " & ex.Message
                End Try

                Dim originalPrecision As Integer = CInt(params.LinearDimensionPrecision)
                Dim tempPrecision As Integer = If(originalPrecision = 3, 2, 3)
                params.LinearDimensionPrecision = tempPrecision
                params.LinearDimensionPrecision = kZeroDecimalPlaceLinearPrecision

                Dim originalDisplay As Integer = CInt(params.DimensionDisplayType)
                Dim tempDisplay As Integer = If(originalDisplay = kDimensionDisplayAsPreciseValue,
                                               kDimensionDisplayAsValue,
                                               kDimensionDisplayAsPreciseValue)
                params.DimensionDisplayType = CType(tempDisplay, DimensionDisplayTypeEnum)
                params.DimensionDisplayType = CType(kDimensionDisplayAsValue, DimensionDisplayTypeEnum)

                Dim originalExpression As Boolean = params.DisplayParameterAsExpression
                params.DisplayParameterAsExpression = False
                params.DisplayParameterAsExpression = True

                Dim hadLinearPrecisionDelta As Boolean = originalPrecision <> kZeroDecimalPlaceLinearPrecision
                Dim hadDisplayTypeDelta As Boolean = originalDisplay <> kDimensionDisplayAsValue
                Dim hadExpressionDelta As Boolean = (Not originalExpression)
                hadAnySettingDelta = hadUomDisplayDelta OrElse hadLinearPrecisionDelta OrElse hadDisplayTypeDelta OrElse hadExpressionDelta

                AddInDiagnostics.Log("ParameterExporterTools", "Update settings applied | Part='" & partPath & "' | UoMDisplayPrecision " & unitsPrecisionTrace & " | Precision " & originalPrecision.ToString() & "->" & CInt(params.LinearDimensionPrecision).ToString() & " | Display " & originalDisplay.ToString() & "->" & CInt(params.DimensionDisplayType).ToString() & " | Expression " & originalExpression.ToString() & "->" & params.DisplayParameterAsExpression.ToString())

                Dim dirtyParamTrace As String = String.Empty
                Dim dirtyParamOk As Boolean = TriggerDirtyFlagByDummyParameter(params, dirtyParamTrace)
                AddInDiagnostics.Log("ParameterExporterTools", "Dirty flag parameter | Part='" & partPath & "' | Success=" & dirtyParamOk.ToString() & " | " & dirtyParamTrace)

                Dim dirtyPropertyTrace As String = String.Empty
                Dim dirtyPropertyOk As Boolean = False
                If Not dirtyParamOk Then
                    dirtyPropertyOk = TriggerDirtyFlagByTempProperty(partDoc, dirtyPropertyTrace)
                    AddInDiagnostics.Log("ParameterExporterTools", "Dirty flag property | Part='" & partPath & "' | Success=" & dirtyPropertyOk.ToString() & " | " & dirtyPropertyTrace)
                End If

                dirtyFlagSucceeded = dirtyParamOk OrElse dirtyPropertyOk

                Dim unitsRefreshTrace As String = String.Empty
                Dim unitsRefreshOk As Boolean = ForceUnitsRefreshEvent(partDoc, unitsRefreshTrace)
                AddInDiagnostics.Log("ParameterExporterTools", "Units refresh | Part='" & partPath & "' | Success=" & unitsRefreshOk.ToString() & " | " & unitsRefreshTrace)

                Dim updateOk As Boolean = False
                Dim updateError As String = String.Empty
                Try
                    partDoc.Update2(True)
                    updateOk = True
                Catch ex As Exception
                    updateError = ex.Message
                End Try

                Dim dirtySetOk As Boolean = False
                Dim dirtyState As String = "UNKNOWN"
                Dim dirtyError As String = String.Empty
                Try
                    partDoc.Dirty = True
                    dirtySetOk = True
                    dirtyState = partDoc.Dirty.ToString()
                Catch ex As Exception
                    dirtyError = ex.Message
                End Try

                Dim saveOk As Boolean = False
                Dim saveError As String = String.Empty
                Try
                    partDoc.Save2(True)
                    saveOk = True
                Catch ex As Exception
                    saveError = ex.Message
                    Throw
                Finally
                    AddInDiagnostics.Log("ParameterExporterTools", "Update settings finalize | Part='" & partPath & "' | Update2=" & updateOk.ToString() & If(String.IsNullOrWhiteSpace(updateError), String.Empty, " | Update2Error='" & updateError & "'") & " | DirtySet=" & dirtySetOk.ToString() & " | DirtyState='" & dirtyState & "'" & If(String.IsNullOrWhiteSpace(dirtyError), String.Empty, " | DirtyError='" & dirtyError & "'") & " | DirtyParam=" & dirtyParamOk.ToString() & " | DirtyProperty=" & dirtyPropertyOk.ToString() & " | DirtyFlagSucceeded=" & dirtyFlagSucceeded.ToString() & " | HadDelta=" & hadAnySettingDelta.ToString() & " | Save=" & saveOk.ToString() & If(String.IsNullOrWhiteSpace(saveError), String.Empty, " | SaveError='" & saveError & "'"))
                End Try

                If hadAnySettingDelta Then
                    status = If(dirtyFlagSucceeded, "UPDATED", "UPDATED_DIRTYFLAG_FAILED")
                Else
                    status = If(dirtyFlagSucceeded, "UPDATED_NO_DELTA", "UPDATED_NO_DELTA_DIRTYFLAG_FAILED")
                End If
                Return True

            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "UpdatePlateDocumentSettingsWithDirtyRefresh failed | Part='" & partPath & "' | " & ex.Message)
                status = "FAILED"
                dirtyFlagSucceeded = False
                hadAnySettingDelta = False
                Return False
            Finally
                If openedByTool AndAlso partDoc IsNot Nothing Then
                    Try
                        partDoc.Close(True)
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Function TriggerDirtyFlagByDummyParameter(ByVal params As Parameters,
                                                          ByRef trace As String) As Boolean
            trace = String.Empty

            If params Is Nothing Then
                trace = "Parameters unavailable"
                Return False
            End If

            Try
                Dim userParams As UserParameters = params.UserParameters
                If userParams Is Nothing Then
                    trace = "UserParameters unavailable"
                    Return False
                End If

                Try
                    Dim existing As UserParameter = userParams.Item(kBomRefreshDummyParameterName)
                    If existing IsNot Nothing Then
                        existing.Delete()
                    End If
                Catch
                End Try

                Dim dummy As UserParameter = userParams.AddByValue(kBomRefreshDummyParameterName,
                                                                    0,
                                                                    "mm")

                If dummy Is Nothing Then
                    trace = "Dummy parameter creation returned Nothing"
                    Return False
                End If

                dummy.Value = 1
                dummy.Delete()

                trace = "Created, changed, and removed dummy parameter"
                Return True
            Catch ex As Exception
                trace = "Exception='" & ex.Message & "'"
                Return False
            End Try
        End Function

        Private Function TriggerDirtyFlagByTempProperty(ByVal partDoc As PartDocument,
                                                        ByRef trace As String) As Boolean
            trace = String.Empty

            If partDoc Is Nothing Then
                trace = "Part document unavailable"
                Return False
            End If

            Try
                Dim userProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
                If userProps Is Nothing Then
                    trace = "User defined property set unavailable"
                    Return False
                End If

                Try
                    Dim existingProp As Inventor.Property = userProps.Item(kBomRefreshDummyPropertyName)
                    If existingProp IsNot Nothing Then
                        existingProp.Delete()
                    End If
                Catch
                End Try

                Dim tempProp As Inventor.Property = userProps.Add("1", kBomRefreshDummyPropertyName)
                If tempProp Is Nothing Then
                    trace = "Temporary property creation returned Nothing"
                    Return False
                End If

                tempProp.Value = "2"
                tempProp.Delete()

                trace = "Created, changed, and removed temporary custom property"
                Return True
            Catch ex As Exception
                trace = "Exception='" & ex.Message & "'"
                Return False
            End Try
        End Function

        Private Function ForceUnitsRefreshEvent(ByVal partDoc As PartDocument,
                                                ByRef trace As String) As Boolean
            trace = String.Empty

            If partDoc Is Nothing Then
                trace = "No part document"
                Return False
            End If

            Try
                Dim units As UnitsOfMeasure = partDoc.UnitsOfMeasure
                If units Is Nothing Then
                    trace = "UnitsOfMeasure unavailable"
                    Return False
                End If

                Dim originalUnits As UnitsTypeEnum = CType(units.LengthUnits, UnitsTypeEnum)
                Dim temporaryUnits As UnitsTypeEnum = If(originalUnits = UnitsTypeEnum.kMillimeterLengthUnits,
                                                         UnitsTypeEnum.kCentimeterLengthUnits,
                                                         UnitsTypeEnum.kMillimeterLengthUnits)

                Dim originalDisplayUnits As UnitsTypeEnum = CType(units.LengthDisplayUnits, UnitsTypeEnum)
                Dim temporaryDisplayUnits As UnitsTypeEnum = If(originalDisplayUnits = UnitsTypeEnum.kMillimeterLengthUnits,
                                                                UnitsTypeEnum.kCentimeterLengthUnits,
                                                                UnitsTypeEnum.kMillimeterLengthUnits)

                If temporaryUnits = originalUnits Then
                    temporaryUnits = UnitsTypeEnum.kInchLengthUnits
                End If

                If temporaryDisplayUnits = originalDisplayUnits Then
                    temporaryDisplayUnits = UnitsTypeEnum.kInchLengthUnits
                End If

                Dim switchedToTemporary As Boolean = False
                Dim switchedBack As Boolean = False
                Dim temporaryUpdateOk As Boolean = False
                Dim restoreUpdateOk As Boolean = False
                Dim displaySwitchedToTemporary As Boolean = False
                Dim displaySwitchedBack As Boolean = False
                Dim displayTemporaryUpdateOk As Boolean = False
                Dim displayRestoreUpdateOk As Boolean = False

                units.LengthUnits = temporaryUnits
                switchedToTemporary = True
                Try
                    partDoc.Update2(True)
                    temporaryUpdateOk = True
                Catch
                End Try

                units.LengthUnits = originalUnits
                switchedBack = True
                Try
                    partDoc.Update2(True)
                    restoreUpdateOk = True
                Catch
                End Try

                Try
                    units.LengthDisplayUnits = temporaryDisplayUnits
                    displaySwitchedToTemporary = True
                Catch
                End Try

                If displaySwitchedToTemporary Then
                    Try
                        partDoc.Update2(True)
                        displayTemporaryUpdateOk = True
                    Catch
                    End Try
                End If

                Try
                    units.LengthDisplayUnits = originalDisplayUnits
                    displaySwitchedBack = True
                Catch
                End Try

                If displaySwitchedBack Then
                    Try
                        partDoc.Update2(True)
                        displayRestoreUpdateOk = True
                    Catch
                    End Try
                End If

                trace = "OriginalUnits='" & originalUnits.ToString() & "' | TemporaryUnits='" & temporaryUnits.ToString() &
                        "' | SetTemp=" & switchedToTemporary.ToString() & " | TempUpdate=" & temporaryUpdateOk.ToString() &
                        " | Restored=" & switchedBack.ToString() & " | RestoreUpdate=" & restoreUpdateOk.ToString() &
                        " | CurrentUnits='" & CType(units.LengthUnits, UnitsTypeEnum).ToString() &
                        "' | OriginalDisplayUnits='" & originalDisplayUnits.ToString() & "' | TemporaryDisplayUnits='" & temporaryDisplayUnits.ToString() &
                        "' | DisplaySetTemp=" & displaySwitchedToTemporary.ToString() & " | DisplayTempUpdate=" & displayTemporaryUpdateOk.ToString() &
                        " | DisplayRestored=" & displaySwitchedBack.ToString() & " | DisplayRestoreUpdate=" & displayRestoreUpdateOk.ToString() &
                        " | CurrentDisplayUnits='" & CType(units.LengthDisplayUnits, UnitsTypeEnum).ToString() & "'"

                Return (switchedToTemporary AndAlso switchedBack) OrElse (displaySwitchedToTemporary AndAlso displaySwitchedBack)
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "ForceUnitsRefreshEvent failed | Part='" & partDoc.FullFileName & "' | " & ex.Message)
                trace = "Exception='" & ex.Message & "'"
                Return False
            End Try
        End Function

        Private Sub RefreshAssemblyBomPrecision(ByVal asmDoc As AssemblyDocument)
            If asmDoc Is Nothing Then
                Return
            End If

            AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision START | Assembly='" & SafeDocumentPath(asmDoc) & "'")

            Dim activateOk As Boolean = False
            Try
                asmDoc.Activate()
                activateOk = True
            Catch
            End Try

            Dim rebuildClassicOk As Boolean = False
            Dim rebuildClassicError As String = String.Empty

            Try
                asmDoc.Rebuild()
                rebuildClassicOk = True
            Catch ex As Exception
                rebuildClassicError = ex.Message
            End Try

            Dim rebuildOk As Boolean = False
            Dim rebuildError As String = String.Empty

            Try
                asmDoc.Rebuild2(True)
                rebuildOk = True
            Catch ex As Exception
                rebuildError = ex.Message
            End Try

            Dim updateOk As Boolean = False
            Dim updateError As String = String.Empty

            Try
                asmDoc.Update2(True)
                updateOk = True
            Catch ex As Exception
                updateError = ex.Message
            End Try

            Try
                Dim bom As BOM = asmDoc.ComponentDefinition.BOM
                If bom Is Nothing Then
                    AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision end (no BOM) | Assembly='" & SafeDocumentPath(asmDoc) & "' | Activated=" & activateOk.ToString() & " | Rebuild=" & rebuildOk.ToString() & If(String.IsNullOrWhiteSpace(rebuildError), String.Empty, " | RebuildError='" & rebuildError & "'") & " | Update=" & updateOk.ToString() & If(String.IsNullOrWhiteSpace(updateError), String.Empty, " | UpdateError='" & updateError & "'"))
                    Return
                End If

                Dim originalStructured As Boolean = bom.StructuredViewEnabled
                Dim originalPartsOnly As Boolean = bom.PartsOnlyViewEnabled
                Dim requiresUpdateBefore As Boolean = False
                Dim assemblyDisplayToggleOk As Boolean = False
                Dim assemblyDisplayToggleTrace As String = String.Empty

                Try
                    requiresUpdateBefore = bom.RequiresUpdate
                Catch
                End Try

                Try
                    Dim asmUnits As UnitsOfMeasure = asmDoc.UnitsOfMeasure
                    If asmUnits IsNot Nothing Then
                        Dim originalDisplayUnits As UnitsTypeEnum = CType(asmUnits.LengthDisplayUnits, UnitsTypeEnum)
                        Dim temporaryDisplayUnits As UnitsTypeEnum = If(originalDisplayUnits = UnitsTypeEnum.kMillimeterLengthUnits,
                                                                       UnitsTypeEnum.kCentimeterLengthUnits,
                                                                       UnitsTypeEnum.kMillimeterLengthUnits)
                        If temporaryDisplayUnits = originalDisplayUnits Then
                            temporaryDisplayUnits = UnitsTypeEnum.kInchLengthUnits
                        End If

                        asmUnits.LengthDisplayUnits = temporaryDisplayUnits
                        Try
                            asmDoc.Update2(True)
                        Catch
                        End Try

                        asmUnits.LengthDisplayUnits = originalDisplayUnits
                        Try
                            asmDoc.Update2(True)
                        Catch
                        End Try

                        assemblyDisplayToggleOk = True
                        assemblyDisplayToggleTrace = "Original='" & originalDisplayUnits.ToString() & "' | Temp='" & temporaryDisplayUnits.ToString() & "' | Current='" & CType(asmUnits.LengthDisplayUnits, UnitsTypeEnum).ToString() & "'"
                    Else
                        assemblyDisplayToggleTrace = "UnitsOfMeasure unavailable"
                    End If
                Catch ex As Exception
                    assemblyDisplayToggleTrace = "Exception='" & ex.Message & "'"
                End Try

                Try
                    If Not bom.StructuredViewEnabled Then
                        bom.StructuredViewEnabled = True
                    End If
                Catch
                End Try

                Try
                    If Not bom.PartsOnlyViewEnabled Then
                        bom.PartsOnlyViewEnabled = True
                    End If
                Catch
                End Try

                Dim structuredToggleOk As Boolean = False
                Dim partsOnlyToggleOk As Boolean = False

                Try
                    bom.StructuredViewEnabled = Not originalStructured
                    bom.StructuredViewEnabled = originalStructured
                    structuredToggleOk = True
                Catch
                End Try

                Try
                    bom.PartsOnlyViewEnabled = Not originalPartsOnly
                    bom.PartsOnlyViewEnabled = originalPartsOnly
                    partsOnlyToggleOk = True
                Catch
                End Try

                Dim structuredRowsOk As Boolean = False
                Dim structuredRowsCount As Integer = -1
                Dim partsOnlyRowsOk As Boolean = False
                Dim partsOnlyRowsCount As Integer = -1
                Dim structuredRenumberOk As Boolean = False
                Dim structuredRenumberError As String = String.Empty
                Dim partsOnlyRenumberOk As Boolean = False
                Dim partsOnlyRenumberError As String = String.Empty

                Try
                    Dim structuredView As BOMView = bom.BOMViews.Item("Structured")
                    If structuredView IsNot Nothing Then
                        structuredRowsCount = structuredView.BOMRows.Count
                        structuredRowsOk = True

                        Try
                            Dim structuredViewObj As Object = structuredView
                            structuredViewObj.Renumber()
                            structuredRenumberOk = True
                        Catch ex As Exception
                            structuredRenumberError = ex.Message
                        End Try
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision structured-row access failed | Assembly='" & SafeDocumentPath(asmDoc) & "' | " & ex.Message)
                End Try

                Try
                    Dim partsOnlyView As BOMView = bom.BOMViews.Item("Parts Only")
                    If partsOnlyView IsNot Nothing Then
                        partsOnlyRowsCount = partsOnlyView.BOMRows.Count
                        partsOnlyRowsOk = True

                        Try
                            Dim partsOnlyViewObj As Object = partsOnlyView
                            partsOnlyViewObj.Renumber()
                            partsOnlyRenumberOk = True
                        Catch ex As Exception
                            partsOnlyRenumberError = ex.Message
                        End Try
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision parts-only row access failed | Assembly='" & SafeDocumentPath(asmDoc) & "' | " & ex.Message)
                End Try

                Dim requiresUpdateAfter As Boolean = False
                If bom.RequiresUpdate Then
                    requiresUpdateAfter = True
                    Try
                        asmDoc.Update2(True)
                    Catch
                    End Try
                Else
                    Try
                        requiresUpdateAfter = bom.RequiresUpdate
                    Catch
                    End Try
                End If

                Dim finalUpdateOk As Boolean = False
                Dim finalUpdateError As String = String.Empty
                Try
                    asmDoc.Update2(True)
                    finalUpdateOk = True
                Catch ex As Exception
                    finalUpdateError = ex.Message
                End Try

                Dim saveOk As Boolean = False
                Dim saveError As String = String.Empty
                Try
                    asmDoc.Save2(True)
                    saveOk = True
                Catch ex As Exception
                    saveError = ex.Message
                End Try

                AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision COMPLETE | Assembly='" & SafeDocumentPath(asmDoc) & "' | Activated=" & activateOk.ToString() & " | RebuildClassic=" & rebuildClassicOk.ToString() & If(String.IsNullOrWhiteSpace(rebuildClassicError), String.Empty, " | RebuildClassicError='" & rebuildClassicError & "'") & " | Rebuild2=" & rebuildOk.ToString() & If(String.IsNullOrWhiteSpace(rebuildError), String.Empty, " | Rebuild2Error='" & rebuildError & "'") & " | Update=" & updateOk.ToString() & If(String.IsNullOrWhiteSpace(updateError), String.Empty, " | UpdateError='" & updateError & "'") & " | AssemblyDisplayToggle=" & assemblyDisplayToggleOk.ToString() & " | AssemblyDisplayTrace=" & assemblyDisplayToggleTrace & " | StructuredToggle=" & structuredToggleOk.ToString() & " | PartsOnlyToggle=" & partsOnlyToggleOk.ToString() & " | StructuredRowsOk=" & structuredRowsOk.ToString() & " | StructuredRows=" & structuredRowsCount.ToString() & " | StructuredRenumber=" & structuredRenumberOk.ToString() & If(String.IsNullOrWhiteSpace(structuredRenumberError), String.Empty, " | StructuredRenumberError='" & structuredRenumberError & "'") & " | PartsOnlyRowsOk=" & partsOnlyRowsOk.ToString() & " | PartsOnlyRows=" & partsOnlyRowsCount.ToString() & " | PartsOnlyRenumber=" & partsOnlyRenumberOk.ToString() & If(String.IsNullOrWhiteSpace(partsOnlyRenumberError), String.Empty, " | PartsOnlyRenumberError='" & partsOnlyRenumberError & "'") & " | RequiresUpdateBefore=" & requiresUpdateBefore.ToString() & " | RequiresUpdateAfter=" & requiresUpdateAfter.ToString() & " | FinalUpdate=" & finalUpdateOk.ToString() & If(String.IsNullOrWhiteSpace(finalUpdateError), String.Empty, " | FinalUpdateError='" & finalUpdateError & "'") & " | Save=" & saveOk.ToString() & If(String.IsNullOrWhiteSpace(saveError), String.Empty, " | SaveError='" & saveError & "'"))
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "RefreshAssemblyBomPrecision failed | Assembly='" & asmDoc.DisplayName & "' | " & ex.Message)
            End Try
        End Sub

        Private Function SafeDocumentPath(ByVal doc As Document) As String
            If doc Is Nothing Then
                Return String.Empty
            End If

            Try
                Dim path As String = doc.FullFileName
                If Not String.IsNullOrWhiteSpace(path) Then
                    Return path
                End If
            Catch
            End Try

            Try
                Return doc.DisplayName
            Catch
            End Try

            Return String.Empty
        End Function

        Private Function ScanAssemblyForNonPlateParts(ByVal asmDoc As AssemblyDocument) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Try
                    If occ Is Nothing OrElse occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then
                        Continue For
                    End If

                    Dim refDoc As Document = occ.Definition.Document
                    If refDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                        Continue For
                    End If

                    Dim partNumber As String = GetPartNumber(refDoc)
                    Dim upperPartNo As String = If(partNumber, String.Empty).ToUpperInvariant()

                    If Not upperPartNo.Contains("PL") AndAlso Not upperPartNo.Contains("S355JR") Then
                        Dim fullPath As String = refDoc.FullFileName
                        If Not String.IsNullOrWhiteSpace(fullPath) AndAlso Not result.ContainsKey(fullPath) Then
                            result.Add(fullPath, partNumber)
                        End If
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("ParameterExporterTools", "Scan non-plate failed for occurrence '" & occ.Name & "' | " & ex.Message)
                End Try
            Next

            Return result
        End Function

        Private Function ScanAssemblyForPlateParts(ByVal asmDoc As AssemblyDocument) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Try
                    If occ Is Nothing OrElse occ.Definition Is Nothing OrElse occ.Definition.Document Is Nothing Then
                        Continue For
                    End If

                    Dim refDoc As Document = occ.Definition.Document
                    If refDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                        Continue For
                    End If

                    Dim partNumber As String = GetPartNumber(refDoc)
                    Dim description As String = GetDescription(refDoc)
                    Dim searchText As String = (If(partNumber, String.Empty) & " " & If(description, String.Empty)).ToUpperInvariant()

                    Dim isPlate As Boolean = searchText.Contains("PL") OrElse
                                             searchText.Contains("PLATE") OrElse
                                             searchText.Contains("VRN") OrElse
                                             searchText.Contains("S355JR")

                    If isPlate Then
                        Dim fullPath As String = refDoc.FullFileName
                        If Not String.IsNullOrWhiteSpace(fullPath) AndAlso Not result.ContainsKey(fullPath) Then
                            result.Add(fullPath, partNumber)
                        End If
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("ParameterExporterTools", "Scan plate failed for occurrence '" & occ.Name & "' | " & ex.Message)
                End Try
            Next

            Return result
        End Function

        Private Sub CollectNonPlatePartsMissingLength(ByVal occurrences As ComponentOccurrences,
                                                      ByVal result As Dictionary(Of String, String))
            If occurrences Is Nothing Then
                Return
            End If

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ Is Nothing OrElse occ.Suppressed Then
                        Continue For
                    End If
                Catch
                End Try

                Dim refDoc As Document = Nothing
                Try
                    refDoc = occ.Definition.Document
                Catch
                    refDoc = Nothing
                End Try

                If refDoc Is Nothing Then
                    Continue For
                End If

                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    CollectNonPlatePartsMissingLength(occ.SubOccurrences, result)
                ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim description As String = GetDescription(refDoc)
                    If description.IndexOf("PL ", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Continue For
                    End If

                    If Not HasUserLengthOrLength2(TryCast(refDoc, PartDocument)) Then
                        Dim fullPath As String = String.Empty
                        Try
                            fullPath = refDoc.FullFileName
                        Catch
                            fullPath = String.Empty
                        End Try

                        If Not String.IsNullOrWhiteSpace(fullPath) AndAlso Not result.ContainsKey(fullPath) Then
                            result.Add(fullPath, refDoc.DisplayName)
                        End If
                    End If
                End If
            Next
        End Sub

        Private Sub CollectPlatePartsByDescription(ByVal occurrences As ComponentOccurrences,
                                                   ByVal result As Dictionary(Of String, String))
            If occurrences Is Nothing Then
                Return
            End If

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ Is Nothing OrElse occ.Suppressed Then
                        Continue For
                    End If
                Catch
                End Try

                Dim refDoc As Document = Nothing
                Try
                    refDoc = occ.Definition.Document
                Catch
                    refDoc = Nothing
                End Try

                If refDoc Is Nothing Then
                    Continue For
                End If

                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    CollectPlatePartsByDescription(occ.SubOccurrences, result)
                ElseIf refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partNumber As String = GetPartNumber(refDoc)
                    Dim description As String = GetDescription(refDoc)
                    Dim searchText As String = (If(partNumber, String.Empty) & " " & If(description, String.Empty)).ToUpperInvariant()
                    Dim isPlate As Boolean = searchText.Contains("PL") OrElse searchText.Contains("VRN") OrElse searchText.Contains("S355JR")
                    If isPlate Then
                        Dim fullPath As String = String.Empty
                        Try
                            fullPath = refDoc.FullFileName
                        Catch
                            fullPath = String.Empty
                        End Try

                        If Not String.IsNullOrWhiteSpace(fullPath) AndAlso Not result.ContainsKey(fullPath) Then
                            result.Add(fullPath, refDoc.DisplayName)
                        End If
                    End If
                End If
            Next
        End Sub

        Private Function HasUserLengthOrLength2(ByVal partDoc As PartDocument) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            Try
                Dim userParams As UserParameters = partDoc.ComponentDefinition.Parameters.UserParameters
                For Each up As UserParameter In userParams
                    If String.Equals(up.Name, "Length", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(up.Name, "Length2", StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        Private Function FixLength2ForPartPath(ByVal partPath As String, ByRef status As String) As Boolean
            status = String.Empty
            If String.IsNullOrWhiteSpace(partPath) Then
                status = "FAILED"
                Return False
            End If

            Dim partDoc As PartDocument = Nothing
            Dim openedByTool As Boolean = False

            Try
                For Each doc As Document In m_InventorApp.Documents
                    If String.Equals(doc.FullFileName, partPath, StringComparison.OrdinalIgnoreCase) Then
                        partDoc = TryCast(doc, PartDocument)
                        Exit For
                    End If
                Next

                If partDoc Is Nothing Then
                    partDoc = TryCast(m_InventorApp.Documents.Open(partPath, False), PartDocument)
                    openedByTool = partDoc IsNot Nothing
                End If

                If partDoc Is Nothing Then
                    status = "FAILED"
                    Return False
                End If

                Dim detail As String = String.Empty
                If Not EnsureLength2FromLongestModelDimension(partDoc, detail) Then
                    status = If(detail = "SKIPPED", "SKIPPED", "FAILED")
                    Return False
                End If

                Try
                    partDoc.Save2(True)
                Catch
                End Try

                status = "UPDATED"
                Return True
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "FixLength2ForPartPath failed | Part='" & partPath & "' | " & ex.Message)
                status = "FAILED"
                Return False
            Finally
                If openedByTool AndAlso partDoc IsNot Nothing Then
                    Try
                        partDoc.Close(True)
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Function EnsureLength2FromLongestModelDimension(ByVal partDoc As PartDocument, ByRef detail As String) As Boolean
            detail = String.Empty
            If partDoc Is Nothing Then
                detail = "FAILED"
                Return False
            End If

            Try
                Dim parameters As Parameters = partDoc.ComponentDefinition.Parameters
                Dim modelParams As ModelParameters = parameters.ModelParameters
                Dim userParams As UserParameters = parameters.UserParameters

                Dim maxParam As Parameter = Nothing
                Dim maxValue As Double = Double.MinValue

                For Each mp As ModelParameter In modelParams
                    Dim unitsValue As String = String.Empty
                    Try
                        unitsValue = Convert.ToString(mp.Units).Trim().ToLowerInvariant()
                    Catch
                        unitsValue = String.Empty
                    End Try

                    If unitsValue = "mm" OrElse unitsValue = String.Empty Then
                        If mp.ModelValue > maxValue Then
                            maxValue = mp.ModelValue
                            maxParam = mp
                        End If
                    End If
                Next

                If maxParam Is Nothing Then
                    detail = "SKIPPED"
                    Return False
                End If

                Dim length2Param As UserParameter = Nothing
                Try
                    length2Param = userParams.Item("Length2")
                    If length2Param IsNot Nothing Then
                        length2Param.Delete()
                    End If
                Catch
                End Try

                length2Param = userParams.AddByValue("Length2", maxParam.Value, UnitsTypeEnum.kMillimeterLengthUnits)
                length2Param.Expression = maxParam.Name
                maxParam.Expression = "Length2"

                detail = "UPDATED"
                Return True
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "EnsureLength2FromLongestModelDimension failed | Part='" & partDoc.FullFileName & "' | " & ex.Message)
                detail = "FAILED"
                Return False
            End Try
        End Function

        Private Function SetPlateBomProperties(ByVal partPath As String, ByRef status As String) As Boolean
            status = String.Empty
            If String.IsNullOrWhiteSpace(partPath) Then
                status = "FAILED"
                Return False
            End If

            Dim partDoc As PartDocument = Nothing
            Dim openedByTool As Boolean = False

            Try
                For Each doc As Document In m_InventorApp.Documents
                    If String.Equals(doc.FullFileName, partPath, StringComparison.OrdinalIgnoreCase) Then
                        partDoc = TryCast(doc, PartDocument)
                        Exit For
                    End If
                Next

                If partDoc Is Nothing Then
                    partDoc = TryCast(m_InventorApp.Documents.Open(partPath, False), PartDocument)
                    openedByTool = partDoc IsNot Nothing
                End If

                If partDoc Is Nothing Then
                    status = "FAILED"
                    Return False
                End If

                If Not IsSheetMetalPartWithFlatPattern(partDoc) Then
                    status = "SKIPPED"
                    Return False
                End If

                SetCustomUserProperty(partDoc, "LENGTH", "=<sheet metal length>")
                SetCustomUserProperty(partDoc, "WIDTH", "=<sheet metal width>")

                Try
                    partDoc.Save2(True)
                Catch
                End Try

                status = "UPDATED"
                Return True
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "SetPlateBomProperties failed | Part='" & partPath & "' | " & ex.Message)
                status = "FAILED"
                Return False
            Finally
                If openedByTool AndAlso partDoc IsNot Nothing Then
                    Try
                        partDoc.Close(True)
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Function IsSheetMetalPartWithFlatPattern(ByVal partDoc As PartDocument) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            Try
                If String.Equals(partDoc.SubType, "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}", StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Catch
            End Try

            Try
                Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
                If compDef IsNot Nothing Then
                    Dim flatPattern As FlatPattern = compDef.FlatPattern
                    If flatPattern IsNot Nothing Then
                        Return True
                    End If
                End If
            Catch
            End Try

            Return False
        End Function

        Private Sub SetCustomUserProperty(ByVal partDoc As PartDocument, ByVal propName As String, ByVal propValue As String)
            If partDoc Is Nothing OrElse String.IsNullOrWhiteSpace(propName) Then
                Return
            End If

            Dim userProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")

            Try
                Dim existingProp As Inventor.Property = userProps.Item(propName)
                existingProp.Value = propValue
            Catch
                userProps.Add(propValue, propName)
            End Try
        End Sub

        Private Sub ApplyFormulaToPart(ByVal partDoc As PartDocument, ByVal formulaValue As String)
            If partDoc Is Nothing Then
                Throw New InvalidOperationException("Part document is required.")
            End If

            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            If designProps Is Nothing Then
                Throw New InvalidOperationException("Design Tracking Properties not found.")
            End If

            designProps.Item("Description").Value = formulaValue
            designProps.Item("Stock Number").Value = formulaValue

            Try
                partDoc.Save2(True)
            Catch
            End Try
        End Sub

        Private Function ProcessParts(ByVal parts As Dictionary(Of String, String),
                                      ByVal parameterName As String,
                                      ByVal progress As ToolProgressForm,
                                      ByVal startProgress As Integer,
                                      ByVal endProgress As Integer) As ProcessingResults
            Dim results As New ProcessingResults()
            Dim progressRange As Integer = endProgress - startProgress

            For Each kvp As KeyValuePair(Of String, String) In parts
                Dim partPath As String = kvp.Key
                Dim partNumber As String = kvp.Value

                Dim totalDone As Integer = results.TotalProcessed
                Dim currentProgress As Integer = startProgress
                If parts.Count > 0 Then
                    currentProgress = startProgress + CInt((CDbl(totalDone) / CDbl(parts.Count)) * progressRange)
                End If
                If currentProgress > endProgress Then currentProgress = endProgress

                progress.UpdateProgress(currentProgress, "Processing " & partNumber & "...")

                Dim partDoc As Document = Nothing
                Try
                    partDoc = m_InventorApp.Documents.Open(partPath, False)

                    If EnableParameterExport(partDoc, parameterName) Then
                        results.ProcessedCount += 1
                        Try
                            partDoc.Save2(True)
                            results.SavedCount += 1
                        Catch ex As Exception
                            AddInDiagnostics.Log("ParameterExporterTools", "Save failed | Part='" & partPath & "' | " & ex.Message)
                        End Try
                    Else
                        results.SkippedCount += 1
                    End If
                Catch ex As Exception
                    AddInDiagnostics.Log("ParameterExporterTools", "Process failed | Part='" & partPath & "' | " & ex.Message)
                    results.FailedCount += 1
                Finally
                    If partDoc IsNot Nothing Then
                        Try
                            partDoc.Close(True)
                        Catch
                        End Try
                    End If
                End Try
            Next

            Return results
        End Function

        Private Function EnableParameterExport(ByVal partDoc As Document, ByVal parameterName As String) As Boolean
            Try
                If partDoc Is Nothing OrElse partDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    Return False
                End If

                Dim partDef As PartComponentDefinition = TryCast(partDoc.ComponentDefinition, PartComponentDefinition)
                If partDef Is Nothing OrElse partDef.Parameters Is Nothing Then
                    Return False
                End If

                Dim found As Boolean = EnableParameterInCollection(partDef.Parameters.UserParameters, parameterName)
                If Not found Then
                    found = EnableParameterInCollection(partDef.Parameters.ModelParameters, parameterName)
                End If

                Return found
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "Enable parameter failed | Part='" & partDoc.FullFileName & "' | " & ex.Message)
                Return False
            End Try
        End Function

        Private Function EnableParameterInCollection(ByVal parameterCollection As Object,
                                                     ByVal parameterName As String) As Boolean
            If parameterCollection Is Nothing Then
                Return False
            End If

            Try
                For Each param As Parameter In parameterCollection
                    If String.Equals(param.Name, parameterName, StringComparison.OrdinalIgnoreCase) Then
                        If Not param.ExportedToSheet Then
                            param.ExportedToSheet = True
                        End If
                        Return True
                    End If
                Next
            Catch ex As Exception
                AddInDiagnostics.Log("ParameterExporterTools", "Parameter collection scan failed | Param='" & parameterName & "' | " & ex.Message)
            End Try

            Return False
        End Function

        Private Function GetPartNumber(ByVal doc As Document) As String
            Try
                Dim partNumber As String = String.Empty
                If doc IsNot Nothing AndAlso doc.PropertySets IsNot Nothing Then
                    partNumber = Convert.ToString(doc.PropertySets("Design Tracking Properties").Item("Part Number").Value)
                End If

                If String.IsNullOrWhiteSpace(partNumber) Then
                    partNumber = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
                End If

                Return partNumber
            Catch
                Return System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
            End Try
        End Function

        Private Function GetDescription(ByVal doc As Document) As String
            Try
                If doc IsNot Nothing AndAlso doc.PropertySets IsNot Nothing Then
                    Return Convert.ToString(doc.PropertySets("Design Tracking Properties").Item("Description").Value)
                End If
            Catch
            End Try
            Return String.Empty
        End Function

        Private Sub ShowResults(ByVal toolName As String, ByVal results As ProcessingResults)
            Dim message As String = toolName & " - Processing Complete" & vbCrLf & vbCrLf &
                                    "Parts processed successfully: " & results.ProcessedCount.ToString() & vbCrLf &
                                    "Parts saved: " & results.SavedCount.ToString() & vbCrLf &
                                    "Parts skipped (no parameter): " & results.SkippedCount.ToString() & vbCrLf &
                                    "Parts failed: " & results.FailedCount.ToString() & vbCrLf & vbCrLf &
                                    "Total processed: " & results.TotalProcessed.ToString() & vbCrLf & vbCrLf &
                                    "Check the log file for details."

            MsgBox(message, MsgBoxStyle.Information, toolName & " Summary")
        End Sub

        Private Class ProcessingResults
            Public ProcessedCount As Integer = 0
            Public SavedCount As Integer = 0
            Public SkippedCount As Integer = 0
            Public FailedCount As Integer = 0

            Public ReadOnly Property TotalProcessed As Integer
                Get
                    Return ProcessedCount + SkippedCount + FailedCount
                End Get
            End Property
        End Class

    End Class

End Namespace
