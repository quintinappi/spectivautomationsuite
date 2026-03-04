Imports Inventor
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Environment = System.Environment
Imports File = System.IO.File
Imports Path = System.IO.Path

Namespace AssemblyClonerAddIn

    Public Class SheetMetalConverterAssemblyTool

        Private Const kSheetMetalSubType As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
        Private Const kDisplayAsValue As Integer = 34821

        Private Shared ReadOnly ThicknessPatterns As Regex() = {
            New Regex("(\d+(?:\.\d+)?)\s*mm", RegexOptions.IgnoreCase),
            New Regex("THK\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase),
            New Regex("THICKNESS\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase)
        }

        Private ReadOnly m_InventorApp As Inventor.Application
        Private ReadOnly m_Log As StringBuilder
        Private m_LogPath As String

        Private NotInheritable Class PlatePartInfo
            Public Property FullPath As String
            Public Property FileName As String
            Public Property Description As String
            Public Property ThicknessMm As Double
        End Class

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            m_Log = New StringBuilder()
            m_LogPath = String.Empty
        End Sub

        Public Sub Execute()
            Try
                InitializeLog()
                LogMessage("=== SHEET METAL CONVERTER STARTED ===")

                If m_InventorApp Is Nothing OrElse m_InventorApp.ActiveDocument Is Nothing Then
                    MessageBox.Show("No active document! Please open an assembly in Inventor.", "Sheet Metal Converter", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SaveLog()
                    Return
                End If

                If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MessageBox.Show("Please open an ASSEMBLY document (.iam file), not a part." & Environment.NewLine & Environment.NewLine &
                                    "The sheet metal converter needs to scan an assembly to find plate parts." & Environment.NewLine & Environment.NewLine &
                                    "Current document: " & m_InventorApp.ActiveDocument.DisplayName,
                                    "Sheet Metal Converter",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                    SaveLog()
                    Return
                End If

                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                If String.IsNullOrWhiteSpace(asmDoc.FullFileName) OrElse Not File.Exists(asmDoc.FullFileName) Then
                    MessageBox.Show("Save the active assembly before running Sheet Metal Converter.", "Sheet Metal Converter", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    SaveLog()
                    Return
                End If

                LogMessage("Processing assembly: " & asmDoc.FullFileName)

                Dim plateGroups As Dictionary(Of String, List(Of PlatePartInfo)) = ScanAssemblyForPlates(asmDoc)
                If plateGroups.Count = 0 Then
                    MessageBox.Show("No parts containing 'PL' or 'S355JR' found in the BOM." & Environment.NewLine & Environment.NewLine &
                                    "Make sure your parts have 'PL' or 'S355JR' in their Description field," & Environment.NewLine &
                                    "and that thickness is specified (for example: 10mm, 5 mm, THK 8).",
                                    "Sheet Metal Converter",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information)
                    SaveLog()
                    Return
                End If

                Dim totalParts As Integer = CountTotalParts(plateGroups)
                Dim confirm As DialogResult = MessageBox.Show(
                    "SHEET METAL BATCH CONVERTER" & Environment.NewLine & Environment.NewLine &
                    "Found " & totalParts.ToString() & " plate parts to convert." & Environment.NewLine & Environment.NewLine &
                    "This migration runs the same output workflow from the add-in:" & Environment.NewLine &
                    "- Convert to sheet metal" & Environment.NewLine &
                    "- Set thickness" & Environment.NewLine &
                    "- Create/repair flat pattern orientation" & Environment.NewLine &
                    "- Apply PLATE LENGTH / PLATE WIDTH formulas" & Environment.NewLine &
                    "- Add assembly PLATE LENGTH / PLATE WIDTH parameters" & Environment.NewLine & Environment.NewLine &
                    "Proceed?",
                    "Sheet Metal Converter",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information)

                If confirm <> DialogResult.OK Then
                    LogMessage("User cancelled conversion")
                    SaveLog()
                    Return
                End If

                Dim asmPath As String = asmDoc.FullFileName
                Try
                    asmDoc.Close(True)
                    LogMessage("Assembly closed for part processing: " & asmPath)
                Catch ex As Exception
                    LogMessage("ERROR: Could not close assembly before processing: " & ex.Message)
                    MessageBox.Show("Could not close the active assembly before conversion." & Environment.NewLine & ex.Message,
                                    "Sheet Metal Converter",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error)
                    SaveLog()
                    Return
                End Try

                Dim maxLengthMm As Double = 0.0
                Dim maxWidthMm As Double = 0.0
                Dim processedCount As Integer = 0
                Dim failedCount As Integer = 0

                Dim orderedGroups As List(Of KeyValuePair(Of String, List(Of PlatePartInfo))) = SortGroups(plateGroups)

                Using progress As New ToolProgressForm("Sheet Metal Converter (Assembly)")
                    progress.Show()

                    Dim partIndex As Integer = 0
                    For Each group As KeyValuePair(Of String, List(Of PlatePartInfo)) In orderedGroups
                        LogMessage("Processing thickness group " & group.Key & " mm with " & group.Value.Count.ToString() & " part(s)")

                        For Each partInfo As PlatePartInfo In group.Value
                            partIndex += 1
                            Dim pct As Integer = CInt((CDbl(partIndex) / CDbl(totalParts)) * 100.0)
                            progress.UpdateProgress(Math.Max(5, Math.Min(95, pct)), "Processing " & partInfo.FileName & "...")

                            Dim partLengthMm As Double = 0.0
                            Dim partWidthMm As Double = 0.0
                            If ProcessPlatePart(partInfo, partLengthMm, partWidthMm) Then
                                processedCount += 1
                                If partLengthMm > maxLengthMm Then
                                    maxLengthMm = partLengthMm
                                End If
                                If partWidthMm > maxWidthMm Then
                                    maxWidthMm = partWidthMm
                                End If
                            Else
                                failedCount += 1
                                LogMessage("FAILED: " & partInfo.FileName)
                            End If
                        Next
                    Next

                    progress.UpdateProgress(100, "Finalizing...")
                End Using

                Dim reopenedAssembly As AssemblyDocument = Nothing
                Try
                    reopenedAssembly = TryCast(m_InventorApp.Documents.Open(asmPath, True), AssemblyDocument)
                Catch ex As Exception
                    reopenedAssembly = Nothing
                    LogMessage("ERROR: Could not reopen assembly: " & ex.Message)
                End Try

                If failedCount > 0 Then
                    SaveLog()
                    MessageBox.Show("Sheet metal conversion FAILED!" & Environment.NewLine & Environment.NewLine &
                                    "Successfully converted: " & processedCount.ToString() & " parts" & Environment.NewLine &
                                    "Failed conversions: " & failedCount.ToString() & " parts" & Environment.NewLine & Environment.NewLine &
                                    "Assembly parameters were not updated." & Environment.NewLine &
                                    "Log: " & m_LogPath,
                                    "Sheet Metal Converter",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error)
                    Return
                End If

                If reopenedAssembly Is Nothing Then
                    SaveLog()
                    MessageBox.Show("All parts converted, but assembly could not be reopened for parameter update." & Environment.NewLine &
                                    "Assembly: " & asmPath,
                                    "Sheet Metal Converter",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning)
                    Return
                End If

                AddPlateParametersToAssembly(reopenedAssembly, maxLengthMm, maxWidthMm)

                Try
                    reopenedAssembly.Update2(True)
                Catch
                End Try

                Try
                    reopenedAssembly.Save2(True)
                Catch
                End Try

                LogMessage("Assembly parameters updated: PLATE LENGTH=" & maxLengthMm.ToString("0.###", CultureInfo.InvariantCulture) & " mm, PLATE WIDTH=" & maxWidthMm.ToString("0.###", CultureInfo.InvariantCulture) & " mm")

                SaveLog()

                MessageBox.Show("Sheet metal conversion completed!" & Environment.NewLine & Environment.NewLine &
                                "Processed " & processedCount.ToString() & " plate parts." & Environment.NewLine &
                                "Added/updated assembly parameters: PLATE LENGTH and PLATE WIDTH." & Environment.NewLine & Environment.NewLine &
                                "Log: " & m_LogPath,
                                "Sheet Metal Converter",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
            Catch ex As Exception
                LogMessage("FATAL: " & ex.ToString())
                SaveLog()
                Throw
            End Try
        End Sub

        Private Function ScanAssemblyForPlates(ByVal asmDoc As AssemblyDocument) As Dictionary(Of String, List(Of PlatePartInfo))
            Dim groups As New Dictionary(Of String, List(Of PlatePartInfo))(StringComparer.OrdinalIgnoreCase)
            Dim uniqueParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ProcessAssemblyForPlates(asmDoc, uniqueParts, groups, "ROOT")
            Return groups
        End Function

        Private Sub ProcessAssemblyForPlates(ByVal asmDoc As AssemblyDocument,
                                             ByVal uniqueParts As HashSet(Of String),
                                             ByVal groups As Dictionary(Of String, List(Of PlatePartInfo)),
                                             ByVal asmLevel As String)
            If asmDoc Is Nothing Then
                Return
            End If

            Dim occurrences As ComponentOccurrences = Nothing
            Try
                occurrences = asmDoc.ComponentDefinition.Occurrences
            Catch ex As Exception
                LogMessage("WARNING: Could not read occurrences for assembly " & asmDoc.DisplayName & ": " & ex.Message)
                Return
            End Try

            For index As Integer = 1 To occurrences.Count
                Dim occ As ComponentOccurrence = Nothing
                Try
                    occ = occurrences.Item(index)
                Catch
                    occ = Nothing
                End Try

                If occ Is Nothing Then
                    Continue For
                End If

                Dim isSuppressed As Boolean = False
                Try
                    isSuppressed = occ.Suppressed
                Catch
                    isSuppressed = False
                End Try

                If isSuppressed Then
                    Continue For
                End If

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch
                    doc = Nothing
                End Try

                If doc Is Nothing Then
                    Continue For
                End If

                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    ProcessPartOccurrence(TryCast(doc, PartDocument), uniqueParts, groups)
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    ProcessAssemblyForPlates(TryCast(doc, AssemblyDocument), uniqueParts, groups, asmLevel & ">" & doc.DisplayName)
                End If
            Next
        End Sub

        Private Sub ProcessPartOccurrence(ByVal partDoc As PartDocument,
                                          ByVal uniqueParts As HashSet(Of String),
                                          ByVal groups As Dictionary(Of String, List(Of PlatePartInfo)))
            If partDoc Is Nothing Then
                Return
            End If

            Dim fullPath As String = partDoc.FullFileName
            If String.IsNullOrWhiteSpace(fullPath) Then
                Return
            End If

            If uniqueParts.Contains(fullPath) Then
                Return
            End If
            uniqueParts.Add(fullPath)

            Dim description As String = GetDescriptionFromIProperty(partDoc)
            If String.IsNullOrWhiteSpace(description) Then
                Return
            End If

            Dim upperDescription As String = description.ToUpperInvariant()
            If upperDescription.IndexOf("PL", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
               upperDescription.IndexOf("S355JR", StringComparison.OrdinalIgnoreCase) < 0 Then
                Return
            End If

            Dim thicknessMm As Double = 0.0
            If Not TryExtractThicknessMm(description, thicknessMm) Then
                LogMessage("WARNING: Could not extract thickness from description for " & Path.GetFileName(fullPath) & ": " & description)
                Return
            End If

            Dim thicknessKey As String = thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)
            If Not groups.ContainsKey(thicknessKey) Then
                groups(thicknessKey) = New List(Of PlatePartInfo)()
            End If

            groups(thicknessKey).Add(New PlatePartInfo With {
                .FullPath = fullPath,
                .FileName = Path.GetFileName(fullPath),
                .Description = description,
                .ThicknessMm = thicknessMm
            })
        End Sub

        Private Function ProcessPlatePart(ByVal partInfo As PlatePartInfo,
                                          ByRef outLengthMm As Double,
                                          ByRef outWidthMm As Double) As Boolean
            outLengthMm = 0.0
            outWidthMm = 0.0

            If partInfo Is Nothing OrElse String.IsNullOrWhiteSpace(partInfo.FullPath) Then
                Return False
            End If

            Dim partDoc As PartDocument = Nothing
            Dim openedVisible As Boolean = False

            Try
                partDoc = TryCast(m_InventorApp.Documents.Open(partInfo.FullPath, False), PartDocument)
                If partDoc Is Nothing Then
                    LogMessage("ERROR: Could not open part: " & partInfo.FullPath)
                    Return False
                End If

                Dim largestFace As Face = GetLargestPlanarFace(partDoc)

                If Not EnsurePartIsSheetMetal(partDoc, largestFace, openedVisible) Then
                    LogMessage("ERROR: Conversion to sheet metal failed for " & partInfo.FileName)
                    Return False
                End If

                Dim thicknessCm As Double = GetThicknessFromGeometryCm(partDoc)
                If thicknessCm <= 0 Then
                    thicknessCm = partInfo.ThicknessMm / 10.0
                End If

                If Not SetSheetMetalThickness(partDoc, thicknessCm) Then
                    LogMessage("WARNING: Could not set thickness for " & partInfo.FileName)
                End If

                If largestFace Is Nothing Then
                    largestFace = GetLargestPlanarFace(partDoc)
                End If

                If Not EnsureFlatPattern(partDoc, largestFace, outLengthMm, outWidthMm) Then
                    LogMessage("ERROR: Could not create/read flat pattern for " & partInfo.FileName)
                    Return False
                End If

                FixFlatPatternOrientation(partDoc, largestFace, outLengthMm, outWidthMm)

                AddPlateCustomProperties(partDoc)
                ApplyDocumentSettingsForZeroDecimals(partDoc)

                Try
                    partDoc.Update2(True)
                Catch
                End Try

                partDoc.Save2(True)
                LogMessage("SUCCESS: " & partInfo.FileName & " | " & outLengthMm.ToString("0.###", CultureInfo.InvariantCulture) & " x " & outWidthMm.ToString("0.###", CultureInfo.InvariantCulture) & " mm")
                Return True
            Catch ex As Exception
                LogMessage("ERROR processing " & partInfo.FileName & ": " & ex.Message)
                Return False
            Finally
                If partDoc IsNot Nothing Then
                    Try
                        partDoc.Close(True)
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Function EnsurePartIsSheetMetal(ByRef partDoc As PartDocument,
                                                ByVal largestFace As Face,
                                                ByRef openedVisible As Boolean) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            If String.Equals(partDoc.SubType, kSheetMetalSubType, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Try
                partDoc.SubType = kSheetMetalSubType
                Try
                    partDoc.Update2(True)
                Catch
                End Try

                If String.Equals(partDoc.SubType, kSheetMetalSubType, StringComparison.OrdinalIgnoreCase) Then
                    LogMessage("Converted via PartDocument.SubType API: " & partDoc.DisplayName)
                    Return True
                End If
            Catch ex As Exception
                LogMessage("Direct SubType conversion failed for " & partDoc.DisplayName & ": " & ex.Message)
            End Try

            Dim partPath As String = partDoc.FullFileName
            Try
                partDoc.Close(True)
            Catch
            End Try

            partDoc = TryCast(m_InventorApp.Documents.Open(partPath, True), PartDocument)
            openedVisible = True
            If partDoc Is Nothing Then
                Return False
            End If

            Return RunLegacyConvertCommandFlow(partDoc, largestFace)
        End Function

        Private Function RunLegacyConvertCommandFlow(ByVal partDoc As PartDocument,
                                                     ByVal selectedFace As Face) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            Try
                partDoc.Activate()
            Catch
            End Try

            Dim largestFace As Face = selectedFace
            If largestFace Is Nothing Then
                largestFace = GetLargestPlanarFace(partDoc)
            End If

            If largestFace IsNot Nothing Then
                Try
                    Dim selectSet As SelectSet = partDoc.SelectSet
                    selectSet.Clear()
                    selectSet.Select(largestFace)
                Catch
                End Try
            End If

            Dim convertCmd As ControlDefinition = Nothing
            Try
                convertCmd = m_InventorApp.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
            Catch
                convertCmd = Nothing
            End Try

            If convertCmd Is Nothing OrElse Not convertCmd.Enabled Then
                LogMessage("ERROR: PartConvertToSheetMetalCmd not available")
                Return False
            End If

            Try
                convertCmd.Execute()
            Catch ex As Exception
                LogMessage("ERROR: Convert command execution failed: " & ex.Message)
                Return False
            End Try

            Dim result As DialogResult = MessageBox.Show(
                "ACTION REQUIRED" & Environment.NewLine & Environment.NewLine &
                "Part: " & partDoc.DisplayName & Environment.NewLine & Environment.NewLine &
                "1) The Convert to Sheet Metal dialog is now open" & Environment.NewLine &
                "2) Click once on the highlighted large face" & Environment.NewLine &
                "3) Click OK here to continue",
                "Sheet Metal Converter",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Exclamation)

            If result <> DialogResult.OK Then
                Return False
            End If

            Try
                SendKeys.SendWait("{ENTER}")
            Catch
            End Try

            Try
                partDoc.Update2(True)
            Catch
            End Try

            Return String.Equals(partDoc.SubType, kSheetMetalSubType, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function SetSheetMetalThickness(ByVal partDoc As PartDocument, ByVal thicknessCm As Double) As Boolean
            If partDoc Is Nothing Then
                Return False
            End If

            Try
                Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef Is Nothing Then
                    Return False
                End If

                Try
                    smDef.UseSheetMetalStyleThickness = False
                Catch
                End Try

                Try
                    smDef.Thickness.Value = thicknessCm
                Catch
                    Return False
                End Try

                Try
                    partDoc.Update2(True)
                Catch
                End Try

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Function EnsureFlatPattern(ByVal partDoc As PartDocument,
                                           ByVal largestFace As Face,
                                           ByRef outLengthMm As Double,
                                           ByRef outWidthMm As Double) As Boolean
            outLengthMm = 0.0
            outWidthMm = 0.0

            If partDoc Is Nothing Then
                Return False
            End If

            Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
            If smDef Is Nothing Then
                Return False
            End If

            If Not smDef.HasFlatPattern Then
                Try
                    If largestFace IsNot Nothing Then
                        Try
                            Dim smDefObj As Object = smDef
                            smDefObj.Unfold(largestFace)
                        Catch
                            smDef.Unfold()
                        End Try
                    Else
                        smDef.Unfold()
                    End If
                Catch
                    Dim cmd As ControlDefinition = Nothing
                    Try
                        cmd = m_InventorApp.CommandManager.ControlDefinitions.Item("PartFlatPatternCmd")
                    Catch
                        cmd = Nothing
                    End Try

                    If cmd IsNot Nothing Then
                        Try
                            cmd.Execute()
                        Catch
                        End Try
                    End If
                End Try
            End If

            If Not smDef.HasFlatPattern Then
                Return False
            End If

            Dim flatPattern As FlatPattern = smDef.FlatPattern
            If flatPattern Is Nothing Then
                Return False
            End If

            outLengthMm = flatPattern.Length * 10.0
            outWidthMm = flatPattern.Width * 10.0
            Return True
        End Function

        Private Sub FixFlatPatternOrientation(ByVal partDoc As PartDocument,
                                              ByVal largestFace As Face,
                                              ByRef outLengthMm As Double,
                                              ByRef outWidthMm As Double)
            If partDoc Is Nothing Then
                Return
            End If

            Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
            If smDef Is Nothing OrElse Not smDef.HasFlatPattern Then
                Return
            End If

            Dim minDim As Double = Math.Min(outLengthMm, outWidthMm)
            If minDim >= 50.0 Then
                Return
            End If

            Try
                smDef.Refold()
                If largestFace IsNot Nothing Then
                    Try
                        Dim smDefObj As Object = smDef
                        smDefObj.Unfold(largestFace)
                    Catch
                        smDef.Unfold()
                    End Try
                Else
                    smDef.Unfold()
                End If
                partDoc.Update2(True)
            Catch
            End Try

            Try
                Dim flatPattern As FlatPattern = smDef.FlatPattern
                If flatPattern IsNot Nothing Then
                    outLengthMm = flatPattern.Length * 10.0
                    outWidthMm = flatPattern.Width * 10.0
                End If
            Catch
            End Try

            If Math.Min(outLengthMm, outWidthMm) >= 50.0 Then
                Return
            End If

            Try
                Dim flatPattern As FlatPattern = smDef.FlatPattern
                If flatPattern IsNot Nothing Then
                    Try
                        flatPattern.Edit()
                    Catch
                    End Try

                    Try
                        Dim orientations As FlatPatternOrientations = flatPattern.FlatPatternOrientations
                        If orientations IsNot Nothing Then
                            Dim activeOrientation As FlatPatternOrientation = orientations.ActiveFlatPatternOrientation
                            If activeOrientation IsNot Nothing Then
                                activeOrientation.FlipBaseFace = Not activeOrientation.FlipBaseFace
                            End If
                        End If
                    Catch
                    End Try

                    Try
                        flatPattern.ExitEdit()
                    Catch
                    End Try

                    Try
                        partDoc.Update2(True)
                    Catch
                    End Try

                    outLengthMm = flatPattern.Length * 10.0
                    outWidthMm = flatPattern.Width * 10.0
                End If
            Catch
            End Try
        End Sub

        Private Sub AddPlateCustomProperties(ByVal partDoc As PartDocument)
            If partDoc Is Nothing Then
                Return
            End If

            Dim userProps As PropertySet = Nothing
            Try
                userProps = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Catch
                userProps = Nothing
            End Try

            If userProps Is Nothing Then
                Return
            End If

            SetOrAddFormulaProperty(userProps, "PLATE LENGTH", "=<SHEET METAL LENGTH>")
            SetOrAddFormulaProperty(userProps, "PLATE WIDTH", "=<SHEET METAL WIDTH>")
        End Sub

        Private Sub SetOrAddFormulaProperty(ByVal userProps As PropertySet,
                                            ByVal propertyName As String,
                                            ByVal propertyValue As String)
            If userProps Is Nothing Then
                Return
            End If

            Dim propValue As Inventor.Property = Nothing
            Try
                propValue = userProps.Item(propertyName)
                propValue.Value = propertyValue
            Catch
                propValue = Nothing
            End Try

            If propValue Is Nothing Then
                Try
                    propValue = userProps.Add(propertyValue, propertyName)
                Catch
                    propValue = Nothing
                End Try
            End If

            If propValue Is Nothing Then
                Return
            End If

            Try
                propValue.DisplayString = "0"
            Catch
            End Try
        End Sub

        Private Sub ApplyDocumentSettingsForZeroDecimals(ByVal partDoc As PartDocument)
            If partDoc Is Nothing Then
                Return
            End If

            Try
                Dim params As Parameters = partDoc.ComponentDefinition.Parameters

                Try
                    params.LinearDimensionPrecision = 0
                Catch
                End Try

                Try
                    params.DimensionDisplayType = CType(kDisplayAsValue, DimensionDisplayTypeEnum)
                Catch
                End Try

                Try
                    params.DisplayParameterAsExpression = True
                Catch
                End Try
            Catch
            End Try
        End Sub

        Private Sub AddPlateParametersToAssembly(ByVal asmDoc As AssemblyDocument,
                                                 ByVal maxLengthMm As Double,
                                                 ByVal maxWidthMm As Double)
            If asmDoc Is Nothing Then
                Return
            End If

            Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            Dim userParams As UserParameters = compDef.Parameters.UserParameters

            SetOrCreateAssemblyParameter(userParams, "PLATE LENGTH", maxLengthMm)
            SetOrCreateAssemblyParameter(userParams, "PLATE WIDTH", maxWidthMm)
        End Sub

        Private Sub SetOrCreateAssemblyParameter(ByVal userParams As UserParameters,
                                                 ByVal parameterName As String,
                                                 ByVal valueMm As Double)
            Dim expression As String = valueMm.ToString("0.###", CultureInfo.InvariantCulture) & " mm"
            Dim existing As UserParameter = Nothing

            Try
                existing = userParams.Item(parameterName)
            Catch
                existing = Nothing
            End Try

            If existing Is Nothing Then
                Try
                    userParams.AddByExpression(parameterName, expression, "mm")
                Catch
                End Try
                Return
            End If

            Try
                existing.Expression = expression
            Catch
            End Try

            Try
                existing.Units = "mm"
            Catch
            End Try
        End Sub

        Private Function GetDescriptionFromIProperty(ByVal doc As Document) As String
            If doc Is Nothing Then
                Return String.Empty
            End If

            Try
                Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim descriptionProp As Inventor.Property = propSet.Item("Description")
                If descriptionProp Is Nothing OrElse descriptionProp.Value Is Nothing Then
                    Return String.Empty
                End If
                Return Convert.ToString(descriptionProp.Value).Trim()
            Catch
                Return String.Empty
            End Try
        End Function

        Private Function TryExtractThicknessMm(ByVal text As String, ByRef thicknessMm As Double) As Boolean
            thicknessMm = 0.0
            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If

            For Each pattern As Regex In ThicknessPatterns
                Dim matchValue As Match = pattern.Match(text)
                If matchValue.Success AndAlso matchValue.Groups.Count > 1 Then
                    If Double.TryParse(matchValue.Groups(1).Value,
                                       NumberStyles.Float,
                                       CultureInfo.InvariantCulture,
                                       thicknessMm) Then
                        Return thicknessMm > 0
                    End If
                End If
            Next

            Return False
        End Function

        Private Function GetThicknessFromGeometryCm(ByVal partDoc As PartDocument) As Double
            If partDoc Is Nothing Then
                Return 0.0
            End If

            Try
                Dim rangeBox As Box = partDoc.ComponentDefinition.RangeBox
                If rangeBox Is Nothing Then
                    Return 0.0
                End If

                Dim dimX As Double = Math.Abs(rangeBox.MaxPoint.X - rangeBox.MinPoint.X)
                Dim dimY As Double = Math.Abs(rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y)
                Dim dimZ As Double = Math.Abs(rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z)

                Dim thickness As Double = Math.Min(dimX, Math.Min(dimY, dimZ))
                Return thickness
            Catch
                Return 0.0
            End Try
        End Function

        Private Function GetLargestPlanarFace(ByVal partDoc As PartDocument) As Face
            If partDoc Is Nothing Then
                Return Nothing
            End If

            Try
                Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
                Dim body As SurfaceBody = partDef.SurfaceBodies.Item(1)

                Dim largestFace As Face = Nothing
                Dim largestArea As Double = 0.0

                For Each face As Face In body.Faces
                    Dim isPlanar As Boolean = False
                    Try
                        isPlanar = (CInt(face.SurfaceType) = 0)
                    Catch
                        isPlanar = False
                    End Try

                    If Not isPlanar Then
                        Continue For
                    End If

                    Dim area As Double = 0.0
                    Try
                        area = face.Evaluator.Area
                    Catch
                        area = 0.0
                    End Try

                    If area > largestArea Then
                        largestArea = area
                        largestFace = face
                    End If
                Next

                Return largestFace
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function CountTotalParts(ByVal groups As Dictionary(Of String, List(Of PlatePartInfo))) As Integer
            Dim total As Integer = 0
            For Each kvp As KeyValuePair(Of String, List(Of PlatePartInfo)) In groups
                total += kvp.Value.Count
            Next
            Return total
        End Function

        Private Function SortGroups(ByVal groups As Dictionary(Of String, List(Of PlatePartInfo))) As List(Of KeyValuePair(Of String, List(Of PlatePartInfo)))
            Dim ordered As New List(Of KeyValuePair(Of String, List(Of PlatePartInfo)))(groups)
            ordered.Sort(Function(a, b)
                             Dim aValue As Double = 0.0
                             Dim bValue As Double = 0.0
                             Double.TryParse(a.Key, NumberStyles.Float, CultureInfo.InvariantCulture, aValue)
                             Double.TryParse(b.Key, NumberStyles.Float, CultureInfo.InvariantCulture, bValue)
                             Return aValue.CompareTo(bValue)
                         End Function)
            Return ordered
        End Function

        Private Sub InitializeLog()
            m_Log.Clear()
            m_LogPath = String.Empty
        End Sub

        Private Sub LogMessage(ByVal message As String)
            Dim line As String = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) & " | " & message
            m_Log.AppendLine(line)
            AddInDiagnostics.Log("SheetMetalConverterAssemblyTool", message)
        End Sub

        Private Sub SaveLog()
            Try
                Dim docsFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                m_LogPath = Path.Combine(docsFolder, "SheetMetalConverter_Log_" & DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) & ".txt")
                File.WriteAllText(m_LogPath, m_Log.ToString())
            Catch
            End Try
        End Sub

    End Class

End Namespace
