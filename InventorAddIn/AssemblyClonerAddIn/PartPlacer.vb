' ==============================================================================
' PART PLACER MODULE - Place Assembly Parts in IDW at 1:1 Scale
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' Workflow:
' 1. User is in an IDW (or clicks button from anywhere)
' 2. Browse to select an assembly file
' 3. Scan that assembly for parts with "PL" and/or "S355JR" in Description iProperty
' 4. Place matching parts as base views in the CURRENT IDW at 1:1 scale
' ==============================================================================

Imports Inventor
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Environment

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Handles scanning assemblies and placing matching parts in IDW drawings
    ''' </summary>
    Public Class PartPlacer
        Private ReadOnly m_InventorApp As Inventor.Application
        Private m_LogPath As String
        Private m_LogBuilder As StringBuilder
        Private m_LogStartTime As DateTime

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            InitializeLogging()
        End Sub

#Region "Logging"

        Private Sub InitializeLogging()
            m_LogStartTime = DateTime.Now
            m_LogBuilder = New StringBuilder()

            Dim logFolder As String = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments),
                "InventorAutomationSuite",
                "Logs")

            If Not System.IO.Directory.Exists(logFolder) Then
                System.IO.Directory.CreateDirectory(logFolder)
            End If

            m_LogPath = System.IO.Path.Combine(logFolder, $"PartPlacer_{m_LogStartTime:yyyyMMdd_HHmmss}.log")

            LogLine(Convert.ToChar("="), 80)
            LogLine("PART PLACER - LOG FILE")
            LogLine(Convert.ToChar("="), 80)
            LogLine($"Start Time: {m_LogStartTime:yyyy-MM-dd HH:mm:ss}")
            LogLine($"Inventor Version: {m_InventorApp.SoftwareVersion.DisplayName}")
            LogLine($"Log File: {m_LogPath}")
            LogLine(Convert.ToChar("="), 80)
            LogLine("")

            FlushLog()
        End Sub

        Private Sub LogLine(ByVal message As String)
            Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss.fff")
            m_LogBuilder.AppendLine($"[{timestamp}] {message}")
        End Sub

        Private Sub LogLine(ByVal separator As Char, ByVal count As Integer)
            m_LogBuilder.AppendLine(New String(separator, count))
        End Sub

        Private Sub LogSection(ByVal title As String)
            LogLine("")
            LogLine(Convert.ToChar("-"), 60)
            LogLine($"SECTION: {title}")
            LogLine(Convert.ToChar("-"), 60)
        End Sub

        Private Sub FlushLog()
            Try
                System.IO.File.AppendAllText(m_LogPath, m_LogBuilder.ToString())
                m_LogBuilder.Clear()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"ERROR: Could not write to log: {ex.Message}")
            End Try
        End Sub

        Private Sub LogError(ByVal context As String, ByVal ex As Exception)
            LogLine("")
            LogLine(Convert.ToChar("="), 40)
            LogLine($"Context: {context}")
            LogLine($"Message: {ex.Message}")
            LogLine($"Type: {ex.GetType().FullName}")
            LogLine($"Stack Trace:")
            LogLine(ex.StackTrace)
            LogLine(Convert.ToChar("="), 40)
            FlushLog()
        End Sub

        Private Sub FinalizeLog()
            Dim duration As TimeSpan = DateTime.Now - m_LogStartTime
            LogLine("")
            LogLine(Convert.ToChar("="), 80)
            LogLine("LOG COMPLETE")
            LogLine(Convert.ToChar("="), 80)
            LogLine($"End Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            LogLine($"Duration: {duration.TotalSeconds:F2} seconds")
            LogLine(Convert.ToChar("="), 80)
            FlushLog()
        End Sub

        Public Function GetLogPath() As String
            Return m_LogPath
        End Function

#End Region

#Region "Main Entry Point"

        ''' <summary>
        ''' Main entry point - browse for assembly, scan parts, place in current IDW
        ''' </summary>
        Public Sub Execute()
            LogSection("EXECUTION STARTED")
            LogLine("User clicked 'Place Parts in IDW' button")
            FlushLog()

            Try
                ' Step 1: Get or create the target IDW document
                LogSection("STEP 1: GET/CREATE TARGET IDW")
                Dim targetIDW As DrawingDocument = GetTargetIDW()
                If targetIDW Is Nothing Then
                    LogLine("FAILED: Could not get or create target IDW")
                    FinalizeLog()
                    Return
                End If
                LogLine($"Target IDW: {targetIDW.DisplayName}")
                FlushLog()

                ' Step 2: Browse for source assembly
                LogSection("STEP 2: BROWSE FOR SOURCE ASSEMBLY")
                Dim asmPath As String = BrowseForAssembly()
                If String.IsNullOrEmpty(asmPath) Then
                    LogLine("USER CANCELLED: No assembly selected")
                    FinalizeLog()
                    Return
                End If
                LogLine($"Selected assembly: {asmPath}")
                FlushLog()

                ' Step 3: Open the assembly (hidden) and scan for matching parts
                LogSection("STEP 3: SCAN ASSEMBLY FOR MATCHING PARTS")
                Dim matchingParts As List(Of PartOccurrenceInfo) = ScanAssemblyFile(asmPath)
                LogLine($"Found {matchingParts.Count} matching parts")
                FlushLog()

                If matchingParts.Count = 0 Then
                    LogLine("No matching parts found - aborting")
                    MessageBox.Show("No parts found containing 'PL' or 'S355JR' in their Description iProperty.", 
                                   "No Matches Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    FinalizeLog()
                    Return
                End If

                ' Step 4: Place parts in the target IDW
                LogSection("STEP 4: PLACE PARTS IN IDW")
                PlacePartsInIDW(targetIDW, matchingParts)

                ' Step 5: Complete
                LogSection("EXECUTION COMPLETED SUCCESSFULLY")
                LogLine($"Total parts placed: {matchingParts.Count}")
                LogLine($"Target IDW: {targetIDW.FullFileName}")
                LogLine($"Log file: {m_LogPath}")

                MessageBox.Show($"Successfully placed {matchingParts.Count} parts in IDW.{vbCrLf}{vbCrLf}Log: {m_LogPath}", 
                               "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                LogError("Execute", ex)
                MessageBox.Show($"An error occurred: {ex.Message}{vbCrLf}{vbCrLf}See log: {m_LogPath}", 
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                FinalizeLog()
            End Try
        End Sub

#End Region

#Region "Get Target IDW"

        ''' <summary>
        ''' Get the target IDW - use active IDW if available, otherwise create new
        ''' </summary>
        Private Function GetTargetIDW() As DrawingDocument
            Try
                ' Check if there's an active document
                If m_InventorApp.ActiveDocument IsNot Nothing Then
                    Dim docType As DocumentTypeEnum = m_InventorApp.ActiveDocument.DocumentType
                    LogLine($"Active document type: {docType}")

                    If docType = DocumentTypeEnum.kDrawingDocumentObject Then
                        ' Use the active IDW
                        Dim idw As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
                        LogLine($"Using active IDW: {idw.DisplayName}")
                        Return idw
                    Else
                        LogLine("Active document is not an IDW, will create new one")
                    End If
                Else
                    LogLine("No active document, will create new IDW")
                End If

                ' Create a new IDW
                Return CreateNewIDW()

            Catch ex As Exception
                LogError("GetTargetIDW", ex)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Create a new IDW document
        ''' </summary>
        Private Function CreateNewIDW() As DrawingDocument
            Try
                Dim templatePath As String = GetDrawingTemplatePath()
                LogLine($"Template: {templatePath}")

                Dim idw As DrawingDocument
                If System.IO.File.Exists(templatePath) Then
                    idw = CType(m_InventorApp.Documents.Add(
                        DocumentTypeEnum.kDrawingDocumentObject, templatePath, True), DrawingDocument)
                Else
                    idw = CType(m_InventorApp.Documents.Add(
                        DocumentTypeEnum.kDrawingDocumentObject), DrawingDocument)
                End If

                LogLine($"Created new IDW: {idw.DisplayName}")
                Return idw

            Catch ex As Exception
                LogError("CreateNewIDW", ex)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Get the path to the drawing template
        ''' </summary>
        Private Function GetDrawingTemplatePath() As String
            Try
                Dim templateDir As String = m_InventorApp.FileOptions.TemplatesPath
                LogLine($"Templates directory: {templateDir}")

                Dim templatePath As String = System.IO.Path.Combine(templateDir, "Standard.idw")
                If System.IO.File.Exists(templatePath) Then Return templatePath

                templatePath = System.IO.Path.Combine(templateDir, "ANSI.idw")
                If System.IO.File.Exists(templatePath) Then Return templatePath

                templatePath = System.IO.Path.Combine(templateDir, "ISO.idw")
                If System.IO.File.Exists(templatePath) Then Return templatePath

                Return System.IO.Path.Combine(templateDir, "Standard.idw")
            Catch ex As Exception
                LogError("GetDrawingTemplatePath", ex)
                Return ""
            End Try
        End Function

#End Region

#Region "Browse for Assembly"

        ''' <summary>
        ''' Browse for an assembly file
        ''' </summary>
        Private Function BrowseForAssembly() As String
            Try
                Using openDialog As New OpenFileDialog()
                    openDialog.Title = "Select Assembly to Scan for Parts"
                    openDialog.Filter = "Inventor Assembly (*.iam)|*.iam"
                    openDialog.DefaultExt = "iam"
                    openDialog.CheckFileExists = True

                    ' Use default location if available
                    If m_InventorApp.ActiveDocument IsNot Nothing Then
                        openDialog.InitialDirectory = System.IO.Path.GetDirectoryName(m_InventorApp.ActiveDocument.FullFileName)
                    End If

                    LogLine("Showing OpenFileDialog...")
                    If openDialog.ShowDialog() = DialogResult.OK Then
                        LogLine($"User selected: {openDialog.FileName}")
                        Return openDialog.FileName
                    Else
                        LogLine("User cancelled the dialog")
                        Return Nothing
                    End If
                End Using

            Catch ex As Exception
                LogError("BrowseForAssembly", ex)
                Return Nothing
            End Try
        End Function

#End Region

#Region "Assembly Scanning"

        Public Class PartOccurrenceInfo
            Public Property FilePath As String
            Public Property PartNumber As String
            Public Property Description As String
            Public Property OccurrenceName As String
            Public Property MatchReason As String

            Public Overrides Function ToString() As String
                Return $"{PartNumber} - {Description} ({MatchReason})"
            End Function
        End Class

        ''' <summary>
        ''' Open assembly file and scan for matching parts
        ''' </summary>
        Private Function ScanAssemblyFile(ByVal asmPath As String) As List(Of PartOccurrenceInfo)
            Dim results As New List(Of PartOccurrenceInfo)()
            Dim asmDoc As AssemblyDocument = Nothing

            Try
                LogLine($"Opening assembly: {asmPath}")
                asmDoc = CType(m_InventorApp.Documents.Open(asmPath, False), AssemblyDocument)
                LogLine($"Assembly opened: {asmDoc.DisplayName}")

                Dim processedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim compDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

                LogLine($"Top-level occurrences: {compDef.Occurrences.Count}")

                ' Process all occurrences recursively
                ProcessOccurrences(compDef.Occurrences, results, processedFiles, 0)

                LogLine("")
                LogLine($"SCAN COMPLETE: Found {results.Count} unique matching parts")
                For i As Integer = 0 To results.Count - 1
                    LogLine($"  [{i + 1}] {results(i).PartNumber} - {results(i).Description} - {results(i).MatchReason}")
                Next

            Catch ex As Exception
                LogError("ScanAssemblyFile", ex)
            Finally
                ' Close the assembly document
                If asmDoc IsNot Nothing Then
                    Try
                        asmDoc.Close(False)
                        LogLine("Assembly document closed")
                    Catch
                        ' Ignore close errors
                    End Try
                End If
            End Try

            Return results
        End Function

        ''' <summary>
        ''' Recursively process component occurrences
        ''' </summary>
        Private Sub ProcessOccurrences(ByVal occurrences As ComponentOccurrences, 
                                        ByVal results As List(Of PartOccurrenceInfo), 
                                        ByVal processedFiles As HashSet(Of String),
                                        ByVal depth As Integer)
            Dim indent As String = New String(Convert.ToChar(" "), depth * 2)

            For Each occ As ComponentOccurrence In occurrences
                Try
                    If occ.Definition Is Nothing Then
                        LogLine($"{indent}Skipping occurrence '{occ.Name}' - no definition")
                        Continue For
                    End If

                    If occ.DefinitionType = ObjectTypeEnum.kPartComponentDefinitionObject Then
                        Dim partDef As PartComponentDefinition = CType(occ.Definition, PartComponentDefinition)
                        Dim doc As PartDocument = CType(partDef.Document, PartDocument)
                        Dim filePath As String = doc.FullFileName

                        LogLine($"{indent}Processing part: {System.IO.Path.GetFileName(filePath)}")

                        If processedFiles.Contains(filePath) Then
                            LogLine($"{indent}  -> Already processed, skipping")
                            Continue For
                        End If

                        Dim partInfo As PartOccurrenceInfo = AnalyzePart(doc, occ)

                        If partInfo IsNot Nothing Then
                            LogLine($"{indent}  -> MATCH FOUND: {partInfo.MatchReason}")
                            LogLine($"{indent}     Description: {partInfo.Description}")
                            results.Add(partInfo)
                            processedFiles.Add(filePath)
                        Else
                            LogLine($"{indent}  -> No match")
                        End If

                    ElseIf occ.DefinitionType = ObjectTypeEnum.kAssemblyComponentDefinitionObject Then
                        Dim subAsmDef As AssemblyComponentDefinition = CType(occ.Definition, AssemblyComponentDefinition)
                        LogLine($"{indent}Entering sub-assembly: {occ.Name}")
                        ProcessOccurrences(subAsmDef.Occurrences, results, processedFiles, depth + 1)
                    End If

                Catch ex As Exception
                    LogError($"ProcessOccurrences - {occ.Name}", ex)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Analyze a part to check if it matches the criteria
        ''' ONLY looks at the Description iProperty for PL and S355JR
        ''' </summary>
        Private Function AnalyzePart(ByVal partDoc As PartDocument, ByVal occurrence As ComponentOccurrence) As PartOccurrenceInfo
            Try
                Dim info As New PartOccurrenceInfo()
                info.FilePath = partDoc.FullFileName
                info.OccurrenceName = occurrence.Name

                ' Get iProperties - matching VBScript logic with error handling
                Dim partNumber As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullFileName)
                Dim description As String = ""

                Try
                    ' Use .Item() method like VBScript does
                    Dim propSet As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
                    
                    ' Part Number (Stock Number)
                    Try
                        partNumber = propSet.Item("Stock Number").Value.ToString()
                    Catch
                        ' Keep default filename
                    End Try

                    ' Description - THIS IS THE KEY FIELD WE CHECK
                    Try
                        description = propSet.Item("Description").Value.ToString()
                        description = description.Trim()
                    Catch
                        description = ""
                    End Try

                Catch ex As Exception
                    LogLine($"    Warning: Could not read iProperties: {ex.Message}")
                End Try

                info.PartNumber = partNumber
                info.Description = description

                LogLine($"    Part Number: {partNumber}")
                LogLine($"    Description: {description}")

                ' IMPORTANT: Only check the Description field for PL and S355JR
                ' Match VBScript logic: Left(desc, 2) = "PL" or contains "S355JR"
                Dim descUpper As String = If(description, "").ToUpper().Trim()

                Dim startsWithPL As Boolean = descUpper.StartsWith("PL")
                Dim containsS355JR As Boolean = descUpper.Contains("S355JR")

                LogLine($"    startsWithPL: {startsWithPL}, containsS355JR: {containsS355JR}")

                ' Determine match reason based on Description ONLY
                If startsWithPL AndAlso containsS355JR Then
                    info.MatchReason = "Description starts with PL and contains S355JR"
                    Return info
                ElseIf containsS355JR Then
                    info.MatchReason = "Description contains S355JR"
                    Return info
                ElseIf startsWithPL Then
                    info.MatchReason = "Description starts with PL"
                    Return info
                End If

                Return Nothing

            Catch ex As Exception
                LogError("AnalyzePart", ex)
                Return Nothing
            End Try
        End Function

        ' Note: IsPLMaterial function removed - now using StartsWith("PL") logic to match VBScript

#End Region

#Region "Place Parts in IDW"

        ''' <summary>
        ''' Place all matching parts as base views in the target IDW
        ''' </summary>
        Private Sub PlacePartsInIDW(ByVal idwDoc As DrawingDocument, ByVal parts As List(Of PartOccurrenceInfo))
            LogLine("")
            LogLine("Placing parts in IDW...")

            Dim sheet As Sheet = idwDoc.ActiveSheet
            LogLine($"Active sheet: {sheet.Name}")
            LogLine($"Sheet size: {sheet.Width:F2} x {sheet.Height:F2} inches")

            ' Calculate grid layout (dimensions in cm)
            Dim viewWidth As Double = 8
            Dim viewHeight As Double = 6
            Dim margin As Double = 1
            Dim spacing As Double = 1

            Dim sheetWidth As Double = sheet.Width * 2.54
            Dim sheetHeight As Double = sheet.Height * 2.54

            Dim availableWidth As Double = sheetWidth - (2 * margin)
            Dim cols As Integer = Math.Max(1, CInt(Math.Floor((availableWidth + spacing) / (viewWidth + spacing))))

            LogLine($"Grid: {cols} columns")

            Dim startX As Double = margin
            Dim startY As Double = sheetHeight - margin - viewHeight

            For i As Integer = 0 To parts.Count - 1
                Dim part As PartOccurrenceInfo = parts(i)
                Dim col As Integer = i Mod cols
                Dim row As Integer = i \ cols

                Dim xPos As Double = startX + (col * (viewWidth + spacing))
                Dim yPos As Double = startY - (row * (viewHeight + spacing))

                LogLine("")
                LogLine($"Placing [{i + 1}/{parts.Count}]: {part.PartNumber}")
                LogLine($"  Description: {part.Description}")
                LogLine($"  Position: ({xPos:F2}cm, {yPos:F2}cm)")

                Try
                    PlaceBaseView(sheet, part, xPos, yPos)
                    LogLine($"  -> Success")
                Catch ex As Exception
                    LogError($"PlacePartsInIDW - {part.PartNumber}", ex)
                    LogLine($"  -> FAILED: {ex.Message}")
                End Try

                FlushLog()
            Next

            LogLine("")
            LogLine($"Finished placing {parts.Count} parts")
        End Sub

        ''' <summary>
        ''' Place a single base view for a part
        ''' </summary>
        Private Sub PlaceBaseView(ByVal sheet As Sheet, ByVal part As PartOccurrenceInfo, 
                                   ByVal xPos As Double, ByVal yPos As Double)
            Try
                ' Open the part document (hidden)
                LogLine($"  Opening: {System.IO.Path.GetFileName(part.FilePath)}")
                Dim partDoc As PartDocument = CType(m_InventorApp.Documents.Open(part.FilePath, False), PartDocument)

                Dim drawingViews As DrawingViews = sheet.DrawingViews
                Dim scale As Double = 1.0

                Dim xInches As Double = xPos / 2.54
                Dim yInches As Double = yPos / 2.54

                Dim baseView As DrawingView = drawingViews.AddBaseView(
                    partDoc, 
                    m_InventorApp.TransientGeometry.CreatePoint2d(xInches, yInches),
                    scale,
                    ViewOrientationTypeEnum.kFrontViewOrientation,
                    DrawingViewStyleEnum.kHiddenLineDrawingViewStyle)

                LogLine($"  View created: {baseView.Name}")
                LogLine($"  Scale: 1:1")

            Catch ex As Exception
                Throw New Exception($"Failed to place base view for {part.PartNumber}: {ex.Message}", ex)
            End Try
        End Sub

#End Region

    End Class

End Namespace
