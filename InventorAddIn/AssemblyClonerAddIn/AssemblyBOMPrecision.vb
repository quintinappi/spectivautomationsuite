' AssemblyBOMPrecision.vb - FIXED VERSION
' Assembly-level BOM Precision Updater
' Uses Custom Property manipulation to reliably trigger document updates
' Author: Quintin de Bruin © 2026

Imports Inventor
Imports System
Imports System.Windows.Forms
Imports System.Threading

Public Class AssemblyBOMPrecision
    Private m_InventorApp As Inventor.Application
    Private m_ProgressForm As ProgressDialog
    Private m_Cancelled As Boolean = False
    Private m_SuccessCount As Integer = 0
    Private m_FailCount As Integer = 0

    Public Sub New(ByVal inventorApp As Inventor.Application)
        m_InventorApp = inventorApp
    End Sub

    ''' <summary>
    ''' Main entry point - shows dialog and processes assembly
    ''' </summary>
    Public Sub Execute(Optional ByVal explicitPartPaths As IEnumerable(Of String) = Nothing)
        Try
            ' Check for active assembly
            If m_InventorApp.ActiveDocument Is Nothing Then
                MessageBox.Show("No active document!" & vbCrLf & "Please open an assembly.", _
                               "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show("Please open an ASSEMBLY document." & vbCrLf & _
                               "Current document: " & m_InventorApp.ActiveDocument.DisplayName, _
                               "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)

            Dim targetPartPaths As List(Of String) = Nothing
            If explicitPartPaths IsNot Nothing Then
                targetPartPaths = NormalizePartPaths(explicitPartPaths)
            Else
                Dim plateParts As List(Of PartDocument) = ScanForPlateParts(asmDoc)
                targetPartPaths = New List(Of String)()
                For Each partDoc As PartDocument In plateParts
                    Try
                        If partDoc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(partDoc.FullFileName) Then
                            targetPartPaths.Add(partDoc.FullFileName)
                        End If
                    Catch
                    End Try
                Next
            End If

            If targetPartPaths.Count = 0 Then
                ' Get diagnostic info
                Dim diagInfo As String = GetDiagnosticInfo(asmDoc)
                
                MessageBox.Show("No plate parts found in this assembly." & vbCrLf & vbCrLf & _
                               "Looking for: PL, PLATE, S355JR, VRN, END PLATE, GUSSET" & vbCrLf & vbCrLf & _
                               "DIAGNOSTIC INFO:" & vbCrLf & diagInfo, _
                               "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Confirm with user
            Dim msg As String = String.Format("Found {0} plate parts to process." & vbCrLf & vbCrLf & _
                            "⚠️  WARNING: This will use Alt+D and automate keyboard." & vbCrLf & _
                            "Do not touch your computer during processing." & vbCrLf & vbCrLf & _
                            "Steps:" & vbCrLf & _
                            "1. Open each plate part" & vbCrLf & _
                            "2. Press Alt+D (Document Settings)" & vbCrLf & _
                            "3. Navigate to Units tab" & vbCrLf & _
                            "4. Toggle precision down/up" & vbCrLf & _
                            "5. Press OK" & vbCrLf & _
                            "6. Save and close part" & vbCrLf & vbCrLf & _
                            "Estimated time: {1} seconds" & vbCrLf & vbCrLf & _
                            "Continue?", _
                            targetPartPaths.Count, targetPartPaths.Count * 5)

            If MessageBox.Show(msg, "BOM Precision Update", _
                              MessageBoxButtons.OKCancel, MessageBoxIcon.Question) <> DialogResult.OK Then
                Return
            End If

            ' Show progress dialog
            m_Cancelled = False
            m_SuccessCount = 0
            m_FailCount = 0

            Using m_ProgressForm = New ProgressDialog()
                m_ProgressForm.Text = "Updating BOM Precision..."
                m_ProgressForm.Maximum = targetPartPaths.Count
                m_ProgressForm.Status = "Starting..."
                m_ProgressForm.Show()

                ' Process each part
                For i As Integer = 0 To targetPartPaths.Count - 1
                    If m_Cancelled Then Exit For

                    Dim partPath As String = targetPartPaths(i)
                    Dim partLabel As String = System.IO.Path.GetFileName(partPath)
                    If String.IsNullOrWhiteSpace(partLabel) Then
                        partLabel = partPath
                    End If

                    m_ProgressForm.Current = i + 1
                    m_ProgressForm.Status = String.Format("Processing {0} ({1}/{2})...", _
                                                         partLabel, i + 1, targetPartPaths.Count)
                    m_ProgressForm.Refresh()

                    ' Process the part using RELIABLE method
                    If UpdatePartPrecisionReliableByPath(partPath) Then
                        m_SuccessCount += 1
                    Else
                        m_FailCount += 1
                    End If

                    System.Windows.Forms.Application.DoEvents()
                    Thread.Sleep(2000)  ' Safety delay between parts
                Next

                m_ProgressForm.Status = "Finalizing..."
                m_ProgressForm.Refresh()
            End Using

            ' Finalize assembly
            FinalizeAssembly(asmDoc)

            ' Show results
            ShowResults(targetPartPaths.Count)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, _
                           "BOM Precision Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Scans assembly for unique plate parts using BOM view
    ''' </summary>
    Private Function ScanForPlateParts(asmDoc As AssemblyDocument) As List(Of PartDocument)
        Dim plateParts As New List(Of PartDocument)()
        Dim processedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Try
            ' Enable BOM structured view
            Dim bom As BOM = asmDoc.ComponentDefinition.BOM
            bom.StructuredViewEnabled = True
            bom.StructuredViewFirstLevelOnly = False
            
            Dim bomView As BOMView = bom.BOMViews("Structured")
            
            For Each row As BOMRow In bomView.BOMRows
                Try
                    ' Get the component definition
                    If row.ComponentDefinitions.Count > 0 Then
                        Dim compDef As ComponentDefinition = row.ComponentDefinitions(1)
                        Dim doc As Document = compDef.Document
                        
                        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                            Dim partDoc As PartDocument = CType(doc, PartDocument)
                            Dim fullPath As String = partDoc.FullFileName
                            
                            ' Skip if already processed
                            If processedPaths.Contains(fullPath) Then Continue For
                            
                            ' Check if plate part
                            If IsPlatePart(partDoc) Then
                                processedPaths.Add(fullPath)
                                plateParts.Add(partDoc)
                            End If
                        End If
                    End If
                Catch
                    ' Skip problematic rows
                End Try
            Next
            
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error scanning BOM: " & ex.Message)
        End Try

        Return plateParts
    End Function

    ''' <summary>
    ''' Gets diagnostic information about what was found in the assembly
    ''' </summary>
    Private Function GetDiagnosticInfo(asmDoc As AssemblyDocument) As String
        Dim info As New System.Text.StringBuilder()
        Dim partCount As Integer = 0
        
        Try
            Dim bom As BOM = asmDoc.ComponentDefinition.BOM
            bom.StructuredViewEnabled = True
            Dim bomView As BOMView = bom.BOMViews("Structured")
            
            info.AppendLine("BOM Rows: " & bomView.BOMRows.Count)
            info.AppendLine("")
            info.AppendLine("First 10 parts found:")
            
            For i As Integer = 0 To Math.Min(9, bomView.BOMRows.Count - 1)
                Try
                    Dim row As BOMRow = bomView.BOMRows(i)
                    If row.ComponentDefinitions.Count > 0 Then
                        Dim doc As Document = row.ComponentDefinitions(1).Document
                        
                        Dim partNum As String = ""
                        Dim desc As String = ""
                        Try
                            partNum = doc.PropertySets("Design Tracking Properties")("Part Number").Value.ToString()
                        Catch
                        End Try
                        Try
                            desc = doc.PropertySets("Design Tracking Properties")("Description").Value.ToString()
                        Catch
                        End Try
                        
                        info.AppendLine("  " & (i + 1) & ". " & partNum)
                        If Not String.IsNullOrEmpty(desc) Then
                            info.AppendLine("      Desc: " & desc)
                        End If
                        partCount += 1
                    End If
                Catch
                End Try
            Next
            
            If partCount = 0 Then
                info.AppendLine("  No parts could be read from BOM")
            End If
            
        Catch ex As Exception
            info.AppendLine("Error reading BOM: " & ex.Message)
        End Try
        
        Return info.ToString()
    End Function

    ''' <summary>
    ''' Determines if a part is a plate based on part number or description
    ''' </summary>
    Private Function IsPlatePart(partDoc As PartDocument) As Boolean
        Try
            Dim partNum As String = ""
            Dim description As String = ""
            
            ' Try to get Part Number
            Try
                partNum = partDoc.PropertySets("Design Tracking Properties")("Part Number").Value.ToString().ToUpper()
            Catch
            End Try
            
            ' Try to get Description
            Try
                description = partDoc.PropertySets("Design Tracking Properties")("Description").Value.ToString().ToUpper()
            Catch
            End Try
            
            ' Combine for checking
            Dim check As String = partNum & " " & description
            
            ' Check for plate indicators (more flexible matching)
            If check.Contains("PL") Then Return True
            If check.Contains("PLATE") Then Return True
            If check.Contains("S355JR") Then Return True
            If check.Contains("VRN") Then Return True
            If check.Contains("END PLATE") Then Return True
            If check.Contains("GUSSET") Then Return True
            
            Return False
            
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' RELIABLE METHOD: Uses Document Settings UI to toggle precision
    ''' This is the ONLY method that reliably triggers BOM refresh
    ''' </summary>
    Private Function UpdatePartPrecisionReliableByPath(ByVal fullPath As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(fullPath) Then
                Return False
            End If
            
            ' Re-open visibly
            Dim openedDoc As Document = m_InventorApp.Documents.Open(fullPath, True)
            If openedDoc Is Nothing OrElse openedDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                Return False
            End If
            
            Dim activePart As PartDocument = CType(openedDoc, PartDocument)
            activePart.Activate()
            Thread.Sleep(500)
            
            ' BRING INVENTOR TO FRONT
            m_InventorApp.Visible = True
            Thread.Sleep(200)
            
            ' Get window handle and activate
            Try
                Dim hwnd As Integer = m_InventorApp.MainFrameHWND
                If hwnd <> 0 Then
                    SetForegroundWindow(hwnd)
                End If
            Catch
            End Try
            Thread.Sleep(300)

            ' BRING INVENTOR TO FRONT
            m_InventorApp.Visible = True
            Thread.Sleep(200)
            
            ' Get window handle and activate
            Try
                Dim hwnd As Integer = m_InventorApp.MainFrameHWND
                If hwnd <> 0 Then
                    SetForegroundWindow(hwnd)
                End If
            Catch
            End Try
            Thread.Sleep(300)

            If Not OpenDocumentSettingsDialog() Then
                Return False
            End If
            
            ' NAVIGATE TO UNITS TAB
            ' Tab 5 times to get to tab control, then Right arrow to Units
            For i As Integer = 1 To 5
                System.Windows.Forms.SendKeys.SendWait("{TAB}")
                Thread.Sleep(150)
            Next
            System.Windows.Forms.SendKeys.SendWait("{RIGHT}")
            Thread.Sleep(400)
            
            ' NAVIGATE TO LENGTH DISPLAY PRECISION DROPDOWN
            ' Tab multiple times to reach the precision field
            For i As Integer = 1 To 6
                System.Windows.Forms.SendKeys.SendWait("{TAB}")
                Thread.Sleep(150)
            Next
            
            ' TOGGLE PRECISION: Down then Up (forces the change event)
            System.Windows.Forms.SendKeys.SendWait("{DOWN}")  ' Change to different value
            Thread.Sleep(300)
            System.Windows.Forms.SendKeys.SendWait("{UP}")    ' Back to original (or close to it)
            Thread.Sleep(300)
            System.Windows.Forms.SendKeys.SendWait("{UP}")    ' Ensure we're at 0
            Thread.Sleep(300)

            ' Press Apply first to force UI event path, then confirm/close.
            System.Windows.Forms.SendKeys.SendWait("%a")
            Thread.Sleep(400)
            
            ' PRESS OK
            System.Windows.Forms.SendKeys.SendWait("{ENTER}")
            Thread.Sleep(800)
            
            ' SAVE (use Save2 to force full save)
            Try
                activePart.Save2(True)
            Catch
                activePart.Save()
            End Try
            
            ' CLOSE
            activePart.Close(True)
            
            Return True
            
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in UI method: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function NormalizePartPaths(ByVal partPaths As IEnumerable(Of String)) As List(Of String)
        Dim result As New List(Of String)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If partPaths Is Nothing Then
            Return result
        End If

        For Each rawPath As String In partPaths
            Dim partPath As String = If(rawPath, String.Empty).Trim()
            If String.IsNullOrWhiteSpace(partPath) Then
                Continue For
            End If

            If Not partPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            If seen.Add(partPath) Then
                result.Add(partPath)
            End If
        Next

        Return result
    End Function

    Private Function OpenDocumentSettingsDialog() As Boolean
        Try
            Dim cmdMgr As CommandManager = m_InventorApp.CommandManager
            Dim ctrlDefs As ControlDefinitions = cmdMgr.ControlDefinitions
            Dim commandIds As String() = {"PartDocumentSettingsCmd", "AppDocumentSettingsCmd", "PartSettingsCmd"}

            For Each cmdId As String In commandIds
                Try
                    Dim ctrlDef As ControlDefinition = ctrlDefs.Item(cmdId)
                    If ctrlDef Is Nothing Then
                        Continue For
                    End If

                    Dim buttonDef As ButtonDefinition = TryCast(ctrlDef, ButtonDefinition)
                    If buttonDef IsNot Nothing Then
                        buttonDef.Execute2(False)
                    Else
                        ctrlDef.Execute()
                    End If

                    Thread.Sleep(1400)
                    Return True
                Catch
                End Try
            Next

            System.Windows.Forms.SendKeys.SendWait("%d")
            Thread.Sleep(1500)
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ' Windows API to bring window to front
    Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer



    ''' <summary>
    ''' Finalizes the assembly after processing all parts
    ''' </summary>
    Private Sub FinalizeAssembly(asmDoc As AssemblyDocument)
        Try
            ' Activate assembly
            asmDoc.Activate()
            Thread.Sleep(500)

            ' Rebuild assembly (forces cache invalidation)
            Try
                asmDoc.Rebuild2(True)
            Catch
            End Try

            ' Update assembly
            Try
                asmDoc.Update2(True)
            Catch
                asmDoc.Update()
            End Try

            ' Refresh BOM - CRITICAL METHODS
            Try
                Dim bom As BOM = asmDoc.ComponentDefinition.BOM
                If bom IsNot Nothing Then
                    ' Method 1: Toggle structured view
                    Dim currentStructured As Boolean = bom.StructuredViewEnabled
                    bom.StructuredViewEnabled = False
                    Thread.Sleep(200)
                    bom.StructuredViewEnabled = True
                    asmDoc.Update()

                    ' Method 2: Access BOM rows (forces cache rebuild)
                    Try
                        Dim structuredView As BOMView = bom.BOMViews("Structured")
                        Dim rowCount As Integer = structuredView.BOMRows.Count
                        System.Diagnostics.Debug.WriteLine("BOM rows: " & rowCount)
                    Catch
                    End Try

                    ' Method 3: Rebuild BOM
                    Try
                        bom.Rebuild()
                    Catch
                    End Try
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("BOM refresh error: " & ex.Message)
            End Try

            ' Final update and save
            asmDoc.Update()
            Try
                asmDoc.Save2(True)
            Catch
                asmDoc.Save()
            End Try

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error finalizing assembly: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Shows results dialog
    ''' </summary>
    Private Sub ShowResults(totalCount As Integer)
        Dim msg As String

        If m_Cancelled Then
            msg = String.Format("Update CANCELLED by user." & vbCrLf & vbCrLf & _
                              "Processed: {0}/{1}" & vbCrLf & _
                              "Successful: {2}" & vbCrLf & _
                              "Failed: {3}", _
                              m_SuccessCount + m_FailCount, totalCount, m_SuccessCount, m_FailCount)
            MessageBox.Show(msg, "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        ElseIf m_FailCount > 0 Then
            msg = String.Format("Update COMPLETE with some failures." & vbCrLf & vbCrLf & _
                              "Total: {0}" & vbCrLf & _
                              "Successful: {1}" & vbCrLf & _
                              "Failed: {2}" & vbCrLf & vbCrLf & _
                              "For failed parts, try manual update.", _
                              totalCount, m_SuccessCount, m_FailCount)
            MessageBox.Show(msg, "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            msg = String.Format("Update COMPLETE!" & vbCrLf & vbCrLf & _
                              "All {0} plate parts processed." & vbCrLf & vbCrLf & _
                              "Document Settings was opened for each part" & vbCrLf & _
                              "and Length Display Precision was toggled." & vbCrLf & vbCrLf & _
                              "⚠️  IMPORTANT:" & vbCrLf & _
                              "The BOM display depends on Inventor's internal cache." & vbCrLf & _
                              "If decimals still appear:" & vbCrLf & _
                              "1. Right-click BOM → Refresh" & vbCrLf & _
                              "2. Save and close assembly, then reopen" & vbCrLf & _
                              "3. Or use Manage → Rebuild All", _
                              totalCount)
            MessageBox.Show(msg, "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class

''' <summary>
''' Simple progress dialog
''' </summary>
Public Class ProgressDialog
    Inherits System.Windows.Forms.Form

    Private progressBar As System.Windows.Forms.ProgressBar
    Private lblStatus As System.Windows.Forms.Label
    Private btnCancel As System.Windows.Forms.Button

    Public Sub New()
        Me.Size = New System.Drawing.Size(400, 150)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen

        ' Progress bar
        progressBar = New System.Windows.Forms.ProgressBar()
        progressBar.Location = New System.Drawing.Point(20, 20)
        progressBar.Size = New System.Drawing.Size(340, 25)
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        Me.Controls.Add(progressBar)

        ' Status label
        lblStatus = New System.Windows.Forms.Label()
        lblStatus.Location = New System.Drawing.Point(20, 55)
        lblStatus.Size = New System.Drawing.Size(340, 20)
        lblStatus.Text = "Starting..."
        Me.Controls.Add(lblStatus)

        ' Cancel button
        btnCancel = New System.Windows.Forms.Button()
        btnCancel.Text = "Cancel"
        btnCancel.Location = New System.Drawing.Point(150, 85)
        btnCancel.Size = New System.Drawing.Size(80, 25)
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click
        Me.Controls.Add(btnCancel)
    End Sub

    Public Property Maximum As Integer
        Get
            Return progressBar.Maximum
        End Get
        Set(value As Integer)
            progressBar.Maximum = value
        End Set
    End Property

    Public Property Current As Integer
        Get
            Return progressBar.Value
        End Get
        Set(value As Integer)
            progressBar.Value = value
        End Set
    End Property

    Public Property Status As String
        Get
            Return lblStatus.Text
        End Get
        Set(value As String)
            lblStatus.Text = value
        End Set
    End Property

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class
