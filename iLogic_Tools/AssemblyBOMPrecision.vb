' AssemblyBOMPrecision.vb
' Assembly-level BOM Precision Updater
' Scans assembly for plate parts and updates precision via reliable methods
' Author: Quintin de Bruin © 2026

Imports Inventor
Imports System
Imports System.Windows.Forms
Imports System.Drawing
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
    Public Sub Execute()
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

            ' Scan for plate parts
            Dim plateParts As List(Of PartDocument) = ScanForPlateParts(asmDoc)

            If plateParts.Count = 0 Then
                MessageBox.Show("No plate parts found in this assembly." & vbCrLf & _
                               "(Looking for parts with 'PL', 'PLATE', 'S355JR', or 'VRN' in part number)", _
                               "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Confirm with user
            Dim msg As String = String.Format("Found {0} plate parts to process." & vbCrLf & vbCrLf & _
                            "This will:" & vbCrLf & _
                            "1. Open each plate part" & vbCrLf & _
                            "2. Update document settings to remove decimals" & vbCrLf & _
                            "3. Save and close each part" & vbCrLf & _
                            "4. Update the assembly BOM" & vbCrLf & vbCrLf & _
                            "Estimated time: {1} seconds" & vbCrLf & vbCrLf & _
                            "Continue?", _
                            plateParts.Count, plateParts.Count * 3)

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
                m_ProgressForm.Maximum = plateParts.Count
                m_ProgressForm.Status = "Starting..."
                m_ProgressForm.Show()

                ' Process each part
                For i As Integer = 0 To plateParts.Count - 1
                    If m_Cancelled Then Exit For

                    Dim partDoc As PartDocument = plateParts(i)
                    m_ProgressForm.Current = i + 1
                    m_ProgressForm.Status = String.Format("Processing {0} ({1}/{2})...", _
                                                         partDoc.DisplayName, i + 1, plateParts.Count)
                    m_ProgressForm.Refresh()

                    ' Process the part
                    If UpdatePartPrecision(partDoc) Then
                        m_SuccessCount += 1
                    Else
                        m_FailCount += 1
                    End If

                    Application.DoEvents()
                Next

                m_ProgressForm.Status = "Finalizing..."
                m_ProgressForm.Refresh()
            End Using

            ' Finalize assembly
            FinalizeAssembly(asmDoc)

            ' Show results
            ShowResults(plateParts.Count)

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, _
                           "BOM Precision Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Scans assembly for unique plate parts
    ''' </summary>
    Private Function ScanForPlateParts(asmDoc As AssemblyDocument) As List(Of PartDocument)
        Dim plateParts As New List(Of PartDocument)()
        Dim processedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Try
            Dim occs As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences

            For Each occ As ComponentOccurrence In occs
                Try
                    Dim refDoc As Document = occ.Definition.Document

                    If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        Dim partDoc As PartDocument = CType(refDoc, PartDocument)

                        ' Get part number
                        Dim partNum As String = ""
                        Try
                            partNum = partDoc.PropertySets("Design Tracking Properties")("Part Number").Value
                        Catch
                            partNum = partDoc.DisplayName
                        End Try

                        ' Check if plate part
                        If IsPlatePart(partNum) Then
                            Dim fullPath As String = partDoc.FullFileName

                            If Not processedPaths.Contains(fullPath) Then
                                processedPaths.Add(fullPath)
                                plateParts.Add(partDoc)
                            End If
                        End If
                    End If
                Catch
                    ' Skip problematic occurrences
                End Try
            Next

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error scanning assembly: " & ex.Message)
        End Try

        Return plateParts
    End Function

    ''' <summary>
    ''' Determines if a part is a plate based on part number
    ''' </summary>
    Private Function IsPlatePart(partNum As String) As Boolean
        Dim check As String = partNum.ToUpper()
        Return check.Contains("PL ") OrElse _
               check.Contains("PLATE") OrElse _
               check.Contains("S355JR") OrElse _
               check.Contains("VRN")
    End Function

    ''' <summary>
    ''' Updates precision for a single part using the MOST RELIABLE method
    ''' </summary>
    Private Function UpdatePartPrecision(partDoc As PartDocument) As Boolean
        Try
            ' METHOD 1: Try the API method first (fastest)
            If TryAPIMethod(partDoc) Then
                Return True
            End If

            ' METHOD 2: Use the Document Settings command via API
            ' This actually opens the dialog and triggers the proper update
            Return TryDocumentSettingsMethod(partDoc)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error updating part: " & partDoc.DisplayName & " - " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' API-only method - may not always trigger BOM refresh but worth trying
    ''' </summary>
    Private Function TryAPIMethod(partDoc As PartDocument) As Boolean
        Try
            ' Update UnitsOfMeasure
            Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
            If uom IsNot Nothing Then
                ' Store original
                Dim origUnits As UnitsTypeEnum = uom.LengthUnits
                Dim origPrecision As Integer = uom.LengthDisplayPrecision

                ' Toggle units (mm -> cm -> mm)
                uom.LengthUnits = UnitsTypeEnum.kCentimeterLengthUnits
                partDoc.Update()

                uom.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                partDoc.Update()

                ' Toggle precision
                uom.LengthDisplayPrecision = 3
                partDoc.Update()

                uom.LengthDisplayPrecision = origPrecision
                partDoc.Update()
            End If

            ' Update Parameters
            Dim params As Parameters = partDoc.ComponentDefinition.Parameters
            If params IsNot Nothing Then
                Dim origPrec As Integer = params.LinearDimensionPrecision
                params.LinearDimensionPrecision = 3
                partDoc.Update()
                params.LinearDimensionPrecision = origPrec
                partDoc.Update()
            End If

            ' Force dirty flag by creating temporary parameter
            Try
                Dim dummyParam As UserParameter = params.UserParameters.AddByValue("_BOM_REFRESH_", 0, UnitsTypeEnum.kMillimeterLengthUnits)
                dummyParam.Value = 1
                params.UserParameters.RemoveByName("_BOM_REFRESH_")
                partDoc.Update()
            Catch
                ' Ignore if parameter operations fail
            End Try

            ' Save
            partDoc.Save()

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Uses the Document Settings dialog - this is the most reliable method
    ''' </summary>
    Private Function TryDocumentSettingsMethod(partDoc As PartDocument) As Boolean
        Try
            ' Save current document
            Dim originalDoc As Document = m_InventorApp.ActiveDocument

            ' Activate the part
            partDoc.Activate()
            Thread.Sleep(500)

            ' Get the Control Definition for Document Settings
            Dim cmdMgr As CommandManager = m_InventorApp.CommandManager
            Dim ctrlDefs As ControlDefinitions = cmdMgr.ControlDefinitions

            ' Execute the Document Settings command (ID varies by Inventor version)
            ' Try multiple possible command IDs
            Dim commandIds As String() = { "PartDocumentSettingsCmd", _
                                          "AppDocumentSettingsCmd", _
                                          "PartSettingsCmd" }

            Dim settingsOpened As Boolean = False

            For Each cmdId As String In commandIds
                Try
                    Dim ctrlDef As ControlDefinition = ctrlDefs.Item(cmdId)
                    If ctrlDef IsNot Nothing Then
                        ctrlDef.Execute()
                        settingsOpened = True
                        Exit For
                    End If
                Catch
                    ' Try next ID
                End Try
            Next

            If Not settingsOpened Then
                ' Fallback: Use keyboard shortcut via SendKeys
                SendKeys.SendWait("%d")  ' Alt+D
                Thread.Sleep(1000)
            End If

            ' Wait for dialog and interact
            Thread.Sleep(1500)

            ' Navigate to Units tab and toggle
            SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}")  ' 5 tabs
            Thread.Sleep(200)
            SendKeys.SendWait("{RIGHT}")  ' Right arrow to Units tab
            Thread.Sleep(500)

            ' Tab to precision field and toggle
            For i As Integer = 1 To 6
                SendKeys.SendWait("{TAB}")
                Thread.Sleep(100)
            Next

            ' Toggle down/up
            SendKeys.SendWait("{DOWN}")
            Thread.Sleep(100)
            SendKeys.SendWait("{UP}")
            Thread.Sleep(100)
            SendKeys.SendWait("{UP}")
            Thread.Sleep(100)

            ' Navigate to OK and press
            For i As Integer = 1 To 6
                SendKeys.SendWait("{TAB}")
                Thread.Sleep(100)
            Next
            SendKeys.SendWait("{ENTER}")
            Thread.Sleep(500)

            ' Save
            partDoc.Save()

            ' Return to original document if needed
            If originalDoc IsNot Nothing Then
                originalDoc.Activate()
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Finalizes the assembly after processing all parts
    ''' </summary>
    Private Sub FinalizeAssembly(asmDoc As AssemblyDocument)
        Try
            ' Activate assembly
            asmDoc.Activate()
            Thread.Sleep(500)

            ' Update assembly
            asmDoc.Update()

            ' Refresh BOM
            Try
                Dim bom As BOM = asmDoc.ComponentDefinition.BOM
                If bom IsNot Nothing Then
                    ' Toggle structured view to force refresh
                    Dim currentStructured As Boolean = bom.StructuredViewEnabled
                    bom.StructuredViewEnabled = False
                    Thread.Sleep(200)
                    bom.StructuredViewEnabled = currentStructured
                    asmDoc.Update()

                    ' Try rebuild
                    Try
                        bom.Rebuild()
                    Catch
                        ' Ignore rebuild errors
                    End Try
                End If
            Catch
                ' Ignore BOM errors
            End Try

            ' Final update
            asmDoc.Update()

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
                              "For failed parts, try manual update or check permissions.", _
                              totalCount, m_SuccessCount, m_FailCount)
            MessageBox.Show(msg, "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            msg = String.Format("Update COMPLETE!" & vbCrLf & vbCrLf & _
                              "All {0} plate parts updated successfully." & vbCrLf & vbCrLf & _
                              "BOM precision should now be updated." & vbCrLf & _
                              "If values still show decimals, try refreshing the BOM view.", _
                              totalCount)
            MessageBox.Show(msg, "BOM Precision Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

End Class

''' <summary>
''' Simple progress dialog
''' </summary>
Public Class ProgressDialog
    Inherits Form

    Private progressBar As ProgressBar
    Private lblStatus As Label
    Private btnCancel As Button

    Public Sub New()
        Me.Size = New Size(400, 150)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Progress bar
        progressBar = New ProgressBar()
        progressBar.Location = New Point(20, 20)
        progressBar.Size = New Size(340, 25)
        progressBar.Minimum = 0
        progressBar.Maximum = 100
        Me.Controls.Add(progressBar)

        ' Status label
        lblStatus = New Label()
        lblStatus.Location = New Point(20, 55)
        lblStatus.Size = New Size(340, 20)
        lblStatus.Text = "Starting..."
        Me.Controls.Add(lblStatus)

        ' Cancel button
        btnCancel = New Button()
        btnCancel.Text = "Cancel"
        btnCancel.Location = New Point(150, 85)
        btnCancel.Size = New Size(80, 25)
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
