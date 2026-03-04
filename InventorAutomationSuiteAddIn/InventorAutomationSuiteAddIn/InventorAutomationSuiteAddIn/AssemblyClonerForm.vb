' ============================================================================
' INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER FORM
' ============================================================================
' Description: Assembly Cloner tool interface
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-20
' ============================================================================

Imports System
Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor

Namespace InventorAutomationSuiteAddIn

    ''' <summary>
    ''' Form for the Assembly Cloner tool
    ''' </summary>
    Public Class AssemblyClonerForm
        Inherits Form

        ' Private members
        Private m_invApp As InventorApplication
        Private m_regManager As RegistryManager

        ' UI Controls
        Private txtPrefix As TextBox
        Private numCloneCount As NumericUpDown
        Private btnScanRegistry As Button
        Private btnClone As Button
        Private btnCancel As Button
        Private grpRegistryStatus As GroupBox
        Private lblPL As Label
        Private lblB As Label
        Private lblCH As Label
        Private lblA As Label
        Private lblFL As Label
        Private ProgressBar As ProgressBar
        Private lblProgress As Label

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(invApp As InventorApplication, regManager As RegistryManager)
            Try
                m_invApp = invApp
                m_regManager = regManager

                ' Initialize form components
                InitializeComponent()

                LogMessage("Assembly Cloner form loaded")

            Catch ex As Exception
                MessageBox.Show("Error initializing Assembly Cloner form: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Initialize form components
        ''' </summary>
        Private Sub InitializeComponent()
            Try
                ' Set form properties
                Me.Text = "Assembly Cloner - Inventor Automation Suite"
                Me.Size = New Size(700, 600)
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.MinimizeBox = False
                Me.MaximizeBox = False
                Me.FormBorderStyle = FormBorderStyle.FixedDialog
                Me.BackColor = Color.White

                ' Create title
                Dim titleLabel As New Label()
                titleLabel.Text = "Assembly Cloner"
                titleLabel.Font = New Font("Segoe UI", 18, FontStyle.Bold)
                titleLabel.ForeColor = Color.FromArgb(102, 126, 234)
                titleLabel.Location = New Point(20, 20)
                titleLabel.Size = New Size(400, 30)
                Me.Controls.Add(titleLabel)

                Dim subtitleLabel As New Label()
                subtitleLabel.Text = "Clone assemblies with automatic numbering continuation"
                subtitleLabel.Font = New Font("Segoe UI", 10)
                subtitleLabel.ForeColor = Color.Gray
                subtitleLabel.Location = New Point(20, 55)
                subtitleLabel.Size = New Size(500, 20)
                Me.Controls.Add(subtitleLabel)

                ' Create prefix input
                Dim lblPrefix As New Label()
                lblPrefix.Text = "Prefix:"
                lblPrefix.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                lblPrefix.Location = New Point(20, 100)
                lblPrefix.Size = New Size(100, 20)
                Me.Controls.Add(lblPrefix)

                txtPrefix = New TextBox()
                txtPrefix.Text = "NCRH01-000-"
                txtPrefix.Font = New Font("Segoe UI", 10)
                txtPrefix.Location = New Point(120, 98)
                txtPrefix.Size = New Size(300, 25)
                Me.Controls.Add(txtPrefix)

                Dim prefixHelp As New Label()
                prefixHelp.Text = "e.g., NCRH01-000-, PLANT1-000-"
                prefixHelp.Font = New Font("Segoe UI", 8)
                prefixHelp.ForeColor = Color.Gray
                prefixHelp.Location = New Point(430, 103)
                prefixHelp.Size = New Size(250, 15)
                Me.Controls.Add(prefixHelp)

                ' Create clone count input
                Dim lblCount As New Label()
                lblCount.Text = "Clone Count:"
                lblCount.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                lblCount.Location = New Point(20, 140)
                lblCount.Size = New Size(100, 20)
                Me.Controls.Add(lblCount)

                numCloneCount = New NumericUpDown()
                numCloneCount.Minimum = 1
                numCloneCount.Maximum = 10
                numCloneCount.Value = 1
                numCloneCount.Font = New Font("Segoe UI", 10)
                numCloneCount.Location = New Point(120, 138)
                numCloneCount.Size = New Size(100, 25)
                Me.Controls.Add(numCloneCount)

                ' Create registry status group
                grpRegistryStatus = New GroupBox()
                grpRegistryStatus.Text = " Current Registry Status "
                grpRegistryStatus.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                grpRegistryStatus.Location = New Point(20, 190)
                grpRegistryStatus.Size = New Size(640, 150)
                grpRegistryStatus.BackColor = Color.FromArgb(248, 249, 250)
                Me.Controls.Add(grpRegistryStatus)

                ' Create registry labels
                Dim y As Integer = 30
                CreateRegistryLabel("PL (Platework):", "Not detected", 30, grpRegistryStatus, lblPL)
                CreateRegistryLabel("B (Beams/Columns):", "Not detected", 60, grpRegistryStatus, lblB)
                CreateRegistryLabel("CH (Channels):", "Not detected", 90, grpRegistryStatus, lblCH)
                CreateRegistryLabel("A (Angles):", "Not detected", 120, grpRegistryStatus, lblA)
                CreateRegistryLabel("FL (Flatbar):", "Not detected", 150, grpRegistryStatus, lblFL)

                ' Create scan registry button
                btnScanRegistry = New Button()
                btnScanRegistry.Text = "Scan Registry"
                btnScanRegistry.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                btnScanRegistry.BackColor = Color.FromArgb(240, 240, 240)
                btnScanRegistry.FlatStyle = FlatStyle.Flat
                btnScanRegistry.FlatAppearance.BorderSize = 0
                btnScanRegistry.Cursor = Cursors.Hand
                btnScanRegistry.Location = New Point(20, 360)
                btnScanRegistry.Size = New Size(150, 40)
                AddHandler btnScanRegistry.Click, AddressOf btnScanRegistry_Click
                Me.Controls.Add(btnScanRegistry)

                ' Create progress bar
                ProgressBar = New ProgressBar()
                ProgressBar.Location = New Point(20, 420)
                ProgressBar.Size = New Size(640, 25)
                ProgressBar.Style = ProgressBarStyle.Continuous
                ProgressBar.Visible = False
                Me.Controls.Add(ProgressBar)

                lblProgress = New Label()
                lblProgress.Text = "Ready"
                lblProgress.Font = New Font("Segoe UI", 9)
                lblProgress.Location = New Point(20, 450)
                lblProgress.Size = New Size(640, 20)
                lblProgress.Visible = False
                Me.Controls.Add(lblProgress)

                ' Create clone button
                btnClone = New Button()
                btnClone.Text = "Clone Assembly"
                btnClone.Font = New Font("Segoe UI", 12, FontStyle.Bold)
                btnClone.BackColor = Color.FromArgb(102, 126, 234)
                btnClone.ForeColor = Color.White
                btnClone.FlatStyle = FlatStyle.Flat
                btnClone.FlatAppearance.BorderSize = 0
                btnClone.Cursor = Cursors.Hand
                btnClone.Location = New Point(460, 490)
                btnClone.Size = New Size(200, 50)
                AddHandler btnClone.Click, AddressOf btnClone_Click
                AddHandler btnClone.MouseEnter, Sub(sender, e)
                                                  btnClone.BackColor = Color.FromArgb(118, 75, 162)
                                              End Sub
                AddHandler btnClone.MouseLeave, Sub(sender, e)
                                                  btnClone.BackColor = Color.FromArgb(102, 126, 234)
                                              End Sub
                Me.Controls.Add(btnClone)

                ' Create cancel button
                btnCancel = New Button()
                btnCancel.Text = "Cancel"
                btnCancel.Font = New Font("Segoe UI", 10)
                btnCancel.BackColor = Color.FromArgb(240, 240, 240)
                btnCancel.FlatStyle = FlatStyle.Flat
                btnCancel.FlatAppearance.BorderSize = 0
                btnCancel.Cursor = Cursors.Hand
                btnCancel.Location = New Point(20, 490)
                btnCancel.Size = New Size(150, 50)
                AddHandler btnCancel.Click, AddressOf btnCancel_Click
                Me.Controls.Add(btnCancel)

            Catch ex As Exception
                MessageBox.Show("Error initializing components: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Create a registry label pair
        ''' </summary>
        Private Sub CreateRegistryLabel(labelText As String, valueText As String, yPos As Integer, parent As GroupBox, ByRef valueLabel As Label)
            Dim lbl As New Label()
            lbl.Text = labelText
            lbl.Font = New Font("Segoe UI", 9)
            lbl.ForeColor = Color.Gray
            lbl.Location = New Point(20, yPos)
            lbl.Size = New Size(200, 20)
            parent.Controls.Add(lbl)

            valueLabel = New Label()
            valueLabel.Text = valueText
            valueLabel.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            valueLabel.ForeColor = Color.Black
            valueLabel.Location = New Point(250, yPos)
            valueLabel.Size = New Size(150, 20)
            parent.Controls.Add(valueLabel)
        End Sub

        ''' <summary>
        ''' Scan registry button click
        ''' </summary>
        Private Sub btnScanRegistry_Click(sender As Object, e As EventArgs)
            Try
                LogMessage("Scanning registry for prefix: " & txtPrefix.Text)

                Dim prefix As String = txtPrefix.Text.Trim()

                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                ' Scan registry
                Dim counters As System.Collections.Generic.Dictionary(Of String, Integer) = _
                    m_regManager.ScanCounters(prefix)

                ' Update display
                If counters.ContainsKey("PL") Then
                    lblPL.Text = If(counters("PL") > 0, counters("PL").ToString(), "Not found")
                Else
                    lblPL.Text = "Not found"
                End If

                If counters.ContainsKey("B") Then
                    lblB.Text = If(counters("B") > 0, counters("B").ToString(), "Not found")
                Else
                    lblB.Text = "Not found"
                End If

                If counters.ContainsKey("CH") Then
                    lblCH.Text = If(counters("CH") > 0, counters("CH").ToString(), "Not found")
                Else
                    lblCH.Text = "Not found"
                End If

                If counters.ContainsKey("A") Then
                    lblA.Text = If(counters("A") > 0, counters("A").ToString(), "Not found")
                Else
                    lblA.Text = "Not found"
                End If

                If counters.ContainsKey("FL") Then
                    lblFL.Text = If(counters("FL") > 0, counters("FL").ToString(), "Not found")
                Else
                    lblFL.Text = "Not found"
                End If

                MessageBox.Show("Registry scan complete for prefix: " & prefix, _
                              "Scan Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)

                LogMessage("Registry scan completed")

            Catch ex As Exception
                MessageBox.Show("Error scanning registry: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Clone button click
        ''' </summary>
        Private Sub btnClone_Click(sender As Object, e As EventArgs)
            Try
                ' Validate inputs
                Dim prefix As String = txtPrefix.Text.Trim()

                If String.IsNullOrEmpty(prefix) Then
                    MessageBox.Show("Please enter a prefix", _
                                  "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                Dim cloneCount As Integer = CInt(numCloneCount.Value)

                ' Show progress
                ProgressBar.Visible = True
                lblProgress.Visible = True
                ProgressBar.Value = 0
                lblProgress.Text = "Initializing..."

                ' Disable buttons
                btnClone.Enabled = False
                btnScanRegistry.Enabled = False
                txtPrefix.Enabled = False
                numCloneCount.Enabled = False

                ' Perform cloning
                LogMessage("Starting assembly cloning: Prefix=" & prefix & ", Count=" & cloneCount)

                Dim cloner As New AssemblyCloner(m_invApp)
                cloner.Clone(prefix, cloneCount, AddressOf UpdateProgress)

                ' Show success
                ProgressBar.Value = 100
                lblProgress.Text = "Completed successfully!"

                MessageBox.Show("Assembly Cloner completed successfully!" & vbCrLf & vbCrLf & _
                              "Prefix: " & prefix & vbCrLf & _
                              "Clones Created: " & cloneCount, _
                              "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                LogMessage("Assembly cloning completed successfully")

                ' Close form
                Me.DialogResult = DialogResult.OK
                Me.Close()

            Catch ex As Exception
                ProgressBar.Visible = False
                lblProgress.Visible = False
                btnClone.Enabled = True
                btnScanRegistry.Enabled = True
                txtPrefix.Enabled = True
                numCloneCount.Enabled = True

                MessageBox.Show("Error during cloning: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                LogMessage("Error during cloning: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Update progress during cloning
        ''' </summary>
        Private Sub UpdateProgress(percent As Integer, status As String)
            Try
                If ProgressBar.InvokeRequired Then
                    ProgressBar.Invoke(New Action(Of Integer, String)(AddressOf UpdateProgress), percent, status)
                Else
                    ProgressBar.Value = percent
                    lblProgress.Text = status
                    Application.DoEvents()
                End If
            Catch ex As Exception
                ' Ignore progress update errors
            End Try
        End Sub

        ''' <summary>
        ''' Cancel button click
        ''' </summary>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Try
                LogMessage("Assembly Cloner cancelled by user")
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            Catch ex As Exception
                ' Ignore errors during close
            End Try
        End Sub

        ''' <summary>
        ''' Log message to file
        ''' </summary>
        Private Sub LogMessage(ByVal message As String)
            Try
                Dim logPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                logPath = System.IO.Path.Combine(logPath, "InventorAutomationSuite_Log.txt")

                Dim logMessage As String = DateTime.Now.ToString() & " - AssemblyCloner - " & message

                System.IO.File.AppendAllText(logPath, logMessage & vbCrLf)

            Catch ex As Exception
                ' Silently fail - logging is not critical
            End Try
        End Sub

    End Class

End Namespace
