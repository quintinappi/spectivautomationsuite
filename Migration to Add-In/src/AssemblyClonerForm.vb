' ============================================================================
' SPECTIV INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER FORM
' ============================================================================
' Description: Windows Form for Assembly Cloner user input
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' UI CONTROLS:
' - Destination folder browser
' - Heritage naming options
' - Prefix input
' - Progress tracking
' - Log output
'
' INTEGRATION:
' - Called from ribbon button
' - Instantiates AssemblyCloner class
' - Displays progress and results
' ============================================================================

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SpectivInventorSuite

Namespace SpectivInventorSuite

    ''' <summary>
    ''' Assembly Cloner UI Form
    ''' </summary>
    Public Class AssemblyClonerForm
        Inherits Form

        ' ========================================================================
        ' PRIVATE MEMBERS
        ' ========================================================================

        Private m_invApp As InventorApplication
        Private m_assemblyCloner As AssemblyCloner
        Private m_destinationFolder As String
        Private m_logFilePath As String

        ' ========================================================================
        ' UI CONTROLS
        ' ========================================================================

        Private WithEvents lblTitle As Label
        Private WithEvents grpDestination As GroupBox
        Private WithEvents txtDestination As TextBox
        Private WithEvents btnBrowse As Button
        Private WithEvents grpOptions As GroupBox
        Private WithEvents chkRename As CheckBox
        Private WithEvents txtPrefix As TextBox
        Private WithEvents grpProgress As GroupBox
        Private WithEvents prgProgress As ProgressBar
        Private WithEvents lblStatus As Label
        Private WithEvents grpLog As GroupBox
        Private WithEvents txtLog As TextBox
        Private WithEvents btnClone As Button
        Private WithEvents btnCancel As Button
        Private WithEvents btnClose As Button

        ' ========================================================================
        ' CONSTRUCTOR
        ' ========================================================================

        Public Sub New(invApp As InventorApplication)
            If invApp Is Nothing Then
                Throw New ArgumentNullException(NameOf(invApp))
            End If

            m_invApp = invApp
            m_assemblyCloner = New AssemblyCloner(invApp)

            InitializeComponent()
            InitializeDefaults()
        End Sub

        ' ========================================================================
        ' FORM INITIALIZATION
        ' ========================================================================

        Private Sub InitializeComponent()
            '
            ' Form
            '
            Me.Text = "Assembly Cloner - Spectiv Inventor Suite"
            Me.Size = New Size(600, 700)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False

            '
            ' Title Label
            '
            lblTitle = New Label()
            lblTitle.Text = "ASSEMBLY CLONER"
            lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
            lblTitle.ForeColor = Color.FromArgb(0, 122, 204)
            lblTitle.Location = New Point(20, 15)
            lblTitle.Size = New Size(560, 30)
            lblTitle.TextAlign = ContentAlignment.MiddleCenter
            Me.Controls.Add(lblTitle)

            '
            ' Destination Group
            '
            grpDestination = New GroupBox()
            grpDestination.Text = "Destination Location"
            grpDestination.Location = New Point(20, 60)
            grpDestination.Size = New Size(560, 70)
            Me.Controls.Add(grpDestination)

            txtDestination = New TextBox()
            txtDestination.Location = New Point(10, 25)
            txtDestination.Size = New Size(450, 25)
            txtDestination.ReadOnly = True
            txtDestination.BackColor = Color.White
            grpDestination.Controls.Add(txtDestination)

            btnBrowse = New Button()
            btnBrowse.Text = "Browse..."
            btnBrowse.Location = New Point(470, 23)
            btnBrowse.Size = New Size(80, 28)
            AddHandler btnBrowse.Click, AddressOf btnBrowse_Click
            grpDestination.Controls.Add(btnBrowse)

            '
            ' Options Group
            '
            grpOptions = New GroupBox()
            grpOptions.Text = "Heritage Naming Options"
            grpOptions.Location = New Point(20, 145)
            grpOptions.Size = New Size(560, 100)
            Me.Controls.Add(grpOptions)

            chkRename = New CheckBox()
            chkRename.Text = "Enable heritage naming (rename parts with prefix + counter)"
            chkRename.Location = New Point(10, 25)
            chkRename.Size = New Size(530, 25)
            chkRename.Checked = False
            AddHandler chkRename.CheckedChanged, AddressOf chkRename_CheckedChanged
            grpOptions.Controls.Add(chkRename)

            Dim lblPrefix As Label = New Label()
            lblPrefix.Text = "Project Prefix:"
            lblPrefix.Location = New Point(30, 60)
            lblPrefix.Size = New Size(100, 20)
            grpOptions.Controls.Add(lblPrefix)

            txtPrefix = New TextBox()
            txtPrefix.Text = "CLONE-001-"
            txtPrefix.Location = New Point(140, 58)
            txtPrefix.Size = New Size(150, 25)
            txtPrefix.Enabled = False
            grpOptions.Controls.Add(txtPrefix)

            '
            ' Progress Group
            '
            grpProgress = New GroupBox()
            grpProgress.Text = "Progress"
            grpProgress.Location = New Point(20, 255)
            grpProgress.Size = New Size(560, 80)
            Me.Controls.Add(grpProgress)

            prgProgress = New ProgressBar()
            prgProgress.Location = New Point(10, 45)
            prgProgress.Size = New Size(540, 20)
            prgProgress.Style = ProgressBarStyle.Continuous
            grpProgress.Controls.Add(prgProgress)

            lblStatus = New Label()
            lblStatus.Text = "Ready"
            lblStatus.Location = New Point(10, 20)
            lblStatus.Size = New Size(540, 20)
            grpProgress.Controls.Add(lblStatus)

            '
            ' Log Group
            '
            grpLog = New GroupBox()
            grpLog.Text = "Activity Log"
            grpLog.Location = New Point(20, 345)
            grpLog.Size = New Size(560, 230)
            Me.Controls.Add(grpLog)

            txtLog = New TextBox()
            txtLog.Location = New Point(10, 20)
            txtLog.Size = New Size(540, 195)
            txtLog.Multiline = True
            txtLog.ScrollBars = ScrollBars.Vertical
            txtLog.ReadOnly = True
            txtLog.BackColor = Color.White
            txtLog.Font = New Font("Consolas", 9)
            grpLog.Controls.Add(txtLog)

            '
            ' Action Buttons
            '
            btnClone = New Button()
            btnClone.Text = "&Clone Assembly"
            btnClone.Location = New Point(20, 590)
            btnClone.Size = New Size(150, 40)
            btnClone.BackColor = Color.FromArgb(0, 122, 204)
            btnClone.ForeColor = Color.White
            btnClone.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            AddHandler btnClone.Click, AddressOf btnClone_Click
            Me.Controls.Add(btnClone)

            btnCancel = New Button()
            btnCancel.Text = "&Cancel"
            btnCancel.Location = New Point(190, 590)
            btnCancel.Size = New Size(120, 40)
            btnCancel.Enabled = False
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            Me.Controls.Add(btnCancel)

            btnClose = New Button()
            btnClose.Text = "&Close"
            btnClose.Location = New Point(440, 590)
            btnClose.Size = New Size(140, 40)
            AddHandler btnClose.Click, AddressOf btnClose_Click
            Me.Controls.Add(btnClose)
        End Sub

        Private Sub InitializeDefaults()
            m_destinationFolder = String.Empty
            UpdateStatus("Ready - Select destination folder and click Clone")
        End Sub

        ' ========================================================================
        ' EVENT HANDLERS
        ' ========================================================================

        Private Sub btnBrowse_Click(sender As Object, e As EventArgs)
            Try
                Using dialog As New FolderBrowserDialog()
                    dialog.Description = "Select DESTINATION folder for the cloned assembly:" & vbCrLf & vbCrLf &
                                       "TIP: Click 'Make New Folder' to create a new destination"

                    ' Start at parent of current document
                    If m_invApp.ActiveDocument IsNot Nothing Then
                        Dim currentPath As String = m_invApp.ActiveDocument.FullFileName
                        Dim parentDir As String = Path.GetDirectoryName(currentPath)
                        If Not String.IsNullOrEmpty(parentDir) Then
                            dialog.SelectedPath = Path.GetDirectoryName(parentDir)
                        End If
                    End If

                    dialog.ShowNewFolderButton = True

                    If dialog.ShowDialog() = DialogResult.OK Then
                        m_destinationFolder = dialog.SelectedPath
                        txtDestination.Text = m_destinationFolder
                        LogMessage("Destination selected: " & m_destinationFolder)
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show(
                    "Error selecting folder: " & ex.Message,
                    "Folder Selection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub chkRename_CheckedChanged(sender As Object, e As EventArgs)
            txtPrefix.Enabled = chkRename.Checked
            UpdateStatus("Ready")
        End Sub

        Private Sub btnClone_Click(sender As Object, e As EventArgs)
            If ValidateInputs() Then
                DisableControls()
                RunClone()
            End If
        End Sub

        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            ' Note: Cancellation would require background worker
            MessageBox.Show(
                "Cancellation not yet implemented. Please let the current operation complete.",
                "Info",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs)
            Me.Close()
        End Sub

        ' ========================================================================
        ' VALIDATION
        ' ========================================================================

        Private Function ValidateInputs() As Boolean
            ' Check destination
            If String.IsNullOrEmpty(m_destinationFolder) OrElse Not Directory.Exists(m_destinationFolder) Then
                MessageBox.Show(
                    "Please select a valid destination folder.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation)
                Return False
            End If

            ' Check prefix if renaming enabled
            If chkRename.Checked AndAlso String.IsNullOrWhiteSpace(txtPrefix.Text) Then
                MessageBox.Show(
                    "Please enter a project prefix for heritage naming.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation)
                Return False
            End If

            ' Check active document
            If m_invApp.ActiveDocument Is Nothing Then
                MessageBox.Show(
                    "No document is currently open in Inventor." & vbCrLf & vbCrLf &
                    "Please open an assembly file first.",
                    "No Active Document",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation)
                Return False
            End If

            If m_invApp.ActiveDocument.Type <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MessageBox.Show(
                    "Current file is not an assembly!" & vbCrLf & vbCrLf &
                    "File: " & m_invApp.ActiveDocument.DisplayName & vbCrLf & vbCrLf &
                    "Please open an assembly file.",
                    "Not an Assembly",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation)
                Return False
            End If

            ' Confirm with user
            Dim sourceDoc As AssemblyDocument = DirectCast(m_invApp.ActiveDocument, AssemblyDocument)
            Dim result As DialogResult = MessageBox.Show(
                "SOURCE ASSEMBLY:" & vbCrLf & vbCrLf &
                "Assembly: " & sourceDoc.DisplayName & vbCrLf &
                "Parts: " & sourceDoc.ComponentDefinition.Occurrences.Count.ToString() & vbCrLf & vbCrLf &
                "DESTINATION:" & vbCrLf & vbCrLf &
                m_destinationFolder & vbCrLf & vbCrLf &
                "OPTIONS:" & vbCrLf &
                "Heritage Naming: " & If(chkRename.Checked, "Enabled (" & txtPrefix.Text & ")", "Disabled") & vbCrLf & vbCrLf &
                "Proceed with cloning?",
                "Confirm Clone",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1)

            Return result = DialogResult.Yes
        End Function

        ' ========================================================================
        ' RUN CLONE OPERATION
        ' ========================================================================

        Private Sub RunClone()
            Try
                UpdateStatus("Starting clone operation...")
                LogMessage("========================================")
                LogMessage("ASSEMBLY CLONER STARTING")
                LogMessage("========================================")

                ' Run clone
                Dim success As Boolean = m_assemblyCloner.Clone(
                    destinationFolder:=m_destinationFolder,
                    renameParts:=chkRename.Checked,
                    prefix:=txtPrefix.Text.ToUpper() & "-"c,
                    progressCallback:=AddressOf ProgressUpdate
                )

                If success Then
                    UpdateStatus("Clone completed successfully!")
                    LogMessage("========================================")
                    LogMessage("CLONE COMPLETED SUCCESSFULLY")
                    LogMessage("========================================")

                    MessageBox.Show(
                        "Assembly cloned successfully!" & vbCrLf & vbCrLf &
                        "Destination: " & m_destinationFolder & vbCrLf & vbCrLf &
                        "Check the log file for details.",
                        "Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

                    ' Open destination folder
                    Process.Start("explorer.exe", m_destinationFolder)
                Else
                    UpdateStatus("Clone failed - check log for details")
                    MessageBox.Show(
                        "Clone operation failed. Please check the log for details.",
                        "Clone Failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
                End If

            Catch ex As Exception
                LogMessage("ERROR: " & ex.Message)
                UpdateStatus("Clone failed with error")
                MessageBox.Show(
                    "Error during clone: " & ex.Message,
                    "Clone Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            Finally
                EnableControls()
            End Try
        End Sub

        ' ========================================================================
        ' PROGRESS CALLBACK
        ' ========================================================================

        Private Sub ProgressUpdate(percent As Integer, status As String)
            ' Update UI from main thread
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Sub()
                                        prgProgress.Value = percent
                                        lblStatus.Text = status
                                        LogMessage(status)
                                    End Sub))
            Else
                prgProgress.Value = percent
                lblStatus.Text = status
                LogMessage(status)
            End If

            ' Process UI events
            Application.DoEvents()
        End Sub

        ' ========================================================================
        ' UI HELPERS
        ' ========================================================================

        Private Sub UpdateStatus(status As String)
            lblStatus.Text = status
            Application.DoEvents()
        End Sub

        Private Sub LogMessage(message As String)
            txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") & " - " & message & vbCrLf)
            txtLog.SelectionStart = txtLog.Text.Length
            txtLog.ScrollToCaret()
            Application.DoEvents()
        End Sub

        Private Sub DisableControls()
            grpDestination.Enabled = False
            grpOptions.Enabled = False
            btnClone.Enabled = False
            btnCancel.Enabled = True
            btnClose.Enabled = False
        End Sub

        Private Sub EnableControls()
            grpDestination.Enabled = True
            grpOptions.Enabled = True
            btnClone.Enabled = True
            btnCancel.Enabled = False
            btnClose.Enabled = True
        End Sub

    End Class

End Namespace
