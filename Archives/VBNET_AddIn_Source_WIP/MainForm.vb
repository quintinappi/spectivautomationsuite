' ============================================================================
' INVENTOR AUTOMATION SUITE - MAIN FORM
' ============================================================================
' Description: Unified launcher form for all automation tools
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
    ''' Main form that provides access to all automation tools
    ''' </summary>
    Public Class MainForm
        Inherits Form

        ' Private members
        Private m_invApp As InventorApplication
        Private m_regManager As RegistryManager

        ' UI Components (will be initialized in InitializeComponent)
        Private btnAssemblyCloner As Button
        Private btnPartRenaming As Button
        Private btnIDWUpdates As Button
        Private btnTitleAutomation As Button
        Private btnRegistryManagement As Button
        Private btnPrefixScanner As Button
        Private btnEmergencyFixer As Button
        Private btnMappingProtection As Button
        Private lblStatus As Label
        Private lblVersion As Label

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(invApp As InventorApplication)
            Try
                m_invApp = invApp
                m_regManager = New RegistryManager()

                ' Initialize form components
                InitializeComponent()

                ' Set form properties
                Me.Text = "Inventor Automation Suite - Professional Edition"
                Me.Size = New Size(1000, 700)
                Me.StartPosition = FormStartPosition.CenterScreen
                Me.MinimizeBox = True
                Me.MaximizeBox = False
                Me.FormBorderStyle = FormBorderStyle.FixedSingle
                Me.BackColor = Color.FromArgb(245, 246, 247)

                LogMessage("Main form loaded")

            Catch ex As Exception
                MessageBox.Show("Error initializing main form: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Initialize all form components
        ''' </summary>
        Private Sub InitializeComponent()
            Try
                ' Create main panel
                Dim mainPanel As New Panel()
                mainPanel.Dock = DockStyle.Fill
                mainPanel.Padding = New Padding(20)
                mainPanel.BackColor = Color.White
                Me.Controls.Add(mainPanel)

                ' Create title label
                Dim titleLabel As New Label()
                titleLabel.Text = "Inventor Automation Suite"
                titleLabel.Font = New Font("Segoe UI", 24, FontStyle.Bold)
                titleLabel.ForeColor = Color.FromArgb(102, 126, 234)
                titleLabel.Location = New Point(20, 20)
                titleLabel.Size = New Size(600, 40)
                mainPanel.Controls.Add(titleLabel)

                ' Create subtitle label
                Dim subtitleLabel As New Label()
                subtitleLabel.Text = "Professional Edition v1.0 - Select a tool to begin"
                subtitleLabel.Font = New Font("Segoe UI", 10)
                subtitleLabel.ForeColor = Color.Gray
                subtitleLabel.Location = New Point(20, 65)
                subtitleLabel.Size = New Size(500, 20)
                mainPanel.Controls.Add(subtitleLabel)

                ' Create Main Workflow section
                CreateToolSection(mainPanel, "Main Workflow", 120, New String() {
                    "Assembly Cloner",
                    "Part Renaming",
                    "IDW Updates",
                    "Title Automation"
                })

                ' Create Management Tools section
                CreateToolSection(mainPanel, "Management Tools", 360, New String() {
                    "Registry Management",
                    "Smart Prefix Scanner"
                })

                ' Create Rescue Tools section
                CreateToolSection(mainPanel, "Rescue Tools", 520, New String() {
                    "Emergency IDW Fixer",
                    "Mapping Protection"
                })

                ' Create status bar at bottom
                Dim statusPanel As New Panel()
                statusPanel.Dock = DockStyle.Bottom
                statusPanel.Height = 40
                statusPanel.BackColor = Color.FromArgb(240, 240, 240)
                Me.Controls.Add(statusPanel)

                ' License status
                Dim licenseLabel As New Label()
                licenseLabel.Text = "● License: Professional Edition"
                licenseLabel.Location = New Point(20, 12)
                licenseLabel.Font = New Font("Segoe UI", 9)
                licenseLabel.ForeColor = Color.Green
                statusPanel.Controls.Add(licenseLabel)

                ' Version info
                lblVersion = New Label()
                lblVersion.Text = "Version 1.0.0 | © 2025 Spectiv Solutions"
                lblVersion.Location = New Point(700, 12)
                lblVersion.Font = New Font("Segoe UI", 9)
                lblVersion.ForeColor = Color.Gray
                statusPanel.Controls.Add(lblVersion)

            Catch ex As Exception
                MessageBox.Show("Error initializing components: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Create a tool section with buttons
        ''' </summary>
        Private Sub CreateToolSection(parentPanel As Panel, sectionTitle As String, yPos As Integer, toolNames As String())
            Try
                ' Create section title
                Dim sectionLabel As New Label()
                sectionLabel.Text = sectionTitle.ToUpper()
                sectionLabel.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                sectionLabel.ForeColor = Color.FromArgb(118, 75, 162)
                sectionLabel.Location = New Point(20, yPos)
                sectionLabel.Size = New Size(500, 25)
                parentPanel.Controls.Add(sectionLabel)

                ' Create buttons for each tool
                Dim i As Integer = 0
                For Each toolName As String In toolNames
                    Dim toolButton As New Button()
                    toolButton.Text = (i + 1) & ". " & toolName
                    toolButton.Font = New Font("Segoe UI", 10)
                    toolButton.BackColor = Color.White
                    toolButton.ForeColor = Color.Black
                    toolButton.FlatStyle = FlatStyle.Flat
                    toolButton.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220)
                    toolButton.FlatAppearance.BorderSize = 2
                    toolButton.Cursor = Cursors.Hand
                    toolButton.Size = New Size(400, 50)
                    toolButton.Location = New Point(20, yPos + 30 + (i * 55))
                    toolButton.TextAlign = ContentAlignment.MiddleLeft
                    toolButton.Padding = New Padding(15, 0, 0, 0)

                    ' Add hover effect
                    AddHandler toolButton.MouseEnter, Sub(sender, e)
                                                        Dim btn As Button = DirectCast(sender, Button)
                                                        btn.BackColor = Color.FromArgb(102, 126, 234)
                                                        btn.ForeColor = Color.White
                                                        btn.FlatAppearance.BorderColor = Color.FromArgb(102, 126, 234)
                                                    End Sub

                    AddHandler toolButton.MouseLeave, Sub(sender, e)
                                                        Dim btn As Button = DirectCast(sender, Button)
                                                        btn.BackColor = Color.White
                                                        btn.ForeColor = Color.Black
                                                        btn.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220)
                                                    End Sub

                    ' Add click handler
                    AddHandler toolButton.Click, AddressOf ToolButton_Click

                    parentPanel.Controls.Add(toolButton)
                    i += 1
                Next

            Catch ex As Exception
                MessageBox.Show("Error creating tool section: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Tool button click handler
        ''' </summary>
        Private Sub ToolButton_Click(sender As Object, e As EventArgs)
            Try
                Dim button As Button = DirectCast(sender, Button)
                Dim toolName As String = button.Text.Substring(3) ' Remove "1. ", "2. ", etc.

                LogMessage("Tool selected: " & toolName)

                Select Case toolName
                    Case "Assembly Cloner"
                        ShowAssemblyCloner()
                    Case "Part Renaming"
                        ShowPartRenaming()
                    Case "IDW Updates"
                        ShowIDWUpdates()
                    Case "Title Automation"
                        ShowTitleAutomation()
                    Case "Registry Management"
                        ShowRegistryManagement()
                    Case "Smart Prefix Scanner"
                        ShowPrefixScanner()
                    Case "Emergency IDW Fixer"
                        ShowEmergencyFixer()
                    Case "Mapping Protection"
                        ShowMappingProtection()
                    Case Else
                        MessageBox.Show("Tool not yet implemented: " & toolName, _
                                      "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Select

            Catch ex As Exception
                MessageBox.Show("Error opening tool: " & ex.Message, _
                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Show Assembly Cloner form
        ''' </summary>
        Private Sub ShowAssemblyCloner()
            Try
                Dim form As New AssemblyClonerForm(m_invApp, m_regManager)
                form.ShowDialog(Me)
            Catch ex As Exception
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Show Part Renaming form
        ''' </summary>
        Private Sub ShowPartRenaming()
            MessageBox.Show("Part Renaming tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing COMPLETE_WORKING_SOLUTION.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show IDW Updates form
        ''' </summary>
        Private Sub ShowIDWUpdates()
            MessageBox.Show("IDW Updates tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing STEP_2_IDW_Updater.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show Title Automation form
        ''' </summary>
        Private Sub ShowTitleAutomation()
            MessageBox.Show("Title Automation tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing Title_Updater.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show Registry Management form
        ''' </summary>
        Private Sub ShowRegistryManagement()
            MessageBox.Show("Registry Management tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing Registry_Manager.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show Smart Prefix Scanner form
        ''' </summary>
        Private Sub ShowPrefixScanner()
            MessageBox.Show("Smart Prefix Scanner tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing Smart_Prefix_Scanner.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show Emergency IDW Fixer form
        ''' </summary>
        Private Sub ShowEmergencyFixer()
            MessageBox.Show("Emergency IDW Fixer tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing Emergency_IDW_Fixer.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Show Mapping Protection form
        ''' </summary>
        Private Sub ShowMappingProtection()
            MessageBox.Show("Mapping Protection tool - Coming soon!" & vbCrLf & vbCrLf & _
                          "This tool will integrate the existing Protect_Mapping_File.vbs functionality.", _
                          "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        ''' <summary>
        ''' Log message to file
        ''' </summary>
        Private Sub LogMessage(ByVal message As String)
            Try
                Dim logPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                logPath = System.IO.Path.Combine(logPath, "InventorAutomationSuite_Log.txt")

                Dim logMessage As String = DateTime.Now.ToString() & " - MainForm - " & message

                System.IO.File.AppendAllText(logPath, logMessage & vbCrLf)

            Catch ex As Exception
                ' Silently fail - logging is not critical
            End Try
        End Sub

        ''' <summary>
        ''' Form closing event
        ''' </summary>
        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            Try
                LogMessage("Main form closing")
                MyBase.OnFormClosing(e)
            Catch ex As Exception
                ' Ignore errors during cleanup
            End Try
        End Sub

    End Class

End Namespace
