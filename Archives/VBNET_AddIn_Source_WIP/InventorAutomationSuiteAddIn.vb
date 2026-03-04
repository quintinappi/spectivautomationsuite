' ============================================================================
' INVENTOR AUTOMATION SUITE - ADD-IN SERVER
' ============================================================================
' Description: Main entry point for Inventor Add-in
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-20
' ============================================================================

Imports System
Imports System.AddIn
Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Win32

Namespace InventorAutomationSuiteAddIn

    ''' <summary>
    ''' Main add-in server class that Inventor loads
    ''' </summary>
    <Guid("YOUR-GUID-HERE-0000-000000000000"), _
    ProgId("InventorAutomationSuiteAddIn.StandardAddInServer"), _
    ClassInterface(ClassInterfaceType.AutoDual)> _
    Public Class StandardAddInServer
        Implements ApplicationAddInServer

        ' Private members
        Private m_inventorApplication As InventorApplication
        Private m_addInGuid As String
        Private m_button As ButtonDefinition

        ' Constants
        Private Const ADDIN_GUID As String = "YOUR-GUID-HERE-0000-000000000000"
        Private Const ADDIN_DISPLAY_NAME As String = "Inventor Automation Suite"
        Private Const ADDIN_DESCRIPTION As String = "Professional automation tools for Autodesk Inventor"

        ''' <summary>
        ''' Called when Inventor loads the add-in
        ''' </summary>
        Public Sub Activate(InventorApplication As InventorApplication, ByVal firstTime As Integer) Implements ApplicationAddInServer.Activate
            Try
                ' Store reference to Inventor application
                m_inventorApplication = InventorApplication

                ' Log activation
                If firstTime = 1 Then
                    LogMessage("Add-in activated for the first time")
                    InitializeFirstTime()
                Else
                    LogMessage("Add-in activated")
                End If

                ' Create UI elements
                CreateUI()

                LogMessage("Add-in activation successful")

            Catch ex As Exception
                MessageBox.Show("Error activating add-in: " & ex.Message, _
                              "Inventor Automation Suite", _
                              MessageBoxButtons.OK, _
                              MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Called when Inventor unloads the add-in
        ''' </summary>
        Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
            Try
                ' Clean up UI elements
                If m_button IsNot Nothing Then
                    m_button.Delete()
                    m_button = Nothing
                End If

                ' Release reference
                m_inventorApplication = Nothing

                LogMessage("Add-in deactivated")

            Catch ex As Exception
                MessageBox.Show("Error deactivating add-in: " & ex.Message, _
                              "Inventor Automation Suite", _
                              MessageBoxButtons.OK, _
                              MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Execute command when button is clicked
        ''' </summary>
        Public Sub ExecuteCommand(ByVal commandIndex As Integer) Implements ApplicationAddInServer.ExecuteCommand
            Try
                LogMessage("Execute command: " & commandIndex)

                ' Show main form
                ShowMainForm()

            Catch ex As Exception
                MessageBox.Show("Error executing command: " & ex.Message, _
                              "Inventor Automation Suite", _
                              MessageBoxButtons.OK, _
                              MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Called automatically by Inventor - not used in this add-in
        ''' </summary>
        Public Sub AutomationCallback(ByVal commandIndex As Integer) Implements ApplicationAddInServer.AutomationCallback
            ' Not used in this add-in
        End Sub

        ''' <summary>
        ''' Create UI elements in Inventor ribbon
        ''' </summary>
        Private Sub CreateUI()
            Try
                ' Get the UserInterfaceManager
                Dim uiManager As UserInterfaceManager
                uiManager = m_inventorApplication.UserInterfaceManager

                ' Get the Ribbon interface
                Dim ribbon As Ribbon
                ribbon = uiManager.Ribbons

                ' Get the Assembly ribbon tab (or create if needed)
                Dim assemblyTab As RibbonTab
                assemblyTab = ribbon.Item("id_AssemblyTab") ' Assembly tab

                ' Create a new panel in the Assembly tab
                Dim panel As RibbonPanel
                panel = assemblyTab.RibbonPanels.Add("Automation Suite", "AutomationSuitePanel", Guid.NewGuid().ToString())

                ' Create button definition
                Dim iconLarge As System.Drawing.Icon
                iconLarge = New System.Drawing.Icon(Me.GetType(), "Resources.AutomationSuiteIcon.ico")

                m_button = m_inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition( _
                    "ShowAutomationSuite", _
                    "ShowAutomationSuite", _
                    CommandTypesEnum.kQueryOnlyCmdType, _
                    ADDIN_GUID, _
                    "Inventor Automation Suite", _
                    "Launch the Inventor Automation Suite", _
                    iconLarge, _
                    iconLarge)

                ' Add button click handler
                AddHandler m_button.OnExecute, AddressOf Button_OnExecute

                ' Add button to panel
                panel.CommandControls.AddButton(m_button, True)

                LogMessage("UI created successfully")

            Catch ex As Exception
                Throw New Exception("Error creating UI: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Button click event handler
        ''' </summary>
        Private Sub Button_OnExecute()
            Try
                ShowMainForm()

            Catch ex As Exception
                MessageBox.Show("Error showing main form: " & ex.Message, _
                              "Inventor Automation Suite", _
                              MessageBoxButtons.OK, _
                              MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Show the main form
        ''' </summary>
        Private Sub ShowMainForm()
            Try
                ' Create and show main form
                Dim mainForm As New MainForm(m_inventorApplication)
                mainForm.Show()

            Catch ex As Exception
                Throw New Exception("Error showing main form: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' First-time initialization
        ''' </summary>
        Private Sub InitializeFirstTime()
            Try
                ' Create registry key for add-in settings
                Dim regKey As RegistryKey
                regKey = Registry.CurrentUser.CreateSubKey("Software\SpectivSolutions\InventorAutomationSuite")

                ' Store installation date
                regKey.SetValue("InstallDate", DateTime.Now.ToString())
                regKey.SetValue("Version", My.Application.Info.Version.ToString())

                regKey.Close()

                LogMessage("First-time initialization completed")

            Catch ex As Exception
                ' Don't throw exception during first-time init
                MessageBox.Show("Warning: Could not complete first-time initialization." & vbCrLf & vbCrLf & _
                              "Error: " & ex.Message, _
                              "Inventor Automation Suite", _
                              MessageBoxButtons.OK, _
                              MessageBoxIcon.Warning)
            End Try
        End Sub

        ''' <summary>
        ''' Log message to file (for debugging)
        ''' </summary>
        Private Sub LogMessage(ByVal message As String)
            Try
                Dim logPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                logPath = System.IO.Path.Combine(logPath, "InventorAutomationSuite_Log.txt")

                Dim logMessage As String = DateTime.Now.ToString() & " - " & message

                System.IO.File.AppendAllText(logPath, logMessage & vbCrLf)

            Catch ex As Exception
                ' Silently fail - logging is not critical
            End Try
        End Sub

    End Class

End Namespace
