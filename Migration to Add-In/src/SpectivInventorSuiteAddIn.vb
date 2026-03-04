' ============================================================================
' SPECTIV INVENTOR AUTOMATION SUITE - ADD-IN
' ============================================================================
' Description: Inventor Add-In for Spectiv automation tools
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' ADD-IN CAPABILITIES:
' - Assembly Cloner ribbon button
' - Assembly Cloner UI Form
'
' INTEGRATION:
' - Requires Inventor 2026 or later
' - Uses Inventor API (Inventor.Interop)
' - Follows Autodesk Add-In guidelines
'
' INSTALLATION:
' - Copy DLL to: %APPDATA%\Autodesk\ApplicationPlugins\SpectivInventorSuite\
' - Include: SpectivInventorSuite.addin manifest file
' ============================================================================

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Inventor

Namespace SpectivInventorSuite

    ' ========================================================================
    ' ADD-IN GUID - Unique identifier for this Add-In
    ' ========================================================================
    <Guid("A8E8C9A2-1234-4567-89AB-CDEF01234567")> _
    Public Class SpectivInventorSuiteAddIn

        Implements ApplicationAddInServer

        ' ========================================================================
        ' PRIVATE MEMBERS
        ' ========================================================================

        Private m_inventorApplication As InventorApplication
        Private m_uiControl As AssemblyClonerForm
        Private m_buttonDefinition As ButtonDefinition

        ' ========================================================================
        ' ACTIVATE - Called when Inventor starts
        ' ========================================================================

        Public Sub Activate(Application As InventorApplication, ByVal AddInSiteObject As Object) Implements ApplicationAddInServer.Actacte
            Try
                m_inventorApplication = Application

                ' Initialize user interface
                InitializeUI()

                LogMessage("Spectiv Inventor Automation Suite loaded successfully")

            Catch ex As Exception
                MessageBox.Show(
                    "Error loading Spectiv Inventor Automation Suite:" & vbCrLf & vbCrLf & ex.Message,
                    "Add-In Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

        ' ========================================================================
        ' DEACTIVATE - Called when Inventor closes
        ' ========================================================================

        Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
            Try
                ' Clean up
                If m_uiControl IsNot Nothing AndAlso Not m_uiControl.IsDisposed Then
                    m_uiControl.Close()
                    m_uiControl.Dispose()
                End If

                m_uiControl = Nothing
                m_inventorApplication = Nothing

            Catch ex As Exception
                ' Silently fail during shutdown
            End Try
        End Sub

        Public ReadOnly Property Automation() As Object Implements ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' ========================================================================
        ' UI INITIALIZATION
        ' ========================================================================

        Private Sub InitializeUI()
            Try
                ' Get the UserInterfaceManager
                Dim uiManager As UserInterfaceManager = m_inventorApplication.UserInterfaceManager

                ' Create Button Definition
                m_buttonDefinition = uiManager.ButtonDefinitions.AddButtonDefinition(
                    "AssemblyClonerButton",
                    CommandTypesEnum.kShapeEditCmdType,
                    "{A8E8C9A2-1234-4567-89AB-CDEF01234568}",
                    "Assembly Cloner",
                    "Clone assembly with all parts to new location",
                    "Assembly Cloner",
                    IconLocation := Path.Combine(GetAssemblyDirectory(), "AssemblyCloner.ico")
                )

                ' Add to Assembly Panel
                Dim asmPanel As CommandBar = Nothing
                Try
                    asmPanel = uiManager.CommandBars("AM:AssemblyPanel")
                Catch
                End Try

                If asmPanel IsNot Nothing Then
                    ' Find the "Manage" command bar controls
                    Dim manageCommand As CommandBarControl = Nothing
                    Try
                        manageCommand = asmPanel.Controls.Item(kManageCommandControl)
                    Catch
                    End Try

                    ' Add button after "Manage" button
                    If manageCommand IsNot Nothing Then
                        Dim assemblyClonerButton As CommandBarControl = asmPanel.Controls.AddButton(
                            m_buttonDefinition,
                            manageCommand.Index + 1
                        )
                    Else
                        ' Add at end if "Manage" not found
                        asmPanel.Controls.AddButton(m_buttonDefinition)
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(
                    "Error initializing UI: " & ex.Message,
                    "UI Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

        ' ========================================================================
        ' BUTTON CLICK HANDLER
        ' ========================================================================

        Private Sub OnButtonClick(Context As NameValueMap) Handles m_buttonDefinition.OnExecute
            Try
                ' Show Assembly Cloner Form
                ShowAssemblyClonerForm()

            Catch ex As Exception
                MessageBox.Show(
                    "Error launching Assembly Cloner: " & ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

        ' ========================================================================
        ' SHOW ASSEMBLY CLONER FORM
        ' ========================================================================

        Private Sub ShowAssemblyClonerForm()
            Try
                ' Create or reuse form
                If m_uiControl Is Nothing OrElse m_uiControl.IsDisposed Then
                    m_uiControl = New AssemblyClonerForm(m_inventorApplication)
                End If

                ' Show form as modal dialog
                m_uiControl.ShowDialog()

            Catch ex As Exception
                MessageBox.Show(
                    "Error showing Assembly Cloner form: " & ex.Message,
                    "Form Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
            End Try
        End Sub

        ' ========================================================================
        ' HELPER FUNCTIONS
        ' ========================================================================

        Private Sub LogMessage(message As String)
            Try
                ' Log to Inventor's message box if available, otherwise console
                If m_inventorApplication IsNot Nothing Then
                    ' Could write to a log file here
                End If
            Catch
            End Try
        End Sub

        Private Function GetAssemblyDirectory() As String
            Try
                ' Get directory where this Add-In DLL is located
                Dim assemblyLocation As String = Assembly.GetExecutingAssembly().Location
                Return Path.GetDirectoryName(assemblyLocation)
            Catch
                Return String.Empty
            End Try
        End Function

    End Class

End Namespace
