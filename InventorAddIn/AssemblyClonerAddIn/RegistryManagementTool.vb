Imports Inventor

Namespace AssemblyClonerAddIn

    Public Class RegistryManagementTool

        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Using form As New RegistryManagementForm(m_InventorApp)
                form.ShowDialog()
            End Using
        End Sub

    End Class

End Namespace
