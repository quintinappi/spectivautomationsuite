Imports Inventor

Namespace AssemblyClonerAddIn

    Public Class AssemblyRenamerTool

        Private ReadOnly m_Renamer As PartRenamer

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_Renamer = New PartRenamer(inventorApp)
        End Sub

        Public Sub Execute()
            m_Renamer.RenameAssemblyParts()
        End Sub

    End Class

End Namespace
