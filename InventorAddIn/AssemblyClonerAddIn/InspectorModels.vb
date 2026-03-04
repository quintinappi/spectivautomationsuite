Imports Inventor

Namespace AssemblyClonerAddIn

    Public Class ParamInfo
        Public Property Name As String
        Public Property Expression As String
        Public Property IsUserParameter As Boolean
        Public Property Document As Document
        Public Property ParamObject As Object
    End Class

    Public Class DocScanResult
        Public Property Doc As Document
        Public Property DocName As String
        Public Property DocPath As String
        Public Property Rules As Dictionary(Of String, String)
        Public Property Params As New System.Collections.Generic.List(Of ParamInfo)
    End Class

End Namespace
