Imports Inventor
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Collections.Generic

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Experimental: Inspects the active assembly, lists parts, and finds iLogic parameters/forms.
    ''' </summary>
    Public Class AssemblySmartInspector
        Private m_InventorApp As Inventor.Application
        Private m_Patcher As iLogicPatcher

        Public Sub New(app As Inventor.Application, Optional patcher As iLogicPatcher = Nothing)
            m_InventorApp = app
            If patcher IsNot Nothing Then
                m_Patcher = patcher
            Else
                m_Patcher = New iLogicPatcher(app)
            End If
        End Sub


        ''' <summary>
        ''' Main entry point for inspection.
        ''' </summary>
        Public Sub InspectActiveAssembly()
            Dim sb As New StringBuilder()

            Try
                Dim asmDoc As AssemblyDocument = TryCast(m_InventorApp.ActiveDocument, AssemblyDocument)
                If asmDoc Is Nothing Then
                    MsgBox("No assembly is active.", MsgBoxStyle.Exclamation)
                    Return
                End If

                sb.AppendLine("=== Assembly Inspector ===")
                sb.AppendLine("Assembly: " & asmDoc.DisplayName)
                sb.AppendLine("Path: " & asmDoc.FullFileName)
                sb.AppendLine()

                ' Gather all referenced documents (assembly + parts + sub-assemblies)
                Dim docs As List(Of Document) = CollectAllReferencedDocs(asmDoc)

                sb.AppendLine("Parts and referenced documents:")
                For Each d As Document In docs
                    sb.AppendLine(" - " & d.DisplayName & " (" & d.FullFileName & ")")
                Next
                sb.AppendLine()

                ' Scan each document for parameters and iLogic rules using our patcher
                Dim docResults As New List(Of DocScanResult)
                For Each d As Document In docs
                    Try
                        Dim dr As New DocScanResult()
                        dr.Doc = d
                        dr.DocName = d.DisplayName
                        dr.DocPath = d.FullFileName
                        dr.Rules = m_Patcher.GetRulesAsDictionary(d)

                        ' Get user parameters for the document (if available)
                        Try
                            Dim paramsObj As Inventor.Parameters = Nothing
                            If d.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                                paramsObj = CType(d, PartDocument).ComponentDefinition.Parameters
                            ElseIf d.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                                paramsObj = CType(d, AssemblyDocument).ComponentDefinition.Parameters
                            End If

                            If paramsObj IsNot Nothing Then
                                Dim userParams As Inventor.UserParameters = paramsObj.UserParameters
                                If userParams IsNot Nothing Then
                                    For Each up As UserParameter In userParams
                                        Dim pi As New ParamInfo()
                                        pi.Name = up.Name
                                        pi.Expression = up.Expression
                                        pi.IsUserParameter = True
                                        pi.Document = d
                                        pi.ParamObject = up
                                        dr.Params.Add(pi)
                                    Next
                                End If
                            End If
                        Catch ex As Exception
                            ' ignore parameter extraction errors
                        End Try

                        docResults.Add(dr)
                    Catch ex As Exception
                        ' ignore docs we cannot read
                    End Try
                Next

                ' Show a form with the collected data
                Dim uiForm As New AssemblyInspectorForm(asmDoc, docResults)
                uiForm.ShowDialog()

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Inspector Error")
            End Try
        End Sub

        ''' <summary>
        ''' Collects the assembly document and all referenced documents recursively.
        ''' Returns a unique list of documents.
        ''' </summary>
        Private Function CollectAllReferencedDocs(ByVal asmDoc As AssemblyDocument) As List(Of Document)
            Dim result As New List(Of Document)()
            Dim processed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Add assembly itself
            result.Add(asmDoc)
            processed.Add(asmDoc.FullFileName)

            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                CollectOccurrenceDocsRecursive(occ, result, processed)
            Next

            Return result
        End Function

        Private Sub CollectOccurrenceDocsRecursive(ByVal occ As ComponentOccurrence, ByVal result As List(Of Document), ByVal processed As HashSet(Of String))
            Try
                If occ Is Nothing OrElse occ.Definition Is Nothing Then Return
                Dim refDoc As Document = occ.Definition.Document
                If refDoc IsNot Nothing AndAlso Not processed.Contains(refDoc.FullFileName) Then
                    processed.Add(refDoc.FullFileName)
                    result.Add(refDoc)

                    ' If sub-assembly, recurse
                    If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        Dim subAsm As AssemblyDocument = CType(refDoc, AssemblyDocument)
                        For Each subOcc As ComponentOccurrence In subAsm.ComponentDefinition.Occurrences
                            CollectOccurrenceDocsRecursive(subOcc, result, processed)
                        Next
                    End If
                End If
            Catch ex As Exception
                ' ignore
            End Try
        End Sub

    End Class

End Namespace
