Imports Inventor
Imports System.Text

Namespace AssemblyClonerAddIn

    Public Class UpdateSameFolderDerivedPartsTool

        Private ReadOnly m_InventorApp As Inventor.Application
        Private m_LogPath As String = String.Empty
        Private m_LogBuffer As StringBuilder

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            m_LogBuffer = New StringBuilder()
        End Sub

        Public Function ExecuteOnActiveAssembly() As Integer
            Const kPartDocumentObject As Integer = 12290
            Const kAssemblyDocumentObject As Integer = 12291

            If m_InventorApp Is Nothing Then
                Throw New InvalidOperationException("Inventor application is not initialized.")
            End If

            Dim activeDoc As Document = m_InventorApp.ActiveDocument
            If activeDoc Is Nothing Then
                Throw New InvalidOperationException("No active document is open.")
            End If

            If activeDoc.DocumentType <> kAssemblyDocumentObject Then
                Throw New InvalidOperationException("Active document is not an assembly.")
            End If

            Dim assemblyDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
            Dim asmFolder As String = System.IO.Path.GetDirectoryName(assemblyDoc.FullFileName)
            Dim mappingPath As String = System.IO.Path.Combine(asmFolder, "STEP_1_MAPPING.txt")

            If Not System.IO.File.Exists(mappingPath) Then
                Throw New System.IO.FileNotFoundException("STEP_1_MAPPING.txt not found.", mappingPath)
            End If

            m_LogPath = System.IO.Path.Combine(asmFolder, "Update_Derived_Log.txt")
            m_LogBuffer.Clear()

            WriteLog("==========================================")
            WriteLog(" UPDATE SAME-FOLDER DERIVED PARTS")
            WriteLog("==========================================")
            WriteLog("Assembly: " & assemblyDoc.DisplayName)
            WriteLog("Mapping: " & mappingPath)
            WriteLog("")

            Dim mapping As Dictionary(Of String, String) = LoadMapping(mappingPath)
            Dim fixCount As Integer = 0

            For Each refDoc As Document In assemblyDoc.AllReferencedDocuments
                If refDoc.DocumentType = kPartDocumentObject Then
                    fixCount += FixPartDerivedRefs(CType(refDoc, PartDocument), asmFolder, mapping)
                End If
            Next

            WriteLog("")
            WriteLog("==========================================")
            WriteLog("COMPLETE")
            WriteLog("==========================================")
            WriteLog("Total fixes: " & fixCount)
            WriteLog("")

            System.IO.File.WriteAllText(m_LogPath, m_LogBuffer.ToString())
            Return fixCount
        End Function

        Private Function LoadMapping(ByVal mappingPath As String) As Dictionary(Of String, String)
            Dim map As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            For Each rawLine As String In System.IO.File.ReadAllLines(mappingPath)
                Dim line As String = rawLine.Trim()
                If line = "" OrElse line.StartsWith("#") Then
                    Continue For
                End If

                Dim parts() As String = line.Split("|"c)
                If parts.Length < 4 Then
                    Continue For
                End If

                Dim originalBase As String = System.IO.Path.GetFileNameWithoutExtension(parts(2).Trim())
                Dim newBase As String = System.IO.Path.GetFileNameWithoutExtension(parts(3).Trim())

                If originalBase <> "" AndAlso newBase <> "" Then
                    map(originalBase) = newBase
                End If
            Next

            Return map
        End Function

        Private Function FixPartDerivedRefs(
            ByVal partDoc As PartDocument,
            ByVal asmFolder As String,
            ByVal mapping As Dictionary(Of String, String)) As Integer

            Dim fixedInPart As Integer = 0

            Try
                Dim partName As String = System.IO.Path.GetFileName(partDoc.FullFileName)
                Dim derivedParts As DerivedPartComponents = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents

                If derivedParts Is Nothing OrElse derivedParts.Count = 0 Then
                    Return 0
                End If

                For i As Integer = 1 To derivedParts.Count
                    Dim derivedComp As DerivedPartComponent = derivedParts.Item(i)
                    If derivedComp Is Nothing OrElse Not derivedComp.LinkedToFile Then
                        Continue For
                    End If

                    Dim docDesc As DocumentDescriptor = derivedComp.ReferencedDocumentDescriptor
                    If docDesc Is Nothing Then
                        Continue For
                    End If

                    Dim basePath As String = docDesc.FullDocumentName
                    Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(basePath)

                    If Not mapping.ContainsKey(baseName) Then
                        Continue For
                    End If

                    Dim newBasePath As String = System.IO.Path.Combine(asmFolder, mapping(baseName) & ".ipt")
                    If Not System.IO.File.Exists(newBasePath) Then
                        WriteLog("Skipping missing mapped file for: " & partName)
                        WriteLog("  Expected: " & newBasePath)
                        WriteLog("")
                        Continue For
                    End If

                    WriteLog("Fixing: " & partName)
                    WriteLog("  Old Base: " & System.IO.Path.GetFileName(basePath))
                    WriteLog("  New Base: " & System.IO.Path.GetFileName(newBasePath))

                    Try
                        derivedComp.Replace(newBasePath, Nothing)
                        partDoc.Update2(True)
                        partDoc.Save2(True)
                        WriteLog("  SUCCESS")
                        fixedInPart += 1
                    Catch ex As Exception
                        WriteLog("  ERROR: " & ex.Message)
                    End Try

                    WriteLog("")
                Next

            Catch ex As Exception
                WriteLog("ERROR scanning part: " & partDoc.DisplayName)
                WriteLog("  " & ex.Message)
                WriteLog("")
            End Try

            Return fixedInPart
        End Function

        Private Sub WriteLog(ByVal text As String)
            m_LogBuffer.AppendLine(text)
        End Sub

    End Class

End Namespace
