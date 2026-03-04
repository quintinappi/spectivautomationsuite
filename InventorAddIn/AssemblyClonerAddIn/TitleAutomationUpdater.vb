Imports Inventor
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class TitleAutomationUpdater
        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            Try
                If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    MessageBox.Show("Open an IDW drawing first.", "Title Automation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
                If Not IsIdwDocument(drawingDoc) Then
                    MessageBox.Show("This command is for IDW drawings only.", "Title Automation", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim processAllSheets As Boolean
                Dim modeChoice As DialogResult = MessageBox.Show(
                    "YES = Process all sheets" & vbCrLf & "NO = Process active sheet only" & vbCrLf & "CANCEL = Abort",
                    "Title Automation - Sheet Scope",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1)

                If modeChoice = DialogResult.Cancel Then
                    Return
                End If
                processAllSheets = (modeChoice = DialogResult.Yes)

                Dim mainAssemblyDoc As AssemblyDocument = Nothing
                Dim openedMainAssemblyPath As String = String.Empty

                Try
                    mainAssemblyDoc = GetMainAssemblyDocument(openedMainAssemblyPath)
                Catch
                    mainAssemblyDoc = Nothing
                End Try

                Dim totalUpdated As Integer = 0

                If processAllSheets Then
                    For Each sheet As Sheet In drawingDoc.Sheets
                        totalUpdated += UpdateSheetViewTitles(sheet, mainAssemblyDoc)
                    Next
                Else
                    Dim activeSheet As Sheet = drawingDoc.ActiveSheet
                    If activeSheet Is Nothing Then
                        MessageBox.Show("No active sheet found.", "Title Automation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    totalUpdated = UpdateSheetViewTitles(activeSheet, mainAssemblyDoc)
                End If

                Try
                    drawingDoc.Update2(True)
                Catch
                End Try

                MessageBox.Show("Updated " & totalUpdated.ToString() & " view titles.", "Title Automation", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If mainAssemblyDoc IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(openedMainAssemblyPath) Then
                    Try
                        mainAssemblyDoc.Close(True)
                    Catch
                    End Try
                End If

            Catch ex As Exception
                MessageBox.Show("Title Automation failed: " & ex.Message,
                                "Title Automation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function UpdateSheetViewTitles(ByVal sheet As Sheet, ByVal mainAssemblyDoc As AssemblyDocument) As Integer
            Dim updated As Integer = 0

            For Each view As DrawingView In sheet.DrawingViews
                If IsBaseView(view) Then
                    If UpdateBaseViewTitle(view, mainAssemblyDoc) Then
                        updated += 1
                    End If
                Else
                    If UpdateProjectedViewTitle(view) Then
                        updated += 1
                    End If
                End If
            Next

            Return updated
        End Function

        Private Function UpdateBaseViewTitle(ByVal view As DrawingView, ByVal mainAssemblyDoc As AssemblyDocument) As Boolean
            Try
                Dim refDoc As Document = Nothing
                Try
                    refDoc = view.ReferencedDocumentDescriptor.ReferencedDocument
                Catch
                    refDoc = Nothing
                End Try

                If refDoc Is Nothing Then
                    Return False
                End If

                Select Case refDoc.DocumentType
                    Case DocumentTypeEnum.kAssemblyDocumentObject
                        Dim quantity As Integer = GetAssemblyQuantity(TryCast(refDoc, AssemblyDocument), mainAssemblyDoc)
                        Return UpdateViewLabel(view, CreateAssemblyTitle(quantity))

                    Case DocumentTypeEnum.kPartDocumentObject
                        Return UpdateViewLabel(view, CreatePartTitle())

                    Case Else
                        Return False
                End Select

            Catch
                Return False
            End Try
        End Function

        Private Function UpdateProjectedViewTitle(ByVal view As DrawingView) As Boolean
            Dim firstLine As String = view.Name

            Try
                If view.Label IsNot Nothing Then
                    Dim labelText As String = view.Label.Text
                    If Not String.IsNullOrWhiteSpace(labelText) Then
                        firstLine = labelText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)(0)
                    End If
                End If
            Catch
            End Try

            firstLine = CleanProjectedTitleLine(firstLine)

            Dim title As String =
                "<StyleOverride FontSize='0.35' Bold='True' Underline='True'>" & firstLine & "</StyleOverride>" & vbCrLf &
                "<StyleOverride FontSize='0.25' Bold='False' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"

            Return UpdateViewLabel(view, title)
        End Function

        Private Function CleanProjectedTitleLine(ByVal value As String) As String
            Dim result As String = If(value, String.Empty).Trim()
            If result = String.Empty Then
                Return result
            End If

            result = Regex.Replace(result, "(?i)\s*SCALE\b.*$", String.Empty).Trim()
            result = Regex.Replace(result, "\s*\d+\s*[:\-]\s*\d+\s*$", String.Empty).Trim()
            result = Regex.Replace(result, "(?i)\s*VIEW\s*$", String.Empty).Trim()

            Return result
        End Function

        Private Function UpdateViewLabel(ByVal view As DrawingView, ByVal formattedTitle As String) As Boolean
            Try
                If Not view.ShowLabel Then
                    view.ShowLabel = True
                End If

                view.Label.FormattedText = formattedTitle
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Function IsBaseView(ByVal view As DrawingView) As Boolean
            Try
                Return view.ParentView Is Nothing
            Catch
                Return True
            End Try
        End Function

        Private Function GetAssemblyQuantity(ByVal assemblyDoc As AssemblyDocument, ByVal mainAssemblyDoc As AssemblyDocument) As Integer
            If assemblyDoc Is Nothing OrElse mainAssemblyDoc Is Nothing Then
                Return 1
            End If

            Try
                Dim count As Integer = 0
                For Each occ As ComponentOccurrence In mainAssemblyDoc.ComponentDefinition.Occurrences
                    Try
                        If String.Equals(occ.ReferencedFileDescriptor.FullFileName, assemblyDoc.FullFileName, StringComparison.OrdinalIgnoreCase) Then
                            count += 1
                        End If
                    Catch
                    End Try
                Next

                Return If(count > 0, count, 1)
            Catch
                Return 1
            End Try
        End Function

        Private Function CreateAssemblyTitle(ByVal quantity As Integer) As String
            Return "<StyleOverride FontSize='0.35' Bold='True' Underline='True'><Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride>" & vbCrLf &
                   "<StyleOverride FontSize='0.25' Bold='True' Underline='False'>" & quantity.ToString() & "-OFF REQ'D</StyleOverride>" & vbCrLf &
                   "<StyleOverride FontSize='0.25' Bold='True' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"
        End Function

        Private Function CreatePartTitle() As String
            Return "<StyleOverride FontSize='0.35' Bold='True' Underline='True'><Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride>" & vbCrLf &
                   "<StyleOverride FontSize='0.25' Bold='False' Underline='False'><Property Document='model' PropertySet='Design Tracking Properties' Property='Stock Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='55'>STOCK NUMBER</Property></StyleOverride>" & vbCrLf &
                   "<StyleOverride FontSize='0.25' Bold='False' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"
        End Function

        Private Function GetMainAssemblyDocument(ByRef openedMainAssemblyPath As String) As AssemblyDocument
            openedMainAssemblyPath = String.Empty

            Try
                Dim activeProject As DesignProject = m_InventorApp.DesignProjectManager.ActiveDesignProject
                If activeProject Is Nothing Then
                    Return Nothing
                End If

                For i As Integer = 1 To activeProject.LibraryPaths.Count
                    Dim projectPath As ProjectPath = activeProject.LibraryPaths.Item(i)
                    Dim searchPath As String = projectPath.Path
                    If String.IsNullOrWhiteSpace(searchPath) Then
                        Continue For
                    End If

                    If searchPath.IndexOf("bom", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Continue For
                    End If

                    Dim structurePath As String = System.IO.Path.Combine(searchPath, "Structure.iam")
                    If Not System.IO.File.Exists(structurePath) Then
                        Continue For
                    End If

                    Dim existing As Document = FindOpenDocument(structurePath)
                    If existing IsNot Nothing Then
                        Return TryCast(existing, AssemblyDocument)
                    End If

                    Dim opened As Document = m_InventorApp.Documents.Open(structurePath, False)
                    openedMainAssemblyPath = structurePath
                    Return TryCast(opened, AssemblyDocument)
                Next
            Catch
            End Try

            Return Nothing
        End Function

        Private Function FindOpenDocument(ByVal fullPath As String) As Document
            For Each doc As Document In m_InventorApp.Documents
                Try
                    If String.Equals(doc.FullFileName, fullPath, StringComparison.OrdinalIgnoreCase) Then
                        Return doc
                    End If
                Catch
                End Try
            Next

            Return Nothing
        End Function

        Private Function IsIdwDocument(ByVal drawingDoc As DrawingDocument) As Boolean
            Dim candidate As String = String.Empty
            Try
                candidate = drawingDoc.FullFileName
            Catch
                candidate = String.Empty
            End Try

            If String.IsNullOrWhiteSpace(candidate) Then
                Try
                    candidate = drawingDoc.DisplayName
                Catch
                    candidate = String.Empty
                End Try
            End If

            Return candidate.EndsWith(".idw", StringComparison.OrdinalIgnoreCase)
        End Function
    End Class

End Namespace