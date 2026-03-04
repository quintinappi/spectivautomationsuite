Imports Inventor
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class ViewIdentifierSetter
        Private ReadOnly m_InventorApp As Inventor.Application

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Open an IDW drawing first.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
            Dim sheet As Sheet = drawingDoc.ActiveSheet
            If sheet Is Nothing OrElse sheet.DrawingViews.Count = 0 Then
                MessageBox.Show("No drawing views found on the active sheet.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim picked As Object = Nothing
            Try
                picked = m_InventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a view to set View Identifier")
            Catch
                picked = Nothing
            End Try

            Dim targetView As DrawingView = TryCast(picked, DrawingView)
            If targetView Is Nothing Then
                MessageBox.Show("No drawing view selected.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim currentIdentifier As String = GetViewIdentifier(targetView)
            Dim desiredIdentifier As String = ShowIdentifierPrompt(currentIdentifier)
            If String.IsNullOrWhiteSpace(desiredIdentifier) Then
                Return
            End If

            desiredIdentifier = NormalizeDisplayIdentifier(desiredIdentifier)
            If String.IsNullOrWhiteSpace(desiredIdentifier) Then
                MessageBox.Show("Enter a valid View Identifier.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim finalIdentifier As String = ResolveUniqueIdentifier(sheet, targetView, desiredIdentifier)

            Try
                targetView.Name = finalIdentifier
            Catch ex As Exception
                MessageBox.Show("Could not set view identifier: " & ex.Message, "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ApplyDisplayedIdentifierLabel(targetView, desiredIdentifier)

            Try
                drawingDoc.Update2(True)
            Catch
            End Try

            If String.Equals(finalIdentifier, desiredIdentifier, StringComparison.OrdinalIgnoreCase) Then
                MessageBox.Show("Displayed title set to '" & desiredIdentifier & " VIEW'.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("'" & desiredIdentifier & "' was already used on this sheet." & vbCrLf &
                                "Displayed title stays as '" & desiredIdentifier & " VIEW'." & vbCrLf &
                                "Inventor internal identifier was set to '" & finalIdentifier & "' to keep names unique.",
                                "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Sub

        Private Sub ApplyDisplayedIdentifierLabel(ByVal view As DrawingView, ByVal desiredIdentifier As String)
            If view Is Nothing Then
                Return
            End If

            Dim normalizedIdentifier As String = NormalizeDisplayIdentifier(desiredIdentifier)
            If String.IsNullOrWhiteSpace(normalizedIdentifier) Then
                Return
            End If

            Try
                If Not view.ShowLabel Then
                    view.ShowLabel = True
                End If

                view.Label.FormattedText = BuildDisplayLabelFormattedText(normalizedIdentifier)
            Catch
            End Try
        End Sub

        Private Function NormalizeDisplayIdentifier(ByVal value As String) As String
            Dim result As String = If(value, String.Empty).Trim()
            If result = String.Empty Then
                Return result
            End If

            result = Regex.Replace(result, "(?i)\s*VIEW\s*$", String.Empty).Trim()
            Return result.ToUpperInvariant()
        End Function

        Private Function BuildDisplayLabelFormattedText(ByVal normalizedIdentifier As String) As String
            Dim escapedIdentifier As String = EscapeFormattedTextValue(normalizedIdentifier)

            Return "<StyleOverride FontSize='0.35' Bold='True' Underline='True'>" & escapedIdentifier & " VIEW</StyleOverride>" & vbCrLf &
                   "<StyleOverride FontSize='0.25' Bold='False' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"
        End Function

        Private Function EscapeFormattedTextValue(ByVal value As String) As String
            Dim escaped As String = If(value, String.Empty)
            escaped = escaped.Replace("&", "&amp;")
            escaped = escaped.Replace("<", "&lt;")
            escaped = escaped.Replace(">", "&gt;")
            escaped = escaped.Replace("""", "&quot;")
            escaped = escaped.Replace("'", "&apos;")
            Return escaped
        End Function

        Private Function GetViewIdentifier(ByVal view As DrawingView) As String
            If view Is Nothing Then
                Return String.Empty
            End If

            Try
                Return If(view.Name, String.Empty)
            Catch
                Return String.Empty
            End Try
        End Function

        Private Function ResolveUniqueIdentifier(ByVal sheet As Sheet, ByVal targetView As DrawingView, ByVal baseIdentifier As String) As String
            Dim normalizedBase As String = baseIdentifier.Trim()
            If normalizedBase = String.Empty Then
                Return normalizedBase
            End If

            Dim usedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each candidateView As DrawingView In sheet.DrawingViews
                If candidateView Is Nothing Then
                    Continue For
                End If

                If candidateView Is targetView Then
                    Continue For
                End If

                Try
                    Dim nameValue As String = candidateView.Name
                    If Not String.IsNullOrWhiteSpace(nameValue) Then
                        usedNames.Add(nameValue.Trim())
                    End If
                Catch
                End Try
            Next

            If Not usedNames.Contains(normalizedBase) Then
                Return normalizedBase
            End If

            For suffix As Integer = 2 To 999
                Dim candidateName As String = normalizedBase & "_" & suffix.ToString()
                If Not usedNames.Contains(candidateName) Then
                    Return candidateName
                End If
            Next

            Return normalizedBase & "_" & DateTime.Now.ToString("HHmmss")
        End Function

        Private Function ShowIdentifierPrompt(ByVal currentIdentifier As String) As String
            Using dialog As New ViewIdentifierDialog(currentIdentifier)
                If dialog.ShowDialog() = DialogResult.OK Then
                    Return dialog.SelectedIdentifier
                End If
            End Using

            Return String.Empty
        End Function

        Private Class ViewIdentifierDialog
            Inherits Form

            Private ReadOnly m_PresetCombo As ComboBox
            Private ReadOnly m_CustomTextBox As System.Windows.Forms.TextBox
            Private m_SelectedIdentifier As String

            Public ReadOnly Property SelectedIdentifier As String
                Get
                    Return m_SelectedIdentifier
                End Get
            End Property

            Public Sub New(ByVal currentIdentifier As String)
                Me.Text = "Set View Identifier"
                Me.Width = 420
                Me.Height = 220
                Me.StartPosition = FormStartPosition.CenterParent
                Me.FormBorderStyle = FormBorderStyle.FixedDialog
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.AutoScaleMode = AutoScaleMode.Font

                Dim promptLabel As New Label() With {
                    .Left = 12,
                    .Top = 12,
                    .Width = 380,
                    .Text = "Choose a preset or enter a custom View Identifier:"
                }

                Dim presetLabel As New Label() With {
                    .Left = 12,
                    .Top = 44,
                    .Width = 90,
                    .Text = "Preset"
                }

                m_PresetCombo = New ComboBox() With {
                    .Left = 108,
                    .Top = 40,
                    .Width = 280,
                    .DropDownStyle = ComboBoxStyle.DropDownList
                }
                m_PresetCombo.Items.AddRange(New Object() {"FRONT", "TOP", "SIDE", "BOTTOM", "Custom"})
                AddHandler m_PresetCombo.SelectedIndexChanged, AddressOf OnPresetChanged

                Dim customLabel As New Label() With {
                    .Left = 12,
                    .Top = 78,
                    .Width = 90,
                    .Text = "Custom"
                }

                m_CustomTextBox = New System.Windows.Forms.TextBox() With {
                    .Left = 108,
                    .Top = 74,
                    .Width = 280
                }

                Dim currentNormalized As String = If(currentIdentifier, String.Empty).Trim()
                Dim presetNames As String() = {"FRONT", "TOP", "SIDE", "BOTTOM"}
                Dim matchedPreset As String = presetNames.FirstOrDefault(Function(nameValue) String.Equals(nameValue, currentNormalized, StringComparison.OrdinalIgnoreCase))

                If String.IsNullOrWhiteSpace(matchedPreset) Then
                    m_PresetCombo.SelectedItem = "Custom"
                    m_CustomTextBox.Text = currentNormalized.ToUpperInvariant()
                Else
                    m_PresetCombo.SelectedItem = matchedPreset
                    m_CustomTextBox.Text = matchedPreset
                End If

                Dim okButton As New Button() With {
                    .Text = "Apply",
                    .Left = 226,
                    .Top = 118,
                    .Width = 78
                }
                AddHandler okButton.Click, AddressOf OnApply

                Dim cancelButton As New Button() With {
                    .Text = "Cancel",
                    .Left = 310,
                    .Top = 118,
                    .Width = 78
                }
                AddHandler cancelButton.Click, Sub() Me.DialogResult = DialogResult.Cancel

                Me.AcceptButton = okButton
                Me.CancelButton = cancelButton

                Me.Controls.Add(promptLabel)
                Me.Controls.Add(presetLabel)
                Me.Controls.Add(m_PresetCombo)
                Me.Controls.Add(customLabel)
                Me.Controls.Add(m_CustomTextBox)
                Me.Controls.Add(okButton)
                Me.Controls.Add(cancelButton)
            End Sub

            Private Sub OnPresetChanged(ByVal sender As Object, ByVal e As EventArgs)
                Dim selectedPreset As String = Convert.ToString(m_PresetCombo.SelectedItem)
                Dim isCustom As Boolean = String.Equals(selectedPreset, "Custom", StringComparison.OrdinalIgnoreCase)

                m_CustomTextBox.Enabled = isCustom
                If Not isCustom AndAlso Not String.IsNullOrWhiteSpace(selectedPreset) Then
                    m_CustomTextBox.Text = selectedPreset.ToUpperInvariant()
                End If
            End Sub

            Private Sub OnApply(ByVal sender As Object, ByVal e As EventArgs)
                Dim selectedPreset As String = Convert.ToString(m_PresetCombo.SelectedItem)
                Dim candidate As String = If(String.Equals(selectedPreset, "Custom", StringComparison.OrdinalIgnoreCase),
                                             m_CustomTextBox.Text,
                                             selectedPreset)

                candidate = If(candidate, String.Empty).Trim()
                If candidate = String.Empty Then
                    MessageBox.Show("Enter a View Identifier.", "Set View Identifier", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                m_SelectedIdentifier = candidate.ToUpperInvariant()
                Me.DialogResult = DialogResult.OK
            End Sub
        End Class
    End Class

End Namespace
