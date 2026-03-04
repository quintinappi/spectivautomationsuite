Imports Inventor
Imports System.Globalization
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class ViewScaleSetter
        Private ReadOnly m_InventorApp As Inventor.Application

        Private Shared ReadOnly ScaleOptions As String() = {
            "1:1", "1:2", "1:3", "1:5", "1:75", "1:10", "1:12.5", "1:15", "1:20", "1:25", "Custom"
        }

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        Public Sub Execute()
            If m_InventorApp.ActiveDocument Is Nothing OrElse m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Open an IDW drawing first.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedScale As String = ShowScalePrompt()
            If String.IsNullOrWhiteSpace(selectedScale) Then
                Return
            End If

            Dim drawingDoc As DrawingDocument = CType(m_InventorApp.ActiveDocument, DrawingDocument)
            Dim sheet As Sheet = drawingDoc.ActiveSheet
            If sheet Is Nothing OrElse sheet.DrawingViews.Count = 0 Then
                MessageBox.Show("No drawing views found on the active sheet.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim picked As Object = Nothing
            Try
                picked = m_InventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a view to set scale " & selectedScale)
            Catch
                picked = Nothing
            End Try

            Dim targetView As DrawingView = TryCast(picked, DrawingView)
            If targetView Is Nothing Then
                MessageBox.Show("No drawing view selected.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Try
                targetView.ScaleFromBase = False
            Catch
            End Try

            Dim appliedScale As String = selectedScale
            If Not TryApplyScale(targetView, selectedScale, appliedScale) Then
                MessageBox.Show("Could not set scale '" & selectedScale & "' on the selected view.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Try
                drawingDoc.Update2(True)
            Catch
            End Try

            MessageBox.Show("View scale set to '" & appliedScale & "'.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        Private Function TryApplyScale(ByVal targetView As DrawingView, ByVal scaleText As String, ByRef appliedScale As String) As Boolean
            If targetView Is Nothing Then
                Return False
            End If

            Dim normalized As String = NormalizeScaleString(scaleText)
            If String.IsNullOrWhiteSpace(normalized) Then
                Return False
            End If

            Try
                targetView.ScaleString = normalized
                appliedScale = normalized
                Return True
            Catch
            End Try

            Dim numericScale As Double
            If TryParseScaleRatio(normalized, numericScale) Then
                Try
                    targetView.Scale = numericScale
                    appliedScale = normalized
                    Return True
                Catch
                End Try
            End If

            Return False
        End Function

        Private Function ShowScalePrompt() As String
            Using dialog As New ViewScaleDialog()
                If dialog.ShowDialog() = DialogResult.OK Then
                    Return dialog.SelectedScale
                End If
            End Using

            Return String.Empty
        End Function

        Private Shared Function NormalizeScaleString(ByVal value As String) As String
            Dim result As String = If(value, String.Empty).Trim()
            If result = String.Empty Then
                Return result
            End If

            result = result.Replace(" ", String.Empty)
            result = result.Replace(",", ".")
            result = result.Replace("/", ":")

            Return result
        End Function

        Private Shared Function IsValidScaleString(ByVal scaleText As String) As Boolean
            Dim ratio As Double
            Return TryParseScaleRatio(scaleText, ratio)
        End Function

        Private Shared Function TryParseScaleRatio(ByVal scaleText As String, ByRef ratio As Double) As Boolean
            ratio = 0.0

            If String.IsNullOrWhiteSpace(scaleText) Then
                Return False
            End If

            Dim parts As String() = scaleText.Split(":"c)
            If parts Is Nothing OrElse parts.Length <> 2 Then
                Return False
            End If

            Dim modelValue As Double
            Dim paperValue As Double

            If Not Double.TryParse(parts(0), NumberStyles.Float, CultureInfo.InvariantCulture, modelValue) Then
                Return False
            End If

            If Not Double.TryParse(parts(1), NumberStyles.Float, CultureInfo.InvariantCulture, paperValue) Then
                Return False
            End If

            If modelValue <= 0 OrElse paperValue <= 0 Then
                Return False
            End If

            ratio = modelValue / paperValue
            Return True
        End Function

        Private Class ViewScaleDialog
            Inherits Form

            Private ReadOnly m_ScaleCombo As ComboBox
            Private ReadOnly m_CustomScaleTextBox As System.Windows.Forms.TextBox
            Private m_SelectedScale As String

            Public ReadOnly Property SelectedScale As String
                Get
                    Return m_SelectedScale
                End Get
            End Property

            Public Sub New()
                Me.Text = "Set View Scale"
                Me.Width = 430
                Me.Height = 220
                Me.StartPosition = FormStartPosition.CenterParent
                Me.FormBorderStyle = FormBorderStyle.FixedDialog
                Me.MaximizeBox = False
                Me.MinimizeBox = False
                Me.AutoScaleMode = AutoScaleMode.Font

                Dim promptLabel As New Label() With {
                    .Left = 12,
                    .Top = 12,
                    .Width = 392,
                    .Text = "Select scale first, then click the target drawing view:"
                }

                Dim scaleLabel As New Label() With {
                    .Left = 12,
                    .Top = 44,
                    .Width = 96,
                    .Text = "Scale"
                }

                m_ScaleCombo = New ComboBox() With {
                    .Left = 108,
                    .Top = 40,
                    .Width = 296,
                    .DropDownStyle = ComboBoxStyle.DropDownList
                }
                m_ScaleCombo.Items.AddRange(ScaleOptions)
                m_ScaleCombo.SelectedIndex = 0
                AddHandler m_ScaleCombo.SelectedIndexChanged, AddressOf OnScaleSelectionChanged

                Dim customLabel As New Label() With {
                    .Left = 12,
                    .Top = 78,
                    .Width = 96,
                    .Text = "Custom"
                }

                m_CustomScaleTextBox = New System.Windows.Forms.TextBox() With {
                    .Left = 108,
                    .Top = 74,
                    .Width = 296,
                    .Enabled = False
                }

                Dim okButton As New Button() With {
                    .Text = "Pick View",
                    .Left = 240,
                    .Top = 118,
                    .Width = 78
                }
                AddHandler okButton.Click, AddressOf OnApply

                Dim cancelButton As New Button() With {
                    .Text = "Cancel",
                    .Left = 326,
                    .Top = 118,
                    .Width = 78
                }
                AddHandler cancelButton.Click, Sub() Me.DialogResult = DialogResult.Cancel

                Me.AcceptButton = okButton
                Me.CancelButton = cancelButton

                Me.Controls.Add(promptLabel)
                Me.Controls.Add(scaleLabel)
                Me.Controls.Add(m_ScaleCombo)
                Me.Controls.Add(customLabel)
                Me.Controls.Add(m_CustomScaleTextBox)
                Me.Controls.Add(okButton)
                Me.Controls.Add(cancelButton)
            End Sub

            Private Sub OnScaleSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
                Dim selectedScale As String = Convert.ToString(m_ScaleCombo.SelectedItem)
                Dim isCustom As Boolean = String.Equals(selectedScale, "Custom", StringComparison.OrdinalIgnoreCase)

                m_CustomScaleTextBox.Enabled = isCustom
                If Not isCustom AndAlso Not String.IsNullOrWhiteSpace(selectedScale) Then
                    m_CustomScaleTextBox.Text = selectedScale
                End If
            End Sub

            Private Sub OnApply(ByVal sender As Object, ByVal e As EventArgs)
                Dim selectedScale As String = Convert.ToString(m_ScaleCombo.SelectedItem)
                Dim candidate As String = If(String.Equals(selectedScale, "Custom", StringComparison.OrdinalIgnoreCase),
                                             m_CustomScaleTextBox.Text,
                                             selectedScale)

                candidate = NormalizeScaleString(candidate)
                If String.IsNullOrWhiteSpace(candidate) Then
                    MessageBox.Show("Select or enter a scale.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If Not IsValidScaleString(candidate) Then
                    MessageBox.Show("Scale format must be like 1:2 or 1:12.5.", "Set View Scale", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                m_SelectedScale = candidate
                Me.DialogResult = DialogResult.OK
            End Sub
        End Class
    End Class

End Namespace
