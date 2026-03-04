Imports System.Windows.Forms
Imports System.Drawing
Imports Inventor
Imports System.Collections.Generic

Namespace AssemblyClonerAddIn

    Public Class AssemblyInspectorForm
        Inherits Form

        Private ReadOnly m_AsmDoc As AssemblyDocument
        Private ReadOnly m_DocResults As List(Of DocScanResult)

        ' Controls
        Private lblTitle As Label
        Private picThumb As PictureBox
        Private lstParts As ListBox
        Private lvParams As ListView
        Private btnEditParam As Button
        Private txtNewValue As System.Windows.Forms.TextBox
        Private btnSaveAll As Button
        Private txtRules As System.Windows.Forms.TextBox
        Private lblRules As Label

        Public Sub New(ByVal asmDoc As AssemblyDocument, ByVal docResults As List(Of DocScanResult))
            Me.m_AsmDoc = asmDoc
            Me.m_DocResults = docResults

            InitializeComponents()
            PopulateData()
        End Sub

        Private Sub InitializeComponents()
            Me.Text = "Assembly Smart Inspector - Experimental"
            Me.Size = New Size(900, 600)
            Me.StartPosition = FormStartPosition.CenterParent

            lblTitle = New Label()
            lblTitle.AutoSize = True
            lblTitle.Location = New System.Drawing.Point(12, 12)
            lblTitle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            lblTitle.Text = m_AsmDoc.DisplayName
            Me.Controls.Add(lblTitle)

            picThumb = New PictureBox()
            picThumb.Size = New Size(240, 180)
            picThumb.Location = New System.Drawing.Point(12, 40)
            picThumb.BorderStyle = BorderStyle.FixedSingle
            Me.Controls.Add(picThumb)

            lstParts = New ListBox()
            lstParts.Location = New System.Drawing.Point(260, 40)
            lstParts.Size = New Size(300, 180)
            lstParts.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            lstParts.DisplayMember = "DocName"
            AddHandler lstParts.SelectedIndexChanged, AddressOf OnPartSelected
            Me.Controls.Add(lstParts)

            lvParams = New ListView()
            lvParams.Location = New System.Drawing.Point(12, 230)
            lvParams.Size = New Size(540, 280)
            lvParams.View = System.Windows.Forms.View.Details
            lvParams.FullRowSelect = True
            lvParams.Columns.Add("Param", 240)
            lvParams.Columns.Add("Document", 200)
            lvParams.Columns.Add("Expression", 100)
            lvParams.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left
            Me.Controls.Add(lvParams)

            ' Rules textbox to the right of params
            lblRules = New Label()
            lblRules.AutoSize = True
            lblRules.Location = New System.Drawing.Point(560, 210)
            lblRules.Text = "iLogic Rules" 
            Me.Controls.Add(lblRules)

            txtRules = New System.Windows.Forms.TextBox()
            txtRules.Location = New System.Drawing.Point(560, 230)
            txtRules.Size = New Size(300, 280)
            txtRules.Multiline = True
            txtRules.ScrollBars = ScrollBars.Both
            txtRules.WordWrap = False
            txtRules.ReadOnly = True
            txtRules.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Left
            Me.Controls.Add(txtRules)

            txtNewValue = New System.Windows.Forms.TextBox()
            txtNewValue.Location = New System.Drawing.Point(12, 520)
            txtNewValue.Size = New Size(300, 22)
            Me.Controls.Add(txtNewValue)

            btnEditParam = New Button()
            btnEditParam.Location = New System.Drawing.Point(320, 518)
            btnEditParam.Size = New Size(120, 26)
            btnEditParam.Text = "Set Selected Param"
            AddHandler btnEditParam.Click, AddressOf OnSetParamClicked
            Me.Controls.Add(btnEditParam)

            btnSaveAll = New Button()
            btnSaveAll.Location = New System.Drawing.Point(460, 518)
            btnSaveAll.Size = New Size(120, 26)
            btnSaveAll.Text = "Save All"
            AddHandler btnSaveAll.Click, AddressOf OnSaveAllClicked
            Me.Controls.Add(btnSaveAll)
        End Sub

        Private Sub PopulateData()
            Try
                ' Thumbnail if available
                Try
                    If m_AsmDoc.Thumbnail IsNot Nothing Then
                        Dim img As Image = CType(m_AsmDoc.Thumbnail, Image)
                        picThumb.Image = New Bitmap(img)
                        picThumb.SizeMode = PictureBoxSizeMode.Zoom
                    End If
                Catch
                    ' ignore thumbnails
                End Try

                ' Populate parts list
                For Each dr As DocScanResult In m_DocResults
                    lstParts.Items.Add(dr)
                Next

                ' Populate parameters (default: show all)
                For Each dr As DocScanResult In m_DocResults
                    For Each p As ParamInfo In dr.Params
                        Dim lvi As New ListViewItem(p.Name)
                        lvi.SubItems.Add(dr.DocName)
                        lvi.SubItems.Add(p.Expression)
                        ' Store the ParamInfo object in Tag so we can set it later
                        lvi.Tag = p
                        lvParams.Items.Add(lvi)
                    Next
                Next

                ' Select the first document to show details
                If lstParts.Items.Count > 0 Then
                    lstParts.SelectedIndex = 0
                End If

            Catch ex As Exception
                MessageBox.Show("Error populating inspector data: " & ex.Message)
            End Try
        End Sub

        Private Sub OnSetParamClicked(ByVal sender As Object, ByVal e As EventArgs)
            If lvParams.SelectedItems.Count = 0 Then
                MessageBox.Show("Select a parameter to edit.")
                Return
            End If

            Dim lvi As ListViewItem = lvParams.SelectedItems(0)
            If lvi Is Nothing Then Return
            Dim p As Object = lvi.Tag
            If p Is Nothing Then Return

            Dim newExpr As String = txtNewValue.Text.Trim()
            If String.IsNullOrEmpty(newExpr) Then
                MessageBox.Show("Enter a new expression, e.g. '100 mm' or '200'.")
                Return
            End If

            Try
                ' Set expression on the underlying UserParameter
                If p.ParamObject IsNot Nothing Then
                    Dim up As UserParameter = CType(p.ParamObject, UserParameter)
                    up.Expression = newExpr
                    lvi.SubItems(2).Text = newExpr
                    MessageBox.Show("Parameter updated in memory. Remember to save the document.")
                Else
                    MessageBox.Show("Selected parameter does not appear to be a user parameter.")
                End If
            Catch ex As Exception
                MessageBox.Show("Error setting parameter: " & ex.Message)
            End Try
        End Sub

        Private Sub OnSaveAllClicked(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Dim savedCount As Integer = 0
                For Each dr As DocScanResult In m_DocResults
                    Try
                        dr.Doc.Save()
                        savedCount += 1
                    Catch
                        ' ignore save errors
                    End Try
                Next

                MessageBox.Show("Saved " & savedCount & " documents.")
            Catch ex As Exception
                MessageBox.Show("Error saving documents: " & ex.Message)
            End Try
        End Sub

        Private Sub OnPartSelected(ByVal sender As Object, ByVal e As EventArgs)
            Try
                If lstParts.SelectedItem Is Nothing Then Return
                Dim dr As DocScanResult = CType(lstParts.SelectedItem, DocScanResult)

                ' Update params list to show only this document's params
                lvParams.Items.Clear()
                For Each p As ParamInfo In dr.Params
                    Dim lvi As New ListViewItem(p.Name)
                    lvi.SubItems.Add(dr.DocName)
                    lvi.SubItems.Add(p.Expression)
                    lvi.Tag = p
                    lvParams.Items.Add(lvi)
                Next

                ' Show rules for this document
                If dr.Rules IsNot Nothing AndAlso dr.Rules.Count > 0 Then
                    Dim sb As New System.Text.StringBuilder()
                    For Each kvp As KeyValuePair(Of String, String) In dr.Rules
                        sb.AppendLine("Rule: " & kvp.Key)
                        sb.AppendLine(New String("-"c, 40))
                        sb.AppendLine(kvp.Value)
                        sb.AppendLine()
                    Next
                    txtRules.Text = sb.ToString()
                Else
                    txtRules.Text = "(No iLogic rules found)"
                End If

            Catch ex As Exception
                MessageBox.Show("Error selecting part: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace
