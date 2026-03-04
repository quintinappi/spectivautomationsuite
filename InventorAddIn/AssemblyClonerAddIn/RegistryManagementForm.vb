Imports Inventor
Imports Microsoft.Win32
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    Public Class RegistryManagementForm
        Inherits Form

        Private ReadOnly m_InventorApp As Inventor.Application

        Private lblAction As System.Windows.Forms.Label
        Private cmbAction As System.Windows.Forms.ComboBox
        Private lblPrefix As System.Windows.Forms.Label
        Private cmbPrefix As System.Windows.Forms.ComboBox
        Private lblGroups As System.Windows.Forms.Label
        Private chkGroups As System.Windows.Forms.CheckedListBox
        Private chkAllGroups As System.Windows.Forms.CheckBox
        Private btnRefresh As System.Windows.Forms.Button
        Private btnExecute As System.Windows.Forms.Button
        Private btnClose As System.Windows.Forms.Button
        Private txtReport As System.Windows.Forms.TextBox

        Private ReadOnly m_GroupCodes As String() = {"CH", "PL", "B", "A", "P", "SQ", "FL", "LPL", "IPE", "R", "OTHER", "PART", "FLG"}
        Private Const RegistryBasePath As String = "Software\InventorRenamer"
        Private Const AllPrefixesLabel As String = "(All Prefixes)"

        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            InitializeUi()
            LoadPrefixOptions()
            UpdateUiState()
        End Sub

        Private Sub InitializeUi()
            Me.Text = "Registry Management"
            Me.Width = 900
            Me.Height = 650
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.StartPosition = FormStartPosition.CenterScreen

            lblAction = New System.Windows.Forms.Label() With {.Text = "Action", .Left = 14, .Top = 14, .Width = 140}
            cmbAction = New System.Windows.Forms.ComboBox() With {.Left = 14, .Top = 34, .Width = 420, .DropDownStyle = ComboBoxStyle.DropDownList}
            cmbAction.Items.Add("Scan Registry Counters")
            cmbAction.Items.Add("Scan Open Assembly and Update Registry")
            cmbAction.Items.Add("Delete Selected Registry Counters")
            cmbAction.Items.Add("Delete ALL Registry Counters")
            cmbAction.SelectedIndex = 0
            AddHandler cmbAction.SelectedIndexChanged, AddressOf OnSelectionChanged

            lblPrefix = New System.Windows.Forms.Label() With {.Text = "Prefix", .Left = 14, .Top = 72, .Width = 140}
            cmbPrefix = New System.Windows.Forms.ComboBox() With {.Left = 14, .Top = 92, .Width = 420, .DropDownStyle = ComboBoxStyle.DropDownList}

            lblGroups = New System.Windows.Forms.Label() With {.Text = "Group Filters", .Left = 14, .Top = 130, .Width = 140}
            chkAllGroups = New System.Windows.Forms.CheckBox() With {.Left = 14, .Top = 150, .Width = 190, .Text = "Select all groups", .Checked = True}
            AddHandler chkAllGroups.CheckedChanged, AddressOf OnAllGroupsChanged

            chkGroups = New System.Windows.Forms.CheckedListBox() With {.Left = 14, .Top = 172, .Width = 420, .Height = 180, .CheckOnClick = True}
            For Each groupCode As String In m_GroupCodes
                chkGroups.Items.Add(groupCode, True)
            Next
            AddHandler chkGroups.ItemCheck, AddressOf OnGroupItemCheck

            btnRefresh = New System.Windows.Forms.Button() With {.Left = 14, .Top = 366, .Width = 130, .Height = 32, .Text = "Refresh Prefixes"}
            AddHandler btnRefresh.Click, AddressOf OnRefreshPrefixes

            btnExecute = New System.Windows.Forms.Button() With {.Left = 304, .Top = 366, .Width = 130, .Height = 32, .Text = "Run"}
            AddHandler btnExecute.Click, AddressOf OnExecuteClick

            btnClose = New System.Windows.Forms.Button() With {.Left = 14, .Top = 570, .Width = 130, .Height = 32, .Text = "Close"}
            AddHandler btnClose.Click, Sub() Me.Close()

            txtReport = New System.Windows.Forms.TextBox() With {
                .Left = 450,
                .Top = 14,
                .Width = 420,
                .Height = 588,
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .Font = New Drawing.Font("Consolas", 9.0F)
            }

            Me.Controls.Add(lblAction)
            Me.Controls.Add(cmbAction)
            Me.Controls.Add(lblPrefix)
            Me.Controls.Add(cmbPrefix)
            Me.Controls.Add(lblGroups)
            Me.Controls.Add(chkAllGroups)
            Me.Controls.Add(chkGroups)
            Me.Controls.Add(btnRefresh)
            Me.Controls.Add(btnExecute)
            Me.Controls.Add(btnClose)
            Me.Controls.Add(txtReport)
        End Sub

        Private Sub LoadPrefixOptions()
            Try
                Dim prefixes As List(Of String) = GetRegistryPrefixes()
                cmbPrefix.Items.Clear()
                cmbPrefix.Items.Add(AllPrefixesLabel)

                For Each prefix As String In prefixes
                    cmbPrefix.Items.Add(prefix)
                Next

                If cmbPrefix.Items.Count > 1 Then
                    cmbPrefix.SelectedIndex = 1
                Else
                    cmbPrefix.SelectedIndex = 0
                End If
            Catch ex As Exception
                AppendReport("Failed to load prefixes: " & ex.Message)
                WriteLog("LoadPrefixOptions", ex)
            End Try
        End Sub

        Private Sub OnSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
            UpdateUiState()
        End Sub

        Private Sub UpdateUiState()
            Dim actionText As String = SelectedAction()
            Dim usePrefix As Boolean = (actionText = "Scan Registry Counters" OrElse actionText = "Delete Selected Registry Counters")
            Dim useGroups As Boolean = usePrefix

            cmbPrefix.Enabled = usePrefix
            lblPrefix.Enabled = usePrefix
            chkAllGroups.Enabled = useGroups
            chkGroups.Enabled = useGroups
            lblGroups.Enabled = useGroups
        End Sub

        Private Sub OnAllGroupsChanged(ByVal sender As Object, ByVal e As EventArgs)
            For i As Integer = 0 To chkGroups.Items.Count - 1
                chkGroups.SetItemChecked(i, chkAllGroups.Checked)
            Next
        End Sub

        Private Sub OnGroupItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs)
            Me.BeginInvoke(New Action(Sub()
                                          Dim allChecked As Boolean = True
                                          For i As Integer = 0 To chkGroups.Items.Count - 1
                                              If Not chkGroups.GetItemChecked(i) Then
                                                  allChecked = False
                                                  Exit For
                                              End If
                                          Next
                                          chkAllGroups.Checked = allChecked
                                      End Sub))
        End Sub

        Private Sub OnRefreshPrefixes(ByVal sender As Object, ByVal e As EventArgs)
            LoadPrefixOptions()
            AppendReport("Refreshed prefix list.")
        End Sub

        Private Sub OnExecuteClick(ByVal sender As Object, ByVal e As EventArgs)
            Try
                Select Case SelectedAction()
                    Case "Scan Registry Counters"
                        ExecuteScanRegistry()
                    Case "Scan Open Assembly and Update Registry"
                        ExecuteScanProjectUpdateRegistry()
                    Case "Delete Selected Registry Counters"
                        ExecuteDeleteSelected()
                    Case "Delete ALL Registry Counters"
                        ExecuteDeleteAll()
                End Select
            Catch ex As Exception
                WriteLog("OnExecuteClick", ex)
                Dim logPath As String = GetLogPath()
                MessageBox.Show("Registry Management failed: " & ex.Message & vbCrLf & vbCrLf &
                                "Troubleshooting:" & vbCrLf &
                                "1) Confirm Inventor is running with proper document state." & vbCrLf &
                                "2) Confirm registry permissions for current user." & vbCrLf &
                                "3) Check log: " & logPath,
                                "Registry Management", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Function SelectedAction() As String
            If cmbAction.SelectedItem Is Nothing Then
                Return String.Empty
            End If
            Return cmbAction.SelectedItem.ToString()
        End Function

        Private Function SelectedGroups() As List(Of String)
            Dim values As New List(Of String)()
            For Each item As Object In chkGroups.CheckedItems
                values.Add(item.ToString())
            Next
            If values.Count = 0 Then
                values.AddRange(m_GroupCodes)
            End If
            Return values
        End Function

        Private Function SelectedPrefixOrDefault() As String
            If cmbPrefix.SelectedItem Is Nothing Then
                Return AllPrefixesLabel
            End If
            Return cmbPrefix.SelectedItem.ToString()
        End Function

        Private Sub ExecuteScanRegistry()
            Dim prefix As String = SelectedPrefixOrDefault()
            Dim groups As List(Of String) = SelectedGroups()

            Dim report As New StringBuilder()
            report.AppendLine("REGISTRY SCAN")
            report.AppendLine(New String("="c, 45))

            Using root As RegistryKey = Registry.CurrentUser.OpenSubKey(RegistryBasePath, False)
                If root Is Nothing Then
                    report.AppendLine("No registry entries found.")
                    AppendReport(report.ToString())
                    Return
                End If

                Dim values As String() = root.GetValueNames()
                Dim matches As New List(Of Tuple(Of String, Integer))()

                For Each valueName As String In values
                    If prefix <> AllPrefixesLabel AndAlso Not valueName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    Dim groupCode As String = ExtractGroupCode(valueName)
                    If groupCode <> String.Empty AndAlso Not groups.Contains(groupCode) Then
                        Continue For
                    End If

                    Dim valueObj As Object = root.GetValue(valueName)
                    Dim valueNumber As Integer
                    If valueObj IsNot Nothing AndAlso Integer.TryParse(valueObj.ToString(), valueNumber) Then
                        matches.Add(Tuple.Create(valueName, valueNumber))
                    End If
                Next

                If matches.Count = 0 Then
                    report.AppendLine("No matching counters found.")
                Else
                    report.AppendLine("Found " & matches.Count.ToString() & " counter(s):")
                    report.AppendLine()
                    For Each pair In matches.OrderBy(Function(x) x.Item1)
                        report.AppendLine(pair.Item1 & " = " & pair.Item2.ToString())
                    Next
                End If
            End Using

            AppendReport(report.ToString())
        End Sub

        Private Sub ExecuteScanProjectUpdateRegistry()
            If m_InventorApp Is Nothing OrElse m_InventorApp.ActiveDocument Is Nothing Then
                Throw New InvalidOperationException("Open an assembly in Inventor before running project scan.")
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                Throw New InvalidOperationException("Active document must be an assembly (.iam).")
            End If

            Dim asmDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
            Dim counters As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim detectedPrefix As String = String.Empty

            ProcessAssemblyOccurrences(asmDoc.ComponentDefinition.Occurrences, counters, detectedPrefix)

            If String.IsNullOrWhiteSpace(detectedPrefix) OrElse counters.Count = 0 Then
                AppendReport("No heritage part names were detected in the active assembly. Registry not updated.")
                Return
            End If

            Dim result As DialogResult = MessageBox.Show(
                "Detected prefix: " & detectedPrefix & vbCrLf &
                "Groups to update: " & counters.Count.ToString() & vbCrLf & vbCrLf &
                "Update registry now?",
                "Confirm Registry Update",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result <> DialogResult.Yes Then
                AppendReport("Registry update cancelled by user.")
                Return
            End If

            Using root As RegistryKey = Registry.CurrentUser.CreateSubKey(RegistryBasePath)
                For Each pair In counters
                    root.SetValue(detectedPrefix & pair.Key, pair.Value, RegistryValueKind.DWord)
                Next
            End Using

            Dim report As New StringBuilder()
            report.AppendLine("PROJECT SCAN + REGISTRY UPDATE")
            report.AppendLine(New String("="c, 45))
            report.AppendLine("Detected prefix: " & detectedPrefix)
            report.AppendLine("Updated groups: " & counters.Count.ToString())
            For Each pair In counters.OrderBy(Function(x) x.Key)
                report.AppendLine("  " & pair.Key & " = " & pair.Value.ToString())
            Next
            AppendReport(report.ToString())

            LoadPrefixOptions()
        End Sub

        Private Sub ExecuteDeleteSelected()
            Dim prefix As String = SelectedPrefixOrDefault()
            If prefix = AllPrefixesLabel Then
                MessageBox.Show("Select a specific prefix for this action.", "Registry Management", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim groups As List(Of String) = SelectedGroups()
            Dim confirm As DialogResult = MessageBox.Show(
                "Delete counters for prefix " & prefix & " for " & groups.Count.ToString() & " selected group(s)?",
                "Confirm Delete Selected",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)

            If confirm <> DialogResult.Yes Then
                AppendReport("Delete selected cancelled by user.")
                Return
            End If

            Dim deleted As Integer = 0
            Using root As RegistryKey = Registry.CurrentUser.OpenSubKey(RegistryBasePath, True)
                If root IsNot Nothing Then
                    For Each groupCode As String In groups
                        Dim keyName As String = prefix & groupCode
                        If root.GetValue(keyName) IsNot Nothing Then
                            root.DeleteValue(keyName, False)
                            deleted += 1
                        End If
                    Next
                End If
            End Using

            AppendReport("Deleted " & deleted.ToString() & " entries for prefix " & prefix & ".")
            LoadPrefixOptions()
        End Sub

        Private Sub ExecuteDeleteAll()
            Dim confirm As DialogResult = MessageBox.Show(
                "Delete ALL Inventor Renamer counters from registry? This cannot be undone.",
                "Confirm Delete ALL",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)

            If confirm <> DialogResult.Yes Then
                AppendReport("Delete all cancelled by user.")
                Return
            End If

            Registry.CurrentUser.DeleteSubKeyTree(RegistryBasePath, False)
            AppendReport("All Inventor Renamer counters were deleted.")
            LoadPrefixOptions()
        End Sub

        Private Sub ProcessAssemblyOccurrences(ByVal occurrences As ComponentOccurrences,
                                              ByVal counters As Dictionary(Of String, Integer),
                                              ByRef detectedPrefix As String)
            For Each occurrence As ComponentOccurrence In occurrences
                If occurrence.Suppressed Then
                    Continue For
                End If

                Dim doc As Document = Nothing
                Try
                    doc = occurrence.Definition.Document
                Catch
                    doc = Nothing
                End Try
                If doc Is Nothing Then
                    Continue For
                End If

                Dim fullPath As String = String.Empty
                Try
                    fullPath = doc.FullFileName
                Catch
                    fullPath = String.Empty
                End Try
                If fullPath.IndexOf("\oldversions\", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Continue For
                End If

                Dim extension As String = System.IO.Path.GetExtension(fullPath)
                If String.Equals(extension, ".ipt", StringComparison.OrdinalIgnoreCase) Then
                    ParseHeritageFileName(System.IO.Path.GetFileNameWithoutExtension(fullPath), counters, detectedPrefix)
                ElseIf String.Equals(extension, ".iam", StringComparison.OrdinalIgnoreCase) Then
                    Dim childAsm As AssemblyDocument = TryCast(doc, AssemblyDocument)
                    If childAsm IsNot Nothing Then
                        ProcessAssemblyOccurrences(childAsm.ComponentDefinition.Occurrences, counters, detectedPrefix)
                    End If
                End If
            Next
        End Sub

        Private Sub ParseHeritageFileName(ByVal baseName As String,
                                          ByVal counters As Dictionary(Of String, Integer),
                                          ByRef detectedPrefix As String)
            Dim match As Match = Regex.Match(baseName, "^(?<prefix>.+-)(?<group>[A-Za-z]+)(?<num>\d+)$")
            If Not match.Success Then
                Return
            End If

            Dim prefix As String = match.Groups("prefix").Value
            Dim groupCode As String = match.Groups("group").Value.ToUpperInvariant()
            Dim numberValue As Integer = Integer.Parse(match.Groups("num").Value)

            If String.IsNullOrWhiteSpace(detectedPrefix) Then
                detectedPrefix = prefix
            ElseIf Not String.Equals(detectedPrefix, prefix, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            If counters.ContainsKey(groupCode) Then
                If numberValue > counters(groupCode) Then
                    counters(groupCode) = numberValue
                End If
            Else
                counters.Add(groupCode, numberValue)
            End If
        End Sub

        Private Function GetRegistryPrefixes() As List(Of String)
            Dim prefixes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Using root As RegistryKey = Registry.CurrentUser.OpenSubKey(RegistryBasePath, False)
                If root IsNot Nothing Then
                    For Each valueName As String In root.GetValueNames()
                        Dim dashIndex As Integer = valueName.LastIndexOf("-"c)
                        If dashIndex > 0 Then
                            prefixes.Add(valueName.Substring(0, dashIndex + 1))
                        End If
                    Next
                End If
            End Using

            If prefixes.Count = 0 Then
                prefixes.Add("NCRH01-000-")
            End If

            Return prefixes.OrderBy(Function(x) x).ToList()
        End Function

        Private Function ExtractGroupCode(ByVal registryKeyName As String) As String
            Dim lastDash As Integer = registryKeyName.LastIndexOf("-"c)
            If lastDash < 0 OrElse lastDash = registryKeyName.Length - 1 Then
                Return String.Empty
            End If
            Return registryKeyName.Substring(lastDash + 1).ToUpperInvariant()
        End Function

        Private Sub AppendReport(ByVal message As String)
            If String.IsNullOrWhiteSpace(message) Then
                Return
            End If

            txtReport.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - " & message & vbCrLf & vbCrLf)
        End Sub

        Private Sub WriteLog(ByVal context As String, ByVal ex As Exception)
            Try
                Dim logPath As String = GetLogPath()
                Dim logDir As String = System.IO.Path.GetDirectoryName(logPath)
                If Not System.IO.Directory.Exists(logDir) Then
                    System.IO.Directory.CreateDirectory(logDir)
                End If

                System.IO.File.AppendAllText(logPath,
                                   DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | " & context & " | " & ex.ToString() & System.Environment.NewLine)
            Catch
            End Try
        End Sub

        Private Function GetLogPath() As String
            Dim baseDir As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Spectiv", "InventorAutomationSuite", "Logs")
            Return System.IO.Path.Combine(baseDir, "RegistryManagement.log")
        End Function

    End Class

End Namespace
