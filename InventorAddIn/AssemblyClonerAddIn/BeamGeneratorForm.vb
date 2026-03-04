' ==============================================================================
' BEAM GENERATOR FORM - User Interface for Beam Assembly Generator
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' Windows Forms dialog for selecting:
'   - Section type (UC, UB, PFC, TFC, L, etc.)
'   - Section size (dynamically populated based on type)
'   - Beam length
'   - Output folder
'   - Part number prefix
'
' ==============================================================================

Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Form for configuring beam assembly generation
    ''' </summary>
    Public Class BeamGeneratorForm
        Inherits Form

        ' Controls
        Private WithEvents lblTitle As Label
        Private WithEvents lblSectionType As Label
        Private WithEvents cboSectionType As ComboBox
        Private WithEvents lblSectionSize As Label
        Private WithEvents cboSectionSize As ComboBox
        Private WithEvents lblBeamLength As Label
        Private WithEvents txtBeamLength As TextBox
        Private WithEvents lblLengthUnit As Label
        Private WithEvents lblOutputFolder As Label
        Private WithEvents txtOutputFolder As TextBox
        Private WithEvents btnBrowseFolder As Button
        Private WithEvents lblPrefix As Label
        Private WithEvents txtPrefix As TextBox
        Private WithEvents lblPreview As Label
        Private WithEvents txtPreview As TextBox
        Private WithEvents grpSectionInfo As GroupBox
        Private WithEvents lblSectionDetails As Label
        Private WithEvents btnOK As Button
        Private WithEvents btnCancel As Button

        ' Properties
        Public Property SelectedSection As SteelSection
        Public Property BeamLength As Double
        Public Property OutputFolder As String
        Public Property PartNumberPrefix As String

        ' Flag to prevent events during initialization
        Private m_IsInitializing As Boolean = True

        Public Sub New()
            m_IsInitializing = True
            InitializeComponent()
            LoadSectionTypes()
            m_IsInitializing = False
            ' Update displays after initialization
            UpdateSectionDetails()
            UpdatePreview()
        End Sub

        ''' <summary>
        ''' Initialize form controls
        ''' </summary>
        Private Sub InitializeComponent()
            Me.Text = "Beam Assembly Generator"
            Me.Size = New Size(550, 550)
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.BackColor = Color.White

            Dim yPos As Integer = 15

            ' Title
            lblTitle = New Label()
            lblTitle.Text = "PARAMETRIC BEAM ASSEMBLY GENERATOR"
            lblTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblTitle.Location = New Point(15, yPos)
            lblTitle.Size = New Size(500, 25)
            lblTitle.ForeColor = Color.FromArgb(33, 150, 243)
            Me.Controls.Add(lblTitle)
            yPos += 35

            ' Section Type
            lblSectionType = New Label()
            lblSectionType.Text = "Section Type:"
            lblSectionType.Location = New Point(15, yPos)
            lblSectionType.Size = New Size(100, 20)
            lblSectionType.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblSectionType)

            cboSectionType = New ComboBox()
            cboSectionType.Location = New Point(130, yPos - 3)
            cboSectionType.Size = New Size(150, 25)
            cboSectionType.DropDownStyle = ComboBoxStyle.DropDownList
            cboSectionType.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(cboSectionType)
            yPos += 35

            ' Section Size
            lblSectionSize = New Label()
            lblSectionSize.Text = "Section Size:"
            lblSectionSize.Location = New Point(15, yPos)
            lblSectionSize.Size = New Size(100, 20)
            lblSectionSize.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblSectionSize)

            cboSectionSize = New ComboBox()
            cboSectionSize.Location = New Point(130, yPos - 3)
            cboSectionSize.Size = New Size(250, 25)
            cboSectionSize.DropDownStyle = ComboBoxStyle.DropDownList
            cboSectionSize.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(cboSectionSize)
            yPos += 35

            ' Section Info GroupBox
            grpSectionInfo = New GroupBox()
            grpSectionInfo.Text = "Section Properties"
            grpSectionInfo.Location = New Point(15, yPos)
            grpSectionInfo.Size = New Size(500, 80)
            grpSectionInfo.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(grpSectionInfo)

            lblSectionDetails = New Label()
            lblSectionDetails.Text = "Select a section to view properties..."
            lblSectionDetails.Location = New Point(10, 20)
            lblSectionDetails.Size = New Size(480, 50)
            lblSectionDetails.Font = New Font("Consolas", 9)
            grpSectionInfo.Controls.Add(lblSectionDetails)
            yPos += 95

            ' Beam Length
            lblBeamLength = New Label()
            lblBeamLength.Text = "Beam Length:"
            lblBeamLength.Location = New Point(15, yPos)
            lblBeamLength.Size = New Size(100, 20)
            lblBeamLength.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblBeamLength)

            txtBeamLength = New TextBox()
            txtBeamLength.Location = New Point(130, yPos - 3)
            txtBeamLength.Size = New Size(100, 25)
            txtBeamLength.Text = "1000"
            txtBeamLength.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(txtBeamLength)

            lblLengthUnit = New Label()
            lblLengthUnit.Text = "mm"
            lblLengthUnit.Location = New Point(235, yPos)
            lblLengthUnit.Size = New Size(30, 20)
            lblLengthUnit.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(lblLengthUnit)
            yPos += 35

            ' Part Number Prefix
            lblPrefix = New Label()
            lblPrefix.Text = "Part Prefix:"
            lblPrefix.Location = New Point(15, yPos)
            lblPrefix.Size = New Size(100, 20)
            lblPrefix.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblPrefix)

            txtPrefix = New TextBox()
            txtPrefix.Location = New Point(130, yPos - 3)
            txtPrefix.Size = New Size(200, 25)
            txtPrefix.Text = "PLANT-000-"
            txtPrefix.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(txtPrefix)
            yPos += 35

            ' Output Folder
            lblOutputFolder = New Label()
            lblOutputFolder.Text = "Output Folder:"
            lblOutputFolder.Location = New Point(15, yPos)
            lblOutputFolder.Size = New Size(100, 20)
            lblOutputFolder.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblOutputFolder)

            txtOutputFolder = New TextBox()
            txtOutputFolder.Location = New Point(130, yPos - 3)
            txtOutputFolder.Size = New Size(330, 25)
            txtOutputFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            txtOutputFolder.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(txtOutputFolder)

            btnBrowseFolder = New Button()
            btnBrowseFolder.Text = "..."
            btnBrowseFolder.Location = New Point(465, yPos - 3)
            btnBrowseFolder.Size = New Size(40, 25)
            btnBrowseFolder.Font = New Font("Segoe UI", 9)
            Me.Controls.Add(btnBrowseFolder)
            yPos += 40

            ' Preview
            lblPreview = New Label()
            lblPreview.Text = "Preview (files to be created):"
            lblPreview.Location = New Point(15, yPos)
            lblPreview.Size = New Size(200, 20)
            lblPreview.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            Me.Controls.Add(lblPreview)
            yPos += 22

            txtPreview = New TextBox()
            txtPreview.Location = New Point(15, yPos)
            txtPreview.Size = New Size(500, 80)
            txtPreview.Multiline = True
            txtPreview.ReadOnly = True
            txtPreview.ScrollBars = ScrollBars.Vertical
            txtPreview.Font = New Font("Consolas", 8)
            txtPreview.BackColor = Color.FromArgb(245, 245, 245)
            Me.Controls.Add(txtPreview)
            yPos += 95

            ' Buttons
            btnOK = New Button()
            btnOK.Text = "Create Assembly"
            btnOK.Location = New Point(300, yPos)
            btnOK.Size = New Size(120, 35)
            btnOK.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnOK.BackColor = Color.FromArgb(76, 175, 80)
            btnOK.ForeColor = Color.White
            btnOK.FlatStyle = FlatStyle.Flat
            btnOK.DialogResult = DialogResult.OK
            Me.Controls.Add(btnOK)

            btnCancel = New Button()
            btnCancel.Text = "Cancel"
            btnCancel.Location = New Point(430, yPos)
            btnCancel.Size = New Size(85, 35)
            btnCancel.Font = New Font("Segoe UI", 9)
            btnCancel.DialogResult = DialogResult.Cancel
            Me.Controls.Add(btnCancel)

            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel
        End Sub

        ''' <summary>
        ''' Load section types into dropdown
        ''' </summary>
        Private Sub LoadSectionTypes()
            If cboSectionType Is Nothing Then Return

            cboSectionType.Items.Clear()
            For Each sectionType As String In SteelSectionDatabase.SectionTypes
                Dim displayName As String
                Select Case sectionType
                    Case "UC" : displayName = "UC - Universal Columns (H-sections)"
                    Case "UB" : displayName = "UB - Universal Beams (I-sections)"
                    Case "PFC" : displayName = "PFC - Parallel Flange Channels"
                    Case "TFC" : displayName = "TFC - Tapered Flange Channels"
                    Case "L" : displayName = "L - Equal Angles"
                    Case "CHS" : displayName = "CHS - Circular Hollow Sections"
                    Case "SHS" : displayName = "SHS - Square Hollow Sections"
                    Case "RHS" : displayName = "RHS - Rectangular Hollow Sections"
                    Case "FL" : displayName = "FL - Flat Bar"
                    Case Else : displayName = sectionType
                End Select
                cboSectionType.Items.Add(New SectionTypeItem(sectionType, displayName))
            Next

            If cboSectionType.Items.Count > 0 Then
                cboSectionType.SelectedIndex = 0
                ' Load initial section sizes
                Dim selectedType As SectionTypeItem = CType(cboSectionType.SelectedItem, SectionTypeItem)
                If selectedType IsNot Nothing Then
                    LoadSectionSizes(selectedType.TypeCode)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Handle section type selection change
        ''' </summary>
        Private Sub cboSectionType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSectionType.SelectedIndexChanged
            If m_IsInitializing Then Return
            If cboSectionType Is Nothing OrElse cboSectionType.SelectedItem Is Nothing Then Return

            Dim selectedType As SectionTypeItem = CType(cboSectionType.SelectedItem, SectionTypeItem)
            LoadSectionSizes(selectedType.TypeCode)
        End Sub

        ''' <summary>
        ''' Load section sizes for selected type
        ''' </summary>
        Private Sub LoadSectionSizes(sectionType As String)
            If cboSectionSize Is Nothing Then Return

            cboSectionSize.Items.Clear()

            Dim sections As List(Of SteelSection) = SteelSectionDatabase.GetSectionsByType(sectionType)
            For Each section As SteelSection In sections
                cboSectionSize.Items.Add(New SectionSizeItem(section))
            Next

            If cboSectionSize.Items.Count > 0 Then
                cboSectionSize.SelectedIndex = 0
            End If

            If Not m_IsInitializing Then
                UpdateSectionDetails()
                UpdatePreview()
            End If
        End Sub

        ''' <summary>
        ''' Handle section size selection change
        ''' </summary>
        Private Sub cboSectionSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSectionSize.SelectedIndexChanged
            If m_IsInitializing Then Return
            UpdateSectionDetails()
            UpdatePreview()
        End Sub

        ''' <summary>
        ''' Update section details display
        ''' </summary>
        Private Sub UpdateSectionDetails()
            ' Skip during initialization
            If m_IsInitializing Then Return
            If lblSectionDetails Is Nothing OrElse cboSectionSize Is Nothing Then Return

            If cboSectionSize.SelectedItem Is Nothing Then
                lblSectionDetails.Text = "Select a section to view properties..."
                Return
            End If

            Dim selectedItem As SectionSizeItem = CType(cboSectionSize.SelectedItem, SectionSizeItem)
            Dim section As SteelSection = selectedItem.Section

            Dim details As String = ""
            Select Case section.SectionType
                Case "UC", "UB"
                    details = String.Format("Height: {0:F1}mm  |  Width: {1:F1}mm  |  Web: {2:F1}mm  |  Flange: {3:F1}mm" & vbCrLf &
                                           "Mass: {4:F1} kg/m  |  Holes: {5}  ({6}mm for {7})" & vbCrLf &
                                           "Backmark X: {8}mm  |  Backmark Y: {9}mm",
                                           section.Height, section.Width, section.WebThickness, section.FlangeThickness,
                                           section.Mass, section.HoleCount, section.HoleDiameter, section.BoltSize,
                                           section.BackmarkX, section.BackmarkY)
                Case "PFC", "TFC"
                    details = String.Format("Height: {0:F1}mm  |  Width: {1:F1}mm  |  Web: {2:F1}mm  |  Flange: {3:F1}mm" & vbCrLf &
                                           "Mass: {4:F1} kg/m  |  Holes: {5}  ({6}mm for {7})" & vbCrLf &
                                           "Backmark: {8}mm",
                                           section.Height, section.Width, section.WebThickness, section.FlangeThickness,
                                           section.Mass, section.HoleCount, section.HoleDiameter, section.BoltSize,
                                           section.BackmarkX)
                Case "L"
                    details = String.Format("Leg A: {0:F1}mm  |  Leg B: {1:F1}mm  |  Thickness: {2:F1}mm" & vbCrLf &
                                           "Mass: {3:F1} kg/m  |  Holes: {4}  ({5}mm for {6})" & vbCrLf &
                                           "Backmark: {7}mm",
                                           section.LegA, section.LegB, section.FlangeThickness,
                                           section.Mass, section.HoleCount, section.HoleDiameter, section.BoltSize,
                                           section.BackmarkX)
                Case "CHS"
                    details = String.Format("Diameter: {0:F1}mm  |  Wall: {1:F1}mm  |  Mass: {2:F2} kg/m" & vbCrLf &
                                           "(No endplate holes - welded connection)",
                                           section.Height, section.WallThickness, section.Mass)
                Case "SHS"
                    details = String.Format("Size: {0:F1}x{0:F1}mm  |  Wall: {1:F1}mm  |  Mass: {2:F2} kg/m" & vbCrLf &
                                           "(No endplate holes - welded connection)",
                                           section.Height, section.WallThickness, section.Mass)
                Case "FL"
                    details = String.Format("Width: {0:F1}mm  |  Thickness: {1:F1}mm  |  Mass: {2:F2} kg/m",
                                           section.BarWidth, section.BarThickness, section.Mass)
                Case Else
                    details = section.GetIPropertyDescription()
            End Select

            lblSectionDetails.Text = details
        End Sub

        ''' <summary>
        ''' Update the preview display
        ''' </summary>
        Private Sub UpdatePreview()
            ' Skip during initialization
            If m_IsInitializing Then Return

            ' Check if controls exist
            If txtPreview Is Nothing OrElse cboSectionSize Is Nothing OrElse txtPrefix Is Nothing Then Return

            If cboSectionSize.SelectedItem Is Nothing Then
                txtPreview.Text = "Select a section to see preview..."
                Return
            End If

            Dim selectedItem As SectionSizeItem = CType(cboSectionSize.SelectedItem, SectionSizeItem)
            Dim section As SteelSection = selectedItem.Section

            Dim prefix As String = If(txtPrefix.Text IsNot Nothing, txtPrefix.Text.Trim(), "PLANT-000-")
            If Not prefix.EndsWith("-") Then prefix &= "-"

            Dim preview As String = ""
            preview &= "ASSEMBLY:  " & prefix & section.SectionType & "-ASSY.iam" & vbCrLf
            preview &= "BEAM:      " & prefix & section.SectionType & "-BEAM.ipt" & vbCrLf
            preview &= "           Description: " & section.GetIPropertyDescription() & vbCrLf
            preview &= "ENDPLATE:  " & prefix & section.SectionType & "-ENDPLATE.ipt" & vbCrLf

            Dim plateThickness As Integer = If(section.Height > 300, 20, 12)
            If section.SectionType = "PFC" OrElse section.SectionType = "TFC" OrElse section.SectionType = "L" Then
                plateThickness = 10
            End If
            preview &= "           Description: PL " & plateThickness & "mm S355JR" & vbCrLf
            preview &= vbCrLf
            preview &= "HOLES:     " & section.HoleCount & " x " & section.HoleDiameter & "mm holes (for " & section.BoltSize & " bolts)"

            txtPreview.Text = preview
        End Sub

        ''' <summary>
        ''' Handle text changes for live preview update
        ''' </summary>
        Private Sub txtPrefix_TextChanged(sender As Object, e As EventArgs) Handles txtPrefix.TextChanged
            If Not m_IsInitializing Then UpdatePreview()
        End Sub

        Private Sub txtBeamLength_TextChanged(sender As Object, e As EventArgs) Handles txtBeamLength.TextChanged
            If Not m_IsInitializing Then UpdatePreview()
        End Sub

        ''' <summary>
        ''' Browse for output folder
        ''' </summary>
        Private Sub btnBrowseFolder_Click(sender As Object, e As EventArgs) Handles btnBrowseFolder.Click
            Using dialog As New FolderBrowserDialog()
                dialog.Description = "Select output folder for beam assembly"
                dialog.SelectedPath = txtOutputFolder.Text
                dialog.ShowNewFolderButton = True

                If dialog.ShowDialog() = DialogResult.OK Then
                    txtOutputFolder.Text = dialog.SelectedPath
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Validate and set properties before closing
        ''' </summary>
        Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
            ' Validate section selection
            If cboSectionSize.SelectedItem Is Nothing Then
                MsgBox("Please select a section size.", MsgBoxStyle.Exclamation)
                Me.DialogResult = DialogResult.None
                Return
            End If

            ' Validate beam length
            Dim length As Double
            If Not Double.TryParse(txtBeamLength.Text, length) OrElse length <= 0 Then
                MsgBox("Please enter a valid beam length (positive number).", MsgBoxStyle.Exclamation)
                Me.DialogResult = DialogResult.None
                Return
            End If

            ' Validate output folder
            If String.IsNullOrEmpty(txtOutputFolder.Text) OrElse Not Directory.Exists(txtOutputFolder.Text) Then
                MsgBox("Please select a valid output folder.", MsgBoxStyle.Exclamation)
                Me.DialogResult = DialogResult.None
                Return
            End If

            ' Set properties
            Dim selectedItem As SectionSizeItem = CType(cboSectionSize.SelectedItem, SectionSizeItem)
            SelectedSection = selectedItem.Section
            BeamLength = length
            OutputFolder = txtOutputFolder.Text
            PartNumberPrefix = txtPrefix.Text.Trim()
            If Not PartNumberPrefix.EndsWith("-") Then PartNumberPrefix &= "-"
        End Sub

    End Class

    ''' <summary>
    ''' Helper class for section type combo items
    ''' </summary>
    Public Class SectionTypeItem
        Public Property TypeCode As String
        Public Property DisplayName As String

        Public Sub New(typeCode As String, displayName As String)
            Me.TypeCode = typeCode
            Me.DisplayName = displayName
        End Sub

        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class

    ''' <summary>
    ''' Helper class for section size combo items
    ''' </summary>
    Public Class SectionSizeItem
        Public Property Section As SteelSection

        Public Sub New(section As SteelSection)
            Me.Section = section
        End Sub

        Public Overrides Function ToString() As String
            Return Section.DisplayName & "  (" & Section.Mass.ToString("F1") & " kg/m)"
        End Function
    End Class

End Namespace
