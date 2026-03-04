' ==============================================================================
' STANDARD ADD-IN SERVER - Entry Point for Inventor Add-In
' ==============================================================================
' Author: Quintin de Bruin © 2025
' 
' This is the main entry point for the Inventor Add-In.
' It creates the ribbon button and handles user interaction.
' ==============================================================================

Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' The main Add-In server class that Inventor loads
    ''' </summary>
    <GuidAttribute("B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B")>
    <ComVisible(True)>
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        ' Inventor application reference
        Private m_InventorApp As Inventor.Application

        ' Button definitions
        Private WithEvents m_CloneButton As ButtonDefinition
        Private WithEvents m_ScanButton As ButtonDefinition
        Private WithEvents m_DocInfoButton As ButtonDefinition

        ' Our cloner and patcher modules
        Private m_Cloner As AssemblyCloner
        Private m_Patcher As iLogicPatcher
        Private m_DocInfoScanner As DocumentInfoScanner

#Region "ApplicationAddInServer Interface"

        ''' <summary>
        ''' Called when the Add-In is loaded by Inventor
        ''' </summary>
        Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) Implements ApplicationAddInServer.Activate
            ' Get reference to Inventor
            m_InventorApp = addInSiteObject.Application

            ' Initialize our modules
            m_Cloner = New AssemblyCloner(m_InventorApp)
            m_Patcher = New iLogicPatcher(m_InventorApp)
            m_DocInfoScanner = New DocumentInfoScanner(m_InventorApp)

            ' Create UI buttons
            Call CreateUserInterface()

            ' Log activation
            System.Diagnostics.Debug.WriteLine("AssemblyClonerAddIn: Activated successfully")
        End Sub

        ''' <summary>
        ''' Called when the Add-In is unloaded
        ''' </summary>
        Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
            ' Clean up
            m_CloneButton = Nothing
            m_ScanButton = Nothing
            m_DocInfoButton = Nothing
            m_Cloner = Nothing
            m_Patcher = Nothing
            m_DocInfoScanner = Nothing
            m_InventorApp = Nothing

            ' Force garbage collection
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub

        ''' <summary>
        ''' Returns automation object (not used)
        ''' </summary>
        Public ReadOnly Property Automation As Object Implements ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ''' <summary>
        ''' Execute command (not used)
        ''' </summary>
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User Interface"

        ''' <summary>
        ''' Create ribbon buttons
        ''' </summary>
        Private Sub CreateUserInterface()
            Try
                Dim controlDefs As ControlDefinitions = m_InventorApp.CommandManager.ControlDefinitions

                ' Inventor shows text-only buttons when no icons are provided.
                ' Create simple in-memory glyph icons (16x16 + 32x32) so the ribbon displays correctly
                ' even if no external PNG/ICO resources are shipped.
                Dim cloneIcon16 As stdole.IPictureDisp = CreateGlyphPicture("C", 16, System.Drawing.Color.FromArgb(33, 150, 243))
                Dim cloneIcon32 As stdole.IPictureDisp = CreateGlyphPicture("C", 32, System.Drawing.Color.FromArgb(33, 150, 243))
                Dim scanIcon16 As stdole.IPictureDisp = CreateGlyphPicture("I", 16, System.Drawing.Color.FromArgb(76, 175, 80))
                Dim scanIcon32 As stdole.IPictureDisp = CreateGlyphPicture("I", 32, System.Drawing.Color.FromArgb(76, 175, 80))
                Dim docIcon16 As stdole.IPictureDisp = CreateGlyphPicture("D", 16, System.Drawing.Color.FromArgb(255, 152, 0))
                Dim docIcon32 As stdole.IPictureDisp = CreateGlyphPicture("D", 32, System.Drawing.Color.FromArgb(255, 152, 0))

                ' Create Clone Assembly button
                m_CloneButton = controlDefs.AddButtonDefinition(
                    "Clone Assembly",
                    "Cmd_CloneAssembly",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Clone assembly with all parts to new folder and patch iLogic rules",
                    "Clone Assembly with iLogic Patching",
                    cloneIcon16,
                    cloneIcon32)

                ' Create Scan iLogic button
                m_ScanButton = controlDefs.AddButtonDefinition(
                    "Scan iLogic",
                    "Cmd_ScaniLogic",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Scan current document for iLogic rules and display details",
                    "Scan iLogic Rules",
                    scanIcon16,
                    scanIcon32)

                ' Create Document Info button
                m_DocInfoButton = controlDefs.AddButtonDefinition(
                    "Document Info",
                    "Cmd_DocInfo",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Display iProperties, Mass, and iLogic rules for current document",
                    "View Document Info",
                    docIcon16,
                    docIcon32)

                ' Add to Assembly ribbon
                Dim assemblyRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Assembly")
                Dim toolsTab As RibbonTab = assemblyRibbon.RibbonTabs("id_TabTools")

                ' Create our panel
                Dim customPanel As RibbonPanel = Nothing
                Try
                    customPanel = toolsTab.RibbonPanels.Add("Cloner Tools", "Pnl_ClonerTools", "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                Catch ex As Exception
                    ' Panel might already exist
                    For Each panel As RibbonPanel In toolsTab.RibbonPanels
                        If panel.InternalName = "Pnl_ClonerTools" Then
                            customPanel = panel
                            Exit For
                        End If
                    Next
                End Try

                If customPanel IsNot Nothing Then
                    customPanel.CommandControls.AddButton(m_CloneButton, True)
                    customPanel.CommandControls.AddButton(m_ScanButton, False)
                    customPanel.CommandControls.AddButton(m_DocInfoButton, False)
                End If

                ' Also add to Part ribbon
                Dim partRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Part")
                Dim partToolsTab As RibbonTab = partRibbon.RibbonTabs("id_TabTools")

                Dim partPanel As RibbonPanel = Nothing
                Try
                    partPanel = partToolsTab.RibbonPanels.Add("Cloner Tools", "Pnl_ClonerToolsPart", "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                Catch ex As Exception
                    For Each panel As RibbonPanel In partToolsTab.RibbonPanels
                        If panel.InternalName = "Pnl_ClonerToolsPart" Then
                            partPanel = panel
                            Exit For
                        End If
                    Next
                End Try

                If partPanel IsNot Nothing Then
                    partPanel.CommandControls.AddButton(m_ScanButton, True)
                    partPanel.CommandControls.AddButton(m_DocInfoButton, False)
                End If

            Catch ex As Exception
                MsgBox("Error creating UI: " & ex.Message, MsgBoxStyle.Critical, "AssemblyClonerAddIn")
            End Try
        End Sub

        ''' <summary>
        ''' Creates a simple colored square icon with a single letter.
        ''' This avoids shipping separate icon files and prevents the ribbon from showing text-only controls.
        ''' </summary>
        Private Shared Function CreateGlyphPicture(ByVal glyph As String, ByVal size As Integer, ByVal backColor As System.Drawing.Color) As stdole.IPictureDisp
            Dim bmp As New Bitmap(size, size, Imaging.PixelFormat.Format32bppArgb)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.SmoothingMode = SmoothingMode.AntiAlias
                g.InterpolationMode = InterpolationMode.HighQualityBicubic
                g.PixelOffsetMode = PixelOffsetMode.HighQuality

                Using bg As New SolidBrush(backColor)
                    g.FillRectangle(bg, 0, 0, size, size)
                End Using

                Using borderPen As New Pen(System.Drawing.Color.FromArgb(40, 0, 0, 0), Math.Max(1.0F, CSng(size) / 16.0F))
                    g.DrawRectangle(borderPen, 0, 0, size - 1, size - 1)
                End Using

                Dim fontSize As Single = If(size <= 16, 9.0F, 18.0F)
                Using f As New Font("Segoe UI", fontSize, FontStyle.Bold, GraphicsUnit.Point)
                    Using fg As New SolidBrush(System.Drawing.Color.White)
                        Dim rect As New RectangleF(0, 0, size, size)
                        Using sf As New StringFormat() With {
                            .Alignment = StringAlignment.Center,
                            .LineAlignment = StringAlignment.Center
                        }
                            g.DrawString(glyph, f, fg, rect, sf)
                        End Using
                    End Using
                End Using
            End Using

            Return PictureDispConverter.ImageToPictureDisp(bmp)
        End Function

        ''' <summary>
        ''' Helper to convert System.Drawing.Image to stdole.IPictureDisp (required by Inventor API).
        ''' </summary>
        Private NotInheritable Class PictureDispConverter
            Inherits AxHost

            Private Sub New()
                MyBase.New(String.Empty)
            End Sub

            Public Shared Function ImageToPictureDisp(ByVal image As Image) As stdole.IPictureDisp
                Return DirectCast(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
            End Function
        End Class

#End Region

#Region "Button Handlers"

        ''' <summary>
        ''' Clone Assembly button clicked
        ''' </summary>
        Private Sub m_CloneButton_OnExecute(ByVal Context As NameValueMap) Handles m_CloneButton.OnExecute
            Try
                ' Check if assembly is open
                If m_InventorApp.ActiveDocument Is Nothing Then
                    MsgBox("Please open an assembly first.", MsgBoxStyle.Exclamation, "Clone Assembly")
                    Return
                End If

                If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    MsgBox("Please open an assembly (.iam) first.", MsgBoxStyle.Exclamation, "Clone Assembly")
                    Return
                End If

                ' Run the cloner
                m_Cloner.CloneAssembly()

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Clone Assembly")
            End Try
        End Sub

        ''' <summary>
        ''' Scan iLogic button clicked
        ''' </summary>
        Private Sub m_ScanButton_OnExecute(ByVal Context As NameValueMap) Handles m_ScanButton.OnExecute
            Try
                ' Check if document is open
                If m_InventorApp.ActiveDocument Is Nothing Then
                    MsgBox("Please open a document first.", MsgBoxStyle.Exclamation, "Scan iLogic")
                    Return
                End If

                ' Run the scanner
                m_Patcher.ScanAndDisplayRules(m_InventorApp.ActiveDocument)

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Scan iLogic")
            End Try
        End Sub

        ''' <summary>
        ''' Document Info button clicked
        ''' </summary>
        Private Sub m_DocInfoButton_OnExecute(ByVal Context As NameValueMap) Handles m_DocInfoButton.OnExecute
            Try
                ' Check if document is open
                If m_InventorApp.ActiveDocument Is Nothing Then
                    MsgBox("Please open a document first.", MsgBoxStyle.Exclamation, "Document Info")
                    Return
                End If

                ' Run the document info scanner
                m_DocInfoScanner.ScanAndDisplayInfo()

            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Document Info")
            End Try
        End Sub

#End Region

    End Class

End Namespace
