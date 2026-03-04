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
Imports System.Diagnostics
Imports System.IO
Imports System.Reflection

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
        Private WithEvents m_ApplicationEvents As ApplicationEvents

        ' Button definitions
        Private WithEvents m_CloneButton As ButtonDefinition
        Private WithEvents m_PartRenamerButton As ButtonDefinition
        Private WithEvents m_TitleAutomationButton As ButtonDefinition
        Private WithEvents m_SetViewIdentifierButton As ButtonDefinition
        Private WithEvents m_SetViewScaleButton As ButtonDefinition
        Private WithEvents m_AutoBalloonerButton As ButtonDefinition
        Private WithEvents m_AutoDetailIDWButton As ButtonDefinition
        Private WithEvents m_RegistryManagementButton As ButtonDefinition
        Private WithEvents m_PopulateDwgRefButton As ButtonDefinition
        Private WithEvents m_PopulateDwgRefAutoPlaceButton As ButtonDefinition
        Private WithEvents m_CreateDxfForModelPlatesButton As ButtonDefinition
        Private WithEvents m_PlacePartsFromOpenAssemblyButton As ButtonDefinition
        Private WithEvents m_CreateSheetPartsListButton As ButtonDefinition
        Private WithEvents m_CreateGAPartsListTopLevelButton As ButtonDefinition
        Private WithEvents m_CleanUpUnusedFilesButton As ButtonDefinition
        Private WithEvents m_LengthParameterExporterButton As ButtonDefinition
        Private WithEvents m_Length2ParameterExporterButton As ButtonDefinition
        Private WithEvents m_ThicknessParameterExporterButton As ButtonDefinition
        Private WithEvents m_FixNonPlatePartsButton As ButtonDefinition
        Private WithEvents m_FixSinglePartLength2Button As ButtonDefinition
        Private WithEvents m_FixBOMPlateDimensionsButton As ButtonDefinition
        Private WithEvents m_ApplyPlateDescStockFormulaButton As ButtonDefinition

        ' Our cloner and patcher modules
        Private m_Cloner As AssemblyCloner
        Private m_AssemblyRenamer As AssemblyRenamerTool
        Private m_TitleAutomation As TitleAutomationUpdater
        Private m_ViewIdentifierSetter As ViewIdentifierSetter
        Private m_ViewScaleSetter As ViewScaleSetter
        Private m_AutoBalloonerTool As AutoBalloonerTool
        Private m_AutoDetailer As AutoDetailer
        Private m_RegistryManagement As RegistryManagementTool
        Private m_PopulateDwgRefTool As PopulateDwgRefTool
        Private m_PopulateDwgRefAutoPlaceTool As PopulateDwgRefTool
        Private m_CreateDxfForModelPlatesTool As CreateDxfForModelPlatesTool
        Private m_PlacePartsFromOpenAssemblyTool As PlacePartsFromOpenAssemblyTool
        Private m_CreateSheetPartsListTool As CreateSheetPartsListTool
        Private m_CreateGAPartsListTopLevelTool As CreateGAPartsListTopLevelTool
        Private m_CleanUpUnusedFilesTool As CleanUpUnusedFilesTool
        Private m_ParameterExporterTools As ParameterExporterTools

#Region "ApplicationAddInServer Interface"

        ''' <summary>
        ''' Called when the Add-In is loaded by Inventor
        ''' </summary>
        Public Sub Activate(ByVal addInSiteObject As ApplicationAddInSite, ByVal firstTime As Boolean) Implements ApplicationAddInServer.Activate
            ' Get reference to Inventor
            m_InventorApp = addInSiteObject.Application
            m_ApplicationEvents = m_InventorApp.ApplicationEvents

            ' Initialize our modules
            m_Cloner = New AssemblyCloner(m_InventorApp)
            m_AssemblyRenamer = New AssemblyRenamerTool(m_InventorApp)
            m_TitleAutomation = New TitleAutomationUpdater(m_InventorApp)
            m_ViewIdentifierSetter = New ViewIdentifierSetter(m_InventorApp)
            m_ViewScaleSetter = New ViewScaleSetter(m_InventorApp)
            m_AutoBalloonerTool = New AutoBalloonerTool(m_InventorApp)
            m_AutoDetailer = New AutoDetailer(m_InventorApp)
            m_RegistryManagement = New RegistryManagementTool(m_InventorApp)
            m_PopulateDwgRefTool = New PopulateDwgRefTool(m_InventorApp, False)
            m_PopulateDwgRefAutoPlaceTool = New PopulateDwgRefTool(m_InventorApp, True)
            m_CreateDxfForModelPlatesTool = New CreateDxfForModelPlatesTool(m_InventorApp)
            m_PlacePartsFromOpenAssemblyTool = New PlacePartsFromOpenAssemblyTool(m_InventorApp)
            m_CreateSheetPartsListTool = New CreateSheetPartsListTool(m_InventorApp)
            m_CreateGAPartsListTopLevelTool = New CreateGAPartsListTopLevelTool(m_InventorApp)
            m_CleanUpUnusedFilesTool = New CleanUpUnusedFilesTool(m_InventorApp)
            m_ParameterExporterTools = New ParameterExporterTools(m_InventorApp)

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
            m_PartRenamerButton = Nothing
            m_TitleAutomationButton = Nothing
            m_SetViewIdentifierButton = Nothing
            m_SetViewScaleButton = Nothing
            m_AutoBalloonerButton = Nothing
            m_AutoDetailIDWButton = Nothing
            m_RegistryManagementButton = Nothing
            m_PopulateDwgRefButton = Nothing
            m_PopulateDwgRefAutoPlaceButton = Nothing
            m_CreateDxfForModelPlatesButton = Nothing
            m_PlacePartsFromOpenAssemblyButton = Nothing
            m_CreateSheetPartsListButton = Nothing
            m_CreateGAPartsListTopLevelButton = Nothing
            m_CleanUpUnusedFilesButton = Nothing
            m_LengthParameterExporterButton = Nothing
            m_Length2ParameterExporterButton = Nothing
            m_ThicknessParameterExporterButton = Nothing
            m_FixNonPlatePartsButton = Nothing
            m_FixSinglePartLength2Button = Nothing
            m_FixBOMPlateDimensionsButton = Nothing
            m_ApplyPlateDescStockFormulaButton = Nothing
            m_ApplicationEvents = Nothing
            m_Cloner = Nothing
            m_AssemblyRenamer = Nothing
            m_TitleAutomation = Nothing
            m_ViewIdentifierSetter = Nothing
            m_ViewScaleSetter = Nothing
            m_AutoBalloonerTool = Nothing
            m_AutoDetailer = Nothing
            m_RegistryManagement = Nothing
            m_PopulateDwgRefTool = Nothing
            m_PopulateDwgRefAutoPlaceTool = Nothing
            m_CreateDxfForModelPlatesTool = Nothing
            m_PlacePartsFromOpenAssemblyTool = Nothing
            m_CreateSheetPartsListTool = Nothing
            m_CreateGAPartsListTopLevelTool = Nothing
            m_CleanUpUnusedFilesTool = Nothing
            m_ParameterExporterTools = Nothing
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

                RemoveLegacyRibbonPanels()

                ' Inventor shows text-only buttons when no icons are provided.
                ' Create simple in-memory glyph icons (16x16 + 32x32) so the ribbon displays correctly
                ' even if no external PNG/ICO resources are shipped.
                Dim cloneIcon16 As stdole.IPictureDisp = CreateGlyphPicture("C", 16, System.Drawing.Color.FromArgb(33, 150, 243))
                Dim cloneIcon32 As stdole.IPictureDisp = CreateGlyphPicture("C", 32, System.Drawing.Color.FromArgb(33, 150, 243))

                ' Create Clone Assembly button
                m_CloneButton = controlDefs.AddButtonDefinition(
                    "Clone Assembly",
                    "Cmd_SpectivCloneAssembly2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Clone assembly with all parts to new folder and patch iLogic rules",
                    "Clone Assembly with iLogic Patching",
                    cloneIcon16,
                    cloneIcon32)

                ' Create Part Renamer button
                Dim renamerIcon16 As stdole.IPictureDisp = CreateGlyphPicture("R", 16, System.Drawing.Color.FromArgb(255, 87, 34))
                Dim renamerIcon32 As stdole.IPictureDisp = CreateGlyphPicture("R", 32, System.Drawing.Color.FromArgb(255, 87, 34))
                m_PartRenamerButton = controlDefs.AddButtonDefinition(
                    "Part Renamer",
                    "Cmd_SpectivPartRenamer2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Rename assembly parts with heritage method",
                    "Part Renamer",
                    renamerIcon16,
                    renamerIcon32)

                Dim titleIcon16 As stdole.IPictureDisp = CreateGlyphPicture("T", 16, System.Drawing.Color.FromArgb(156, 39, 176))
                Dim titleIcon32 As stdole.IPictureDisp = CreateGlyphPicture("T", 32, System.Drawing.Color.FromArgb(156, 39, 176))
                m_TitleAutomationButton = controlDefs.AddButtonDefinition(
                    "Title Automation (IDW)",
                    "Cmd_SpectivTitleAutomationIDW2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Update IDW view titles to Spectiv title format",
                    "Title Automation (IDW only)",
                    titleIcon16,
                    titleIcon32)

                Dim viewIdentifierIcon16 As stdole.IPictureDisp = CreateGlyphPicture("VI", 16, System.Drawing.Color.FromArgb(0, 150, 136))
                Dim viewIdentifierIcon32 As stdole.IPictureDisp = CreateGlyphPicture("VI", 32, System.Drawing.Color.FromArgb(0, 150, 136))
                m_SetViewIdentifierButton = controlDefs.AddButtonDefinition(
                    "Set View Identifier",
                    "Cmd_SpectivSetViewIdentifier2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Pick a drawing view and set its View Identifier using preset names or custom text",
                    "Set View Identifier",
                    viewIdentifierIcon16,
                    viewIdentifierIcon32)

                Dim viewScaleIcon16 As stdole.IPictureDisp = CreateGlyphPicture("VS", 16, System.Drawing.Color.FromArgb(0, 137, 123))
                Dim viewScaleIcon32 As stdole.IPictureDisp = CreateGlyphPicture("VS", 32, System.Drawing.Color.FromArgb(0, 137, 123))
                m_SetViewScaleButton = controlDefs.AddButtonDefinition(
                    "Set View Scale",
                    "Cmd_SpectivSetViewScale2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Choose a scale from presets or custom, then click a drawing view to apply it",
                    "Set View Scale",
                    viewScaleIcon16,
                    viewScaleIcon32)

                Dim autoBalloonIcon16 As stdole.IPictureDisp = CreateGlyphPicture("B", 16, System.Drawing.Color.FromArgb(0, 121, 107))
                Dim autoBalloonIcon32 As stdole.IPictureDisp = CreateGlyphPicture("B", 32, System.Drawing.Color.FromArgb(0, 121, 107))
                m_AutoBalloonerButton = controlDefs.AddButtonDefinition(
                    "Auto Ballooner",
                    "Cmd_SpectivAutoBallooner2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Auto-place balloons on the active sheet for visible assembly parts using nearest-edge leader placement",
                    "Auto Ballooner",
                    autoBalloonIcon16,
                    autoBalloonIcon32)

                Dim autoDetailIcon16 As stdole.IPictureDisp = CreateGlyphPicture("AD", 16, System.Drawing.Color.FromArgb(63, 81, 181))
                Dim autoDetailIcon32 As stdole.IPictureDisp = CreateGlyphPicture("AD", 32, System.Drawing.Color.FromArgb(63, 81, 181))
                m_AutoDetailIDWButton = controlDefs.AddButtonDefinition(
                    "Auto Detail IDW",
                    "Cmd_SpectivAutoDetailIDW2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Auto-place overall and feature dimensions for a selected drawing view",
                    "Auto Detail IDW",
                    autoDetailIcon16,
                    autoDetailIcon32)

                Dim registryIcon16 As stdole.IPictureDisp = CreateGlyphPicture("G", 16, System.Drawing.Color.FromArgb(0, 150, 136))
                Dim registryIcon32 As stdole.IPictureDisp = CreateGlyphPicture("G", 32, System.Drawing.Color.FromArgb(0, 150, 136))
                m_RegistryManagementButton = controlDefs.AddButtonDefinition(
                    "Registry Management",
                    "Cmd_SpectivRegistryManagement2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Manage Inventor Renamer counters and prefix registry data",
                    "Registry Management",
                    registryIcon16,
                    registryIcon32)

                Dim dwgRefIcon16 As stdole.IPictureDisp = CreateGlyphPicture("D", 16, System.Drawing.Color.FromArgb(63, 81, 181))
                Dim dwgRefIcon32 As stdole.IPictureDisp = CreateGlyphPicture("D", 32, System.Drawing.Color.FromArgb(63, 81, 181))
                m_PopulateDwgRefButton = controlDefs.AddButtonDefinition(
                    "Populate DWG REF from Parts Lists",
                    "Cmd_SpectivPopulateDWGRefFromPartsLists2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Scan non-DXF sheets and update DWG REF values from parts lists",
                    "Populate DWG REF from Parts Lists",
                    dwgRefIcon16,
                    dwgRefIcon32)

                Dim autoPlaceIcon16 As stdole.IPictureDisp = CreateGlyphPicture("A", 16, System.Drawing.Color.FromArgb(0, 188, 212))
                Dim autoPlaceIcon32 As stdole.IPictureDisp = CreateGlyphPicture("A", 32, System.Drawing.Color.FromArgb(0, 188, 212))
                m_PopulateDwgRefAutoPlaceButton = controlDefs.AddButtonDefinition(
                    "Populate DWG REF + Auto-place Missing Parts",
                    "Cmd_SpectivPopulateDWGRefAutoPlace2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Run DWG REF update and auto-place missing parts on selected non-DXF sheet",
                    "Populate DWG REF + Auto-place Missing Parts",
                    autoPlaceIcon16,
                    autoPlaceIcon32)

                Dim dxfIcon16 As stdole.IPictureDisp = CreateGlyphPicture("X", 16, System.Drawing.Color.FromArgb(121, 85, 72))
                Dim dxfIcon32 As stdole.IPictureDisp = CreateGlyphPicture("X", 32, System.Drawing.Color.FromArgb(121, 85, 72))
                m_CreateDxfForModelPlatesButton = controlDefs.AddButtonDefinition(
                    "CREATE DXF FOR MODEL PLATES",
                    "Cmd_SpectivCreateDXFForModelPlates2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Select source sheet, collect plate parts, and create DXF sheet(s)",
                    "CREATE DXF FOR MODEL PLATES",
                    dxfIcon16,
                    dxfIcon32)

                Dim placeIcon16 As stdole.IPictureDisp = CreateGlyphPicture("P", 16, System.Drawing.Color.FromArgb(76, 175, 80))
                Dim placeIcon32 As stdole.IPictureDisp = CreateGlyphPicture("P", 32, System.Drawing.Color.FromArgb(76, 175, 80))
                m_PlacePartsFromOpenAssemblyButton = controlDefs.AddButtonDefinition(
                    "Place parts from open Assembly",
                    "Cmd_SpectivPlacePartsFromOpenAssembly2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Build ordered assembly/subassembly detailing sheets and parts placement",
                    "Place parts from open Assembly",
                    placeIcon16,
                    placeIcon32)

                Dim sheetListIcon16 As stdole.IPictureDisp = CreateGlyphPicture("L", 16, System.Drawing.Color.FromArgb(255, 152, 0))
                Dim sheetListIcon32 As stdole.IPictureDisp = CreateGlyphPicture("L", 32, System.Drawing.Color.FromArgb(255, 152, 0))
                m_CreateSheetPartsListButton = controlDefs.AddButtonDefinition(
                    "Create Sheet Parts List",
                    "Cmd_SpectivCreateSheetPartsList2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Standalone tool: create parts list for active sheet from selected assembly and filter to visible sheet parts",
                    "Create Sheet Parts List",
                    sheetListIcon16,
                    sheetListIcon32)

                Dim gaTopIcon16 As stdole.IPictureDisp = CreateGlyphPicture("G", 16, System.Drawing.Color.FromArgb(233, 30, 99))
                Dim gaTopIcon32 As stdole.IPictureDisp = CreateGlyphPicture("G", 32, System.Drawing.Color.FromArgb(233, 30, 99))
                m_CreateGAPartsListTopLevelButton = controlDefs.AddButtonDefinition(
                    "Create GA Parts List (Top Level)",
                    "Cmd_SpectivCreateGAPartsListTopLevel2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Create GA parts list showing only first-level components of the selected assembly",
                    "Create GA Parts List (Top Level)",
                    gaTopIcon16,
                    gaTopIcon32)

                Dim cleanupIcon16 As stdole.IPictureDisp = CreateGlyphPicture("U", 16, System.Drawing.Color.FromArgb(244, 67, 54))
                Dim cleanupIcon32 As stdole.IPictureDisp = CreateGlyphPicture("U", 32, System.Drawing.Color.FromArgb(244, 67, 54))
                m_CleanUpUnusedFilesButton = controlDefs.AddButtonDefinition(
                    "Clean Up Unused Files",
                    "Cmd_SpectivCleanUpUnusedFiles2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Scan IDW drawing for referenced parts and move unreferenced .ipt files to backup folder",
                    "Clean Up Unused Files",
                    cleanupIcon16,
                    cleanupIcon32)

                Dim lengthIcon16 As stdole.IPictureDisp = CreateGlyphPicture("L", 16, System.Drawing.Color.FromArgb(76, 175, 80))
                Dim lengthIcon32 As stdole.IPictureDisp = CreateGlyphPicture("L", 32, System.Drawing.Color.FromArgb(76, 175, 80))
                m_LengthParameterExporterButton = controlDefs.AddButtonDefinition(
                    "Length Parameter Exporter",
                    "Cmd_SpectivLengthParameterExporter2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Enable export for Length parameter on non-plate parts in active assembly",
                    "Length Parameter Exporter",
                    lengthIcon16,
                    lengthIcon32)

                Dim length2Icon16 As stdole.IPictureDisp = CreateGlyphPicture("L2", 16, System.Drawing.Color.FromArgb(139, 195, 74))
                Dim length2Icon32 As stdole.IPictureDisp = CreateGlyphPicture("L2", 32, System.Drawing.Color.FromArgb(139, 195, 74))
                m_Length2ParameterExporterButton = controlDefs.AddButtonDefinition(
                    "Length2 Parameter Exporter",
                    "Cmd_SpectivLength2ParameterExporter2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Enable export for Length2 parameter on non-plate parts in active assembly",
                    "Length2 Parameter Exporter",
                    length2Icon16,
                    length2Icon32)

                Dim thicknessIcon16 As stdole.IPictureDisp = CreateGlyphPicture("T", 16, System.Drawing.Color.FromArgb(255, 152, 0))
                Dim thicknessIcon32 As stdole.IPictureDisp = CreateGlyphPicture("T", 32, System.Drawing.Color.FromArgb(255, 152, 0))
                m_ThicknessParameterExporterButton = controlDefs.AddButtonDefinition(
                    "Thickness Parameter Exporter",
                    "Cmd_SpectivThicknessParameterExporter2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Enable export for Thickness parameter on plate parts in active assembly",
                    "Thickness Parameter Exporter",
                    thicknessIcon16,
                    thicknessIcon32)

                Dim fixNonPlateIcon16 As stdole.IPictureDisp = CreateGlyphPicture("F", 16, System.Drawing.Color.FromArgb(255, 193, 7))
                Dim fixNonPlateIcon32 As stdole.IPictureDisp = CreateGlyphPicture("F", 32, System.Drawing.Color.FromArgb(255, 193, 7))
                m_FixNonPlatePartsButton = controlDefs.AddButtonDefinition(
                    "Fix Non-Plate Parts",
                    "Cmd_SpectivFixNonPlateParts2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Add/update Length2 on non-plate parts in active assembly",
                    "Fix Non-Plate Parts",
                    fixNonPlateIcon16,
                    fixNonPlateIcon32)

                Dim fixSingleIcon16 As stdole.IPictureDisp = CreateGlyphPicture("S", 16, System.Drawing.Color.FromArgb(96, 125, 139))
                Dim fixSingleIcon32 As stdole.IPictureDisp = CreateGlyphPicture("S", 32, System.Drawing.Color.FromArgb(96, 125, 139))
                m_FixSinglePartLength2Button = controlDefs.AddButtonDefinition(
                    "Fix Single Part Length2",
                    "Cmd_SpectivFixSinglePartLength22026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Add/update Length2 on active part and link to longest model dimension",
                    "Fix Single Part Length2",
                    fixSingleIcon16,
                    fixSingleIcon32)

                Dim fixBomIcon16 As stdole.IPictureDisp = CreateGlyphPicture("B", 16, System.Drawing.Color.FromArgb(3, 169, 244))
                Dim fixBomIcon32 As stdole.IPictureDisp = CreateGlyphPicture("B", 32, System.Drawing.Color.FromArgb(3, 169, 244))
                m_FixBOMPlateDimensionsButton = controlDefs.AddButtonDefinition(
                    "Fix BOM Plate Dimensions",
                    "Cmd_SpectivFixBOMPlateDimensions2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "Set plate part LENGTH/WIDTH custom properties from sheet metal dimensions",
                    "Fix BOM Plate Dimensions",
                    fixBomIcon16,
                    fixBomIcon32)

                Dim finalStepIcon16 As stdole.IPictureDisp = CreateGlyphPicture("!", 16, System.Drawing.Color.FromArgb(233, 30, 99))
                Dim finalStepIcon32 As stdole.IPictureDisp = CreateGlyphPicture("!", 32, System.Drawing.Color.FromArgb(233, 30, 99))
                m_ApplyPlateDescStockFormulaButton = controlDefs.AddButtonDefinition(
                    "Apply Plate Desc/Stock Formula",
                    "Cmd_SpectivApplyPlateDescStockFormula2026",
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}",
                    "LAST STEP ONLY: set Description and Stock Number formula for plate model(s)",
                    "Apply Plate Desc/Stock Formula",
                    finalStepIcon16,
                    finalStepIcon32)

                ' Add to Assembly ribbon
                Dim assemblyRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Assembly")
                Dim toolsTab As RibbonTab = assemblyRibbon.RibbonTabs("id_TabTools")

                DeletePanelsByPrefix(toolsTab, "Pnl_ClonerTools")
                DeletePanelsByPrefix(toolsTab, "Pnl_SpectivTools2026")
                DeletePanelIfExists(toolsTab, "Pnl_SpectivParamTools2026")

                ' Create our panel fresh each activation
                Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("Spectiv Inventor Automation Suite 2026", "Pnl_SpectivTools2026", "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")

                If customPanel IsNot Nothing Then
                    customPanel.CommandControls.AddButton(m_CloneButton, True)
                    customPanel.CommandControls.AddButton(m_PartRenamerButton, False)
                    customPanel.CommandControls.AddButton(m_RegistryManagementButton, False)
                End If

                Dim parameterPanel As RibbonPanel = Nothing
                Try
                    parameterPanel = toolsTab.RibbonPanels.Add("Spectiv Parameter Management 2026", "Pnl_SpectivParamTools2026", "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                Catch
                    parameterPanel = Nothing
                End Try

                If parameterPanel IsNot Nothing Then
                    parameterPanel.CommandControls.AddButton(m_LengthParameterExporterButton, True)
                    parameterPanel.CommandControls.AddButton(m_Length2ParameterExporterButton, True)
                    parameterPanel.CommandControls.AddButton(m_ThicknessParameterExporterButton, True)
                    parameterPanel.CommandControls.AddButton(m_FixNonPlatePartsButton, False)
                    parameterPanel.CommandControls.AddButton(m_FixSinglePartLength2Button, False)
                    parameterPanel.CommandControls.AddButton(m_FixBOMPlateDimensionsButton, False)
                    parameterPanel.CommandControls.AddButton(m_ApplyPlateDescStockFormulaButton, False)
                End If

                EnsureDrawingRibbonButtons()

                ' Intentionally hide all other tool buttons for now.
                ' Migration to the add-in will re-enable tools incrementally.

            Catch ex As Exception
                MsgBox("Error creating UI: " & ex.Message, MsgBoxStyle.Critical, "AssemblyClonerAddIn")
            End Try
        End Sub

        Private Sub RemoveLegacyRibbonPanels()
            Try
                Dim assemblyRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Assembly")
                If assemblyRibbon IsNot Nothing Then
                    Dim assemblyToolsTab As RibbonTab = Nothing
                    Try
                        assemblyToolsTab = assemblyRibbon.RibbonTabs("id_TabTools")
                    Catch
                        assemblyToolsTab = Nothing
                    End Try

                    If assemblyToolsTab IsNot Nothing Then
                        DeletePanelsByPrefix(assemblyToolsTab, "Pnl_ClonerTools")
                        DeletePanelsByPrefix(assemblyToolsTab, "Pnl_SpectivTools2026")
                    End If
                End If
            Catch
            End Try

            Try
                Dim drawingRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Drawing")
                If drawingRibbon IsNot Nothing Then
                    For Each tab As RibbonTab In drawingRibbon.RibbonTabs
                        DeletePanelsByPrefix(tab, "Pnl_ClonerTools")
                        DeletePanelsByPrefix(tab, "Pnl_SpectivTools2026")
                    Next
                End If
            Catch
            End Try
        End Sub

        Private Sub DeletePanelsByPrefix(ByVal tab As RibbonTab, ByVal panelInternalNamePrefix As String)
            If tab Is Nothing Then
                Return
            End If

            For i As Integer = tab.RibbonPanels.Count To 1 Step -1
                Try
                    Dim panel As RibbonPanel = tab.RibbonPanels.Item(i)
                    If panel IsNot Nothing AndAlso panel.InternalName IsNot Nothing AndAlso panel.InternalName.StartsWith(panelInternalNamePrefix, StringComparison.OrdinalIgnoreCase) Then
                        panel.Delete()
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Sub DeletePanelIfExists(ByVal tab As RibbonTab, ByVal panelInternalName As String)
            If tab Is Nothing Then
                Return
            End If

            For i As Integer = tab.RibbonPanels.Count To 1 Step -1
                Try
                    Dim panel As RibbonPanel = tab.RibbonPanels.Item(i)
                    If String.Equals(panel.InternalName, panelInternalName, StringComparison.OrdinalIgnoreCase) Then
                        panel.Delete()
                    End If
                Catch
                End Try
            Next
        End Sub

        Private Sub EnsureDrawingRibbonButtons()
            Try
                Dim drawingRibbon As Ribbon = m_InventorApp.UserInterfaceManager.Ribbons("Drawing")
                If drawingRibbon Is Nothing Then
                    Return
                End If

                For Each tab As RibbonTab In drawingRibbon.RibbonTabs
                    If tab Is Nothing Then
                        Continue For
                    End If

                    Dim legacyPanelInternalName As String = "Pnl_SpectivTools2026_Drawing_" & tab.InternalName
                    Dim corePanelInternalName As String = "Pnl_SpectivDrawingCore2026_" & tab.InternalName
                    Dim refPanelInternalName As String = "Pnl_SpectivDrawingRef2026_" & tab.InternalName
                    Dim listPanelInternalName As String = "Pnl_SpectivDrawingList2026_" & tab.InternalName

                    DeletePanelIfExists(tab, legacyPanelInternalName)
                    DeletePanelIfExists(tab, corePanelInternalName)
                    DeletePanelIfExists(tab, refPanelInternalName)
                    DeletePanelIfExists(tab, listPanelInternalName)

                    Dim corePanel As RibbonPanel = Nothing
                    Try
                        corePanel = tab.RibbonPanels.Add("Spectiv Drawing Core 2026", corePanelInternalName, "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                    Catch
                        corePanel = Nothing
                    End Try

                    If corePanel IsNot Nothing Then
                        corePanel.CommandControls.AddButton(m_TitleAutomationButton, True)
                        corePanel.CommandControls.AddButton(m_SetViewIdentifierButton, False)
                        corePanel.CommandControls.AddButton(m_SetViewScaleButton, False)
                        corePanel.CommandControls.AddButton(m_AutoBalloonerButton, False)
                        corePanel.CommandControls.AddButton(m_AutoDetailIDWButton, False)
                    End If

                    Dim refPanel As RibbonPanel = Nothing
                    Try
                        refPanel = tab.RibbonPanels.Add("Spectiv DWG REF & DXF 2026", refPanelInternalName, "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                    Catch
                        refPanel = Nothing
                    End Try

                    If refPanel IsNot Nothing Then
                        refPanel.CommandControls.AddButton(m_PopulateDwgRefButton, True)
                        refPanel.CommandControls.AddButton(m_PopulateDwgRefAutoPlaceButton, False)
                        refPanel.CommandControls.AddButton(m_CreateDxfForModelPlatesButton, False)
                    End If

                    Dim listPanel As RibbonPanel = Nothing
                    Try
                        listPanel = tab.RibbonPanels.Add("Spectiv Lists & Placement 2026", listPanelInternalName, "{B8F4E2A1-3C5D-4E6F-9A8B-1C2D3E4F5A6B}")
                    Catch
                        listPanel = Nothing
                    End Try

                    If listPanel IsNot Nothing Then
                        listPanel.CommandControls.AddButton(m_PlacePartsFromOpenAssemblyButton, True)
                        listPanel.CommandControls.AddButton(m_CreateSheetPartsListButton, False)
                        listPanel.CommandControls.AddButton(m_CreateGAPartsListTopLevelButton, False)
                        listPanel.CommandControls.AddButton(m_CleanUpUnusedFilesButton, False)
                    End If
                Next
            Catch
            End Try
        End Sub

        Private Sub m_ApplicationEvents_OnActivateDocument(ByVal DocumentObject As _Document,
                                                           ByVal BeforeOrAfter As EventTimingEnum,
                                                           ByVal Context As NameValueMap,
                                                           ByRef HandlingCode As HandlingCodeEnum) Handles m_ApplicationEvents.OnActivateDocument
            If BeforeOrAfter <> EventTimingEnum.kAfter Then
                Return
            End If

            Try
                If DocumentObject IsNot Nothing AndAlso DocumentObject.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                    EnsureDrawingRibbonButtons()
                End If
            Catch
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

        Private Function GuardToolEnabled(ByVal tool As AddInTool, ByVal toolDisplayName As String) As Boolean
            If AddInFeatureFlags.IsEnabled(tool) Then
                Return True
            End If

            MsgBox(AddInFeatureFlags.DisabledMessage(toolDisplayName), MsgBoxStyle.Information, toolDisplayName)
            Return False
        End Function

        Private Function ConfirmAssemblyRenamerPreflight() As Boolean
            If m_InventorApp.ActiveDocument Is Nothing Then
                MsgBox("Please open the assembly you want to rename before running Assembly Renamer.", MsgBoxStyle.Exclamation, "Assembly Renamer")
                Return False
            End If

            If m_InventorApp.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                MsgBox("Assembly Renamer requires an open .iam assembly as the active document.", MsgBoxStyle.Exclamation, "Assembly Renamer")
                Return False
            End If

            Dim activeDoc As Document = m_InventorApp.ActiveDocument
            If String.IsNullOrWhiteSpace(activeDoc.FullFileName) OrElse Not System.IO.File.Exists(activeDoc.FullFileName) Then
                MsgBox("Save the assembly to disk before running Assembly Renamer.", MsgBoxStyle.Exclamation, "Assembly Renamer")
                Return False
            End If

            Dim preflight As String =
                "Assembly Renamer - Preflight Checklist" & vbCrLf & vbCrLf &
                "Before continuing, make sure:" & vbCrLf &
                "1) The correct top-level assembly (.iam) is active." & vbCrLf &
                "2) Assembly and referenced files are saved and writable." & vbCrLf &
                "3) You have a backup/copy of the project folder." & vbCrLf &
                "4) No external rename/update scripts are running." & vbCrLf &
                "5) You are ready to let the tool update IDW references and write STEP_1_MAPPING.txt." & vbCrLf & vbCrLf &
                "Continue with Assembly Renamer?"

            Return MsgBox(preflight, MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Assembly Renamer") = MsgBoxResult.Yes
        End Function

        ''' <summary>
        ''' Clone Assembly button clicked
        ''' </summary>
        Private Sub m_CloneButton_OnExecute(ByVal Context As NameValueMap) Handles m_CloneButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.CloneAssembly, "Clone Assembly") Then Return

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
        ''' Handles the Part Renamer button click
        ''' </summary>
        Private Sub m_PartRenamerButton_OnExecute(ByVal Context As NameValueMap) Handles m_PartRenamerButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.AssemblyRenamer, "Assembly Renamer") Then Return
                If Not ConfirmAssemblyRenamerPreflight() Then Return

                ' Run the assembly renamer tool (native add-in implementation)
                m_AssemblyRenamer.Execute()

            Catch ex As Exception
                MsgBox("Error running Assembly Renamer: " & ex.Message, MsgBoxStyle.Critical, "Assembly Renamer")
            End Try
        End Sub

        Private Sub m_TitleAutomationButton_OnExecute(ByVal Context As NameValueMap) Handles m_TitleAutomationButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.TitleAutomationIDWOnly, "Title Automation (IDW)") Then Return
                m_TitleAutomation.Execute()
            Catch ex As Exception
                MsgBox("Error running Title Automation (IDW): " & ex.Message, MsgBoxStyle.Critical, "Title Automation (IDW)")
            End Try
        End Sub

        Private Sub m_SetViewIdentifierButton_OnExecute(ByVal Context As NameValueMap) Handles m_SetViewIdentifierButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.SetViewIdentifier, "Set View Identifier") Then Return
                m_ViewIdentifierSetter.Execute()
            Catch ex As Exception
                LogToolError("SetViewIdentifier", ex)
                MsgBox("Error running Set View Identifier: " & ex.Message, MsgBoxStyle.Critical, "Set View Identifier")
            End Try
        End Sub

        Private Sub m_SetViewScaleButton_OnExecute(ByVal Context As NameValueMap) Handles m_SetViewScaleButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.SetViewScale, "Set View Scale") Then Return
                m_ViewScaleSetter.Execute()
            Catch ex As Exception
                LogToolError("SetViewScale", ex)
                MsgBox("Error running Set View Scale: " & ex.Message, MsgBoxStyle.Critical, "Set View Scale")
            End Try
        End Sub

        Private Sub m_AutoBalloonerButton_OnExecute(ByVal Context As NameValueMap) Handles m_AutoBalloonerButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.AutoBalloonLeaders, "Auto Ballooner") Then Return
                m_AutoBalloonerTool.Execute()
            Catch ex As Exception
                LogToolError("AutoBallooner", ex)
                MsgBox("Error running Auto Ballooner: " & ex.Message, MsgBoxStyle.Critical, "Auto Ballooner")
            End Try
        End Sub

        Private Sub m_AutoDetailIDWButton_OnExecute(ByVal Context As NameValueMap) Handles m_AutoDetailIDWButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.AutoDetailIDW, "Auto Detail IDW") Then Return
                m_AutoDetailer.Execute()
            Catch ex As Exception
                LogToolError("AutoDetailIDW", ex)
                MsgBox("Error running Auto Detail IDW: " & ex.Message, MsgBoxStyle.Critical, "Auto Detail IDW")
            End Try
        End Sub

        Private Sub m_RegistryManagementButton_OnExecute(ByVal Context As NameValueMap) Handles m_RegistryManagementButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.RegistryManagement, "Registry Management") Then Return
                m_RegistryManagement.Execute()
            Catch ex As Exception
                MsgBox("Error running Registry Management: " & ex.Message, MsgBoxStyle.Critical, "Registry Management")
            End Try
        End Sub

        Private Sub m_PopulateDwgRefButton_OnExecute(ByVal Context As NameValueMap) Handles m_PopulateDwgRefButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.PopulateDWGRefFromPartsLists, "Populate DWG REF from Parts Lists") Then Return
                m_PopulateDwgRefTool.Execute()
            Catch ex As Exception
                LogToolError("PopulateDWGRefFromPartsLists", ex)
                MsgBox("Error running Populate DWG REF from Parts Lists: " & ex.Message, MsgBoxStyle.Critical, "Populate DWG REF")
            End Try
        End Sub

        Private Sub m_PopulateDwgRefAutoPlaceButton_OnExecute(ByVal Context As NameValueMap) Handles m_PopulateDwgRefAutoPlaceButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.PopulateDWGRefAutoPlaceMissingParts, "Populate DWG REF + Auto-place Missing Parts") Then Return
                m_PopulateDwgRefAutoPlaceTool.Execute()
            Catch ex As Exception
                LogToolError("PopulateDWGRefAutoPlaceMissingParts", ex)
                MsgBox("Error running Populate DWG REF + Auto-place Missing Parts: " & ex.Message, MsgBoxStyle.Critical, "Populate DWG REF + Auto-place")
            End Try
        End Sub

        Private Sub m_CreateDxfForModelPlatesButton_OnExecute(ByVal Context As NameValueMap) Handles m_CreateDxfForModelPlatesButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.CreateDXFForModelPlates, "CREATE DXF FOR MODEL PLATES") Then Return
                m_CreateDxfForModelPlatesTool.Execute()
            Catch ex As Exception
                LogToolError("CreateDXFForModelPlates", ex)
                MsgBox("Error running CREATE DXF FOR MODEL PLATES: " & ex.Message, MsgBoxStyle.Critical, "CREATE DXF FOR MODEL PLATES")
            End Try
        End Sub

        Private Sub m_PlacePartsFromOpenAssemblyButton_OnExecute(ByVal Context As NameValueMap) Handles m_PlacePartsFromOpenAssemblyButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.PartPlacer, "Place parts from open Assembly") Then Return
                m_PlacePartsFromOpenAssemblyTool.Execute()
            Catch ex As Exception
                LogToolError("PlacePartsFromOpenAssembly", ex)
                MsgBox("Error running Place parts from open Assembly: " & ex.Message, MsgBoxStyle.Critical, "Place parts from open Assembly")
            End Try
        End Sub

        Private Sub m_CreateSheetPartsListButton_OnExecute(ByVal Context As NameValueMap) Handles m_CreateSheetPartsListButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.CreateSheetPartsList, "Create Sheet Parts List") Then Return
                m_CreateSheetPartsListTool.Execute()
            Catch ex As Exception
                LogToolError("CreateSheetPartsList", ex)
                MsgBox("Error running Create Sheet Parts List: " & ex.Message, MsgBoxStyle.Critical, "Create Sheet Parts List")
            End Try
        End Sub

        Private Sub m_CreateGAPartsListTopLevelButton_OnExecute(ByVal Context As NameValueMap) Handles m_CreateGAPartsListTopLevelButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.CreateGAPartsListTopLevel, "Create GA Parts List (Top Level)") Then Return
                m_CreateGAPartsListTopLevelTool.Execute()
            Catch ex As Exception
                LogToolError("CreateGAPartsListTopLevel", ex)
                MsgBox("Error running Create GA Parts List (Top Level): " & ex.Message, MsgBoxStyle.Critical, "Create GA Parts List (Top Level)")
            End Try
        End Sub

        Private Sub m_CleanUpUnusedFilesButton_OnExecute(ByVal Context As NameValueMap) Handles m_CleanUpUnusedFilesButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.CleanUpUnusedFiles, "Clean Up Unused Files") Then Return
                m_CleanUpUnusedFilesTool.Execute()
            Catch ex As Exception
                LogToolError("CleanUpUnusedFiles", ex)
                MsgBox("Error running Clean Up Unused Files: " & ex.Message, MsgBoxStyle.Critical, "Clean Up Unused Files")
            End Try
        End Sub

        Private Sub m_LengthParameterExporterButton_OnExecute(ByVal Context As NameValueMap) Handles m_LengthParameterExporterButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.LengthParameterExporter, "Length Parameter Exporter") Then Return
                m_ParameterExporterTools.ExecuteLengthParameterExporter()
            Catch ex As Exception
                LogToolError("LengthParameterExporter", ex)
                MsgBox("Error running Length Parameter Exporter: " & ex.Message, MsgBoxStyle.Critical, "Length Parameter Exporter")
            End Try
        End Sub

        Private Sub m_Length2ParameterExporterButton_OnExecute(ByVal Context As NameValueMap) Handles m_Length2ParameterExporterButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.Length2ParameterExporter, "Length2 Parameter Exporter") Then Return
                m_ParameterExporterTools.ExecuteLength2ParameterExporter()
            Catch ex As Exception
                LogToolError("Length2ParameterExporter", ex)
                MsgBox("Error running Length2 Parameter Exporter: " & ex.Message, MsgBoxStyle.Critical, "Length2 Parameter Exporter")
            End Try
        End Sub

        Private Sub m_ThicknessParameterExporterButton_OnExecute(ByVal Context As NameValueMap) Handles m_ThicknessParameterExporterButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.ThicknessParameterExporter, "Thickness Parameter Exporter") Then Return
                m_ParameterExporterTools.ExecuteThicknessParameterExporter()
            Catch ex As Exception
                LogToolError("ThicknessParameterExporter", ex)
                MsgBox("Error running Thickness Parameter Exporter: " & ex.Message, MsgBoxStyle.Critical, "Thickness Parameter Exporter")
            End Try
        End Sub

        Private Sub m_FixNonPlatePartsButton_OnExecute(ByVal Context As NameValueMap) Handles m_FixNonPlatePartsButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.FixNonPlateParts, "Fix Non-Plate Parts") Then Return
                m_ParameterExporterTools.ExecuteFixNonPlateParts()
            Catch ex As Exception
                LogToolError("FixNonPlateParts", ex)
                MsgBox("Error running Fix Non-Plate Parts: " & ex.Message, MsgBoxStyle.Critical, "Fix Non-Plate Parts")
            End Try
        End Sub

        Private Sub m_FixSinglePartLength2Button_OnExecute(ByVal Context As NameValueMap) Handles m_FixSinglePartLength2Button.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.FixSinglePartLength2, "Fix Single Part Length2") Then Return
                m_ParameterExporterTools.ExecuteFixSinglePartLength2()
            Catch ex As Exception
                LogToolError("FixSinglePartLength2", ex)
                MsgBox("Error running Fix Single Part Length2: " & ex.Message, MsgBoxStyle.Critical, "Fix Single Part Length2")
            End Try
        End Sub

        Private Sub m_FixBOMPlateDimensionsButton_OnExecute(ByVal Context As NameValueMap) Handles m_FixBOMPlateDimensionsButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.FixBOMPlateDimensions, "Fix BOM Plate Dimensions") Then Return
                m_ParameterExporterTools.ExecuteFixBOMPlateDimensions()
            Catch ex As Exception
                LogToolError("FixBOMPlateDimensions", ex)
                MsgBox("Error running Fix BOM Plate Dimensions: " & ex.Message, MsgBoxStyle.Critical, "Fix BOM Plate Dimensions")
            End Try
        End Sub

        Private Sub m_ApplyPlateDescStockFormulaButton_OnExecute(ByVal Context As NameValueMap) Handles m_ApplyPlateDescStockFormulaButton.OnExecute
            Try
                If Not GuardToolEnabled(AddInTool.ApplyPlateDescStockFormula, "Apply Plate Desc/Stock Formula") Then Return
                m_ParameterExporterTools.ExecuteApplyPlateDescStockFormula()
            Catch ex As Exception
                LogToolError("ApplyPlateDescStockFormula", ex)
                MsgBox("Error running Apply Plate Desc/Stock Formula: " & ex.Message, MsgBoxStyle.Critical, "Apply Plate Desc/Stock Formula")
            End Try
        End Sub

        Private Sub LogToolError(ByVal context As String, ByVal ex As Exception)
            Try
                Dim logDir As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "Spectiv", "InventorAutomationSuite", "Logs")
                If Not System.IO.Directory.Exists(logDir) Then
                    System.IO.Directory.CreateDirectory(logDir)
                End If

                Dim logPath As String = System.IO.Path.Combine(logDir, "AddInTools.log")
                Dim message As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | " & context & " | " & ex.ToString() & System.Environment.NewLine
                System.IO.File.AppendAllText(logPath, message)
            Catch
            End Try
        End Sub


#End Region

    End Class

End Namespace
