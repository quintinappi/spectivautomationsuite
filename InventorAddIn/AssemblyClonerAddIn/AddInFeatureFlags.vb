Imports System.Collections.Generic

Namespace AssemblyClonerAddIn

    Public Module AddInFeatureFlags

        Private ReadOnly EnabledTools As HashSet(Of AddInTool) = New HashSet(Of AddInTool) From {
            AddInTool.CloneAssembly,
            AddInTool.AssemblyRenamer,
            AddInTool.TitleAutomationIDWOnly,
            AddInTool.SetViewIdentifier,
            AddInTool.SetViewScale,
            AddInTool.AutoBalloonLeaders,
            AddInTool.AutoDetailIDW,
            AddInTool.RegistryManagement,
            AddInTool.PopulateDWGRefFromPartsLists,
            AddInTool.PopulateDWGRefAutoPlaceMissingParts,
            AddInTool.CreateDXFForModelPlates,
            AddInTool.PartPlacer,
            AddInTool.CreateSheetPartsList,
            AddInTool.CreateGAPartsListTopLevel,
            AddInTool.CleanUpUnusedFiles,
            AddInTool.LengthParameterExporter,
            AddInTool.Length2ParameterExporter,
            AddInTool.ThicknessParameterExporter,
            AddInTool.FixNonPlateParts,
            AddInTool.FixSinglePartLength2,
            AddInTool.FixBOMPlateDimensions,
            AddInTool.ApplyPlateDescStockFormula
        }

        Public Function IsEnabled(ByVal tool As AddInTool) As Boolean
            Return EnabledTools.Contains(tool)
        End Function

        Public Function DisabledMessage(ByVal toolDisplayName As String) As String
            Return toolDisplayName & " is temporarily hidden during migration to the Inventor Add-In." & vbCrLf & vbCrLf &
                   "Status: command exists but is not enabled yet." & vbCrLf &
                                 "Enabled right now: Clone Assembly, Assembly Renamer, Title Automation (IDW), Set View Identifier, Set View Scale, Auto Ballooner, Auto Detail IDW, Registry Management, Populate DWG REF tools, CREATE DXF FOR MODEL PLATES, Place parts from open Assembly, Create Sheet Parts List, Create GA Parts List (Top Level), Clean Up Unused Files, Length Parameter Exporter, Length2 Parameter Exporter, Thickness Parameter Exporter, Fix Non-Plate Parts, Fix Single Part Length2, Fix BOM Plate Dimensions, Apply Plate Desc/Stock Formula."
        End Function

    End Module

End Namespace
