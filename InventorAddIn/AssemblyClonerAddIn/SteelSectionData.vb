' ==============================================================================
' SANS STEEL SECTION DATA - Comprehensive Steel Section Database
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' Contains all SANS (South African National Standards) steel section data
' including dimensions, properties, and backmark specifications for:
'   - Universal Beams (UB)
'   - Universal Columns (UC)
'   - Parallel Flange Channels (PFC)
'   - Tapered Flange Channels (TFC)
'   - Equal Angles (L)
'   - Unequal Angles
'   - Circular Hollow Sections (CHS)
'   - Square Hollow Sections (SHS)
'   - Rectangular Hollow Sections (RHS)
'   - Flat Bar (FL)
'
' All dimensions in millimeters (mm)
' ==============================================================================

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Represents a steel section with its properties
    ''' </summary>
    Public Class SteelSection
        Public Property SectionType As String       ' UB, UC, PFC, TFC, L, CHS, SHS, RHS, FL
        Public Property SectionName As String       ' e.g., "UC 203x203x46"
        Public Property DisplayName As String       ' e.g., "UC203x203x46"
        Public Property Description As String       ' e.g., "UC 203 x 203 x 46"
        Public Property Height As Double            ' Section depth (mm)
        Public Property Width As Double             ' Section/flange width (mm)
        Public Property WebThickness As Double      ' Web thickness (mm)
        Public Property FlangeThickness As Double   ' Flange thickness (mm)
        Public Property RootRadius As Double        ' Root radius/fillet (mm)
        Public Property Mass As Double              ' Mass per meter (kg/m)
        Public Property BackmarkY As Double         ' Backmark from center to hole (mm) - vertical
        Public Property BackmarkX As Double         ' Backmark from flange edge (mm) - horizontal
        Public Property HoleCount As Integer        ' Number of holes per row (2 or 4)
        Public Property BoltSize As String          ' M16, M20, M24
        Public Property HoleDiameter As Double      ' 18mm for M16, 22mm for M20, 26mm for M24

        ' For hollow sections
        Public Property WallThickness As Double     ' Wall thickness for CHS/SHS/RHS

        ' For angles
        Public Property LegA As Double              ' First leg (mm)
        Public Property LegB As Double              ' Second leg (mm) - same as LegA for equal angles

        ' For flat bar
        Public Property BarWidth As Double          ' Flat bar width (mm)
        Public Property BarThickness As Double      ' Flat bar thickness (mm)

        Public Sub New()
            ' Default values
            HoleCount = 4
            BoltSize = "M20"
            HoleDiameter = 22
        End Sub

        ''' <summary>
        ''' Get iProperty description format
        ''' </summary>
        Public Function GetIPropertyDescription() As String
            Select Case SectionType
                Case "UB", "UC"
                    Return SectionType & " " & CInt(Height) & "x" & CInt(Width) & "x" & CInt(Mass)
                Case "PFC"
                    Return "PFC " & CInt(Height) & "x" & CInt(Width)
                Case "TFC"
                    Return "TFC " & CInt(Height) & "x" & CInt(Width) & "x" & CInt(FlangeThickness)
                Case "L"
                    If LegA = LegB Then
                        Return "L " & CInt(LegA) & "x" & CInt(LegA) & "x" & CInt(FlangeThickness)
                    Else
                        Return "L " & CInt(LegA) & "x" & CInt(LegB) & "x" & CInt(FlangeThickness)
                    End If
                Case "CHS"
                    Return "CHS " & CInt(Height) & "x" & CInt(WallThickness)
                Case "SHS"
                    Return "SHS " & CInt(Height) & "x" & CInt(Height) & "x" & CInt(WallThickness)
                Case "RHS"
                    Return "RHS " & CInt(Height) & "x" & CInt(Width) & "x" & CInt(WallThickness)
                Case "FL"
                    Return "FL " & CInt(BarWidth) & "x" & CInt(BarThickness)
                Case Else
                    Return DisplayName
            End Select
        End Function
    End Class

    ''' <summary>
    ''' Static database of all SANS steel sections
    ''' </summary>
    Public Class SteelSectionDatabase

        Private Shared _sections As Dictionary(Of String, List(Of SteelSection))
        Private Shared _initialized As Boolean = False

        ''' <summary>
        ''' Get all section types available
        ''' </summary>
        Public Shared ReadOnly Property SectionTypes As String()
            Get
                Return New String() {"UC", "UB", "PFC", "TFC", "L", "CHS", "SHS", "RHS", "FL"}
            End Get
        End Property

        ''' <summary>
        ''' Get sections by type
        ''' </summary>
        Public Shared Function GetSectionsByType(sectionType As String) As List(Of SteelSection)
            EnsureInitialized()
            If _sections.ContainsKey(sectionType.ToUpper()) Then
                Return _sections(sectionType.ToUpper())
            End If
            Return New List(Of SteelSection)()
        End Function

        ''' <summary>
        ''' Get section by name
        ''' </summary>
        Public Shared Function GetSectionByName(sectionName As String) As SteelSection
            EnsureInitialized()
            For Each kvp In _sections
                For Each section In kvp.Value
                    If section.DisplayName.ToUpper() = sectionName.ToUpper() OrElse
                       section.SectionName.ToUpper() = sectionName.ToUpper() Then
                        Return section
                    End If
                Next
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Ensure database is initialized
        ''' </summary>
        Private Shared Sub EnsureInitialized()
            If Not _initialized Then
                Initialize()
                _initialized = True
            End If
        End Sub

        ''' <summary>
        ''' Initialize the steel section database
        ''' </summary>
        Private Shared Sub Initialize()
            _sections = New Dictionary(Of String, List(Of SteelSection))()

            ' Initialize all section types
            InitializeUniversalColumns()
            InitializeUniversalBeams()
            InitializeParallelFlangeChannels()
            InitializeTaperFlangeChannels()
            InitializeEqualAngles()
            InitializeCircularHollowSections()
            InitializeSquareHollowSections()
            InitializeFlatBar()
        End Sub

        ''' <summary>
        ''' Initialize Universal Columns (UC) - H-sections
        ''' </summary>
        Private Shared Sub InitializeUniversalColumns()
            Dim ucList As New List(Of SteelSection)()

            ' UC 152x152 series
            ucList.Add(CreateUC("UC152x152x23", 152, 152.2, 5.8, 6.8, 7.6, 23.0, 90, 45, 4))
            ucList.Add(CreateUC("UC152x152x30", 157.6, 152.9, 6.5, 9.4, 7.6, 30.0, 90, 45, 4))
            ucList.Add(CreateUC("UC152x152x37", 161.8, 154.4, 8.0, 11.5, 7.6, 37.0, 90, 45, 4))

            ' UC 203x203 series
            ucList.Add(CreateUC("UC203x203x46", 203.2, 203.6, 7.2, 11.0, 10.2, 46.0, 110, 55, 4))
            ucList.Add(CreateUC("UC203x203x52", 206.2, 204.3, 7.9, 12.5, 10.2, 52.0, 110, 55, 4))
            ucList.Add(CreateUC("UC203x203x60", 209.6, 205.8, 9.4, 14.2, 10.2, 60.0, 110, 55, 4))
            ucList.Add(CreateUC("UC203x203x71", 215.8, 206.4, 10.0, 17.3, 10.2, 71.0, 110, 55, 4))
            ucList.Add(CreateUC("UC203x203x86", 222.2, 209.1, 12.7, 20.5, 10.2, 86.0, 110, 55, 4))

            ' UC 254x254 series
            ucList.Add(CreateUC("UC254x254x73", 254.1, 254.6, 8.6, 14.2, 12.7, 73.0, 140, 70, 4))
            ucList.Add(CreateUC("UC254x254x89", 260.3, 256.3, 10.3, 17.3, 12.7, 89.0, 140, 70, 4))
            ucList.Add(CreateUC("UC254x254x107", 266.7, 258.8, 12.8, 20.5, 12.7, 107.0, 140, 70, 4))
            ucList.Add(CreateUC("UC254x254x132", 276.3, 261.3, 15.3, 25.3, 12.7, 132.0, 140, 70, 4))
            ucList.Add(CreateUC("UC254x254x167", 289.1, 265.2, 19.2, 31.7, 12.7, 167.0, 140, 70, 4))

            ' UC 305x305 series
            ucList.Add(CreateUC("UC305x305x97", 307.9, 305.3, 9.9, 15.4, 15.2, 97.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x118", 314.5, 307.4, 12.0, 18.7, 15.2, 118.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x137", 320.5, 309.2, 13.8, 21.7, 15.2, 137.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x158", 327.1, 311.2, 15.8, 25.0, 15.2, 158.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x198", 339.9, 314.5, 19.1, 31.4, 15.2, 198.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x240", 352.5, 318.4, 23.0, 37.7, 15.2, 240.0, 165, 80, 4))
            ucList.Add(CreateUC("UC305x305x283", 365.3, 322.2, 26.8, 44.1, 15.2, 283.0, 165, 80, 4))

            ' UC 356x368 series
            ucList.Add(CreateUC("UC356x368x129", 355.6, 368.6, 10.4, 17.5, 15.2, 129.0, 190, 95, 4))
            ucList.Add(CreateUC("UC356x368x153", 362.0, 370.5, 12.3, 20.7, 15.2, 153.0, 190, 95, 4))
            ucList.Add(CreateUC("UC356x368x177", 368.2, 372.6, 14.4, 23.8, 15.2, 177.0, 190, 95, 4))
            ucList.Add(CreateUC("UC356x368x202", 374.6, 374.7, 16.5, 27.0, 15.2, 202.0, 190, 95, 4))

            _sections.Add("UC", ucList)
        End Sub

        ''' <summary>
        ''' Initialize Universal Beams (UB) - I-sections
        ''' </summary>
        Private Shared Sub InitializeUniversalBeams()
            Dim ubList As New List(Of SteelSection)()

            ' UB 127x76 series
            ubList.Add(CreateUB("UB127x76x13", 127.0, 76.0, 4.0, 7.6, 7.6, 13.0, 55, 38, 2))

            ' UB 152x89 series
            ubList.Add(CreateUB("UB152x89x16", 152.4, 88.7, 4.5, 7.7, 7.6, 16.0, 60, 44, 2))

            ' UB 178x102 series
            ubList.Add(CreateUB("UB178x102x19", 177.8, 101.2, 4.8, 7.9, 7.6, 19.0, 65, 50, 2))

            ' UB 203x102 series
            ubList.Add(CreateUB("UB203x102x23", 203.2, 101.8, 5.4, 9.3, 7.6, 23.0, 70, 50, 2))

            ' UB 203x133 series
            ubList.Add(CreateUB("UB203x133x25", 203.2, 133.2, 5.7, 7.8, 7.6, 25.0, 75, 65, 4))
            ubList.Add(CreateUB("UB203x133x30", 206.8, 133.9, 6.4, 9.6, 7.6, 30.0, 75, 65, 4))

            ' UB 254x102 series
            ubList.Add(CreateUB("UB254x102x22", 254.0, 101.6, 5.7, 6.8, 7.6, 22.0, 80, 50, 2))
            ubList.Add(CreateUB("UB254x102x25", 257.2, 101.9, 6.0, 8.4, 7.6, 25.0, 80, 50, 2))
            ubList.Add(CreateUB("UB254x102x28", 260.4, 102.2, 6.3, 10.0, 7.6, 28.0, 80, 50, 2))

            ' UB 254x146 series
            ubList.Add(CreateUB("UB254x146x31", 251.4, 146.1, 6.0, 8.6, 7.6, 31.0, 90, 70, 4))
            ubList.Add(CreateUB("UB254x146x37", 256.0, 146.4, 6.3, 10.9, 7.6, 37.0, 90, 70, 4))
            ubList.Add(CreateUB("UB254x146x43", 259.6, 147.3, 7.2, 12.7, 7.6, 43.0, 90, 70, 4))

            ' UB 305x102 series
            ubList.Add(CreateUB("UB305x102x25", 305.1, 101.6, 5.8, 7.0, 7.6, 25.0, 90, 50, 2))
            ubList.Add(CreateUB("UB305x102x28", 308.7, 101.8, 6.0, 8.8, 7.6, 28.0, 90, 50, 2))
            ubList.Add(CreateUB("UB305x102x33", 312.7, 102.4, 6.6, 10.8, 7.6, 33.0, 90, 50, 2))

            ' UB 305x127 series
            ubList.Add(CreateUB("UB305x127x37", 304.4, 123.4, 7.1, 10.7, 8.9, 37.0, 100, 60, 2))
            ubList.Add(CreateUB("UB305x127x42", 307.2, 124.3, 8.0, 12.1, 8.9, 42.0, 100, 60, 2))
            ubList.Add(CreateUB("UB305x127x48", 311.0, 125.3, 9.0, 14.0, 8.9, 48.0, 100, 60, 2))

            ' UB 305x165 series
            ubList.Add(CreateUB("UB305x165x40", 303.4, 165.0, 6.0, 10.2, 8.9, 40.0, 110, 80, 4))
            ubList.Add(CreateUB("UB305x165x46", 306.6, 165.7, 6.7, 11.8, 8.9, 46.0, 110, 80, 4))
            ubList.Add(CreateUB("UB305x165x54", 310.4, 166.9, 7.9, 13.7, 8.9, 54.0, 110, 80, 4))

            ' UB 356x127 series
            ubList.Add(CreateUB("UB356x127x33", 349.0, 125.4, 6.0, 8.5, 10.2, 33.0, 110, 60, 2))
            ubList.Add(CreateUB("UB356x127x39", 353.4, 126.0, 6.6, 10.7, 10.2, 39.0, 110, 60, 2))

            ' UB 356x171 series
            ubList.Add(CreateUB("UB356x171x45", 351.4, 171.1, 7.0, 9.7, 10.2, 45.0, 120, 85, 4))
            ubList.Add(CreateUB("UB356x171x51", 355.0, 171.5, 7.4, 11.5, 10.2, 51.0, 120, 85, 4))
            ubList.Add(CreateUB("UB356x171x57", 358.0, 172.2, 8.1, 13.0, 10.2, 57.0, 120, 85, 4))
            ubList.Add(CreateUB("UB356x171x67", 363.4, 173.2, 9.1, 15.7, 10.2, 67.0, 120, 85, 4))

            ' UB 406x140 series
            ubList.Add(CreateUB("UB406x140x39", 398.0, 141.8, 6.4, 8.6, 10.2, 39.0, 120, 70, 2))
            ubList.Add(CreateUB("UB406x140x46", 403.2, 142.2, 6.8, 11.2, 10.2, 46.0, 120, 70, 2))

            ' UB 406x178 series
            ubList.Add(CreateUB("UB406x178x54", 402.6, 177.7, 7.7, 10.9, 10.2, 54.0, 130, 90, 4))
            ubList.Add(CreateUB("UB406x178x60", 406.4, 177.9, 7.9, 12.8, 10.2, 60.0, 130, 90, 4))
            ubList.Add(CreateUB("UB406x178x67", 409.4, 178.8, 8.8, 14.3, 10.2, 67.0, 130, 90, 4))
            ubList.Add(CreateUB("UB406x178x74", 412.8, 179.5, 9.5, 16.0, 10.2, 74.0, 130, 90, 4))

            ' UB 457x152 series
            ubList.Add(CreateUB("UB457x152x52", 449.8, 152.4, 7.6, 10.9, 10.2, 52.0, 130, 75, 2))
            ubList.Add(CreateUB("UB457x152x60", 454.6, 152.9, 8.1, 13.3, 10.2, 60.0, 130, 75, 2))
            ubList.Add(CreateUB("UB457x152x67", 458.0, 153.8, 9.0, 15.0, 10.2, 67.0, 130, 75, 2))
            ubList.Add(CreateUB("UB457x152x74", 462.0, 154.4, 9.6, 17.0, 10.2, 74.0, 130, 75, 2))
            ubList.Add(CreateUB("UB457x152x82", 465.8, 155.3, 10.5, 18.9, 10.2, 82.0, 130, 75, 2))

            ' UB 457x191 series
            ubList.Add(CreateUB("UB457x191x67", 453.4, 189.9, 8.5, 12.7, 10.2, 67.0, 140, 95, 4))
            ubList.Add(CreateUB("UB457x191x74", 457.0, 190.4, 9.0, 14.5, 10.2, 74.0, 140, 95, 4))
            ubList.Add(CreateUB("UB457x191x82", 460.0, 191.3, 9.9, 16.0, 10.2, 82.0, 140, 95, 4))
            ubList.Add(CreateUB("UB457x191x89", 463.4, 191.9, 10.5, 17.7, 10.2, 89.0, 140, 95, 4))
            ubList.Add(CreateUB("UB457x191x98", 467.2, 192.8, 11.4, 19.6, 10.2, 98.0, 140, 95, 4))

            ' UB 533x210 series
            ubList.Add(CreateUB("UB533x210x82", 528.3, 208.8, 9.6, 13.2, 12.7, 82.0, 150, 105, 4))
            ubList.Add(CreateUB("UB533x210x92", 533.1, 209.3, 10.1, 15.6, 12.7, 92.0, 150, 105, 4))
            ubList.Add(CreateUB("UB533x210x101", 536.7, 210.0, 10.8, 17.4, 12.7, 101.0, 150, 105, 4))
            ubList.Add(CreateUB("UB533x210x109", 539.5, 210.8, 11.6, 18.8, 12.7, 109.0, 150, 105, 4))
            ubList.Add(CreateUB("UB533x210x122", 544.5, 211.9, 12.7, 21.3, 12.7, 122.0, 150, 105, 4))

            ' UB 610x229 series
            ubList.Add(CreateUB("UB610x229x101", 602.6, 227.6, 10.5, 14.8, 12.7, 101.0, 160, 115, 4))
            ubList.Add(CreateUB("UB610x229x113", 607.6, 228.2, 11.1, 17.3, 12.7, 113.0, 160, 115, 4))
            ubList.Add(CreateUB("UB610x229x125", 612.2, 229.0, 11.9, 19.6, 12.7, 125.0, 160, 115, 4))
            ubList.Add(CreateUB("UB610x229x140", 617.2, 230.2, 13.1, 22.1, 12.7, 140.0, 160, 115, 4))

            ' UB 610x305 series
            ubList.Add(CreateUB("UB610x305x149", 612.4, 304.8, 11.8, 19.7, 16.5, 149.0, 170, 150, 4))
            ubList.Add(CreateUB("UB610x305x179", 620.2, 307.1, 14.1, 23.6, 16.5, 179.0, 170, 150, 4))
            ubList.Add(CreateUB("UB610x305x238", 635.8, 311.4, 18.4, 31.4, 16.5, 238.0, 170, 150, 4))

            _sections.Add("UB", ubList)
        End Sub

        ''' <summary>
        ''' Initialize Parallel Flange Channels (PFC)
        ''' </summary>
        Private Shared Sub InitializeParallelFlangeChannels()
            Dim pfcList As New List(Of SteelSection)()

            ' PFC series - Height x Width
            pfcList.Add(CreatePFC("PFC100x50", 100.0, 50.0, 5.0, 8.5, 12.0, 10.2, 35))
            pfcList.Add(CreatePFC("PFC125x65", 125.0, 65.0, 5.5, 9.5, 14.2, 12.0, 40))
            pfcList.Add(CreatePFC("PFC150x75", 150.0, 75.0, 5.5, 10.0, 17.9, 12.0, 45))
            pfcList.Add(CreatePFC("PFC150x90", 150.0, 90.0, 6.5, 12.0, 24.0, 12.0, 50))
            pfcList.Add(CreatePFC("PFC180x75", 180.0, 75.0, 6.0, 10.5, 20.3, 12.0, 45))
            pfcList.Add(CreatePFC("PFC180x90", 180.0, 90.0, 6.5, 12.5, 26.1, 12.0, 50))
            pfcList.Add(CreatePFC("PFC200x75", 200.0, 75.0, 6.0, 11.5, 23.4, 15.0, 45))
            pfcList.Add(CreatePFC("PFC200x90", 200.0, 90.0, 7.0, 14.0, 29.4, 15.0, 50))
            pfcList.Add(CreatePFC("PFC230x75", 230.0, 75.0, 6.5, 12.5, 25.7, 15.0, 45))
            pfcList.Add(CreatePFC("PFC230x90", 230.0, 90.0, 7.5, 14.0, 32.2, 15.0, 50))
            pfcList.Add(CreatePFC("PFC260x75", 260.0, 75.0, 7.0, 13.5, 28.0, 15.0, 45))
            pfcList.Add(CreatePFC("PFC260x90", 260.0, 90.0, 8.0, 14.0, 34.8, 15.0, 50))
            pfcList.Add(CreatePFC("PFC300x90", 300.0, 90.0, 9.0, 16.0, 41.8, 18.0, 50))
            pfcList.Add(CreatePFC("PFC300x100", 300.0, 100.0, 9.0, 16.5, 45.5, 18.0, 55))
            pfcList.Add(CreatePFC("PFC380x100", 380.0, 100.0, 9.5, 17.5, 54.0, 18.0, 55))
            pfcList.Add(CreatePFC("PFC400x100", 400.0, 100.0, 10.5, 18.0, 59.5, 18.0, 55))

            _sections.Add("PFC", pfcList)
        End Sub

        ''' <summary>
        ''' Initialize Tapered Flange Channels (TFC)
        ''' </summary>
        Private Shared Sub InitializeTaperFlangeChannels()
            Dim tfcList As New List(Of SteelSection)()

            ' TFC series - Height x Width x Flange
            tfcList.Add(CreateTFC("TFC76x38", 76.2, 38.1, 5.1, 6.8, 6.7, 40))
            tfcList.Add(CreateTFC("TFC102x51", 101.6, 50.8, 6.1, 7.6, 10.4, 45))
            tfcList.Add(CreateTFC("TFC127x64", 127.0, 63.5, 6.4, 9.1, 14.9, 50))
            tfcList.Add(CreateTFC("TFC152x76", 152.4, 76.2, 6.4, 9.0, 17.9, 55))
            tfcList.Add(CreateTFC("TFC152x89", 152.4, 88.9, 7.1, 11.2, 23.8, 60))
            tfcList.Add(CreateTFC("TFC178x54", 177.8, 53.8, 7.0, 10.3, 20.3, 40))
            tfcList.Add(CreateTFC("TFC178x76", 177.8, 76.2, 6.6, 10.3, 24.0, 55))
            tfcList.Add(CreateTFC("TFC178x89", 177.8, 88.9, 7.6, 12.3, 30.6, 60))
            tfcList.Add(CreateTFC("TFC203x76", 203.2, 76.2, 7.1, 11.2, 29.8, 55))
            tfcList.Add(CreateTFC("TFC203x89", 203.2, 88.9, 8.1, 13.7, 37.9, 60))
            tfcList.Add(CreateTFC("TFC229x76", 228.6, 76.2, 7.6, 12.4, 33.2, 55))
            tfcList.Add(CreateTFC("TFC229x89", 228.6, 88.9, 8.6, 14.2, 41.7, 60))
            tfcList.Add(CreateTFC("TFC254x76", 254.0, 76.2, 8.1, 13.2, 37.2, 55))
            tfcList.Add(CreateTFC("TFC254x89", 254.0, 88.9, 9.1, 15.2, 44.4, 60))
            tfcList.Add(CreateTFC("TFC305x89", 304.8, 88.9, 10.2, 16.5, 52.0, 60))
            tfcList.Add(CreateTFC("TFC305x102", 304.8, 101.6, 10.8, 17.1, 58.1, 65))
            tfcList.Add(CreateTFC("TFC381x102", 381.0, 101.6, 11.2, 17.5, 66.0, 65))

            _sections.Add("TFC", tfcList)
        End Sub

        ''' <summary>
        ''' Initialize Equal Angles (L)
        ''' </summary>
        Private Shared Sub InitializeEqualAngles()
            Dim angleList As New List(Of SteelSection)()

            ' Equal angles - LegxLegxThickness
            angleList.Add(CreateAngle("L25x25x3", 25, 25, 3, 1.12, 15))
            angleList.Add(CreateAngle("L25x25x4", 25, 25, 4, 1.46, 15))
            angleList.Add(CreateAngle("L25x25x5", 25, 25, 5, 1.77, 15))
            angleList.Add(CreateAngle("L30x30x3", 30, 30, 3, 1.37, 18))
            angleList.Add(CreateAngle("L30x30x4", 30, 30, 4, 1.79, 18))
            angleList.Add(CreateAngle("L30x30x5", 30, 30, 5, 2.18, 18))
            angleList.Add(CreateAngle("L40x40x3", 40, 40, 3, 1.85, 25))
            angleList.Add(CreateAngle("L40x40x4", 40, 40, 4, 2.42, 25))
            angleList.Add(CreateAngle("L40x40x5", 40, 40, 5, 2.97, 25))
            angleList.Add(CreateAngle("L40x40x6", 40, 40, 6, 3.52, 25))
            angleList.Add(CreateAngle("L50x50x5", 50, 50, 5, 3.77, 30))
            angleList.Add(CreateAngle("L50x50x6", 50, 50, 6, 4.47, 30))
            angleList.Add(CreateAngle("L50x50x8", 50, 50, 8, 5.82, 30))
            angleList.Add(CreateAngle("L60x60x5", 60, 60, 5, 4.57, 35))
            angleList.Add(CreateAngle("L60x60x6", 60, 60, 6, 5.42, 35))
            angleList.Add(CreateAngle("L60x60x8", 60, 60, 8, 7.09, 35))
            angleList.Add(CreateAngle("L60x60x10", 60, 60, 10, 8.69, 35))
            angleList.Add(CreateAngle("L65x65x6", 65, 65, 6, 5.91, 40))
            angleList.Add(CreateAngle("L65x65x8", 65, 65, 8, 7.73, 40))
            angleList.Add(CreateAngle("L65x65x10", 65, 65, 10, 9.49, 40))
            angleList.Add(CreateAngle("L70x70x6", 70, 70, 6, 6.38, 45))
            angleList.Add(CreateAngle("L70x70x8", 70, 70, 8, 8.36, 45))
            angleList.Add(CreateAngle("L70x70x10", 70, 70, 10, 10.3, 45))
            angleList.Add(CreateAngle("L75x75x6", 75, 75, 6, 6.85, 45))
            angleList.Add(CreateAngle("L75x75x8", 75, 75, 8, 8.99, 45))
            angleList.Add(CreateAngle("L75x75x10", 75, 75, 10, 11.0, 45))
            angleList.Add(CreateAngle("L80x80x8", 80, 80, 8, 9.63, 50))
            angleList.Add(CreateAngle("L80x80x10", 80, 80, 10, 11.9, 50))
            angleList.Add(CreateAngle("L90x90x8", 90, 90, 8, 10.9, 55))
            angleList.Add(CreateAngle("L90x90x10", 90, 90, 10, 13.4, 55))
            angleList.Add(CreateAngle("L90x90x12", 90, 90, 12, 15.9, 55))
            angleList.Add(CreateAngle("L100x100x8", 100, 100, 8, 12.2, 60))
            angleList.Add(CreateAngle("L100x100x10", 100, 100, 10, 15.0, 60))
            angleList.Add(CreateAngle("L100x100x12", 100, 100, 12, 17.8, 60))
            angleList.Add(CreateAngle("L100x100x15", 100, 100, 15, 21.9, 60))
            angleList.Add(CreateAngle("L120x120x10", 120, 120, 10, 18.2, 70))
            angleList.Add(CreateAngle("L120x120x12", 120, 120, 12, 21.6, 70))
            angleList.Add(CreateAngle("L120x120x15", 120, 120, 15, 26.6, 70))
            angleList.Add(CreateAngle("L150x150x10", 150, 150, 10, 23.0, 90))
            angleList.Add(CreateAngle("L150x150x12", 150, 150, 12, 27.3, 90))
            angleList.Add(CreateAngle("L150x150x15", 150, 150, 15, 33.8, 90))
            angleList.Add(CreateAngle("L150x150x18", 150, 150, 18, 40.1, 90))
            angleList.Add(CreateAngle("L200x200x16", 200, 200, 16, 48.5, 120))
            angleList.Add(CreateAngle("L200x200x20", 200, 200, 20, 59.9, 120))
            angleList.Add(CreateAngle("L200x200x24", 200, 200, 24, 71.1, 120))

            _sections.Add("L", angleList)
        End Sub

        ''' <summary>
        ''' Initialize Circular Hollow Sections (CHS)
        ''' </summary>
        Private Shared Sub InitializeCircularHollowSections()
            Dim chsList As New List(Of SteelSection)()

            ' CHS series - Diameter x Wall
            chsList.Add(CreateCHS("CHS33.7x2.0", 33.7, 2.0, 1.56))
            chsList.Add(CreateCHS("CHS33.7x3.0", 33.7, 3.0, 2.27))
            chsList.Add(CreateCHS("CHS42.4x2.0", 42.4, 2.0, 1.99))
            chsList.Add(CreateCHS("CHS42.4x3.0", 42.4, 3.0, 2.91))
            chsList.Add(CreateCHS("CHS48.3x2.5", 48.3, 2.5, 2.82))
            chsList.Add(CreateCHS("CHS48.3x3.0", 48.3, 3.0, 3.35))
            chsList.Add(CreateCHS("CHS60.3x2.5", 60.3, 2.5, 3.56))
            chsList.Add(CreateCHS("CHS60.3x3.0", 60.3, 3.0, 4.24))
            chsList.Add(CreateCHS("CHS60.3x4.0", 60.3, 4.0, 5.55))
            chsList.Add(CreateCHS("CHS76.1x3.0", 76.1, 3.0, 5.41))
            chsList.Add(CreateCHS("CHS76.1x4.0", 76.1, 4.0, 7.11))
            chsList.Add(CreateCHS("CHS76.1x5.0", 76.1, 5.0, 8.77))
            chsList.Add(CreateCHS("CHS88.9x3.0", 88.9, 3.0, 6.36))
            chsList.Add(CreateCHS("CHS88.9x4.0", 88.9, 4.0, 8.38))
            chsList.Add(CreateCHS("CHS88.9x5.0", 88.9, 5.0, 10.4))
            chsList.Add(CreateCHS("CHS114.3x3.0", 114.3, 3.0, 8.23))
            chsList.Add(CreateCHS("CHS114.3x4.0", 114.3, 4.0, 10.9))
            chsList.Add(CreateCHS("CHS114.3x5.0", 114.3, 5.0, 13.5))
            chsList.Add(CreateCHS("CHS114.3x6.0", 114.3, 6.0, 16.0))
            chsList.Add(CreateCHS("CHS139.7x4.0", 139.7, 4.0, 13.4))
            chsList.Add(CreateCHS("CHS139.7x5.0", 139.7, 5.0, 16.6))
            chsList.Add(CreateCHS("CHS139.7x6.0", 139.7, 6.0, 19.8))
            chsList.Add(CreateCHS("CHS168.3x5.0", 168.3, 5.0, 20.1))
            chsList.Add(CreateCHS("CHS168.3x6.0", 168.3, 6.0, 24.0))
            chsList.Add(CreateCHS("CHS168.3x8.0", 168.3, 8.0, 31.6))

            _sections.Add("CHS", chsList)
        End Sub

        ''' <summary>
        ''' Initialize Square Hollow Sections (SHS)
        ''' </summary>
        Private Shared Sub InitializeSquareHollowSections()
            Dim shsList As New List(Of SteelSection)()

            ' SHS series - Size x Wall
            shsList.Add(CreateSHS("SHS25x25x2.0", 25, 2.0, 1.36))
            shsList.Add(CreateSHS("SHS30x30x2.0", 30, 2.0, 1.68))
            shsList.Add(CreateSHS("SHS30x30x2.5", 30, 2.5, 2.03))
            shsList.Add(CreateSHS("SHS40x40x2.0", 40, 2.0, 2.31))
            shsList.Add(CreateSHS("SHS40x40x2.5", 40, 2.5, 2.82))
            shsList.Add(CreateSHS("SHS40x40x3.0", 40, 3.0, 3.30))
            shsList.Add(CreateSHS("SHS50x50x2.5", 50, 2.5, 3.60))
            shsList.Add(CreateSHS("SHS50x50x3.0", 50, 3.0, 4.25))
            shsList.Add(CreateSHS("SHS50x50x4.0", 50, 4.0, 5.45))
            shsList.Add(CreateSHS("SHS60x60x3.0", 60, 3.0, 5.19))
            shsList.Add(CreateSHS("SHS60x60x4.0", 60, 4.0, 6.71))
            shsList.Add(CreateSHS("SHS60x60x5.0", 60, 5.0, 8.13))
            shsList.Add(CreateSHS("SHS75x75x3.0", 75, 3.0, 6.60))
            shsList.Add(CreateSHS("SHS75x75x4.0", 75, 4.0, 8.59))
            shsList.Add(CreateSHS("SHS75x75x5.0", 75, 5.0, 10.5))
            shsList.Add(CreateSHS("SHS75x75x6.0", 75, 6.0, 12.3))
            shsList.Add(CreateSHS("SHS100x100x3.0", 100, 3.0, 8.96))
            shsList.Add(CreateSHS("SHS100x100x4.0", 100, 4.0, 11.7))
            shsList.Add(CreateSHS("SHS100x100x5.0", 100, 5.0, 14.4))
            shsList.Add(CreateSHS("SHS100x100x6.0", 100, 6.0, 17.0))
            shsList.Add(CreateSHS("SHS100x100x8.0", 100, 8.0, 21.7))
            shsList.Add(CreateSHS("SHS150x150x5.0", 150, 5.0, 22.3))
            shsList.Add(CreateSHS("SHS150x150x6.0", 150, 6.0, 26.4))
            shsList.Add(CreateSHS("SHS150x150x8.0", 150, 8.0, 34.2))
            shsList.Add(CreateSHS("SHS150x150x10.0", 150, 10.0, 41.3))
            shsList.Add(CreateSHS("SHS200x200x6.0", 200, 6.0, 35.8))
            shsList.Add(CreateSHS("SHS200x200x8.0", 200, 8.0, 46.7))
            shsList.Add(CreateSHS("SHS200x200x10.0", 200, 10.0, 56.8))

            _sections.Add("SHS", shsList)
        End Sub

        ''' <summary>
        ''' Initialize Flat Bar (FL)
        ''' </summary>
        Private Shared Sub InitializeFlatBar()
            Dim flList As New List(Of SteelSection)()

            ' Common flat bar sizes - Width x Thickness
            Dim widths() As Double = {20, 25, 30, 40, 50, 60, 65, 75, 80, 100, 120, 150, 200}
            Dim thicknesses() As Double = {3, 5, 6, 8, 10, 12, 15, 20, 25, 30}

            For Each w In widths
                For Each t In thicknesses
                    If t <= w / 2 Then ' Reasonable aspect ratio
                        Dim mass As Double = (w * t * 7.85) / 1000 ' kg/m
                        flList.Add(CreateFlatBar("FL" & CInt(w) & "x" & CInt(t), w, t, mass))
                    End If
                Next
            Next

            _sections.Add("FL", flList)
        End Sub

#Region "Helper Functions"

        Private Shared Function CreateUC(name As String, h As Double, w As Double, tw As Double, tf As Double, r As Double, mass As Double, backmarkY As Double, backmarkX As Double, holeCount As Integer) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "UC"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = h
            section.Width = w
            section.WebThickness = tw
            section.FlangeThickness = tf
            section.RootRadius = r
            section.Mass = mass
            section.BackmarkY = backmarkY
            section.BackmarkX = backmarkX
            section.HoleCount = holeCount
            section.BoltSize = "M20"
            section.HoleDiameter = 22
            Return section
        End Function

        Private Shared Function CreateUB(name As String, h As Double, w As Double, tw As Double, tf As Double, r As Double, mass As Double, backmarkY As Double, backmarkX As Double, holeCount As Integer) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "UB"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = h
            section.Width = w
            section.WebThickness = tw
            section.FlangeThickness = tf
            section.RootRadius = r
            section.Mass = mass
            section.BackmarkY = backmarkY
            section.BackmarkX = backmarkX
            section.HoleCount = holeCount
            section.BoltSize = If(w >= 150, "M20", "M16")
            section.HoleDiameter = If(w >= 150, 22, 18)
            Return section
        End Function

        Private Shared Function CreatePFC(name As String, h As Double, w As Double, tw As Double, tf As Double, mass As Double, r As Double, backmark As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "PFC"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = h
            section.Width = w
            section.WebThickness = tw
            section.FlangeThickness = tf
            section.RootRadius = r
            section.Mass = mass
            section.BackmarkY = 0 ' Channels have single row of holes
            section.BackmarkX = backmark
            section.HoleCount = 2 ' 2 holes for channels (one per end, single row)
            section.BoltSize = "M16"
            section.HoleDiameter = 18
            Return section
        End Function

        Private Shared Function CreateTFC(name As String, h As Double, w As Double, tw As Double, tf As Double, mass As Double, backmark As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "TFC"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = h
            section.Width = w
            section.WebThickness = tw
            section.FlangeThickness = tf
            section.RootRadius = 0
            section.Mass = mass
            section.BackmarkY = 0
            section.BackmarkX = backmark
            section.HoleCount = 2 ' 2 holes for channels
            section.BoltSize = "M16"
            section.HoleDiameter = 18
            Return section
        End Function

        Private Shared Function CreateAngle(name As String, legA As Double, legB As Double, t As Double, mass As Double, backmark As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "L"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.LegA = legA
            section.LegB = legB
            section.Height = legA
            section.Width = legB
            section.FlangeThickness = t
            section.Mass = mass
            section.BackmarkX = backmark
            section.BackmarkY = backmark
            section.HoleCount = 2 ' 2 holes for angles
            section.BoltSize = If(legA >= 100, "M20", "M16")
            section.HoleDiameter = If(legA >= 100, 22, 18)
            Return section
        End Function

        Private Shared Function CreateCHS(name As String, diameter As Double, wall As Double, mass As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "CHS"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = diameter
            section.Width = diameter
            section.WallThickness = wall
            section.Mass = mass
            section.HoleCount = 0 ' No holes for hollow sections (welded)
            Return section
        End Function

        Private Shared Function CreateSHS(name As String, size As Double, wall As Double, mass As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "SHS"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.Height = size
            section.Width = size
            section.WallThickness = wall
            section.Mass = mass
            section.HoleCount = 0 ' No holes for hollow sections (welded)
            Return section
        End Function

        Private Shared Function CreateFlatBar(name As String, width As Double, thickness As Double, mass As Double) As SteelSection
            Dim section As New SteelSection()
            section.SectionType = "FL"
            section.SectionName = name
            section.DisplayName = name.Replace("x", " x ")
            section.BarWidth = width
            section.BarThickness = thickness
            section.Height = width
            section.Width = thickness
            section.Mass = mass
            section.HoleCount = 0 ' Flat bar holes depend on use
            Return section
        End Function

#End Region

    End Class

End Namespace
