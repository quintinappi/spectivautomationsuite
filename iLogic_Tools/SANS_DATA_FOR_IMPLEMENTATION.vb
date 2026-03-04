' ==============================================================================
' SANS STEEL SECTION DATA - Extracted from Beam_Factory.ipt
' ==============================================================================
' This data can be directly copied into BeamAssemblyGenerator.vb
' Source: iLogic rule in Beam_Factory.ipt
' Extracted: December 15, 2025
' ==============================================================================

Public Class SteelSectionDatabase

    ''' <summary>
    ''' SANS I-Beams (UB sections) - 27 standard sizes
    ''' </summary>
    Public Shared Function GetSANS_IBeams() As List(Of SteelSection)
        Dim sections As New List(Of SteelSection)

        ' Format: New SteelSection(type, designation, height_mm, width_mm, web_thickness_mm, flange_thickness_mm, root_radius_mm, mass_kg_per_m)

        sections.Add(New SteelSection("UB", "203 x 133 x 25", 203.2, 133.2, 5.7, 7.8, 7.6, 25))
        sections.Add(New SteelSection("UB", "203 x 133 x 30", 206.8, 133.9, 6.4, 9.6, 7.6, 30))
        sections.Add(New SteelSection("UB", "254 x 146 x 31", 251.4, 146.1, 6.0, 8.6, 7.6, 31))
        sections.Add(New SteelSection("UB", "254 x 146 x 37", 256.0, 146.4, 6.3, 10.9, 7.6, 37))
        sections.Add(New SteelSection("UB", "254 x 146 x 43", 259.6, 147.3, 7.2, 12.7, 7.6, 43))
        sections.Add(New SteelSection("UB", "305 x 102 x 25", 305.1, 101.6, 5.8, 6.7, 7.6, 25))
        sections.Add(New SteelSection("UB", "305 x 102 x 28", 308.9, 101.9, 6.1, 8.9, 7.6, 28))
        sections.Add(New SteelSection("UB", "305 x 102 x 33", 312.7, 102.4, 6.6, 10.8, 7.6, 33))
        sections.Add(New SteelSection("UB", "305 x 165 x 40", 303.8, 165.1, 6.1, 10.2, 8.9, 40))
        sections.Add(New SteelSection("UB", "305 x 165 x 46", 307.1, 165.7, 6.7, 11.8, 8.9, 46))
        sections.Add(New SteelSection("UB", "305 x 165 x 54", 310.9, 166.8, 7.7, 13.7, 8.9, 54))
        sections.Add(New SteelSection("UB", "356 x 171 x 45", 352.0, 171.0, 6.9, 9.7, 10.2, 45))
        sections.Add(New SteelSection("UB", "356 x 171 x 51", 355.6, 171.5, 7.3, 11.5, 10.2, 51))
        sections.Add(New SteelSection("UB", "356 x 171 x 57", 358.6, 172.1, 8.0, 13.0, 10.2, 57))
        sections.Add(New SteelSection("UB", "356 x 171 x 67", 364.0, 173.2, 9.1, 15.7, 10.2, 67))
        sections.Add(New SteelSection("UB", "406 x 140 x 39", 397.3, 141.8, 6.3, 8.6, 10.2, 39))
        sections.Add(New SteelSection("UB", "406 x 140 x 46", 402.3, 142.4, 6.9, 11.2, 10.2, 46))
        sections.Add(New SteelSection("UB", "406 x 178 x 54", 402.6, 177.6, 7.6, 10.9, 10.2, 54))
        sections.Add(New SteelSection("UB", "406 x 178 x 60", 406.4, 177.8, 7.8, 12.8, 10.2, 60))
        sections.Add(New SteelSection("UB", "406 x 178 x 67", 409.4, 178.8, 8.8, 14.3, 10.2, 67))
        sections.Add(New SteelSection("UB", "406 x 178 x 74", 412.8, 179.7, 9.7, 16.0, 10.2, 74))
        sections.Add(New SteelSection("UB", "457 x 191 x 67", 453.6, 189.9, 8.5, 12.7, 10.2, 67))
        sections.Add(New SteelSection("UB", "457 x 191 x 74", 457.2, 190.5, 9.1, 14.5, 10.2, 74))
        sections.Add(New SteelSection("UB", "457 x 191 x 82", 460.2, 191.3, 9.9, 16.0, 10.2, 82))
        sections.Add(New SteelSection("UB", "457 x 191 x 89", 463.6, 192.0, 10.6, 17.7, 10.2, 89))
        sections.Add(New SteelSection("UB", "457 x 191 x 98", 467.6, 192.8, 11.4, 19.6, 10.2, 98))
        sections.Add(New SteelSection("UB", "533 x 210 x 82", 528.3, 208.7, 9.6, 13.2, 12.7, 82))
        sections.Add(New SteelSection("UB", "533 x 210 x 92", 533.1, 209.3, 10.2, 15.6, 12.7, 92))
        sections.Add(New SteelSection("UB", "533 x 210 x 101", 536.7, 210.1, 10.9, 17.4, 12.7, 101))
        sections.Add(New SteelSection("UB", "533 x 210 x 109", 539.5, 210.7, 11.6, 18.8, 12.7, 109))
        sections.Add(New SteelSection("UB", "533 x 210 x 122", 544.6, 211.9, 12.8, 21.3, 12.7, 122))

        Return sections
    End Function

    ''' <summary>
    ''' SANS H-Beams (UC sections) - 17 standard sizes
    ''' </summary>
    Public Shared Function GetSANS_HBeams() As List(Of SteelSection)
        Dim sections As New List(Of SteelSection)

        ' Format: New SteelSection(type, designation, height_mm, width_mm, web_thickness_mm, flange_thickness_mm, root_radius_mm, mass_kg_per_m)

        sections.Add(New SteelSection("UC", "152 x 152 x 23", 152.4, 152.4, 6.1, 6.8, 7.6, 23))
        sections.Add(New SteelSection("UC", "152 x 152 x 30", 157.5, 152.9, 6.6, 9.4, 7.6, 30))
        sections.Add(New SteelSection("UC", "152 x 152 x 37", 161.8, 154.4, 8.1, 11.5, 7.6, 37))
        sections.Add(New SteelSection("UC", "203 x 203 x 46", 203.2, 203.2, 7.3, 11.0, 10.2, 46))
        sections.Add(New SteelSection("UC", "203 x 203 x 52", 206.2, 203.9, 8.0, 12.5, 10.2, 52))
        sections.Add(New SteelSection("UC", "203 x 203 x 60", 209.6, 205.2, 9.3, 14.2, 10.2, 60))
        sections.Add(New SteelSection("UC", "203 x 203 x 71", 215.9, 206.2, 10.3, 17.3, 10.2, 71))
        sections.Add(New SteelSection("UC", "203 x 203 x 86", 222.3, 208.8, 13.0, 20.5, 10.2, 86))
        sections.Add(New SteelSection("UC", "254 x 254 x 73", 254.2, 254.0, 8.6, 14.2, 12.7, 73))
        sections.Add(New SteelSection("UC", "254 x 254 x 89", 260.4, 255.9, 10.5, 17.3, 12.7, 89))
        sections.Add(New SteelSection("UC", "254 x 254 x 107", 266.7, 258.3, 13.0, 20.5, 12.7, 107))
        sections.Add(New SteelSection("UC", "254 x 254 x 132", 276.4, 261.0, 15.6, 25.1, 12.7, 132))
        sections.Add(New SteelSection("UC", "254 x 254 x 167", 289.1, 264.5, 19.2, 31.7, 12.7, 167))
        sections.Add(New SteelSection("UC", "305 x 305 x 97", 307.8, 304.8, 9.9, 15.4, 15.2, 97))
        sections.Add(New SteelSection("UC", "305 x 305 x 118", 314.5, 306.8, 11.9, 18.7, 15.2, 118))
        sections.Add(New SteelSection("UC", "305 x 305 x 137", 320.5, 308.7, 13.8, 21.7, 15.2, 137))
        sections.Add(New SteelSection("UC", "305 x 305 x 158", 327.2, 310.6, 15.7, 25.0, 15.2, 158))

        Return sections
    End Function

    ''' <summary>
    ''' SANS IPE Sections - 6 standard sizes
    ''' </summary>
    Public Shared Function GetSANS_IPESections() As List(Of SteelSection)
        Dim sections As New List(Of SteelSection)

        ' Format: New SteelSection(type, designation, height_mm, width_mm, web_thickness_mm, flange_thickness_mm, root_radius_mm, mass_kg_per_m)
        ' Note: Original iLogic doesn't include mass for IPE, using 0 as placeholder

        sections.Add(New SteelSection("IPE", "IPE100", 100.0, 55.0, 4.1, 5.7, 7.0, 0))
        sections.Add(New SteelSection("IPE", "IPE120", 120.0, 64.0, 4.4, 6.3, 7.0, 0))
        sections.Add(New SteelSection("IPE", "IPE140", 140.0, 73.0, 4.7, 6.9, 7.0, 0))
        sections.Add(New SteelSection("IPE", "IPE160", 160.0, 82.0, 5.0, 7.4, 9.0, 0))
        sections.Add(New SteelSection("IPE", "IPE180", 180.0, 91.0, 5.3, 8.0, 9.0, 0))
        sections.Add(New SteelSection("IPE", "IPE200", 200.0, 100.0, 5.6, 8.5, 12.0, 0))

        Return sections
    End Function

    ''' <summary>
    ''' Get all SANS sections combined
    ''' </summary>
    Public Shared Function GetAllSANSSections() As List(Of SteelSection)
        Dim allSections As New List(Of SteelSection)
        allSections.AddRange(GetSANS_IBeams())
        allSections.AddRange(GetSANS_HBeams())
        allSections.AddRange(GetSANS_IPESections())
        Return allSections
    End Function

    ''' <summary>
    ''' Find a section by designation string
    ''' </summary>
    Public Shared Function FindSectionByDesignation(designation As String) As SteelSection
        For Each section As SteelSection In GetAllSANSSections()
            If section.Designation = designation Then
                Return section
            End If
        Next
        Return Nothing
    End Function

End Class

' ==============================================================================
' UPDATED STEELSECTION CLASS - Add this to your existing SteelSection.vb
' ==============================================================================

Public Class SteelSection

    ' Existing properties
    Public Property SectionType As String ' "UC", "UB", "PFC", "TFC", "L", "IPE"
    Public Property Designation As String ' "203 x 133 x 25"
    Public Property Height As Double ' mm
    Public Property Width As Double ' mm (flange width for I/H, leg for L)
    Public Property WebThickness As Double ' mm
    Public Property FlangeThickness As Double ' mm

    ' NEW PROPERTY - Critical for accurate I-beam profiles!
    Public Property RootRadius As Double ' mm (fillet radius at web-flange junction)

    ' NEW PROPERTY - Mass per meter for weight calculations
    Public Property MassPerMeter As Double ' kg/m

    ' For angles only
    Public Property LegA As Double ' mm
    Public Property LegB As Double ' mm

    ' Hole parameters
    Public Property HoleDiameter As Double ' mm
    Public Property HoleCount As Integer
    Public Property BackmarkX As Double ' mm
    Public Property BackmarkY As Double ' mm

    ''' <summary>
    ''' Constructor for SANS standard sections (with root radius and mass)
    ''' </summary>
    Public Sub New(sectionType As String, designation As String,
                   height As Double, width As Double,
                   webThickness As Double, flangeThickness As Double,
                   rootRadius As Double, massPerMeter As Double)
        Me.SectionType = sectionType
        Me.Designation = designation
        Me.Height = height
        Me.Width = width
        Me.WebThickness = webThickness
        Me.FlangeThickness = flangeThickness
        Me.RootRadius = rootRadius
        Me.MassPerMeter = massPerMeter

        ' Set default hole parameters based on section type
        SetDefaultHoleParameters()
    End Sub

    ''' <summary>
    ''' Get iProperty description string (for Description field)
    ''' </summary>
    Public Function GetIPropertyDescription() As String
        Select Case SectionType
            Case "UB"
                Return Designation & " PFI"
            Case "UC"
                Return Designation & " PFH"
            Case "IPE"
                Return Designation
            Case "PFC", "TFC"
                Return Designation
            Case "L"
                Return "L" & LegA & "x" & LegB & "x" & FlangeThickness
            Case Else
                Return Designation
        End Select
    End Function

    Private Sub SetDefaultHoleParameters()
        ' Default hole parameters based on section type
        ' Can be customized per section if needed
        Select Case SectionType
            Case "UC", "UB"
                HoleDiameter = 22
                HoleCount = 4
                ' Backmarks will be set based on actual flange geometry
                BackmarkX = Width / 2.0 - 30 ' 30mm from edge
                BackmarkY = Height / 2.0 - 50 ' 50mm from edge
            Case "PFC", "TFC"
                HoleDiameter = 18
                HoleCount = 2
                BackmarkX = Height / 4.0
                BackmarkY = Height / 4.0
            Case "L"
                HoleDiameter = 16
                HoleCount = 2
                BackmarkX = LegA / 2.0
                BackmarkY = LegB / 2.0
            Case Else
                HoleCount = 0
        End Select
    End Sub

End Class

' ==============================================================================
' USAGE EXAMPLE
' ==============================================================================
'
' ' Get a specific section
' Dim section As SteelSection = SteelSectionDatabase.FindSectionByDesignation("203 x 133 x 25")
'
' ' Create beam from template
' Dim partDoc As PartDocument = CreateBeamFromTemplate(section)
'
' Function CreateBeamFromTemplate(section As SteelSection) As PartDocument
'     ' Copy template file
'     Dim templatePath As String = "C:\Templates\Template_IBeam.ipt"
'     Dim newPartPath As String = "C:\Output\NewBeam.ipt"
'     System.IO.File.Copy(templatePath, newPartPath, True)
'
'     ' Open template
'     Dim partDoc As PartDocument = m_InventorApp.Documents.Open(newPartPath, False)
'
'     ' Update parameters
'     Dim params As Parameters = partDoc.ComponentDefinition.Parameters.UserParameters
'     params.Item("h_").Value = section.Height / 10.0     ' Convert mm to cm
'     params.Item("b_").Value = section.Width / 10.0
'     params.Item("tw").Value = section.WebThickness / 10.0
'     params.Item("tf").Value = section.FlangeThickness / 10.0
'     params.Item("r1").Value = section.RootRadius / 10.0  ' ← THIS IS THE KEY!
'     params.Item("Length").Value = 1000 / 10.0            ' 1000mm = 100cm
'
'     ' Inventor automatically regenerates the geometry with correct fillets!
'     partDoc.Save2(True)
'
'     Return partDoc
' End Function
' ==============================================================================
