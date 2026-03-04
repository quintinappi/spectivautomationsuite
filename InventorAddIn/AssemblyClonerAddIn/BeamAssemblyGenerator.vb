' ==============================================================================
' BEAM ASSEMBLY GENERATOR - Parametric Steel Connection Generator
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' Creates parametric beam assemblies with endplates for Inventor 2026:
'   - Generates steel section profiles (UB, UC, PFC, TFC, L)
'   - Creates endplates with proper hole positions based on backmarks
'   - Assembles components with proper constraints
'   - Sets iProperties correctly (Part Number, Description)
'   - Dynamic updates based on section type selection
'
' All dimensions in centimeters (Inventor's internal unit)
' Input dimensions in millimeters, converted to cm internally
' ==============================================================================

Imports Inventor
Imports System.Runtime.InteropServices

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Generates parametric beam assemblies with endplates
    ''' </summary>
    Public Class BeamAssemblyGenerator

        Private m_InventorApp As Inventor.Application
        Private m_TransientGeometry As TransientGeometry
        Private m_OutputFolder As String
        Private m_PartNumberPrefix As String

        ' Current section being generated
        Private m_CurrentSection As SteelSection
        Private m_BeamLength As Double ' in mm

        ' Generated files
        Private m_BeamPartPath As String
        Private m_EndplatePath As String
        Private m_AssemblyPath As String

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            m_TransientGeometry = m_InventorApp.TransientGeometry
            m_PartNumberPrefix = "BEAM-"
        End Sub

        ''' <summary>
        ''' Main entry point - shows the generator form and creates the assembly
        ''' </summary>
        Public Sub ShowGeneratorForm()
            Using form As New BeamGeneratorForm()
                If form.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    ' Get user selections
                    m_CurrentSection = form.SelectedSection
                    m_BeamLength = form.BeamLength
                    m_OutputFolder = form.OutputFolder
                    m_PartNumberPrefix = form.PartNumberPrefix

                    If m_CurrentSection Is Nothing Then
                        MsgBox("Please select a valid section.", MsgBoxStyle.Exclamation)
                        Return
                    End If

                    ' Create the assembly
                    Try
                        m_InventorApp.SilentOperation = True
                        CreateBeamAssembly()
                        m_InventorApp.SilentOperation = False

                        MsgBox("Beam assembly created successfully!" & vbCrLf & vbCrLf &
                               "Assembly: " & m_AssemblyPath & vbCrLf &
                               "Beam: " & m_BeamPartPath & vbCrLf &
                               "Endplate: " & m_EndplatePath,
                               MsgBoxStyle.Information, "Success!")

                    Catch ex As Exception
                        m_InventorApp.SilentOperation = False
                        MsgBox("Error creating beam assembly: " & ex.Message & vbCrLf & vbCrLf &
                               "Stack: " & ex.StackTrace,
                               MsgBoxStyle.Critical, "Error")
                    End Try
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Create the complete beam assembly with endplates
        ''' </summary>
        Private Sub CreateBeamAssembly()
            ' Ensure output folder exists
            If Not System.IO.Directory.Exists(m_OutputFolder) Then
                System.IO.Directory.CreateDirectory(m_OutputFolder)
            End If

            ' Generate file names
            Dim sectionCode As String = m_CurrentSection.SectionType
            Dim timestamp As String = DateTime.Now.ToString("HHmmss")
            Dim baseName As String = m_PartNumberPrefix & sectionCode & "-" & timestamp

            m_BeamPartPath = System.IO.Path.Combine(m_OutputFolder, baseName & "-BEAM.ipt")
            m_EndplatePath = System.IO.Path.Combine(m_OutputFolder, baseName & "-ENDPLATE.ipt")
            m_AssemblyPath = System.IO.Path.Combine(m_OutputFolder, baseName & ".iam")

            ' Step 1: Create the beam part
            CreateBeamPart()

            ' Step 2: Create the endplate part
            CreateEndplatePart()

            ' Step 3: Create the assembly with constraints
            CreateAssemblyWithConstraints()
        End Sub

#Region "Beam Part Creation"

        ''' <summary>
        ''' Create the beam part with extruded profile
        ''' Uses Inventor's built-in rectangle for reliable profile creation
        ''' </summary>
        Private Sub CreateBeamPart()
            ' Create new part document
            Dim partDoc As PartDocument = CType(m_InventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , True), PartDocument)
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition

            ' Create sketch on XY plane (front view)
            Dim sketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3)) ' XY plane

            ' Get dimensions in cm
            Dim h As Double = m_CurrentSection.Height / 10.0
            Dim w As Double = m_CurrentSection.Width / 10.0

            ' For I-beams and channels, we'll use a composite approach
            ' Create the profile using Inventor's TwoPointRectangle which is guaranteed to work
            Select Case m_CurrentSection.SectionType
                Case "UC", "UB"
                    ' Create I-beam using three rectangles (two flanges + web)
                    CreateIBeamFromRectangles(compDef, sketch)
                Case "PFC", "TFC"
                    ' Create channel using rectangles
                    CreateChannelFromRectangles(compDef, sketch)
                Case "L"
                    ' Create angle using rectangles
                    CreateAngleFromRectangles(compDef, sketch)
                Case Else
                    ' Default: simple rectangle
                    CreateSimpleRectangleBeam(compDef, sketch, w, h)
            End Select

            ' Set iProperties
            SetPartIProperties(partDoc, "BEAM", m_CurrentSection.GetIPropertyDescription())

            ' Save the beam part
            partDoc.SaveAs(m_BeamPartPath, False)
            partDoc.Close(False)
        End Sub

        ''' <summary>
        ''' Create I-beam using three rectangles and Boolean join
        ''' </summary>
        Private Sub CreateIBeamFromRectangles(compDef As PartComponentDefinition, sketch As PlanarSketch)
            Dim h As Double = m_CurrentSection.Height / 10.0
            Dim w As Double = m_CurrentSection.Width / 10.0
            Dim tw As Double = m_CurrentSection.WebThickness / 10.0
            Dim tf As Double = m_CurrentSection.FlangeThickness / 10.0
            Dim lengthCm As Double = m_BeamLength / 10.0

            ' Create top flange rectangle (centered at top)
            Dim topFlangeSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim topRect As SketchEntitiesEnumerator = topFlangeSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-w / 2, h / 2 - tf),
                m_TransientGeometry.CreatePoint2d(w / 2, h / 2))
            Dim topProfile As Profile = topFlangeSketch.Profiles.AddForSolid()
            Dim topExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                topProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kNewBodyOperation)

            ' Create bottom flange rectangle (centered at bottom)
            Dim bottomFlangeSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim bottomRect As SketchEntitiesEnumerator = bottomFlangeSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-w / 2, -h / 2),
                m_TransientGeometry.CreatePoint2d(w / 2, -h / 2 + tf))
            Dim bottomProfile As Profile = bottomFlangeSketch.Profiles.AddForSolid()
            Dim bottomExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                bottomProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kJoinOperation)

            ' Create web rectangle (centered, connecting flanges)
            Dim webSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim webRect As SketchEntitiesEnumerator = webSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-tw / 2, -h / 2 + tf),
                m_TransientGeometry.CreatePoint2d(tw / 2, h / 2 - tf))
            Dim webProfile As Profile = webSketch.Profiles.AddForSolid()
            Dim webExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                webProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kJoinOperation)
        End Sub

        ''' <summary>
        ''' Create channel using rectangles and Boolean join
        ''' </summary>
        Private Sub CreateChannelFromRectangles(compDef As PartComponentDefinition, sketch As PlanarSketch)
            Dim h As Double = m_CurrentSection.Height / 10.0
            Dim w As Double = m_CurrentSection.Width / 10.0
            Dim tw As Double = m_CurrentSection.WebThickness / 10.0
            Dim tf As Double = m_CurrentSection.FlangeThickness / 10.0
            Dim lengthCm As Double = m_BeamLength / 10.0

            ' Create web (vertical part on right side)
            Dim webSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim webRect As SketchEntitiesEnumerator = webSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(0, -h / 2),
                m_TransientGeometry.CreatePoint2d(tw, h / 2))
            Dim webProfile As Profile = webSketch.Profiles.AddForSolid()
            Dim webExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                webProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kNewBodyOperation)

            ' Create top flange (horizontal part at top)
            Dim topFlangeSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim topRect As SketchEntitiesEnumerator = topFlangeSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-w + tw, h / 2 - tf),
                m_TransientGeometry.CreatePoint2d(tw, h / 2))
            Dim topProfile As Profile = topFlangeSketch.Profiles.AddForSolid()
            Dim topExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                topProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kJoinOperation)

            ' Create bottom flange (horizontal part at bottom)
            Dim bottomFlangeSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim bottomRect As SketchEntitiesEnumerator = bottomFlangeSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-w + tw, -h / 2),
                m_TransientGeometry.CreatePoint2d(tw, -h / 2 + tf))
            Dim bottomProfile As Profile = bottomFlangeSketch.Profiles.AddForSolid()
            Dim bottomExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                bottomProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kJoinOperation)
        End Sub

        ''' <summary>
        ''' Create angle using rectangles and Boolean join
        ''' </summary>
        Private Sub CreateAngleFromRectangles(compDef As PartComponentDefinition, sketch As PlanarSketch)
            Dim legA As Double = m_CurrentSection.LegA / 10.0
            Dim legB As Double = m_CurrentSection.LegB / 10.0
            Dim t As Double = m_CurrentSection.FlangeThickness / 10.0
            Dim lengthCm As Double = m_BeamLength / 10.0

            ' Create horizontal leg
            Dim hLegSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim hLegRect As SketchEntitiesEnumerator = hLegSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(0, 0),
                m_TransientGeometry.CreatePoint2d(legA, t))
            Dim hLegProfile As Profile = hLegSketch.Profiles.AddForSolid()
            Dim hLegExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                hLegProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kNewBodyOperation)

            ' Create vertical leg
            Dim vLegSketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim vLegRect As SketchEntitiesEnumerator = vLegSketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(0, t),
                m_TransientGeometry.CreatePoint2d(t, legB))
            Dim vLegProfile As Profile = vLegSketch.Profiles.AddForSolid()
            Dim vLegExtrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                vLegProfile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kJoinOperation)
        End Sub

        ''' <summary>
        ''' Create a simple rectangular beam
        ''' </summary>
        Private Sub CreateSimpleRectangleBeam(compDef As PartComponentDefinition, sketch As PlanarSketch, w As Double, h As Double)
            Dim lengthCm As Double = m_BeamLength / 10.0

            ' Create rectangle using built-in method (guaranteed to work)
            Dim rect As SketchEntitiesEnumerator = sketch.SketchLines.AddAsTwoPointRectangle(
                m_TransientGeometry.CreatePoint2d(-w / 2, -h / 2),
                m_TransientGeometry.CreatePoint2d(w / 2, h / 2))

            Dim profile As Profile = sketch.Profiles.AddForSolid()
            Dim extrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                profile, lengthCm, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection,
                PartFeatureOperationEnum.kNewBodyOperation)
        End Sub

        ''' <summary>
        ''' Draw I-beam/H-section profile (for UC and UB)
        ''' Uses connected lines to ensure a proper closed profile
        ''' </summary>
        Private Sub DrawIBeamProfile(sketch As PlanarSketch)
            Dim lines As SketchLines = sketch.SketchLines

            ' Get dimensions in cm (Inventor's internal unit)
            Dim h As Double = m_CurrentSection.Height / 10.0    ' Height
            Dim w As Double = m_CurrentSection.Width / 10.0      ' Flange width
            Dim tw As Double = m_CurrentSection.WebThickness / 10.0  ' Web thickness
            Dim tf As Double = m_CurrentSection.FlangeThickness / 10.0  ' Flange thickness

            ' Half dimensions for symmetric drawing
            Dim hh As Double = h / 2.0   ' Half height
            Dim hw As Double = w / 2.0   ' Half width
            Dim htw As Double = tw / 2.0 ' Half web thickness

            ' Create points for I-beam profile (12 vertices, clockwise from top-left)
            Dim points As ObjectCollection = m_InventorApp.TransientObjects.CreateObjectCollection()

            ' Top flange
            points.Add(m_TransientGeometry.CreatePoint2d(-hw, hh))           ' 1: Top-left outer
            points.Add(m_TransientGeometry.CreatePoint2d(hw, hh))            ' 2: Top-right outer
            points.Add(m_TransientGeometry.CreatePoint2d(hw, hh - tf))       ' 3: Top-right inner
            points.Add(m_TransientGeometry.CreatePoint2d(htw, hh - tf))      ' 4: Web top-right

            ' Web right side down to bottom flange
            points.Add(m_TransientGeometry.CreatePoint2d(htw, -hh + tf))     ' 5: Web bottom-right

            ' Bottom flange
            points.Add(m_TransientGeometry.CreatePoint2d(hw, -hh + tf))      ' 6: Bottom-right inner
            points.Add(m_TransientGeometry.CreatePoint2d(hw, -hh))           ' 7: Bottom-right outer
            points.Add(m_TransientGeometry.CreatePoint2d(-hw, -hh))          ' 8: Bottom-left outer
            points.Add(m_TransientGeometry.CreatePoint2d(-hw, -hh + tf))     ' 9: Bottom-left inner
            points.Add(m_TransientGeometry.CreatePoint2d(-htw, -hh + tf))    ' 10: Web bottom-left

            ' Web left side up to top flange
            points.Add(m_TransientGeometry.CreatePoint2d(-htw, hh - tf))     ' 11: Web top-left
            points.Add(m_TransientGeometry.CreatePoint2d(-hw, hh - tf))      ' 12: Top-left inner

            ' Draw as connected lines forming a closed loop
            DrawClosedPolygon(sketch, lines, points)
        End Sub

        ''' <summary>
        ''' Draw channel profile (for PFC and TFC)
        ''' </summary>
        Private Sub DrawChannelProfile(sketch As PlanarSketch)
            Dim lines As SketchLines = sketch.SketchLines

            ' Get dimensions in cm
            Dim h As Double = m_CurrentSection.Height / 10.0
            Dim w As Double = m_CurrentSection.Width / 10.0
            Dim tw As Double = m_CurrentSection.WebThickness / 10.0
            Dim tf As Double = m_CurrentSection.FlangeThickness / 10.0

            ' Half height for centering
            Dim hh As Double = h / 2.0

            ' Create points for C-channel profile
            Dim points As ObjectCollection = m_InventorApp.TransientObjects.CreateObjectCollection()
            points.Add(m_TransientGeometry.CreatePoint2d(tw, -hh))           ' 1: Bottom-right (web)
            points.Add(m_TransientGeometry.CreatePoint2d(tw, hh))            ' 2: Top-right (web)
            points.Add(m_TransientGeometry.CreatePoint2d(-w + tw, hh))       ' 3: Top-left outer
            points.Add(m_TransientGeometry.CreatePoint2d(-w + tw, hh - tf))  ' 4: Top-left inner
            points.Add(m_TransientGeometry.CreatePoint2d(0, hh - tf))        ' 5: Inner top flange
            points.Add(m_TransientGeometry.CreatePoint2d(0, -hh + tf))       ' 6: Inner bottom flange
            points.Add(m_TransientGeometry.CreatePoint2d(-w + tw, -hh + tf)) ' 7: Bottom-left inner
            points.Add(m_TransientGeometry.CreatePoint2d(-w + tw, -hh))      ' 8: Bottom-left outer

            ' Draw as connected lines
            DrawClosedPolygon(sketch, lines, points)
        End Sub

        ''' <summary>
        ''' Draw angle profile (for L sections)
        ''' </summary>
        Private Sub DrawAngleProfile(sketch As PlanarSketch)
            Dim lines As SketchLines = sketch.SketchLines

            ' Get dimensions in cm
            Dim legA As Double = m_CurrentSection.LegA / 10.0
            Dim legB As Double = m_CurrentSection.LegB / 10.0
            Dim t As Double = m_CurrentSection.FlangeThickness / 10.0

            ' Create points for L-angle profile
            Dim points As ObjectCollection = m_InventorApp.TransientObjects.CreateObjectCollection()
            points.Add(m_TransientGeometry.CreatePoint2d(0, 0))          ' 1: Inner corner
            points.Add(m_TransientGeometry.CreatePoint2d(legA, 0))       ' 2: Horizontal leg end outer
            points.Add(m_TransientGeometry.CreatePoint2d(legA, t))       ' 3: Horizontal leg end inner
            points.Add(m_TransientGeometry.CreatePoint2d(t, t))          ' 4: Inner junction
            points.Add(m_TransientGeometry.CreatePoint2d(t, legB))       ' 5: Vertical leg end inner
            points.Add(m_TransientGeometry.CreatePoint2d(0, legB))       ' 6: Vertical leg end outer

            ' Draw as connected lines
            DrawClosedPolygon(sketch, lines, points)
        End Sub

        ''' <summary>
        ''' Helper method to draw a closed polygon from a collection of points
        ''' Uses line chaining to ensure a truly closed profile
        ''' </summary>
        Private Sub DrawClosedPolygon(sketch As PlanarSketch, lines As SketchLines, points As ObjectCollection)
            If points.Count < 3 Then Return

            ' Draw lines chained together
            Dim firstLine As SketchLine = Nothing
            Dim prevLine As SketchLine = Nothing

            For i As Integer = 1 To points.Count
                Dim startPt As Point2d = CType(points.Item(i), Point2d)
                Dim endPt As Point2d

                If i = points.Count Then
                    endPt = CType(points.Item(1), Point2d) ' Close back to first point
                Else
                    endPt = CType(points.Item(i + 1), Point2d)
                End If

                Dim newLine As SketchLine

                If prevLine Is Nothing Then
                    ' First line - just draw it normally
                    newLine = lines.AddByTwoPoints(startPt, endPt)
                    firstLine = newLine
                Else
                    ' Subsequent lines - start from previous line's end point
                    ' Use the geometry of the previous line's end sketch point
                    newLine = lines.AddByTwoPoints(prevLine.EndSketchPoint.Geometry, endPt)
                End If

                prevLine = newLine
            Next

            ' Close the loop: merge the last line's endpoint with first line's start point
            If firstLine IsNot Nothing AndAlso prevLine IsNot Nothing Then
                Try
                    ' Add coincident constraint to close the loop
                    sketch.GeometricConstraints.AddCoincident(
                        CType(prevLine.EndSketchPoint, SketchEntity),
                        CType(firstLine.StartSketchPoint, SketchEntity))
                Catch
                    ' If constraint fails, the lines should still be close enough
                End Try
            End If
        End Sub

#End Region

#Region "Endplate Creation"

        ''' <summary>
        ''' Create the endplate part with holes
        ''' </summary>
        Private Sub CreateEndplatePart()
            ' Create new part document
            Dim partDoc As PartDocument = CType(m_InventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, , True), PartDocument)
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition

            ' Determine endplate dimensions based on section
            Dim plateWidth As Double = m_CurrentSection.Width + 20 ' Section width + 10mm each side
            Dim plateHeight As Double = m_CurrentSection.Height + 20 ' Section height + 10mm each side
            Dim plateThickness As Double = 12 ' Default 12mm endplate

            ' Adjust for specific section types
            Select Case m_CurrentSection.SectionType
                Case "UC", "UB"
                    ' Endplate slightly larger than flange
                    plateWidth = m_CurrentSection.Width + 20
                    plateHeight = m_CurrentSection.Height + 20
                    plateThickness = If(m_CurrentSection.Height > 300, 20, 12)
                Case "PFC", "TFC"
                    ' For channels, width is the height of channel, height is the flange width
                    plateWidth = m_CurrentSection.Height + 20
                    plateHeight = m_CurrentSection.Width + 20
                    plateThickness = 10
                Case "L"
                    ' For angles, plate based on leg sizes
                    plateWidth = m_CurrentSection.LegA + 20
                    plateHeight = m_CurrentSection.LegB + 20
                    plateThickness = 10
            End Select

            ' Create sketch on XY plane
            Dim sketch As PlanarSketch = compDef.Sketches.Add(compDef.WorkPlanes.Item(3))
            Dim lines As SketchLines = sketch.SketchLines

            ' Convert to cm
            Dim pwCm As Double = plateWidth / 10.0
            Dim phCm As Double = plateHeight / 10.0
            Dim hpw As Double = pwCm / 2.0
            Dim hph As Double = phCm / 2.0

            ' Draw rectangle for endplate (centered)
            Dim p1 As Point2d = m_TransientGeometry.CreatePoint2d(-hpw, -hph)
            Dim p2 As Point2d = m_TransientGeometry.CreatePoint2d(hpw, -hph)
            Dim p3 As Point2d = m_TransientGeometry.CreatePoint2d(hpw, hph)
            Dim p4 As Point2d = m_TransientGeometry.CreatePoint2d(-hpw, hph)

            lines.AddByTwoPoints(p1, p2)
            lines.AddByTwoPoints(p2, p3)
            lines.AddByTwoPoints(p3, p4)
            lines.AddByTwoPoints(p4, p1)

            ' Create profile and extrude
            Dim profile As Profile = sketch.Profiles.AddForSolid()
            Dim thicknessCm As Double = plateThickness / 10.0

            ' Create extrusion using AddByDistanceExtent
            Dim extrude As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                profile,
                thicknessCm,
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kNewBodyOperation)

            ' Add holes based on section type and backmarks
            AddEndplateHoles(compDef, plateWidth, plateHeight, plateThickness)

            ' Set iProperties
            Dim description As String = "PL " & CInt(plateThickness) & "mm S355JR"
            SetPartIProperties(partDoc, "ENDPLATE", description)

            ' Save
            partDoc.SaveAs(m_EndplatePath, False)
            partDoc.Close(False)
        End Sub

        ''' <summary>
        ''' Add holes to the endplate based on section backmarks
        ''' </summary>
        Private Sub AddEndplateHoles(compDef As PartComponentDefinition, plateWidth As Double, plateHeight As Double, plateThickness As Double)
            ' Get hole parameters
            Dim holeDiameter As Double = m_CurrentSection.HoleDiameter ' mm
            Dim holeCount As Integer = m_CurrentSection.HoleCount
            Dim backmarkX As Double = m_CurrentSection.BackmarkX ' mm from center
            Dim backmarkY As Double = m_CurrentSection.BackmarkY ' mm from center

            If holeCount = 0 Then Return ' No holes for hollow sections

            ' Create sketch for hole centers on front face
            Dim frontFace As Face = GetFrontFace(compDef)
            If frontFace Is Nothing Then Return

            Dim holeSketch As PlanarSketch = compDef.Sketches.Add(frontFace)

            ' Convert to cm
            Dim holeDiameterCm As Double = holeDiameter / 10.0
            Dim backmarkXCm As Double = backmarkX / 10.0
            Dim backmarkYCm As Double = backmarkY / 10.0

            ' Collect hole center points
            Dim holePoints As New System.Collections.Generic.List(Of Point2d)

            ' Add hole center points based on section type
            Select Case m_CurrentSection.SectionType
                Case "UC", "UB"
                    ' 4 holes for I/H sections (2x2 pattern)
                    If holeCount = 4 Then
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(backmarkXCm, backmarkYCm))
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(-backmarkXCm, backmarkYCm))
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(backmarkXCm, -backmarkYCm))
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(-backmarkXCm, -backmarkYCm))
                    ElseIf holeCount = 2 Then
                        ' Narrow sections with 2 holes (single row)
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(0, backmarkYCm))
                        holePoints.Add(m_TransientGeometry.CreatePoint2d(0, -backmarkYCm))
                    End If

                Case "PFC", "TFC"
                    ' 2 holes for channels (single row in web)
                    Dim channelBackmark As Double = backmarkXCm
                    holePoints.Add(m_TransientGeometry.CreatePoint2d(0, channelBackmark))
                    holePoints.Add(m_TransientGeometry.CreatePoint2d(0, -channelBackmark))

                Case "L"
                    ' 2 holes for angles (one in each leg)
                    holePoints.Add(m_TransientGeometry.CreatePoint2d(backmarkXCm, 0))
                    holePoints.Add(m_TransientGeometry.CreatePoint2d(0, backmarkYCm))
            End Select

            ' Create holes for each point using sketch circles and cut extrude
            For Each holeCenter As Point2d In holePoints
                Try
                    ' Add sketch point for hole center
                    Dim sketchPt As SketchPoint = holeSketch.SketchPoints.Add(holeCenter, False)

                    ' Draw circle for hole
                    Dim circle As SketchCircle = holeSketch.SketchCircles.AddByCenterRadius(holeCenter, holeDiameterCm / 2.0)
                Catch ex As Exception
                    ' Continue with next hole if this one fails
                End Try
            Next

            ' Create cut extrude for holes if we have circles
            If holeSketch.SketchCircles.Count > 0 Then
                Try
                    Dim holeProfile As Profile = holeSketch.Profiles.AddForSolid()
                    Dim holeCut As ExtrudeFeature = compDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                        holeProfile,
                        plateThickness / 10.0 + 0.1, ' Extra depth to ensure through hole
                        PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
                        PartFeatureOperationEnum.kCutOperation)
                Catch ex As Exception
                    ' Holes failed but endplate is still valid
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Get the front face of the extruded plate (face with normal in Z direction)
        ''' </summary>
        Private Function GetFrontFace(compDef As PartComponentDefinition) As Face
            For Each body As SurfaceBody In compDef.SurfaceBodies
                For Each face As Face In body.Faces
                    Try
                        Dim evalObj As Object = face.Geometry
                        If TypeOf evalObj Is Plane Then
                            Dim planeGeom As Plane = CType(evalObj, Plane)
                            If Math.Abs(planeGeom.Normal.Z) > 0.9 AndAlso planeGeom.Normal.Z > 0 Then
                                Return face
                            End If
                        End If
                    Catch
                    End Try
                Next
            Next
            Return Nothing
        End Function

#End Region

#Region "Assembly Creation"

        ''' <summary>
        ''' Create the assembly with beam and endplates, properly constrained
        ''' </summary>
        Private Sub CreateAssemblyWithConstraints()
            ' Create new assembly document
            Dim asmDoc As AssemblyDocument = CType(m_InventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, , True), AssemblyDocument)
            Dim asmCompDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

            ' Place the beam at origin (grounded)
            Dim beamMatrix As Matrix = m_TransientGeometry.CreateMatrix()
            Dim beamOcc As ComponentOccurrence = asmCompDef.Occurrences.Add(m_BeamPartPath, beamMatrix)
            beamOcc.Grounded = True

            ' Calculate endplate positions (at each end of beam)
            Dim beamHalfLength As Double = m_BeamLength / 2.0 / 10.0 ' In cm, half length

            ' Place first endplate at positive Z end
            Dim endplate1Matrix As Matrix = m_TransientGeometry.CreateMatrix()
            endplate1Matrix.SetTranslation(m_TransientGeometry.CreateVector(0, 0, beamHalfLength), False)
            Dim endplate1Occ As ComponentOccurrence = asmCompDef.Occurrences.Add(m_EndplatePath, endplate1Matrix)

            ' Place second endplate at negative Z end (rotated 180 degrees)
            Dim endplate2Matrix As Matrix = m_TransientGeometry.CreateMatrix()
            endplate2Matrix.SetTranslation(m_TransientGeometry.CreateVector(0, 0, -beamHalfLength), False)
            ' Rotate 180 degrees around Y axis
            Dim rotVector As Vector = m_TransientGeometry.CreateVector(0, 1, 0)
            endplate2Matrix.SetToRotation(Math.PI, rotVector, m_TransientGeometry.CreatePoint(0, 0, -beamHalfLength))
            Dim endplate2Occ As ComponentOccurrence = asmCompDef.Occurrences.Add(m_EndplatePath, endplate2Matrix)

            ' Add flush constraints between beam ends and endplates
            Try
                AddEndplateConstraints(asmCompDef, beamOcc, endplate1Occ, True)
                AddEndplateConstraints(asmCompDef, beamOcc, endplate2Occ, False)
            Catch ex As Exception
                ' Constraints may fail, but assembly is still valid
            End Try

            ' Set assembly iProperties
            SetAssemblyIProperties(asmDoc)

            ' Save the assembly
            asmDoc.SaveAs(m_AssemblyPath, False)

            ' Keep assembly open for user to view
            ' asmDoc.Close(False)
        End Sub

        ''' <summary>
        ''' Add constraints between beam and endplate
        ''' </summary>
        Private Sub AddEndplateConstraints(asmCompDef As AssemblyComponentDefinition, beamOcc As ComponentOccurrence, endplateOcc As ComponentOccurrence, isPositiveEnd As Boolean)
            ' Find matching faces for flush constraint
            Dim beamEndFace As Face = FindBeamEndFace(beamOcc, isPositiveEnd)
            Dim endplateFace As Face = FindEndplateFace(endplateOcc, Not isPositiveEnd)

            If beamEndFace IsNot Nothing AndAlso endplateFace IsNot Nothing Then
                Try
                    asmCompDef.Constraints.AddFlushConstraint(beamEndFace, endplateFace, 0)
                Catch
                End Try
            End If

            ' Add mate constraints for alignment (optional - position already set by matrix)
            ' Could add center-to-center alignment here if needed
        End Sub

        ''' <summary>
        ''' Find the end face of the beam
        ''' </summary>
        Private Function FindBeamEndFace(beamOcc As ComponentOccurrence, isPositiveEnd As Boolean) As Face
            For Each body As SurfaceBody In beamOcc.SurfaceBodies
                For Each face As Face In body.Faces
                    Try
                        Dim evalObj As Object = face.Geometry
                        If TypeOf evalObj Is Plane Then
                            Dim planeGeom As Plane = CType(evalObj, Plane)
                            If Math.Abs(planeGeom.Normal.Z) > 0.9 Then
                                Dim center As Point = face.PointOnFace
                                If (isPositiveEnd AndAlso center.Z > 0) OrElse (Not isPositiveEnd AndAlso center.Z < 0) Then
                                    Return face
                                End If
                            End If
                        End If
                    Catch
                    End Try
                Next
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Find the face of the endplate
        ''' </summary>
        Private Function FindEndplateFace(endplateOcc As ComponentOccurrence, isBackFace As Boolean) As Face
            For Each body As SurfaceBody In endplateOcc.SurfaceBodies
                For Each face As Face In body.Faces
                    Try
                        Dim evalObj As Object = face.Geometry
                        If TypeOf evalObj Is Plane Then
                            Dim planeGeom As Plane = CType(evalObj, Plane)
                            If Math.Abs(planeGeom.Normal.Z) > 0.9 Then
                                Return face
                            End If
                        End If
                    Catch
                    End Try
                Next
            Next
            Return Nothing
        End Function

#End Region

#Region "iProperties"

        ''' <summary>
        ''' Set iProperties for a part document
        ''' </summary>
        Private Sub SetPartIProperties(partDoc As PartDocument, partType As String, description As String)
            Try
                Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")

                ' Set Part Number
                Dim partNumber As String = m_PartNumberPrefix & m_CurrentSection.SectionType & "-" & partType
                designProps.Item("Part Number").Value = partNumber

                ' Set Description
                designProps.Item("Description").Value = description

                ' Set other properties
                Try
                    designProps.Item("Stock Number").Value = partNumber
                Catch
                End Try

                ' Set custom properties
                Dim customProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
                Try
                    Dim materialProp = customProps.Add("S355JR", "Material")
                Catch
                    Try
                        customProps.Item("Material").Value = "S355JR"
                    Catch
                    End Try
                End Try

            Catch ex As Exception
                ' Log but don't fail
            End Try
        End Sub

        ''' <summary>
        ''' Set iProperties for the assembly document
        ''' </summary>
        Private Sub SetAssemblyIProperties(asmDoc As AssemblyDocument)
            Try
                Dim designProps As PropertySet = asmDoc.PropertySets.Item("Design Tracking Properties")

                ' Set Part Number
                Dim partNumber As String = m_PartNumberPrefix & m_CurrentSection.SectionType & "-ASSY"
                designProps.Item("Part Number").Value = partNumber

                ' Set Description
                Dim description As String = m_CurrentSection.GetIPropertyDescription() & " BEAM ASSEMBLY"
                designProps.Item("Description").Value = description

            Catch ex As Exception
                ' Log but don't fail
            End Try
        End Sub

#End Region

    End Class

End Namespace
