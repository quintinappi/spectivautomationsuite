' ==============================================================================
' DOCUMENT INFO SCANNER - Read iProperties, Mass, and iLogic Rules
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' This module provides diagnostic information about the current document:
' 1. iProperties (Part Number, Description, Stock Number, etc.)
' 2. Mass Properties (Mass, Volume, Surface Area)
' 3. iLogic Rules (Names, Active state, and source code)
'
' Used for debugging and verification before modifying rules.
' ==============================================================================

Imports Inventor
Imports System.IO
Imports System.Text

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Scans documents and extracts iProperties, mass properties, and iLogic rules
    ''' </summary>
    Public Class DocumentInfoScanner

        Private Const ILOGIC_ADDIN_GUID As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

        Private m_InventorApp As Inventor.Application
        Private m_iLogicAddIn As ApplicationAddIn
        Private m_iLogicAuto As Object  ' Late binding for IiLogicAutomation
        Private m_LogPath As String

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            InitializeiLogic()
        End Sub

        ''' <summary>
        ''' Initialize connection to iLogic Add-In
        ''' </summary>
        Private Sub InitializeiLogic()
            Try
                m_iLogicAddIn = m_InventorApp.ApplicationAddIns.ItemById(ILOGIC_ADDIN_GUID)
                If m_iLogicAddIn IsNot Nothing Then
                    m_iLogicAddIn.Activate()
                    m_iLogicAuto = m_iLogicAddIn.Automation
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("DocumentInfoScanner: Could not initialize iLogic - " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Check if iLogic is available
        ''' </summary>
        Public ReadOnly Property IsiLogicAvailable As Boolean
            Get
                Return m_iLogicAuto IsNot Nothing
            End Get
        End Property

#Region "Main Scan Method"

        ''' <summary>
        ''' Scan the active document and display all info in an alert + log
        ''' </summary>
        Public Sub ScanAndDisplayInfo()
            If m_InventorApp.ActiveDocument Is Nothing Then
                MsgBox("No document is open.", MsgBoxStyle.Exclamation, "Document Info Scanner")
                Return
            End If

            Dim doc As Document = m_InventorApp.ActiveDocument
            Dim report As New StringBuilder()
            Dim summary As New StringBuilder()

            ' Initialize log file
            InitializeLog(doc)

            report.AppendLine("=" & New String("="c, 60))
            report.AppendLine("DOCUMENT INFO SCANNER REPORT")
            report.AppendLine("=" & New String("="c, 60))
            report.AppendLine()
            report.AppendLine("Document: " & doc.DisplayName)
            report.AppendLine("Full Path: " & doc.FullFileName)
            report.AppendLine("Type: " & GetDocumentTypeName(doc.DocumentType))
            report.AppendLine("Scanned: " & DateTime.Now.ToString())
            report.AppendLine()

            ' Section 1: iProperties
            report.AppendLine("-" & New String("-"c, 60))
            report.AppendLine("1. iPROPERTIES")
            report.AppendLine("-" & New String("-"c, 60))
            Dim propsInfo As String = GetAllProperties(doc)
            report.AppendLine(propsInfo)
            report.AppendLine()

            ' Section 2: Mass Properties (if part or assembly)
            If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse
               doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                report.AppendLine("-" & New String("-"c, 60))
                report.AppendLine("2. MASS PROPERTIES")
                report.AppendLine("-" & New String("-"c, 60))
                Dim massInfo As String = GetMassProperties(doc)
                report.AppendLine(massInfo)
                report.AppendLine()
            End If

            ' Section 3: iLogic Rules
            report.AppendLine("-" & New String("-"c, 60))
            report.AppendLine("3. iLOGIC RULES")
            report.AppendLine("-" & New String("-"c, 60))
            Dim rulesInfo As String = GetAllRulesInfo(doc)
            report.AppendLine(rulesInfo)
            report.AppendLine()

            report.AppendLine("=" & New String("="c, 60))
            report.AppendLine("END OF REPORT")
            report.AppendLine("=" & New String("="c, 60))

            ' Write to log file
            WriteToLog(report.ToString())

            ' Build summary for message box (shorter version)
            summary.AppendLine("DOCUMENT INFO SCAN COMPLETE")
            summary.AppendLine()
            summary.AppendLine("Document: " & doc.DisplayName)
            summary.AppendLine()
            summary.AppendLine("--- Key iProperties ---")
            summary.AppendLine(GetKeyPropertiesSummary(doc))
            summary.AppendLine()
            summary.AppendLine("--- Mass Properties ---")
            summary.AppendLine(GetMassPropertiesSummary(doc))
            summary.AppendLine()
            summary.AppendLine("--- iLogic Rules ---")
            summary.AppendLine(GetRulesSummary(doc))
            summary.AppendLine()
            summary.AppendLine("Full details logged to:")
            summary.AppendLine(m_LogPath)

            ' Show summary in message box
            MsgBox(summary.ToString(), MsgBoxStyle.Information, "Document Info Scanner")

            ' Also write to debug output
            System.Diagnostics.Debug.WriteLine(report.ToString())
        End Sub

#End Region

#Region "iProperties Methods"

        ''' <summary>
        ''' Get all properties from all property sets
        ''' </summary>
        Private Function GetAllProperties(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            Try
                Dim propSets As PropertySets = doc.PropertySets

                For Each propSet As PropertySet In propSets
                    sb.AppendLine("  [" & propSet.DisplayName & "]")

                    For Each prop As [Property] In propSet
                        Try
                            Dim propValue As String = ""
                            If prop.Value IsNot Nothing Then
                                propValue = prop.Value.ToString()
                            End If

                            ' Only show non-empty properties
                            If Not String.IsNullOrWhiteSpace(propValue) Then
                                sb.AppendLine("    " & prop.DisplayName & ": " & propValue)
                            End If
                        Catch ex As Exception
                            ' Some properties may not be readable
                        End Try
                    Next

                    sb.AppendLine()
                Next

            Catch ex As Exception
                sb.AppendLine("  ERROR reading properties: " & ex.Message)
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Get key properties summary for message box
        ''' </summary>
        Private Function GetKeyPropertiesSummary(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            Try
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")

                ' Read key properties
                Dim partNumber As String = SafeGetProperty(designProps, "Part Number")
                Dim description As String = SafeGetProperty(designProps, "Description")
                Dim stockNumber As String = SafeGetProperty(designProps, "Stock Number")
                Dim material As String = SafeGetProperty(designProps, "Material")

                sb.AppendLine("Part Number: " & If(String.IsNullOrEmpty(partNumber), "(empty)", partNumber))
                sb.AppendLine("Description: " & If(String.IsNullOrEmpty(description), "(empty)", description))
                sb.AppendLine("Stock Number: " & If(String.IsNullOrEmpty(stockNumber), "(empty)", stockNumber))
                sb.AppendLine("Material: " & If(String.IsNullOrEmpty(material), "(empty)", material))

            Catch ex As Exception
                sb.AppendLine("(Could not read properties: " & ex.Message & ")")
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Safely get a property value
        ''' </summary>
        Private Function SafeGetProperty(ByVal propSet As PropertySet, ByVal propName As String) As String
            Try
                Dim prop As [Property] = propSet.Item(propName)
                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    Return prop.Value.ToString()
                End If
            Catch ex As Exception
                ' Property doesn't exist or can't be read
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' Read a specific Design Tracking Property by name
        ''' </summary>
        Public Function ReadDesignTrackingProperty(ByVal doc As Document, ByVal propertyName As String) As String
            Try
                Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Return SafeGetProperty(designProps, propertyName)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Read a specific User Defined Property by name
        ''' </summary>
        Public Function ReadUserDefinedProperty(ByVal doc As Document, ByVal propertyName As String) As String
            Try
                Dim userProps As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
                Return SafeGetProperty(userProps, propertyName)
            Catch ex As Exception
                Return ""
            End Try
        End Function

#End Region

#Region "Mass Properties Methods"

        ''' <summary>
        ''' Get all mass properties
        ''' </summary>
        Private Function GetMassProperties(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            Try
                Dim compDef As ComponentDefinition = Nothing

                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partDoc As PartDocument = CType(doc, PartDocument)
                    compDef = partDoc.ComponentDefinition
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
                    compDef = asmDoc.ComponentDefinition
                End If

                If compDef IsNot Nothing Then
                    Dim massProps As MassProperties = compDef.MassProperties

                    ' Mass in kg (Inventor internal units)
                    Dim massKg As Double = massProps.Mass
                    sb.AppendLine("  Mass: " & massKg.ToString("F3") & " kg")

                    ' Volume in cm³
                    Dim volumeCm3 As Double = massProps.Volume
                    sb.AppendLine("  Volume: " & volumeCm3.ToString("F2") & " cm³")

                    ' Surface area in cm²
                    Dim areaCm2 As Double = massProps.Area
                    sb.AppendLine("  Surface Area: " & areaCm2.ToString("F2") & " cm²")

                    ' Center of gravity
                    Try
                        Dim cog As Point = massProps.CenterOfMass
                        sb.AppendLine("  Center of Mass: (" & cog.X.ToString("F3") & ", " & cog.Y.ToString("F3") & ", " & cog.Z.ToString("F3") & ") cm")
                    Catch
                        ' COG might not be available
                    End Try

                    ' Accuracy
                    sb.AppendLine("  Accuracy: " & massProps.Accuracy.ToString())
                End If

            Catch ex As Exception
                sb.AppendLine("  ERROR reading mass properties: " & ex.Message)
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Get mass properties summary for message box
        ''' </summary>
        Private Function GetMassPropertiesSummary(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            Try
                Dim compDef As ComponentDefinition = Nothing

                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partDoc As PartDocument = CType(doc, PartDocument)
                    compDef = partDoc.ComponentDefinition
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
                    compDef = asmDoc.ComponentDefinition
                End If

                If compDef IsNot Nothing Then
                    Dim massProps As MassProperties = compDef.MassProperties
                    sb.AppendLine("Mass: " & massProps.Mass.ToString("F3") & " kg")
                    sb.AppendLine("Volume: " & massProps.Volume.ToString("F2") & " cm³")
                    sb.AppendLine("Area: " & massProps.Area.ToString("F2") & " cm²")
                Else
                    sb.AppendLine("(Not available for this document type)")
                End If

            Catch ex As Exception
                sb.AppendLine("(Could not read: " & ex.Message & ")")
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Get mass of a document in kg
        ''' </summary>
        Public Function GetMassKg(ByVal doc As Document) As Double
            Try
                Dim compDef As ComponentDefinition = Nothing

                If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim partDoc As PartDocument = CType(doc, PartDocument)
                    compDef = partDoc.ComponentDefinition
                ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
                    compDef = asmDoc.ComponentDefinition
                End If

                If compDef IsNot Nothing Then
                    Return compDef.MassProperties.Mass
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("GetMassKg error: " & ex.Message)
            End Try

            Return 0
        End Function

#End Region

#Region "iLogic Rules Methods"

        ''' <summary>
        ''' Get all iLogic rules info
        ''' </summary>
        Private Function GetAllRulesInfo(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            If Not IsiLogicAvailable Then
                sb.AppendLine("  iLogic Add-In is not available.")
                Return sb.ToString()
            End If

            Try
                Dim rules As Object = m_iLogicAuto.Rules(doc)

                If rules Is Nothing Then
                    sb.AppendLine("  No iLogic rules found in this document.")
                    Return sb.ToString()
                End If

                Dim ruleCount As Integer = 0

                For Each rule As Object In rules
                    ruleCount += 1
                    Dim ruleName As String = rule.Name
                    Dim ruleText As String = rule.Text
                    Dim isActive As Boolean = rule.IsActive

                    sb.AppendLine("  RULE " & ruleCount & ": " & ruleName)
                    sb.AppendLine("    Active: " & isActive.ToString())
                    sb.AppendLine("    Source Code Length: " & ruleText.Length & " characters")
                    sb.AppendLine("    --- SOURCE CODE ---")

                    ' Indent the source code
                    Dim lines As String() = ruleText.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                    For Each line As String In lines
                        sb.AppendLine("    | " & line)
                    Next

                    sb.AppendLine("    --- END SOURCE ---")
                    sb.AppendLine()
                Next

                If ruleCount = 0 Then
                    sb.AppendLine("  No iLogic rules found in this document.")
                Else
                    sb.AppendLine("  Total Rules: " & ruleCount)
                End If

            Catch ex As Exception
                sb.AppendLine("  ERROR reading iLogic rules: " & ex.Message)
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Get rules summary for message box
        ''' </summary>
        Private Function GetRulesSummary(ByVal doc As Document) As String
            Dim sb As New StringBuilder()

            If Not IsiLogicAvailable Then
                sb.AppendLine("iLogic Add-In not available")
                Return sb.ToString()
            End If

            Try
                Dim rules As Object = m_iLogicAuto.Rules(doc)

                If rules Is Nothing Then
                    sb.AppendLine("No rules found")
                    Return sb.ToString()
                End If

                Dim ruleCount As Integer = 0
                Dim activeCount As Integer = 0

                For Each rule As Object In rules
                    ruleCount += 1
                    If CBool(rule.IsActive) Then
                        activeCount += 1
                    End If
                    sb.AppendLine("• " & rule.Name & If(CBool(rule.IsActive), " (active)", " (inactive)"))
                Next

                If ruleCount = 0 Then
                    sb.AppendLine("No rules found")
                Else
                    sb.AppendLine()
                    sb.AppendLine("Total: " & ruleCount & " rule(s), " & activeCount & " active")
                End If

            Catch ex As Exception
                sb.AppendLine("(Could not read rules: " & ex.Message & ")")
            End Try

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Get a list of all rule names in a document
        ''' </summary>
        Public Function GetRuleNames(ByVal doc As Document) As List(Of String)
            Dim result As New List(Of String)

            If Not IsiLogicAvailable Then Return result

            Try
                Dim rules As Object = m_iLogicAuto.Rules(doc)
                If rules IsNot Nothing Then
                    For Each rule As Object In rules
                        result.Add(CStr(rule.Name))
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("GetRuleNames error: " & ex.Message)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Get the source code of a specific rule by name
        ''' </summary>
        Public Function GetRuleText(ByVal doc As Document, ByVal ruleName As String) As String
            If Not IsiLogicAvailable Then Return ""

            Try
                Dim rules As Object = m_iLogicAuto.Rules(doc)
                If rules IsNot Nothing Then
                    For Each rule As Object In rules
                        If CStr(rule.Name) = ruleName Then
                            Return CStr(rule.Text)
                        End If
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("GetRuleText error: " & ex.Message)
            End Try

            Return ""
        End Function

        ''' <summary>
        ''' Set the source code of a specific rule (MODIFIES THE RULE!)
        ''' </summary>
        ''' <returns>True if successful</returns>
        Public Function SetRuleText(ByVal doc As Document, ByVal ruleName As String, ByVal newText As String) As Boolean
            If Not IsiLogicAvailable Then Return False

            Try
                Dim rules As Object = m_iLogicAuto.Rules(doc)
                If rules IsNot Nothing Then
                    For Each rule As Object In rules
                        If CStr(rule.Name) = ruleName Then
                            rule.Text = newText
                            System.Diagnostics.Debug.WriteLine("SetRuleText: Successfully modified rule '" & ruleName & "'")
                            Return True
                        End If
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("SetRuleText error: " & ex.Message)
            End Try

            Return False
        End Function

        ''' <summary>
        ''' Run a specific iLogic rule by name
        ''' </summary>
        Public Function RunRule(ByVal doc As Document, ByVal ruleName As String) As Boolean
            If Not IsiLogicAvailable Then Return False

            Try
                m_iLogicAuto.RunRule(doc, ruleName)
                System.Diagnostics.Debug.WriteLine("RunRule: Successfully ran rule '" & ruleName & "'")
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("RunRule error: " & ex.Message)
                Return False
            End Try
        End Function

#End Region

#Region "Logging Methods"

        ''' <summary>
        ''' Initialize log file
        ''' </summary>
        Private Sub InitializeLog(ByVal doc As Document)
            Try
                Dim docFolder As String = System.IO.Path.GetDirectoryName(doc.FullFileName)
                Dim logsFolder As String = System.IO.Path.Combine(docFolder, "Logs")

                If Not Directory.Exists(logsFolder) Then
                    Directory.CreateDirectory(logsFolder)
                End If

                Dim timestamp As String = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
                Dim docName As String = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
                m_LogPath = System.IO.Path.Combine(logsFolder, "DocInfo_" & docName & "_" & timestamp & ".txt")

            Catch ex As Exception
                ' Fallback to temp folder
                m_LogPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DocInfo_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
            End Try
        End Sub

        ''' <summary>
        ''' Write content to log file
        ''' </summary>
        Private Sub WriteToLog(ByVal content As String)
            Try
                Using writer As New StreamWriter(m_LogPath, False, Encoding.UTF8)
                    writer.Write(content)
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("WriteToLog error: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Utility Methods"

        ''' <summary>
        ''' Get document type as friendly name
        ''' </summary>
        Private Function GetDocumentTypeName(ByVal docType As DocumentTypeEnum) As String
            Select Case docType
                Case DocumentTypeEnum.kPartDocumentObject
                    Return "Part (.ipt)"
                Case DocumentTypeEnum.kAssemblyDocumentObject
                    Return "Assembly (.iam)"
                Case DocumentTypeEnum.kDrawingDocumentObject
                    Return "Drawing (.idw/.dwg)"
                Case DocumentTypeEnum.kPresentationDocumentObject
                    Return "Presentation (.ipn)"
                Case Else
                    Return "Unknown"
            End Select
        End Function

#End Region

    End Class

End Namespace
