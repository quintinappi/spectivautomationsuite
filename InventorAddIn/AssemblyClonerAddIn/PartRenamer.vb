' ==============================================================================
' PART RENAMER - Dynamic Heritage-Based Solution
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' VB.NET port of Assembly_Renamer.vbs (Option 1 from Main_Launcher.bat)
' Uses EXACT same logic as the VBScript version:
'   1. Detects currently open assembly in Inventor
'   2. Flattens hierarchy and groups similar components
'   3. Asks user for naming scheme per component group
'   4. Uses proven heritage method for model + IDW updates
' ==============================================================================

Imports Inventor
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Part Renamer - Dynamic heritage-based renaming solution
    ''' Port of Assembly_Renamer.vbs with EXACT same logic
    ''' </summary>
    Public Class PartRenamer

        Private m_InventorApp As Inventor.Application

        ' Tracking dictionaries (same as VBScript)
        Private m_ComponentGroups As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)  ' groupCode -> (fullPath -> "fullPath|description|fileName")
        Private m_NamingSchemes As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' groupCode -> namingScheme
        Private m_FileNameMapping As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' originalPath -> newFileName
        Private m_ComprehensiveMapping As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' originalPath -> "newPath|originalFile|newFile|group|description"
        Private m_AssemblyPathMapping As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' originalAssemblyPath -> newAssemblyPath
        Private m_PlantSection As String = ""

        ' Logging
        Private m_LogFile As StreamWriter
        Private m_LogPath As String

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
        End Sub

        ''' <summary>
        ''' Main entry point - rename parts in active assembly
        ''' </summary>
        Public Sub RenameAssemblyParts()
            Try
                ' Initialize logging
                InitializeLog()
                LogMessage("=== DYNAMIC HERITAGE-BASED SOLUTION ===")
                LogMessage("Auto-detecting open model and creating dynamic renaming workflow")

                ' Reset tracking
                m_ComponentGroups.Clear()
                m_NamingSchemes.Clear()
                m_FileNameMapping.Clear()
                m_ComprehensiveMapping.Clear()
                m_AssemblyPathMapping.Clear()

                ' Confirm with user
                Dim confirmResult As MsgBoxResult = MsgBox(
                    "DYNAMIC INVENTOR RENAMING TOOL" & vbCrLf & vbCrLf &
                    "This will:" & vbCrLf &
                    "1. Detect your currently open assembly" & vbCrLf &
                    "2. Group similar components automatically" & vbCrLf &
                    "3. Let you define naming schemes per group" & vbCrLf &
                    "4. Update models AND drawings automatically" & vbCrLf & vbCrLf &
                    "Make sure your target assembly is open in Inventor!" & vbCrLf & vbCrLf &
                    "Continue?",
                    MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                    "Dynamic Heritage Solution")

                If confirmResult = MsgBoxResult.No Then
                    LogMessage("User cancelled workflow")
                    CloseLog()
                    Return
                End If

                ' Step 1: Detect open assembly
                LogMessage("STEP 1: Detecting open assembly and analyzing components")
                Dim activeDoc As AssemblyDocument = DetectOpenAssembly()
                If activeDoc Is Nothing Then
                    CloseLog()
                    Return
                End If

                ' Step 2: Flatten hierarchy and group components
                LogMessage("STEP 2: Flattening hierarchy and grouping components")
                FlattenAndGroupComponents(activeDoc)

                ' Step 3: Get plant section naming convention
                LogMessage("STEP 3: Getting plant section naming convention")
                If Not GetPlantSectionNaming() Then
                    CloseLog()
                    Return
                End If

                ' Step 4: Show groups summary
                LogMessage("STEP 4: Showing component groups summary")
                If Not ShowGroupsSummary() Then
                    CloseLog()
                    Return
                End If

                ' Step 5: Get user input for naming schemes
                LogMessage("STEP 5: Getting user naming schemes for component groups")
                GetUserNamingSchemes()

                ' Step 6: Create heritage-based copies
                LogMessage("STEP 6: Creating heritage-based copies with user naming")
                CreateDynamicHeritageBasedCopies(activeDoc)

                ' Step 7: Update assembly references
                LogMessage("STEP 7: Updating assembly references")
                UpdateDynamicAssemblyReferences(activeDoc)

                ' Step 7.5: Rename main assembly with project prefix
                LogMessage("STEP 7.5: Renaming main assembly with project prefix")
                RenameMainAssemblyWithPrefix(activeDoc)

                ' Step 8: Update IDW files
                LogMessage("STEP 8: Auto-detecting and updating IDW files")
                UpdateAllIDWsInDirectory(activeDoc)

                ' Step 8.5: Save mapping file
                LogMessage("STEP 8.5: Saving mapping file for external IDW updater")
                SaveMappingFile(activeDoc)

                ' Step 9: Complete
                LogMessage("STEP 9: Keeping original files for safety (skipping cleanup)")
                LogMessage("=== DYNAMIC HERITAGE-BASED SOLUTION COMPLETED ===")

                CloseLog()

                MsgBox(
                    "DYNAMIC SOLUTION COMPLETED!" & vbCrLf & vbCrLf &
                    "Components analyzed and grouped" & vbCrLf &
                    "Heritage-based copies created" & vbCrLf &
                    "Assembly references updated" & vbCrLf &
                    "IDW drawings updated automatically" & vbCrLf &
                    "Mapping file saved for STEP 2" & vbCrLf &
                    "Fully automated workflow!" & vbCrLf & vbCrLf &
                    "Log: " & m_LogPath,
                    MsgBoxStyle.Information,
                    "Success!")

            Catch ex As Exception
                Dim errorMsg As String = "ERROR: " & ex.Message & vbCrLf & "Stack: " & ex.StackTrace
                LogMessage(errorMsg)
                CloseLog()
                MsgBox("PART RENAMER FAILED!" & vbCrLf & vbCrLf &
                       "Error: " & ex.Message,
                       MsgBoxStyle.Critical, "Rename Failed")
            End Try
        End Sub

        ''' <summary>
        ''' Detect the open assembly (same as VBScript DetectOpenAssembly)
        ''' </summary>
        Private Function DetectOpenAssembly() As AssemblyDocument
            Try
                If m_InventorApp.ActiveDocument Is Nothing Then
                    LogMessage("No active document found")
                    MsgBox("ERROR: No assembly is currently open in Inventor!" & vbCrLf &
                           "Please open your target assembly first.", MsgBoxStyle.Critical)
                    Return Nothing
                End If

                Dim activeDoc As Document = m_InventorApp.ActiveDocument

                ' Check by file extension
                If Not activeDoc.FullFileName.ToLower().EndsWith(".iam") Then
                    LogMessage("File extension is not .iam: " & activeDoc.FullFileName)
                    MsgBox("FILE TYPE ISSUE" & vbCrLf & vbCrLf &
                           "Current file: " & activeDoc.DisplayName & vbCrLf &
                           "Need: Assembly file (.iam extension)", MsgBoxStyle.Exclamation)
                    Return Nothing
                End If

                LogMessage("DETECTED: Active assembly - " & activeDoc.DisplayName)
                LogMessage("DETECTED: Full path - " & activeDoc.FullFileName)

                Dim asmDoc As AssemblyDocument = CType(activeDoc, AssemblyDocument)
                Dim occCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
                Dim folderPath As String = System.IO.Path.GetDirectoryName(activeDoc.FullFileName)

                ' Show validation prompt
                Dim confirmResult As MsgBoxResult = MsgBox(
                    "ASSEMBLY DETECTED" & vbCrLf & vbCrLf &
                    "Assembly: " & activeDoc.DisplayName & vbCrLf &
                    "Parts Count: " & occCount & " occurrences" & vbCrLf &
                    "Location: " & folderPath & vbCrLf & vbCrLf &
                    "Is this the correct assembly to process?" & vbCrLf & vbCrLf &
                    "WARNING: This will create heritage files for all parts!",
                    MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                    "Confirm Assembly")

                If confirmResult = MsgBoxResult.No Then
                    LogMessage("USER CANCELLED: Assembly validation failed")
                    Return Nothing
                End If

                LogMessage("USER CONFIRMED: Proceeding with assembly processing")
                Return asmDoc

            Catch ex As Exception
                LogMessage("ERROR detecting assembly: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Flatten and group components (same as VBScript FlattenAndGroupComponents)
        ''' </summary>
        Private Sub FlattenAndGroupComponents(ByVal asmDoc As AssemblyDocument)
            LogMessage("ANALYZE: Recursively flattening ENTIRE model hierarchy and reading iProperty descriptions")

            Dim uniqueParts As New Dictionary(Of String, Boolean)
            ProcessAssemblyRecursively(asmDoc, uniqueParts, "ROOT", True)

            ' Fallback: if everything was filtered by legacy heritage-name safety check,
            ' run a second pass that allows those files so user still gets groups.
            If m_ComponentGroups.Count = 0 Then
                LogMessage("ANALYZE: No groups found on first pass. Retrying including heritage-named files...")
                uniqueParts.Clear()
                ProcessAssemblyRecursively(asmDoc, uniqueParts, "ROOT", False)
            End If

            LogMessage("ANALYZE: Recursive processing completed - Total unique parts processed: " & uniqueParts.Count)
            LogMessage("ANALYZE: Created " & m_ComponentGroups.Count & " component groups")

            ' Debug log all groups
            For Each kvp In m_ComponentGroups
                Dim groupName As String = kvp.Key
                Dim groupDict As Dictionary(Of String, String) = kvp.Value
                LogMessage("DEBUG GROUP: '" & groupName & "' contains " & groupDict.Count & " components")
            Next
        End Sub

        ''' <summary>
        ''' Recursively process assembly (same as VBScript ProcessAssemblyRecursively)
        ''' </summary>
        Private Sub ProcessAssemblyRecursively(ByVal asmDoc As AssemblyDocument, ByVal uniqueParts As Dictionary(Of String, Boolean), ByVal asmLevel As String, ByVal skipHeritageNamed As Boolean)
            LogMessage("ANALYZE: Processing assembly - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")")

            Dim occurrences As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences
            LogMessage("ANALYZE: Found " & occurrences.Count & " occurrences in " & asmDoc.DisplayName)

            For i As Integer = 1 To occurrences.Count
                Try
                    Dim occ As ComponentOccurrence = occurrences.Item(i)

                    If occ.Suppressed Then
                        LogMessage("ANALYZE: SKIPPING (suppressed occurrence in " & asmDoc.DisplayName & ")")
                        Continue For
                    End If

                    Dim doc As Document = occ.Definition.Document
                    Dim fileName As String = System.IO.Path.GetFileName(doc.FullFileName)
                    Dim fullPath As String = doc.FullFileName

                    If fileName.ToLower().EndsWith(".ipt") Then
                        ' CRITICAL SAFETY CHECK: Skip if already heritage-renamed
                        If skipHeritageNamed AndAlso (fileName.Contains("-000-") OrElse fileName.Contains("-751-") OrElse fileName.Contains("-752-")) Then
                            LogMessage("ANALYZE: SKIPPING (already heritage-renamed) - " & fileName)
                            Continue For
                        End If

                        If Not uniqueParts.ContainsKey(fullPath) Then
                            uniqueParts.Add(fullPath, True)

                            ' Read Description from Design Tracking Properties
                            Dim description As String = GetDescriptionFromIProperty(doc)

                            If String.IsNullOrEmpty(description) Then
                                LogMessage("ANALYZE: WARNING - No description found for " & fileName)
                                description = "NO DESCRIPTION"
                            End If

                            ' Group by description (or fallback marker)
                            Dim groupCode As String = ClassifyByDescription(description)

                            If groupCode = "SKIP" Then
                                LogMessage("ANALYZE: SKIPPING " & fileName & " (hardware/bolts) - Description: " & description)
                            Else
                                LogMessage("ANALYZE: PART - " & fileName & " -> Description: " & description & " -> Group: " & groupCode)

                                ' Add to component groups
                                If Not m_ComponentGroups.ContainsKey(groupCode) Then
                                    m_ComponentGroups.Add(groupCode, New Dictionary(Of String, String))
                                    LogMessage("ANALYZE: Created new group - " & groupCode)
                                End If

                                Dim groupDict As Dictionary(Of String, String) = m_ComponentGroups(groupCode)
                                If Not groupDict.ContainsKey(fullPath) Then
                                    groupDict.Add(fullPath, fullPath & "|" & description & "|" & fileName)
                                End If
                            End If
                        Else
                            LogMessage("ANALYZE: DUPLICATE PART SKIPPED - " & fileName)
                        End If

                    ElseIf fileName.ToLower().EndsWith(".iam") Then
                        ' Sub-assembly - recurse (skip bolted connections)
                        If fileName.ToLower().Contains("bolted connection") Then
                            LogMessage("ANALYZE: SKIPPING " & fileName & " (bolted connection assembly)")
                        Else
                            LogMessage("ANALYZE: RECURSING into sub-assembly - " & fileName)
                            Dim subAsm As AssemblyDocument = CType(doc, AssemblyDocument)
                            ProcessAssemblyRecursively(subAsm, uniqueParts, asmLevel & ">" & fileName, skipHeritageNamed)
                        End If
                    End If

                Catch ex As Exception
                    LogMessage("ANALYZE: Error processing occurrence: " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Get description from iProperty (same as VBScript GetDescriptionFromIProperty)
        ''' </summary>
        Private Function GetDescriptionFromIProperty(ByVal doc As Document) As String
            Try
                Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
                Dim descProp As [Property] = propSet.Item("Description")
                Return descProp.Value.ToString().Trim()
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Classify by description (EXACT same logic as VBScript ClassifyByDescription)
        ''' </summary>
        Private Function ClassifyByDescription(ByVal description As String) As String
            Dim desc As String = description.ToUpper().Trim()

            ' Skip hardware and bolts first
            If desc.Contains("BOLT") OrElse desc.Contains("SCREW") OrElse desc.Contains("WASHER") OrElse desc.Contains("NUT") Then
                Return "SKIP"
            End If

            ' Client's grouping logic - exact requirements
            If desc.StartsWith("UB") Then
                Return "B"  ' I and H sections - UB beams
            ElseIf desc.StartsWith("UC") Then
                Return "B"  ' I and H sections - UC columns
            ElseIf desc.StartsWith("PL") Then
                ' Check if it's platework (PL + S355JR) or liners (PL + NOT S355JR)
                If desc.Contains("S355JR") Then
                    Return "PL"  ' Platework
                Else
                    Return "LPL" ' Liners
                End If
            ElseIf desc.StartsWith("L") AndAlso (desc.Contains("X") OrElse desc.Contains(" X ")) Then
                Return "A"   ' Angles
            ElseIf desc.StartsWith("PFC") Then
                Return "CH"  ' Parallel flange channels
            ElseIf desc.StartsWith("TFC") Then
                Return "CH"  ' Taper flange channels
            ElseIf desc.StartsWith("CHS") Then
                Return "P"   ' Circular hollow sections
            ElseIf desc.StartsWith("SHS") Then
                Return "SQ"  ' Square/rectangular hollow sections
            ElseIf desc.StartsWith("FL") AndAlso Not desc.Contains("FLOOR") Then
                Return "FL"  ' Flatbar (but not floor grating)
            ElseIf desc.StartsWith("IPE") Then
                Return "IPE"  ' European I-beams
            Else
                Return "OTHER"
            End If
        End Function

        ''' <summary>
        ''' Get plant section naming from user (same as VBScript GetPlantSectionNaming)
        ''' </summary>
        Private Function GetPlantSectionNaming() As Boolean
            Dim plantInput As String = InputBox(
                "STEP 3: DEFINE PREFIX" & vbCrLf & vbCrLf &
                "Enter the project prefix (as per drawing register):" & vbCrLf & vbCrLf &
                "Examples:" & vbCrLf &
                "  PLANT1-000-    (for Plant 1)" & vbCrLf &
                "  AREA2-000-     (for Area 2)" & vbCrLf &
                "  SEC-A-000-     (for Section A)" & vbCrLf & vbCrLf &
                "This will create part numbers like:" & vbCrLf &
                "  PLANT1-000-B1, PLANT1-000-PL1, PLANT1-000-CH1, etc." & vbCrLf & vbCrLf &
                "Leave blank for default TEST-000-",
                "Define Project Prefix",
                "PLANT1-000-")

            If String.IsNullOrEmpty(plantInput) Then
                m_PlantSection = "TEST-000-"
                LogMessage("PLANT: Using default naming convention: TEST-000-")
            Else
                plantInput = plantInput.Trim()

                ' Validate prefix
                Dim isValid As Boolean = True
                For Each c As Char In plantInput
                    If Not Char.IsLetterOrDigit(c) AndAlso c <> "-"c AndAlso c <> "_"c Then
                        isValid = False
                        Exit For
                    End If
                Next

                If Not isValid Then
                    MsgBox("Invalid prefix!" & vbCrLf & vbCrLf &
                           "Prefix can only contain:" & vbCrLf &
                           "• Letters (A-Z)" & vbCrLf &
                           "• Numbers (0-9)" & vbCrLf &
                           "• Dash (-)" & vbCrLf &
                           "• Underscore (_)",
                           MsgBoxStyle.Critical)
                    Return False
                End If

                ' Ensure ends with dash
                If Not plantInput.EndsWith("-") Then
                    plantInput = plantInput & "-"
                End If

                m_PlantSection = plantInput.ToUpper()
                LogMessage("PLANT: Using custom naming convention: " & m_PlantSection)
            End If

            MsgBox("PROJECT PREFIX SET" & vbCrLf & vbCrLf &
                   "Your project prefix: " & m_PlantSection & vbCrLf & vbCrLf &
                   "Example part numbers will be:" & vbCrLf &
                   "  " & m_PlantSection & "B1 (I/H sections)" & vbCrLf &
                   "  " & m_PlantSection & "PL1 (Platework)" & vbCrLf &
                   "  " & m_PlantSection & "CH1 (Channels)" & vbCrLf &
                   "  " & m_PlantSection & "A1 (Angles)",
                   MsgBoxStyle.Information,
                   "Prefix Set")

            Return True
        End Function

        ''' <summary>
        ''' Show groups summary (same as VBScript ShowGroupsSummary)
        ''' </summary>
        Private Function ShowGroupsSummary() As Boolean
            If m_ComponentGroups.Count = 0 Then
                MsgBox("No component groups found!" & vbCrLf & vbCrLf &
                       "Make sure your assembly contains part files (.ipt) with Description properties",
                       MsgBoxStyle.Exclamation)
                Return False
            End If

            Dim summaryMsg As String = "STEP 4: IDENTIFY QTY OF SECTIONS" & vbCrLf & vbCrLf
            summaryMsg &= "Found " & m_ComponentGroups.Count & " component groups:" & vbCrLf & vbCrLf

            For Each kvp In m_ComponentGroups
                Dim groupName As String = kvp.Key
                Dim groupDict As Dictionary(Of String, String) = kvp.Value

                Dim groupDescription As String
                Select Case groupName
                    Case "B" : groupDescription = "I and H sections (UB/UC)"
                    Case "PL" : groupDescription = "Platework (PL + S355JR)"
                    Case "LPL" : groupDescription = "Liners (PL + other materials)"
                    Case "A" : groupDescription = "Angles (L sections)"
                    Case "CH" : groupDescription = "Channels (PFC/TFC)"
                    Case "P" : groupDescription = "Circular hollow sections (CHS)"
                    Case "SQ" : groupDescription = "Square/rectangular hollow (SHS)"
                    Case "FL" : groupDescription = "Flatbar (FL)"
                    Case "IPE" : groupDescription = "European I-beams (IPE)"
                    Case Else : groupDescription = "Other components"
                End Select

                summaryMsg &= "[" & groupName & "] " & groupDescription & vbCrLf
                summaryMsg &= "   Quantity: " & groupDict.Count & " components" & vbCrLf

                ' Show first 3 examples
                Dim exampleCount As Integer = Math.Min(3, groupDict.Count)
                Dim idx As Integer = 0
                For Each compKvp In groupDict
                    If idx >= exampleCount Then Exit For
                    Dim parts() As String = compKvp.Value.Split("|"c)
                    Dim description As String = parts(1)
                    Dim fileName As String = parts(2)
                    summaryMsg &= "   - " & fileName & " (" & description & ")" & vbCrLf
                    idx += 1
                Next

                If groupDict.Count > exampleCount Then
                    summaryMsg &= "   - (and " & (groupDict.Count - exampleCount) & " more...)" & vbCrLf
                End If

                summaryMsg &= vbCrLf
            Next

            summaryMsg &= "Continue with renaming each group?"

            Dim result As MsgBoxResult = MsgBox(summaryMsg, MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Component Groups - Step 4")
            If result = MsgBoxResult.No Then
                LogMessage("SUMMARY: User cancelled after reviewing groups")
                Return False
            End If

            LogMessage("SUMMARY: User approved component groups")
            Return True
        End Function

        ''' <summary>
        ''' Get user naming schemes (same as VBScript GetUserNamingSchemes)
        ''' </summary>
        Private Sub GetUserNamingSchemes()
            LogMessage("INPUT: Getting user naming schemes for " & m_ComponentGroups.Count & " groups")

            For Each kvp In m_ComponentGroups
                Dim groupName As String = kvp.Key
                Dim groupDict As Dictionary(Of String, String) = kvp.Value

                ' Build component list
                Dim componentList As String = ""
                For Each compKvp In groupDict
                    Dim parts() As String = compKvp.Value.Split("|"c)
                    Dim description As String = parts(1)
                    Dim fileName As String = parts(2)
                    componentList &= "  - " & fileName & " (" & description & ")" & vbCrLf
                Next

                ' Generate default scheme
                Dim defaultScheme As String
                Select Case groupName
                    Case "B" : defaultScheme = m_PlantSection & "B{N}"
                    Case "PL" : defaultScheme = m_PlantSection & "PL{N}"
                    Case "LPL" : defaultScheme = m_PlantSection & "LPL{N}"
                    Case "A" : defaultScheme = m_PlantSection & "A{N}"
                    Case "CH" : defaultScheme = m_PlantSection & "CH{N}"
                    Case "P" : defaultScheme = m_PlantSection & "P{N}"
                    Case "SQ" : defaultScheme = m_PlantSection & "SQ{N}"
                    Case "FL" : defaultScheme = m_PlantSection & "FL{N}"
                    Case "IPE" : defaultScheme = m_PlantSection & "IPE{N}"
                    Case Else : defaultScheme = m_PlantSection & "PART{N}"
                End Select

                Dim userInput As String = InputBox(
                    "COMPONENT GROUP: " & groupName & " (" & groupDict.Count & " components)" & vbCrLf & vbCrLf &
                    "Plant Section: " & m_PlantSection & vbCrLf & vbCrLf &
                    "Components in this group:" & vbCrLf & componentList & vbCrLf &
                    "Enter naming scheme:" & vbCrLf &
                    "IMPORTANT: Use {N} for auto-numbering!" & vbCrLf & vbCrLf &
                    "Examples:" & vbCrLf &
                    "  " & m_PlantSection & "B{N}   -> " & m_PlantSection & "B1, " & m_PlantSection & "B2..." & vbCrLf &
                    "  " & m_PlantSection & "PL{N}  -> " & m_PlantSection & "PL1, " & m_PlantSection & "PL2..." & vbCrLf & vbCrLf &
                    "WITHOUT {N}, all parts get the SAME name!",
                    "Naming Scheme for Group: " & groupName,
                    defaultScheme)

                If String.IsNullOrEmpty(userInput) Then
                    userInput = defaultScheme
                End If

                m_NamingSchemes.Add(groupName, userInput)
                LogMessage("INPUT: Group '" & groupName & "' -> Scheme: " & userInput)
            Next
        End Sub

        ''' <summary>
        ''' Create dynamic heritage-based copies (same as VBScript CreateDynamicHeritageBasedCopies)
        ''' </summary>
        Private Sub CreateDynamicHeritageBasedCopies(ByVal asmDoc As AssemblyDocument)
            LogMessage("HERITAGE: Creating dynamic heritage-based copies")

            Dim asmDir As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName) & "\"

            ' Scan registry for existing counters
            Dim existingCounters As New Dictionary(Of String, Integer)
            ScanRegistryForCounters(existingCounters, m_PlantSection)

            ' Process each group
            For Each kvp In m_ComponentGroups
                Dim groupName As String = kvp.Key
                Dim namingScheme As String = m_NamingSchemes(groupName)
                Dim groupDict As Dictionary(Of String, String) = kvp.Value

                ' Get starting counter
                Dim prefixGroupKey As String = ExtractPrefixFromScheme(namingScheme) & groupName
                LogMessage("HERITAGE: Checking for existing counter with key: '" & prefixGroupKey & "'")

                Dim startingCounter As Integer
                If existingCounters.ContainsKey(prefixGroupKey) Then
                    startingCounter = existingCounters(prefixGroupKey) + 1
                    LogMessage("HERITAGE: Group '" & groupName & "' continuing from number " & startingCounter)
                Else
                    startingCounter = 1
                    LogMessage("HERITAGE: Group '" & groupName & "' starting from number 1")
                End If

                LogMessage("HERITAGE: Processing group '" & groupName & "' with scheme: " & namingScheme)

                Dim componentCounter As Integer = startingCounter

                For Each compKvp In groupDict
                    Dim pathAndDescAndFile As String = compKvp.Value
                    Dim parts() As String = pathAndDescAndFile.Split("|"c)
                    Dim originalPath As String = parts(0)
                    Dim description As String = parts(1)
                    Dim originalFileName As String = parts(2)

                    ' Generate new filename
                    Dim newFileName As String = GenerateNewFileName(namingScheme, componentCounter)
                    componentCounter += 1

                    ' Create heritage file in same directory as original
                    Dim originalDir As String = System.IO.Path.GetDirectoryName(originalPath) & "\"
                    Dim newPath As String = originalDir & newFileName

                    LogMessage("HERITAGE: " & originalFileName & " -> " & newFileName)

                    ' Store mappings
                    m_FileNameMapping.Add(originalPath, newFileName)
                    Dim mappingValue As String = newPath & "|" & originalFileName & "|" & newFileName & "|" & groupName & "|" & description
                    m_ComprehensiveMapping.Add(originalPath, mappingValue)

                    LogMessage("MAPPING: " & originalPath & " -> " & newPath)

                    ' Create heritage file if it doesn't exist
                    If System.IO.File.Exists(newPath) Then
                        LogMessage("HERITAGE: File already exists: " & newFileName)
                    Else
                        LogMessage("HERITAGE: Creating new file: " & newFileName)

                        Try
                            Dim partDoc As Document = m_InventorApp.Documents.Open(originalPath, False)
                            partDoc.SaveAs(newPath, True)
                            LogMessage("HERITAGE: SUCCESS - Created " & newFileName)
                            partDoc.Close()
                        Catch ex As Exception
                            LogMessage("HERITAGE: ERROR - " & ex.Message)
                        End Try
                    End If
                Next

                ' Save final counter to Registry
                Dim finalCounter As Integer = componentCounter - 1
                SaveCounterToRegistry(prefixGroupKey, finalCounter)
            Next
        End Sub

        ''' <summary>
        ''' Generate new filename (same as VBScript GenerateNewFileName)
        ''' </summary>
        Private Function GenerateNewFileName(ByVal scheme As String, ByVal number As Integer) As String
            Dim result As String = scheme.Replace("{N}", number.ToString())
            If Not result.ToLower().EndsWith(".ipt") Then
                result &= ".ipt"
            End If
            Return result
        End Function

        ''' <summary>
        ''' Update assembly references (same as VBScript UpdateDynamicAssemblyReferences)
        ''' </summary>
        Private Sub UpdateDynamicAssemblyReferences(ByVal asmDoc As AssemblyDocument)
            LogMessage("ASSEMBLY: Recursively updating assembly references across entire model hierarchy")

            UpdateAssemblyReferencesRecursively(asmDoc, "ROOT")

            ' Save main assembly
            LogMessage("ASSEMBLY: Saving assembly with updated references...")
            Try
                asmDoc.Save()
                LogMessage("ASSEMBLY: Successfully saved assembly")
            Catch ex As Exception
                LogMessage("ASSEMBLY: ERROR saving assembly: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Recursively update assembly references (same as VBScript)
        ''' </summary>
        Private Sub UpdateAssemblyReferencesRecursively(ByVal asmDoc As AssemblyDocument, ByVal asmLevel As String)
            LogMessage("ASSEMBLY: Updating references in - " & asmDoc.DisplayName & " (Level: " & asmLevel & ")")

            For i As Integer = 1 To asmDoc.ComponentDefinition.Occurrences.Count
                Try
                    Dim occ As ComponentOccurrence = asmDoc.ComponentDefinition.Occurrences.Item(i)

                    If occ.Suppressed Then
                        Continue For
                    End If

                    Dim doc As Document = occ.Definition.Document
                    Dim fullPath As String = doc.FullFileName
                    Dim fileName As String = System.IO.Path.GetFileName(fullPath)

                    If fileName.ToLower().EndsWith(".ipt") Then
                        If m_ComprehensiveMapping.ContainsKey(fullPath) Then
                            Dim mappingValue As String = m_ComprehensiveMapping(fullPath)
                            Dim mappingParts() As String = mappingValue.Split("|"c)

                            Dim newPath As String = mappingParts(0)
                            Dim originalFileName As String = mappingParts(1)
                            Dim newFileName As String = mappingParts(2)
                            Dim groupName As String = mappingParts(3)

                            LogMessage("ASSEMBLY: Replacing " & originalFileName & " -> " & newFileName & " [" & groupName & "]")

                            Try
                                occ.Replace(newPath, True)
                                LogMessage("ASSEMBLY: SUCCESS - Updated to " & newFileName)
                            Catch ex As Exception
                                LogMessage("ASSEMBLY: ERROR - Replace failed: " & ex.Message)
                            End Try
                        End If

                    ElseIf fileName.ToLower().EndsWith(".iam") Then
                        If Not fileName.ToLower().Contains("bolted connection") Then
                            LogMessage("ASSEMBLY: RECURSING into sub-assembly - " & fileName)
                            Dim subAsm As AssemblyDocument = CType(doc, AssemblyDocument)
                            UpdateAssemblyReferencesRecursively(subAsm, asmLevel & ">" & fileName)

                            Try
                                doc.Save()
                                LogMessage("ASSEMBLY: Saved sub-assembly - " & fileName)
                            Catch ex As Exception
                                LogMessage("ASSEMBLY: ERROR saving sub-assembly: " & ex.Message)
                            End Try
                        End If
                    End If

                Catch ex As Exception
                    LogMessage("ASSEMBLY: Error processing occurrence: " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Update all IDWs in directory (same as VBScript UpdateAllIDWsInDirectory)
        ''' </summary>
        Private Sub UpdateAllIDWsInDirectory(ByVal asmDoc As AssemblyDocument)
            LogMessage("IDW: Auto-detecting IDW files in assembly directory")

            Dim asmDir As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
            Dim idwFiles() As String = Directory.GetFiles(asmDir, "*.idw")

            If idwFiles.Length = 0 Then
                LogMessage("IDW: No IDW files found in directory")
                Return
            End If

            For Each idwPath As String In idwFiles
                LogMessage("IDW: Found drawing file - " & System.IO.Path.GetFileName(idwPath))
                UpdateSingleIDWWithDynamicReferences(idwPath, asmDir)
            Next
        End Sub

        ''' <summary>
        ''' Update single IDW (same as VBScript UpdateSingleIDWWithDynamicReferences)
        ''' </summary>
        Private Sub UpdateSingleIDWWithDynamicReferences(ByVal idwPath As String, ByVal asmDir As String)
            Try
                ' DO NOT call CloseAll() - it closes the active assembly and breaks subsequent steps!
                ' Only close IDW documents after processing them

                LogMessage("IDW: Opening " & System.IO.Path.GetFileName(idwPath))
                Dim idwDoc As DrawingDocument = CType(m_InventorApp.Documents.Open(idwPath, False), DrawingDocument)

                Dim fileDescriptors As Object = idwDoc.File.ReferencedFileDescriptors
                LogMessage("IDW: Found " & fileDescriptors.Count & " referenced files")

                For i As Integer = 1 To fileDescriptors.Count
                    Dim fd As Object = fileDescriptors.Item(i)
                    Dim currentFullPath As String = fd.FullFileName
                    Dim currentFileName As String = System.IO.Path.GetFileName(currentFullPath)

                    If m_ComprehensiveMapping.ContainsKey(currentFullPath) Then
                        Dim mappingValue As String = m_ComprehensiveMapping(currentFullPath)
                        Dim mappingParts() As String = mappingValue.Split("|"c)
                        Dim newPath As String = mappingParts(0)
                        Dim newFileName As String = System.IO.Path.GetFileName(newPath)

                        LogMessage("IDW: Updating reference " & currentFileName & " -> " & newFileName)

                        Try
                            fd.ReplaceReference(newPath)
                            LogMessage("IDW: SUCCESS - Reference updated")
                        Catch ex As Exception
                            LogMessage("IDW: ERROR - ReplaceReference failed: " & ex.Message)
                        End Try
                    ElseIf m_AssemblyPathMapping.ContainsKey(currentFullPath) Then
                        Dim mappedAsmPath As String = m_AssemblyPathMapping(currentFullPath)
                        Dim mappedAsmName As String = System.IO.Path.GetFileName(mappedAsmPath)
                        LogMessage("IDW: Updating main assembly reference " & currentFileName & " -> " & mappedAsmName)

                        Try
                            fd.ReplaceReference(mappedAsmPath)
                            LogMessage("IDW: SUCCESS - Main assembly reference updated")
                        Catch ex As Exception
                            LogMessage("IDW: ERROR - Main assembly ReplaceReference failed: " & ex.Message)
                        End Try
                    Else
                        LogMessage("IDW: INFO - No mapping found for " & currentFileName)
                    End If
                Next

                idwDoc.Save()
                LogMessage("IDW: Saved " & System.IO.Path.GetFileName(idwPath))

            Catch ex As Exception
                LogMessage("IDW: ERROR - " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Rename the main assembly with the selected project prefix and track mapping for IDW updates.
        ''' </summary>
        Private Sub RenameMainAssemblyWithPrefix(ByVal asmDoc As AssemblyDocument)
            Try
                Dim originalAsmPath As String = asmDoc.FullFileName
                Dim originalAsmNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(originalAsmPath)

                Dim newAsmNameNoExt As String = originalAsmNameNoExt
                If Not originalAsmNameNoExt.StartsWith(m_PlantSection, StringComparison.OrdinalIgnoreCase) Then
                    newAsmNameNoExt = m_PlantSection & originalAsmNameNoExt
                End If

                Dim newAsmPath As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(originalAsmPath), newAsmNameNoExt & ".iam")

                If String.Equals(originalAsmPath, newAsmPath, StringComparison.OrdinalIgnoreCase) Then
                    LogMessage("ASSEMBLY RENAME: Main assembly already prefixed - " & System.IO.Path.GetFileName(originalAsmPath))
                    Return
                End If

                If System.IO.File.Exists(newAsmPath) Then
                    LogMessage("ASSEMBLY RENAME: Target already exists, skipping main assembly rename - " & System.IO.Path.GetFileName(newAsmPath))
                    Return
                End If

                asmDoc.SaveAs(newAsmPath, True)
                m_AssemblyPathMapping(originalAsmPath) = newAsmPath
                LogMessage("ASSEMBLY RENAME: " & System.IO.Path.GetFileName(originalAsmPath) & " -> " & System.IO.Path.GetFileName(newAsmPath))

            Catch ex As Exception
                LogMessage("ASSEMBLY RENAME: ERROR - " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Save mapping file (same as VBScript SaveMappingFile)
        ''' </summary>
        Private Sub SaveMappingFile(ByVal activeDoc As AssemblyDocument)
            LogMessage("MAPPING: Saving comprehensive mapping file for STEP 2")

            Dim asmDir As String = System.IO.Path.GetDirectoryName(activeDoc.FullFileName)
            Dim mappingFilePath As String = System.IO.Path.Combine(asmDir, "STEP_1_MAPPING.txt")

            LogMessage("MAPPING: Saving mapping file to: " & mappingFilePath)

            Using writer As New StreamWriter(mappingFilePath, True) ' Append mode
                If Not System.IO.File.Exists(mappingFilePath) OrElse New FileInfo(mappingFilePath).Length = 0 Then
                    writer.WriteLine("# STEP 1 MAPPING FILE - Generated: " & DateTime.Now.ToString())
                    writer.WriteLine("# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description")
                    writer.WriteLine("")
                End If

                For Each kvp In m_ComprehensiveMapping
                    Dim originalPath As String = kvp.Key
                    Dim mappingValue As String = kvp.Value
                    Dim parts() As String = mappingValue.Split("|"c)

                    If parts.Length >= 5 Then
                        Dim newPath As String = parts(0)
                        Dim originalFile As String = parts(1)
                        Dim newFile As String = parts(2)
                        Dim groupName As String = parts(3)
                        Dim description As String = parts(4)

                        writer.WriteLine(originalPath & "|" & newPath & "|" & originalFile & "|" & newFile & "|" & groupName & "|" & description)
                        LogMessage("MAPPING: " & originalFile & " -> " & newFile)
                    End If
                Next

                writer.WriteLine("")
                writer.WriteLine("# End of mapping file")
            End Using

            LogMessage("MAPPING: Saved comprehensive mapping (" & m_ComprehensiveMapping.Count & " mappings)")
        End Sub

        ''' <summary>
        ''' Scan registry for counters (same as VBScript ScanRegistryForCounters)
        ''' </summary>
        Private Sub ScanRegistryForCounters(ByVal existingCounters As Dictionary(Of String, Integer), ByVal userPrefix As String)
            LogMessage("SCAN: Scanning Registry for existing counters with prefix: " & userPrefix)

            Dim counterKeys() As String = {
                userPrefix & "CH", userPrefix & "PL", userPrefix & "B", userPrefix & "A",
                userPrefix & "P", userPrefix & "SQ", userPrefix & "FL", userPrefix & "LPL",
                userPrefix & "IPE", userPrefix & "OTHER"
            }

            Dim foundCount As Integer = 0

            For Each keyName As String In counterKeys
                Try
                    Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer")
                        If regKey IsNot Nothing Then
                            Dim value As Object = regKey.GetValue(keyName)
                            If value IsNot Nothing Then
                                existingCounters.Add(keyName, CInt(value))
                                LogMessage("SCAN: Found existing counter: " & keyName & " = " & value.ToString())
                                foundCount += 1
                            End If
                        End If
                    End Using
                Catch
                    ' Key doesn't exist
                End Try
            Next

            If foundCount > 0 Then
                LogMessage("SCAN: Loaded " & foundCount & " existing counters from Registry")
            Else
                LogMessage("SCAN: No existing counters found in Registry - starting fresh")
            End If
        End Sub

        ''' <summary>
        ''' Save counter to registry (same as VBScript SaveCounterToRegistry)
        ''' </summary>
        Private Sub SaveCounterToRegistry(ByVal prefixGroupKey As String, ByVal finalCounter As Integer)
            Try
                Using regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\InventorRenamer")
                    regKey.SetValue(prefixGroupKey, finalCounter, RegistryValueKind.DWord)
                    LogMessage("REGISTRY: Saved " & prefixGroupKey & " = " & finalCounter)
                End Using
            Catch ex As Exception
                LogMessage("REGISTRY: ERROR - Could not save " & prefixGroupKey & ": " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Extract prefix from scheme (same as VBScript ExtractPrefixFromScheme)
        ''' </summary>
        Private Function ExtractPrefixFromScheme(ByVal namingScheme As String) As String
            Dim lastDashPos As Integer = namingScheme.LastIndexOf("-"c)

            If lastDashPos > 0 Then
                Return namingScheme.Substring(0, lastDashPos + 1)
            Else
                Dim nPos As Integer = namingScheme.IndexOf("{N}")
                If nPos > 0 Then
                    For i As Integer = nPos - 1 To 0 Step -1
                        If Not Char.IsLetter(namingScheme(i)) Then
                            Return namingScheme.Substring(0, i + 1)
                        End If
                    Next
                End If
                Return namingScheme
            End If
        End Function

        ''' <summary>
        ''' Initialize logging
        ''' </summary>
        Private Sub InitializeLog()
            Try
                Dim logsFolder As String = "C:\Users\Quintin\Documents\Spectiv\3. Working\22. DMS plant\REF\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\Logs"
                If Not Directory.Exists(logsFolder) Then
                    Directory.CreateDirectory(logsFolder)
                End If

                Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                m_LogPath = System.IO.Path.Combine(logsFolder, "PartRenamer_" & timestamp & ".log")
                m_LogFile = New StreamWriter(m_LogPath, False, System.Text.Encoding.UTF8)
                m_LogFile.AutoFlush = True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to initialize log: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Log a message
        ''' </summary>
        Private Sub LogMessage(ByVal message As String)
            Try
                If m_LogFile IsNot Nothing Then
                    m_LogFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | " & message)
                End If
                System.Diagnostics.Debug.WriteLine("PartRenamer: " & message)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Close log
        ''' </summary>
        Private Sub CloseLog()
            Try
                If m_LogFile IsNot Nothing Then
                    m_LogFile.Flush()
                    m_LogFile.Close()
                    m_LogFile = Nothing
                End If
            Catch
            End Try
        End Sub

    End Class

End Namespace
