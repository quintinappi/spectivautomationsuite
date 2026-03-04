Imports Inventor
Imports System.IO
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

''' <summary>
''' Part Renamer - Renames parts in place using heritage method.
''' Matches Option 1 (Assembly_Renamer.vbs) flow but uses plugin's silent replacement.
''' NO COPYING - just rename and replace in same location.
''' </summary>
Public Class PartRenamer
    Private m_inventorApp As Inventor.Application
    Private m_log As System.Text.StringBuilder
    Private m_logPath As String

    ' Component grouping
    Private m_componentGroups As Dictionary(Of String, Dictionary(Of String, String))  ' groupCode -> (originalPath -> "path|desc|filename")
    Private m_namingSchemes As Dictionary(Of String, String)  ' groupCode -> naming scheme
    Private m_plantPrefix As String  ' User-defined prefix like "PLANT1-000-"

    ' Mapping dictionaries
    Private m_comprehensiveMapping As Dictionary(Of String, String)  ' originalPath -> "newPath|origFile|newFile|group|desc"
    Private m_partNameMapping As Dictionary(Of String, String)  ' occBaseName -> newOccBaseName (for iLogic)

    ' Registry path for counter persistence
    Private Const REGISTRY_PATH As String = "HKEY_CURRENT_USER\Software\InventorRenamer\"

    Public Sub New(inventorApp As Inventor.Application)
        m_inventorApp = inventorApp
        m_log = New System.Text.StringBuilder()
        m_componentGroups = New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)
        m_namingSchemes = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        m_comprehensiveMapping = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        m_partNameMapping = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Sub

    Private Sub Log(message As String)
        Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss")
        m_log.AppendLine(timestamp & " | " & message)
    End Sub

    Private Sub SaveLog(asmDir As String)
        Try
            m_logPath = System.IO.Path.Combine(asmDir, "PartRenamer_Log_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
            System.IO.File.WriteAllText(m_logPath, m_log.ToString())
        Catch
        End Try
    End Sub

    Public Sub RenameAndReplace()
        Try
            Log("=== PART RENAMER STARTED ===")
            Log("Rename parts in place using heritage method")

            ' Check active document
            Dim activeDoc As Document = m_inventorApp.ActiveDocument
            If activeDoc Is Nothing OrElse activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                Log("ERROR: No assembly document open")
                MsgBox("Please open an assembly document first.", MsgBoxStyle.Exclamation)
                Return
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(activeDoc, AssemblyDocument)
            Dim asmPath As String = asmDoc.FullFileName
            Dim asmDir As String = System.IO.Path.GetDirectoryName(asmPath)

            Log("Assembly: " & asmDoc.DisplayName)
            Log("Location: " & asmDir)

            ' Confirm with user
            Dim occCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
            Dim confirmResult As MsgBoxResult = MsgBox(
                "PART RENAMER" & vbCrLf & vbCrLf &
                "Assembly: " & asmDoc.DisplayName & vbCrLf &
                "Parts: " & occCount & " occurrences" & vbCrLf &
                "Location: " & asmDir & vbCrLf & vbCrLf &
                "This will:" & vbCrLf &
                "1. Group parts by description (PL, B, CH, etc.)" & vbCrLf &
                "2. Create heritage copies with new names" & vbCrLf &
                "3. Update assembly references" & vbCrLf &
                "4. Update IDW drawings" & vbCrLf & vbCrLf &
                "Continue?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Part Renamer")

            If confirmResult = MsgBoxResult.No Then
                Log("User cancelled")
                Return
            End If

            ' STEP 1: Get plant prefix from user
            Log("STEP 1: Getting plant prefix")
            m_plantPrefix = GetPlantPrefix()
            If String.IsNullOrEmpty(m_plantPrefix) Then
                Log("User cancelled - no prefix")
                Return
            End If
            Log("Prefix: " & m_plantPrefix)

            ' STEP 2: Flatten hierarchy and group components (assembly STAYS OPEN)
            Log("STEP 2: Grouping components by description")
            GroupComponentsByDescription(asmDoc)
            Log("Groups created: " & m_componentGroups.Count)

            If m_componentGroups.Count = 0 Then
                Log("ERROR: No components found to rename!")
                Log("This could mean:")
                Log("  - All parts already have heritage names (-000-, -001-, etc.)")
                Log("  - All parts are hardware (bolts, nuts, washers)")
                Log("  - No .ipt files in assembly")
                SaveLog(asmDir)
                MsgBox("No components found to rename!" & vbCrLf & vbCrLf &
                       "Check log file for details:" & vbCrLf & m_logPath, MsgBoxStyle.Exclamation)
                Return
            End If

            ' STEP 3: Show groups summary and confirm
            Log("STEP 3: Showing groups summary")
            If Not ShowGroupsSummary() Then
                Log("User cancelled after groups summary")
                Return
            End If

            ' STEP 4: Get naming schemes per group
            Log("STEP 4: Getting naming schemes")
            GetNamingSchemes()

            ' STEP 5: Create heritage copies (SaveAs in SAME directory)
            Log("STEP 5: Creating heritage copies")
            CreateHeritageCopies()

            ' STEP 6: Update assembly references (assembly STILL OPEN)
            Log("STEP 6: Updating assembly references")
            UpdateAssemblyReferencesInOpenDoc(asmDoc)

            ' STEP 7: Save assembly
            Log("STEP 7: Saving assembly")
            asmDoc.Save()

            ' STEP 8: Update IDW files
            Log("STEP 8: Updating IDW files")
            UpdateIDWFiles(asmDir)

            ' STEP 9: Save mapping file
            Log("STEP 9: Saving mapping file")
            SaveMappingFile(asmDir)

            Log("=== PART RENAMER COMPLETED ===")
            SaveLog(asmDir)

            MsgBox("Part Renamer Complete!" & vbCrLf & vbCrLf &
                   "Parts renamed: " & m_comprehensiveMapping.Count & vbCrLf &
                   "Assembly updated" & vbCrLf &
                   "IDWs updated" & vbCrLf &
                   "Mapping file saved" & vbCrLf & vbCrLf &
                   "Log: " & m_logPath, MsgBoxStyle.Information, "Success!")

        Catch ex As Exception
            Log("EXCEPTION: " & ex.Message)
            Log("Stack: " & ex.StackTrace)

            Try
                Dim errorLogPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PartRenamer_Error_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
                System.IO.File.WriteAllText(errorLogPath, m_log.ToString())
                MsgBox("Error: " & ex.Message & vbCrLf & vbCrLf & "Log: " & errorLogPath, MsgBoxStyle.Critical)
            Catch
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End Try
    End Sub

    Private Function GetPlantPrefix() As String
        Dim userInput As String = InputBox(
            "DEFINE PROJECT PREFIX" & vbCrLf & vbCrLf &
            "Enter the project prefix (as per drawing register):" & vbCrLf & vbCrLf &
            "Examples:" & vbCrLf &
            "  PLANT1-000-" & vbCrLf &
            "  AREA2-000-" & vbCrLf &
            "  SSCR05-001-" & vbCrLf & vbCrLf &
            "This will create part numbers like:" & vbCrLf &
            "  PREFIX-B1, PREFIX-PL1, PREFIX-CH1, etc.",
            "Define Project Prefix", "PLANT1-000-")

        If String.IsNullOrEmpty(userInput) Then
            Return Nothing
        End If

        userInput = userInput.Trim().ToUpper()

        ' Ensure ends with dash
        If Not userInput.EndsWith("-") Then
            userInput = userInput & "-"
        End If

        Return userInput
    End Function

    Private Sub GroupComponentsByDescription(asmDoc As AssemblyDocument)
        Dim uniqueParts As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        ProcessAssemblyRecursively(asmDoc, uniqueParts, 0)
    End Sub

    Private Sub ProcessAssemblyRecursively(asmDoc As AssemblyDocument, uniqueParts As Dictionary(Of String, Boolean), level As Integer)
        Dim indent As String = New String(" "c, level * 2)

        Log(indent & "Processing: " & asmDoc.DisplayName & " (" & asmDoc.ComponentDefinition.Occurrences.Count & " occurrences)")

        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Try
                If occ.Suppressed Then
                    Log(indent & "  SKIP (suppressed): " & occ.Name)
                    Continue For
                End If

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch ex As Exception
                    Log(indent & "  SKIP (no doc): " & occ.Name & " - " & ex.Message)
                    Continue For
                End Try

                If doc Is Nothing Then
                    Log(indent & "  SKIP (null doc): " & occ.Name)
                    Continue For
                End If

                Dim fullPath As String = doc.FullFileName
                Dim fileName As String = System.IO.Path.GetFileName(fullPath)

                Log(indent & "  Checking: " & fileName)

                If fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    ' Skip if already heritage-renamed
                    ' Check if filename matches heritage pattern: PREFIX-###-GROUPNUM.ipt
                    ' e.g., NCRH01-000-PL1.ipt, PLANT1-000-B23.ipt
                    ' But NOT: Support Beam-2_Clone_PLANT1-000-PL1.ipt (this is a cloned file, not heritage)
                    If IsHeritageRenamed(fileName) Then
                        Log(indent & "    SKIP (already heritage renamed): " & fileName)
                        Continue For
                    End If

                    ' Skip if already processed
                    If uniqueParts.ContainsKey(fullPath) Then
                        Log(indent & "    SKIP (duplicate): " & fileName)
                        Continue For
                    End If

                    uniqueParts.Add(fullPath, True)

                    ' Get description from iProperty
                    Dim description As String = GetDescriptionFromIProperty(doc)
                    Log(indent & "    Description: '" & description & "'")

                    If String.IsNullOrEmpty(description) Then
                        Log(indent & "    WARNING: No description - adding to OTHER group")
                        description = "UNKNOWN"
                    End If

                    ' Classify by description
                    Dim groupCode As String = ClassifyByDescription(description)
                    Log(indent & "    Classified as: " & groupCode)

                    If groupCode = "SKIP" Then
                        Log(indent & "    SKIP (hardware): " & fileName & " - " & description)
                        Continue For
                    End If

                    Log(indent & "    ADDING: " & fileName & " -> [" & groupCode & "] " & description)

                    ' Add to group
                    If Not m_componentGroups.ContainsKey(groupCode) Then
                        m_componentGroups.Add(groupCode, New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))
                    End If

                    Dim groupDict As Dictionary(Of String, String) = m_componentGroups(groupCode)
                    If Not groupDict.ContainsKey(fullPath) Then
                        groupDict.Add(fullPath, fullPath & "|" & description & "|" & fileName)
                    End If

                ElseIf fullPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    ' Sub-assembly - recurse
                    If fullPath.ToLower().Contains("bolted connection") Then
                        Log(indent & "    SKIP (bolted connection): " & fileName)
                    Else
                        Log(indent & "    RECURSING: " & fileName)
                        Dim subAsm As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                        ProcessAssemblyRecursively(subAsm, uniqueParts, level + 1)
                    End If
                Else
                    Log(indent & "    SKIP (unknown type): " & fileName)
                End If

            Catch ex As Exception
                Log(indent & "  ERROR: " & ex.Message)
            End Try
        Next

        Log(indent & "Done with: " & asmDoc.DisplayName & " - Total groups so far: " & m_componentGroups.Count)
    End Sub

    Private Function GetDescriptionFromIProperty(doc As Document) As String
        Try
            Dim propSet As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Dim descProp As [Property] = propSet.Item("Description")
            Return If(descProp.Value IsNot Nothing, descProp.Value.ToString().Trim(), "")
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Check if filename is already a heritage-renamed file.
    ''' Heritage pattern: starts with PREFIX-###-GROUP (e.g., NCRH01-000-PL1.ipt, PLANT1-000-B23.ipt)
    ''' NOT heritage: Support Beam-2_Clone_PLANT1-000-PL1.ipt (contains prefix but doesn't start with it)
    ''' </summary>
    Private Function IsHeritageRenamed(fileName As String) As Boolean
        ' Remove extension
        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(fileName)

        ' Heritage files follow pattern: LETTERS/NUMBERS - ### - GROUPCODE + NUMBER
        ' e.g., NCRH01-000-PL1, PLANT1-000-B23, SSCR05-001-CH5

        ' Must start with alphanumeric prefix, then have -###- pattern near the start
        ' Pattern: ^[A-Z0-9]+-\d{3}-[A-Z]+\d+$

        Try
            ' Simple check: if it starts with something like "XXXX-000-" or "XXXX-001-" etc.
            ' and the total length is relatively short (heritage names are compact)

            ' Split by dash
            Dim parts() As String = baseName.Split("-"c)

            ' Heritage pattern has at least 3 parts: PREFIX-###-GROUPNUM
            If parts.Length < 3 Then Return False

            ' Second part should be 3 digits (000, 001, 751, 752, etc.)
            Dim numPart As String = parts(1)
            If numPart.Length <> 3 Then Return False
            If Not IsNumeric(numPart) Then Return False

            ' Third part should start with letters (group code) followed by numbers
            Dim groupPart As String = parts(2)
            If groupPart.Length < 2 Then Return False

            ' Check if it starts with known group codes
            Dim knownGroups() As String = {"PL", "B", "CH", "A", "P", "SQ", "FL", "LPL", "IPE", "OTHER"}
            Dim startsWithGroup As Boolean = False
            For Each grp As String In knownGroups
                If groupPart.StartsWith(grp, StringComparison.OrdinalIgnoreCase) Then
                    ' Check that what follows is a number
                    Dim remainder As String = groupPart.Substring(grp.Length)
                    If remainder.Length > 0 AndAlso IsNumeric(remainder) Then
                        startsWithGroup = True
                        Exit For
                    End If
                End If
            Next

            Return startsWithGroup

        Catch
            Return False
        End Try
    End Function

    Private Function IsNumeric(s As String) As Boolean
        If String.IsNullOrEmpty(s) Then Return False
        For Each c As Char In s
            If Not Char.IsDigit(c) Then Return False
        Next
        Return True
    End Function

    Private Function ClassifyByDescription(description As String) As String
        Dim desc As String = description.ToUpper().Trim()

        ' Skip hardware
        If desc.Contains("BOLT") OrElse desc.Contains("SCREW") OrElse desc.Contains("WASHER") OrElse desc.Contains("NUT") Then
            Return "SKIP"
        End If

        ' Classification logic matching client requirements
        If desc.StartsWith("UB") OrElse desc.StartsWith("UC") Then
            Return "B"  ' Beams/Columns
        ElseIf desc.StartsWith("PL") Then
            If desc.Contains("S355JR") Then
                Return "PL"  ' Platework
            Else
                Return "LPL" ' Liners
            End If
        ElseIf desc.StartsWith("L") AndAlso desc.Contains("X") Then
            Return "A"  ' Angles
        ElseIf desc.StartsWith("PFC") OrElse desc.StartsWith("TFC") Then
            Return "CH" ' Channels
        ElseIf desc.StartsWith("CHS") Then
            Return "P"  ' Circular hollow
        ElseIf desc.StartsWith("SHS") Then
            Return "SQ" ' Square hollow
        ElseIf desc.StartsWith("FL") AndAlso Not desc.Contains("FLOOR") Then
            Return "FL" ' Flatbar
        ElseIf desc.StartsWith("IPE") Then
            Return "IPE" ' European I-beams
        Else
            Return "OTHER"
        End If
    End Function

    Private Function ShowGroupsSummary() As Boolean
        Dim summary As String = "COMPONENT GROUPS FOUND:" & vbCrLf & vbCrLf

        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In m_componentGroups
            Dim groupCode As String = kvp.Key
            Dim groupDict As Dictionary(Of String, String) = kvp.Value

            Dim groupDesc As String = GetGroupDescription(groupCode)
            summary &= "[" & groupCode & "] " & groupDesc & vbCrLf
            summary &= "   Quantity: " & groupDict.Count & " parts" & vbCrLf

            ' Show first 3 examples
            Dim count As Integer = 0
            For Each item As KeyValuePair(Of String, String) In groupDict
                If count >= 3 Then Exit For
                Dim parts() As String = item.Value.Split("|"c)
                Dim fileName As String = parts(2)
                Dim desc As String = parts(1)
                summary &= "   - " & fileName & " (" & desc & ")" & vbCrLf
                count += 1
            Next

            If groupDict.Count > 3 Then
                summary &= "   - (and " & (groupDict.Count - 3) & " more...)" & vbCrLf
            End If

            summary &= vbCrLf
        Next

        summary &= "Continue with renaming?"

        Dim result As MsgBoxResult = MsgBox(summary, MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Component Groups")
        Return result = MsgBoxResult.Yes
    End Function

    Private Function GetGroupDescription(groupCode As String) As String
        Select Case groupCode
            Case "B" : Return "I and H sections (UB/UC)"
            Case "PL" : Return "Platework (PL + S355JR)"
            Case "LPL" : Return "Liners (PL + other)"
            Case "A" : Return "Angles (L sections)"
            Case "CH" : Return "Channels (PFC/TFC)"
            Case "P" : Return "Circular hollow (CHS)"
            Case "SQ" : Return "Square hollow (SHS)"
            Case "FL" : Return "Flatbar (FL)"
            Case "IPE" : Return "European I-beams (IPE)"
            Case Else : Return "Other components"
        End Select
    End Function

    Private Sub GetNamingSchemes()
        For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In m_componentGroups
            Dim groupCode As String = kvp.Key
            Dim defaultScheme As String = m_plantPrefix & groupCode & "{N}"

            Dim userInput As String = InputBox(
                "NAMING SCHEME FOR: " & groupCode & vbCrLf & vbCrLf &
                GetGroupDescription(groupCode) & vbCrLf &
                "Count: " & kvp.Value.Count & " parts" & vbCrLf & vbCrLf &
                "Use {N} for auto-numbering:" & vbCrLf &
                "  " & m_plantPrefix & groupCode & "{N} -> " & m_plantPrefix & groupCode & "1, " & m_plantPrefix & groupCode & "2..." & vbCrLf & vbCrLf &
                "Enter naming scheme:",
                "Naming Scheme - " & groupCode, defaultScheme)

            If String.IsNullOrEmpty(userInput) Then
                userInput = defaultScheme
            End If

            m_namingSchemes.Add(groupCode, userInput)
            Log("SCHEME: " & groupCode & " -> " & userInput)
        Next
    End Sub

    Private Sub CreateHeritageCopies()
        ' Load existing counters from registry
        Dim existingCounters As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        LoadCountersFromRegistry(existingCounters)

        For Each groupKvp As KeyValuePair(Of String, Dictionary(Of String, String)) In m_componentGroups
            Dim groupCode As String = groupKvp.Key
            Dim groupDict As Dictionary(Of String, String) = groupKvp.Value
            Dim namingScheme As String = m_namingSchemes(groupCode)

            ' Get starting counter
            Dim prefixGroupKey As String = m_plantPrefix & groupCode
            Dim counter As Integer = 1

            If existingCounters.ContainsKey(prefixGroupKey) Then
                counter = existingCounters(prefixGroupKey) + 1
                Log("GROUP " & groupCode & ": Continuing from " & counter)
            Else
                Log("GROUP " & groupCode & ": Starting from 1")
            End If

            For Each partKvp As KeyValuePair(Of String, String) In groupDict
                Dim originalPath As String = partKvp.Key
                Dim parts() As String = partKvp.Value.Split("|"c)
                Dim description As String = parts(1)
                Dim originalFileName As String = parts(2)

                ' Generate new filename
                Dim newFileName As String = namingScheme.Replace("{N}", counter.ToString())
                If Not newFileName.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    newFileName &= ".ipt"
                End If

                ' New path is in SAME directory as original
                Dim originalDir As String = System.IO.Path.GetDirectoryName(originalPath)
                Dim newPath As String = System.IO.Path.Combine(originalDir, newFileName)

                Log("HERITAGE: " & originalFileName & " -> " & newFileName)

                ' Store mapping
                Dim mappingValue As String = newPath & "|" & originalFileName & "|" & newFileName & "|" & groupCode & "|" & description
                m_comprehensiveMapping.Add(originalPath, mappingValue)

                ' Build occurrence name mapping for iLogic
                Dim origOccName As String = System.IO.Path.GetFileNameWithoutExtension(originalFileName)
                Dim newOccName As String = System.IO.Path.GetFileNameWithoutExtension(newFileName)
                If Not m_partNameMapping.ContainsKey(origOccName) Then
                    m_partNameMapping.Add(origOccName, newOccName)
                End If

                ' Create heritage file if it doesn't exist
                If Not System.IO.File.Exists(newPath) Then
                    Try
                        Dim partDoc As PartDocument = DirectCast(m_inventorApp.Documents.Open(originalPath, False), PartDocument)
                        partDoc.SaveAs(newPath, True)  ' True = copy
                        partDoc.Close()
                        Log("  CREATED: " & newFileName)
                    Catch ex As Exception
                        Log("  ERROR creating " & newFileName & ": " & ex.Message)
                    End Try
                Else
                    Log("  EXISTS: " & newFileName)
                End If

                counter += 1
            Next

            ' Save final counter to registry
            SaveCounterToRegistry(prefixGroupKey, counter - 1)
        Next
    End Sub

    Private Sub LoadCountersFromRegistry(counters As Dictionary(Of String, Integer))
        Dim groupCodes() As String = {"B", "PL", "LPL", "A", "CH", "P", "SQ", "FL", "IPE", "OTHER"}

        For Each code As String In groupCodes
            Dim keyName As String = m_plantPrefix & code
            Try
                Dim value As Object = Microsoft.Win32.Registry.GetValue("HKEY_CURRENT_USER\Software\InventorRenamer", keyName, Nothing)
                If value IsNot Nothing Then
                    counters.Add(keyName, CInt(value))
                    Log("REGISTRY: Loaded " & keyName & " = " & value.ToString())
                End If
            Catch
            End Try
        Next
    End Sub

    Private Sub SaveCounterToRegistry(keyName As String, value As Integer)
        Try
            Microsoft.Win32.Registry.SetValue("HKEY_CURRENT_USER\Software\InventorRenamer", keyName, value, Microsoft.Win32.RegistryValueKind.DWord)
            Log("REGISTRY: Saved " & keyName & " = " & value)
        Catch ex As Exception
            Log("REGISTRY ERROR: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateAssemblyReferencesInOpenDoc(asmDoc As AssemblyDocument)
        Log("  Updating references recursively...")
        UpdateReferencesRecursively(asmDoc, 0)
    End Sub

    Private Sub UpdateReferencesRecursively(asmDoc As AssemblyDocument, level As Integer)
        Dim indent As String = New String(" "c, level * 2)

        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Try
                If occ.Suppressed Then Continue For

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch
                    Continue For
                End Try

                If doc Is Nothing Then Continue For

                Dim fullPath As String = doc.FullFileName

                If fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    If m_comprehensiveMapping.ContainsKey(fullPath) Then
                        Dim mappingParts() As String = m_comprehensiveMapping(fullPath).Split("|"c)
                        Dim newPath As String = mappingParts(0)
                        Dim newFileName As String = mappingParts(2)

                        Log(indent & "REPLACING: " & System.IO.Path.GetFileName(fullPath) & " -> " & newFileName)

                        ' Key: occ.Replace on OPEN document works silently
                        occ.Replace(newPath, True)
                    End If

                ElseIf fullPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    If Not fullPath.ToLower().Contains("bolted connection") Then
                        Dim subAsm As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                        UpdateReferencesRecursively(subAsm, level + 1)
                        subAsm.Save()
                    End If
                End If

            Catch ex As Exception
                Log(indent & "ERROR: " & ex.Message)
            End Try
        Next
    End Sub

    Private Sub UpdateIDWFiles(asmDir As String)
        Dim idwFiles() As String = System.IO.Directory.GetFiles(asmDir, "*.idw")
        Log("  Found " & idwFiles.Length & " IDW files")

        If idwFiles.Length = 0 Then Return

        m_inventorApp.SilentOperation = True

        For Each idwPath As String In idwFiles
            Try
                Log("  Processing: " & System.IO.Path.GetFileName(idwPath))

                Dim idwDoc As DrawingDocument = DirectCast(m_inventorApp.Documents.Open(idwPath, False), DrawingDocument)

                Dim updatedCount As Integer = 0
                For Each fd As FileDescriptor In idwDoc.File.ReferencedFileDescriptors
                    Try
                        Dim refPath As String = fd.FullFileName

                        If m_comprehensiveMapping.ContainsKey(refPath) Then
                            Dim mappingParts() As String = m_comprehensiveMapping(refPath).Split("|"c)
                            Dim newPath As String = mappingParts(0)

                            fd.ReplaceReference(newPath)
                            updatedCount += 1
                        End If
                    Catch
                    End Try
                Next

                idwDoc.Save()
                idwDoc.Close()
                Log("    Updated " & updatedCount & " references")

            Catch ex As Exception
                Log("  IDW ERROR: " & ex.Message)
            End Try
        Next

        m_inventorApp.SilentOperation = False
    End Sub

    Private Sub SaveMappingFile(asmDir As String)
        Dim mappingPath As String = System.IO.Path.Combine(asmDir, "STEP_1_MAPPING.txt")

        Using writer As New System.IO.StreamWriter(mappingPath, True)  ' Append mode
            writer.WriteLine("# Part Renamer Mapping - " & DateTime.Now.ToString())
            writer.WriteLine("# Format: OriginalPath|NewPath|OriginalFile|NewFile|Group|Description")
            writer.WriteLine()

            For Each kvp As KeyValuePair(Of String, String) In m_comprehensiveMapping
                Dim originalPath As String = kvp.Key
                Dim mappingParts() As String = kvp.Value.Split("|"c)
                writer.WriteLine(originalPath & "|" & kvp.Value)
            Next

            writer.WriteLine()
        End Using

        Log("MAPPING: Saved to " & mappingPath)
    End Sub
End Class
