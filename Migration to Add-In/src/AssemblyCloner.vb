' ============================================================================
' INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER
' ============================================================================
' Description: Clone assembly with all sub-assemblies and parts to new location
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' Ported from: Assembly_Cloner.vbs
'
' MAIN FUNCTIONALITY:
' 1. Detect currently open assembly in Inventor
' 2. Ask for destination folder
' 3. Copy assembly, ALL sub-assemblies, parts, and IDW files to destination
' 4. Update all references in copied assemblies to use local copies
' 5. Recursively update ALL IDW files in ALL subfolders
' 6. Optionally apply heritage renaming to all parts
' 7. Generate STEP_1_MAPPING.txt for reference tracking
' ============================================================================
'
' MIGRATION STATUS: Step-by-step port in progress
'
' PHASE 1: Core Infrastructure ✅ COMPLETE
'   ✅ Logger class
'   ✅ FileHelper class
'   ✅ RegistryManager class
'   ✅ ValidateActiveDocument() - validates active .iam with confirmation
'   ✅ GetDestinationFolder() - folder browser with validation
'   ✅ GetPrefixFromUser() - heritage naming prefix input
'
' PHASE 2: File Collection ✅ COMPLETE
'   ✅ CollectReferencedParts() - wrapper for recursive assembly scan
'   ✅ CollectPartsRecursively() - recursive assembly traversal
'   ✅ CollectIDWFiles() - wrapper for IDW discovery
'   ✅ CollectIDWFilesRecursive() - recursive folder scan
'   ✅ GetDescriptionFromIProperty() - read Description iProperty
'
' PHASE 3: File Operations ✅ COMPLETE
'   ✅ CopyAllFiles() - copy with optional heritage renaming
'   ✅ GroupPartsForRenaming() - classify parts into groups
'   ✅ InitializeNamingSchemes() - create naming schemes
'   ✅ PartClassifier class - part classification logic (separate file)
'
' PHASE 4: Reference Updates ✅ COMPLETE
'   ✅ UpdateAssemblyReferences() - preload, open, update, save
'   ✅ UpdateReferencesInAssembly() - single assembly reference update
'   ✅ UpdateIDWReferences() - update drawing references
'   ✅ FindIDWFilesRecursive() - scan for IDW files
'
' PHASE 5: Registry & Mapping ✅ COMPLETE
'   ✅ WriteMappingFile() - create STEP_1_MAPPING.txt files
'   ✅ ValidateClone() - final validation and report
'   ✅ Registry operations - in RegistryManager.vb
'   ✅ File inventory - in FileHelper.vb
'
' ============================================================================
' MIGRATION STATUS: ✅ COMPLETE - ALL 5 PHASES IMPLEMENTED
' ============================================================================
' Date Completed: 2025-01-21
' Total Phases: 5/5 (100%)
' ============================================================================
'
' ============================================================================

Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.VisualBasic

Namespace SpectivInventorSuite

    ''' <summary>
    ''' Main Assembly Cloner class - ports Assembly_Cloner.vbs functionality
    ''' </summary>
    Public Class AssemblyCloner

        ' ========================================================================
        ' PRIVATE MEMBERS
        ' ========================================================================

        ' Core objects
        Private m_invApp As InventorApplication
        Private m_logger As Logger
        Private m_registryManager As RegistryManager

        ' Tracking dictionaries
        Private m_copiedFiles As Dictionary(Of String, String)    ' originalPath -> newPath
        Private m_componentGroups As Dictionary(Of String, String) ' partName -> group
        Private m_namingSchemes As Dictionary(Of String, String)   ' group -> scheme

        ' Settings
        Private m_plantSection As String   ' Prefix for heritage naming (e.g., "CLONE-001-")
        Private m_doRename As Boolean      ' Whether to rename parts
        Private m_sourceRoot As String     ' Source root folder

        ' ========================================================================
        ' CONSTRUCTOR
        ' ========================================================================

        ''' <summary>
        ''' Constructor - takes Inventor application reference
        ''' </summary>
        Public Sub New(invApp As InventorApplication)
            If invApp Is Nothing Then
                Throw New ArgumentNullException(NameOf(invApp))
            End If

            m_invApp = invApp
            m_logger = New Logger()
            m_registryManager = New RegistryManager()

            ' Initialize dictionaries
            m_copiedFiles = New Dictionary(Of String, String)()
            m_componentGroups = New Dictionary(Of String, String)()
            m_namingSchemes = New Dictionary(Of String, String)()

            ' Initialize logger
            m_logger.Initialize()
        End Sub

        ' ========================================================================
        ' MAIN ENTRY POINT
        ' ========================================================================

        ''' <summary>
        ''' Main clone method - equivalent to ASSEMBLY_CLONER_MAIN()
        ''' This is the primary entry point for the cloning operation
        ''' </summary>
        ''' <param name="destinationFolder">Destination folder path (empty = ask user)</param>
        ''' <param name="renameParts">Whether to rename parts with heritage naming</param>
        ''' <param name="prefix">Prefix for heritage naming (used if renameParts = true)</param>
        ''' <param name="progressCallback">Optional progress callback</param>
        ''' <returns>True if successful, false otherwise</returns>
        Public Function Clone( _
            destinationFolder As String, _
            renameParts As Boolean, _
            prefix As String, _
            Optional progressCallback As ProgressCallback = Nothing _
        ) As Boolean

            Try
                m_logger.Log("========================================")
                m_logger.Log("ASSEMBLY CLONER - STARTING")
                m_logger.Log("========================================")

                ' ========================================================================
                ' STEP 1: Validate Active Document
                ' VBScript: DetectOpenAssembly()
                ' ========================================================================
                UpdateProgress(progressCallback, 5, "Validating active assembly...")
                If Not ValidateActiveDocument() Then
                    Throw New Exception("No active assembly document found")
                End If

                Dim sourceDoc As AssemblyDocument = DirectCast(m_invApp.ActiveDocument, AssemblyDocument)
                m_logger.Log("Source Assembly: " & sourceDoc.FullFileName)

                ' ========================================================================
                ' STEP 2: Get Destination Folder
                ' VBScript: GetDestinationFolder()
                ' ========================================================================
                UpdateProgress(progressCallback, 10, "Getting destination folder...")
                Dim destFolder As String = destinationFolder

                If String.IsNullOrWhiteSpace(destFolder) Then
                    destFolder = GetDestinationFolder()
                    If String.IsNullOrWhiteSpace(destFolder) Then
                        m_logger.Log("User cancelled - no destination folder")
                        Return False
                    End If
                End If

                m_logger.Log("Destination Folder: " & destFolder)

                ' ========================================================================
                ' STEP 3: Configure Renaming Options
                ' VBScript: GetPlantSectionNaming(), GetUserNamingSchemes()
                ' ========================================================================
                UpdateProgress(progressCallback, 15, "Configuring renaming options...")
                m_doRename = renameParts

                If m_doRename Then
                    m_plantSection = prefix
                    If String.IsNullOrWhiteSpace(m_plantSection) Then
                        m_plantSection = GetPrefixFromUser()
                    End If
                    m_logger.Log("Heritage Prefix: " & m_plantSection)
                End If

                ' ========================================================================
                ' STEP 4: Analyze Assembly Structure
                ' VBScript: CollectAllReferencedParts(), CollectIDWFiles()
                ' ========================================================================
                UpdateProgress(progressCallback, 20, "Analyzing assembly structure...")
                Dim sourceDir As String = Path.GetDirectoryName(sourceDoc.FullFileName)
                m_sourceRoot = sourceDir

                ' Collect all referenced parts and sub-assemblies
                Dim allParts As New Dictionary(Of String, String)()
                CollectReferencedParts(sourceDoc, allParts)

                ' Collect IDW files
                CollectIDWFiles(sourceDir, allParts)

                m_logger.Log("Found " & allParts.Count & " unique files to copy")

                ' ========================================================================
                ' STEP 5: Group Parts for Renaming (if enabled)
                ' VBScript: GroupPartsForRenaming()
                ' ========================================================================
                If m_doRename Then
                    UpdateProgress(progressCallback, 25, "Grouping parts for heritage naming...")
                    GroupPartsForRenaming(allParts)
                End If

                ' ========================================================================
                ' STEP 6: Copy Assembly File
                ' ========================================================================
                UpdateProgress(progressCallback, 30, "Copying assembly file...")
                Dim newAsmPath As String = Path.Combine(destFolder, Path.GetFileName(sourceDoc.FullFileName))

                ' Close source document before copying
                Dim sourceFullPath As String = sourceDoc.FullFileName
                sourceDoc.Close()

                ' Copy assembly file
                File.Copy(sourceFullPath, newAsmPath, True)
                m_logger.Log("Copied: " & Path.GetFileName(sourceFullPath) & " -> " & Path.GetFileName(newAsmPath))

                ' Store in copied files dictionary
                m_copiedFiles(sourceFullPath) = newAsmPath

                ' ========================================================================
                ' STEP 7: Copy All Parts and Sub-Assemblies
                ' VBScript: CopyAllFiles()
                ' ========================================================================
                UpdateProgress(progressCallback, 40, "Copying parts and sub-assemblies...")
                CopyAllFiles(allParts, destFolder)

                ' ========================================================================
                ' STEP 8: Update Assembly References
                ' VBScript: UpdateInMemoryAssemblyReferences()
                ' ========================================================================
                UpdateProgress(progressCallback, 60, "Updating assembly references...")
                UpdateAssemblyReferences(newAsmPath)

                ' ========================================================================
                ' STEP 9: Update IDW Drawing References
                ' VBScript: UpdateIDWReferences()
                ' ========================================================================
                UpdateProgress(progressCallback, 80, "Updating IDW references...")
                UpdateIDWReferences(destFolder)

                ' ========================================================================
                ' STEP 10: Write Mapping File
                ' VBScript: WriteMappingFile()
                ' ========================================================================
                UpdateProgress(progressCallback, 90, "Writing mapping file...")
                WriteMappingFile(destFolder)

                ' ========================================================================
                ' STEP 11: Validation and Cleanup
                ' VBScript: ValidateCloneAndLog()
                ' ========================================================================
                UpdateProgress(progressCallback, 95, "Validating clone...")
                ValidateClone(sourceDir, destFolder)

                ' ========================================================================
                ' COMPLETE
                ' ========================================================================
                UpdateProgress(progressCallback, 100, "Completed successfully!")
                m_logger.Log("========================================")
                m_logger.Log("ASSEMBLY CLONER - COMPLETED SUCCESSFULLY")
                m_logger.Log("========================================")
                m_logger.Log("Files copied: " & m_copiedFiles.Count)

                Return True

            Catch ex As Exception
                m_logger.LogError("Clone failed: " & ex.Message)
                UpdateProgress(progressCallback, 0, "Failed: " & ex.Message)
                Return False
            End Try
        End Function

        ' ========================================================================
        ' PROGRESS CALLBACK DELEGATE
        ' ========================================================================

        ''' <summary>
        ''' Progress callback delegate
        ''' </summary>
        Public Delegate Sub ProgressCallback(percent As Integer, status As String)

        ''' <summary>
        ''' Update progress safely
        ''' </summary>
        Private Sub UpdateProgress(callback As ProgressCallback, percent As Integer, status As String)
            Try
                If callback IsNot Nothing Then
                    callback(percent, status)
                End If
            Catch ex As Exception
                ' Ignore progress errors
            End Try
        End Sub

        ' ========================================================================
        ' PHASE 1: CORE INFRASTRUCTURE METHODS
        ' ========================================================================

        ''' <summary>
        ''' Validate that we have an active assembly document
        ''' VBScript: DetectOpenAssembly()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Function ValidateActiveDocument() As Boolean
            Try
                ' Check ActiveDocument exists
                If m_invApp.ActiveDocument Is Nothing Then
                    m_logger.Log("ERROR: No active document found")
                    MessageBox.Show(
                        "No document is currently open in Inventor!" & vbCrLf & vbCrLf &
                        "Please open an assembly file first.",
                        "No Active Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
                    Return False
                End If

                Dim activeDoc As Document = m_invApp.ActiveDocument

                ' Check document type
                If activeDoc.Type <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    m_logger.Log("ERROR: Document is not an assembly. Type: " & activeDoc.Type.ToString())
                    MessageBox.Show(
                        "Current file is not an assembly (.iam)!" & vbCrLf & vbCrLf &
                        "File: " & activeDoc.DisplayName & vbCrLf &
                        "Type: " & activeDoc.Type.ToString() & vbCrLf & vbCrLf &
                        "Please open an assembly file.",
                        "Not an Assembly",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
                    Return False
                End If

                ' Check file extension
                Dim extension As String = Path.GetExtension(activeDoc.FullFileName).ToLower()
                If extension <> ".iam" Then
                    m_logger.Log("ERROR: File extension is not .iam: " & activeDoc.FullFileName)
                    MessageBox.Show(
                        "Current file does not have .iam extension!" & vbCrLf & vbCrLf &
                        "File: " & activeDoc.FullFileName & vbCrLf & vbCrLf &
                        "Please open an assembly file.",
                        "Invalid File Type",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation)
                    Return False
                End If

                ' Get assembly details
                Dim asmDoc As AssemblyDocument = DirectCast(activeDoc, AssemblyDocument)
                Dim occCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
                Dim folderPath As String = Path.GetDirectoryName(activeDoc.FullFileName)

                m_logger.Log("DETECTED: " & activeDoc.DisplayName)
                m_logger.Log("DETECTED: Full path - " & activeDoc.FullFileName)
                m_logger.Log("DETECTED: Occurrences - " & occCount.ToString())

                ' Confirm with user
                Dim confirmMsg As String =
                    "SOURCE ASSEMBLY DETECTED" & vbCrLf & vbCrLf &
                    "Assembly: " & activeDoc.DisplayName & vbCrLf &
                    "Parts Count: " & occCount.ToString() & " occurrences" & vbCrLf &
                    "Location: " & folderPath & vbCrLf & vbCrLf &
                    "Clone this assembly to a new location?"

                Dim result As DialogResult = MessageBox.Show(
                    confirmMsg,
                    "Confirm Source Assembly",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If result <> DialogResult.Yes Then
                    m_logger.Log("User cancelled assembly confirmation")
                    Return False
                End If

                m_logger.Log("SUCCESS: Assembly validated and confirmed")
                Return True

            Catch ex As Exception
                m_logger.LogError("ValidateActiveDocument failed: " & ex.Message)
                MessageBox.Show(
                    "Error validating active document: " & ex.Message,
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Get destination folder from user via folder browser dialog
        ''' VBScript: GetDestinationFolder()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Function GetDestinationFolder() As String
            Try
                ' Get source directory from active document
                Dim sourceDoc As Document = m_invApp.ActiveDocument
                If sourceDoc Is Nothing Then
                    m_logger.Log("ERROR: No active document for destination folder")
                    Return String.Empty
                End If

                Dim sourceDir As String = Path.GetDirectoryName(sourceDoc.FullFileName)
                Dim startFolder As String = Path.GetDirectoryName(sourceDir)

                ' If no parent directory, use the source directory itself
                If String.IsNullOrEmpty(startFolder) Then
                    startFolder = sourceDir
                End If

                ' Create and configure folder browser dialog
                Using dialog As New FolderBrowserDialog()
                    dialog.Description = "Select DESTINATION folder for the cloned assembly:" & vbCrLf &
                                       vbCrLf &
                                       "Source: " & sourceDir & vbCrLf &
                                       vbCrLf &
                                       "TIP: Click 'Make New Folder' to create a new destination"
                    dialog.SelectedPath = startFolder
                    dialog.ShowNewFolderButton = True

                    ' Show dialog
                    Dim result As DialogResult = dialog.ShowDialog()

                    If result <> DialogResult.OK Then
                        m_logger.Log("User cancelled destination folder selection")
                        Return String.Empty
                    End If

                    Dim destPath As String = dialog.SelectedPath

                    ' Validate destination is different from source
                    If String.Equals(destPath, sourceDir, StringComparison.OrdinalIgnoreCase) Then
                        MessageBox.Show(
                            "Destination cannot be the same as source folder!" & vbCrLf & vbCrLf &
                            "Please select a different folder.",
                            "Invalid Destination",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation)
                        m_logger.Log("ERROR: Destination same as source")
                        Return String.Empty
                    End If

                    m_logger.Log("DESTINATION: " & destPath)
                    Return destPath
                End Using

            Catch ex As Exception
                m_logger.LogError("GetDestinationFolder failed: " & ex.Message)
                MessageBox.Show(
                    "Error selecting destination folder: " & ex.Message,
                    "Folder Selection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Get prefix for heritage naming from user
        ''' VBScript: GetPlantSectionNaming()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Function GetPrefixFromUser() As String
            Try
                m_logger.Log("PLANT: Getting plant section naming convention from user")

                Dim prompt As String =
                    "DEFINE PROJECT PREFIX" & vbCrLf & vbCrLf &
                    "Enter the project prefix for heritage naming:" & vbCrLf & vbCrLf &
                    "Examples:" & vbCrLf &
                    "  WALKWAY-3-    (for Walkway 3)" & vbCrLf &
                    "  PLANT1-000-   (for Plant 1)" & vbCrLf &
                    "  SSCR05-001-   (for Section 1)" & vbCrLf & vbCrLf &
                    "This will create part numbers like:" & vbCrLf &
                    "  PREFIX-PL1, PREFIX-CH1, PREFIX-B1, etc."

                ' Get input from user
                Dim plantInput As String = Interaction.InputBox(
                    prompt,
                    "Define Project Prefix",
                    "CLONE-001-",
                    -1, -1)

                ' Handle empty input (user cancelled)
                If String.IsNullOrWhiteSpace(plantInput) Then
                    m_logger.Log("PLANT: User cancelled, using default prefix: CLONE-001-")
                    Return "CLONE-001-"
                End If

                ' Trim and format the prefix
                plantInput = plantInput.Trim()
                plantInput = plantInput.ToUpper()

                ' Ensure prefix ends with "-"
                If Not plantInput.EndsWith("-") Then
                    plantInput = plantInput & "-"
                End If

                m_logger.Log("PLANT: Using prefix: " & plantInput)
                Return plantInput

            Catch ex As Exception
                m_logger.LogError("GetPrefixFromUser failed: " & ex.Message)
                ' Return default on error
                m_logger.Log("PLANT: Error occurred, using default prefix: CLONE-001-")
                Return "CLONE-001-"
            End Try
        End Function

        ' ========================================================================
        ' PHASE 2: FILE COLLECTION METHODS
        ' ========================================================================

        ''' <summary>
        ''' Collect all referenced parts and sub-assemblies recursively
        ''' VBScript: CollectAllReferencedParts()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub CollectReferencedParts(asmDoc As AssemblyDocument, allParts As Dictionary(Of String, String))
            m_logger.Log("COLLECT: Scanning assembly for all referenced parts and sub-assemblies...")

            ' Start recursive collection
            CollectPartsRecursively(asmDoc, allParts, "ROOT")

            m_logger.Log("COLLECT: Found " & allParts.Count.ToString() & " unique files")
        End Sub

        ''' <summary>
        ''' Collect all IDW drawing files from source directory tree
        ''' VBScript: CollectIDWFiles()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub CollectIDWFiles(sourceDir As String, allParts As Dictionary(Of String, String))
            m_logger.Log("COLLECT: Scanning for IDW drawing files...")

            ' Start recursive IDW collection
            CollectIDWFilesRecursive(sourceDir, allParts)
        End Sub

        ''' <summary>
        ''' Recursively traverse assembly hierarchy to collect all parts and sub-assemblies
        ''' VBScript: CollectPartsRecursively()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub CollectPartsRecursively(asmDoc As AssemblyDocument, allParts As Dictionary(Of String, String), level As String)
            Try
                Dim occurrences As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences

                For i As Integer = 1 To occurrences.Count
                    Dim occ As ComponentOccurrence = occurrences.Item(i)

                    ' Skip suppressed occurrences
                    If occ.Suppressed Then
                        m_logger.Log("COLLECT: Skipping suppressed occurrence " & occ.Name & " at " & level)
                        Continue For
                    End If

                    Try
                        ' Get the document referenced by this occurrence
                        Dim doc As Document = occ.Definition.Document

                        If doc Is Nothing Then
                            m_logger.Log("COLLECT: Could not access document for occurrence " & occ.Name & " at " & level)
                            Continue For
                        End If

                        Dim fullPath As String = doc.FullFileName

                        ' Skip files from OldVersions folders
                        If fullPath.ToLower().Contains("\oldversions\") Then
                            m_logger.Log("COLLECT: Skipping OldVersions file: " & Path.GetFileName(fullPath))
                            Continue For
                        End If

                        Dim fileName As String = Path.GetFileName(fullPath)
                        Dim extension As String = Path.GetExtension(fullPath).ToLower()

                        ' Process based on file type
                        If extension = ".ipt" Then
                            ' It's a part file
                            If Not allParts.ContainsKey(fullPath) Then
                                ' Get description for grouping
                                Dim description As String = GetDescriptionFromIProperty(doc)
                                allParts(fullPath) = description
                                m_logger.Log("COLLECT: PART " & fileName & " (" & description & ") at " & level)
                            End If

                        ElseIf extension = ".iam" Then
                            ' It's a sub-assembly - add it and recurse
                            If Not allParts.ContainsKey(fullPath) Then
                                allParts(fullPath) = "SUB-ASSEMBLY"
                                m_logger.Log("COLLECT: SUB-ASSEMBLY " & fileName & " at " & level)
                            End If

                            ' Recurse into sub-assembly
                            Dim subAsmDoc As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                            CollectPartsRecursively(subAsmDoc, allParts, level & ">" & fileName)
                        End If

                    Catch ex As Exception
                        m_logger.LogWarning("COLLECT: Error processing occurrence " & occ.Name & ": " & ex.Message)
                    End Try
                Next

            Catch ex As Exception
                m_logger.LogError("CollectPartsRecursively failed at " & level & ": " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Recursively scan directory for IDW files
        ''' VBScript: CollectIDWFilesRecursive()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub CollectIDWFilesRecursive(folderPath As String, allParts As Dictionary(Of String, String))
            Try
                ' Skip if directory doesn't exist
                If Not Directory.Exists(folderPath) Then
                    Exit Sub
                End If

                ' Check files in current folder
                Dim files As String() = Directory.GetFiles(folderPath, "*.idw")
                For Each file As String In files
                    If Not allParts.ContainsKey(file) Then
                        allParts(file) = "DRAWING"
                        m_logger.Log("COLLECT: DRAWING " & Path.GetFileName(file))
                    End If
                Next

                ' Recurse into subfolders (skip OldVersions)
                Dim subFolders As String() = Directory.GetDirectories(folderPath)
                For Each subFolder As String In subFolders
                    Dim folderName As String = Path.GetFileName(subFolder)

                    ' Skip OldVersions folders completely
                    If folderName.ToLower() <> "oldversions" Then
                        CollectIDWFilesRecursive(subFolder, allParts)
                    Else
                        m_logger.Log("COLLECT: Skipping OldVersions folder: " & subFolder)
                    End If
                Next

            Catch ex As Exception
                m_logger.LogWarning("CollectIDWFilesRecursive error in " & folderPath & ": " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get Description iProperty from document
        ''' VBScript: GetDescriptionFromIProperty()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Function GetDescriptionFromIProperty(doc As Document) As String
            Try
                ' Access Design Tracking Properties
                Dim propSets As PropertySets = doc.PropertySets
                Dim designProps As PropertySet = propSets.Item("Design Tracking Properties")

                If designProps Is Nothing Then
                    Return String.Empty
                End If

                ' Get Description property
                Dim descProp As [Property] = designProps.Item("Description")

                If descProp Is Nothing Then
                    Return String.Empty
                End If

                Dim description As String = Trim(descProp.Value.ToString())

                If description = "N/A" OrElse description = "[N/A]" OrElse description = "" Then
                    Return String.Empty
                End If

                Return description

            Catch ex As Exception
                ' Return empty string on any error
                Return String.Empty
            End Try
        End Function

        ' ========================================================================
        ' PHASE 3: FILE OPERATION METHODS
        ' ========================================================================

        ''' <summary>
        ''' Group parts for heritage renaming by classification
        ''' VBScript: GroupPartsForRenaming()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub GroupPartsForRenaming(allParts As Dictionary(Of String, String))
            m_logger.Log("GROUP: Grouping parts for heritage naming...")

            ' Clear existing groups
            m_componentGroups.Clear()

            ' Process each part
            For Each kvp As KeyValuePair(Of String, String) In allParts
                Dim partPath As String = kvp.Key
                Dim description As String = kvp.Value
                Dim fileName As String = Path.GetFileName(partPath)
                Dim extension As String = Path.GetExtension(fileName).ToLower()

                ' Skip sub-assemblies and drawings - only process parts
                If extension = ".iam" Then
                    m_logger.Log("GROUP: Skipping sub-assembly " & fileName)
                ElseIf extension = ".idw" Then
                    m_logger.Log("GROUP: Skipping drawing " & fileName)
                Else
                    ' It's a part file - classify by description
                    Dim groupCode As String = PartClassifier.ClassifyByDescription(description)

                    If groupCode <> "SKIP" Then
                        ' Add to component groups (dictionary of dictionaries)
                        If Not m_componentGroups.ContainsKey(groupCode) Then
                            m_componentGroups(groupCode) = New Dictionary(Of String, String)()
                        End If

                        Dim groupDict As Dictionary(Of String, String) = m_componentGroups(groupCode)

                        ' Store part info: "path|description|fileName"
                        If Not groupDict.ContainsKey(partPath) Then
                            groupDict(partPath) = partPath & "|" & description & "|" & fileName
                        End If

                        m_logger.Log("GROUP: " & fileName & " -> " & groupCode & " (" & description & ")")
                    Else
                        m_logger.Log("GROUP: Skipping hardware " & fileName & " (" & description & ")")
                    End If
                End If
            Next

            m_logger.Log("GROUP: Created " & m_componentGroups.Count.ToString() & " groups")
        End Sub

        ''' <summary>
        ''' Copy all files to destination with optional renaming
        ''' VBScript: CopyAllFiles()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub CopyAllFiles(allParts As Dictionary(Of String, String), destFolder As String)
            m_logger.Log("COPY: Starting file copy process...")

            ' Group counters for renaming (groupCode -> currentNumber)
            Dim groupCounters As New Dictionary(Of String, Integer)()

            ' Initialize naming schemes
            If m_doRename Then
                InitializeNamingSchemes()
            End If

            ' Initialize counters from Registry if renaming
            If m_doRename Then
                m_logger.Log("REGISTRY: Loading existing counters for prefix: " & m_plantSection)

                For Each groupCode As String In m_componentGroups.Keys
                    Dim prefixGroupKey As String = m_plantSection & groupCode
                    Dim existingCounter As Integer = m_registryManager.GetCounter(m_plantSection, groupCode)

                    Dim startingCounter As Integer
                    If existingCounter > 0 Then
                        startingCounter = existingCounter + 1
                        m_logger.Log("REGISTRY: Group '" & groupCode & "' continuing from number " & startingCounter.ToString() & " (found existing highest = " & existingCounter.ToString() & ")")
                    Else
                        startingCounter = 1
                        m_logger.Log("REGISTRY: Group '" & groupCode & "' starting from number 1 (new prefix or group)")
                    End If

                    groupCounters(groupCode) = startingCounter
                Next
            End If

            ' Copy each file
            For Each kvp As KeyValuePair(Of String, String) In allParts
                Dim originalPath As String = kvp.Key
                Dim originalFileName As String = Path.GetFileName(originalPath)
                Dim newFileName As String
                Dim newPath As String

                ' Determine new name based on file type
                Dim extension As String = Path.GetExtension(originalFileName).ToLower()

                If extension = ".iam" Then
                    ' Sub-assembly - keep original name
                    newFileName = originalFileName
                    m_logger.Log("COPY: SUB-ASSEMBLY " & originalFileName & " (keeping original name)")
                ElseIf extension = ".idw" Then
                    ' Drawing file - keep original name
                    newFileName = originalFileName
                    m_logger.Log("COPY: DRAWING " & originalFileName & " (keeping original name)")
                ElseIf m_doRename Then
                    ' Part file with renaming enabled
                    Dim description As String = kvp.Value
                    Dim groupCode As String = PartClassifier.ClassifyByDescription(description)

                    If groupCode = "SKIP" OrElse groupCode = "OTHER" Then
                        ' Hardware or unclassified - keep original name
                        newFileName = originalFileName
                    Else
                        ' Generate new name using scheme
                        Dim scheme As String = m_namingSchemes(groupCode)
                        Dim counter As Integer = groupCounters(groupCode)

                        newFileName = scheme.Replace("{N}", counter.ToString())

                        ' Ensure .ipt extension
                        If Path.GetExtension(newFileName).ToLower() <> ".ipt" Then
                            newFileName = newFileName & ".ipt"
                        End If

                        ' Increment counter
                        groupCounters(groupCode) = counter + 1

                        m_logger.Log("COPY: PART " & originalFileName & " -> " & newFileName & " (" & groupCode & ")")
                    End If
                Else
                    ' Part file without renaming - keep original name
                    newFileName = originalFileName
                End If

                ' Compute destination path (preserve folder structure)
                newPath = GetDestinationPath(originalPath, destFolder, newFileName)

                ' Create target directory if needed
                Dim targetDir As String = Path.GetDirectoryName(newPath)
                If Not Directory.Exists(targetDir) Then
                    Directory.CreateDirectory(targetDir)
                    m_logger.Log("FOLDER: Created " & FileHelper.GetRelativePath(destFolder, targetDir))
                End If

                ' Copy the file
                Try
                    File.Copy(originalPath, newPath, True)
                    m_logger.Log("COPIED: " & originalFileName & " -> " & newFileName)
                    m_copiedFiles(originalPath) = newPath
                Catch ex As Exception
                    m_logger.LogError("Could not copy " & originalFileName & ": " & ex.Message)
                End Try
            Next

            ' Save final counters to Registry if renaming was enabled
            If m_doRename AndAlso groupCounters.Count > 0 Then
                m_logger.Log("REGISTRY: Saving final counters to Registry")

                For Each kvp2 As KeyValuePair(Of String, Integer) In groupCounters
                    Dim groupCode As String = kvp2.Key
                    Dim finalCounter As Integer = kvp2.Value - 1  ' Last used number (counter was incremented after use)

                    m_registryManager.SetCounter(m_plantSection, groupCode, finalCounter)
                    m_logger.Log("REGISTRY: Saved " & m_plantSection & groupCode & " = " & finalCounter.ToString())
                Next

                m_logger.Log("REGISTRY: All counters saved successfully")
            End If

            m_logger.Log("COPY: File copy process completed. Total files: " & m_copiedFiles.Count.ToString())
        End Sub

        ''' <summary>
        ''' Get destination path preserving folder structure
        ''' Helper for CopyAllFiles
        ''' </summary>
        Private Function GetDestinationPath(originalPath As String, destFolder As String, newFileName As String) As String
            Dim originalDir As String = Path.GetDirectoryName(originalPath)

            ' Check if file is in a subfolder of source root
            If originalDir.Length > m_sourceRoot.Length AndAlso
               originalDir.StartsWith(m_sourceRoot, StringComparison.OrdinalIgnoreCase) Then

                ' File is in a subfolder - preserve the subfolder structure
                Dim relativePath As String = originalDir.Substring(m_sourceRoot.Length + 1) ' Skip trailing backslash
                Return Path.Combine(destFolder, relativePath, newFileName)
            Else
                ' File is at source root level or from different location
                Return Path.Combine(destFolder, newFileName)
            End If
        End Function

        ''' <summary>
        ''' Initialize naming schemes for each group
        ''' VBScript: GetUserNamingSchemes()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub InitializeNamingSchemes()
            m_namingSchemes.Clear()

            For Each groupCode As String In m_componentGroups.Keys
                ' Generate scheme: PREFIX + GROUP + {N}
                Dim scheme As String = m_plantSection & groupCode & "{N}"
                m_namingSchemes(groupCode) = scheme
                m_logger.Log("SCHEME: " & groupCode & " -> " & scheme)
            Next
        End Sub

        ' ========================================================================
        ' PHASE 4: REFERENCE UPDATE METHODS
        ' ========================================================================

        ''' <summary>
        ''' Update assembly references to use local copies
        ''' VBScript: UpdateInMemoryAssemblyReferences() + assembly-by-assembly processing
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub UpdateAssemblyReferences(asmPath As String)
            Try
                m_logger.Log("ASM UPDATE: Starting reference update process...")
                m_logger.Log("ASM UPDATE: Target assembly: " & Path.GetFileName(asmPath))

                ' Store original settings
                Dim originalSilent As Boolean = m_invApp.SilentOperation

                ' Enable silent mode
                m_invApp.SilentOperation = True
                m_logger.Log("ASM UPDATE: Enabled SilentOperation mode")

                ' Build lookup dictionaries for faster matching
                Dim fileNameLookup As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                Dim pathLookup As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    Dim origFileName As String = Path.GetFileName(kvp.Key)
                    If Not fileNameLookup.ContainsKey(origFileName) Then
                        fileNameLookup(origFileName) = kvp.Value
                    End If
                    pathLookup(kvp.Key) = kvp.Value
                Next

                m_logger.Log("ASM UPDATE: Built lookup with " & fileNameLookup.Count.ToString() & " filenames, " & pathLookup.Count.ToString() & " paths")

                ' ====================================================================
                ' STEP 1: Preload all copied PARTS into memory first
                ' Parts have no references so they open cleanly without dialogs
                ' ====================================================================
                m_logger.Log("ASM UPDATE: Step 1 - Preloading all copied PARTS into memory...")

                Dim partsOpened As Integer = 0
                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    Dim newPath As String = kvp.Value

                    If Path.GetExtension(newPath).ToLower() = ".ipt" Then
                        Try
                            Dim partDoc As Document = m_invApp.Documents.Open(newPath, False)
                            partsOpened += 1
                            m_logger.Log("ASM UPDATE: Preloaded PART " & Path.GetFileName(newPath))
                        Catch ex As Exception
                            m_logger.LogWarning("Could not preload part " & Path.GetFileName(newPath) & ": " & ex.Message)
                        End Try
                    End If
                Next

                m_logger.Log("ASM UPDATE: Preloaded " & partsOpened.ToString() & " parts into memory")

                ' ====================================================================
                ' STEP 2: Open all copied SUB-ASSEMBLIES (except main)
                ' ====================================================================
                m_logger.Log("ASM UPDATE: Step 2 - Opening sub-assemblies...")

                Dim assembliesOpened As Integer = 0
                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    Dim newPath As String = kvp.Value

                    If Path.GetExtension(newPath).ToLower() = ".iam" AndAlso
                       Not String.Equals(newPath, asmPath, StringComparison.OrdinalIgnoreCase) Then
                        Try
                            Dim subAsmDoc As AssemblyDocument = DirectCast(m_invApp.Documents.Open(newPath, False), AssemblyDocument)
                            assembliesOpened += 1
                            m_logger.Log("ASM UPDATE: Opened SUB-ASSEMBLY " & Path.GetFileName(newPath))
                        Catch ex As Exception
                            m_logger.LogWarning("Could not open sub-assembly " & Path.GetFileName(newPath) & ": " & ex.Message)
                        End Try
                    End If
                Next

                m_logger.Log("ASM UPDATE: Opened " & assembliesOpened.ToString() & " sub-assemblies")

                ' ====================================================================
                ' STEP 3: Open main assembly and update ALL assemblies
                ' ====================================================================
                m_logger.Log("ASM UPDATE: Step 3 - Opening main assembly...")

                Dim mainAsmDoc As AssemblyDocument = Nothing
                Try
                    mainAsmDoc = DirectCast(m_invApp.Documents.Open(asmPath, False), AssemblyDocument)
                    m_logger.Log("ASM UPDATE: Opened main assembly: " & Path.GetFileName(asmPath))
                Catch ex As Exception
                    m_logger.LogError("Could not open main assembly: " & ex.Message)
                    m_invApp.SilentOperation = originalSilent
                    Return
                End Try

                ' Update all assemblies including the main one
                m_logger.Log("ASM UPDATE: Step 4 - Updating references in ALL assemblies...")

                Dim totalUpdated As Integer = 0
                Dim assembliesProcessed As Integer = 0

                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    Dim asmPathToUpdate As String = kvp.Value

                    If Path.GetExtension(asmPathToUpdate).ToLower() = ".iam" Then
                        Dim updatedInThisAssembly As Integer = UpdateReferencesInAssembly(asmPathToUpdate, fileNameLookup, pathLookup)
                        totalUpdated += updatedInThisAssembly
                        assembliesProcessed += 1
                        m_logger.Log("ASM UPDATE: [" & assembliesProcessed.ToString() & "] Updated " & updatedInThisAssembly.ToString() & " refs in " & Path.GetFileName(asmPathToUpdate))
                    End If
                Next

                ' ====================================================================
                ' STEP 5: Save all modified documents
                ' ====================================================================
                m_logger.Log("ASM UPDATE: Step 5 - Saving all modified documents...")

                Dim savedCount As Integer = 0
                For Each doc As Document In m_invApp.Documents
                    If doc.Dirty Then
                        doc.Save()
                        savedCount += 1
                        m_logger.Log("ASM UPDATE: Saved " & Path.GetFileName(doc.FullFileName))
                    End If
                Next

                ' ====================================================================
                ' COMPLETE
                ' ====================================================================
                m_logger.Log("ASM UPDATE: Complete - Updated " & totalUpdated.ToString() & " references in " & assembliesProcessed.ToString() & " assemblies")
                m_logger.Log("ASM UPDATE: Saved " & savedCount.ToString() & " documents")

                ' Restore original settings
                m_invApp.SilentOperation = originalSilent

            Catch ex As Exception
                m_logger.LogError("UpdateAssemblyReferences failed: " & ex.Message)
                ' Try to restore settings
                Try
                    m_invApp.SilentOperation = True
                Catch
                End Try
            End Try
        End Sub

        ''' <summary>
        ''' Update references in a single assembly document
        ''' Helper for UpdateAssemblyReferences
        ''' </summary>
        Private Function UpdateReferencesInAssembly(asmPath As String, fileNameLookup As Dictionary(Of String, String), pathLookup As Dictionary(Of String, String)) As Integer
            Try
                ' Find the assembly in the open documents collection
                Dim asmDoc As AssemblyDocument = Nothing
                For Each doc As Document In m_invApp.Documents
                    If String.Equals(doc.FullFileName, asmPath, StringComparison.OrdinalIgnoreCase) Then
                        asmDoc = DirectCast(doc, AssemblyDocument)
                        Exit For
                    End If
                Next

                If asmDoc Is Nothing Then
                    m_logger.Log("ASM UPDATE: Assembly not found in memory: " & Path.GetFileName(asmPath))
                    Return 0
                End If

                ' Get referenced file descriptors
                Dim refDescs As FileDescriptors = asmDoc.File.ReferencedFileDescriptors

                Dim updatedCount As Integer = 0
                For Each fd As FileDescriptor In refDescs
                    Try
                        Dim refPath As String = fd.FullFileName
                        Dim refFileName As String = Path.GetFileName(refPath)
                        Dim newRefPath As String = String.Empty

                        ' Try exact path match first
                        If pathLookup.ContainsKey(refPath) Then
                            newRefPath = pathLookup(refPath)
                            m_logger.Log("ASM UPDATE: Found by EXACT PATH for " & refFileName)
                        ElseIf fileNameLookup.ContainsKey(refFileName) Then
                            newRefPath = fileNameLookup(refFileName)
                            m_logger.Log("ASM UPDATE: Found by FILENAME for " & refFileName)
                        End If

                        If Not String.IsNullOrEmpty(newRefPath) Then
                            ' Only replace if paths are different
                            If Not String.Equals(refPath, newRefPath, StringComparison.OrdinalIgnoreCase) Then
                                m_logger.Log("ASM UPDATE: Replacing " & refFileName & " -> " & Path.GetFileName(newRefPath))
                                fd.ReplaceReference(newRefPath)
                                updatedCount += 1
                            Else
                                m_logger.Log("ASM UPDATE: SKIP (already correct path) for " & refFileName)
                            End If
                        Else
                            ' Only log if it's a part/assembly we care about
                            Dim ext As String = Path.GetExtension(refFileName).ToLower()
                            If ext = ".ipt" OrElse ext = ".iam" Then
                                m_logger.Log("ASM UPDATE: No mapping for " & refFileName)
                            End If
                        End If

                    Catch ex As Exception
                        m_logger.LogWarning("Error processing reference: " & ex.Message)
                    End Try
                Next

                Return updatedCount

            Catch ex As Exception
                m_logger.LogError("UpdateReferencesInAssembly failed: " & ex.Message)
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Update IDW drawing references to new assemblies/parts
        ''' VBScript: UpdateIDWReferences()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub UpdateIDWReferences(destFolder As String)
            Try
                m_logger.Log("IDW UPDATE: Starting IDW reference update process...")
                m_logger.Log("IDW UPDATE: Target folder: " & destFolder)

                ' Store original settings
                Dim originalSilent As Boolean = m_invApp.SilentOperation
                Dim originalResolve As FileResolutionEnum = m_invApp.FileOptions.ResolveFileOption

                ' Enable silent mode
                m_invApp.SilentOperation = True
                m_invApp.FileOptions.ResolveFileOption = FileResolutionEnum.kSkipUnresolvedFiles
                m_logger.Log("IDW UPDATE: Enabled SilentOperation mode")

                ' Build lookup dictionary
                Dim fileNameLookup As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    Dim origFileName As String = Path.GetFileName(kvp.Key)
                    If Not fileNameLookup.ContainsKey(origFileName) Then
                        fileNameLookup(origFileName) = kvp.Value
                    End If
                Next

                m_logger.Log("IDW UPDATE: Built lookup with " & fileNameLookup.Count.ToString() & " entries")

                ' Find all IDW files in destination folder (recursive)
                Dim idwFiles As List(Of String) = New List(Of String)()
                FindIDWFilesRecursive(destFolder, idwFiles)

                m_logger.Log("IDW UPDATE: Found " & idwFiles.Count.ToString() & " IDW files to process")

                Dim totalUpdated As Integer = 0

                ' Process each IDW file
                For Each idwPath As String In idwFiles
                    Try
                        m_logger.Log("IDW UPDATE: Processing " & Path.GetFileName(idwPath))

                        ' Open IDW document
                        Dim idwDoc As DrawingDocument = DirectCast(m_invApp.Documents.Open(idwPath, False), DrawingDocument)

                        ' Get referenced file descriptors
                        Dim refDescs As FileDescriptors = idwDoc.File.ReferencedFileDescriptors
                        Dim updatedInThisFile As Integer = 0

                        For Each fd As FileDescriptor In refDescs
                            Try
                                Dim refPath As String = fd.FullFileName
                                Dim refFileName As String = Path.GetFileName(refPath)
                                Dim newRefPath As String = String.Empty

                                ' Try exact path match first
                                If m_copiedFiles.ContainsKey(refPath) Then
                                    newRefPath = m_copiedFiles(refPath)
                                    m_logger.Log("IDW UPDATE: Found by EXACT PATH for " & refFileName)
                                ElseIf fileNameLookup.ContainsKey(refFileName) Then
                                    newRefPath = fileNameLookup(refFileName)
                                    m_logger.Log("IDW UPDATE: Found by FILENAME for " & refFileName)
                                End If

                                If Not String.IsNullOrEmpty(newRefPath) Then
                                    m_logger.Log("IDW UPDATE: Replacing " & refFileName & " -> " & Path.GetFileName(newRefPath))
                                    fd.ReplaceReference(newRefPath)

                                    If Err.Number = 0 Then
                                        updatedInThisFile += 1
                                    Else
                                        m_logger.LogWarning("ReplaceReference failed: " & Err.Description)
                                    End If
                                    Err.Clear()
                                Else
                                    m_logger.Log("IDW UPDATE: No mapping for " & refFileName)
                                End If

                            Catch ex As Exception
                                m_logger.LogWarning("Error processing IDW reference: " & ex.Message)
                            End Try
                        Next

                        ' Save IDW document
                        idwDoc.Save()
                        totalUpdated += updatedInThisFile
                        m_logger.Log("IDW UPDATE: Saved " & Path.GetFileName(idwPath) & " (" & updatedInThisFile.ToString() & " refs updated)")

                        idwDoc.Close()

                    Catch ex As Exception
                        m_logger.LogError("Could not process IDW " & Path.GetFileName(idwPath) & ": " & ex.Message)
                    End Try
                Next

                m_logger.Log("IDW UPDATE: Complete - Updated " & totalUpdated.ToString() & " references in " & idwFiles.Count.ToString() & " IDW files")

                ' Restore original settings
                m_invApp.SilentOperation = originalSilent
                m_invApp.FileOptions.ResolveFileOption = originalResolve

            Catch ex As Exception
                m_logger.LogError("UpdateIDWReferences failed: " & ex.Message)
                ' Try to restore settings
                Try
                    m_invApp.SilentOperation = True
                Catch
                End Try
            End Try
        End Sub

        ''' <summary>
        ''' Recursively find all IDW files in a folder
        ''' Helper for UpdateIDWReferences
        ''' </summary>
        Private Sub FindIDWFilesRecursive(folderPath As String, idwFiles As List(Of String))
            Try
                ' Add IDW files in current folder
                Dim files As String() = Directory.GetFiles(folderPath, "*.idw")
                idwFiles.AddRange(files)

                ' Recurse into subfolders
                Dim subFolders As String() = Directory.GetDirectories(folderPath)
                For Each subFolder As String In subFolders
                    ' Skip OldVersions folders
                    If Not Path.GetFileName(subFolder).ToLower().Equals("oldversions") Then
                        FindIDWFilesRecursive(subFolder, idwFiles)
                    End If
                Next

            Catch ex As Exception
                m_logger.LogWarning("FindIDWFilesRecursive error: " & ex.Message)
            End Try
        End Sub

        ' ========================================================================
        ' PHASE 5: REGISTRY & MAPPING METHODS
        ' ========================================================================

        ''' <summary>
        ''' Write STEP_1_MAPPING.txt file for reference tracking
        ''' VBScript: WriteMappingFile()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub WriteMappingFile(destFolder As String)
            Try
                m_logger.Log("MAPPING: Writing mapping files...")

                ' File 1: Filename-based mapping (for IDW updater filename matching)
                Dim mappingPath As String = Path.Combine(destFolder, "STEP_1_MAPPING.txt")

                Using writer As New StreamWriter(mappingPath, False, System.Text.Encoding.UTF8)
                    writer.WriteLine("# Filename-based mapping file for cloned folder: " & destFolder)
                    writer.WriteLine("# Generated by Assembly Cloner on " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    writer.WriteLine("# Format: OLD_FILENAME|NEW_FILENAME")
                    writer.WriteLine("# For filename-only matching")
                    writer.WriteLine()

                    For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                        Dim oldFileName As String = Path.GetFileName(kvp.Key)
                        Dim newFileName As String = Path.GetFileName(kvp.Value)
                        writer.WriteLine(oldFileName.ToLower() & "|" & newFileName)
                    Next
                End Using

                m_logger.Log("MAPPING: Wrote " & m_copiedFiles.Count.ToString() & " entries to STEP_1_MAPPING.txt")

                ' File 2: Full path mapping (for Assembly Renamer compatibility)
                Dim fullPathMappingPath As String = Path.Combine(destFolder, "STEP_1_MAPPING_FULLPATH.txt")

                Using fullPathWriter As New StreamWriter(fullPathMappingPath, False, System.Text.Encoding.UTF8)
                    fullPathWriter.WriteLine("# Full path mapping file for cloned folder: " & destFolder)
                    fullPathWriter.WriteLine("# Generated by Assembly Cloner on " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    fullPathWriter.WriteLine("# Format: OLD_FULLPATH|NEW_FULLPATH")
                    fullPathWriter.WriteLine("# For exact path matching")
                    fullPathWriter.WriteLine()

                    For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                        fullPathWriter.WriteLine(kvp.Key & "|" & kvp.Value)
                    Next
                End Using

                m_logger.Log("MAPPING: Wrote " & m_copiedFiles.Count.ToString() & " entries to STEP_1_MAPPING_FULLPATH.txt")

            Catch ex As Exception
                m_logger.LogError("WriteMappingFile failed: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Validate clone and generate final report
        ''' VBScript: ValidateCloneAndLog()
        ''' STATUS: ✅ IMPLEMENTED
        ''' </summary>
        Private Sub ValidateClone(sourceDir As String, destFolder As String)
            Try
                m_logger.Log("VALIDATE: Validating clone and generating final report...")
                m_logger.Log("VALIDATE: Source: " & sourceDir)
                m_logger.Log("VALIDATE: Destination: " & destFolder)

                ' Build file inventories
                Dim sourceInventory As Dictionary(Of String, Integer) = FileHelper.BuildFileInventory(sourceDir)
                Dim destInventory As Dictionary(Of String, Integer) = FileHelper.BuildFileInventory(destFolder)

                ' Count files by type
                m_logger.Log("VALIDATE: Source file count: " & sourceInventory.Values.Sum().ToString() & " files")
                m_logger.Log("VALIDATE: Destination file count: " & destInventory.Values.Sum().ToString() & " files")

                ' Count by type
                For Each kvp As KeyValuePair(Of String, Integer) In sourceInventory
                    m_logger.Log("VALIDATE: Source - " & kvp.Key & ": " & kvp.Value.ToString())
                Next

                For Each kvp As KeyValuePair(Of String, Integer) In destInventory
                    m_logger.Log("VALIDATE: Dest - " & kvp.Key & ": " & kvp.Value.ToString())
                Next

                ' Validate that all copied files exist in destination
                Dim missingFiles As Integer = 0
                For Each kvp As KeyValuePair(Of String, String) In m_copiedFiles
                    If Not File.Exists(kvp.Value) Then
                        m_logger.LogError("VALIDATE: MISSING FILE in destination: " & kvp.Value)
                        missingFiles += 1
                    End If
                Next

                If missingFiles = 0 Then
                    m_logger.Log("VALIDATE: SUCCESS - All copied files verified in destination")
                Else
                    m_logger.LogError("VALIDATE: FAILED - " & missingFiles.ToString() & " files missing from destination")
                End If

                ' Generate summary
                m_logger.Log("========================================")
                m_logger.Log("CLONE VALIDATION SUMMARY")
                m_logger.Log("========================================")
                m_logger.Log("Total files copied: " & m_copiedFiles.Count.ToString())
                m_logger.Log("Source files: " & sourceInventory.Values.Sum().ToString())
                m_logger.Log("Destination files: " & destInventory.Values.Sum().ToString())
                m_logger.Log("Missing files: " & missingFiles.ToString())
                m_logger.Log("Validation status: " & If(missingFiles = 0, "PASSED", "FAILED"))
                m_logger.Log("========================================")

            Catch ex As Exception
                m_logger.LogError("ValidateClone failed: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace
