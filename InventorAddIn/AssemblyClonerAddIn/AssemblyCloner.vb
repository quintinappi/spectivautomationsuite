' ==============================================================================
' ASSEMBLY CLONER - Clone Assembly with Parts and Patch iLogic
' ==============================================================================
' Author: Quintin de Bruin © 2025
'
' CRITICAL: Uses EXACT methodology from Assembly_Cloner.vbs (Option 9):
'   1. fso.CopyFile - Simple file copy for parts (NOT SaveAs!)
'   2. ApprenticeServer or fd.ReplaceReference - Update assembly references
'   3. fd.ReplaceReference - Update IDW references
'   4. iLogicPatcher - Patch iLogic rules with new part names
'
' ==============================================================================

Imports Inventor
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Namespace AssemblyClonerAddIn

    ''' <summary>
    ''' Main assembly cloning logic - Uses EXACT methodology from VBScript Option 9
    ''' </summary>
    Public Class AssemblyCloner

        Private m_InventorApp As Inventor.Application
        Private m_Patcher As iLogicPatcher

        ' Tracking dictionaries
        Private m_CopiedFiles As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' originalPath -> newPath
        Private m_OccurrenceRenames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' oldOccName -> newOccName (for iLogic patching)
        Private m_FileRenames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)  ' oldFileName (no ext) -> newFileName (no ext)
        Private m_DoRename As Boolean = False
        Private m_PlantSection As String = ""
        Private m_iLogicPatchLog As New System.Text.StringBuilder()  ' Log for iLogic patching

        ' Logging
        Private m_LogFile As StreamWriter
        Private m_LogPath As String

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(ByVal inventorApp As Inventor.Application)
            m_InventorApp = inventorApp
            m_Patcher = New iLogicPatcher(inventorApp)
        End Sub

        ''' <summary>
        ''' Main entry point - clone the active assembly
        ''' </summary>
        Public Sub CloneAssembly()
            Try
                ' Initialize logging
                InitializeLog()
                LogMessage("=== ASSEMBLY CLONER STARTED ===")
                LogMessage("Using EXACT methodology from VBScript Option 9")

                ' Reset tracking
                m_CopiedFiles.Clear()
                m_OccurrenceRenames.Clear()
                m_FileRenames.Clear()
                m_iLogicPatchLog.Clear()

                ' Get source assembly
                Dim sourceDoc As AssemblyDocument = CType(m_InventorApp.ActiveDocument, AssemblyDocument)
                Dim sourceDir As String = System.IO.Path.GetDirectoryName(sourceDoc.FullFileName)
                Dim sourceFileName As String = System.IO.Path.GetFileName(sourceDoc.FullFileName)

                LogMessage("Source assembly: " & sourceDoc.FullFileName)
                LogMessage("Source directory: " & sourceDir)

                ' Confirm with user
                Dim occCount As Integer = sourceDoc.ComponentDefinition.Occurrences.Count
                Dim confirmResult As MsgBoxResult = MsgBox(
                    "ASSEMBLY CLONER" & vbCrLf & vbCrLf &
                    "Source: " & sourceDoc.DisplayName & vbCrLf &
                    "Parts: " & occCount & " occurrences" & vbCrLf &
                    "Location: " & sourceDir & vbCrLf & vbCrLf &
                    "This will:" & vbCrLf &
                    "• Copy assembly + all parts to new folder" & vbCrLf &
                    "• Update all references to local copies" & vbCrLf &
                    "• Patch iLogic rules with new part names" & vbCrLf &
                    "• Copy and update IDW drawings" & vbCrLf & vbCrLf &
                    "Continue?",
                    MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                    "Clone Assembly")

                If confirmResult = MsgBoxResult.No Then
                    LogMessage("User cancelled at confirmation")
                    CloseLog()
                    Return
                End If

                LogMessage("User confirmed clone operation")

                ' Get destination folder
                LogMessage("Opening folder browser dialog...")
                Dim destFolder As String = GetDestinationFolder(sourceDir)
                If String.IsNullOrEmpty(destFolder) Then
                    LogMessage("User cancelled folder selection")
                    CloseLog()
                    Return
                End If

                LogMessage("Destination folder selected: " & destFolder)

                ' Ask about heritage renaming
                Dim renameResult As MsgBoxResult = MsgBox(
                    "HERITAGE RENAMING" & vbCrLf & vbCrLf &
                    "Do you want to rename parts with heritage naming?" & vbCrLf & vbCrLf &
                    "YES = Rename parts (e.g., PLANT-001-PL1, PLANT-001-CH1)" & vbCrLf &
                    "NO = Keep original part names" & vbCrLf & vbCrLf &
                    "Either way, parts will be copied to new folder.",
                    MsgBoxStyle.YesNo + MsgBoxStyle.Question,
                    "Rename Parts?")

                m_DoRename = (renameResult = MsgBoxResult.Yes)
                LogMessage("Rename parts: " & m_DoRename.ToString())

                If m_DoRename Then
                    LogMessage("Prompting for plant section prefix...")
                    m_PlantSection = InputBox(
                        "Enter the plant section prefix:" & vbCrLf & vbCrLf &
                        "Example: DMS-SSCR05-STAIR" & vbCrLf & vbCrLf &
                        "This will be used for heritage naming.",
                        "Plant Section Prefix",
                        System.IO.Path.GetFileNameWithoutExtension(destFolder))

                    If String.IsNullOrEmpty(m_PlantSection) Then
                        LogMessage("User cancelled at prefix input")
                        CloseLog()
                        Return
                    End If

                    LogMessage("Plant section prefix: " & m_PlantSection)
                End If

                ' Ask for new assembly name
                Dim destFolderName As String = System.IO.Path.GetFileName(destFolder)
                Dim newAsmName As String = InputBox(
                    "Enter the new assembly name:" & vbCrLf & vbCrLf &
                    "This will be the filename for the cloned assembly." & vbCrLf &
                    "Do NOT include .iam extension.",
                    "New Assembly Name",
                    destFolderName)

                If String.IsNullOrEmpty(newAsmName) Then
                    LogMessage("User cancelled at assembly name input")
                    CloseLog()
                    Return
                End If

                LogMessage("New assembly name: " & newAsmName)

                ' Get new assembly filename
                Dim newAsmFileName As String = newAsmName & ".iam"
                Dim newAsmPath As String = System.IO.Path.Combine(destFolder, newAsmFileName)
                LogMessage("New assembly path: " & newAsmPath)

                ' Collect all referenced parts AND their occurrence names
                Dim allParts As New Dictionary(Of String, String)  ' filePath -> description
                Dim occurrenceNames As New Dictionary(Of String, String)  ' filePath -> occurrenceName (without :1)
                Dim allSubAssemblies As New Dictionary(Of String, String)  ' filePath -> fileName
                CollectAllReferencedParts(sourceDoc, allParts, occurrenceNames, allSubAssemblies)

                LogMessage("Collected " & allParts.Count & " parts")

                ' CRITICAL: Close source document before copying (same as VBScript)
                Dim sourceFullPath As String = sourceDoc.FullFileName
                LogMessage("Closing source assembly for safe copy...")
                sourceDoc.Close()

                ' Copy assembly file using File.Copy (same as VBScript fso.CopyFile)
                LogMessage("Copying assembly file...")
                System.IO.File.Copy(sourceFullPath, newAsmPath, True)
                m_CopiedFiles.Add(sourceFullPath, newAsmPath)
                LogMessage("COPIED: " & sourceFileName & " -> " & newAsmFileName)

                ' Copy all referenced sub-assemblies using same filenames
                LogMessage("Copying " & allSubAssemblies.Count & " sub-assemblies...")
                CopyAllSubAssemblies(allSubAssemblies, destFolder)

                ' Copy all parts using File.Copy (same as VBScript fso.CopyFile)
                LogMessage("Copying " & allParts.Count & " parts...")
                CopyAllParts(allParts, occurrenceNames, destFolder)

                ' Relink copied sub-assemblies to copied files before patching top-level assembly
                LogMessage("Relinking copied sub-assembly references...")
                UpdateCopiedSubAssemblyReferences(newAsmPath)

                ' CRITICAL FIX: Update references AND patch iLogic in SAME session
                ' Don't close and re-open assembly - that causes Inventor to re-resolve references!
                LogMessage("Updating assembly references and patching iLogic...")
                Dim patchedRulesCount As Integer = UpdateAssemblyAndPatchILogic(newAsmPath)

                ' Process IDW files using fd.ReplaceReference (same as VBScript)
                LogMessage("Processing IDW files...")
                ProcessIDWFilesWithReferenceUpdate(sourceDir, destFolder, sourceFullPath, newAsmPath)

                ' Update iProperties for all copied documents
                LogMessage("Updating iProperties for copied documents...")
                UpdateIPropertiesForCopiedDocuments()

                ' Save iLogic patch log
                LogMessage("Saving iLogic patch log...")
                SaveiLogicPatchLog(destFolder)

                ' Done!
                Dim iLogicMsg As String = ""
                If patchedRulesCount > 0 Then
                    iLogicMsg = "✓ " & patchedRulesCount & " iLogic rule(s) patched with " & m_OccurrenceRenames.Count & " mappings"
                ElseIf m_OccurrenceRenames.Count > 0 Then
                    iLogicMsg = "✓ iLogic checked (no rules needed patching)"
                Else
                    iLogicMsg = "• No part renames - iLogic unchanged"
                End If

                LogMessage("=== ASSEMBLY CLONE COMPLETED SUCCESSFULLY ===")
                LogMessage("Log saved to: " & m_LogPath)
                CloseLog()

                MsgBox(
                    "ASSEMBLY CLONE COMPLETED!" & vbCrLf & vbCrLf &
                    "✓ Assembly copied to: " & destFolder & vbCrLf &
                    "✓ " & m_CopiedFiles.Count & " files copied" & vbCrLf &
                    "✓ References updated" & vbCrLf &
                    iLogicMsg & vbCrLf & vbCrLf &
                    "The new assembly is now completely isolated!" & vbCrLf &
                    "Log saved to Logs folder.",
                    MsgBoxStyle.Information,
                    "Success!")

            Catch ex As Exception
                Dim errorMsg As String = "ERROR: " & ex.Message & vbCrLf & "Stack: " & ex.StackTrace
                LogMessage(errorMsg)
                CloseLog()

                MsgBox("ASSEMBLY CLONE FAILED!" & vbCrLf & vbCrLf &
                       "Error: " & ex.Message & vbCrLf & vbCrLf &
                       "Check log file for details:" & vbCrLf & m_LogPath,
                       MsgBoxStyle.Critical, "Clone Failed")
            End Try
        End Sub

        ''' <summary>
        ''' Update references inside copied sub-assemblies so nested links resolve to copied files.
        ''' </summary>
        Private Sub UpdateCopiedSubAssemblyReferences(ByVal topAssemblyPath As String)
            For Each kvp As KeyValuePair(Of String, String) In m_CopiedFiles
                Dim newPath As String = kvp.Value
                If Not newPath.ToLower().EndsWith(".iam") Then Continue For
                If String.Equals(newPath, topAssemblyPath, StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    Dim asmDoc As AssemblyDocument = CType(m_InventorApp.Documents.Open(newPath, False), AssemblyDocument)
                    Dim fileDescriptors As Object = asmDoc.File.ReferencedFileDescriptors
                    Dim updatedCount As Integer = 0

                    For i As Integer = 1 To fileDescriptors.Count
                        Dim fd As Object = fileDescriptors.Item(i)
                        Dim refPath As String = fd.FullFileName

                        If m_CopiedFiles.ContainsKey(refPath) Then
                            Dim newRefPath As String = m_CopiedFiles(refPath)
                            If Not String.Equals(refPath, newRefPath, StringComparison.OrdinalIgnoreCase) Then
                                fd.ReplaceReference(newRefPath)
                                updatedCount += 1
                            End If
                        End If
                    Next

                    asmDoc.Save()
                    asmDoc.Close(False)
                    LogMessage("SUBASM RELINK: " & System.IO.Path.GetFileName(newPath) & " updated " & updatedCount & " refs")
                Catch ex As Exception
                    LogMessage("SUBASM RELINK ERROR: " & System.IO.Path.GetFileName(newPath) & " - " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Get destination folder from user using folder browser dialog
        ''' </summary>
        Private Function GetDestinationFolder(ByVal sourceDir As String) As String
            Dim parentDir As String = System.IO.Path.GetDirectoryName(sourceDir)

            ' Use Windows folder browser dialog
            Using folderDialog As New System.Windows.Forms.FolderBrowserDialog()
                folderDialog.Description = "Select destination folder for cloned assembly"
                folderDialog.SelectedPath = parentDir
                folderDialog.ShowNewFolderButton = True

                If folderDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    Return folderDialog.SelectedPath
                Else
                    Return Nothing
                End If
            End Using
        End Function

        ''' <summary>
        ''' Collect all referenced parts from assembly (recursive)
        ''' Also captures occurrence names for iLogic patching
        ''' </summary>
        Private Sub CollectAllReferencedParts(ByVal asmDoc As AssemblyDocument, ByVal partsDict As Dictionary(Of String, String), ByVal occNames As Dictionary(Of String, String), ByVal subAsmDict As Dictionary(Of String, String))
            CollectPartsRecursively(asmDoc, partsDict, occNames, subAsmDict, "ROOT")
        End Sub

        ''' <summary>
        ''' Recursively collect parts from assembly
        ''' </summary>
        Private Sub CollectPartsRecursively(ByVal asmDoc As AssemblyDocument, ByVal partsDict As Dictionary(Of String, String), ByVal occNames As Dictionary(Of String, String), ByVal subAsmDict As Dictionary(Of String, String), ByVal level As String)
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                Try
                    If Not occ.Suppressed Then
                        Dim refDoc As Document = occ.Definition.Document
                        If refDoc IsNot Nothing Then
                            Dim fullPath As String = refDoc.FullFileName
                            Dim fileName As String = System.IO.Path.GetFileName(fullPath)

                            If fileName.ToLower().EndsWith(".ipt") Then
                                ' It's a part file
                                If Not partsDict.ContainsKey(fullPath) Then
                                    ' Get description from iProperty
                                    Dim description As String = GetDescriptionFromIProperty(refDoc)
                                    partsDict.Add(fullPath, description)

                                    ' Capture occurrence name (strip :1)
                                    Dim occName As String = occ.Name
                                    Dim colonPos As Integer = occName.LastIndexOf(":"c)
                                    If colonPos > 0 Then
                                        occName = occName.Substring(0, colonPos)
                                    End If
                                    If Not occNames.ContainsKey(fullPath) Then
                                        occNames.Add(fullPath, occName)
                                    End If

                                    LogMessage("COLLECT: " & fileName & " (" & description & ")")
                                End If

                            ElseIf fileName.ToLower().EndsWith(".iam") Then
                                ' Sub-assembly - recurse (skip bolted connections)
                                If Not fileName.ToLower().Contains("bolted connection") Then
                                    If Not subAsmDict.ContainsKey(fullPath) Then
                                        subAsmDict.Add(fullPath, fileName)
                                        LogMessage("COLLECT SUBASM: " & fileName)
                                    End If
                                    Dim subAsm As AssemblyDocument = CType(refDoc, AssemblyDocument)
                                    CollectPartsRecursively(subAsm, partsDict, occNames, subAsmDict, level & ">" & fileName)
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    ' Continue with next occurrence
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Copy all referenced sub-assemblies to destination folder (same filename)
        ''' </summary>
        Private Sub CopyAllSubAssemblies(ByVal allSubAssemblies As Dictionary(Of String, String), ByVal destFolder As String)
            Dim subAsmCounter As Integer = 0

            For Each kvp As KeyValuePair(Of String, String) In allSubAssemblies
                Dim sourcePath As String = kvp.Key
                Dim sourceFileName As String = kvp.Value
                Dim destFileName As String = sourceFileName

                If m_DoRename Then
                    subAsmCounter += 1
                    destFileName = m_PlantSection & "-ASM" & subAsmCounter.ToString() & ".iam"
                End If

                Dim destPath As String = System.IO.Path.Combine(destFolder, destFileName)

                Try
                    System.IO.File.Copy(sourcePath, destPath, True)
                    If Not m_CopiedFiles.ContainsKey(sourcePath) Then
                        m_CopiedFiles.Add(sourcePath, destPath)
                    End If
                    LogMessage("COPIED SUBASM: " & sourceFileName & " -> " & destFileName)
                Catch ex As Exception
                    LogMessage("ERROR copying sub-assembly " & sourceFileName & ": " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Get description from iProperty
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
        ''' Copy all parts to destination folder using File.Copy (same as VBScript fso.CopyFile)
        ''' Builds occurrence name mappings for iLogic patching
        ''' </summary>
        Private Sub CopyAllParts(ByVal allParts As Dictionary(Of String, String), ByVal occurrenceNames As Dictionary(Of String, String), ByVal destFolder As String)
            Dim counter As Integer = 0

            For Each kvp As KeyValuePair(Of String, String) In allParts
                Dim sourcePath As String = kvp.Key
                Dim description As String = kvp.Value
                Dim sourceFileName As String = System.IO.Path.GetFileName(sourcePath)
                Dim destPath As String

                If m_DoRename Then
                    ' Generate heritage name based on description
                    counter += 1
                    Dim suffix As String = ClassifyByDescription(description)
                    If suffix = "SKIP" Then
                        suffix = "PL" ' Default for hardware
                    End If
                    Dim newName As String = m_PlantSection & "-" & suffix & counter.ToString() & ".ipt"
                    Dim newNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(newName)
                    destPath = System.IO.Path.Combine(destFolder, newName)

                    ' Get the occurrence name for iLogic patching
                    Dim oldOccName As String = ""
                    If occurrenceNames.ContainsKey(sourcePath) Then
                        oldOccName = occurrenceNames(sourcePath)
                    Else
                        oldOccName = System.IO.Path.GetFileNameWithoutExtension(sourcePath)
                    End If

                    ' Track rename for iLogic patching
                    If Not m_OccurrenceRenames.ContainsKey(oldOccName) Then
                        m_OccurrenceRenames.Add(oldOccName, newNameNoExt)
                    End If

                    ' Also track filename renames
                    Dim oldFileNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(sourcePath)
                    If Not m_FileRenames.ContainsKey(oldFileNameNoExt) Then
                        m_FileRenames.Add(oldFileNameNoExt, newNameNoExt)
                    End If
                Else
                    destPath = System.IO.Path.Combine(destFolder, sourceFileName)
                End If

                ' CRITICAL: Use File.Copy (same as VBScript fso.CopyFile)
                Try
                    System.IO.File.Copy(sourcePath, destPath, True)
                    m_CopiedFiles.Add(sourcePath, destPath)
                    LogMessage("COPIED: " & sourceFileName & " -> " & System.IO.Path.GetFileName(destPath))
                Catch ex As Exception
                    LogMessage("ERROR copying " & sourceFileName & ": " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Classify part by description (same as VBScript)
        ''' </summary>
        Private Function ClassifyByDescription(ByVal description As String) As String
            Dim desc As String = description.ToUpper().Trim()

            ' Skip hardware
            If desc.Contains("BOLT") OrElse desc.Contains("SCREW") OrElse desc.Contains("WASHER") OrElse desc.Contains("NUT") Then
                Return "SKIP"
            End If

            ' Classification logic (same as Assembly_Cloner.vbs)
            If desc.StartsWith("UB") OrElse desc.StartsWith("UC") Then
                Return "B"
            ElseIf desc.StartsWith("PL") Then
                If desc.Contains("S355JR") Then
                    Return "PL"
                Else
                    Return "LPL"
                End If
            ElseIf desc.StartsWith("L") AndAlso (desc.Contains("X") OrElse desc.Contains(" X ")) Then
                Return "A"
            ElseIf desc.StartsWith("PFC") OrElse desc.StartsWith("TFC") Then
                Return "CH"
            ElseIf desc.StartsWith("CHS") Then
                Return "P"
            ElseIf desc.StartsWith("SHS") Then
                Return "SQ"
            ElseIf desc.StartsWith("FL") AndAlso Not desc.Contains("FLOOR") Then
                Return "FL"
            ElseIf desc.StartsWith("IPE") Then
                Return "IPE"
            Else
                Return "OTHER"
            End If
        End Function

        ''' <summary>
        ''' Update assembly references AND patch iLogic in SAME session
        ''' CRITICAL: Don't close and re-open - that causes Inventor to re-resolve references!
        ''' </summary>
        Private Function UpdateAssemblyAndPatchILogic(ByVal asmPath As String) As Integer
            Dim patchedRulesCount As Integer = 0

            LogMessage("ASSEMBLY UPDATE: Using fd.ReplaceReference method (same as VBScript)")
            m_InventorApp.SilentOperation = True

            Try
                ' Open the assembly
                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.Documents.Open(asmPath, False), AssemblyDocument)
                LogMessage("ASSEMBLY UPDATE: Opened " & System.IO.Path.GetFileName(asmPath))

                ' STEP 1: Update file references
                Dim fileDescriptors As Object = asmDoc.File.ReferencedFileDescriptors
                Dim fdCount As Integer = fileDescriptors.Count
                LogMessage("ASSEMBLY UPDATE: Found " & fdCount & " file references")

                Dim updatedCount As Integer = 0

                For i As Integer = 1 To fdCount
                    Dim fd As Object = fileDescriptors.Item(i)
                    Dim refPath As String = fd.FullFileName
                    Dim refFileName As String = System.IO.Path.GetFileName(refPath)

                    If m_CopiedFiles.ContainsKey(refPath) Then
                        Dim newRefPath As String = m_CopiedFiles(refPath)
                        Dim newFileName As String = System.IO.Path.GetFileName(newRefPath)

                        LogMessage("ASSEMBLY UPDATE: Replacing " & refFileName & " -> " & newFileName)

                        Try
                            fd.ReplaceReference(newRefPath)
                            updatedCount += 1
                            LogMessage("ASSEMBLY UPDATE: SUCCESS")
                        Catch ex As Exception
                            LogMessage("ASSEMBLY UPDATE: ERROR - " & ex.Message)
                        End Try
                    Else
                        LogMessage("ASSEMBLY UPDATE: No mapping for " & refFileName)
                    End If
                Next

                LogMessage("ASSEMBLY UPDATE: Updated " & updatedCount & " references")

                ' CRITICAL FIX: Rename occurrences to match new filenames
                ' fd.ReplaceReference updates FILE path but NOT occurrence name!
                LogMessage("ASSEMBLY UPDATE: Renaming occurrences to match new filenames...")
                Dim renameCount As Integer = RenameOccurrencesToMatchFiles(asmDoc)
                LogMessage("ASSEMBLY UPDATE: Renamed " & renameCount & " occurrences")

                ' CRITICAL: Update the assembly to force Inventor to resolve all new occurrences
                ' This ensures occurrence names match the new filenames before iLogic patching
                LogMessage("ASSEMBLY UPDATE: Forcing assembly update to resolve occurrences...")
                asmDoc.Update()
                LogMessage("ASSEMBLY UPDATE: Assembly updated and resolved")

                ' STEP 2: Patch iLogic rules (in SAME session, don't close/reopen!)
                If m_Patcher.IsAvailable AndAlso m_OccurrenceRenames.Count > 0 Then
                    m_iLogicPatchLog.AppendLine("=== iLOGIC PATCHING ===")
                    m_iLogicPatchLog.AppendLine("Mappings to apply (" & m_OccurrenceRenames.Count & "):")
                    For Each kvp In m_OccurrenceRenames
                        m_iLogicPatchLog.AppendLine("  " & kvp.Key & " -> " & kvp.Value)
                    Next
                    m_iLogicPatchLog.AppendLine()

                    LogMessage("Patching iLogic rules (in same session)...")
                    patchedRulesCount = m_Patcher.PatchRulesRecursive(asmDoc, m_OccurrenceRenames)
                    m_iLogicPatchLog.AppendLine("Rules patched: " & patchedRulesCount)
                    LogMessage("Patched " & patchedRulesCount & " iLogic rules")
                End If

                ' STEP 3: Save and close (only NOW after everything is done)
                LogMessage("ASSEMBLY UPDATE: Saving assembly...")
                asmDoc.Save()
                asmDoc.Close(False)
                LogMessage("ASSEMBLY UPDATE: Complete - references updated and iLogic patched")

            Catch ex As Exception
                LogMessage("ASSEMBLY UPDATE: ERROR - " & ex.Message)
            Finally
                m_InventorApp.SilentOperation = False
            End Try

            Return patchedRulesCount
        End Function

        ''' <summary>
        ''' OLD METHOD - DEPRECATED - Use UpdateAssemblyAndPatchILogic instead
        ''' </summary>
        Private Sub UpdateAssemblyReferencesWithFileDescriptors(ByVal asmPath As String)
            LogMessage("ASSEMBLY UPDATE: Using fd.ReplaceReference method (same as VBScript)")

            m_InventorApp.SilentOperation = True

            Try
                ' Open the assembly
                Dim asmDoc As AssemblyDocument = CType(m_InventorApp.Documents.Open(asmPath, False), AssemblyDocument)
                LogMessage("ASSEMBLY UPDATE: Opened " & System.IO.Path.GetFileName(asmPath))

                ' CRITICAL FIX: Use late binding (Object) to avoid COM cast issues
                ' VBScript doesn't have strict typing so it works fine there
                Dim fileDescriptors As Object = asmDoc.File.ReferencedFileDescriptors
                Dim fdCount As Integer = fileDescriptors.Count
                LogMessage("ASSEMBLY UPDATE: Found " & fdCount & " file references")

                Dim updatedCount As Integer = 0

                For i As Integer = 1 To fdCount
                    Dim fd As Object = fileDescriptors.Item(i)
                    Dim refPath As String = fd.FullFileName
                    Dim refFileName As String = System.IO.Path.GetFileName(refPath)

                    ' Check if we have a mapping for this file
                    If m_CopiedFiles.ContainsKey(refPath) Then
                        Dim newRefPath As String = m_CopiedFiles(refPath)
                        Dim newFileName As String = System.IO.Path.GetFileName(newRefPath)

                        LogMessage("ASSEMBLY UPDATE: Replacing " & refFileName & " -> " & newFileName)

                        Try
                            ' EXACT same method as VBScript: fd.ReplaceReference newRefPath
                            fd.ReplaceReference(newRefPath)
                            updatedCount += 1
                            LogMessage("ASSEMBLY UPDATE: SUCCESS")
                        Catch ex As Exception
                            LogMessage("ASSEMBLY UPDATE: ERROR - " & ex.Message)
                        End Try
                    Else
                        LogMessage("ASSEMBLY UPDATE: No mapping for " & refFileName)
                    End If
                Next

                ' Save the assembly
                LogMessage("ASSEMBLY UPDATE: Saving assembly...")
                asmDoc.Save()
                asmDoc.Close(False)
                LogMessage("ASSEMBLY UPDATE: Saved successfully (" & updatedCount & " references updated)")

            Catch ex As Exception
                LogMessage("ASSEMBLY UPDATE: ERROR - " & ex.Message)
            Finally
                m_InventorApp.SilentOperation = False
            End Try
        End Sub

        ''' <summary>
        ''' Rename all occurrences to match their new filenames
        ''' CRITICAL: fd.ReplaceReference updates file path but NOT occurrence name!
        ''' </summary>
        Private Function RenameOccurrencesToMatchFiles(ByVal asmDoc As AssemblyDocument) As Integer
            Dim renameCount As Integer = 0

            Try
                Dim occurrences As ComponentOccurrences = asmDoc.ComponentDefinition.Occurrences

                For i As Integer = 1 To occurrences.Count
                    Try
                        Dim occ As ComponentOccurrence = occurrences.Item(i)
                        Dim refPath As String = occ.ReferencedFileDescriptor.FullFileName
                        Dim currentFileName As String = System.IO.Path.GetFileNameWithoutExtension(refPath)

                        ' Get current occurrence name (without instance number)
                        Dim currentOccName As String = occ.Name
                        Dim colonPos As Integer = currentOccName.LastIndexOf(":"c)
                        Dim occNameWithoutInstance As String = currentOccName
                        Dim instanceNum As String = ""
                        If colonPos > 0 Then
                            occNameWithoutInstance = currentOccName.Substring(0, colonPos)
                            instanceNum = currentOccName.Substring(colonPos) ' Keep ":1", ":2", etc.
                        End If

                        ' Check if occurrence name needs updating
                        If occNameWithoutInstance <> currentFileName Then
                            ' Rename occurrence to match filename
                            Dim newOccName As String = currentFileName & instanceNum
                            occ.Name = newOccName
                            renameCount += 1
                            LogMessage("  Renamed: """ & currentOccName & """ -> """ & newOccName & """")
                        End If

                    Catch ex As Exception
                        ' Skip this occurrence if error
                        LogMessage("  Error renaming occurrence: " & ex.Message)
                    End Try
                Next

            Catch ex As Exception
                LogMessage("ERROR in RenameOccurrencesToMatchFiles: " & ex.Message)
            End Try

            Return renameCount
        End Function

        ''' <summary>
        ''' Process IDW files using fd.ReplaceReference (EXACT same as VBScript ProcessIDWFilesWithReferenceUpdate)
        ''' </summary>
        Private Sub ProcessIDWFilesWithReferenceUpdate(ByVal sourceDir As String, ByVal destFolder As String, ByVal sourceAsmPath As String, ByVal newAsmPath As String)
            Try
                ' Find all IDW files in source directory
                Dim idwFiles As String() = Directory.GetFiles(sourceDir, "*.idw")
                LogMessage("IDW PROCESS: Found " & idwFiles.Length & " IDW files")

                m_InventorApp.SilentOperation = True

                For Each idwPath As String In idwFiles
                    Dim idwFileName As String = System.IO.Path.GetFileName(idwPath)
                    Dim originalName As String = System.IO.Path.GetFileNameWithoutExtension(idwPath)

                    ' Ask for new IDW name
                    Dim newIdwName As String = InputBox(
                        "Enter new name for IDW '" & originalName & "':" & vbCrLf & vbCrLf &
                        "This will be the filename for the cloned drawing." & vbCrLf &
                        "Do NOT include .idw extension.",
                        "New IDW Name",
                        originalName)

                    If String.IsNullOrEmpty(newIdwName) Then
                        newIdwName = originalName
                    End If

                    Dim newIdwPath As String = System.IO.Path.Combine(destFolder, newIdwName & ".idw")

                    LogMessage("IDW PROCESS: Processing " & idwFileName & " -> " & newIdwName & ".idw")

                    ' Close all documents first (same as VBScript)
                    m_InventorApp.Documents.CloseAll()

                    ' Open the ORIGINAL IDW (it has valid references to original parts)
                    ' Same as VBScript: Set idwDoc = invApp.Documents.Open(file.Path, False)
                    LogMessage("IDW PROCESS: Opening original IDW from source location...")
                    Dim idwDoc As DrawingDocument = CType(m_InventorApp.Documents.Open(idwPath, False), DrawingDocument)

                    LogMessage("IDW PROCESS: Opened successfully, now updating references...")

                    ' Update references using fd.ReplaceReference (EXACT same as VBScript)
                    ' CRITICAL FIX: Use late binding (Object) to avoid COM cast issues
                    Dim fileDescriptors As Object = idwDoc.File.ReferencedFileDescriptors
                    Dim fdCount As Integer = fileDescriptors.Count
                    LogMessage("IDW PROCESS: Found " & fdCount & " references")

                    Dim updatedCount As Integer = 0

                    For i As Integer = 1 To fdCount
                        Dim fd As Object = fileDescriptors.Item(i)
                        Dim refPath As String = fd.FullFileName
                        Dim refFileName As String = System.IO.Path.GetFileName(refPath)

                        Dim newRefPath As String = ""

                        ' Check if we have a mapping for this file (original -> new)
                        If TryGetMappedPath(refPath, newRefPath) Then

                            LogMessage("IDW PROCESS: Updating " & refFileName & " -> " & System.IO.Path.GetFileName(newRefPath))

                            Try
                                ' EXACT same method as VBScript: fd.ReplaceReference newRefPath
                                fd.ReplaceReference(newRefPath)
                                updatedCount += 1
                                LogMessage("IDW PROCESS: SUCCESS")
                            Catch ex As Exception
                                LogMessage("IDW PROCESS: ERROR - " & ex.Message)
                            End Try
                        ElseIf refFileName.Equals(System.IO.Path.GetFileName(sourceAsmPath), StringComparison.OrdinalIgnoreCase) Then
                            LogMessage("IDW PROCESS: Main assembly fallback mapping " & refFileName & " -> " & System.IO.Path.GetFileName(newAsmPath))

                            Try
                                fd.ReplaceReference(newAsmPath)
                                updatedCount += 1
                                LogMessage("IDW PROCESS: SUCCESS")
                            Catch ex As Exception
                                LogMessage("IDW PROCESS: ERROR - " & ex.Message)
                            End Try
                        Else
                            LogMessage("IDW PROCESS: No mapping for " & refFileName)
                        End If
                    Next

                    ' Patch iLogic in IDW too
                    If m_Patcher.IsAvailable AndAlso m_OccurrenceRenames.Count > 0 Then
                        LogMessage("IDW PROCESS: Patching iLogic in IDW...")
                        Dim patchedInIdw As Integer = m_Patcher.PatchRules(idwDoc, m_OccurrenceRenames)
                        LogMessage("IDW PROCESS: Patched " & patchedInIdw & " rules in IDW")
                    End If

                    ' Update iProperties in the IDW (Part Number and Description)
                    Try
                        LogMessage("IDW PROCESS: Updating iProperties for " & idwFileName)
                        Dim newAsmName As String = System.IO.Path.GetFileNameWithoutExtension(newAsmPath)

                        ' Update Part Number AND Description in Design Tracking Properties
                        ' NOTE: "Description" is in Design Tracking Properties (Project tab)
                        '       "Comments" is in Inventor Summary Information (Summary tab) - DIFFERENT field!
                        Try
                            Dim designProps As PropertySet = idwDoc.PropertySets.Item("Design Tracking Properties")

                            ' Update Part Number to IDW's new name
                            designProps.Item("Part Number").Value = newIdwName
                            LogMessage("IDW PROCESS: Updated Part Number to '" & newIdwName & "'")

                            ' Update Description to assembly's new name
                            designProps.Item("Description").Value = newAsmName
                            LogMessage("IDW PROCESS: Updated Description to '" & newAsmName & "'")
                        Catch ex As Exception
                            LogMessage("IDW PROCESS: Error updating Design Tracking Properties: " & ex.Message)
                        End Try

                        LogMessage("IDW PROCESS: iProperties update complete for " & idwFileName)
                    Catch ex As Exception
                        LogMessage("IDW PROCESS: Error updating iProperties: " & ex.Message)
                    End Try

                    ' Save to NEW location (same as VBScript: idwDoc.SaveAs destIdwPath, False)
                    LogMessage("IDW PROCESS: Saving to destination: " & newIdwPath)
                    idwDoc.SaveAs(newIdwPath, False)
                    LogMessage("IDW PROCESS: Successfully saved " & idwFileName & " as " & newIdwName & ".idw (" & updatedCount & " references updated)")

                    idwDoc.Close()
                Next

                m_InventorApp.SilentOperation = False
                LogMessage("IDW PROCESS: All IDW files processed successfully")

            Catch ex As Exception
                LogMessage("IDW PROCESS: ERROR - " & ex.Message)
                m_InventorApp.SilentOperation = False
            End Try
        End Sub

        ''' <summary>
        ''' Resolve mapped path using exact full path first, then filename fallback for normalized path mismatches.
        ''' </summary>
        Private Function TryGetMappedPath(ByVal sourcePath As String, ByRef mappedPath As String) As Boolean
            mappedPath = ""

            If m_CopiedFiles.ContainsKey(sourcePath) Then
                mappedPath = m_CopiedFiles(sourcePath)
                Return True
            End If

            Dim sourceFile As String = System.IO.Path.GetFileName(sourcePath)
            For Each kvp As KeyValuePair(Of String, String) In m_CopiedFiles
                If System.IO.Path.GetFileName(kvp.Key).Equals(sourceFile, StringComparison.OrdinalIgnoreCase) Then
                    mappedPath = kvp.Value
                    Return True
                End If
            Next

            Return False
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
                m_LogPath = System.IO.Path.Combine(logsFolder, "AssemblyCloner_" & timestamp & ".log")
                m_LogFile = New StreamWriter(m_LogPath, False, System.Text.Encoding.UTF8)
                m_LogFile.AutoFlush = True

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Failed to initialize log: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Log a message with timestamp
        ''' </summary>
        Private Sub LogMessage(ByVal message As String)
            Try
                If m_LogFile IsNot Nothing Then
                    m_LogFile.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | " & message)
                End If
                System.Diagnostics.Debug.WriteLine("AssemblyCloner: " & message)
            Catch ex As Exception
                ' Ignore logging errors
            End Try
        End Sub

        ''' <summary>
        ''' Close the log file
        ''' </summary>
        Private Sub CloseLog()
            Try
                If m_LogFile IsNot Nothing Then
                    m_LogFile.Flush()
                    m_LogFile.Close()
                    m_LogFile = Nothing
                End If
            Catch ex As Exception
                ' Ignore
            End Try
        End Sub

        ''' <summary>
        ''' Save the iLogic patch log to a file
        ''' </summary>
        Private Sub SaveiLogicPatchLog(ByVal destFolder As String)
            Try
                Dim logsFolder As String = System.IO.Path.Combine(destFolder, "Logs")
                If Not Directory.Exists(logsFolder) Then
                    Directory.CreateDirectory(logsFolder)
                End If

                Dim timestamp As String = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")
                Dim logPath As String = System.IO.Path.Combine(logsFolder, "iLogic_Patch_" & timestamp & ".txt")

                Using writer As New StreamWriter(logPath)
                    writer.WriteLine("=== iLOGIC PATCHING LOG ===")
                    writer.WriteLine("Generated: " & DateTime.Now.ToString())
                    writer.WriteLine("Destination: " & destFolder)
                    writer.WriteLine()
                    writer.WriteLine(m_iLogicPatchLog.ToString())
                    writer.WriteLine()
                    writer.WriteLine("=== ALL OCCURRENCE MAPPINGS ===")
                    For Each kvp In m_OccurrenceRenames
                        writer.WriteLine("  """ & kvp.Key & """ -> """ & kvp.Value & """")
                    Next
                    writer.WriteLine()
                    writer.WriteLine("=== ALL FILE MAPPINGS ===")
                    For Each kvp In m_FileRenames
                        writer.WriteLine("  " & kvp.Key & " -> " & kvp.Value)
                    Next
                End Using

                System.Diagnostics.Debug.WriteLine("iLogic patch log saved to: " & logPath)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error saving iLogic patch log: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Update iProperties for all copied documents (parts and assemblies)
        ''' </summary>
        Private Sub UpdateIPropertiesForCopiedDocuments()
            For Each kvp As KeyValuePair(Of String, String) In m_CopiedFiles
                Dim oldPath As String = kvp.Key
                Dim newPath As String = kvp.Value
                Dim ext As String = System.IO.Path.GetExtension(newPath).ToLower()

                If ext = ".ipt" Or ext = ".iam" Then
                    Try
                        Dim doc As Document = m_InventorApp.Documents.Open(newPath, False)
                        Dim oldName As String = System.IO.Path.GetFileNameWithoutExtension(oldPath)
                        Dim newName As String = System.IO.Path.GetFileNameWithoutExtension(newPath)

                        ' Find differing suffix
                        Dim minLen As Integer = Math.Min(oldName.Length, newName.Length)
                        Dim diffIndex As Integer = 0
                        While diffIndex < minLen AndAlso oldName(diffIndex) = newName(diffIndex)
                            diffIndex += 1
                        End While
                        Dim oldSuffix As String = oldName.Substring(diffIndex)
                        Dim newSuffix As String = newName.Substring(diffIndex)

                        Dim replaced As Boolean = False
                        For Each propSet As PropertySet In doc.PropertySets
                            For Each prop As Inventor.Property In propSet
                                If prop.Value IsNot Nothing AndAlso TypeOf prop.Value Is String Then
                                    Dim valueStr As String = prop.Value.ToString()
                                    ' For parts, skip Description (Comments property)
                                    If ext = ".ipt" AndAlso prop.Name = "Comments" Then Continue For

                                    If oldSuffix <> "" AndAlso Regex.IsMatch(valueStr, Regex.Escape(oldSuffix), RegexOptions.IgnoreCase) Then
                                        prop.Value = Regex.Replace(valueStr, Regex.Escape(oldSuffix), newSuffix, RegexOptions.IgnoreCase)
                                        replaced = True
                                    ElseIf oldName <> newName AndAlso Regex.IsMatch(valueStr, Regex.Escape(oldName), RegexOptions.IgnoreCase) Then
                                        prop.Value = Regex.Replace(valueStr, Regex.Escape(oldName), newName, RegexOptions.IgnoreCase)
                                        replaced = True
                                    End If
                                End If
                            Next
                        Next

                        If replaced Then
                            LogMessage("Updated iProperties for " & System.IO.Path.GetFileName(newPath))
                        End If

                        doc.Save()
                        doc.Close(False)

                    Catch ex As Exception
                        LogMessage("Error updating iProperties for " & System.IO.Path.GetFileName(newPath) & ": " & ex.Message)
                    End Try
                End If
            Next
        End Sub

    End Class

End Namespace
