Imports Inventor
Imports System.IO
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic

Public Class AssemblyCloner
    Private m_inventorApp As Inventor.Application
    Private m_copiedFiles As Dictionary(Of String, String)  ' originalPath -> newPath
    Private m_partNameMapping As Dictionary(Of String, String)  ' old occurrence name -> new occurrence name
    Private m_log As System.Text.StringBuilder
    Private m_logPath As String

    Public Sub New(inventorApp As Inventor.Application)
        m_inventorApp = inventorApp
        m_copiedFiles = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        m_partNameMapping = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        m_log = New System.Text.StringBuilder()
    End Sub

    Private Sub Log(message As String)
        Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss")
        m_log.AppendLine(timestamp & " | " & message)
    End Sub

    Private Sub SaveLog(destFolder As String)
        Try
            m_logPath = System.IO.Path.Combine(destFolder, "AssemblyCloner_Log_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
            System.IO.File.WriteAllText(m_logPath, m_log.ToString())
        Catch
        End Try
    End Sub

    Public Sub CloneAndPatch()
        Try
            Log("=== ASSEMBLY CLONER STARTED ===")

            Dim activeDoc As Document = m_inventorApp.ActiveDocument
            If activeDoc Is Nothing Then
                Log("ERROR: No active document")
                MsgBox("No active document!", MsgBoxStyle.Critical)
                Return
            End If

            Log("Active document: " & activeDoc.FullFileName)
            Log("Document type: " & activeDoc.DocumentType.ToString())

            If activeDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                Log("ERROR: Not an assembly document")
                MsgBox("Please open an assembly document.", MsgBoxStyle.Exclamation)
                Return
            End If

            Dim asmDoc As AssemblyDocument = DirectCast(activeDoc, AssemblyDocument)
            Dim originalAsmPath As String = asmDoc.FullFileName
            Dim originalAsmName As String = System.IO.Path.GetFileNameWithoutExtension(originalAsmPath)
            Dim sourceDir As String = System.IO.Path.GetDirectoryName(originalAsmPath)

            Log("Source assembly: " & originalAsmName)
            Log("Source directory: " & sourceDir)

            ' Step 1: Get new assembly name from user
            Log("STEP 1: Getting new assembly name from user")
            Dim newName As String = InputBox("Enter new assembly name (without extension):" & vbCrLf & vbCrLf & _
                                             "Current: " & originalAsmName, "Clone Assembly", originalAsmName & "_Clone")
            If String.IsNullOrEmpty(newName) Then
                Log("User cancelled - no name entered")
                Return
            End If
            Log("New name: " & newName)

            ' Step 2: Get destination folder
            Log("STEP 2: Getting destination folder")
            Dim destFolder As String = GetDestinationFolder(sourceDir)
            If String.IsNullOrEmpty(destFolder) Then
                Log("User cancelled - no destination folder")
                Return
            End If
            Log("Destination folder: " & destFolder)

            ' ============================================================
            ' FLOW MATCHING OPTION 1 (Assembly_Renamer.vbs)
            ' Key: Keep assembly OPEN, update refs, then SaveAs
            ' This avoids "where is file?" dialogs
            ' ============================================================

            ' Step 3: Collect all part info while assembly is OPEN
            Log("STEP 3: Collecting part information (assembly stays OPEN)")
            Dim partInfo As New Dictionary(Of String, String)
            CollectPartInfo(asmDoc, partInfo)
            Log("Parts collected: " & partInfo.Count)

            For Each kvp As KeyValuePair(Of String, String) In partInfo
                Log("  PART: " & System.IO.Path.GetFileName(kvp.Key) & " -> OccName: " & kvp.Value)
            Next

            ' Step 4: Copy all parts to destination (assembly still OPEN)
            Log("STEP 4: Copying parts to destination")
            CopyAllParts(partInfo, destFolder, newName)
            Log("Parts copied: " & m_copiedFiles.Count)

            For Each kvp As KeyValuePair(Of String, String) In m_partNameMapping
                Log("  MAPPING: """ & kvp.Key & """ -> """ & kvp.Value & """")
            Next

            ' Step 5: Update assembly references IN THE OPEN ASSEMBLY using occ.Replace
            ' This is the key difference - we update while Inventor has everything loaded
            Log("STEP 5: Updating assembly references (assembly still OPEN)")
            UpdateAssemblyReferencesInOpenDoc(asmDoc)

            ' Step 6: SaveAs the assembly to new location with new name
            Log("STEP 6: Saving assembly to new location")
            Dim newAsmPath As String = System.IO.Path.Combine(destFolder, newName & ".iam")
            asmDoc.SaveAs(newAsmPath, False)
            m_copiedFiles.Add(originalAsmPath, newAsmPath)
            Log("Assembly saved to: " & newAsmPath)

            ' Step 7: Process IDW files (open from source, update refs, SaveAs to dest)
            Log("STEP 7: Processing IDW files")
            ProcessIDWFilesFromSource(sourceDir, destFolder, newName)

            ' Step 8: Patch iLogic rules in the new assembly
            Log("STEP 8: Patching iLogic rules")
            Dim patcher As New iLogicPatcher(m_inventorApp, m_log)
            Dim patchedCount As Integer = patcher.PatchRulesInAssembly(asmDoc, m_partNameMapping)
            Log("iLogic patching complete. Replacements: " & patchedCount)

            ' Step 9: Save again after iLogic patching
            Log("STEP 9: Final save")
            asmDoc.Save()
            Log("Assembly saved")

            Log("=== ASSEMBLY CLONER COMPLETED ===")

            ' Save log file
            SaveLog(destFolder)

            MsgBox("Assembly cloned successfully!" & vbCrLf & vbCrLf & _
                   "Assembly: " & newName & ".iam" & vbCrLf & _
                   "Parts copied: " & m_copiedFiles.Count & vbCrLf & _
                   "Part name mappings: " & m_partNameMapping.Count & vbCrLf & _
                   "iLogic replacements: " & patchedCount & vbCrLf & _
                   "Destination: " & destFolder & vbCrLf & vbCrLf & _
                   "Log file: " & m_logPath, MsgBoxStyle.Information, "Success!")

        Catch ex As Exception
            Log("EXCEPTION: " & ex.Message)
            Log("Stack trace: " & ex.StackTrace)

            ' Try to save log even on error
            Try
                Dim errorLogPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "AssemblyCloner_Error_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
                System.IO.File.WriteAllText(errorLogPath, m_log.ToString())
                MsgBox("Error: " & ex.Message & vbCrLf & vbCrLf & "Log saved to: " & errorLogPath, MsgBoxStyle.Critical)
            Catch
                MsgBox("Error: " & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, MsgBoxStyle.Critical)
            End Try
        End Try
    End Sub

    Private Function GetDestinationFolder(sourceDir As String) As String
        ' Use the modern Windows Vista+ dialog (IFileOpenDialog) which allows full navigation
        ' Fall back to manual entry if COM dialog fails

        Try
            ' Try to use the modern folder picker via COM
            Dim modernPath As String = ShowModernFolderPicker(sourceDir)
            If Not String.IsNullOrEmpty(modernPath) Then
                If modernPath.Equals(sourceDir, StringComparison.OrdinalIgnoreCase) Then
                    MsgBox("Destination cannot be the same as source folder!", MsgBoxStyle.Exclamation)
                    Return Nothing
                End If
                Return modernPath
            End If
        Catch ex As Exception
            Log("Modern folder picker failed: " & ex.Message)
        End Try

        ' Fall back to manual path entry - more flexible than old dialog
        Log("Using manual path entry as fallback")
        Dim parentDir As String = System.IO.Path.GetDirectoryName(sourceDir)

        Dim userPath As String = InputBox(
            "DESTINATION FOLDER" & vbCrLf & vbCrLf & _
            "Source: " & sourceDir & vbCrLf & vbCrLf & _
            "Enter the FULL PATH to destination folder:" & vbCrLf & _
            "(Folder will be created if it doesn't exist)",
            "Select Destination",
            parentDir & "\NewAssembly")

        If String.IsNullOrEmpty(userPath) Then
            Return Nothing
        End If

        userPath = userPath.Trim()

        ' Validate
        If userPath.Equals(sourceDir, StringComparison.OrdinalIgnoreCase) Then
            MsgBox("Destination cannot be the same as source folder!", MsgBoxStyle.Exclamation)
            Return Nothing
        End If

        ' Create folder if it doesn't exist
        If Not System.IO.Directory.Exists(userPath) Then
            Dim createResult As MsgBoxResult = MsgBox(
                "Folder does not exist:" & vbCrLf & userPath & vbCrLf & vbCrLf & _
                "Create this folder?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Create Folder?")

            If createResult = MsgBoxResult.Yes Then
                Try
                    System.IO.Directory.CreateDirectory(userPath)
                    Log("Created destination folder: " & userPath)
                Catch dirEx As Exception
                    MsgBox("ERROR: Could not create folder!" & vbCrLf & dirEx.Message, MsgBoxStyle.Critical)
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End If

        Return userPath
    End Function

    Private Function ShowModernFolderPicker(startPath As String) As String
        ' Use Windows Vista+ IFileOpenDialog for modern folder selection
        ' This allows full navigation unlike the old FolderBrowserDialog

        Try
            Dim dialog As New System.Windows.Forms.OpenFileDialog()

            ' Configure it to select folders instead of files
            ' We use a hack: set ValidateNames and CheckFileExists to false,
            ' and set FileName to "Folder Selection"
            dialog.ValidateNames = False
            dialog.CheckFileExists = False
            dialog.CheckPathExists = True
            dialog.FileName = "Select Folder"
            dialog.Title = "Select Destination Folder (navigate to folder and click Open)"

            ' Set initial directory to parent of source
            If Not String.IsNullOrEmpty(startPath) AndAlso System.IO.Directory.Exists(startPath) Then
                dialog.InitialDirectory = startPath
            End If

            If dialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                ' Get the selected path (file or folder)
                Dim selectedPath As String = System.IO.Path.GetDirectoryName(dialog.FileName)
                Return selectedPath
            End If

            Return Nothing
        Catch ex As Exception
            Log("OpenFileDialog folder picker failed: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub CollectPartInfo(asmDoc As AssemblyDocument, partInfo As Dictionary(Of String, String))
        Log("  Scanning assembly occurrences...")
        Dim occCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
        Log("  Top-level occurrences: " & occCount)

        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            CollectPartInfoRecursively(occ, partInfo, 1)
        Next
    End Sub

    Private Sub CollectPartInfoRecursively(occ As ComponentOccurrence, partInfo As Dictionary(Of String, String), level As Integer)
        Try
            Dim indent As String = New String(" "c, level * 2)

            If occ.Suppressed Then
                Log(indent & "SKIP (suppressed): " & occ.Name)
                Return
            End If

            Dim doc As Document = Nothing
            Try
                doc = occ.Definition.Document
            Catch ex As Exception
                Log(indent & "SKIP (no document): " & occ.Name & " - " & ex.Message)
                Return
            End Try

            If doc Is Nothing Then
                Log(indent & "SKIP (null document): " & occ.Name)
                Return
            End If

            Dim fullPath As String = doc.FullFileName
            Dim occName As String = occ.Name

            Log(indent & "Processing: " & occName & " (" & System.IO.Path.GetFileName(fullPath) & ")")

            If fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                If Not partInfo.ContainsKey(fullPath) Then
                    Dim baseName As String = GetBaseOccurrenceName(occName)
                    partInfo.Add(fullPath, baseName)
                    Log(indent & "  ADDED: " & baseName)
                Else
                    Log(indent & "  ALREADY EXISTS")
                End If
            ElseIf fullPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                Log(indent & "  SUB-ASSEMBLY - recursing...")
                Dim subAsm As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                For Each subOcc As ComponentOccurrence In subAsm.ComponentDefinition.Occurrences
                    CollectPartInfoRecursively(subOcc, partInfo, level + 1)
                Next
            End If
        Catch ex As Exception
            Log("  ERROR in CollectPartInfoRecursively: " & ex.Message)
        End Try
    End Sub

    Private Function GetBaseOccurrenceName(occName As String) As String
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos > 0 Then
            Return occName.Substring(0, colonPos)
        End If
        Return occName
    End Function

    Private Sub CopyAllParts(partInfo As Dictionary(Of String, String), destFolder As String, newPrefix As String)
        For Each kvp As KeyValuePair(Of String, String) In partInfo
            Dim originalPath As String = kvp.Key
            Dim occBaseName As String = kvp.Value

            If System.IO.File.Exists(originalPath) Then
                Dim originalFileName As String = System.IO.Path.GetFileNameWithoutExtension(originalPath)
                Dim newFileName As String = newPrefix & "_" & originalFileName & ".ipt"
                Dim newPath As String = System.IO.Path.Combine(destFolder, newFileName)

                Try
                    System.IO.File.Copy(originalPath, newPath, True)
                    m_copiedFiles.Add(originalPath, newPath)
                    Log("  COPIED: " & originalFileName & ".ipt -> " & newFileName)

                    ' Build the occurrence name mapping for iLogic patching
                    Dim newOccBaseName As String = newPrefix & "_" & occBaseName
                    If Not m_partNameMapping.ContainsKey(occBaseName) Then
                        m_partNameMapping.Add(occBaseName, newOccBaseName)
                        Log("  MAPPING ADDED: " & occBaseName & " -> " & newOccBaseName)
                    End If
                Catch ex As Exception
                    Log("  ERROR copying " & originalFileName & ": " & ex.Message)
                End Try
            Else
                Log("  FILE NOT FOUND: " & originalPath)
            End If
        Next
    End Sub

    Private Sub UpdateAssemblyReferencesWithApprentice(asmPath As String)
        Try
            Log("  Creating ApprenticeServerComponent...")
            Dim apprentice As Object = CreateObject("Inventor.ApprenticeServerComponent")
            Log("  ApprenticeServer created")

            Log("  Opening assembly with ApprenticeServer...")
            Dim appDoc As Object = apprentice.Open(asmPath)
            Log("  Assembly opened")

            Dim fileDescriptors As Object = appDoc.ReferencedFileDescriptors
            Dim fdCount As Integer = CInt(fileDescriptors.Count)
            Log("  File descriptors count: " & fdCount)

            Dim updatedCount As Integer = 0
            For i As Integer = 1 To fdCount
                Try
                    Dim fd As Object = fileDescriptors.Item(i)
                    Dim refPath As String = CStr(fd.FullFileName)

                    If m_copiedFiles.ContainsKey(refPath) Then
                        Dim newRefPath As String = m_copiedFiles(refPath)
                        fd.ReplaceReference(newRefPath)
                        updatedCount += 1
                        Log("  REPLACED: " & System.IO.Path.GetFileName(refPath) & " -> " & System.IO.Path.GetFileName(newRefPath))
                    End If
                Catch ex As Exception
                    Log("  ERROR on descriptor " & i & ": " & ex.Message)
                End Try
            Next

            Log("  Saving assembly...")
            appDoc.SaveAs(asmPath, False)
            Log("  Assembly saved. Updated " & updatedCount & " references")

            appDoc.Close()
            apprentice = Nothing

        Catch ex As Exception
            Log("  ApprenticeServer FAILED: " & ex.Message)
            Log("  Falling back to Inventor method...")
            UpdateAssemblyReferencesWithInventor(asmPath)
        End Try
    End Sub

    Private Sub UpdateAssemblyReferencesWithInventor(asmPath As String)
        Try
            Log("  Opening assembly with Inventor...")
            Dim asmDoc As AssemblyDocument = DirectCast(m_inventorApp.Documents.Open(asmPath, False), AssemblyDocument)
            Log("  Assembly opened")

            UpdateReferencesRecursively(asmDoc)

            asmDoc.Save()
            asmDoc.Close()
            Log("  Inventor method complete")
        Catch ex As Exception
            Log("  Inventor method FAILED: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateReferencesRecursively(asmDoc As AssemblyDocument)
        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Try
                If occ.Suppressed Then Continue For

                Dim doc As Document = occ.Definition.Document
                If doc Is Nothing Then Continue For

                Dim fullPath As String = doc.FullFileName

                If fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    If m_copiedFiles.ContainsKey(fullPath) Then
                        Dim newPath As String = m_copiedFiles(fullPath)
                        occ.Replace(newPath, True)
                    End If
                ElseIf fullPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    Dim subAsm As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                    UpdateReferencesRecursively(subAsm)
                    subAsm.Save()
                End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Private Sub ProcessIDWFiles(sourceDir As String, destFolder As String, newPrefix As String)
        ' Use ApprenticeServer for IDW processing - completely SILENT, no dialogs!
        ' 1. Copy IDW to destination first
        ' 2. Use ApprenticeServer to update references (no Inventor dialogs)

        Dim idwFiles() As String = System.IO.Directory.GetFiles(sourceDir, "*.idw")
        Log("  Found " & idwFiles.Length & " IDW files in source folder")

        If idwFiles.Length = 0 Then Return

        Try
            Log("  Creating ApprenticeServer for IDW processing...")
            Dim apprentice As Object = CreateObject("Inventor.ApprenticeServerComponent")
            Log("  ApprenticeServer created")

            For Each idwPath As String In idwFiles
                Try
                    Dim idwFileName As String = System.IO.Path.GetFileName(idwPath)
                    Dim newIdwName As String = newPrefix & "_" & idwFileName
                    Dim newIdwPath As String = System.IO.Path.Combine(destFolder, newIdwName)

                    Log("  Processing IDW: " & idwFileName)

                    ' Step 1: Copy IDW file to destination
                    Log("    Copying IDW to destination...")
                    System.IO.File.Copy(idwPath, newIdwPath, True)
                    Log("    Copied to: " & newIdwPath)

                    ' Step 2: Open the COPIED IDW with ApprenticeServer (silent!)
                    Log("    Opening with ApprenticeServer...")
                    Dim appDoc As Object = apprentice.Open(newIdwPath)
                    Log("    IDW opened with ApprenticeServer")

                    ' Step 3: Update references
                    Dim fileDescriptors As Object = appDoc.ReferencedFileDescriptors
                    Dim fdCount As Integer = CInt(fileDescriptors.Count)
                    Log("    File descriptors: " & fdCount)

                    Dim updatedCount As Integer = 0
                    For i As Integer = 1 To fdCount
                        Try
                            Dim fd As Object = fileDescriptors.Item(i)
                            Dim refPath As String = CStr(fd.FullFileName)

                            If m_copiedFiles.ContainsKey(refPath) Then
                                Dim newRefPath As String = m_copiedFiles(refPath)
                                fd.ReplaceReference(newRefPath)
                                updatedCount += 1
                                Log("    REPLACED: " & System.IO.Path.GetFileName(refPath))
                            End If
                        Catch ex As Exception
                            Log("    ERROR on descriptor " & i & ": " & ex.Message)
                        End Try
                    Next

                    ' Step 4: Save
                    Log("    Saving IDW...")
                    appDoc.SaveAs(newIdwPath, False)
                    appDoc.Close()
                    Log("    IDW complete. Updated " & updatedCount & " references")

                Catch ex As Exception
                    Log("  IDW ERROR: " & ex.Message)
                End Try
            Next

            apprentice = Nothing
            Log("  ApprenticeServer released")

        Catch ex As Exception
            Log("  ApprenticeServer ERROR: " & ex.Message)
            Log("  Falling back to Inventor method...")
            ProcessIDWFilesWithInventor(sourceDir, destFolder, newPrefix)
        End Try
    End Sub

    Private Sub ProcessIDWFilesWithInventor(sourceDir As String, destFolder As String, newPrefix As String)
        ' Fallback method using Inventor (may show dialogs)
        m_inventorApp.Documents.CloseAll()

        For Each idwPath As String In System.IO.Directory.GetFiles(sourceDir, "*.idw")
            Try
                Dim idwFileName As String = System.IO.Path.GetFileName(idwPath)
                Dim newIdwName As String = newPrefix & "_" & idwFileName
                Dim newIdwPath As String = System.IO.Path.Combine(destFolder, newIdwName)

                Log("  [Inventor Fallback] Processing: " & idwFileName)

                Dim idwDoc As DrawingDocument = DirectCast(m_inventorApp.Documents.Open(idwPath, False), DrawingDocument)

                For Each fd As FileDescriptor In idwDoc.File.ReferencedFileDescriptors
                    Try
                        Dim refPath As String = fd.FullFileName
                        If m_copiedFiles.ContainsKey(refPath) Then
                            fd.ReplaceReference(m_copiedFiles(refPath))
                        End If
                    Catch
                    End Try
                Next

                idwDoc.SaveAs(newIdwPath, False)
                idwDoc.Close()
            Catch ex As Exception
                Log("  [Inventor Fallback] ERROR: " & ex.Message)
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Update assembly references while the document is OPEN in Inventor.
    ''' This matches Option 1 (Assembly_Renamer.vbs) flow - no "where is file?" dialogs.
    ''' </summary>
    Private Sub UpdateAssemblyReferencesInOpenDoc(asmDoc As AssemblyDocument)
        Log("  Updating references recursively in OPEN assembly...")
        UpdateReferencesInOpenDocRecursively(asmDoc, 0)
        Log("  Reference update complete")
    End Sub

    Private Sub UpdateReferencesInOpenDocRecursively(asmDoc As AssemblyDocument, level As Integer)
        Dim indent As String = New String(" "c, level * 2)

        For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            Try
                If occ.Suppressed Then
                    Log(indent & "  SKIP (suppressed): " & occ.Name)
                    Continue For
                End If

                Dim doc As Document = Nothing
                Try
                    doc = occ.Definition.Document
                Catch
                    Continue For
                End Try

                If doc Is Nothing Then Continue For

                Dim fullPath As String = doc.FullFileName
                Dim fileName As String = System.IO.Path.GetFileName(fullPath)

                If fullPath.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase) Then
                    ' Part file - check if we have a mapping
                    If m_copiedFiles.ContainsKey(fullPath) Then
                        Dim newPath As String = m_copiedFiles(fullPath)
                        Dim newFileName As String = System.IO.Path.GetFileName(newPath)

                        Log(indent & "  REPLACING: " & fileName & " -> " & newFileName)

                        ' This is the key - occ.Replace on OPEN document works silently
                        occ.Replace(newPath, True)

                        Log(indent & "  SUCCESS")
                    Else
                        Log(indent & "  NO MAPPING: " & fileName)
                    End If

                ElseIf fullPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    ' Sub-assembly - recurse into it
                    If Not fullPath.ToLower().Contains("bolted connection") Then
                        Log(indent & "  RECURSING: " & fileName)
                        Dim subAsm As AssemblyDocument = DirectCast(doc, AssemblyDocument)
                        UpdateReferencesInOpenDocRecursively(subAsm, level + 1)

                        ' Save sub-assembly after updating
                        subAsm.Save()
                        Log(indent & "  SAVED: " & fileName)
                    End If
                End If

            Catch ex As Exception
                Log(indent & "  ERROR: " & ex.Message)
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Process IDW files: Open from SOURCE, update references, SaveAs to DESTINATION.
    ''' This matches Option 9 (Assembly_Cloner.vbs) flow - avoids "non-unique project file names" dialog.
    ''' </summary>
    Private Sub ProcessIDWFilesFromSource(sourceDir As String, destFolder As String, newPrefix As String)
        Dim idwFiles() As String = System.IO.Directory.GetFiles(sourceDir, "*.idw")
        Log("  Found " & idwFiles.Length & " IDW files in source folder")

        If idwFiles.Length = 0 Then Return

        ' Suppress dialogs during IDW processing
        m_inventorApp.SilentOperation = True

        For Each idwPath As String In idwFiles
            Try
                Dim idwFileName As String = System.IO.Path.GetFileName(idwPath)
                Dim newIdwName As String = newPrefix & "_" & idwFileName
                Dim newIdwPath As String = System.IO.Path.Combine(destFolder, newIdwName)

                Log("  Processing IDW: " & idwFileName)

                ' Open the ORIGINAL IDW from SOURCE (it has valid references)
                Log("    Opening from source...")
                Dim idwDoc As DrawingDocument = DirectCast(m_inventorApp.Documents.Open(idwPath, False), DrawingDocument)
                Log("    Opened successfully")

                ' Update references to point to NEW paths in destination
                Dim fileDescriptors As FileDescriptorsEnumerator = idwDoc.File.ReferencedFileDescriptors
                Log("    Found " & fileDescriptors.Count & " references")

                Dim updatedCount As Integer = 0
                For Each fd As FileDescriptor In fileDescriptors
                    Try
                        Dim refPath As String = fd.FullFileName

                        If m_copiedFiles.ContainsKey(refPath) Then
                            Dim newRefPath As String = m_copiedFiles(refPath)
                            Log("    UPDATING: " & System.IO.Path.GetFileName(refPath) & " -> " & System.IO.Path.GetFileName(newRefPath))

                            fd.ReplaceReference(newRefPath)
                            updatedCount += 1
                        End If
                    Catch ex As Exception
                        Log("    REF ERROR: " & ex.Message)
                    End Try
                Next

                ' SaveAs to destination with new name
                Log("    Saving to: " & newIdwPath)
                idwDoc.SaveAs(newIdwPath, False)
                idwDoc.Close()

                Log("    IDW complete. Updated " & updatedCount & " references")

            Catch ex As Exception
                Log("  IDW ERROR: " & ex.Message)
            End Try
        Next

        ' Restore normal operation
        m_inventorApp.SilentOperation = False
    End Sub
End Class
