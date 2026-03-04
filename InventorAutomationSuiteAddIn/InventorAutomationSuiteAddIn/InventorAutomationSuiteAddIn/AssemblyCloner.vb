' ============================================================================
' INVENTOR AUTOMATION SUITE - ASSEMBLY CLONER
' ============================================================================
' Description: Core logic for cloning assemblies with numbering continuation
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-20
' ============================================================================

Imports System
Imports System.IO
Imports Inventor

Namespace InventorAutomationSuiteAddIn

    ''' <summary>
    ''' Core assembly cloning logic
    ''' </summary>
    Public Class AssemblyCloner

        ' Private members
        Private m_invApp As InventorApplication
        Private m_regManager As RegistryManager
        Private m_logPath As String

        ' Progress callback
        Public Delegate Sub ProgressCallback(percent As Integer, status As String)
        Private m_progressCallback As ProgressCallback

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New(invApp As InventorApplication)
            Try
                m_invApp = invApp
                m_regManager = New RegistryManager()

                ' Set log file path
                m_logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), _
                                        "Assembly_Cloner_Log.txt")

                ' Clear log file
                File.WriteAllText(m_logPath, String.Empty)

            Catch ex As Exception
                Throw New Exception("Error initializing AssemblyCloner: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Clone assembly with given prefix and count
        ''' </summary>
        Public Sub Clone(prefix As String, cloneCount As Integer, progressCallback As ProgressCallback)
            Try
                m_progressCallback = progressCallback

                LogMessage("==========================================")
                LogMessage("ASSEMBLY CLONER - STARTING")
                LogMessage("==========================================")
                LogMessage("Prefix: " & prefix)
                LogMessage("Clone Count: " & cloneCount)

                ' Step 1: Validate active document
                UpdateProgress(5, "Validating active document...")
                If Not ValidateActiveDocument() Then
                    Throw New Exception("No active assembly document found")
                End If

                Dim asmDoc As AssemblyDocument = DirectCast(m_invApp.ActiveDocument, AssemblyDocument)
                LogMessage("Active Document: " & asmDoc.FullFileName)

                ' Step 2: Scan current registry
                UpdateProgress(10, "Scanning registry...")
                ScanRegistry(prefix)

                ' Step 3: Get original assembly info
                UpdateProgress(15, "Analyzing assembly structure...")
                Dim originalPath As String = asmDoc.FullFileName
                Dim originalDir As String = Path.GetDirectoryName(originalPath)
                Dim originalName As String = Path.GetFileName(originalPath)
                Dim originalBaseName As String = Path.GetFileNameWithoutExtension(originalPath)

                LogMessage("Original Assembly:")
                LogMessage("  Path: " & originalPath)
                LogMessage("  Directory: " & originalDir)
                LogMessage("  Name: " & originalName)

                ' Step 4: Create clones
                For i As Integer = 1 To cloneCount
                    Dim percent As Integer = 15 + CInt((70 * i) / cloneCount)
                    UpdateProgress(percent, "Creating clone " & i & " of " & cloneCount & "...")

                    LogMessage("==========================================")
                    LogMessage("CREATING CLONE #" & i & " OF " & cloneCount)
                    LogMessage("==========================================")

                    CreateSingleClone(i, originalDir, originalBaseName, prefix)
                Next

                ' Step 5: Finalize
                UpdateProgress(90, "Finalizing...")
                LogMessage("==========================================")
                LogMessage("ASSEMBLY CLONER - COMPLETED SUCCESSFULLY")
                LogMessage("==========================================")

                UpdateProgress(100, "Completed successfully!")

            Catch ex As Exception
                LogMessage("ERROR: " & ex.Message)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Validate that we have an active assembly document
        ''' </summary>
        Private Function ValidateActiveDocument() As Boolean
            Try
                If m_invApp.ActiveDocument Is Nothing Then
                    Return False
                End If

                If m_invApp.ActiveDocument.Type <> DocumentTypeEnum.kAssemblyDocumentObject Then
                    Return False
                End If

                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Scan registry and display current counters
        ''' </summary>
        Private Sub ScanRegistry(prefix As String)
            Try
                LogMessage("==========================================")
                LogMessage("SCANNING REGISTRY FOR PREFIX: " & prefix)
                LogMessage("==========================================")

                Dim counters As System.Collections.Generic.Dictionary(Of String, Integer) = _
                    m_regManager.ScanCounters(prefix)

                ' Log top 5 counters
                Dim topGroups As String() = {"PL", "B", "CH", "A", "FL"}

                For Each group As String In topGroups
                    If counters.ContainsKey(group) Then
                        LogMessage("  " & group & ": " & counters(group))
                    End If
                Next

                LogMessage("Registry scan complete")

            Catch ex As Exception
                LogMessage("ERROR scanning registry: " & ex.Message)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Create a single clone
        ''' </summary>
        Private Sub CreateSingleClone(cloneNumber As Integer, originalDir As String, originalBaseName As String, prefix As String)
            Try
                Dim cloneName As String = originalBaseName & "_CLONE" & cloneNumber & ".iam"
                Dim clonePath As String = Path.Combine(originalDir, cloneName)

                LogMessage("Creating clone: " & cloneName)
                LogMessage("  Path: " & clonePath)

                ' Check if file already exists
                If File.Exists(clonePath) Then
                    LogMessage("  WARNING: File already exists, skipping")
                    Throw New Exception("Clone file already exists: " & clonePath)
                End If

                ' Get reference to active document
                Dim asmDoc As AssemblyDocument = DirectCast(m_invApp.ActiveDocument, AssemblyDocument)

                ' Save assembly as clone
                asmDoc.SaveAs(clonePath, False)

                LogMessage("  Clone created successfully")

                ' NOTE: The actual part cloning and renaming logic will be implemented
                ' This is a placeholder that creates the assembly clone only
                LogMessage("  NOTE: Full part cloning logic to be implemented in next phase")
                LogMessage("  Currently creating assembly clone only")

            Catch ex As Exception
                LogMessage("ERROR creating clone: " & ex.Message)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Update progress
        ''' </summary>
        Private Sub UpdateProgress(percent As Integer, status As String)
            Try
                If m_progressCallback IsNot Nothing Then
                    m_progressCallback(percent, status)
                End If
            Catch ex As Exception
                ' Ignore progress update errors
            End Try
        End Sub

        ''' <summary>
        ''' Log message to file
        ''' </summary>
        Private Sub LogMessage(message As String)
            Try
                Dim logMessage As String = DateTime.Now.ToString() & " - " & message
                File.AppendAllText(m_logPath, logMessage & vbCrLf)
            Catch ex As Exception
                ' Silently fail - logging is not critical
            End Try
        End Sub

    End Class

End Namespace
