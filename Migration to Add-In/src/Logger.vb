' ============================================================================
' INVENTOR AUTOMATION SUITE - LOGGER
' ============================================================================
' Description: Logging system for Assembly Cloner
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' Ported from: Assembly_Cloner.vbs (StartLogging, LogMessage, StopLogging)
' ============================================================================

Imports System
Imports System.IO
Imports System.Text

Namespace SpectivInventorSuite

    ''' <summary>
    ''' Thread-safe logger for Assembly Cloner operations
    ''' </summary>
    Public Class Logger

        ' Private members
        Private m_logPath As String
        Private m_logWriter As StreamWriter
        Private m_lock As New Object()
        Private m_isInitialized As Boolean = False

        ''' <summary>
        ''' Constructor
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Initialize logger with default path
        ''' </summary>
        Public Sub Initialize()
            Initialize(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Assembly_Cloner_Log.txt"))
        End Sub

        ''' <summary>
        ''' Initialize logger with specific path
        ''' </summary>
        Public Sub Initialize(logPath As String)
            Try
                SyncLock m_lock
                    ' Close existing writer if open
                    If m_logWriter IsNot Nothing Then
                        m_logWriter.Close()
                        m_logWriter.Dispose()
                    End If

                    ' Store path
                    m_logPath = logPath

                    ' Create new log file (overwrite existing)
                    m_logWriter = New StreamWriter(m_logPath, False, Encoding.UTF8)
                    m_logWriter.AutoFlush = True

                    m_isInitialized = True

                    ' Write header
                    Log("========================================")
                    Log("ASSEMBLY CLONER - LOG STARTED")
                    Log("Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    Log("Log File: " & m_logPath)
                    Log("========================================")
                End SyncLock

            Catch ex As Exception
                Throw New Exception("Failed to initialize logger: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Log a message with timestamp
        ''' </summary>
        Public Sub Log(message As String)
            Try
                SyncLock m_lock
                    If Not m_isInitialized OrElse m_logWriter Is Nothing Then
                        ' Fallback to console if not initialized
                        Console.WriteLine(message)
                        Return
                    End If

                    Dim logEntry As String = DateTime.Now.ToString("HH:mm:ss") & " - " & message
                    m_logWriter.WriteLine(logEntry)
                End SyncLock
            Catch ex As Exception
                ' Silently fail - logging errors shouldn't break the app
                Console.WriteLine("Logging error: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Log an error message
        ''' </summary>
        Public Sub LogError(message As String)
            Log("ERROR: " & message)
        End Sub

        ''' <summary>
        ''' Log a warning message
        ''' </summary>
        Public Sub LogWarning(message As String)
            Log("WARNING: " & message)
        End Sub

        ''' <summary>
        ''' Log an info message
        ''' </summary>
        Public Sub LogInfo(message As String)
            Log("INFO: " & message)
        End Sub

        ''' <summary>
        ''' Log a success message
        ''' </summary>
        Public Sub LogSuccess(message As String)
            Log("SUCCESS: " & message)
        End Sub

        ''' <summary>
        ''' Close logger and flush all pending writes
        ''' </summary>
        Public Sub Close()
            Try
                SyncLock m_lock
                    If m_logWriter IsNot Nothing Then
                        Log("========================================")
                        Log("ASSEMBLY CLONER - LOG ENDED")
                        Log("Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                        Log("========================================")
                        Log("")

                        m_logWriter.Close()
                        m_logWriter.Dispose()
                        m_logWriter = Nothing
                    End If

                    m_isInitialized = False
                End SyncLock
            Catch ex As Exception
                Console.WriteLine("Error closing logger: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get the log file path
        ''' </summary>
        Public ReadOnly Property LogPath As String
            Get
                Return m_logPath
            End Get
        End Property

        ''' <summary>
        ''' Check if logger is initialized
        ''' </summary>
        Public ReadOnly Property IsInitialized As Boolean
            Get
                Return m_isInitialized
            End Get
        End Property

    End Class

End Namespace
