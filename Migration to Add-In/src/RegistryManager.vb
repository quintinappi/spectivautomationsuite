' ============================================================================
' INVENTOR AUTOMATION SUITE - REGISTRY MANAGER
' ============================================================================
' Description: Windows Registry operations for part numbering counters
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' Ported from: Registry_Manager.vbs
' Functions: ScanRegistryForCounters, SaveCounterToRegistry,
'            CheckIfPrefixExistsInRegistry
'
' Registry Path: HKEY_CURRENT_USER\Software\InventorRenamer\
' Format: PREFIX-GROUP = Number (e.g., NCRH01-000-PL = 173)
' ============================================================================

Imports System
Imports Microsoft.Win32

Namespace SpectivInventorSuite

    ''' <summary>
    ''' Windows Registry manager for part numbering counters
    ''' </summary>
    Public Class RegistryManager

        ' Registry constants
        Private Const REGISTRY_ROOT As String = "HKEY_CURRENT_USER\Software\InventorRenamer\"
        Private Const COMMON_GROUPS As String() = {"PL", "B", "CH", "A", "FL", "LPL", "SQ", "P", "IPE", "R", "FLG"}

        ''' <summary>
        ''' Scan registry for counters matching a prefix
        ''' VBScript: ScanRegistryForCounters()
        ''' </summary>
        ''' <param name="prefix">Prefix to scan for (e.g., "NCRH01-000-")</param>
        ''' <returns>Dictionary of group -> counter value</returns>
        Public Function ScanCounters(prefix As String) As Dictionary(Of String, Integer)
            Dim counters As New Dictionary(Of String, Integer)()

            Try
                ' Ensure prefix ends with dash
                If Not String.IsNullOrEmpty(prefix) AndAlso Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If

                ' Open registry key
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer", False)
                    If regKey Is Nothing Then
                        ' Registry path doesn't exist yet
                        Return counters
                    End If

                    ' Get all value names
                    Dim valueNames As String() = regKey.GetValueNames()

                    ' Filter by prefix if provided
                    For Each valueName As String In valueNames
                        If String.IsNullOrEmpty(prefix) OrElse valueName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                            Dim value As Object = regKey.GetValue(valueName)
                            If TypeOf value Is Integer Then
                                counters(valueName) = CInt(value)
                            End If
                        End If
                    Next
                End Using

            Catch ex As Exception
                ' Return empty dictionary on error
            End Try

            Return counters
        End Function

        ''' <summary>
        ''' Save counter to registry
        ''' VBScript: SaveCounterToRegistry()
        ''' </summary>
        ''' <param name="prefixGroupKey">Full key (e.g., "NCRH01-000-PL")</param>
        ''' <param name="value">Counter value to save</param>
        Public Sub SaveCounter(prefixGroupKey As String, value As Integer)
            Try
                ' Ensure key ends with dash before group
                If Not prefixGroupKey.Contains("-") Then
                    Throw New ArgumentException("Invalid prefix group key format")
                End If

                ' Open or create registry key
                Using regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\InventorRenamer")
                    regKey.SetValue(prefixGroupKey, value, RegistryValueKind.DWord)
                End Using

            Catch ex As Exception
                Throw New Exception("Failed to save counter to registry: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Check if a prefix exists in registry
        ''' VBScript: CheckIfPrefixExistsInRegistry()
        ''' </summary>
        ''' <param name="prefix">Prefix to check</param>
        ''' <returns>True if prefix has any counters in registry</returns>
        Public Function PrefixExists(prefix As String) As Boolean
            Try
                ' Ensure prefix ends with dash
                If Not String.IsNullOrEmpty(prefix) AndAlso Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If

                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer", False)
                    If regKey Is Nothing Then
                        Return False
                    End If

                    ' Check if any values start with this prefix
                    For Each valueName As String In regKey.GetValueNames()
                        If valueName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                            Return True
                        End If
                    Next
                End Using

                Return False

            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Get counter value for a specific prefix and group
        ''' </summary>
        ''' <param name="prefix">Prefix (e.g., "NCRH01-000-")</param>
        ''' <param name="group">Group code (e.g., "PL")</param>
        ''' <returns>Current counter value, or 0 if not found</returns>
        Public Function GetCounter(prefix As String, group As String) As Integer
            Try
                ' Ensure prefix ends with dash
                If Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If

                Dim fullKey As String = prefix & group

                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer", False)
                    If regKey Is Nothing Then
                        Return 0
                    End If

                    Dim value As Object = regKey.GetValue(fullKey)
                    If TypeOf value Is Integer Then
                        Return CInt(value)
                    End If
                End Using

            Catch ex As Exception
            End Try

            Return 0
        End Function

        ''' <summary>
        ''' Set counter value for a specific prefix and group
        ''' </summary>
        ''' <param name="prefix">Prefix (e.g., "NCRH01-000-")</param>
        ''' <param name="group">Group code (e.g., "PL")</param>
        ''' <param name="value">Counter value to set</param>
        Public Sub SetCounter(prefix As String, group As String, value As Integer)
            Try
                ' Ensure prefix ends with dash
                If Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If

                Dim fullKey As String = prefix & group

                Using regKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\InventorRenamer")
                    regKey.SetValue(fullKey, value, RegistryValueKind.DWord)
                End Using

            Catch ex As Exception
                Throw New Exception("Failed to set counter: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Increment counter and return new value
        ''' </summary>
        ''' <param name="prefix">Prefix (e.g., "NCRH01-000-")</param>
        ''' <param name="group">Group code (e.g., "PL")</param>
        ''' <returns>New counter value after increment</returns>
        Public Function IncrementCounter(prefix As String, group As String) As Integer
            Dim current As Integer = GetCounter(prefix, group)
            Dim newValue As Integer = current + 1
            SetCounter(prefix, group, newValue)
            Return newValue
        End Function

        ''' <summary>
        ''' Clear all counters for a specific prefix
        ''' </summary>
        ''' <param name="prefix">Prefix to clear</param>
        ''' <returns>Number of counters cleared</returns>
        Public Function ClearPrefix(prefix As String) As Integer
            Dim clearedCount As Integer = 0

            Try
                ' Ensure prefix ends with dash
                If Not String.IsNullOrEmpty(prefix) AndAlso Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If

                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer", True)
                    If regKey Is Nothing Then
                        Return 0
                    End If

                    ' Get all value names
                    Dim valueNames As String() = regKey.GetValueNames()

                    ' Delete values matching prefix
                    For Each valueName As String In valueNames
                        If valueName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                            regKey.DeleteValue(valueName)
                            clearedCount += 1
                        End If
                    Next
                End Using

            Catch ex As Exception
                Throw New Exception("Failed to clear prefix: " & ex.Message)
            End Try

            Return clearedCount
        End Function

        ''' <summary>
        ''' Get all counters as formatted string for display
        ''' </summary>
        ''' <param name="prefix">Prefix to filter by (empty = show all)</param>
        ''' <returns>Formatted string of all counters</returns>
        Public Function GetCountersDisplay(prefix As String) As String
            Dim counters As Dictionary(Of String, Integer) = ScanCounters(prefix)
            Dim result As New System.Text.StringBuilder()

            If String.IsNullOrEmpty(prefix) Then
                result.AppendLine("ALL REGISTRY ENTRIES:")
                result.AppendLine(new String("="c, 50))
                result.AppendLine()
            Else
                If Not prefix.EndsWith("-") Then
                    prefix = prefix & "-"
                End If
                result.AppendLine("REGISTRY ENTRIES FOR PREFIX: " & prefix)
                result.AppendLine(new String("="c, 50))
                result.AppendLine()
            End If

            If counters.Count = 0 Then
                result.AppendLine("No counters found - database is empty.")
            Else
                result.AppendLine("Found " & counters.Count & " counter(s):")
                result.AppendLine()

                For Each kvp As KeyValuePair(Of String, Integer) In counters
                    result.AppendLine(kvp.Key & " = " & kvp.Value)
                Next
            End If

            Return result.ToString()
        End Function

        ''' <summary>
        ''' Get list of common group codes
        ''' </summary>
        Public Shared Function GetCommonGroups() As String()
            Return DirectCast(COMMON_GROUPS.Clone(), String())
        End Function

    End Class

End Namespace
