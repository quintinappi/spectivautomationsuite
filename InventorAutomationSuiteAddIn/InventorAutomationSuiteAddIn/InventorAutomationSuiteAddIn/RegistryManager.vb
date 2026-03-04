' ============================================================================
' INVENTOR AUTOMATION SUITE - REGISTRY MANAGER
' ============================================================================
' Description: Manages Windows Registry for numbering counters
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-20
' ============================================================================

Imports System
Imports Microsoft.Win32
Imports System.Collections.Generic

Namespace InventorAutomationSuiteAddIn

    ''' <summary>
    ''' Manages registry operations for the add-in
    ''' </summary>
    Public Class RegistryManager

        ' Constants
        Private Const REGISTRY_BASE_PATH As String = "HKEY_CURRENT_USER\Software\InventorRenamer\"

        ''' <summary>
        ''' Scan registry for counters with given prefix
        ''' </summary>
        Public Function ScanCounters(prefix As String) As Dictionary(Of String, Integer)
            Try
                Dim counters As New Dictionary(Of String, Integer)()

                ' Ensure prefix ends with dash
                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                ' Define counter groups
                Dim groups As String() = {"PL", "B", "CH", "A", "FL", "LPL", "P", "SQ", "R", "FLG", "IPE"}

                ' Read each counter from registry
                For Each group As String In groups
                    Dim regKey As String = REGISTRY_BASE_PATH & prefix & group
                    Dim value As Integer = 0

                    Try
                        Dim regValue As Object = Registry.GetValue(regKey, Nothing, 0)
                        If regValue IsNot Nothing Then
                            value = Convert.ToInt32(regValue)
                        End If
                    Catch ex As Exception
                        ' Key doesn't exist, use default value of 0
                        value = 0
                    End Try

                    counters(group) = value
                Next

                Return counters

            Catch ex As Exception
                Throw New Exception("Error scanning registry: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Get a specific counter value
        ''' </summary>
        Public Function GetCounter(prefix As String, group As String) As Integer
            Try
                ' Ensure prefix ends with dash
                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                Dim regKey As String = REGISTRY_BASE_PATH & prefix & group
                Dim value As Integer = 0

                Try
                    Dim regValue As Object = Registry.GetValue(regKey, Nothing, 0)
                    If regValue IsNot Nothing Then
                        value = Convert.ToInt32(regValue)
                    End If
                Catch ex As Exception
                    ' Key doesn't exist, use default value of 0
                    value = 0
                End Try

                Return value

            Catch ex As Exception
                Throw New Exception("Error getting counter: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Set a specific counter value
        ''' </summary>
        Public Sub SetCounter(prefix As String, group As String, value As Integer)
            Try
                ' Ensure prefix ends with dash
                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                Dim regKey As String = REGISTRY_BASE_PATH & prefix & group

                ' Write value to registry
                Registry.SetValue(regKey, value, RegistryValueKind.DWord)

            Catch ex As Exception
                Throw New Exception("Error setting counter: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Increment a counter and return the new value
        ''' </summary>
        Public Function IncrementCounter(prefix As String, group As String) As Integer
            Try
                Dim currentValue As Integer = GetCounter(prefix, group)
                Dim newValue As Integer = currentValue + 1

                SetCounter(prefix, group, newValue)

                Return newValue

            Catch ex As Exception
                Throw New Exception("Error incrementing counter: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Clear all counters for a given prefix
        ''' </summary>
        Public Sub ClearCounters(prefix As String)
            Try
                ' Ensure prefix ends with dash
                If Right(prefix, 1) <> "-" Then
                    prefix = prefix & "-"
                End If

                ' Define counter groups
                Dim groups As String() = {"PL", "B", "CH", "A", "FL", "LPL", "P", "SQ", "R", "FLG", "IPE"}

                ' Delete each counter from registry
                For Each group As String In groups
                    Try
                        Dim regKey As String = "HKEY_CURRENT_USER\Software\InventorRenamer\" & prefix & group
                        Registry.CurrentUser.DeleteValue("Software\InventorRenamer\" & prefix & group, False)
                    Catch ex As Exception
                        ' Key doesn't exist, ignore
                    End Try
                Next

            Catch ex As Exception
                Throw New Exception("Error clearing counters: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get all prefixes in registry
        ''' </summary>
        Public Function GetAllPrefixes() As List(Of String)
            Try
                Dim prefixes As New List(Of String)()

                ' Open the base registry key
                Dim baseKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\InventorRenamer")

                If baseKey IsNot Nothing Then
                    ' Get all value names
                    Dim valueNames As String() = baseKey.GetValueNames()

                    ' Extract unique prefixes
                    Dim prefixSet As New HashSet(Of String)()

                    For Each valueName As String In valueNames
                        ' Parse prefix from value name (format: PREFIX-GROUP)
                        Dim lastDashIndex As Integer = valueName.LastIndexOf("-")

                        If lastDashIndex > 0 Then
                            Dim prefix As String = valueName.Substring(0, lastDashIndex + 1)
                            prefixSet.Add(prefix)
                        End If
                    Next

                    prefixes = New List(Of String)(prefixSet)
                    prefixes.Sort()

                    baseKey.Close()
                End If

                Return prefixes

            Catch ex As Exception
                Throw New Exception("Error getting prefixes: " & ex.Message)
            End Try
        End Function

    End Class

End Namespace
