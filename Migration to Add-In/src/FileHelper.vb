' ============================================================================
' INVENTOR AUTOMATION SUITE - FILE HELPER
' ============================================================================
' Description: File and folder operations helper
' Author: Spectiv Solutions
' Version: 1.0.0
' Date: 2025-01-21
'
' Ported from: Assembly_Cloner.vbs utility functions
' Functions: GetFileNameFromPath, CreateFolderRecursive, IncrementFileName,
'            BuildFileInventory
' ============================================================================

Imports System
Imports System.IO
Imports System.Collections.Generic

Namespace SpectivInventorSuite

    ''' <summary>
    ''' File and folder operations utility class
    ''' </summary>
    Public Class FileHelper

        ''' <summary>
        ''' Extract filename from full path
        ''' VBScript: GetFileNameFromPath()
        ''' </summary>
        Public Shared Function GetFileNameFromPath(fullPath As String) As String
            If String.IsNullOrEmpty(fullPath) Then
                Return String.Empty
            End If

            Try
                Return Path.GetFileName(fullPath)
            Catch ex As Exception
                Return fullPath
            End Try
        End Function

        ''' <summary>
        ''' Get filename without extension
        ''' </summary>
        Public Shared Function GetFileNameWithoutExtension(fullPath As String) As String
            If String.IsNullOrEmpty(fullPath) Then
                Return String.Empty
            End If

            Try
                Return Path.GetFileNameWithoutExtension(fullPath)
            Catch ex As Exception
                Return fullPath
            End Try
        End Function

        ''' <summary>
        ''' Get directory from full path
        ''' </summary>
        Public Shared Function GetDirectoryFromPath(fullPath As String) As String
            If String.IsNullOrEmpty(fullPath) Then
                Return String.Empty
            End If

            Try
                Return Path.GetDirectoryName(fullPath)
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Create folder recursively (creates all parent directories if needed)
        ''' VBScript: CreateFolderRecursive()
        ''' </summary>
        Public Shared Sub CreateFolderRecursive(folderPath As String)
            If String.IsNullOrWhiteSpace(folderPath) Then
                Throw New ArgumentException("Folder path cannot be empty")
            End If

            Try
                If Not Directory.Exists(folderPath) Then
                    Directory.CreateDirectory(folderPath)
                End If
            Catch ex As Exception
                Throw New Exception("Failed to create folder: " & folderPath & " - " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Increment filename if file exists (e.g., file.iam -> file (2).iam)
        ''' VBScript: IncrementFileName()
        ''' </summary>
        Public Shared Function IncrementFileName(fileName As String) As String
            If String.IsNullOrEmpty(fileName) Then
                Return fileName
            End If

            Try
                Dim dir As String = Path.GetDirectoryName(fileName)
                Dim name As String = Path.GetFileNameWithoutExtension(fileName)
                Dim ext As String = Path.GetExtension(fileName)

                ' Check if already has increment pattern
                Dim incrementPattern As String = " (\d+)$"
                Dim match As System.Text.RegularExpressions.Match = _
                    System.Text.RegularExpressions.Regex.Match(name, incrementPattern)

                If match.Success Then
                    ' Increment existing number
                    Dim currentNum As Integer = Integer.Parse(match.Groups(1).Value)
                    Dim newName As String = System.Text.RegularExpressions.Regex.Replace(name, incrementPattern, " (" & (currentNum + 1) & ")")
                    Return Path.Combine(dir, newName & ext)
                Else
                    ' Add increment
                    Return Path.Combine(dir, name & " (2)" & ext)
                End If

            Catch ex As Exception
                ' If increment fails, return original
                Return fileName
            End Try
        End Function

        ''' <summary>
        ''' Build file inventory - count files by type in a folder
        ''' VBScript: BuildFileInventory()
        ''' </summary>
        Public Shared Function BuildFileInventory(folderPath As String) As Dictionary(Of String, Integer)
            Dim inventory As New Dictionary(Of String, Integer)()

            If Not Directory.Exists(folderPath) Then
                Return inventory
            End Try

            Try
                Dim files As String() = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)

                For Each file As String In files
                    Dim ext As String = Path.GetExtension(file).ToLower()

                    If inventory.ContainsKey(ext) Then
                        inventory(ext) += 1
                    Else
                        inventory(ext) = 1
                    End If
                Next

            Catch ex As Exception
                ' Return whatever we collected
            End Try

            Return inventory
        End Function

        ''' <summary>
        ''' Copy file with overwrite, creating target directory if needed
        ''' </summary>
        Public Shared Sub CopyFileSafe(sourcePath As String, destPath As Boolean)
            Try
                ' Ensure destination directory exists
                Dim destDir As String = Path.GetDirectoryName(destPath)
                If Not Directory.Exists(destDir) Then
                    Directory.CreateDirectory(destDir)
                End If

                ' Copy file
                File.Copy(sourcePath, destPath, True)

            Catch ex As Exception
                Throw New Exception("Failed to copy file: " & sourcePath & " -> " & destPath & " - " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Check if a file is an Inventor part file
        ''' </summary>
        Public Shared Function IsPartFile(filePath As String) As Boolean
            Return String.Equals(Path.GetExtension(filePath), ".ipt", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Check if a file is an Inventor assembly file
        ''' </summary>
        Public Shared Function IsAssemblyFile(filePath As String) As Boolean
            Return String.Equals(Path.GetExtension(filePath), ".iam", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Check if a file is an Inventor drawing file
        ''' </summary>
        Public Shared Function IsDrawingFile(filePath As String) As Boolean
            Return String.Equals(Path.GetExtension(filePath), ".idw", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(Path.GetExtension(filePath), ".dwg", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Get relative path from base path
        ''' </summary>
        Public Shared Function GetRelativePath(basePath As String, fullPath As String) As String
            Try
                Dim baseUri As New Uri(basePath & "\")
                Dim fullUri As New Uri(fullPath)
                Dim relativeUri As Uri = baseUri.MakeRelativeUri(fullUri)
                Return Uri.UnescapeDataString(relativeUri.ToString().Replace("/", "\"))
            Catch ex As Exception
                Return fullPath
            End Try
        End Function

        ''' <summary>
        ''' Combine paths safely
        ''' </summary>
        Public Shared Function CombinePaths(paramArray paths As String()) As String
            If paths Is Nothing OrElse paths.Length = 0 Then
                Return String.Empty
            End If

            Return Path.Combine(paths)
        End Function

    End Class

End Namespace
