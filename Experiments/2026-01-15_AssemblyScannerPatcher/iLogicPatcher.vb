Imports Inventor
Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualBasic

Public Class iLogicPatcher
    Private m_inventorApp As Inventor.Application
    Private m_log As System.Text.StringBuilder
    Private Const ILOGIC_ADDIN_GUID As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

    Public Sub New(inventorApp As Inventor.Application, log As System.Text.StringBuilder)
        m_inventorApp = inventorApp
        m_log = log
    End Sub

    Private Sub Log(message As String)
        If m_log IsNot Nothing Then
            Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss")
            m_log.AppendLine(timestamp & " | " & message)
        End If
    End Sub

    Public Function PatchRulesInAssembly(asmDoc As AssemblyDocument, partNameMapping As Dictionary(Of String, String)) As Integer
        Dim totalReplacements As Integer = 0

        Log("  === iLogic Patcher Started ===")
        Log("  Part name mappings to apply: " & partNameMapping.Count)

        If partNameMapping.Count = 0 Then
            Log("  WARNING: Part name mapping is EMPTY!")
            Return 0
        End If

        ' Log all mappings for debugging
        For Each kvp As KeyValuePair(Of String, String) In partNameMapping
            Log("    MAPPING: """ & kvp.Key & """ -> """ & kvp.Value & """")
        Next

        Try
            ' Get iLogic add-in
            Log("  Getting iLogic add-in...")
            Dim iLogicAddIn As ApplicationAddIn = Nothing
            Try
                iLogicAddIn = m_inventorApp.ApplicationAddIns.ItemById(ILOGIC_ADDIN_GUID)
            Catch ex As Exception
                Log("  ERROR getting iLogic add-in: " & ex.Message)
            End Try

            If iLogicAddIn Is Nothing Then
                Log("  ERROR: iLogic Add-In not found!")
                Return 0
            End If
            Log("  iLogic add-in found: " & iLogicAddIn.DisplayName)

            ' Get iLogic automation
            Dim iLogicAuto As Object = iLogicAddIn.Automation
            If iLogicAuto Is Nothing Then
                Log("  ERROR: Failed to get iLogic automation interface!")
                Return 0
            End If
            Log("  iLogic automation interface obtained")

            ' Patch rules in the main assembly
            Log("  Processing main assembly rules...")
            Dim mainRuleCount As Integer = 0
            Dim mainReplacements As Integer = PatchRulesInDocument(iLogicAuto, asmDoc, partNameMapping, mainRuleCount)
            totalReplacements += mainReplacements
            Log("  Main assembly: " & mainRuleCount & " rules, " & mainReplacements & " replacements")

            ' Patch rules in all referenced documents
            Log("  Processing referenced documents...")
            Dim refDocCount As Integer = 0
            For Each refDoc As Document In asmDoc.AllReferencedDocuments
                If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Or
                   refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    Dim docRuleCount As Integer = 0
                    Dim docReplacements As Integer = PatchRulesInDocument(iLogicAuto, refDoc, partNameMapping, docRuleCount)
                    If docRuleCount > 0 Then
                        Log("    " & System.IO.Path.GetFileName(refDoc.FullFileName) & ": " & docRuleCount & " rules, " & docReplacements & " replacements")
                        refDocCount += 1
                        totalReplacements += docReplacements
                    End If
                End If
            Next
            Log("  Referenced documents with rules: " & refDocCount)

        Catch ex As Exception
            Log("  EXCEPTION in PatchRulesInAssembly: " & ex.Message)
            Log("  Stack trace: " & ex.StackTrace)
        End Try

        Log("  === iLogic Patcher Complete. Total replacements: " & totalReplacements & " ===")
        Return totalReplacements
    End Function

    Private Function PatchRulesInDocument(iLogicAuto As Object, doc As Document, partNameMapping As Dictionary(Of String, String), ByRef ruleCount As Integer) As Integer
        Dim replacements As Integer = 0
        ruleCount = 0

        Try
            ' Get rules collection
            Dim rules As Object = Nothing
            Try
                rules = iLogicAuto.Rules(doc)
            Catch ex As Exception
                ' No rules in this document
                Return 0
            End Try

            If rules Is Nothing Then
                Return 0
            End If

            ' Get count
            Try
                ruleCount = CInt(rules.Count)
            Catch
                Return 0
            End Try

            If ruleCount = 0 Then
                Return 0
            End If

            ' Iterate through rules
            For i As Integer = 1 To ruleCount
                Try
                    Dim rule As Object = rules.Item(i)
                    Dim ruleName As String = ""
                    Try
                        ruleName = CStr(rule.Name)
                    Catch
                    End Try

                    Dim originalCode As String = CStr(rule.Text)
                    Dim localReplacements As Integer = 0
                    Dim patchedCode As String = PatchCode(originalCode, partNameMapping, localReplacements, ruleName)

                    If localReplacements > 0 Then
                        Log("      Rule '" & ruleName & "': " & localReplacements & " replacements")
                        rule.Text = patchedCode
                        replacements += localReplacements
                    End If
                Catch ruleEx As Exception
                    Log("      ERROR processing rule " & i & ": " & ruleEx.Message)
                End Try
            Next

        Catch ex As Exception
            Log("    ERROR in PatchRulesInDocument: " & ex.Message)
        End Try

        Return replacements
    End Function

    Private Function PatchCode(code As String, partNameMapping As Dictionary(Of String, String), ByRef replacementCount As Integer, ruleName As String) As String
        Dim result As String = code
        replacementCount = 0

        ' Log the first 500 chars of rule code for debugging
        Dim cleanCode As String = code.Replace(vbCrLf, " ").Replace(vbLf, " ")
        Dim previewLen As Integer = Math.Min(500, cleanCode.Length)
        Log("        RULE CODE PREVIEW: " & cleanCode.Substring(0, previewLen))

        For Each kvp As KeyValuePair(Of String, String) In partNameMapping
            Dim oldBaseName As String = kvp.Key
            Dim newBaseName As String = kvp.Value

            ' PATTERN 1: Replace occurrences with numbers :1 through :50
            ' e.g., "Part1 TFC S-2:1" -> "NewAsm_Part1 TFC S-2:1"
            For occNum As Integer = 1 To 50
                Dim oldPattern As String = """" & oldBaseName & ":" & occNum.ToString() & """"
                Dim newPattern As String = """" & newBaseName & ":" & occNum.ToString() & """"

                If result.Contains(oldPattern) Then
                    Dim countBefore As Integer = CountOccurrences(result, oldPattern)
                    result = result.Replace(oldPattern, newPattern)
                    replacementCount += countBefore
                    Log("        REPLACED (with colon): " & oldPattern & " -> " & newPattern)
                End If
            Next

            ' PATTERN 2: Replace base name without occurrence number
            ' e.g., "Part1 TFC S-2" -> "NewAsm_Part1 TFC S-2"
            ' This catches cases where the occurrence number might not be included
            Dim oldBasePattern As String = """" & oldBaseName & """"
            Dim newBasePattern As String = """" & newBaseName & """"

            If result.Contains(oldBasePattern) Then
                Dim countBefore As Integer = CountOccurrences(result, oldBasePattern)
                result = result.Replace(oldBasePattern, newBasePattern)
                replacementCount += countBefore
                Log("        REPLACED (base only): " & oldBasePattern & " -> " & newBasePattern)
            End If

            ' PATTERN 3: Check if the old name appears anywhere (for debugging)
            If result.Contains(oldBaseName) Then
                Log("        FOUND (unquoted): '" & oldBaseName & "' still exists in rule - may need different pattern")
            End If
        Next

        Return result
    End Function

    Private Function CountOccurrences(text As String, pattern As String) As Integer
        Dim count As Integer = 0
        Dim index As Integer = 0

        While True
            index = text.IndexOf(pattern, index)
            If index < 0 Then Exit While
            count += 1
            index += pattern.Length
        End While

        Return count
    End Function
End Class
