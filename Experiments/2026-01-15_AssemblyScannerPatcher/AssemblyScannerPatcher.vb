Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Reflection

' AssemblyScannerPatcher - Unified tool for scanning, patching, and comparing Inventor assemblies
' Created: 2026-01-15
' Purpose: 
'   - SCAN: Extract iLogic rules, detect derived parts, analyze assembly structure
'   - PATCH: Apply mapping file to patch iLogic rules using iLogicPatcher
'   - COMPARE: Generate comparison report between before/after states

Module AssemblyScannerPatcher
    Const ILOGIC_ADDIN_GUID As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"

    Sub Main(args As String())
        ' If no args or just "SCAN", auto-detect open assembly
        If args.Length = 0 OrElse (args.Length = 1 AndAlso args(0).ToUpper() = "SCAN") Then
            ' Auto-detect mode
            Console.WriteLine("=== AUTO-DETECT MODE ===")
            Console.WriteLine("Detecting open assembly in Inventor...")

            Try
                Dim inventorApp As Object = GetInventorApp()
                If inventorApp Is Nothing Then
                    Console.WriteLine("ERROR: Inventor not running!")
                    Return
                End If

                Dim activeDoc As Object = inventorApp.ActiveDocument
                If activeDoc Is Nothing Then
                    Console.WriteLine("ERROR: No document open in Inventor!")
                    Return
                End If

                Dim assemblyPath As String = CStr(activeDoc.FullFileName)
                If assemblyPath.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) Then
                    Console.WriteLine("DETECTED: " & activeDoc.DisplayName)
                    Console.WriteLine("PATH: " & assemblyPath)

                    Dim outputFolder As String = Path.GetDirectoryName(assemblyPath) & "\Scan_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    ScanAssembly(assemblyPath, outputFolder)
                Else
                    Console.WriteLine("ERROR: Current document is not an assembly (.iam)!")
                    Console.WriteLine("TYPE: " & Path.GetExtension(assemblyPath))
                    Return
                End If

            Catch ex As Exception
                Console.WriteLine("ERROR: " & ex.Message)
                Return
            End Try
            Return
        End If

        ' Manual mode with explicit paths
        Dim mode As String = args(0).ToUpper()
        Dim outputPath As String = ""

        Select Case mode
            Case "SCAN"
                If args.Length < 3 Then
                    Console.WriteLine("SCAN mode requires: SCAN <assembly_path> <output_folder>")
                    Return
                End If
                ScanAssembly(args(1), args(2))

            Case "PATCH"
                If args.Length < 2 Then
                    Console.WriteLine("PATCH mode requires: PATCH <assembly_path>")
                    Return
                End If
                PatchAssembly(args(1))

            Case "COMPARE"
                If args.Length < 3 Then
                    Console.WriteLine("COMPARE mode requires: COMPARE <before_folder> <after_folder> [report_folder]")
                    Return
                End If
                outputPath = If(args.Length >= 4, args(3), Path.GetDirectoryName(args(1)) & "\Comparison Report")
                CompareScans(args(1), args(2), outputPath)

            Case Else
                Console.WriteLine("Unknown mode: " & mode)
                ShowHelp()
        End Select
    End Sub

    Sub ShowHelp()
        Console.WriteLine("AssemblyScannerPatcher - Scan, Patch, and Compare Inventor Assemblies")
        Console.WriteLine()
        Console.WriteLine("USAGE:")
        Console.WriteLine("  AssemblyScannerPatcher.exe SCAN <assembly_path> <output_folder>")
        Console.WriteLine("  AssemblyScannerPatcher.exe PATCH <assembly_path>")
        Console.WriteLine("  AssemblyScannerPatcher.exe COMPARE <before_folder> <after_folder> [report_folder]")
        Console.WriteLine()
        Console.WriteLine("MODES:")
        Console.WriteLine("  SCAN    - Extract iLogic rules, detect derived parts, analyze structure")
        Console.WriteLine("  PATCH   - Apply mapping file to patch iLogic rules")
        Console.WriteLine("  COMPARE - Generate comparison report between before/after scans")
        Console.WriteLine()
        Console.WriteLine("EXAMPLES:")
        Console.WriteLine("  AssemblyScannerPatcher.exe SCAN ""C:\Assembly\Part.iam"" ""C:\Output\Before """)
        Console.WriteLine("  AssemblyScannerPatcher.exe PATCH ""C:\Assembly\Part.iam""")
        Console.WriteLine("  AssemblyScannerPatcher.exe COMPARE ""C:\Output\Before"" ""C:\Output\After"" ""C:\Output\Report""")
    End Sub

    Sub ScanAssembly(assemblyPath As String, outputFolder As String)
        Console.WriteLine("=== SCAN MODE STARTED ===")
        Console.WriteLine("Assembly: " & assemblyPath)
        Console.WriteLine("Output: " & outputFolder)
        Console.WriteLine()

        ' Create output folder
        Directory.CreateDirectory(outputFolder)

        Dim log As New StringBuilder()
        Dim inventorApp As Object = Nothing

        Try
            ' Connect to Inventor
            Console.WriteLine("Connecting to Inventor...")
            inventorApp = GetInventorApp()
            If inventorApp Is Nothing Then
                Console.WriteLine("ERROR: Could not connect to Inventor!")
                Return
            End If

            ' Open assembly
            Console.WriteLine("Opening assembly...")
            Dim assemblyDoc As Object = OpenDocument(inventorApp, assemblyPath)
            If assemblyDoc Is Nothing Then
                Console.WriteLine("ERROR: Could not open assembly!")
                Return
            End If

            ' Scan iLogic rules
            Console.WriteLine("Scanning iLogic rules...")
            ScaniLogicRules(inventorApp, assemblyDoc, outputFolder, log)

            ' Scan derived parts
            Console.WriteLine("Scanning derived parts...")
            ScanDerivedParts(inventorApp, assemblyDoc, outputFolder, log)

            ' Scan assembly structure
            Console.WriteLine("Scanning assembly structure...")
            ScanAssemblyStructure(assemblyDoc, outputFolder, log)

            ' Save log
            Dim logPath As String = Path.Combine(outputFolder, "Scan_Log.txt")
            File.WriteAllText(logPath, log.ToString())
            Console.WriteLine("Scan log saved: " & logPath)

            ' Save summary
            SaveScanSummary(outputFolder, log.ToString())

            Console.WriteLine()
            Console.WriteLine("=== SCAN COMPLETED ===")
            Console.WriteLine("Results saved to: " & outputFolder)

        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Console.WriteLine(ex.StackTrace)
        Finally
            ' Cleanup if needed
        End Try
    End Sub

    Sub PatchAssembly(assemblyPath As String)
        Console.WriteLine("=== PATCH MODE STARTED ===")
        Console.WriteLine("Assembly: " & assemblyPath)

        ' Find mapping file in same directory as assembly
        Dim assemblyDir As String = Path.GetDirectoryName(assemblyPath)
        Dim mappingFilePath As String = Path.Combine(assemblyDir, "STEP_1_MAPPING.txt")

        If Not File.Exists(mappingFilePath) Then
            Console.WriteLine("ERROR: Mapping file not found!")
            Console.WriteLine("Expected: " & mappingFilePath)
            Return
        End If

        Console.WriteLine("Mapping file: " & mappingFilePath)

        Try
            ' Read mapping file
            Dim partNameMapping As New Dictionary(Of String, String)()
            ReadMappingFile(mappingFilePath, partNameMapping)

            Console.WriteLine("Loaded " & partNameMapping.Count & " part name mappings")

            If partNameMapping.Count = 0 Then
                Console.WriteLine("WARNING: No mappings found in file!")
                Return
            End If

            ' Connect to Inventor
            Console.WriteLine("Connecting to Inventor...")
            Dim inventorApp As Object = GetInventorApp()
            If inventorApp Is Nothing Then
                Console.WriteLine("ERROR: Could not connect to Inventor!")
                Return
            End If

            ' Get or open assembly
            Console.WriteLine("Checking for open assembly...")
            Dim assemblyDoc As Object = GetOpenAssembly(inventorApp, assemblyPath)
            If assemblyDoc Is Nothing Then
                Console.WriteLine("ERROR: Assembly not open! Please open it in Inventor before running.")
                Return
            End If

            ' Create iLogic patcher
            Dim ilogicPatcher As Object = Nothing
            ilogicPatcher = CreateiLogicPatcher(inventorApp)

            If ilogicPatcher Is Nothing Then
                Console.WriteLine("ERROR: Could not create iLogicPatcher!")
                Return
            End If

            ' Patch rules
            Console.WriteLine("Patching iLogic rules...")
            Dim result As Integer = CInt(ilogicPatcher.PatchRulesInAssembly(assemblyDoc, partNameMapping))

            Console.WriteLine("=== PATCH COMPLETED ===")
            Console.WriteLine("Total replacements made: " & result)

        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub

    Sub CompareScans(beforeFolder As String, afterFolder As String, reportFolder As String)
        Console.WriteLine("=== COMPARE MODE STARTED ===")
        Console.WriteLine("Before: " & beforeFolder)
        Console.WriteLine("After: " & afterFolder)
        Console.WriteLine("Report: " & reportFolder)

        Try
            ' Create report folder
            Directory.CreateDirectory(reportFolder)

            Dim report As New StringBuilder()

            report.AppendLine("<!DOCTYPE html>")
            report.AppendLine("<html>")
            report.AppendLine("<head>")
            report.AppendLine("<title>Assembly Comparison Report</title>")
            report.AppendLine("<style>")
            report.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }")
            report.AppendLine("h1 { color: #333; }")
            report.AppendLine("h2 { color: #666; border-bottom: 2px solid #ddd; padding-bottom: 10px; }")
            report.AppendLine(".section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }")
            report.AppendLine(".changed { background-color: #fff3cd; }")
            report.AppendLine(".unchanged { background-color: #f8f9fa; }")
            report.AppendLine(".error { background-color: #f8d7da; }")
            report.AppendLine("table { border-collapse: collapse; width: 100%; margin: 10px 0; }")
            report.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }")
            report.AppendLine("th { background-color: #4CAF50; color: white; }")
            report.AppendLine("tr:nth-child(even) { background-color: #f2f2f2; }")
            report.AppendLine("</style>")
            report.AppendLine("</head>")
            report.AppendLine("<body>")
            report.AppendLine("<h1>Assembly Comparison Report</h1>")
            report.AppendLine("<p>Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "</p>")

            ' Compare files
            CompareFileStructure(beforeFolder, afterFolder, report)

            ' Compare iLogic rules
            CompareiLogicRules(beforeFolder, afterFolder, report)

            ' Compare derived parts
            CompareDerivedParts(beforeFolder, afterFolder, report)

            ' Compare assembly structure
            CompareAssemblyStructure(beforeFolder, afterFolder, report)

            report.AppendLine("</body>")
            report.AppendLine("</html>")

            ' Save report
            Dim reportPath As String = Path.Combine(reportFolder, "Comparison_Report.html")
            File.WriteAllText(reportPath, report.ToString())

            Console.WriteLine("=== COMPARISON COMPLETED ===")
            Console.WriteLine("Report saved: " & reportPath)

        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub

    ' Helper Functions

    ' Late-binding helper functions for COM Object access
    Function GetComProperty(obj As Object, propName As String) As Object
        Try
            Return CallByName(obj, propName, CallType.Get)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Function GetComPropertyInteger(obj As Object, propName As String) As Integer
        Try
            Dim val As Object = GetComProperty(obj, propName)
            If val IsNot Nothing AndAlso TypeOf val Is Integer Then
                Return CInt(val)
            End If
        Catch
        End Try
        Return 0
    End Function

    Function GetComPropertyString(obj As Object, propName As String) As String
        Try
            Dim result As Object = CallByName(obj, propName, CallType.Get)
            If result IsNot Nothing Then
                Return CStr(result)
            End If
            Return ""
        Catch ex As Exception
            ' Return empty for failed property reads
            Return ""
        End Try
    End Function

    Function GetComPropertyObject(obj As Object, propName As String) As Object
        Try
            Dim val As Object = GetComProperty(obj, propName)
            Return val
        Catch
            Return Nothing
        End Try
    End Function

    Function GetComPropertyCollection(obj As Object, propName As String) As Object
        Try
            Dim val As Object = GetComProperty(obj, propName)
            Return val
        Catch
            Return Nothing
        End Try
    End Function

    Function InvokeComMethod(obj As Object, methodName As String, ParamArray args As Object()) As Object
        Try
            Return CallByName(obj, methodName, CallType.Method, args)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Function GetInventorApp() As Object
        Try
            ' Try to get running instance first (GetObject equivalent)
            Return Marshal.GetActiveObject("Inventor.Application")
        Catch ex As COMException
            Console.WriteLine("Error getting running Inventor instance: " & ex.Message)
            Console.WriteLine("Please ensure Inventor is running and an assembly is open")
            Return Nothing
        Catch ex As Exception
            Console.WriteLine("Error getting Inventor: " & ex.Message)
            Return Nothing
        End Try
        Return Nothing
    End Function

    Function OpenDocument(inventorApp As Object, filePath As String) As Object
        Try
            Return inventorApp.Documents.Open(filePath, True)
        Catch ex As Exception
            Console.WriteLine("Error opening document: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Function GetOpenAssembly(inventorApp As Object, filePath As String) As Object
        Try
            For Each doc As Object In inventorApp.Documents
                If doc.FullFileName.ToLower() = filePath.ToLower() Then
                    Return doc
                End If
            Next
        Catch ex As Exception
            Console.WriteLine("Error checking for open assembly: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Function CreateiLogicPatcher(inventorApp As Object) As Object
        Try
            ' Try to load iLogicPatcher class from same directory as EXE
            Dim scriptFolder As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim dllPath As String = Path.Combine(scriptFolder, "iLogicPatcher.dll")

            If Not File.Exists(dllPath) Then
                Console.WriteLine("Note: iLogicPatcher.dll not found in: " & dllPath)
                Return Nothing
            End If

            Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(dllPath)
            Dim patcherType As Type = assembly.GetType("iLogicPatcher")
            Return System.Activator.CreateInstance(patcherType, inventorApp, Nothing)
        Catch ex As Exception
            Console.WriteLine("Note: iLogicPatcher.dll not found or could not load: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Sub ReadMappingFile(filePath As String, ByRef partNameMapping As Dictionary(Of String, String))
        Dim lines As String() = File.ReadAllLines(filePath)
        For Each line As String In lines
            If line.StartsWith("#") OrElse String.IsNullOrEmpty(line) Then Continue For
            Dim parts As String() = line.Split("|"c)
            If parts.Length >= 4 Then
                Dim originalFile As String = Path.GetFileNameWithoutExtension(parts(2))
                Dim newFile As String = Path.GetFileNameWithoutExtension(parts(3))
                If Not partNameMapping.ContainsKey(originalFile) Then
                    partNameMapping.Add(originalFile, newFile)
                End If
            End If
        Next
    End Sub

    Sub ScaniLogicRules(inventorApp As Object, assemblyDoc As Object, outputFolder As String, log As StringBuilder)
        Try
            log.AppendLine("iLogic: Searching for iLogic add-in...")

            Dim addIns As Object = GetComPropertyCollection(inventorApp, "ApplicationAddIns")
            If addIns Is Nothing Then
                log.AppendLine("iLogic: Could not access ApplicationAddIns")
                Return
            End If

            Dim addInCount As Integer = GetComPropertyInteger(addIns, "Count")
            Dim iLogicAddInItem As Object = Nothing
            log.AppendLine("iLogic: Checking " & addInCount & " add-ins...")

            For i As Integer = 1 To addInCount
                Try
                    Dim addIn As Object = InvokeComMethod(addIns, "Item", i)
                    Dim displayName As String = GetComPropertyString(addIn, "DisplayName")

                    If Not String.IsNullOrEmpty(displayName) AndAlso displayName.ToLower().Contains("ilogic") Then
                        iLogicAddInItem = addIn
                        log.AppendLine("iLogic: Found add-in - " & displayName)
                        Exit For
                    End If
                Catch ex As Exception
                    ' Skip errors
                End Try
            Next

            If iLogicAddInItem Is Nothing Then
                log.AppendLine("iLogic: Add-in not found (checked " & addInCount & " add-ins)")
                Return
            End If

            Dim iLogicAuto As Object = GetComPropertyObject(iLogicAddInItem, "Automation")
            If iLogicAuto Is Nothing Then
                log.AppendLine("iLogic: Could not get automation interface")
                Return
            End If

            log.AppendLine("iLogic: Automation interface obtained")

            ' Create iLogic rules folder
            Dim iLogicFolder As String = Path.Combine(outputFolder, "iLogic_Rules")
            Directory.CreateDirectory(iLogicFolder)

            ' Process assembly
            ProcessDocumentiLogic(iLogicAuto, assemblyDoc, iLogicFolder, log)

            ' Process referenced documents
            Dim refDocs As Object = GetComPropertyCollection(assemblyDoc, "AllReferencedDocuments")
            If refDocs IsNot Nothing Then
                For Each refDoc As Object In refDocs
                    ProcessDocumentiLogic(iLogicAuto, refDoc, iLogicFolder, log)
                Next
            End If

        Catch ex As Exception
            log.AppendLine("iLogic: ERROR - " & ex.Message)
            log.AppendLine("   Stack: " & ex.StackTrace)
        End Try
    End Sub

    Sub ProcessDocumentiLogic(iLogicAuto As Object, doc As Object, iLogicFolder As String, log As StringBuilder)
        Try
            Dim docName As String = Path.GetFileName(GetComPropertyString(doc, "FullFileName"))
            log.AppendLine("iLogic: Processing " & docName)

            ' Get rules collection - pass doc as parameter (same pattern as VBScript)
            Dim rules As Object = Nothing
            Try
                rules = InvokeComMethod(iLogicAuto, "Rules", doc)
            Catch ex As Exception
                log.AppendLine("iLogic:   ERROR calling Rules(doc): " & ex.Message)
                Return
            End Try

            If rules Is Nothing Then
                log.AppendLine("iLogic:   No rules in " & docName)
                Return
            End If

            Dim ruleCount As Integer = GetComPropertyInteger(rules, "Count")
            If ruleCount = 0 Then
                log.AppendLine("iLogic:   No rules in " & docName)
                Return
            End If

            ' Create file for this document's rules
            Dim docRulePath As String = Path.Combine(iLogicFolder, docName & "_Rules.txt")
            Dim docRules As New StringBuilder()

            docRules.AppendLine("=== iLogic RULES FOR: " & docName & " ===")
            docRules.AppendLine("Extracted: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            docRules.AppendLine("Total rules: " & ruleCount)
            docRules.AppendLine()

            For i As Integer = 1 To ruleCount
                Try
                    Dim rule As Object = InvokeComMethod(rules, "Item", i)
                    Dim ruleName As String = GetComPropertyString(rule, "Name")
                    Dim ruleText As String = GetComPropertyString(rule, "Text")

                    docRules.AppendLine("------------------------------------------------------------------")
                    docRules.AppendLine("RULE " & i & ": " & ruleName)
                    docRules.AppendLine("------------------------------------------------------------------")
                    docRules.AppendLine(ruleText)
                    docRules.AppendLine()

                    log.AppendLine("iLogic:   Found rule: " & ruleName)
                Catch ruleEx As Exception
                    log.AppendLine("iLogic:   ERROR - " & ruleEx.Message)
                End Try
            Next

            File.WriteAllText(docRulePath, docRules.ToString())
            log.AppendLine("iLogic:   Saved rules to: " & docRulePath)

        Catch ex As Exception
            log.AppendLine("iLogic:   ERROR - " & ex.Message)
        End Try
    End Sub

    Sub ScanDerivedParts(inventorApp As Object, assemblyDoc As Object, outputFolder As String, log As StringBuilder)
        Try
            Dim derivedFolder As String = Path.Combine(outputFolder, "Derived_Parts")
            Directory.CreateDirectory(derivedFolder)

            Dim derivedReport As New StringBuilder()
            derivedReport.AppendLine("=== DERIVED PARTS AUDIT ===")
            derivedReport.AppendLine("Assembly: " & GetComPropertyString(assemblyDoc, "DisplayName"))
            derivedReport.AppendLine("Scanned: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            derivedReport.AppendLine()

            ' Get component definition
            Dim compDef As Object = GetComPropertyObject(assemblyDoc, "ComponentDefinition")
            If compDef Is Nothing Then
                log.AppendLine("Derived: ERROR - Cannot get ComponentDefinition")
                Return
            End If

            ' Get occurrences collection
            Dim occurrences As Object = GetComPropertyCollection(compDef, "Occurrences")
            If occurrences Is Nothing Then
                log.AppendLine("Derived: ERROR - Cannot get Occurrences")
                Return
            End If

            Dim occCount As Integer = GetComPropertyInteger(occurrences, "Count")
            Dim hasDerived As Boolean = False

            For i As Integer = 1 To occCount
                Try
                    Dim comp As Object = InvokeComMethod(occurrences, "Item", i)
                    ScanPartForDerived(comp, derivedReport, log, hasDerived)
                Catch ex As Exception
                    log.AppendLine("Derived: ERROR scanning occurrence " & i & ": " & ex.Message)
                End Try
            Next

            ' Save report
            Dim reportPath As String = Path.Combine(derivedFolder, "Derived_Parts_Report.txt")
            File.WriteAllText(reportPath, derivedReport.ToString())

            If hasDerived Then
                log.AppendLine("Derived: Report saved to: " & reportPath)
            Else
                log.AppendLine("Derived: No derived parts found")
                derivedReport.AppendLine("RESULT: No derived parts detected in assembly")
                File.WriteAllText(reportPath, derivedReport.ToString())
            End If

        Catch ex As Exception
            log.AppendLine("Derived: CRITICAL ERROR - " & ex.Message)
        End Try
    End Sub

    Sub ScanPartForDerived(occ As Object, report As StringBuilder, log As StringBuilder, ByRef hasDerived As Boolean)
        Try
            ' Get DefinitionDocument
            Dim defDoc As Object = GetComPropertyObject(occ, "DefinitionDocument")
            If defDoc Is Nothing Then Return

            Dim docType As Integer = GetComPropertyInteger(defDoc, "Type")
            If docType <> 12288 Then Return ' Not a part document (12288 = kPartDocumentObject)

            Dim partDoc As Object = defDoc
            Dim partDocName As String = Path.GetFileName(GetComPropertyString(partDoc, "FullFileName"))

            ' Get component definition
            Dim partCompDef As Object = GetComPropertyObject(partDoc, "ComponentDefinition")
            If partCompDef Is Nothing Then Return

            ' Get features collection
            Dim features As Object = GetComPropertyCollection(partCompDef, "Features")
            If features Is Nothing Then Return

            Dim featCount As Integer = GetComPropertyInteger(features, "Count")

            For i As Integer = 1 To featCount
                Try
                    Dim feat As Object = InvokeComMethod(features, "Item", i)
                    Dim featType As Integer = GetComPropertyInteger(feat, "Type")

                    ' kDerivedPartFeatureObject = 20480
                    If featType = 20480 Then
                        Dim baseFile As String = ""
                        Try
                            Dim baseComp As Object = GetComPropertyObject(feat, "BaseComponent")
                            If baseComp IsNot Nothing Then
                                baseFile = GetComPropertyString(baseComp, "FullFileName")
                            End If
                        Catch
                            baseFile = "[Unknown - Could not access BaseComponent]"
                        End Try

                        Dim partDir As String = Path.GetDirectoryName(GetComPropertyString(partDoc, "FullFileName"))
                        Dim isExternal As Boolean = baseFile.IndexOf(partDir, StringComparison.OrdinalIgnoreCase) < 0
                        Dim status As String = If(isExternal, "EXTERNAL DERIVED", "LOCAL DERIVED")

                        report.AppendLine(status)
                        report.AppendLine("  Part: " & partDocName)
                        report.AppendLine("  Base Component: " & baseFile)
                        report.AppendLine("  Status: " & If(isExternal, "External reference - needs fix", "Local reference - OK"))
                        report.AppendLine()

                        log.AppendLine("Derived: Found " & status & " - " & partDocName)
                        hasDerived = True
                    End If
                Catch featEx As Exception
                    log.AppendLine("Derived: ERROR checking feature " & i & ": " & featEx.Message)
                End Try
            Next
        Catch ex As Exception
            log.AppendLine("Derived: ERROR scanning part: " & ex.Message)
        End Try
    End Sub

    Sub ScanAssemblyStructure(assemblyDoc As Object, outputFolder As String, log As StringBuilder)
        Try
            Dim structureFile As String = Path.Combine(outputFolder, "Assembly_Structure.txt")
            Dim structureTxt As New StringBuilder()

            structureTxt.AppendLine("=== ASSEMBLY STRUCTURE ===")
            structureTxt.AppendLine("Assembly: " & GetComPropertyString(assemblyDoc, "DisplayName"))
            structureTxt.AppendLine("Path: " & GetComPropertyString(assemblyDoc, "FullFileName"))
            structureTxt.AppendLine("Scanned: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            structureTxt.AppendLine()

            ScanComponentOccurrences(assemblyDoc, structureTxt, "", 0, log)

            File.WriteAllText(structureFile, structureTxt.ToString())
            log.AppendLine("Structure: Saved to: " & structureFile)

        Catch ex As Exception
            log.AppendLine("Structure: ERROR - " & ex.Message)
        End Try
    End Sub

    Sub ScanComponentOccurrences(assemblyDoc As Object, output As StringBuilder, indent As String, level As Integer, log As StringBuilder)
        Try
            Dim compDef As Object = GetComPropertyObject(assemblyDoc, "ComponentDefinition")
            If compDef Is Nothing Then Return

            Dim occurrences As Object = GetComPropertyCollection(compDef, "Occurrences")
            If occurrences Is Nothing Then Return

            Dim occCount As Integer = GetComPropertyInteger(occurrences, "Count")

            For i As Integer = 1 To occCount
                Try
                    Dim occ As Object = InvokeComMethod(occurrences, "Item", i)
                    Dim occName As String = GetComPropertyString(occ, "Name")

                    Dim defDoc As Object = GetComPropertyObject(occ, "DefinitionDocument")
                    Dim defDocName As String = "[Unknown]"
                    If defDoc IsNot Nothing Then
                        defDocName = Path.GetFileName(GetComPropertyString(defDoc, "FullFileName"))
                    End If

                    output.AppendLine(indent & "- " & occName & " [" & defDocName & "]")

                    ' Check if this is a subassembly
                    If defDoc IsNot Nothing Then
                        Dim docType As Integer = GetComPropertyInteger(defDoc, "Type")
                        If docType = 12291 Then ' kAssemblyDocumentObject
                            ScanComponentOccurrences(defDoc, output, indent & "  ", level + 1, log)
                        End If
                    End If
                Catch occEx As Exception
                    log.AppendLine("Structure: ERROR processing occurrence " & i & ": " & occEx.Message)
                End Try
            Next
        Catch ex As Exception
            output.AppendLine(indent & "ERROR: " & ex.Message)
        End Try
    End Sub

    Sub SaveScanSummary(outputFolder As String, log As String)
        Dim summaryPath As String = Path.Combine(outputFolder, "Scan_Summary.html")

        Dim html As New StringBuilder()
        html.AppendLine("<!DOCTYPE html>")
        html.AppendLine("<html>")
        html.AppendLine("<head>")
        html.AppendLine("<style>")
        html.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }")
        html.AppendLine("h1 { color: #333; }")
        html.AppendLine(".section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }")
        html.AppendLine("pre { background-color: #f8f9fa; padding: 10px; border-radius: 5px; overflow-x: auto; }")
        html.AppendLine("</style>")
        html.AppendLine("</head>")
        html.AppendLine("<body>")
        html.AppendLine("<h1>Assembly Scan Summary</h1>")
        html.AppendLine("<p>Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "</p>")

        html.AppendLine("<div class=""section"">")
        html.AppendLine("<h2>Scan Log</h2>")
        html.AppendLine("<pre>" & Replace(log, vbCrLf, "<br/>") & "</pre>")
        html.AppendLine("</div>")

        html.AppendLine("</body>")
        html.AppendLine("</html>")

        File.WriteAllText(summaryPath, html.ToString())
    End Sub

    ' Comparison Functions

    Sub CompareFileStructure(beforeFolder As String, afterFolder As String, report As StringBuilder)
        report.AppendLine("<div class=""section"">")
        report.AppendLine("<h2>File Structure Comparison</h2>")
        report.AppendLine("<table>")
        report.AppendLine("<tr><th>File</th><th>Status</th></tr>")

        ' This is a simplified comparison - can be enhanced
        report.AppendLine("<tr><td>Comparison not yet implemented</td><td class=""unchanged"">Pending</td></tr>")

        report.AppendLine("</table>")
        report.AppendLine("</div>")
    End Sub

    Sub CompareiLogicRules(beforeFolder As String, afterFolder As String, report As StringBuilder)
        report.AppendLine("<div class=""section"">")
        report.AppendLine("<h2>iLogic Rules Comparison</h2>")

        Dim beforeLogicFolder As String = Path.Combine(beforeFolder, "iLogic_Rules")
        Dim afterLogicFolder As String = Path.Combine(afterFolder, "iLogic_Rules")

        If Not Directory.Exists(beforeLogicFolder) OrElse Not Directory.Exists(afterLogicFolder) Then
            report.AppendLine("<p class=""error"">iLogic rules folders not found</p>")
        Else
            Dim beforeFiles As String() = Directory.GetFiles(beforeLogicFolder, "*_Rules.txt")
            Dim afterFiles As String() = Directory.GetFiles(afterLogicFolder, "*_Rules.txt")

            report.AppendLine("<p>Before: " & beforeFiles.Length & " documents with iLogic rules</p>")
            report.AppendLine("<p>After: " & afterFiles.Length & " documents with iLogic rules</p>")

            ' Detailed comparison would go here
            report.AppendLine("<p><em>Detailed comparison pending implementation</em></p>")
        End If

        report.AppendLine("</div>")
    End Sub

    Sub CompareDerivedParts(beforeFolder As String, afterFolder As String, report As StringBuilder)
        report.AppendLine("<div class=""section"">")
        report.AppendLine("<h2>Derived Parts Comparison</h2>")

        Dim beforeFile As String = Path.Combine(beforeFolder, "Derived_Parts\Derived_Parts_Report.txt")
        Dim afterFile As String = Path.Combine(afterFolder, "Derived_Parts\Derived_Parts_Report.txt")

        If File.Exists(beforeFile) AndAlso File.Exists(afterFile) Then
            Dim beforeText As String = File.ReadAllText(beforeFile)
            Dim afterText As String = File.ReadAllText(afterFile)

            If beforeText.Equals(afterText) Then
                report.AppendLine("<p class=""unchanged"">No changes in derived parts</p>")
            Else
                report.AppendLine("<p class=""changed"">Derived parts have changed</p>")
            End If
        Else
            report.AppendLine("<p class=""error"">Derived parts reports not found</p>")
        End If

        report.AppendLine("</div>")
    End Sub

    Sub CompareAssemblyStructure(beforeFolder As String, afterFolder As String, report As StringBuilder)
        report.AppendLine("<div class=""section"">")
        report.AppendLine("<h2>Assembly Structure Comparison</h2>")

        Dim beforeFile As String = Path.Combine(beforeFolder, "Assembly_Structure.txt")
        Dim afterFile As String = Path.Combine(afterFolder, "Assembly_Structure.txt")

        If File.Exists(beforeFile) AndAlso File.Exists(afterFile) Then
            Dim beforeText As String = File.ReadAllText(beforeFile)
            Dim afterText As String = File.ReadAllText(afterFile)

            If beforeText.Equals(afterText) Then
                report.AppendLine("<p class=""unchanged"">Assembly structure unchanged</p>")
            Else
                report.AppendLine("<p class=""changed"">Assembly structure has changed</p>")
            End If
        Else
            report.AppendLine("<p class=""error"">Assembly structure files not found</p>")
        End If

        report.AppendLine("</div>")
    End Sub

End Module
