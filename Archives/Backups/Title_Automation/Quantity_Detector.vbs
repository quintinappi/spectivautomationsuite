' ===============================================================================
' PROJECT QUANTITY DETECTOR - AUTO PROJECT DETECTION
' ===============================================================================
' Automatically finds current project, locates Structure.iam, counts occurrences
' ===============================================================================

Option Explicit

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Sub AutoDetectProjectQuantities()
    WScript.Echo "PROJECT QUANTITY DETECTOR"
    WScript.Echo "========================="

    ' Step 1: Get current project
    Dim currentProject
    Set currentProject = GetCurrentProject()

    If currentProject Is Nothing Then
        Exit Sub
    End If

    ' Step 2: Find Structure.iam in project
    Dim structurePath
    structurePath = FindStructureInProject(currentProject)

    If structurePath = "" Then
        Exit Sub
    End If

    ' Step 3: Open Structure.iam
    Dim structureDoc
    Set structureDoc = OpenStructureAssembly(structurePath)

    If structureDoc Is Nothing Then
        Exit Sub
    End If

    ' Step 4: Analyze current drawing's base views and get quantities
    Call AnalyzeDrawingQuantities(structureDoc)

    ' Clean up - close Structure.iam if we opened it
    WScript.Echo ""
    WScript.Echo "Closing Structure.iam..."
    structureDoc.Close
End Sub

Function GetCurrentProject()
    WScript.Echo "Detecting current project..."

    On Error Resume Next

    Dim activeProject
    Set activeProject = invApp.DesignProjectManager.ActiveDesignProject

    If Err.Number <> 0 Or activeProject Is Nothing Then
        WScript.Echo "ERROR: No active project found"
        WScript.Echo "Please ensure a project is open in Inventor"
        Set GetCurrentProject = Nothing
        Exit Function
    End If

    WScript.Echo "Active Project: " & activeProject.Name
    WScript.Echo "Workspace: " & activeProject.WorkspacePath
    WScript.Echo ""

    Set GetCurrentProject = activeProject
End Function

Function FindStructureInProject(project)
    WScript.Echo "Searching for main model assembly in project..."

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Common main assembly names to search for
    Dim mainAssemblyNames
    mainAssemblyNames = Array("Structure.iam", "Main.iam", "Assembly.iam", "Model.iam", "Project.iam")

    ' Method 1: Search in project library paths (but exclude BOM folders)
    Dim libraryPaths
    Set libraryPaths = project.LibraryPaths

    Dim i, j
    For i = 1 To libraryPaths.Count
        Dim searchPath
        searchPath = libraryPaths.Item(i)

        ' Skip BOM folders
        If InStr(LCase(searchPath), "bom") = 0 Then
            For j = 0 To UBound(mainAssemblyNames)
                Dim structureFile
                structureFile = searchPath & "\" & mainAssemblyNames(j)

                If fso.FileExists(structureFile) Then
                    WScript.Echo "FOUND MAIN MODEL: " & structureFile
                    FindStructureInProject = structureFile
                    Exit Function
                End If
            Next
        End If
    Next

    ' Method 2: Search recursively in workspace (exclude BOM folders)
    Dim workspacePath
    workspacePath = project.WorkspacePath

    WScript.Echo "Searching recursively in workspace (excluding BOM folders): " & workspacePath

    For j = 0 To UBound(mainAssemblyNames)
        Dim foundFile
        foundFile = SearchMainAssemblyRecursively(workspacePath, mainAssemblyNames(j))
        If foundFile <> "" Then
            WScript.Echo "FOUND MAIN MODEL: " & foundFile
            FindStructureInProject = foundFile
            Exit Function
        End If
    Next

    WScript.Echo "ERROR: Main model assembly not found in project"
    WScript.Echo "Searched for: Structure.iam, Main.iam, Assembly.iam, Model.iam, Project.iam"
    WScript.Echo "Searched paths (excluding BOM folders):"
    For i = 1 To libraryPaths.Count
        WScript.Echo "  " & libraryPaths.Item(i)
    Next
    WScript.Echo "  " & workspacePath & " (recursive)"

    FindStructureInProject = ""
End Function

Function SearchMainAssemblyRecursively(startPath, fileName)
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Skip BOM folders
    If InStr(LCase(startPath), "bom") > 0 Then
        SearchMainAssemblyRecursively = ""
        Exit Function
    End If

    ' Check current directory
    Dim targetFile
    targetFile = startPath & "\" & fileName
    If fso.FileExists(targetFile) Then
        SearchMainAssemblyRecursively = targetFile
        Exit Function
    End If

    ' Search subdirectories (except BOM folders)
    If fso.FolderExists(startPath) Then
        Dim folder
        Set folder = fso.GetFolder(startPath)

        Dim subFolder
        For Each subFolder In folder.SubFolders
            ' Skip BOM folders
            If InStr(LCase(subFolder.Name), "bom") = 0 Then
                Dim result
                result = SearchMainAssemblyRecursively(subFolder.Path, fileName)
                If result <> "" Then
                    SearchMainAssemblyRecursively = result
                    Exit Function
                End If
            End If
        Next
    End If

    SearchMainAssemblyRecursively = ""
End Function

Function SearchFileRecursively(startPath, fileName)
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check current directory
    Dim targetFile
    targetFile = startPath & "\" & fileName
    If fso.FileExists(targetFile) Then
        SearchFileRecursively = targetFile
        Exit Function
    End If

    ' Search subdirectories
    If fso.FolderExists(startPath) Then
        Dim folder
        Set folder = fso.GetFolder(startPath)

        Dim subFolder
        For Each subFolder In folder.SubFolders
            Dim result
            result = SearchFileRecursively(subFolder.Path, fileName)
            If result <> "" Then
                SearchFileRecursively = result
                Exit Function
            End If
        Next
    End If

    SearchFileRecursively = ""
End Function

Function OpenStructureAssembly(structurePath)
    WScript.Echo "Opening Structure.iam..."

    On Error Resume Next

    Dim structureDoc
    Set structureDoc = invApp.Documents.Open(structurePath, False)

    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Could not open Structure.iam - " & Err.Description
        Set OpenStructureAssembly = Nothing
        Exit Function
    End If

    If structureDoc.DocumentType <> 12291 Then  ' Not an assembly
        WScript.Echo "ERROR: Structure.iam is not an assembly document"
        structureDoc.Close
        Set OpenStructureAssembly = Nothing
        Exit Function
    End If

    WScript.Echo "SUCCESS: Structure.iam opened"
    WScript.Echo ""

    Set OpenStructureAssembly = structureDoc
End Function

Sub AnalyzeDrawingQuantities(structureDoc)
    WScript.Echo "ANALYZING DRAWING QUANTITIES"
    WScript.Echo "============================="

    ' Find the drawing document (not the Structure.iam we just opened)
    Dim drawingDoc
    Set drawingDoc = FindDrawingDocument()

    If drawingDoc Is Nothing Then
        WScript.Echo "ERROR: Cannot find drawing document"
        Exit Sub
    End If

    Dim sheets
    Set sheets = drawingDoc.Sheets

    Dim i
    For i = 1 To sheets.Count
        Dim sheet
        Set sheet = sheets.Item(i)

        WScript.Echo "SHEET " & i & ": " & sheet.Name
        WScript.Echo "--------------------"

        Call ProcessSheetQuantities(sheet, structureDoc)
        WScript.Echo ""
    Next
End Sub

Sub ProcessSheetQuantities(sheet, structureDoc)
    Dim drawingViews
    Set drawingViews = sheet.DrawingViews

    Dim baseViewCount
    baseViewCount = 0

    Dim i
    For i = 1 To drawingViews.Count
        Dim view
        Set view = drawingViews.Item(i)

        If IsBaseView(view) Then
            baseViewCount = baseViewCount + 1

            Dim refDoc
            Set refDoc = view.ReferencedDocumentDescriptor.ReferencedDocument

            If Not refDoc Is Nothing Then
                WScript.Echo "  " & baseViewCount & ". " & view.Name
                WScript.Echo "     Assembly: " & refDoc.DisplayName

                ' Count occurrences in Structure.iam
                Dim quantity
                quantity = CountOccurrencesInStructure(refDoc.FullFileName, structureDoc)

                WScript.Echo "     Quantity in Structure.iam: " & quantity
                WScript.Echo ""
            End If
        End If
    Next
End Sub

Function CountOccurrencesInStructure(targetFilePath, structureDoc)
    CountOccurrencesInStructure = 0

    On Error Resume Next

    Dim compDef
    Set compDef = structureDoc.ComponentDefinition

    If Err.Number <> 0 Then
        Exit Function
    End If

    Dim occurrences
    Set occurrences = compDef.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Compare file paths (case insensitive)
        If LCase(occ.ReferencedFileDescriptor.FullFileName) = LCase(targetFilePath) Then
            CountOccurrencesInStructure = CountOccurrencesInStructure + 1
        End If
    Next

    Err.Clear
End Function

Function FindDrawingDocument()
    Set FindDrawingDocument = Nothing

    ' Look through all open documents for a drawing
    Dim docs
    Set docs = invApp.Documents

    Dim i
    For i = 1 To docs.Count
        Dim doc
        Set doc = docs.Item(i)

        If doc.DocumentType = 12292 Then  ' Drawing document
            Set FindDrawingDocument = doc
            WScript.Echo "Using drawing: " & doc.DisplayName
            Exit Function
        End If
    Next

    WScript.Echo "ERROR: No drawing document found in open documents"
End Function

Function IsBaseView(view)
    IsBaseView = False

    On Error Resume Next
    Dim parentView
    Set parentView = view.ParentView

    If Err.Number <> 0 Or parentView Is Nothing Then
        IsBaseView = True
    End If

    Err.Clear
End Function

Call AutoDetectProjectQuantities()