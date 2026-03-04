' ===============================================================================
' BASE VIEW TITLE UPDATER - PRODUCTION SCRIPT
' ===============================================================================
' Updates base view titles with exact format requirements:
'
' PARTS:                    ASSEMBLIES:
' NCHR01-000-PL1           NCHR01-000-BA1
' SCALE 1:5                7-OFF REQ'D
'                          SCALE 1:20
'
' Font: Part/Assembly names = 3.5mm (0.35cm), other text = 2.5mm (0.25cm)
' ===============================================================================

Option Explicit

Dim invApp
Set invApp = GetObject(, "Inventor.Application")

Sub UpdateAllBaseViewTitles()
    WScript.Echo "BASE VIEW TITLE UPDATER"
    WScript.Echo "======================="

    On Error Resume Next

    ' Get current drawing
    Dim drawingDoc
    Set drawingDoc = FindDrawingDocument()

    If drawingDoc Is Nothing Then
        WScript.Echo "ERROR: No drawing document found"
        Exit Sub
    End If

    WScript.Echo "Updating titles in: " & drawingDoc.DisplayName
    WScript.Echo ""

    ' Get main assembly for quantity calculations
    Dim mainAssemblyDoc
    Set mainAssemblyDoc = GetMainAssemblyDocument()

    If mainAssemblyDoc Is Nothing Then
        WScript.Echo "WARNING: Cannot find main assembly - quantities will show as 1"
        WScript.Echo ""
    Else
        WScript.Echo "Using main assembly: " & mainAssemblyDoc.DisplayName
        WScript.Echo ""
    End If

    ' Process each sheet
    Dim sheets
    Set sheets = drawingDoc.Sheets

    Dim totalUpdated
    totalUpdated = 0

    Dim i
    For i = 1 To sheets.Count
        Dim sheet
        Set sheet = sheets.Item(i)

        WScript.Echo "SHEET " & i & ": " & sheet.Name
        WScript.Echo "--------------------"

        Dim sheetUpdated
        sheetUpdated = UpdateSheetBaseViewTitles(sheet, mainAssemblyDoc)
        totalUpdated = totalUpdated + sheetUpdated

        WScript.Echo ""
    Next

    WScript.Echo "==========================="
    WScript.Echo "SUMMARY: Updated " & totalUpdated & " view titles (base + non-base)"
    WScript.Echo "==========================="
End Sub

Function UpdateSheetBaseViewTitles(sheet, mainAssemblyDoc)
    UpdateSheetBaseViewTitles = 0

    Dim drawingViews
    Set drawingViews = sheet.DrawingViews

    Dim i
    For i = 1 To drawingViews.Count
        Dim view
        Set view = drawingViews.Item(i)

        If IsBaseView(view) Then
            ' Update base views with full title formatting
            If UpdateSingleViewTitle(view, mainAssemblyDoc) Then
                UpdateSheetBaseViewTitles = UpdateSheetBaseViewTitles + 1
            End If
        Else
            ' Update non-base views with proper part/assembly information
            If UpdateNonBaseViewTitle(view) Then
                UpdateSheetBaseViewTitles = UpdateSheetBaseViewTitles + 1
            End If
        End If
    Next
End Function

Function UpdateSingleViewTitle(view, mainAssemblyDoc)
    UpdateSingleViewTitle = False

    On Error Resume Next

    ' Get referenced document
    Dim refDoc
    Set refDoc = view.ReferencedDocumentDescriptor.ReferencedDocument

    If refDoc Is Nothing Then
        WScript.Echo "  SKIP: " & view.Name & " - No referenced document"
        Exit Function
    End If

    ' Get part/assembly name from Part Number property
    Dim partNumber
    partNumber = GetPartNumber(refDoc)

    ' Get scale string
    Dim scaleString
    scaleString = GetScaleString(view)

    ' Create title based on document type
    Dim newTitle
    If refDoc.DocumentType = 12291 Then  ' Assembly
        ' Assembly format: NAME + QTY-OFF REQ'D + SCALE
        Dim quantity
        quantity = GetAssemblyQuantity(refDoc, mainAssemblyDoc)
        newTitle = CreateAssemblyTitle(partNumber, quantity, scaleString)
        WScript.Echo "  ASSEMBLY: " & view.Name & " -> " & partNumber & " (Qty: " & quantity & ")"

    ElseIf refDoc.DocumentType = 12290 Then  ' Part
        ' Part format: NAME + SCALE
        newTitle = CreatePartTitle(partNumber, scaleString)
        WScript.Echo "  PART: " & view.Name & " -> " & partNumber

    Else
        WScript.Echo "  SKIP: " & view.Name & " - Unknown document type"
        Exit Function
    End If

    ' Update the view label
    If UpdateViewLabel(view, newTitle) Then
        UpdateSingleViewTitle = True
    End If

    Err.Clear
End Function

Function UpdateNonBaseViewTitle(view)
    UpdateNonBaseViewTitle = False

    On Error Resume Next

    ' Get parent base view
    Dim parentView
    Set parentView = view.ParentView
    
    If parentView Is Nothing Then
        WScript.Echo "  NON-BASE: " & view.Name & " - Cannot find parent base view"
        Exit Function
    End If

    ' Get referenced document from parent
    Dim refDoc
    Set refDoc = parentView.ReferencedDocumentDescriptor.ReferencedDocument

    If refDoc Is Nothing Then
        WScript.Echo "  NON-BASE: " & view.Name & " - No referenced document in parent view"
        Exit Function
    End If

    ' Get part/assembly name from Part Number property
    Dim partNumber
    partNumber = GetPartNumber(refDoc)

    ' Get scale string
    Dim scaleString
    scaleString = GetScaleString(view)

    ' Create title based on document type (same as parent)
    Dim newTitle
    If refDoc.DocumentType = 12291 Then  ' Assembly
        ' Assembly format: NAME + QTY-OFF REQ'D + SCALE
        Dim quantity
        quantity = 1  ' Projected views show qty as 1
        newTitle = CreateAssemblyTitle(partNumber, quantity, scaleString)
        WScript.Echo "  NON-BASE (ASSEMBLY): " & view.Name & " -> " & partNumber

    ElseIf refDoc.DocumentType = 12290 Then  ' Part
        ' Part format: NAME + SCALE
        newTitle = CreatePartTitle(partNumber, scaleString)
        WScript.Echo "  NON-BASE (PART): " & view.Name & " -> " & partNumber

    Else
        WScript.Echo "  NON-BASE: " & view.Name & " - Unknown document type"
        Exit Function
    End If

    ' Update the view label
    If UpdateViewLabel(view, newTitle) Then
        UpdateNonBaseViewTitle = True
    End If

    Err.Clear
End Function

Function GetPartNumber(doc)
    On Error Resume Next

    Dim propSets
    Set propSets = doc.PropertySets

    Dim designProps
    Set designProps = propSets.Item("Design Tracking Properties")

    If Err.Number = 0 Then
        Dim partNumberProp
        Set partNumberProp = designProps.Item("Part Number")
        If Err.Number = 0 Then
            GetPartNumber = partNumberProp.Value
        Else
            GetPartNumber = "[No Part Number]"
        End If
    Else
        GetPartNumber = "[No Properties]"
    End If

    Err.Clear
End Function

Function GetScaleString(view)
    On Error Resume Next

    Dim scale
    scale = view.Scale

    ' Format scale for display
    If scale = 1 Then
        GetScaleString = "1:1"
    ElseIf scale < 1 Then
        GetScaleString = "1:" & CStr(Int(1 / scale))
    Else
        GetScaleString = CStr(Int(scale)) & ":1"
    End If

    Err.Clear
End Function

Function GetAssemblyQuantity(assemblyDoc, mainAssemblyDoc)
    GetAssemblyQuantity = 1  ' Default quantity

    If mainAssemblyDoc Is Nothing Then
        Exit Function
    End If

    On Error Resume Next

    ' Count occurrences in main assembly
    Dim compDef
    Set compDef = mainAssemblyDoc.ComponentDefinition

    If Err.Number <> 0 Then
        Exit Function
    End If

    Dim occurrences
    Set occurrences = compDef.Occurrences

    Dim count
    count = 0

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        ' Compare file paths (case insensitive)
        If LCase(occ.ReferencedFileDescriptor.FullFileName) = LCase(assemblyDoc.FullFileName) Then
            count = count + 1
        End If
    Next

    If count > 0 Then
        GetAssemblyQuantity = count
    End If

    Err.Clear
End Function

Function CreateAssemblyTitle(partNumber, quantity, scaleString)
    ' Assembly format with mixed font sizes:
    ' NCHR01-000-BA1      (3.5mm = 0.35cm) - UNDERLINED
    ' 7-OFF REQ'D         (2.5mm = 0.25cm)
    ' SCALE 1:20          (2.5mm = 0.25cm)

    Dim title
    title = "<StyleOverride FontSize='0.35' Bold='True' Underline='True'><Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride>" & vbCrLf & _
            "<StyleOverride FontSize='0.25' Bold='True' Underline='False'>" & quantity & "-OFF REQ'D</StyleOverride>" & vbCrLf & _
            "<StyleOverride FontSize='0.25' Bold='True' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"

    CreateAssemblyTitle = title
End Function

Function CreatePartTitle(partNumber, scaleString)
    ' Part format with mixed font sizes:
    ' NCHR01-000-PL1      (3.5mm = 0.35cm) - UNDERLINED
    ' SCALE 1:5           (2.5mm = 0.25cm)

    Dim title
    title = "<StyleOverride FontSize='0.35' Bold='True' Underline='True'><Property Document='model' PropertySet='Design Tracking Properties' Property='Part Number' FormatID='{32853F0F-3444-11D1-9E93-0060B03C1CA6}' PropertyID='5'>PART NUMBER</Property></StyleOverride>" & vbCrLf & _
            "<StyleOverride FontSize='0.25' Bold='True' Underline='False'>SCALE <DrawingViewScale/></StyleOverride>"

    CreatePartTitle = title
End Function

Function UpdateViewLabel(view, newTitle)
    UpdateViewLabel = False

    On Error Resume Next

    ' Get the view label
    Dim viewLabel
    Set viewLabel = view.Label

    If Err.Number <> 0 Then
        WScript.Echo "    ERROR: Cannot access view label"
        Err.Clear
        Exit Function
    End If

    ' Update the formatted text
    viewLabel.FormattedText = newTitle

    If Err.Number = 0 Then
        WScript.Echo "    SUCCESS: Title updated"
        UpdateViewLabel = True
    Else
        WScript.Echo "    ERROR: Failed to update title - " & Err.Description
        Err.Clear
    End If
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
            Exit Function
        End If
    Next
End Function

Function GetMainAssemblyDocument()
    Set GetMainAssemblyDocument = Nothing

    On Error Resume Next

    ' Try to get current project
    Dim activeProject
    Set activeProject = invApp.DesignProjectManager.ActiveDesignProject

    If Err.Number <> 0 Or activeProject Is Nothing Then
        Exit Function
    End If

    ' Find main assembly in project
    Dim mainAssemblyPath
    mainAssemblyPath = FindMainAssemblyInProject(activeProject)

    If mainAssemblyPath <> "" Then
        ' Try to open main assembly
        Set GetMainAssemblyDocument = invApp.Documents.Open(mainAssemblyPath, False)
    End If

    Err.Clear
End Function

Function FindMainAssemblyInProject(project)
    FindMainAssemblyInProject = ""

    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Search for Structure.iam in project paths (excluding BOM folders)
    Dim libraryPaths
    Set libraryPaths = project.LibraryPaths

    Dim i
    For i = 1 To libraryPaths.Count
        Dim searchPath
        searchPath = libraryPaths.Item(i)

        ' Skip BOM folders
        If InStr(LCase(searchPath), "bom") = 0 Then
            Dim structureFile
            structureFile = searchPath & "\Structure.iam"

            If fso.FileExists(structureFile) Then
                FindMainAssemblyInProject = structureFile
                Exit Function
            End If
        End If
    Next

    Err.Clear
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

' Execute the title updates
Call UpdateAllBaseViewTitles()