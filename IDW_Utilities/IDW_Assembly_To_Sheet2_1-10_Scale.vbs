' ==================================================================================
' IDW ASSEMBLY TO NEXT SHEET - 1:10 SCALE PLACER
' ==================================================================================
' Asks user which sheet to scan, identifies the assembly and ALL its parts,
' then places them on the next available sheet at scale 1:10 (0.1)
'
' Features:
' - User selects which sheet to scan from a list
' - Extracts ALL parts from the assembly (recursively through sub-assemblies)
' - Creates base views on the NEXT sheet for each unique part
' - Sets scale to 1:10 (0.1) for all placed views
' - Auto-positions views in grid layout
' - Creates new sheet if next one doesn't exist
'
' Requirements:
' - Inventor must be running with the IDW file open
' - Selected sheet must have at least one view of the assembly
'
' API Used: Inventor Application API (full API, NOT Apprentice)
' Reason: ApprenticeServer cannot create/modify drawing views
' ==================================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim logPath

' Configuration
Const VIEWS_PER_SHEET = 12     ' Max views per sheet (3 columns x 4 rows)
Const COLUMN_SPACING = 8       ' cm between columns
Const ROW_SPACING = 6          ' cm between rows
Const MARGIN_LEFT = 2          ' cm left margin
Const MARGIN_BOTTOM = 18       ' cm from bottom (above title block)
Const TARGET_SCALE = 0.1       ' 1:10 scale = 0.1

' Main execution
Main

Sub Main()
    On Error Resume Next

    ' Setup logging
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\IDW_Assembly_To_Sheet2.log"
    Set logFile = fso.CreateTextFile(logPath, True)

    LogMessage "=== IDW ASSEMBLY TO SHEET 2 (1:10 SCALE) ==="
    LogMessage "Starting operation..."
    LogMessage "Target scale: 1:10 (" & TARGET_SCALE & ")"

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor."
        MsgBox "ERROR: Could not connect to Inventor." & vbCrLf & _
               "Please make sure Inventor is running with an IDW file open.", vbCritical, "Inventor Not Found"
        WScript.Quit
    End If
    LogMessage "SUCCESS: Connected to Inventor"

    ' Get active document
    Set invDoc = invApp.ActiveDocument
    If Err.Number <> 0 Or invDoc Is Nothing Then
        LogMessage "ERROR: No active document found."
        MsgBox "ERROR: No active document found." & vbCrLf & _
               "Please open an IDW drawing file.", vbCritical, "No Document"
        WScript.Quit
    End If

    ' Check if it's a drawing document
    If invDoc.DocumentType <> 12294 And invDoc.DocumentType <> 12292 Then
        LogMessage "ERROR: Active document is not a drawing (Type: " & invDoc.DocumentType & ")"
        MsgBox "ERROR: Active document is not a drawing file." & vbCrLf & _
               "Please open an IDW or DWG drawing.", vbCritical, "Wrong Document Type"
        WScript.Quit
    End If
    LogMessage "SUCCESS: Active document is a drawing: " & invDoc.DisplayName

    ' Check if any sheets exist
    If invDoc.Sheets.Count < 1 Then
        LogMessage "ERROR: Drawing has no sheets!"
        MsgBox "ERROR: This drawing has no sheets!", vbCritical, "No Sheets"
        WScript.Quit
    End If

    ' Let user select which sheet to scan
    Dim sourceSheet, sourceSheetIndex
    sourceSheetIndex = SelectSheetToScan(invDoc)
    
    If sourceSheetIndex = 0 Then
        LogMessage "User cancelled sheet selection"
        WScript.Quit
    End If
    
    Set sourceSheet = invDoc.Sheets.Item(sourceSheetIndex)
    LogMessage "User selected Sheet " & sourceSheetIndex & ": " & sourceSheet.Name

    If sourceSheet.DrawingViews.Count = 0 Then
        LogMessage "ERROR: Selected sheet has no views!"
        MsgBox "ERROR: The selected sheet has no views!" & vbCrLf & _
               "Please place the assembly view on the sheet first.", vbExclamation, "No Views"
        WScript.Quit
    End If
    LogMessage "Source sheet has " & sourceSheet.DrawingViews.Count & " view(s)"

    ' Get the first view (should be the assembly)
    Dim firstView
    Set firstView = sourceSheet.DrawingViews.Item(1)

    ' Get the referenced document (assembly)
    Dim assemblyDoc
    Set assemblyDoc = firstView.ReferencedDocumentDescriptor.ReferencedDocument

    If assemblyDoc Is Nothing Then
        LogMessage "ERROR: Could not get referenced document from view."
        MsgBox "ERROR: Could not get referenced document from view." & vbCrLf & _
               "Make sure the view on Sheet 1 references an assembly.", vbCritical, "No Reference"
        WScript.Quit
    End If

    ' Check if it's an assembly
    If assemblyDoc.DocumentType <> 12291 Then
        LogMessage "WARNING: Referenced document is not an assembly (Type: " & assemblyDoc.DocumentType & ")"
        LogMessage "Will process the single document..."
    Else
        LogMessage "SUCCESS: Found assembly: " & assemblyDoc.DisplayName
    End If

    ' Collect all unique parts from the assembly
    LogMessage ""
    LogMessage "Scanning for all parts in assembly..."
    Dim partList
    partList = GetAllParts(assemblyDoc)

    Dim partCount
    partCount = UBound(partList) + 1
    LogMessage "Found " & partCount & " unique part(s)"

    If partCount = 0 Then
        LogMessage "ERROR: No parts found in assembly!"
        MsgBox "ERROR: No parts found in the assembly!" & vbCrLf & _
               "Check that the assembly contains parts.", vbExclamation, "No Parts"
        WScript.Quit
    End If

    ' Determine target sheet (next sheet after source)
    Dim targetSheet, targetSheetIndex
    targetSheetIndex = sourceSheetIndex + 1
    
    If invDoc.Sheets.Count < targetSheetIndex Then
        LogMessage "Creating Sheet " & targetSheetIndex & "..."
        Set targetSheet = invDoc.Sheets.Add()
        targetSheet.Name = "Sheet:" & targetSheetIndex
        LogMessage "Created Sheet " & targetSheetIndex
    Else
        Set targetSheet = invDoc.Sheets.Item(targetSheetIndex)
        LogMessage "Using existing Sheet " & targetSheetIndex
    End If

    ' Activate target sheet before placing views
    targetSheet.Activate
    LogMessage "Activated target sheet: " & targetSheet.Name

    ' Place parts on target sheet
    LogMessage ""
    LogMessage "=== PLACING PART VIEWS AT 1:10 SCALE ==="

    Dim viewsPlaced, viewsFailed
    viewsPlaced = 0
    viewsFailed = 0

    Dim viewIndexOnSheet
    viewIndexOnSheet = 0

    Dim i, partPath
    For i = 0 To UBound(partList)
        partPath = partList(i)

        ' Create base view for this part at 1:10 scale
        Dim partFileName
        partFileName = fso.GetFileName(partPath)
        LogMessage ""
        LogMessage "[" & (i + 1) & "/" & partCount & "] Placing: " & partFileName

        If PlacePartViewAtScale(targetSheet, partPath, viewIndexOnSheet, TARGET_SCALE) Then
            viewsPlaced = viewsPlaced + 1
            viewIndexOnSheet = viewIndexOnSheet + 1
        Else
            viewsFailed = viewsFailed + 1
        End If
    Next
    
    ' Force screen update
    invApp.ScreenUpdating = True
    
    ' Verify views were actually placed by counting them
    Dim actualViewCount
    actualViewCount = targetSheet.DrawingViews.Count
    LogMessage ""
    LogMessage "Verification: Target sheet now has " & actualViewCount & " view(s)"

    ' Save the drawing
    LogMessage ""
    LogMessage "Saving drawing..."
    invDoc.Save2 True
    LogMessage "SUCCESS: Drawing saved"

    ' Final report
    LogMessage ""
    LogMessage "=== PLACEMENT COMPLETE ==="
    LogMessage "Source Sheet: " & sourceSheet.Name
    LogMessage "Target Sheet: " & targetSheet.Name
    LogMessage "Total parts found: " & partCount
    LogMessage "Views placed at 1:10: " & viewsPlaced
    LogMessage "Views failed: " & viewsFailed
    LogMessage "Actual views on target sheet: " & actualViewCount

    logFile.Close

    ' Show message to user
    Dim msg
    msg = "Part Placement Complete!" & vbCrLf & vbCrLf
    msg = msg & "Source: " & sourceSheet.Name & vbCrLf
    msg = msg & "Target: " & targetSheet.Name & vbCrLf
    msg = msg & "Assembly: " & assemblyDoc.DisplayName & vbCrLf
    msg = msg & "Parts found: " & partCount & vbCrLf
    msg = msg & "Views placed: " & viewsPlaced & vbCrLf
    msg = msg & "Actual views on " & targetSheet.Name & ": " & actualViewCount & vbCrLf
    If viewsFailed > 0 Then
        msg = msg & "Views failed: " & viewsFailed & vbCrLf & vbCrLf
        msg = msg & "Check log file for details:" & vbCrLf & logPath
    End If
    
    ' Warn if counts don't match
    If actualViewCount <> viewsPlaced Then
        msg = msg & vbCrLf & "WARNING: View count mismatch!" & vbCrLf
        msg = msg & "Expected " & viewsPlaced & " but found " & actualViewCount
    End If

    MsgBox msg, vbInformation, "Placement Complete"

End Sub

' Function to get all unique parts from an assembly (recursively)
Function GetAllParts(assemblyDoc)
    Dim parts
    ReDim parts(0)
    Dim count
    count = 0

    ' If it's a part, just return it
    If assemblyDoc.DocumentType = 12290 Then ' kPartDocumentObject
        parts(0) = assemblyDoc.FullFileName
        GetAllParts = parts
        Exit Function
    End If

    ' It's an assembly - collect all parts
    Dim occurrences
    Set occurrences = assemblyDoc.ComponentDefinition.Occurrences

    Dim i, occ, refDoc, partPath
    Dim isDuplicate

    For i = 1 To occurrences.Count
        Set occ = occurrences.Item(i)
        Set refDoc = occ.ReferencedDocumentDescriptor.ReferencedDocument

        ' Check if it's a part
        If refDoc.DocumentType = 12290 Then ' kPartDocumentObject
            partPath = refDoc.FullFileName

            ' Check for duplicate
            isDuplicate = False
            Dim j
            For j = 0 To count - 1
                If LCase(parts(j)) = LCase(partPath) Then
                    isDuplicate = True
                    Exit For
                End If
            Next

            ' Add if not duplicate
            If Not isDuplicate Then
                If count > 0 Then
                    ReDim Preserve parts(count)
                End If
                parts(count) = partPath
                count = count + 1
            End If

        ' It's a sub-assembly - recurse
        ElseIf refDoc.DocumentType = 12291 Then ' kAssemblyDocumentObject
            Dim subParts
            subParts = GetAllParts(refDoc)

            Dim k
            For k = LBound(subParts) To UBound(subParts)
                partPath = subParts(k)

                ' Check for duplicate
                isDuplicate = False
                For j = 0 To count - 1
                    If LCase(parts(j)) = LCase(partPath) Then
                        isDuplicate = True
                        Exit For
                    End If
                Next

                ' Add if not duplicate
                If Not isDuplicate Then
                    If count > 0 Then
                        ReDim Preserve parts(count)
                    End If
                    parts(count) = partPath
                    count = count + 1
                End If
            Next
        End If
    Next

    ' Trim array
    If count > 0 Then
        ReDim Preserve parts(count - 1)
    Else
        ReDim parts(0)
    End If

    GetAllParts = parts
End Function

' Function to place a single part view on a sheet at specified scale
Function PlacePartViewAtScale(sheet, partPath, viewIndex, viewScale)
    On Error Resume Next
    PlacePartViewAtScale = False

    ' Check if part is already open in Inventor
    Dim partDoc
    Set partDoc = Nothing

    Dim i
    For i = 1 To invApp.Documents.Count
        Dim doc
        Set doc = invApp.Documents.Item(i)
        If LCase(doc.FullFileName) = LCase(partPath) Then
            Set partDoc = doc
            LogMessage "  Using already-open part document"
            Exit For
        End If
    Next

    ' If not already open, open it
    Dim weOpenedIt
    weOpenedIt = False

    If partDoc Is Nothing Then
        Set partDoc = invApp.Documents.Open(partPath, True)
        weOpenedIt = True

        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "  ERROR: Could not open part: " & Err.Description
            Err.Clear
            Exit Function
        End If
    End If

    ' Calculate position based on view index (grid layout)
    ' Inventor coordinates: Origin (0,0) is bottom-left, Y increases UP
    Dim col, row
    col = viewIndex Mod 3  ' 3 columns (0, 1, 2)
    row = viewIndex \ 3    ' rows (0, 1, 2, 3...)

    Dim x, y
    x = MARGIN_LEFT + (col * COLUMN_SPACING)
    y = MARGIN_BOTTOM + (row * ROW_SPACING)  ' Add to go UP the sheet

    LogMessage "  Position: Column " & (col + 1) & ", Row " & (row + 1) & " (X:" & x & "cm, Y:" & y & "cm from bottom)"
    LogMessage "  Target scale: 1:" & (1/viewScale)

    ' Create point for view placement
    Dim position
    Set position = invApp.TransientGeometry.CreatePoint2d(x, y)

    ' Activate the sheet first (required when adding views)
    sheet.Activate
    LogMessage "  Activated sheet: " & sheet.Name & " (now has " & invDoc.ActiveSheet.DrawingViews.Count & " views)"

    ' Count views before adding
    Dim viewsBefore
    viewsBefore = sheet.DrawingViews.Count
    LogMessage "  Views before: " & viewsBefore

    ' Create base view with specified scale - use literal 1 for now to test
    Dim baseView
    If viewScale = 0.1 Then
        Set baseView = sheet.DrawingViews.AddBaseView(partDoc, position, 0.1)
    Else
        Set baseView = sheet.DrawingViews.AddBaseView(partDoc, position, 1)
    End If

    ' Force update and verify view was actually added
    invApp.ScreenUpdating = True
    WScript.Sleep 50
    Dim viewsAfter
    viewsAfter = sheet.DrawingViews.Count
    LogMessage "  Views after: " & viewsAfter

    ' Check if view was created
    Dim isBaseViewNothing
    isBaseViewNothing = (baseView Is Nothing)
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR: AddBaseView failed with Err #" & Err.Number & ": " & Err.Description
        Err.Clear
        If weOpenedIt Then partDoc.Close
        Exit Function
    End If
    
    If isBaseViewNothing Then
        LogMessage "  ERROR: baseView is Nothing"
        If weOpenedIt Then partDoc.Close
        Exit Function
    End If
    
    If viewsAfter <= viewsBefore Then
        LogMessage "  ERROR: View count didn't increase (still " & viewsAfter & ")"
        If weOpenedIt Then partDoc.Close
        Exit Function
    End If
    
    LogMessage "  SUCCESS: View created at 1:" & (1/viewScale) & " scale"

    ' Only close if we opened it
    If weOpenedIt Then
        partDoc.Close
    End If

    PlacePartViewAtScale = True
End Function

' Function to let user select which sheet to scan
Function SelectSheetToScan(drawingDoc)
    On Error Resume Next
    
    Dim sheetCount, i, sheetList, promptText
    sheetCount = drawingDoc.Sheets.Count
    
    ' Build list of sheets
    sheetList = ""
    For i = 1 To sheetCount
        Dim s
        Set s = drawingDoc.Sheets.Item(i)
        sheetList = sheetList & i & ". " & s.Name & " (" & s.DrawingViews.Count & " views)" & vbCrLf
    Next
    
    ' Build prompt text
    promptText = "Select which sheet contains the assembly view:" & vbCrLf & vbCrLf
    promptText = promptText & sheetList & vbCrLf
    promptText = promptText & "Enter sheet number (1-" & sheetCount & "):"
    
    ' Show input dialog
    Dim userInput, selectedIndex
    userInput = InputBox(promptText, "Select Source Sheet", "1")
    
    ' Check if user cancelled
    If userInput = "" Then
        SelectSheetToScan = 0
        Exit Function
    End If
    
    ' Validate input
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation, "Invalid Input"
        SelectSheetToScan = 0
        Exit Function
    End If
    
    selectedIndex = CInt(userInput)
    
    If selectedIndex < 1 Or selectedIndex > sheetCount Then
        MsgBox "Sheet number must be between 1 and " & sheetCount & ".", vbExclamation, "Invalid Input"
        SelectSheetToScan = 0
        Exit Function
    End If
    
    ' Verify selected sheet has views
    Dim selectedSheet
    Set selectedSheet = drawingDoc.Sheets.Item(selectedIndex)
    If selectedSheet.DrawingViews.Count = 0 Then
        MsgBox "The selected sheet has no views. Please select a sheet with at least one view.", _
               vbExclamation, "No Views"
        SelectSheetToScan = 0
        Exit Function
    End If
    
    SelectSheetToScan = selectedIndex
End Function

Sub LogMessage(msg)
    Dim timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)

    If Not logFile Is Nothing Then
        logFile.WriteLine timestamp & " - " & msg
    End If
End Sub
