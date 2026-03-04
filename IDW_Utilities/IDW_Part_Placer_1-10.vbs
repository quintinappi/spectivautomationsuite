' ==================================================================================
' IDW PART PLACER - 1:10 SCALE VERSION
' ==================================================================================
' Automatically creates base views for ALL parts from the assembly on Sheet 1
' Places them on Sheet 2 (and Sheet 3 if needed) in a grid layout
'
' Features:
' - Reads assembly model from Sheet 1
' - Scans ALL parts in the assembly (including sub-assemblies)
' - Creates base views for each unique part
' - Auto-positions views in grid layout
' - Adds Sheet 3 if Sheet 2 runs out of space
' - Smart spacing and alignment
'
' Usage:
' 1. Open IDW drawing with assembly placed on Sheet 1
' 2. Run this script
' 3. All parts will be placed on Sheet 2 (and Sheet 3 if needed)
'
' ==================================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim logPath

' Configuration
Const VIEWS_PER_SHEET = 12  ' Max views per sheet (3 columns x 4 rows)
Const COLUMN_SPACING = 20   ' cm between columns
Const ROW_SPACING = 15      ' cm between rows
Const MARGIN_LEFT = 3       ' cm left margin
Const MARGIN_TOP = 27       ' cm top margin (below title block)
Const MARGIN_BOTTOM = 5     ' cm bottom margin

' Main execution
Main

Sub Main()
    On Error Resume Next

    ' Setup logging
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\IDW_Part_Placer_1-10.log"
    Set logFile = fso.CreateTextFile(logPath, True)

    LogMessage "=== IDW PART PLACER ==="
    LogMessage "Starting part placement operation..."

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Could not connect to Inventor."
        MsgBox "ERROR: Could not connect to Inventor." & vbCrLf & _
               "Please make sure Inventor is running with an IDW file open.", vbCritical, "Inventor Not Found"
        WScript.Quit
    End If
    LogMessage "SUCCESS: Connected to Inventor"

    ' Get active document - or search for open drawings
    Set invDoc = invApp.ActiveDocument
    If Err.Number <> 0 Or invDoc Is Nothing Then
        LogMessage "WARNING: No active document found. Searching for open drawings..."

        ' Try to find any open drawing document
        Dim i, foundDoc
        foundDoc = False
        For i = 1 To invApp.Documents.Count
            Dim doc
            Set doc = invApp.Documents.Item(i)
            LogMessage "  Open doc [" & i & "]: " & doc.DisplayName & " (Type: " & doc.DocumentType & ")"

            If doc.DocumentType = 12294 Or doc.DocumentType = 12292 Then ' kDrawingDocumentObject (IDW or DWG)
                Set invDoc = doc
                foundDoc = True
                LogMessage "  -> Found drawing! Using this document."
                Exit For
            End If
        Next

        If Not foundDoc Then
            MsgBox "ERROR: No active document found and no drawings open." & vbCrLf & _
                   "Please open an IDW drawing file.", vbCritical, "No Document"
            WScript.Quit
        End If
    End If

    ' Check if it's a drawing document (accept both IDW 12294 and DWG 12292)
    If invDoc.DocumentType <> 12294 And invDoc.DocumentType <> 12292 Then ' kDrawingDocumentObject
        LogMessage "WARNING: Active document is not a drawing (Type: " & invDoc.DocumentType & ")"
        LogMessage "Active document: " & invDoc.DisplayName
        LogMessage "Searching for open drawing documents instead..."

        ' Try to find any open drawing document
        Dim j, foundDoc2
        foundDoc2 = False
        For j = 1 To invApp.Documents.Count
            Dim doc2
            Set doc2 = invApp.Documents.Item(j)
            LogMessage "  Open doc [" & j & "]: " & doc2.DisplayName & " (Type: " & doc2.DocumentType & ")"

            If doc2.DocumentType = 12294 Or doc2.DocumentType = 12292 Then ' kDrawingDocumentObject (IDW or DWG)
                Set invDoc = doc2
                foundDoc2 = True
                LogMessage "  -> Found drawing! Using this document."
                Exit For
            End If
        Next

        If Not foundDoc2 Then
            MsgBox "ERROR: Active document is not a drawing file." & vbCrLf & _
                   "Active: " & invDoc.DisplayName & vbCrLf & vbCrLf & _
                   "Please open an IDW or DWG drawing and make it the active window.", vbCritical, "Wrong Document Type"
            WScript.Quit
        End If
    End If
    LogMessage "SUCCESS: Using drawing: " & invDoc.DisplayName

    ' Check if Sheet 1 exists and has views
    If invDoc.Sheets.Count < 1 Then
        LogMessage "ERROR: Drawing has no sheets!"
        MsgBox "ERROR: This drawing has no sheets!", vbCritical, "No Sheets"
        WScript.Quit
    End If

    Dim sourceSheet
    Set sourceSheet = invDoc.Sheets.Item(1)

    If sourceSheet.DrawingViews.Count = 0 Then
        LogMessage "ERROR: Sheet 1 has no views!"
        MsgBox "ERROR: Sheet 1 has no views!" & vbCrLf & _
               "Please place the assembly view on Sheet 1 first.", vbExclamation, "No Views"
        WScript.Quit
    End If
    LogMessage "Sheet 1 has " & sourceSheet.DrawingViews.Count & " view(s)"

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
    If assemblyDoc.DocumentType <> 12291 Then ' kAssemblyDocumentObject
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

    ' Ensure we have Sheet 2
    Dim targetSheet, sheetNum
    sheetNum = 2
    If invDoc.Sheets.Count < 2 Then
        LogMessage "Creating Sheet 2..."
        Set targetSheet = invDoc.Sheets.Add()
        targetSheet.Name = "Sheet:2"
        LogMessage "Created Sheet 2"
    Else
        Set targetSheet = invDoc.Sheets.Item(2)
        LogMessage "Using existing Sheet 2"
    End If

    ' Place parts on sheets
    LogMessage ""
    LogMessage "=== PLACING PART VIEWS ==="

    Dim viewsPlaced, viewsFailed
    viewsPlaced = 0
    viewsFailed = 0

    Dim currentSheet, currentSheetIndex
    currentSheetIndex = 2
    Set currentSheet = targetSheet

    Dim viewIndexOnSheet
    viewIndexOnSheet = 0

    For i = 0 To UBound(partList)
        Dim partPath
        partPath = partList(i)

        ' Check if we need a new sheet
        If viewIndexOnSheet >= VIEWS_PER_SHEET Then
            currentSheetIndex = currentSheetIndex + 1
            viewIndexOnSheet = 0

            If invDoc.Sheets.Count < currentSheetIndex Then
                LogMessage "Creating Sheet " & currentSheetIndex & "..."
                Set currentSheet = invDoc.Sheets.Add()
                currentSheet.Name = "Sheet:" & currentSheetIndex
                LogMessage "Created Sheet " & currentSheetIndex
            Else
                Set currentSheet = invDoc.Sheets.Item(currentSheetIndex)
                LogMessage "Switching to existing Sheet " & currentSheetIndex
            End If
        End If

        ' Create base view for this part
        Dim partFileName
        partFileName = fso.GetFileName(partPath)
        LogMessage ""
        LogMessage "[" & (i + 1) & "/" & partCount & "] Placing: " & partFileName

        If PlacePartView(currentSheet, partPath, viewIndexOnSheet) Then
            viewsPlaced = viewsPlaced + 1
            viewIndexOnSheet = viewIndexOnSheet + 1
        Else
            viewsFailed = viewsFailed + 1
        End If
    Next

    ' Save the drawing
    LogMessage ""
    LogMessage "Saving drawing..."
    invDoc.Save2 True
    LogMessage "SUCCESS: Drawing saved"

    ' Final report
    LogMessage ""
    LogMessage "=== PLACEMENT COMPLETE ==="
    LogMessage "Total parts found: " & partCount
    LogMessage "Views placed: " & viewsPlaced
    LogMessage "Views failed: " & viewsFailed
    LogMessage "Sheets used: " & (currentSheetIndex - 1)

    logFile.Close

    ' Show message to user
    Dim msg
    msg = "Part Placement Complete!" & vbCrLf & vbCrLf
    msg = msg & "Assembly: " & assemblyDoc.DisplayName & vbCrLf
    msg = msg & "Parts found: " & partCount & vbCrLf
    msg = msg & "Views placed: " & viewsPlaced & vbCrLf
    If viewsFailed > 0 Then
        msg = msg & "Views failed: " & viewsFailed & vbCrLf & vbCrLf
        msg = msg & "Check log file for details:" & vbCrLf & logPath
    End If
    msg = msg & vbCrLf & "Sheets used: " & (currentSheetIndex - 1)

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

' Function to place a single part view on a sheet
Function PlacePartView(sheet, partPath, viewIndex)
    On Error Resume Next
    PlacePartView = False

    ' Check if part is already open in Inventor (faster, more reliable)
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

    ' If not already open, open it (visible to ensure proper loading)
    Dim weOpenedIt
    weOpenedIt = False

    If partDoc Is Nothing Then
        Set partDoc = invApp.Documents.Open(partPath, True)  ' Open visible!
        weOpenedIt = True

        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "  ERROR: Could not open part: " & Err.Description
            Err.Clear
            Exit Function
        End If
    End If

    ' Calculate position based on view index (grid layout)
    Dim col, row
    col = viewIndex Mod 3  ' 3 columns
    row = viewIndex \ 3    ' 4 rows per sheet

    Dim x, y
    x = MARGIN_LEFT + (col * COLUMN_SPACING)
    y = MARGIN_TOP - (row * ROW_SPACING)

    LogMessage "  Position: Column " & (col + 1) & ", Row " & (row + 1) & " (" & x & "cm, " & y & "cm)"

    ' Create point for view placement
    Dim position
    Set position = invApp.TransientGeometry.CreatePoint2d(x, y)

    ' Activate the sheet first (required when adding views to non-active sheets)
    sheet.Activate

    ' Count views before adding
    Dim viewsBefore
    viewsBefore = sheet.DrawingViews.Count

    ' Try simple call first with just required parameters
    Dim baseView
    Set baseView = sheet.DrawingViews.AddBaseView(partDoc, position, 0.1)

    ' Verify view was actually added
    Dim viewsAfter
    viewsAfter = sheet.DrawingViews.Count

    If Err.Number = 0 And Not baseView Is Nothing And viewsAfter > viewsBefore Then
        LogMessage "  SUCCESS: View created at 1:10 (views: " & viewsBefore & " -> " & viewsAfter & ")"
    Else
        LogMessage "  ERROR: Could not create base view (Err #" & Err.Number & "): " & Err.Description
        Err.Clear
        ' Only close if we opened it
        If weOpenedIt Then
            partDoc.Close
        End If
        Exit Function
    End If

    ' Only close if we opened it (not if it was already open)
    If weOpenedIt Then
        partDoc.Close
    End If

    PlacePartView = True
End Function

Sub LogMessage(msg)
    Dim timestamp
    timestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
                Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2)

    If Not logFile Is Nothing Then
        logFile.WriteLine timestamp & " - " & msg
    End If
End Sub
