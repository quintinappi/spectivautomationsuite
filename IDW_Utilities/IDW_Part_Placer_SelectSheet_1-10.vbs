' ==================================================================================
' IDW PART PLACER - SELECT SHEET + 1:10 SCALE
' ==================================================================================
' User selects which sheet to scan, then places parts on next sheet at 1:10 scale
'
' Features:
' - User selects source sheet from list
' - Extracts assembly from selected sheet
' - Places all parts on NEXT sheet at 1:10 scale (0.1)
' - Creates new sheets as needed
' ==================================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim logPath

' Configuration
Const VIEWS_PER_SHEET = 12  ' Max views per sheet
Const COLUMN_SPACING = 8    ' cm between columns
Const ROW_SPACING = 6       ' cm between rows  
Const MARGIN_LEFT = 2       ' cm left margin
Const MARGIN_TOP = 27       ' cm from top (Inventor uses top-down for this script)

' Main execution
Main

Sub Main()
    On Error Resume Next

    ' Setup logging
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\IDW_SelectSheet_1-10.log"
    Set logFile = fso.CreateTextFile(logPath, True)

    LogMessage "=== IDW PART PLACER - SELECT SHEET + 1:10 SCALE ==="

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "ERROR: Could not connect to Inventor." & vbCrLf & _
               "Please make sure Inventor is running with an IDW file open.", vbCritical, "Inventor Not Found"
        WScript.Quit
    End If

    ' Get active document
    Set invDoc = invApp.ActiveDocument
    If invDoc Is Nothing Then
        MsgBox "ERROR: No active document found.", vbCritical, "No Document"
        WScript.Quit
    End If

    ' Check if it's a drawing document
    If invDoc.DocumentType <> 12294 And invDoc.DocumentType <> 12292 Then
        MsgBox "ERROR: Active document is not a drawing file.", vbCritical, "Wrong Document Type"
        WScript.Quit
    End If
    
    LogMessage "Drawing: " & invDoc.DisplayName

    ' Check sheets exist
    If invDoc.Sheets.Count < 1 Then
        MsgBox "ERROR: Drawing has no sheets!", vbCritical, "No Sheets"
        WScript.Quit
    End If

    ' Let user select which sheet to scan
    Dim sourceSheet, sourceSheetIndex
    sourceSheetIndex = SelectSheetToScan(invDoc)
    
    If sourceSheetIndex = 0 Then
        LogMessage "User cancelled"
        WScript.Quit
    End If
    
    Set sourceSheet = invDoc.Sheets.Item(sourceSheetIndex)
    LogMessage "Source sheet: " & sourceSheet.Name & " (" & sourceSheet.DrawingViews.Count & " views)"

    If sourceSheet.DrawingViews.Count = 0 Then
        MsgBox "ERROR: Selected sheet has no views!", vbExclamation, "No Views"
        WScript.Quit
    End If

    ' Get the first view's assembly
    Dim firstView, assemblyDoc
    Set firstView = sourceSheet.DrawingViews.Item(1)
    Set assemblyDoc = firstView.ReferencedDocumentDescriptor.ReferencedDocument

    If assemblyDoc Is Nothing Then
        MsgBox "ERROR: Could not get referenced document from view.", vbCritical, "No Reference"
        WScript.Quit
    End If
    
    LogMessage "Assembly: " & assemblyDoc.DisplayName

    ' Collect all unique parts
    Dim partList
    partList = GetAllParts(assemblyDoc)

    Dim partCount
    partCount = UBound(partList) + 1
    LogMessage "Found " & partCount & " unique part(s)"

    If partCount = 0 Then
        MsgBox "ERROR: No parts found in the assembly!", vbExclamation, "No Parts"
        WScript.Quit
    End If

    ' Determine target sheet (next sheet after source)
    Dim targetSheet, targetSheetIndex
    targetSheetIndex = sourceSheetIndex + 1
    
    If invDoc.Sheets.Count < targetSheetIndex Then
        LogMessage "Creating Sheet " & targetSheetIndex & "..."
        Set targetSheet = invDoc.Sheets.Add()
        targetSheet.Name = "Sheet:" & targetSheetIndex
    Else
        Set targetSheet = invDoc.Sheets.Item(targetSheetIndex)
        LogMessage "Using existing Sheet " & targetSheetIndex
    End If

    ' Place parts
    LogMessage ""
    LogMessage "=== PLACING PARTS AT 1:10 SCALE ==="

    Dim viewsPlaced, viewsFailed, viewIndexOnSheet
    viewsPlaced = 0
    viewsFailed = 0
    viewIndexOnSheet = 0

    Dim i, partPath
    For i = 0 To UBound(partList)
        partPath = partList(i)

        LogMessage ""
        LogMessage "[" & (i + 1) & "/" & partCount & "] " & fso.GetFileName(partPath)

        If PlacePartViewAtScale(targetSheet, partPath, viewIndexOnSheet) Then
            viewsPlaced = viewsPlaced + 1
            viewIndexOnSheet = viewIndexOnSheet + 1
        Else
            viewsFailed = viewsFailed + 1
        End If
    Next

    ' Save
    invDoc.Save2 True
    LogMessage ""
    LogMessage "=== COMPLETE ==="
    LogMessage "Placed: " & viewsPlaced & ", Failed: " & viewsFailed

    logFile.Close

    MsgBox "Complete!" & vbCrLf & vbCrLf & _
           "Source: " & sourceSheet.Name & vbCrLf & _
           "Target: " & targetSheet.Name & vbCrLf & _
           "Parts: " & partCount & vbCrLf & _
           "Placed at 1:10: " & viewsPlaced, vbInformation, "Done"

End Sub

' Let user select which sheet to scan
Function SelectSheetToScan(drawingDoc)
    On Error Resume Next
    
    Dim sheetCount, i, sheetList, promptText
    sheetCount = drawingDoc.Sheets.Count
    
    sheetList = ""
    For i = 1 To sheetCount
        Dim s
        Set s = drawingDoc.Sheets.Item(i)
        sheetList = sheetList & i & ". " & s.Name & " (" & s.DrawingViews.Count & " views)" & vbCrLf
    Next
    
    promptText = "Select sheet with the assembly view:" & vbCrLf & vbCrLf & sheetList & vbCrLf & "Enter number:"
    
    Dim userInput
    userInput = InputBox(promptText, "Select Source Sheet", "1")
    
    If userInput = "" Or Not IsNumeric(userInput) Then
        SelectSheetToScan = 0
        Exit Function
    End If
    
    Dim selectedIndex
    selectedIndex = CInt(userInput)
    
    If selectedIndex < 1 Or selectedIndex > sheetCount Then
        MsgBox "Invalid sheet number.", vbExclamation, "Error"
        SelectSheetToScan = 0
        Exit Function
    End If
    
    Dim selectedSheet
    Set selectedSheet = drawingDoc.Sheets.Item(selectedIndex)
    If selectedSheet.DrawingViews.Count = 0 Then
        MsgBox "Selected sheet has no views.", vbExclamation, "Error"
        SelectSheetToScan = 0
        Exit Function
    End If
    
    SelectSheetToScan = selectedIndex
End Function

' Get all unique parts from assembly recursively
Function GetAllParts(assemblyDoc)
    Dim parts()
    ReDim parts(0)
    Dim count
    count = 0

    If assemblyDoc.DocumentType = 12290 Then ' kPartDocumentObject
        parts(0) = assemblyDoc.FullFileName
        GetAllParts = parts
        Exit Function
    End If

    Dim occurrences
    Set occurrences = assemblyDoc.ComponentDefinition.Occurrences

    Dim i, occ, refDoc, partPath, isDuplicate, j, k

    For i = 1 To occurrences.Count
        Set occ = occurrences.Item(i)
        Set refDoc = occ.ReferencedDocumentDescriptor.ReferencedDocument

        If refDoc.DocumentType = 12290 Then ' Part
            partPath = refDoc.FullFileName
            isDuplicate = False
            For j = 0 To count - 1
                If LCase(parts(j)) = LCase(partPath) Then
                    isDuplicate = True
                    Exit For
                End If
            Next
            If Not isDuplicate Then
                If count > 0 Then ReDim Preserve parts(count)
                parts(count) = partPath
                count = count + 1
            End If
        ElseIf refDoc.DocumentType = 12291 Then ' Sub-assembly
            Dim subParts
            subParts = GetAllParts(refDoc)
            For k = LBound(subParts) To UBound(subParts)
                partPath = subParts(k)
                isDuplicate = False
                For j = 0 To count - 1
                    If LCase(parts(j)) = LCase(partPath) Then
                        isDuplicate = True
                        Exit For
                    End If
                Next
                If Not isDuplicate Then
                    If count > 0 Then ReDim Preserve parts(count)
                    parts(count) = partPath
                    count = count + 1
                End If
            Next
        End If
    Next

    If count > 0 Then ReDim Preserve parts(count - 1)
    GetAllParts = parts
End Function

' Place a single part view at 1:10 scale - EXACT copy from working script
Function PlacePartViewAtScale(sheet, partPath, viewIndex)
    On Error Resume Next
    PlacePartViewAtScale = False

    ' Find or open part document
    Dim partDoc, i, doc, weOpenedIt
    Set partDoc = Nothing
    weOpenedIt = False

    For i = 1 To invApp.Documents.Count
        Set doc = invApp.Documents.Item(i)
        If LCase(doc.FullFileName) = LCase(partPath) Then
            Set partDoc = doc
            LogMessage "  Using already-open part"
            Exit For
        End If
    Next

    If partDoc Is Nothing Then
        Set partDoc = invApp.Documents.Open(partPath, True)
        weOpenedIt = True
        If Err.Number <> 0 Or partDoc Is Nothing Then
            LogMessage "  ERROR: Could not open part"
            Err.Clear
            Exit Function
        End If
    End If

    ' Calculate position (original script uses top-down Y)
    Dim col, row, x, y
    col = viewIndex Mod 3
    row = viewIndex \ 3
    x = MARGIN_LEFT + (col * COLUMN_SPACING)
    y = MARGIN_TOP - (row * ROW_SPACING)

    LogMessage "  Position: (" & x & ", " & y & ")"

    ' Create position point
    Dim position
    Set position = invApp.TransientGeometry.CreatePoint2d(x, y)

    ' Activate sheet
    sheet.Activate

    ' Count before
    Dim viewsBefore
    viewsBefore = sheet.DrawingViews.Count

    ' THE CRITICAL LINE - EXACTLY as in working script, just with 0.1 instead of 1
    Dim baseView
    Set baseView = sheet.DrawingViews.AddBaseView(partDoc, position, 0.1)

    ' Check result
    Dim viewsAfter
    viewsAfter = sheet.DrawingViews.Count

    If Err.Number = 0 And Not baseView Is Nothing And viewsAfter > viewsBefore Then
        LogMessage "  SUCCESS at 1:10"
        PlacePartViewAtScale = True
    Else
        LogMessage "  ERROR: Err=" & Err.Number & ", baseView=" & TypeName(baseView) & ", views=" & viewsBefore & "->" & viewsAfter
        Err.Clear
        PlacePartViewAtScale = False
    End If

    If weOpenedIt Then partDoc.Close
End Function

Sub LogMessage(msg)
    If Not logFile Is Nothing Then
        logFile.WriteLine Now & " - " & msg
    End If
End Sub
