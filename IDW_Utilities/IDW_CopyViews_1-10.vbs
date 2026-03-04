' ==================================================================================
' IDW COPY VIEWS TO NEXT SHEET - 1:10 SCALE
' ==================================================================================
' Copies views from selected sheet to next sheet, then changes scale to 1:10
' Uses CopyTo method (AddBaseView is broken in Inventor 2026)
'
' Features:
' - User selects source sheet
' - Copies first view from source as template
' - Places all parts on NEXT sheet using CopyTo
' - Sets scale to 1:10 for all copied views
' - Creates new sheets as needed
' ==================================================================================

Option Explicit

' Global variables
Dim invApp, invDoc
Dim fso, logFile
Dim logPath

' Configuration
Const TARGET_SCALE = 0.1       ' 1:10 scale
Const COLUMN_SPACING = 8       ' cm between columns
Const ROW_SPACING = 6          ' cm between rows
Const MARGIN_LEFT = 2          ' cm left margin
Const MARGIN_TOP = 27          ' cm from top

' Main execution
Main

Sub Main()
    On Error Resume Next

    ' Setup logging
    Set fso = CreateObject("Scripting.FileSystemObject")
    logPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\IDW_CopyViews_1-10.log"
    Set logFile = fso.CreateTextFile(logPath, True)

    LogMessage "=== IDW COPY VIEWS TO NEXT SHEET - 1:10 SCALE ==="

    ' Connect to Inventor
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        MsgBox "ERROR: Could not connect to Inventor.", vbCritical, "Error"
        WScript.Quit
    End If

    ' Get active document
    Set invDoc = invApp.ActiveDocument
    If invDoc Is Nothing Then
        MsgBox "ERROR: No active document.", vbCritical, "Error"
        WScript.Quit
    End If

    If invDoc.DocumentType <> 12294 And invDoc.DocumentType <> 12292 Then
        MsgBox "ERROR: Not a drawing file.", vbCritical, "Error"
        WScript.Quit
    End If
    
    LogMessage "Drawing: " & invDoc.DisplayName

    ' Check sheets
    If invDoc.Sheets.Count < 1 Then
        MsgBox "ERROR: No sheets!", vbCritical, "Error"
        WScript.Quit
    End If

    ' Let user select source sheet
    Dim sourceSheet, sourceSheetIndex
    sourceSheetIndex = SelectSheetToScan(invDoc)
    
    If sourceSheetIndex = 0 Then
        LogMessage "User cancelled"
        WScript.Quit
    End If
    
    Set sourceSheet = invDoc.Sheets.Item(sourceSheetIndex)
    LogMessage "Source: " & sourceSheet.Name & " (" & sourceSheet.DrawingViews.Count & " views)"

    If sourceSheet.DrawingViews.Count = 0 Then
        MsgBox "ERROR: Selected sheet has no views!", vbExclamation, "Error"
        WScript.Quit
    End If

    ' Get assembly from first view
    Dim firstView, assemblyDoc
    Set firstView = sourceSheet.DrawingViews.Item(1)
    Set assemblyDoc = firstView.ReferencedDocumentDescriptor.ReferencedDocument

    If assemblyDoc Is Nothing Then
        MsgBox "ERROR: No assembly reference found.", vbCritical, "Error"
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
        MsgBox "ERROR: No parts found!", vbExclamation, "Error"
        WScript.Quit
    End If

    ' Get/create target sheet (next after source)
    Dim targetSheet, targetSheetIndex
    targetSheetIndex = sourceSheetIndex + 1
    
    If invDoc.Sheets.Count < targetSheetIndex Then
        LogMessage "Creating Sheet " & targetSheetIndex & "..."
        Set targetSheet = invDoc.Sheets.Add()
        targetSheet.Name = "Sheet:" & targetSheetIndex
    Else
        Set targetSheet = invDoc.Sheets.Item(targetSheetIndex)
        LogMessage "Using Sheet " & targetSheetIndex
    End If

    ' Copy views from SOURCE SHEET for each part (not from template)
    LogMessage ""
    LogMessage "=== PLACING PART VIEWS AT 1:10 ==="

    Dim viewsPlaced, viewsFailed, viewIndex
    viewsPlaced = 0
    viewsFailed = 0
    viewIndex = 0

    Dim i, partPath
    For i = 0 To UBound(partList)
        partPath = partList(i)

        LogMessage ""
        LogMessage "[" & (i + 1) & "/" & partCount & "] " & fso.GetFileName(partPath)

        ' Copy from the ORIGINAL source view each time
        If PlacePartByCopy(targetSheet, firstView, partPath, viewIndex) Then
            viewsPlaced = viewsPlaced + 1
            viewIndex = viewIndex + 1
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

' Copy template view and set scale to 1:10
Function PlacePartByCopy(sheet, templateView, partPath, viewIndex)
    On Error Resume Next
    PlacePartByCopy = False

    ' Calculate position (using same grid as before)
    Dim col, row, x, y
    col = viewIndex Mod 3
    row = viewIndex \ 3
    x = MARGIN_LEFT + (col * COLUMN_SPACING)
    y = MARGIN_TOP - (row * ROW_SPACING)

    LogMessage "  Position: (" & x & ", " & y & ")"

    ' Copy the template view
    Dim newView
    Set newView = templateView.CopyTo(sheet)
    
    If Err.Number <> 0 Or newView Is Nothing Then
        LogMessage "  ERROR: CopyTo failed"
        Err.Clear
        Exit Function
    End If
    
    LogMessage "  View copied: " & newView.Name

    ' Position the view
    Dim position
    Set position = invApp.TransientGeometry.CreatePoint2d(x, y)
    newView.Position = position
    LogMessage "  Position set"

    ' Change scale to 1:10
    newView.Scale = TARGET_SCALE
    
    If Err.Number <> 0 Then
        LogMessage "  ERROR: Could not set scale: " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    LogMessage "  Scale set to 1:10"
    LogMessage "  SUCCESS"
    
    PlacePartByCopy = True
End Function

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
    
    promptText = "Select sheet with assembly view:" & vbCrLf & vbCrLf & sheetList & vbCrLf & "Enter number:"
    
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

    If assemblyDoc.DocumentType = 12290 Then
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

        If refDoc.DocumentType = 12290 Then
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
        ElseIf refDoc.DocumentType = 12291 Then
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

Sub LogMessage(msg)
    If Not logFile Is Nothing Then
        logFile.WriteLine Now & " - " & msg
    End If
End Sub
