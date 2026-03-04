' ==============================================================================
' POPULATE DWG REF FROM PARTS LISTS
' ==============================================================================
' - Scans Parts Lists on non-DXF sheets
' - Finds where referenced models are actually placed (DrawingViews on non-DXF)
' - Writes DWG. REF. column values as IDWNAME-01/02 format
' - Writes model user iProperties (DWG REF aliases)
' - Warns which PARTS (.ipt) are not detailed (no view placement)
' ==============================================================================

Option Explicit

Const kDrawingDocumentObject = 12292
Const FORCE_TEST_MODE = False
Const FORCE_TEST_VALUE = "test"
Const WRITE_PARTS_LIST_CELLS_DIRECT = True
Const ENABLE_PLACE_FIRST_UNDETAILED_DEFAULT = False
Const DEFAULT_DWG_REF_ONLY_MODE = True
Const TARGET_SHEET_NAME_DEFAULT = "Sheet:5"
Const TARGET_ABSOLUTE_SHEET_INDEX_DEFAULT = 5
Const TARGET_NON_DXF_SHEET_INDEX_FALLBACK = 5
Const AUTO_PLACE_SCALE = 0.1
Const kFrontViewOrientation = 10764
Const kHiddenLineRemovedDrawingViewStyle = 32258
Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Const AUTO_LAYOUT_COLUMNS = 3
Const AUTO_LAYOUT_ROWS_PER_SHEET = 3
Const ENABLE_DEBUG_ROW_MATCH_LOGS = False

Dim g_LogFile
Dim g_LogPath

Call Main()

Sub Main()
    On Error Resume Next

    StartLogging
    LogMessage "=== DWG REF UPDATER START ==="
    LogMessage "FORCE_TEST_MODE=" & CStr(FORCE_TEST_MODE)

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: Inventor is not running"
        MsgBox "Inventor is not running.", vbCritical, "DWG REF Updater"
        StopLogging
        Exit Sub
    End If
    Err.Clear

    Dim drawingDoc
    Set drawingDoc = invApp.ActiveDocument
    If drawingDoc Is Nothing Then
        LogMessage "ERROR: No active document found"
        MsgBox "No active document found.", vbCritical, "DWG REF Updater"
        StopLogging
        Exit Sub
    End If

    If drawingDoc.DocumentType <> kDrawingDocumentObject Then
        LogMessage "ERROR: Active document is not IDW"
        MsgBox "Active document is not an IDW drawing.", vbCritical, "DWG REF Updater"
        StopLogging
        Exit Sub
    End If

    Dim idwBaseName
    idwBaseName = GetBaseName(SafeText(drawingDoc.FullFileName))
    LogMessage "Drawing: " & SafeText(drawingDoc.FullFileName)

    Dim modelSheets
    Set modelSheets = CreateObject("Scripting.Dictionary")
    modelSheets.CompareMode = 1

    Dim modelDocs
    Set modelDocs = CreateObject("Scripting.Dictionary")
    modelDocs.CompareMode = 1

    Dim modelPartNumbers
    Set modelPartNumbers = CreateObject("Scripting.Dictionary")
    modelPartNumbers.CompareMode = 1

    Dim sheetCount, partsListCount, rowCount
    sheetCount = 0
    partsListCount = 0
    rowCount = 0

    CollectModelsFromPartsLists drawingDoc, modelSheets, modelDocs, modelPartNumbers, sheetCount, partsListCount, rowCount
    CollectModelViewPlacements drawingDoc, modelSheets
    LogMessage "Collected models: " & CStr(modelSheets.Count)

    Dim dwgRefOnlyMode
    dwgRefOnlyMode = ParseOnOffArg("dwgrefonly", DEFAULT_DWG_REF_ONLY_MODE)

    Dim autoPlaceEnabled
    autoPlaceEnabled = ParseOnOffArg("autoplace", ENABLE_PLACE_FIRST_UNDETAILED_DEFAULT)
    If WScript.Arguments.Named.Exists("autoplace") = False Then
        autoPlaceEnabled = Not dwgRefOnlyMode
    End If

    Dim autoPlaceMessage
    autoPlaceMessage = "Disabled"

    LogMessage "DWG_REF_ONLY_MODE=" & CStr(dwgRefOnlyMode)
    LogMessage "AUTO_PLACE_ENABLED=" & CStr(autoPlaceEnabled)

    If autoPlaceEnabled Then
        Dim selectedSheetName
        selectedSheetName = GetPlacementTargetSheetName(drawingDoc)
        If selectedSheetName = "" Then
            autoPlaceMessage = "Placement cancelled by user"
            LogMessage autoPlaceMessage
        Else
            Dim preMsg
            preMsg = "Auto placement is enabled." & vbCrLf & vbCrLf & _
                     "An open sheet for part placement will be used/created before scanning continues." & vbCrLf & _
                     "Target sheet: " & selectedSheetName & vbCrLf & vbCrLf & _
                     "Continue?"
            If MsgBox(preMsg, vbYesNo + vbQuestion, "DWG REF Updater") = vbYes Then
                Dim placedCount
                placedCount = 0
                TryPlaceUndetailedParts invApp, drawingDoc, modelSheets, modelDocs, modelPartNumbers, selectedSheetName, placedCount, autoPlaceMessage
                LogMessage autoPlaceMessage
            Else
                autoPlaceMessage = "Placement cancelled by user"
                LogMessage autoPlaceMessage
            End If
        End If
    Else
        LogMessage "Auto-place disabled by mode/argument"
    End If

    Dim undetailedCount, undetailedReport
    GetUndetailedPartsReport modelSheets, modelPartNumbers, undetailedCount, undetailedReport
    LogMessage "Undetailed parts found: " & CStr(undetailedCount)

    Dim partsListCellsUpdated
    partsListCellsUpdated = 0
    If WRITE_PARTS_LIST_CELLS_DIRECT Then
        InjectPartsListDwgRefCells drawingDoc, modelSheets, idwBaseName, partsListCellsUpdated
        LogMessage "Parts list DWG REF cells updated: " & CStr(partsListCellsUpdated)
    End If

    If modelSheets.Count = 0 Then
        LogMessage "ERROR: No referenced models found in parts lists"
        MsgBox "No referenced parts/assemblies found in any parts list.", vbExclamation, "DWG REF Updater"
        StopLogging
        Exit Sub
    End If

    Dim totalUpdated, totalErrors
    totalUpdated = 0
    totalErrors = 0
    UpdateModelDwgRefs modelSheets, modelDocs, idwBaseName, totalUpdated, totalErrors

    Dim drawingSummary
    drawingSummary = BuildDrawingSummary(modelSheets, modelPartNumbers, idwBaseName)
    SetUserDefinedProperty drawingDoc, "DWG REF", drawingSummary

    drawingDoc.Update
    drawingDoc.Save

    LogMessage "Sheets scanned: " & CStr(sheetCount)
    LogMessage "Parts lists scanned: " & CStr(partsListCount)
    LogMessage "Rows scanned: " & CStr(rowCount)
    LogMessage "Models updated: " & CStr(totalUpdated)
    LogMessage "Model update errors: " & CStr(totalErrors)
    LogMessage "Log path: " & g_LogPath

    Dim summary
    summary = "DWG REF model updates completed." & vbCrLf & vbCrLf & _
              "Sheets scanned (non-DXF): " & sheetCount & vbCrLf & _
              "Parts lists scanned: " & partsListCount & vbCrLf & _
              "Rows scanned: " & rowCount & vbCrLf & _
              "Unique models found: " & modelSheets.Count & vbCrLf & _
              "Undetailed parts: " & undetailedCount & vbCrLf & _
              "Auto-place missing parts: " & autoPlaceMessage & vbCrLf & _
              "Parts list cells updated: " & partsListCellsUpdated & vbCrLf & _
              "Models updated: " & totalUpdated & vbCrLf & _
              "Model update errors: " & totalErrors & vbCrLf & _
              "Log file: " & g_LogPath

    If undetailedCount > 0 Then
        summary = summary & vbCrLf & vbCrLf & _
                  "WARNING - Undetailed parts (not placed on non-DXF sheets):" & vbCrLf & _
                  undetailedReport
    End If

    MsgBox summary, vbInformation, "DWG REF Updater"
    StopLogging
End Sub

Sub TryPlaceUndetailedParts(invApp, drawingDoc, modelSheets, modelDocs, modelPartNumbers, targetSheetName, ByRef placedCount, ByRef resultMessage)
    On Error Resume Next

    placedCount = 0
    resultMessage = "Skipped"

    Dim missingParts
    Set missingParts = FindUndetailedPartPaths(modelSheets)
    If missingParts Is Nothing Or missingParts.Count = 0 Then
        resultMessage = "No undetailed parts found"
        Exit Sub
    End If

    Dim baseSheet
    Set baseSheet = ResolveTargetSheetForPlacement(drawingDoc, targetSheetName, TARGET_ABSOLUTE_SHEET_INDEX_DEFAULT, TARGET_NON_DXF_SHEET_INDEX_FALLBACK)
    If baseSheet Is Nothing Then
        resultMessage = "Target sheet not found (name=" & targetSheetName & ")"
        Exit Sub
    End If

    Dim maxSlots
    maxSlots = AUTO_LAYOUT_COLUMNS * AUTO_LAYOUT_ROWS_PER_SHEET

    Dim currentSheet
    Set currentSheet = baseSheet

    Dim slotOnSheet
    slotOnSheet = 0

    Dim extraSheetCount
    extraSheetCount = 0

    Dim totalViewsPlaced
    totalViewsPlaced = 0

    Dim failures
    failures = ""

    Dim missingKeys
    missingKeys = missingParts.Keys
    SortStringArray missingKeys

    Dim i
    For i = 0 To UBound(missingKeys)
        Dim partPath
        partPath = CStr(missingKeys(i))

        Dim canProcess
        canProcess = True

        Dim modelDoc
        Set modelDoc = Nothing
        If modelDocs.Exists(partPath) Then Set modelDoc = modelDocs(partPath)

        If modelDoc Is Nothing Then
            Err.Clear
            Set modelDoc = invApp.Documents.Open(partPath, True)
            If Err.Number <> 0 Or modelDoc Is Nothing Then
                AppendFailure failures, "Open failed: " & GetBaseName(partPath)
                Err.Clear
                canProcess = False
            End If
        End If

        If canProcess Then
        Dim isPlate
        isPlate = IsSheetMetalPart(modelDoc)

        Dim requiredSlots
        requiredSlots = 1
        If isPlate Then requiredSlots = 2

        If slotOnSheet + requiredSlots > maxSlots Then
            extraSheetCount = extraSheetCount + 1
            Set currentSheet = CreateAutoPlacementSheet(drawingDoc, baseSheet, extraSheetCount)
            slotOnSheet = 0
        End If

        currentSheet.Activate

        Dim viewsPlacedForPart
        viewsPlacedForPart = 0

        If isPlate Then
            EnsureFlatPatternExists modelDoc

            Dim ptFolded, ptFlat
            Set ptFolded = GetAutoPlacementPoint(invApp, currentSheet, slotOnSheet)
            Set ptFlat = GetAutoPlacementPoint(invApp, currentSheet, slotOnSheet + 1)

            Dim foldedView, flatView, foldErr, flatErr
            Set foldedView = Nothing
            Set flatView = Nothing
            foldErr = ""
            flatErr = ""

            AddSheetMetalViewWithOption invApp, currentSheet, modelDoc, ptFolded, AUTO_PLACE_SCALE, True, foldedView, foldErr
            AddSheetMetalViewWithOption invApp, currentSheet, modelDoc, ptFlat, AUTO_PLACE_SCALE, False, flatView, flatErr

            If Not foldedView Is Nothing Then
                viewsPlacedForPart = viewsPlacedForPart + 1
                totalViewsPlaced = totalViewsPlaced + 1
            End If
            If Not flatView Is Nothing Then
                viewsPlacedForPart = viewsPlacedForPart + 1
                totalViewsPlaced = totalViewsPlaced + 1
            End If

            If foldedView Is Nothing And flatView Is Nothing Then
                AppendFailure failures, "Plate failed: " & GetBaseName(partPath) & " | folded=" & foldErr & " | flat=" & flatErr
            End If

            slotOnSheet = slotOnSheet + 2
        Else
            Dim pt
            Set pt = GetAutoPlacementPoint(invApp, currentSheet, slotOnSheet)

            Dim newView, addErr
            Set newView = Nothing
            addErr = ""
            AddBaseViewWithFallback currentSheet, modelDoc, pt, AUTO_PLACE_SCALE, newView, addErr

            If Not newView Is Nothing Then
                viewsPlacedForPart = viewsPlacedForPart + 1
                totalViewsPlaced = totalViewsPlaced + 1
            Else
                AppendFailure failures, "Part failed: " & GetBaseName(partPath) & " | " & addErr
            End If

            slotOnSheet = slotOnSheet + 1
        End If

        If viewsPlacedForPart > 0 Then
            placedCount = placedCount + 1

            Dim sheetNo
            sheetNo = GetSheetNumberToken(drawingDoc, currentSheet)
            If sheetNo = "" Then sheetNo = "05"
            If modelSheets.Exists(partPath) Then
                If Not modelSheets(partPath).Exists(sheetNo) Then
                    modelSheets(partPath).Add sheetNo, True
                End If
            End If
        End If
        End If

    Next

    If placedCount = 0 Then
        resultMessage = "No missing parts were placed"
    Else
        resultMessage = "Placed " & CStr(placedCount) & " missing parts (" & CStr(totalViewsPlaced) & " views)"
        If extraSheetCount > 0 Then
            resultMessage = resultMessage & " across " & CStr(extraSheetCount + 1) & " sheets"
        End If
        If failures <> "" Then
            resultMessage = resultMessage & " | Some failed: " & failures
        End If
    End If
End Sub

Sub AddSheetMetalViewWithOption(invApp, targetSheet, modelDoc, pt, scaleValue, foldedFlag, ByRef outView, ByRef outErr)
    On Error Resume Next

    Set outView = Nothing
    outErr = ""

    Dim scales
    scales = Array(scaleValue, 0.2, 0.05)

    Dim i
    For i = 0 To UBound(scales)
        Dim opts
        Set opts = CreateSheetMetalAddOptions(invApp, foldedFlag)

        Dim beforeCount
        beforeCount = targetSheet.DrawingViews.Count

        Err.Clear
        targetSheet.DrawingViews.AddBaseView modelDoc, pt, CDbl(scales(i)), kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle, "", Nothing, opts
        If Err.Number = 0 And targetSheet.DrawingViews.Count > beforeCount Then
            Set outView = targetSheet.DrawingViews.Item(targetSheet.DrawingViews.Count)
            If foldedFlag Then
                outErr = "OK folded scale=" & CStr(scales(i))
            Else
                outErr = "OK flat scale=" & CStr(scales(i))
            End If
            Exit Sub
        Else
            If foldedFlag Then
                outErr = outErr & " folded(scale=" & CStr(scales(i)) & ")=" & Err.Description
            Else
                outErr = outErr & " flat(scale=" & CStr(scales(i)) & ")=" & Err.Description
            End If
            Err.Clear
        End If
    Next
End Sub

Function CreateSheetMetalAddOptions(invApp, foldedFlag)
    On Error Resume Next

    Set CreateSheetMetalAddOptions = Nothing

    Dim opts
    Set opts = invApp.TransientObjects.CreateNameValueMap
    If Err.Number <> 0 Or opts Is Nothing Then
        Err.Clear
        Exit Function
    End If

    opts.Add "SheetMetalFoldedModel", CBool(foldedFlag)
    Set CreateSheetMetalAddOptions = opts
End Function

Function IsSheetMetalPart(modelDoc)
    On Error Resume Next

    IsSheetMetalPart = False
    If modelDoc Is Nothing Then Exit Function

    If LCase(Right(SafeText(modelDoc.FullFileName), 4)) <> ".ipt" Then Exit Function

    Dim subtype
    subtype = UCase(SafeText(modelDoc.SubType))
    If subtype = UCase(kSheetMetalSubType) Then
        IsSheetMetalPart = True
    End If
End Function

Sub EnsureFlatPatternExists(modelDoc)
    On Error Resume Next

    If modelDoc Is Nothing Then Exit Sub
    If Not IsSheetMetalPart(modelDoc) Then Exit Sub

    Dim compDef
    Set compDef = modelDoc.ComponentDefinition
    If compDef Is Nothing Then Exit Sub

    Err.Clear
    If Not compDef.HasFlatPattern Then
        compDef.Unfold
        Err.Clear
    End If
End Sub

Sub AddBaseViewWithFallback(targetSheet, modelDoc, pt, scaleValue, ByRef outView, ByRef outErr)
    On Error Resume Next

    Set outView = Nothing
    outErr = "No attempts made"

    TryAddBaseView targetSheet, modelDoc, pt, scaleValue, outView, outErr
    If Not outView Is Nothing Then Exit Sub

    TryAddBaseView targetSheet, modelDoc, pt, 0.2, outView, outErr
    If Not outView Is Nothing Then Exit Sub

    TryAddBaseView targetSheet, modelDoc, pt, 0.05, outView, outErr
End Sub

Sub TryAddBaseView(targetSheet, modelDoc, pt, scaleValue, ByRef outView, ByRef outErr)
    On Error Resume Next

    Dim beforeCount
    beforeCount = targetSheet.DrawingViews.Count

    Err.Clear
    targetSheet.DrawingViews.AddBaseView modelDoc, pt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle
    If Err.Number = 0 And targetSheet.DrawingViews.Count > beforeCount Then
        Set outView = targetSheet.DrawingViews.Item(targetSheet.DrawingViews.Count)
        outErr = "OK: 5-arg scale=" & CStr(scaleValue)
        Exit Sub
    ElseIf Err.Number <> 0 Then
        outErr = "fiveArg(scale=" & CStr(scaleValue) & ")=" & Err.Description
    End If
    Err.Clear

    beforeCount = targetSheet.DrawingViews.Count
    targetSheet.DrawingViews.AddBaseView modelDoc, pt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle, "", Nothing, Nothing
    If Err.Number = 0 And targetSheet.DrawingViews.Count > beforeCount Then
        Set outView = targetSheet.DrawingViews.Item(targetSheet.DrawingViews.Count)
        outErr = "OK: 8-arg scale=" & CStr(scaleValue)
        Exit Sub
    ElseIf Err.Number <> 0 Then
        outErr = outErr & "; eightArg(scale=" & CStr(scaleValue) & ")=" & Err.Description
    End If
    Err.Clear
End Sub

Sub CollectModelsFromPartsLists(drawingDoc, modelSheets, modelDocs, modelPartNumbers, ByRef sheetCount, ByRef partsListCount, ByRef rowCount)
    On Error Resume Next

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            sheetCount = sheetCount + 1

            Dim partsList
            For Each partsList In sheet.PartsLists
                partsListCount = partsListCount + 1

                Dim row
                For Each row In partsList.PartsListRows
                    rowCount = rowCount + 1
                    CollectRowModels row, modelSheets, modelDocs, modelPartNumbers
                Next
            Next
        Else
            LogMessage "SKIP DXF SHEET (parts list scan): " & SafeText(sheet.Name)
        End If
    Next
End Sub

Function FindUndetailedPartPaths(modelSheets)
    On Error Resume Next

    Set FindUndetailedPartPaths = CreateObject("Scripting.Dictionary")
    FindUndetailedPartPaths.CompareMode = 1

    Dim keys
    keys = modelSheets.Keys
    SortStringArray keys

    Dim i
    For i = 0 To UBound(keys)
        Dim modelPath
        modelPath = CStr(keys(i))
        If LCase(Right(modelPath, 4)) = ".ipt" Then
            If modelSheets(modelPath).Count = 0 Then
                If Not FindUndetailedPartPaths.Exists(modelPath) Then
                    FindUndetailedPartPaths.Add modelPath, True
                End If
            End If
        End If
    Next
End Function

Sub AppendFailure(ByRef textValue, itemText)
    If textValue = "" Then
        textValue = itemText
    Else
        textValue = textValue & " || " & itemText
    End If
End Sub

Function GetPlacementTargetSheetName(drawingDoc)
    On Error Resume Next

    Dim defaultName
    defaultName = TARGET_SHEET_NAME_DEFAULT
    If defaultName = "" Then
        Dim fallbackSheet
        Set fallbackSheet = GetNonDxfSheetByIndex(drawingDoc, TARGET_NON_DXF_SHEET_INDEX_FALLBACK)
        If Not fallbackSheet Is Nothing Then defaultName = SafeText(fallbackSheet.Name)
    End If

    Dim sheetNames
    sheetNames = BuildNonDxfSheetNameArray(drawingDoc)
    If IsEmpty(sheetNames) Then
        GetPlacementTargetSheetName = ""
        Exit Function
    End If

    Dim picked
    picked = NormalizePickerValue(SelectFromDropdown("DWG REF Updater - Select Sheet", "Select target sheet for part placement:", sheetNames, defaultName))

    If picked = "" Then
        GetPlacementTargetSheetName = ""
        Exit Function
    End If

    If SheetExistsNonDxf(drawingDoc, picked) Then
        GetPlacementTargetSheetName = picked
    Else
        MsgBox "Sheet not found or is DXF: " & picked, vbExclamation, "DWG REF Updater"
        GetPlacementTargetSheetName = ""
    End If
End Function

Function BuildNonDxfSheetNameArray(drawingDoc)
    On Error Resume Next

    BuildNonDxfSheetNameArray = Empty

    Dim temp
    Set temp = CreateObject("Scripting.Dictionary")
    temp.CompareMode = 1

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            If Not temp.Exists(CStr(sheet.Name)) Then
                temp.Add CStr(sheet.Name), True
            End If
        End If
    Next

    If temp.Count = 0 Then Exit Function
    BuildNonDxfSheetNameArray = temp.Keys
End Function

Function SelectFromDropdown(windowTitle, promptText, optionsArray, defaultValue)
    On Error Resume Next

    SelectFromDropdown = ""
    If IsEmpty(optionsArray) Then Exit Function

    Dim fso, shell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    Dim tempPs
    tempPs = shell.ExpandEnvironmentStrings("%TEMP%") & "\inv_dropdown_" & Replace(CStr(Timer), ".", "") & ".ps1"

    Dim tf
    Set tf = fso.CreateTextFile(tempPs, True)

    tf.WriteLine "Add-Type -AssemblyName System.Windows.Forms"
    tf.WriteLine "Add-Type -AssemblyName System.Drawing"
    tf.WriteLine "$form = New-Object System.Windows.Forms.Form"
    tf.WriteLine "$form.Text = '" & EscapePsSingle(windowTitle) & "'"
    tf.WriteLine "$form.Size = New-Object System.Drawing.Size(520,170)"
    tf.WriteLine "$form.StartPosition = 'CenterScreen'"
    tf.WriteLine "$form.TopMost = $true"
    tf.WriteLine "$lbl = New-Object System.Windows.Forms.Label"
    tf.WriteLine "$lbl.Text = '" & EscapePsSingle(promptText) & "'"
    tf.WriteLine "$lbl.Location = New-Object System.Drawing.Point(12,12)"
    tf.WriteLine "$lbl.Size = New-Object System.Drawing.Size(490,20)"
    tf.WriteLine "$combo = New-Object System.Windows.Forms.ComboBox"
    tf.WriteLine "$combo.Location = New-Object System.Drawing.Point(12,40)"
    tf.WriteLine "$combo.Size = New-Object System.Drawing.Size(490,24)"
    tf.WriteLine "$combo.DropDownStyle = 'DropDownList'"

    Dim i
    For i = 0 To UBound(optionsArray)
        tf.WriteLine "$null = $combo.Items.Add('" & EscapePsSingle(CStr(optionsArray(i))) & "')"
    Next

    tf.WriteLine "$default = '" & EscapePsSingle(defaultValue) & "'"
    tf.WriteLine "$idx = $combo.Items.IndexOf($default)"
    tf.WriteLine "if ($idx -ge 0) { $combo.SelectedIndex = $idx } else { if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 } }"
    tf.WriteLine "$ok = New-Object System.Windows.Forms.Button"
    tf.WriteLine "$ok.Text = 'OK'"
    tf.WriteLine "$ok.Location = New-Object System.Drawing.Point(346,80)"
    tf.WriteLine "$ok.DialogResult = [System.Windows.Forms.DialogResult]::OK"
    tf.WriteLine "$cancel = New-Object System.Windows.Forms.Button"
    tf.WriteLine "$cancel.Text = 'Cancel'"
    tf.WriteLine "$cancel.Location = New-Object System.Drawing.Point(426,80)"
    tf.WriteLine "$cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel"
    tf.WriteLine "$form.Controls.Add($lbl)"
    tf.WriteLine "$form.Controls.Add($combo)"
    tf.WriteLine "$form.Controls.Add($ok)"
    tf.WriteLine "$form.Controls.Add($cancel)"
    tf.WriteLine "$form.AcceptButton = $ok"
    tf.WriteLine "$form.CancelButton = $cancel"
    tf.WriteLine "$result = $form.ShowDialog()"
    tf.WriteLine "if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $combo.SelectedItem) { Write-Output $combo.SelectedItem.ToString() }"

    tf.Close

    Dim cmd, exec, output
    cmd = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & tempPs & """"
    Set exec = shell.Exec(cmd)
    Do While exec.Status = 0
        WScript.Sleep 50
    Loop

    output = NormalizePickerValue(exec.StdOut.ReadAll)
    If output <> "" Then SelectFromDropdown = output

    On Error Resume Next
    If fso.FileExists(tempPs) Then fso.DeleteFile tempPs, True
End Function

Function EscapePsSingle(value)
    EscapePsSingle = Replace(CStr(value), "'", "''")
End Function

Function NormalizePickerValue(value)
    Dim s
    s = CStr(value)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr(160), " ")

    If Len(s) > 0 Then
        If AscW(Left(s, 1)) = &HFEFF Then
            s = Mid(s, 2)
        End If
    End If

    NormalizePickerValue = Trim(s)
End Function

Function BuildNonDxfSheetList(drawingDoc)
    BuildNonDxfSheetList = ""

    Dim idx
    idx = 0

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            idx = idx + 1
            If BuildNonDxfSheetList = "" Then
                BuildNonDxfSheetList = CStr(idx) & ") " & SafeText(sheet.Name)
            Else
                BuildNonDxfSheetList = BuildNonDxfSheetList & vbCrLf & CStr(idx) & ") " & SafeText(sheet.Name)
            End If
        End If
    Next
End Function

Function SheetExistsNonDxf(drawingDoc, sheetName)
    SheetExistsNonDxf = False

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If UCase(SafeText(sheet.Name)) = UCase(Trim(CStr(sheetName))) Then
            If Not IsDxfSheetName(sheet.Name) Then
                SheetExistsNonDxf = True
                Exit Function
            End If
        End If
    Next
End Function

Function GetAutoPlacementPoint(invApp, targetSheet, slotIndex)
    Dim col, row
    col = slotIndex Mod AUTO_LAYOUT_COLUMNS
    row = slotIndex \ AUTO_LAYOUT_COLUMNS

    Dim marginX, marginY, usableW, usableH, cellW, cellH
    marginX = targetSheet.Width * 0.08
    marginY = targetSheet.Height * 0.08
    usableW = targetSheet.Width - (2 * marginX)
    usableH = targetSheet.Height - (2 * marginY)
    cellW = usableW / AUTO_LAYOUT_COLUMNS
    cellH = usableH / AUTO_LAYOUT_ROWS_PER_SHEET

    Dim x, y
    x = marginX + ((col + 0.5) * cellW)
    y = targetSheet.Height - marginY - ((row + 0.5) * cellH)

    Set GetAutoPlacementPoint = invApp.TransientGeometry.CreatePoint2d(x, y)
End Function

Function CreateAutoPlacementSheet(drawingDoc, templateSheet, suffixIndex)
    On Error Resume Next

    Set CreateAutoPlacementSheet = drawingDoc.Sheets.Add(templateSheet.Size)
    If CreateAutoPlacementSheet Is Nothing Then
        Set CreateAutoPlacementSheet = templateSheet
        Exit Function
    End If

    Err.Clear
    CreateAutoPlacementSheet.Name = "AUTO-MISSING-" & CStr(suffixIndex)
End Function

Function GetNonDxfSheetByIndex(drawingDoc, nonDxfIndex)
    On Error Resume Next

    Set GetNonDxfSheetByIndex = Nothing

    Dim count
    count = 0

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            count = count + 1
            If count = nonDxfIndex Then
                Set GetNonDxfSheetByIndex = sheet
                Exit Function
            End If
        End If
    Next
End Function

Function ResolveTargetSheetForPlacement(drawingDoc, preferredName, absoluteIndex, fallbackNonDxfIndex)
    On Error Resume Next

    Set ResolveTargetSheetForPlacement = Nothing

    Dim sheet
    If Trim(CStr(preferredName)) <> "" Then
        For Each sheet In drawingDoc.Sheets
            If UCase(SafeText(sheet.Name)) = UCase(Trim(CStr(preferredName))) Then
                If Not IsDxfSheetName(sheet.Name) Then
                    Set ResolveTargetSheetForPlacement = sheet
                    Exit Function
                End If
            End If
        Next
    End If

    If absoluteIndex > 0 Then
        If drawingDoc.Sheets.Count >= absoluteIndex Then
            Set sheet = drawingDoc.Sheets.Item(absoluteIndex)
            If Not sheet Is Nothing Then
                If Not IsDxfSheetName(sheet.Name) Then
                    Set ResolveTargetSheetForPlacement = sheet
                    Exit Function
                End If
            End If
        End If
    End If

    Set ResolveTargetSheetForPlacement = GetNonDxfSheetByIndex(drawingDoc, fallbackNonDxfIndex)
End Function

Function GetSheetNumberToken(drawingDoc, targetSheet)
    On Error Resume Next

    GetSheetNumberToken = ""
    If targetSheet Is Nothing Then Exit Function

    Dim directIndex
    directIndex = 0

    Err.Clear
    directIndex = CLng(targetSheet.Index)
    If Err.Number = 0 Then
        GetSheetNumberToken = Right("0" & CStr(directIndex), 2)
        Exit Function
    End If
    Err.Clear

    Dim count
    count = 0

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        count = count + 1
        If UCase(SafeText(sheet.Name)) = UCase(SafeText(targetSheet.Name)) Then
            GetSheetNumberToken = Right("0" & CStr(count), 2)
            Exit Function
        End If
    Next
End Function

Function ParseOnOffArg(argName, defaultValue)
    On Error Resume Next

    ParseOnOffArg = defaultValue

    Dim raw
    raw = GetNamedArgValue(argName, "")
    raw = UCase(Trim(CStr(raw)))

    If raw = "" Then Exit Function
    If raw = "1" Or raw = "TRUE" Or raw = "ON" Or raw = "YES" Then
        ParseOnOffArg = True
        Exit Function
    End If
    If raw = "0" Or raw = "FALSE" Or raw = "OFF" Or raw = "NO" Then
        ParseOnOffArg = False
    End If
End Function

Function GetNamedArgValue(argName, defaultValue)
    On Error Resume Next

    GetNamedArgValue = defaultValue

    If WScript.Arguments.Named.Exists(argName) Then
        GetNamedArgValue = WScript.Arguments.Named.Item(argName)
    End If
End Function

Sub CollectRowModels(partsListRow, modelSheets, modelDocs, modelPartNumbers)
    On Error Resume Next

    Dim preferredPaths
    Set preferredPaths = GetPreferredRowModelPaths(partsListRow)
    If preferredPaths Is Nothing Then Exit Sub
    If preferredPaths.Count = 0 Then Exit Sub

    Dim keys
    keys = preferredPaths.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim modelPath
        modelPath = CStr(keys(i))

        If Not modelSheets.Exists(modelPath) Then
            Dim sheetDict
            Set sheetDict = CreateObject("Scripting.Dictionary")
            sheetDict.CompareMode = 1
            modelSheets.Add modelPath, sheetDict

            Dim modelDoc
            Set modelDoc = Nothing
            Set modelDoc = preferredPaths(modelPath)
            If Not modelDoc Is Nothing Then
                Set modelDocs(modelPath) = modelDoc
            End If

            Dim partNo
            partNo = GetPartNumberFromDocument(modelDoc)
            If partNo = "" Then partNo = GetBaseName(modelPath)
            modelPartNumbers.Add modelPath, partNo
        End If
    Next
End Sub

Sub CollectModelViewPlacements(drawingDoc, modelSheets)
    On Error Resume Next

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            Dim sheetNo
            sheetNo = GetSheetNumberToken(drawingDoc, sheet)

            Dim view
            For Each view In sheet.DrawingViews
                Dim modelPath
                modelPath = ResolveViewModelPath(view)

                If ENABLE_DEBUG_ROW_MATCH_LOGS Then
                    LogMessage "VIEWMAP: " & sheetNo & " | " & SafeText(modelPath)
                End If

                If modelPath <> "" Then
                    If modelSheets.Exists(modelPath) Then
                        If Not modelSheets(modelPath).Exists(sheetNo) Then
                            modelSheets(modelPath).Add sheetNo, True
                        End If
                    End If
                End If
            Next
        Else
            LogMessage "SKIP DXF SHEET (view placement): " & SafeText(sheet.Name)
        End If
    Next
End Sub

Sub InjectPartsListDwgRefCells(drawingDoc, modelSheets, idwBaseName, ByRef updatedCells)
    On Error Resume Next

    updatedCells = 0

    Dim sheet
    For Each sheet In drawingDoc.Sheets
        If Not IsDxfSheetName(sheet.Name) Then
            Dim partsList
            For Each partsList In sheet.PartsLists
                Dim targetColumnIndex
                targetColumnIndex = FindDwgRefColumnIndex(partsList)

                If targetColumnIndex > 0 Then
                    Dim row
                    For Each row In partsList.PartsListRows
                        Dim cell
                        Set cell = row.Item(targetColumnIndex)
                        If Not cell Is Nothing Then
                            Dim cellValue
                            If FORCE_TEST_MODE Then
                                cellValue = FORCE_TEST_VALUE
                            Else
                                cellValue = GetRowSheetRefs(row, modelSheets, idwBaseName)
                            End If

                            Err.Clear
                            cell.Value = cellValue
                            If Err.Number = 0 Then
                                updatedCells = updatedCells + 1
                            Else
                                LogMessage "WARN: Failed to set cell on sheet " & SafeText(sheet.Name) & ": " & Err.Description
                                Err.Clear
                            End If
                        End If
                    Next
                Else
                    LogMessage "INFO: No DWG REF column found on sheet " & SafeText(sheet.Name)
                End If
            Next
        Else
            LogMessage "SKIP DXF SHEET (cell update): " & SafeText(sheet.Name)
        End If
    Next
End Sub

Function GetRowSheetRefs(partsListRow, modelSheets, idwBaseName)
    On Error Resume Next

    GetRowSheetRefs = ""

    Dim aggregateSheets
    Set aggregateSheets = CreateObject("Scripting.Dictionary")
    aggregateSheets.CompareMode = 1

    Dim preferredPaths
    Set preferredPaths = GetPreferredRowModelPaths(partsListRow)
    If preferredPaths Is Nothing Then Exit Function
    If preferredPaths.Count = 0 Then Exit Function

    If ENABLE_DEBUG_ROW_MATCH_LOGS Then
        LogRowReferenceDebug partsListRow, preferredPaths
    End If

    Dim modelKeys
    modelKeys = preferredPaths.Keys

    Dim modelIdx
    For modelIdx = 0 To UBound(modelKeys)
        Dim modelPath
        modelPath = CStr(modelKeys(modelIdx))

        If modelSheets.Exists(modelPath) Then
            Dim sheetDict
            Set sheetDict = modelSheets(modelPath)

            Dim keys
            keys = sheetDict.Keys
            Dim i
            For i = 0 To UBound(keys)
                If Not aggregateSheets.Exists(keys(i)) Then
                    aggregateSheets.Add keys(i), True
                End If
            Next
        End If
    Next

    If aggregateSheets.Count > 0 Then
        GetRowSheetRefs = BuildSheetRefValue(aggregateSheets, idwBaseName)
    End If

    If ENABLE_DEBUG_ROW_MATCH_LOGS Then
        LogMessage "ROWREF_RESULT: " & GetRowDebugLabel(partsListRow) & " => " & SafeText(GetRowSheetRefs)
    End If
End Function

Function GetPreferredRowModelPaths(partsListRow)
    On Error Resume Next

    Set GetPreferredRowModelPaths = CreateObject("Scripting.Dictionary")
    GetPreferredRowModelPaths.CompareMode = 1

    Dim iptPaths
    Set iptPaths = CreateObject("Scripting.Dictionary")
    iptPaths.CompareMode = 1

    Dim iamPaths
    Set iamPaths = CreateObject("Scripting.Dictionary")
    iamPaths.CompareMode = 1

    Dim refFiles
    Set refFiles = partsListRow.ReferencedFiles
    If refFiles Is Nothing Then Exit Function

    Dim refFile
    For Each refFile In refFiles
        Dim modelPath
        modelPath = SafeText(refFile.FullFileName)
        If modelPath = "" Then
            Err.Clear
            modelPath = SafeText(refFile.ReferencedDocument.FullFileName)
        End If

        If modelPath <> "" Then
            Dim ext
            ext = LCase(Right(modelPath, 4))

            Dim refDoc
            Set refDoc = Nothing
            Set refDoc = refFile.ReferencedDocument

            If ext = ".ipt" Then
                If Not iptPaths.Exists(modelPath) Then iptPaths.Add modelPath, refDoc
            ElseIf ext = ".iam" Then
                If Not iamPaths.Exists(modelPath) Then iamPaths.Add modelPath, refDoc
            End If
        End If
    Next

    Dim sourceDict
    Set sourceDict = Nothing
    If iptPaths.Count > 0 Then
        Set sourceDict = iptPaths
    ElseIf iamPaths.Count > 0 Then
        Set sourceDict = iamPaths
    Else
        Exit Function
    End If

    Dim keys
    keys = sourceDict.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim k
        k = CStr(keys(i))
        If Not GetPreferredRowModelPaths.Exists(k) Then
            GetPreferredRowModelPaths.Add k, sourceDict(k)
        End If
    Next
End Function

Sub LogRowReferenceDebug(partsListRow, preferredPaths)
    On Error Resume Next

    Dim rawRefs
    rawRefs = GetRowReferencedModelList(partsListRow)

    Dim chosen
    chosen = JoinDictionaryKeys(preferredPaths)

    Dim chosenType
    chosenType = "none"
    If preferredPaths.Count > 0 Then
        Dim firstKey
        firstKey = CStr(preferredPaths.Keys()(0))
        chosenType = LCase(Right(firstKey, 4))
    End If

    LogMessage "ROWREF: " & GetRowDebugLabel(partsListRow) & " | RAW=" & rawRefs & " | CHOSEN_TYPE=" & chosenType & " | CHOSEN=" & chosen
End Sub

Function GetRowReferencedModelList(partsListRow)
    On Error Resume Next

    GetRowReferencedModelList = ""

    Dim refs
    Set refs = CreateObject("Scripting.Dictionary")
    refs.CompareMode = 1

    Dim refFiles
    Set refFiles = partsListRow.ReferencedFiles
    If refFiles Is Nothing Then Exit Function

    Dim refFile
    For Each refFile In refFiles
        Dim modelPath
        modelPath = SafeText(refFile.FullFileName)
        If modelPath = "" Then
            Err.Clear
            modelPath = SafeText(refFile.ReferencedDocument.FullFileName)
        End If

        If modelPath <> "" Then
            Dim ext
            ext = LCase(Right(modelPath, 4))
            If ext = ".ipt" Or ext = ".iam" Then
                If Not refs.Exists(modelPath) Then refs.Add modelPath, True
            End If
        End If
    Next

    GetRowReferencedModelList = JoinDictionaryKeys(refs)
End Function

Function JoinDictionaryKeys(dictObj)
    On Error Resume Next

    JoinDictionaryKeys = ""
    If dictObj Is Nothing Then Exit Function
    If dictObj.Count = 0 Then Exit Function

    Dim keys
    keys = dictObj.Keys
    SortStringArray keys

    Dim i, output
    output = ""
    For i = 0 To UBound(keys)
        If output = "" Then
            output = CStr(keys(i))
        Else
            output = output & " | " & CStr(keys(i))
        End If
    Next

    JoinDictionaryKeys = output
End Function

Function GetRowDebugLabel(partsListRow)
    On Error Resume Next

    GetRowDebugLabel = "Item?"

    Dim itemCell
    Set itemCell = Nothing
    Set itemCell = partsListRow.Item("ITEM")
    If Err.Number = 0 And Not itemCell Is Nothing Then
        Dim v
        v = SafeText(itemCell.Value)
        If v <> "" Then
            GetRowDebugLabel = "ITEM=" & v
            Exit Function
        End If
    End If
    Err.Clear

    Dim itemNumber
    itemNumber = SafeText(partsListRow.ItemNumber)
    If itemNumber <> "" Then
        GetRowDebugLabel = "ITEMNUM=" & itemNumber
    End If
End Function

Sub UpdateModelDwgRefs(modelSheets, modelDocs, idwBaseName, ByRef totalUpdated, ByRef totalErrors)
    On Error Resume Next

    Dim keys
    keys = modelSheets.Keys

    Dim i
    For i = 0 To UBound(keys)
        Dim modelPath
        modelPath = keys(i)

        Dim dwgRefValue
        If FORCE_TEST_MODE Then
            dwgRefValue = FORCE_TEST_VALUE
        Else
            dwgRefValue = BuildSheetRefValue(modelSheets(modelPath), idwBaseName)
        End If

        If modelDocs.Exists(modelPath) Then
            Dim modelDoc
            Set modelDoc = modelDocs(modelPath)

            If SetDwgRefAliasesOnDocument(modelDoc, dwgRefValue) Then
                totalUpdated = totalUpdated + 1
                LogMessage "UPDATED: " & SafeText(modelPath) & " => " & dwgRefValue
            Else
                totalErrors = totalErrors + 1
                LogMessage "ERROR: Failed updating " & SafeText(modelPath)
            End If
        Else
            totalErrors = totalErrors + 1
            LogMessage "ERROR: Missing model document handle for " & SafeText(modelPath)
        End If
    Next
End Sub

Sub GetUndetailedPartsReport(modelSheets, modelPartNumbers, ByRef partCount, ByRef reportText)
    On Error Resume Next

    partCount = 0
    reportText = ""

    Dim keys
    keys = modelSheets.Keys
    SortStringArray keys

    Dim i
    For i = 0 To UBound(keys)
        Dim modelPath
        modelPath = keys(i)

        If LCase(Right(modelPath, 4)) = ".ipt" Then
            If modelSheets(modelPath).Count = 0 Then
                partCount = partCount + 1

                Dim partNo
                partNo = modelPartNumbers(modelPath)
                If partNo = "" Then partNo = GetBaseName(modelPath)

                If reportText = "" Then
                    reportText = partNo
                Else
                    reportText = reportText & vbCrLf & partNo
                End If
            End If
        End If
    Next

    If reportText = "" Then reportText = "None"
End Sub

Function FindDwgRefColumnIndex(partsList)
    On Error Resume Next

    FindDwgRefColumnIndex = 0
    Dim idx
    For idx = 1 To partsList.PartsListColumns.Count
        Dim col
        Set col = partsList.PartsListColumns.Item(idx)
        If IsDwgRefTitle(SafeText(col.Title)) Then
            FindDwgRefColumnIndex = idx
            Exit Function
        End If
    Next
End Function

Function IsDwgRefTitle(value)
    Dim s
    s = UCase(CStr(value))
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    IsDwgRefTitle = (InStr(s, "DWGREF") > 0)
End Function

Function BuildSheetRefValue(sheetDict, idwBaseName)
    Dim keys
    keys = sheetDict.Keys
    SortStringArray keys

    Dim i, output
    output = ""
    For i = 0 To UBound(keys)
        If output = "" Then
            output = keys(i)
        Else
            output = output & "/" & keys(i)
        End If
    Next

    If output = "" Then
        BuildSheetRefValue = ""
    Else
        BuildSheetRefValue = idwBaseName & "-" & output
    End If
End Function

Function BuildDrawingSummary(modelSheets, modelPartNumbers, idwBaseName)
    Dim keys
    keys = modelSheets.Keys
    SortStringArray keys

    Dim i, outText
    outText = ""
    For i = 0 To UBound(keys)
        Dim modelPath, partNo, refText
        modelPath = keys(i)
        partNo = modelPartNumbers(modelPath)
        refText = BuildSheetRefValue(modelSheets(modelPath), idwBaseName)
        If outText = "" Then
            outText = partNo & "=" & refText
        Else
            outText = outText & "; " & partNo & "=" & refText
        End If
    Next
    BuildDrawingSummary = outText
End Function

Function SetDwgRefAliasesOnDocument(targetDoc, propertyValue)
    On Error Resume Next

    SetDwgRefAliasesOnDocument = False
    If targetDoc Is Nothing Then Exit Function

    Dim names
    names = Array("DWG REF", "DWG. REF.", "DWG_REF", "DWGREF")

    Dim i
    For i = 0 To UBound(names)
        If SetUserDefinedPropertyOnDocument(targetDoc, CStr(names(i)), propertyValue) Then
            SetDwgRefAliasesOnDocument = True
        End If
    Next
End Function

Function SetUserDefinedPropertyOnDocument(targetDoc, propertyName, propertyValue)
    On Error Resume Next

    SetUserDefinedPropertyOnDocument = False
    If targetDoc Is Nothing Then Exit Function

    Dim userProps
    Set userProps = targetDoc.PropertySets.Item("Inventor User Defined Properties")
    If Err.Number <> 0 Or userProps Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Dim targetProp
    Set targetProp = Nothing
    Set targetProp = userProps.Item(propertyName)

    If Err.Number <> 0 Or targetProp Is Nothing Then
        Err.Clear
        userProps.Add propertyValue, propertyName
    Else
        targetProp.Value = propertyValue
    End If

    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    targetDoc.Save
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    SetUserDefinedPropertyOnDocument = True
End Function

Sub SetUserDefinedProperty(drawingDoc, propertyName, propertyValue)
    On Error Resume Next

    Dim userProps
    Set userProps = drawingDoc.PropertySets.Item("Inventor User Defined Properties")

    Dim targetProp
    Set targetProp = Nothing
    Set targetProp = userProps.Item(propertyName)

    If Err.Number <> 0 Or targetProp Is Nothing Then
        Err.Clear
        userProps.Add propertyValue, propertyName
    Else
        targetProp.Value = propertyValue
    End If
End Sub

Function ResolveViewModelPath(view)
    On Error Resume Next

    ResolveViewModelPath = ""

    Err.Clear
    ResolveViewModelPath = SafeText(view.ReferencedDocumentDescriptor.FullDocumentName)
    If ResolveViewModelPath <> "" Then Exit Function

    Err.Clear
    ResolveViewModelPath = SafeText(view.ReferencedDocumentDescriptor.FullFileName)
    If ResolveViewModelPath <> "" Then Exit Function

    Err.Clear
    ResolveViewModelPath = SafeText(view.ReferencedDocumentDescriptor.ReferencedDocument.FullFileName)
End Function

Function GetPartNumberFromDocument(modelDoc)
    On Error Resume Next

    GetPartNumberFromDocument = ""
    If modelDoc Is Nothing Then Exit Function

    Dim designProps
    Set designProps = modelDoc.PropertySets.Item("Design Tracking Properties")
    If Err.Number <> 0 Or designProps Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Dim partNumberProp
    Set partNumberProp = designProps.Item("Part Number")
    If Err.Number <> 0 Or partNumberProp Is Nothing Then
        Err.Clear
        Exit Function
    End If

    GetPartNumberFromDocument = Trim(CStr(partNumberProp.Value))
End Function

Function IsDxfSheetName(sheetName)
    Dim s
    s = UCase(Trim(CStr(sheetName)))
    IsDxfSheetName = (Left(s, 3) = "DXF")
End Function

Function SafeText(value)
    On Error Resume Next
    SafeText = Trim(CStr(value))
    If Err.Number <> 0 Then
        SafeText = ""
        Err.Clear
    End If
End Function

Function GetBaseName(fullPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    GetBaseName = fso.GetBaseName(fullPath)
End Function

Sub SortStringArray(ByRef arr)
    On Error Resume Next

    If IsEmpty(arr) Then Exit Sub
    If UBound(arr) <= 0 Then Exit Sub

    Dim i, j, temp
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If UCase(CStr(arr(i))) > UCase(CStr(arr(j))) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next
    Next
End Sub

Sub StartLogging()
    On Error Resume Next

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim scriptDir
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

    Dim logsDir
    logsDir = scriptDir & "\Logs"
    If Not fso.FolderExists(logsDir) Then fso.CreateFolder logsDir

    g_LogPath = logsDir & "\Populate_DWG_REF_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
    Set g_LogFile = fso.CreateTextFile(g_LogPath, True)
End Sub

Sub LogMessage(message)
    On Error Resume Next
    If Not g_LogFile Is Nothing Then
        g_LogFile.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & " | " & message
    End If
End Sub

Sub StopLogging()
    On Error Resume Next
    If Not g_LogFile Is Nothing Then
        g_LogFile.Close
        Set g_LogFile = Nothing
    End If
End Sub
