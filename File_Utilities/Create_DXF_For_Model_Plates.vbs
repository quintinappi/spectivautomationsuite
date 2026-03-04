' ==============================================================================
' CREATE DXF FOR MODEL PLATES
' ==============================================================================
' - Prompts user with dropdown to choose a non-DXF source sheet
' - Finds assembly model referenced on that sheet
' - Creates sheet: DXF FOR {MODEL NAME}
' - Places all PLATE parts from that assembly at 1:1 scale
' - For sheet metal plates, places FLAT pattern view
' - Creates parts list and filters rows to plates only
' ==============================================================================

Option Explicit

Const kDrawingDocumentObject = 12292
Const kAssemblyDocumentObject = 12291
Const kFrontViewOrientation = 10764
Const kHiddenLineRemovedDrawingViewStyle = 32258
Const kSheetMetalSubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
Const VIEW_SCALE = 1
Const LAYOUT_COLS = 4
Const LAYOUT_ROWS = 3

Dim g_LogFile
Dim g_LogPath

Call Main()

Sub Main()
    On Error Resume Next

    StartLogging
    LogMessage "=== CREATE DXF FOR MODEL PLATES START ==="

    Dim invApp
    Set invApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Or invApp Is Nothing Then
        LogMessage "ERROR: Inventor is not running"
        MsgBox "Inventor is not running.", vbCritical, "CREATE DXF FOR MODEL PLATES"
        StopLogging
        Exit Sub
    End If
    Err.Clear

    Dim drawDoc
    Set drawDoc = invApp.ActiveDocument
    If drawDoc Is Nothing Or drawDoc.DocumentType <> kDrawingDocumentObject Then
        LogMessage "ERROR: Active document is not IDW"
        MsgBox "Active document must be an IDW.", vbCritical, "CREATE DXF FOR MODEL PLATES"
        StopLogging
        Exit Sub
    End If

    LogMessage "Drawing: " & SafeText(drawDoc.FullFileName)

    Dim sourceSheetName
    sourceSheetName = SelectSourceSheetName(drawDoc)
    If sourceSheetName = "" Then
        LogMessage "Cancelled: no source sheet selected"
        StopLogging
        Exit Sub
    End If
    LogMessage "Selected source sheet: " & sourceSheetName

    Dim sourceSheet
    Set sourceSheet = GetSheetByName(drawDoc, sourceSheetName)
    If sourceSheet Is Nothing Then
        LogMessage "ERROR: Selected sheet not found after normalization"
        MsgBox "Selected sheet not found.", vbExclamation, "CREATE DXF FOR MODEL PLATES"
        StopLogging
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = FindAssemblyDocumentOnSheet(invApp, sourceSheet)
    If asmDoc Is Nothing Then
        LogMessage "ERROR: No assembly view found on selected sheet"
        MsgBox "No assembly view found on selected sheet.", vbExclamation, "CREATE DXF FOR MODEL PLATES"
        StopLogging
        Exit Sub
    End If
    LogMessage "Assembly detected: " & SafeText(asmDoc.FullFileName)

    Dim modelName
    modelName = GetBaseName(SafeText(asmDoc.FullFileName))

    Dim targetSheet
    Set targetSheet = drawDoc.Sheets.Add(sourceSheet.Size)
    If targetSheet Is Nothing Then
        MsgBox "Failed to create DXF sheet.", vbCritical, "CREATE DXF FOR MODEL PLATES"
        Exit Sub
    End If

    Dim targetName
    targetName = BuildDxfSheetName(modelName)
    Err.Clear
    targetSheet.Name = targetName
    Err.Clear

    Dim plates
    Set plates = CollectPlateParts(asmDoc, sourceSheet)
    LogMessage "Plate candidates collected: " & CStr(plates.Count)
    If plates.Count = 0 Then
        LogMessage "ERROR: No plate parts found after assembly + parts-list scan"
        MsgBox "No plate parts found in assembly: " & modelName, vbInformation, "CREATE DXF FOR MODEL PLATES"
        MsgBox "No plate parts found. Check log:" & vbCrLf & g_LogPath, vbExclamation, "CREATE DXF FOR MODEL PLATES"
        StopLogging
        Exit Sub
    End If

    Dim firstDxfSheet
    Set firstDxfSheet = targetSheet

    Dim maxSlots, slot, extraSheetIndex
    maxSlots = LAYOUT_COLS * LAYOUT_ROWS
    slot = 0
    extraSheetIndex = 0

    Dim keys
    keys = plates.Keys
    SortStringArray keys

    Dim i, placedViews
    placedViews = 0

    For i = 0 To UBound(keys)
        Dim partPath
        partPath = CStr(keys(i))

        Dim pDoc
        Set pDoc = plates(partPath)
        If pDoc Is Nothing Then
            Set pDoc = invApp.Documents.Open(partPath, True)
        End If

        If Not pDoc Is Nothing Then
            If slot >= maxSlots Then
                extraSheetIndex = extraSheetIndex + 1
                Set targetSheet = drawDoc.Sheets.Add(sourceSheet.Size)
                targetSheet.Name = BuildDxfSheetName(modelName) & "-" & CStr(extraSheetIndex + 1)
                slot = 0
            End If

            Dim pt
            Set pt = GetGridPoint(invApp, targetSheet, slot, LAYOUT_COLS, LAYOUT_ROWS)

            Dim viewOk
            viewOk = AddPlateView(invApp, targetSheet, pDoc, pt, VIEW_SCALE)
            If viewOk Then placedViews = placedViews + 1

            slot = slot + 1
        End If
    Next

    Dim partsListCount
    partsListCount = CreateAndFilterPlatePartsList(invApp, firstDxfSheet, asmDoc, plates, modelName)

    drawDoc.Update
    drawDoc.Save

    MsgBox "CREATE DXF FOR MODEL PLATES complete." & vbCrLf & vbCrLf & _
           "Source sheet: " & sourceSheetName & vbCrLf & _
           "Assembly: " & modelName & vbCrLf & _
           "Plate parts found: " & plates.Count & vbCrLf & _
           "Views placed: " & placedViews & vbCrLf & _
           "Parts list rows visible: " & partsListCount, vbInformation, "CREATE DXF FOR MODEL PLATES"

    LogMessage "Views placed: " & CStr(placedViews)
    LogMessage "Parts list rows visible: " & CStr(partsListCount)
    LogMessage "Log path: " & g_LogPath
    StopLogging
End Sub

Function SelectSourceSheetName(drawDoc)
    On Error Resume Next

    SelectSourceSheetName = ""

    Dim names
    names = BuildNonDxfSheetNameArray(drawDoc)
    If IsEmpty(names) Then Exit Function

    SelectSourceSheetName = NormalizePickerValue(SelectFromDropdown("CREATE DXF FOR MODEL PLATES", "Select source sheet (must contain assembly view):", names, CStr(names(0))))
End Function

Function FindAssemblyDocumentOnSheet(invApp, sheet)
    On Error Resume Next

    Set FindAssemblyDocumentOnSheet = Nothing

    Dim v
    For Each v In sheet.DrawingViews
        Dim path
        path = ResolveViewModelPath(v)
        If LCase(Right(path, 4)) = ".iam" Then
            Dim d
            Set d = Nothing
            Err.Clear
            Set d = v.ReferencedDocumentDescriptor.ReferencedDocument
            If d Is Nothing Then
                Err.Clear
                Set d = invApp.Documents.Open(path, True)
            End If
            If Not d Is Nothing Then
                Set FindAssemblyDocumentOnSheet = d
                Exit Function
            End If
        End If
    Next
End Function

Function CollectPlateParts(asmDoc, sourceSheet)
    On Error Resume Next

    Set CollectPlateParts = CreateObject("Scripting.Dictionary")
    CollectPlateParts.CompareMode = 1

    Dim fromAsm
    Set fromAsm = CollectPlatePartsFromAssembly(asmDoc)

    Dim k
    For Each k In fromAsm.Keys
        If Not CollectPlateParts.Exists(CStr(k)) Then
            CollectPlateParts.Add CStr(k), fromAsm(k)
        End If
    Next

    Dim fromPl
    Set fromPl = CollectPlatePartsFromSourcePartsLists(sourceSheet)
    For Each k In fromPl.Keys
        If Not CollectPlateParts.Exists(CStr(k)) Then
            CollectPlateParts.Add CStr(k), fromPl(k)
        End If
    Next
End Function

Function CollectPlatePartsFromAssembly(asmDoc)
    On Error Resume Next

    Set CollectPlatePartsFromAssembly = CreateObject("Scripting.Dictionary")
    CollectPlatePartsFromAssembly.CompareMode = 1

    If asmDoc Is Nothing Then Exit Function
    If asmDoc.DocumentType <> kAssemblyDocumentObject Then Exit Function

    Dim occ
    For Each occ In asmDoc.ComponentDefinition.Occurrences.AllLeafOccurrences
        Dim partDoc
        Set partDoc = Nothing
        Err.Clear
        Set partDoc = occ.Definition.Document

        If Not partDoc Is Nothing Then
            If LCase(Right(SafeText(partDoc.FullFileName), 4)) = ".ipt" Then
                If IsPlatePart(partDoc) Then
                    If Not CollectPlatePartsFromAssembly.Exists(partDoc.FullFileName) Then
                        CollectPlatePartsFromAssembly.Add partDoc.FullFileName, partDoc
                        LogMessage "PLATE+ASM: " & GetBaseName(partDoc.FullFileName) & " | Desc=" & GetDescription(partDoc)
                    End If
                Else
                    LogMessage "SKIP+ASM: " & GetBaseName(partDoc.FullFileName) & " | Desc=" & GetDescription(partDoc)
                End If
            End If
        End If
    Next
End Function

Function CollectPlatePartsFromSourcePartsLists(sourceSheet)
    On Error Resume Next

    Set CollectPlatePartsFromSourcePartsLists = CreateObject("Scripting.Dictionary")
    CollectPlatePartsFromSourcePartsLists.CompareMode = 1

    If sourceSheet Is Nothing Then Exit Function

    Dim pl
    For Each pl In sourceSheet.PartsLists
        Dim row
        For Each row In pl.PartsListRows
            Dim refs
            Set refs = row.ReferencedFiles
            If Not refs Is Nothing Then
                Dim rf
                For Each rf In refs
                    Dim pth
                    pth = SafeText(rf.FullFileName)
                    If pth = "" Then
                        Err.Clear
                        pth = SafeText(rf.ReferencedDocument.FullFileName)
                    End If

                    If LCase(Right(pth, 4)) = ".ipt" Then
                        Dim pd
                        Set pd = Nothing
                        Err.Clear
                        Set pd = rf.ReferencedDocument

                        If Not pd Is Nothing Then
                            If IsPlatePart(pd) Then
                                If Not CollectPlatePartsFromSourcePartsLists.Exists(pth) Then
                                    CollectPlatePartsFromSourcePartsLists.Add pth, pd
                                    LogMessage "PLATE+PLIST: " & GetBaseName(pth) & " | Desc=" & GetDescription(pd)
                                End If
                            Else
                                LogMessage "SKIP+PLIST: " & GetBaseName(pth) & " | Desc=" & GetDescription(pd)
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next
End Function

Function IsPlatePart(partDoc)
    On Error Resume Next

    IsPlatePart = False
    If partDoc Is Nothing Then Exit Function

    Dim nameText, partNo, descText, matText
    nameText = UCase(GetBaseName(SafeText(partDoc.FullFileName)))
    partNo = UCase(GetPartNumber(partDoc))
    descText = UCase(GetDescription(partDoc))
    matText = UCase(GetMaterial(partDoc))

    If UCase(SafeText(partDoc.SubType)) = UCase(kSheetMetalSubType) Then
        IsPlatePart = True
        Exit Function
    End If

    If InStr(nameText, "-PL") > 0 Or InStr(nameText, "-LPL") > 0 Or Left(nameText, 2) = "PL" Then
        IsPlatePart = True
        Exit Function
    End If

    If InStr(partNo, "-PL") > 0 Or InStr(partNo, "-LPL") > 0 Or Left(partNo, 2) = "PL" Then
        IsPlatePart = True
        Exit Function
    End If

    If InStr(descText, "PL") > 0 Or InStr(descText, "PLA") > 0 Or InStr(descText, "VRN") > 0 Or InStr(descText, "S355JR") > 0 Then
        IsPlatePart = True
        Exit Function
    End If

    If InStr(matText, "PL") > 0 Or InStr(matText, "VRN") > 0 Or InStr(matText, "S355JR") > 0 Then
        IsPlatePart = True
    End If
End Function

Function AddPlateView(invApp, targetSheet, partDoc, pt, scaleValue)
    On Error Resume Next

    AddPlateView = False

    If partDoc Is Nothing Then Exit Function

    If UCase(SafeText(partDoc.SubType)) = UCase(kSheetMetalSubType) Then
        EnsureFlatPatternExists partDoc
        Dim opts
        Set opts = invApp.TransientObjects.CreateNameValueMap
        opts.Add "SheetMetalFoldedModel", False

        Err.Clear
        targetSheet.DrawingViews.AddBaseView partDoc, pt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle, "", Nothing, opts
        If Err.Number = 0 Then AddPlateView = True
        Err.Clear
    Else
        Err.Clear
        targetSheet.DrawingViews.AddBaseView partDoc, pt, scaleValue, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle
        If Err.Number = 0 Then AddPlateView = True
        Err.Clear
    End If
End Function

Function CreateAndFilterPlatePartsList(invApp, targetSheet, asmDoc, plates, modelName)
    On Error Resume Next

    CreateAndFilterPlatePartsList = 0

    Dim tg
    Set tg = invApp.TransientGeometry

    Dim asmPt
    Set asmPt = tg.CreatePoint2d(targetSheet.Width - 2, 2)

    Dim asmView
    Set asmView = Nothing
    Err.Clear
    Set asmView = targetSheet.DrawingViews.AddBaseView(asmDoc, asmPt, 0.01, kFrontViewOrientation, kHiddenLineRemovedDrawingViewStyle)
    If Err.Number <> 0 Or asmView Is Nothing Then
        Err.Clear
        Exit Function
    End If

    Dim plPt
    Set plPt = tg.CreatePoint2d(targetSheet.Width - 40, targetSheet.Height - 40)

    Dim pl
    Set pl = targetSheet.PartsLists.Add(asmView, plPt)
    If pl Is Nothing Then Exit Function

    Err.Clear
    pl.Title = modelName & "- BEAM PARTS LIST"
    Err.Clear

    Dim plateNames
    Set plateNames = CreateObject("Scripting.Dictionary")
    plateNames.CompareMode = 1

    Dim k
    For Each k In plates.Keys
        Dim bn
        bn = UCase(GetBaseName(CStr(k)))
        If Not plateNames.Exists(bn) Then plateNames.Add bn, True
    Next

    Dim visibleCount
    visibleCount = 0

    Dim row
    For Each row In pl.PartsListRows
        Dim rowKey
        rowKey = UCase(Trim(CStr(row.Item(2).Value)))
        If rowKey = "" Then rowKey = UCase(Trim(CStr(row.Item(1).Value)))

        If plateNames.Exists(rowKey) Then
            row.Visible = True
            visibleCount = visibleCount + 1
        Else
            row.Visible = False
        End If
    Next

    CreateAndFilterPlatePartsList = visibleCount
End Function

Sub EnsureFlatPatternExists(partDoc)
    On Error Resume Next

    If partDoc Is Nothing Then Exit Sub
    If UCase(SafeText(partDoc.SubType)) <> UCase(kSheetMetalSubType) Then Exit Sub

    Dim smDef
    Set smDef = partDoc.ComponentDefinition
    If smDef Is Nothing Then Exit Sub

    If Not smDef.HasFlatPattern Then
        smDef.Unfold
        Err.Clear
    End If
End Sub

Function GetGridPoint(invApp, targetSheet, slotIndex, cols, rows)
    Dim col, row
    col = slotIndex Mod cols
    row = slotIndex \ cols

    Dim marginX, marginY, usableW, usableH, cellW, cellH
    marginX = targetSheet.Width * 0.08
    marginY = targetSheet.Height * 0.08
    usableW = targetSheet.Width - (2 * marginX)
    usableH = targetSheet.Height - (2 * marginY)
    cellW = usableW / cols
    cellH = usableH / rows

    Dim x, y
    x = marginX + ((col + 0.5) * cellW)
    y = targetSheet.Height - marginY - ((row + 0.5) * cellH)

    Set GetGridPoint = invApp.TransientGeometry.CreatePoint2d(x, y)
End Function

Function BuildDxfSheetName(modelName)
    Dim n
    n = "DXF FOR " & modelName
    If Len(n) > 31 Then n = Left(n, 31)
    BuildDxfSheetName = n
End Function

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

Function GetPartNumber(partDoc)
    On Error Resume Next

    GetPartNumber = ""
    If partDoc Is Nothing Then Exit Function

    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")
    If designProps Is Nothing Then Exit Function

    Dim p
    Set p = designProps.Item("Part Number")
    If p Is Nothing Then Exit Function

    GetPartNumber = Trim(CStr(p.Value))
End Function

Function GetDescription(partDoc)
    On Error Resume Next

    GetDescription = ""
    If partDoc Is Nothing Then Exit Function

    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")
    If designProps Is Nothing Then Exit Function

    Dim p
    Set p = designProps.Item("Description")
    If p Is Nothing Then Exit Function

    GetDescription = Trim(CStr(p.Value))
End Function

Function GetMaterial(partDoc)
    On Error Resume Next

    GetMaterial = ""
    If partDoc Is Nothing Then Exit Function

    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")
    If designProps Is Nothing Then Exit Function

    Dim p
    Set p = designProps.Item("Material")
    If p Is Nothing Then Exit Function

    GetMaterial = Trim(CStr(p.Value))
End Function

Function GetSheetByName(drawDoc, sheetName)
    Set GetSheetByName = Nothing

    Dim target
    target = NormalizePickerValue(sheetName)
    If target = "" Then Exit Function

    Dim s
    For Each s In drawDoc.Sheets
        If UCase(NormalizePickerValue(s.Name)) = UCase(target) Then
            Set GetSheetByName = s
            Exit Function
        End If
    Next
End Function

Function BuildNonDxfSheetNameArray(drawDoc)
    BuildNonDxfSheetNameArray = Empty

    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim s
    For Each s In drawDoc.Sheets
        If Not IsDxfSheetName(s.Name) Then
            If Not d.Exists(CStr(s.Name)) Then d.Add CStr(s.Name), True
        End If
    Next

    If d.Count > 0 Then BuildNonDxfSheetNameArray = d.Keys
End Function

Function SelectFromDropdown(windowTitle, promptText, optionsArray, defaultValue)
    On Error Resume Next

    SelectFromDropdown = ""
    If IsEmpty(optionsArray) Then Exit Function

    Dim fso, shell
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    Dim tempPs
    tempPs = shell.ExpandEnvironmentStrings("%TEMP%") & "\\inv_dropdown_" & Replace(CStr(Timer), ".", "") & ".ps1"

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

    output = Trim(CStr(exec.StdOut.ReadAll))
    output = NormalizePickerValue(output)
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

    g_LogPath = logsDir & "\Create_DXF_For_Model_Plates_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & ".log"
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
