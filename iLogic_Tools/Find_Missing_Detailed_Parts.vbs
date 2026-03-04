' Find Missing Detailed Parts
' Scans Parts List on page 1, then checks all other pages to find parts that haven't been detailed
' Also detects parts detailed with wrong file references (prefix mismatches)
' Author: Quintin de Bruin © 2026

Option Explicit

Const kDrawingDocumentObject = 12292

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FIND MISSING DETAILED PARTS ==="
    WScript.Echo ""

    ' Get Inventor application
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        WScript.Quit 1
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        WScript.Quit 1
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kDrawingDocumentObject Then
        WScript.Echo "ERROR: Not a drawing document"
        WScript.Quit 1
    End If

    Dim drawDoc
    Set drawDoc = m_InventorApp.ActiveDocument

    If drawDoc.Sheets.Count < 1 Then
        WScript.Echo "ERROR: Drawing has no sheets"
        WScript.Quit 1
    End If

    ' Step 1: Get all components from Parts List on page 1
    WScript.Echo "STEP 1: Scanning Parts List on page 1..."
    WScript.Echo "========================================"
    Dim sheet1
    Set sheet1 = drawDoc.Sheets.Item(1)

    Dim allComponents
    Set allComponents = GetComponentsFromPartsList(sheet1)

    If allComponents.Count = 0 Then
        WScript.Echo "No Parts List found on page 1."
        WScript.Quit 1
    End If

    WScript.Echo "Found " & allComponents.Count & " items in Parts List"
    WScript.Echo ""

    ' Step 2: Scan all other pages to find which parts are detailed
    WScript.Echo "STEP 2: Scanning all other pages for detailed parts..."
    WScript.Echo "======================================================="

    Dim detailedComponents
    Set detailedComponents = CreateObject("Scripting.Dictionary")

    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim i
    For i = 2 To drawDoc.Sheets.Count
        Dim sheet
        Set sheet = drawDoc.Sheets.Item(i)

        WScript.Echo "Scanning Sheet " & i & ": " & sheet.Name

        Dim componentsOnSheet
        Set componentsOnSheet = GetComponentsFromSheetViews(sheet)

        ' Add to detailed components set with sheet number
        Dim key
        For Each key In componentsOnSheet.Keys
            If Not detailedComponents.Exists(key) Then
                detailedComponents.Add key, "Sheet " & i
            End If
        Next

        WScript.Echo "  Found " & componentsOnSheet.Count & " components"
    Next

    WScript.Echo ""
    WScript.Echo "Total detailed components found: " & detailedComponents.Count
    WScript.Echo ""

    ' Step 3: Find missing parts and incorrectly detailed parts (wrong prefix)
    WScript.Echo "STEP 3: Analyzing parts..."
    WScript.Echo "============================"

    Dim missingParts
    Set missingParts = CreateObject("Scripting.Dictionary")

    Dim incorrectParts
    Set incorrectParts = CreateObject("Scripting.Dictionary")

    Dim partName
    Dim plKey, detailKey

    For Each plKey In allComponents.Keys
        partName = fso.GetFileName(plKey)
        Dim partNumber
        partNumber = ExtractPartNumber(partName)  ' Get "FL60" from "NSCR05-779-FL60.ipt"

        ' Check if exact match exists in detailed components
        If detailedComponents.Exists(plKey) Then
            ' Perfect match - part is correctly detailed
        Else
            ' No exact match - check if there's a part number match with different prefix (wrong file)
            Dim foundAsWrongFile
            foundAsWrongFile = False
            Dim wrongFilePath

            For Each detailKey In detailedComponents.Keys
                Dim detailFileName
                detailFileName = fso.GetFileName(detailKey)
                Dim detailPartNumber
                detailPartNumber = ExtractPartNumber(detailFileName)

                If detailPartNumber = partNumber And detailFileName <> partName Then
                    ' Same part number but different filename - wrong file reference!
                    foundAsWrongFile = True
                    wrongFilePath = detailKey
                    Exit For
                End If
            Next

            If foundAsWrongFile Then
                incorrectParts.Add partName, "PL: " & GetShortName(plKey) & " | View: " & GetShortName(wrongFilePath)
            Else
                ' No match at all - part is not detailed
                missingParts.Add partName, plKey
            End If
        End If
    Next

    WScript.Echo ""

    ' Show incorrectly detailed parts first (these need fixing!)
    If incorrectParts.Count > 0 Then
        WScript.Echo "*** WARNING: INCORRECTLY DETAILED PARTS (" & incorrectParts.Count & " parts with wrong file references) ***"
        WScript.Echo ""
        WScript.Echo "These parts are detailed but reference the WRONG FILE (wrong prefix/mismatch):"
        WScript.Echo ""

        Dim count
        count = 1
        For Each partName In incorrectParts.Keys
            WScript.Echo "  " & count & ". " & partName
            WScript.Echo "     " & incorrectParts(partName)
            count = count + 1
        Next
        WScript.Echo ""
    End If

    ' Check for extra parts in views that are NOT in Parts List
    Dim extraParts
    Set extraParts = CreateObject("Scripting.Dictionary")

    For Each detailKey In detailedComponents.Keys
        detailFileName = fso.GetFileName(detailKey)
        Dim foundInPL
        foundInPL = False

        For Each plKey In allComponents.Keys
            If fso.GetFileName(plKey) = detailFileName Then
                foundInPL = True
                Exit For
            End If
        Next

        If Not foundInPL Then
            If Not extraParts.Exists(detailFileName) Then
                extraParts.Add detailFileName, detailKey
            End If
        End If
    Next

    ' Show extra parts
    If extraParts.Count > 0 Then
        WScript.Echo "*** WARNING: EXTRA PARTS IN VIEWS (" & extraParts.Count & " parts detailed but NOT in Parts List) ***"
        WScript.Echo ""
        WScript.Echo "These parts are detailed but don't appear in the Parts List:"
        WScript.Echo ""

        count = 1
        Dim partKey
        For Each partName In extraParts.Keys
            partKey = extraParts(partName)
            WScript.Echo "  " & count & ". " & partName & " (" & detailedComponents(partKey) & ")"
            count = count + 1
        Next
        WScript.Echo ""
    End If

    ' Show missing parts
    If missingParts.Count > 0 Then
        WScript.Echo "*** MISSING PARTS (" & missingParts.Count & " parts not detailed) ***"
        WScript.Echo ""

        count = 1
        For Each partName In missingParts.Keys
            WScript.Echo "  " & count & ". " & partName
            count = count + 1
        Next
        WScript.Echo ""
    End If

    ' Show success if no issues
    If incorrectParts.Count = 0 And missingParts.Count = 0 And extraParts.Count = 0 Then
        WScript.Echo "*** ALL PARTS CORRECTLY DETAILED! ***"
        WScript.Echo ""
    End If

    WScript.Echo "Summary:"
    WScript.Echo "  Parts List items: " & allComponents.Count
    WScript.Echo "  Detailed components: " & detailedComponents.Count
    WScript.Echo "  Incorrectly detailed (wrong file): " & incorrectParts.Count
    WScript.Echo "  Extra parts (not in Parts List): " & extraParts.Count
    WScript.Echo "  Missing (not detailed): " & missingParts.Count

End Sub

' Helper function to get short filename from full path
Function GetShortName(fullPath)
    Dim fso2
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    GetShortName = fso2.GetFileName(fullPath)
End Function

' Helper function to extract part number from filename
' "NSCR05-779-FL60.ipt" → "FL60"
' "N1SCR04-780-PL5.ipt" → "PL5"
Function ExtractPartNumber(fileName)
    Dim baseName
    baseName = fileName  ' fileName is just "NSCR05-779-FL60.ipt" (no path)

    ' Remove extension
    Dim pos
    pos = InStrRev(baseName, ".")
    If pos > 0 Then
        baseName = Left(baseName, pos - 1)
    End If

    ' Extract part number (everything after the last dash)
    pos = InStrRev(baseName, "-")
    If pos > 0 Then
        ExtractPartNumber = Mid(baseName, pos + 1)
    Else
        ExtractPartNumber = baseName
    End If
End Function

Function GetComponentsFromPartsList(sheet)
    Dim components
    Set components = CreateObject("Scripting.Dictionary")

    On Error Resume Next

    ' Get all Parts Lists on this sheet
    Dim partsLists
    Set partsLists = sheet.PartsLists

    If partsLists Is Nothing Or partsLists.Count = 0 Then
        WScript.Echo "  No Parts List found on this sheet."
        Set GetComponentsFromPartsList = components
        Exit Function
    End If

    ' Process each Parts List
    Dim pl
    For Each pl In partsLists
        Dim partsList
        Set partsList = pl

        ' Iterate through all rows in the Parts List
        Dim row
        For Each row In partsList.PartsListRows
            Dim partsListRow
            Set partsListRow = row

            ' Get the referenced files for this row
            Dim refFiles
            Set refFiles = partsListRow.ReferencedFiles

            If Not refFiles Is Nothing Then
                Dim refFile
                For Each refFile In refFiles
                    Dim fileName
                    fileName = refFile.FullFileName

                    If Not components.Exists(fileName) Then
                        components.Add fileName, True
                    End If
                Next
            End If
        Next
    Next

    On Error GoTo 0
    Set GetComponentsFromPartsList = components
End Function

Function GetComponentsFromSheetViews(sheet)
    Dim components
    Set components = CreateObject("Scripting.Dictionary")

    Dim view
    For Each view In sheet.DrawingViews
        On Error Resume Next
        If Not view Is Nothing Then
            Dim refDocDesc
            Set refDocDesc = view.ReferencedDocumentDescriptor

            If Err.Number <> 0 Then
                Err.Clear
            ElseIf Not refDocDesc Is Nothing Then
                Dim doc
                Set doc = refDocDesc.ReferencedDocument

                If Err.Number <> 0 Then
                    Err.Clear
                ElseIf Not doc Is Nothing Then
                    If LCase(Right(doc.FullFileName, 4)) = ".ipt" Then
                        ' Only count individual part views
                        If Not components.Exists(doc.FullFileName) Then
                            components.Add doc.FullFileName, True
                        End If
                    End If
                    ' Ignore assembly views (.iam) on detail sheets
                End If
            End If
        End If
        On Error GoTo 0
    Next

    Set GetComponentsFromSheetViews = components
End Function

Main
