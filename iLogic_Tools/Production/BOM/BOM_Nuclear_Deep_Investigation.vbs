' BOM Deep Investigation - Nuclear Diagnostic
' Comprehensive investigation into BOM display formatting and caching
' This script will reveal EXACTLY why formulas don't refresh
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp
Dim m_Report

Sub Main()
    On Error Resume Next

    m_Report = ""

    LogLine "=== BOM DEEP INVESTIGATION - NUCLEAR DIAGNOSTIC ==="
    LogLine "Date: " & Date & " " & Time
    LogLine ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        LogLine "ERROR: Inventor not running"
        ShowReport
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        LogLine "ERROR: No active document"
        ShowReport
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        LogLine "ERROR: Not an assembly - this diagnostic requires an assembly"
        ShowReport
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogLine "Assembly: " & asmDoc.DisplayName
    LogLine "Path: " & asmDoc.FullFileName
    LogLine ""

    ' === PHASE 1: BOM STRUCTURE ANALYSIS ===
    LogLine "=== PHASE 1: BOM STRUCTURE ANALYSIS ==="
    LogLine ""

    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM
    If bom Is Nothing Then
        LogLine "ERROR: Could not access BOM"
        ShowReport
        Exit Sub
    End If

    LogLine "BOM.StructuredViewEnabled: " & bom.StructuredViewEnabled
    LogLine "BOM.PartsOnlyViewEnabled: " & bom.PartsOnlyViewEnabled
    LogLine ""

    ' Enable structured view if needed
    If Not bom.StructuredViewEnabled Then
        LogLine "Enabling structured view..."
        bom.StructuredViewEnabled = True
        If Err.Number <> 0 Then
            LogLine "ERROR: Could not enable structured view - " & Err.Description
            Err.Clear
        End If
    End If

    ' Get BOM view
    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        LogLine "ERROR: Could not access Structured BOMView"
        ShowReport
        Exit Sub
    End If

    LogLine "BOMView.ViewState: " & bomView.ViewState
    LogLine "BOMView.BOMRows.Count: " & bomView.BOMRows.Count
    LogLine ""

    ' === PHASE 2: BOM CELL INVESTIGATION ===
    LogLine "=== PHASE 2: BOM CELL INVESTIGATION ==="
    LogLine ""

    If bomView.BOMRows.Count = 0 Then
        LogLine "WARNING: No BOM rows found"
    Else
        ' Investigate first few rows
        Dim rowCount
        rowCount = bomView.BOMRows.Count
        If rowCount > 3 Then rowCount = 3

        Dim i
        For i = 1 To rowCount
            Dim bomRow
            Set bomRow = bomView.BOMRows.Item(i)

            LogLine "--- BOMRow " & i & " ---"
            LogLine "  Component: " & bomRow.ComponentDefinitions.Item(1).Definition.Name
            LogLine "  ItemQuantity: " & bomRow.ItemQuantity
            LogLine "  TotalQuantity: " & bomRow.TotalQuantity

            ' Check BOMQuantity object
            On Error Resume Next
            Dim bomQty
            Set bomQty = bomRow.BOMQuantity

            If Err.Number = 0 Then
                LogLine "  BOMQuantity object: EXISTS"

                ' Try all possible properties
                Dim props
                props = Array("Precision", "DisplayFormat", "Expression", "Value", "Format", "DisplayPrecision", "DecimalPlaces", "NumberFormat")

                Dim prop
                For Each prop In props
                    Dim val
                    Execute "val = bomQty." & prop
                    If Err.Number = 0 Then
                        LogLine "    BOMQuantity." & prop & ": " & val
                    Else
                        LogLine "    BOMQuantity." & prop & ": (not found)"
                        Err.Clear
                    End If
                Next

                ' Check if it's a formula
                If bomQty.Expression <> "" Then
                    LogLine "    Expression: " & bomQty.Expression
                    LogLine "    Value: " & bomQty.Value
                End If
            Else
                LogLine "  BOMQuantity object: NOT ACCESSIBLE"
                Err.Clear
            End If

            LogLine ""
        Next
    End If

    ' === PHASE 3: BOM COLUMN INVESTIGATION ===
    LogLine "=== PHASE 3: BOM COLUMN INVESTIGATION ==="
    LogLine ""

    On Error Resume Next
    Dim bomColumns
    Set bomColumns = bomView.BOMColumns

    If Err.Number = 0 Then
        LogLine "BOMColumns.Count: " & bomColumns.Count

        Dim j
        For j = 1 To bomColumns.Count
            Dim col
            Set col = bomColumns.Item(j)
            LogLine "  Column " & j & ": " & col.Title & " (Property: " & col.Property & ")"

            ' Check for precision-related properties
            Dim colProps
            colProps = Array("Precision", "DisplayFormat", "Format", "DecimalPlaces")

            For Each prop In colProps
                Dim colVal
                Execute "colVal = col." & prop
                If Err.Number = 0 Then
                    LogLine "    Column." & prop & ": " & colVal
                Else
                    Err.Clear
                End If
            Next
        Next
    Else
        LogLine "BOMColumns: Not accessible"
        Err.Clear
    End If

    LogLine ""

    ' === PHASE 4: SAMPLE PART INVESTIGATION ===
    LogLine "=== PHASE 4: SAMPLE PART PRECISION INVESTIGATION ==="
    LogLine ""

    ' Find a plate part
    Dim foundPlate
    foundPlate = False

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim k
    For k = 1 To occurrences.Count
        If foundPlate Then Exit For

        Dim occ
        Set occ = occurrences.Item(k)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                Dim partNumber
                partNumber = ""
                On Error Resume Next
                partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                Err.Clear

                If InStr(UCase(partNumber), "PL") > 0 Then
                    LogLine "Sample plate part: " & partNumber
                    LogLine "  Path: " & refDoc.FullFileName

                    ' Check part precision settings
                    Dim partParams
                    Set partParams = refDoc.ComponentDefinition.Parameters

                    If Not partParams Is Nothing Then
                        LogLine "  LinearDimensionPrecision: " & partParams.LinearDimensionPrecision
                        LogLine "  DimensionDisplayType: " & partParams.DimensionDisplayType
                        LogLine "  DisplayParameterAsExpression: " & partParams.DisplayParameterAsExpression
                    End If

                    ' Check units
                    Dim partUnits
                    Set partUnits = refDoc.UnitsOfMeasure
                    If Not partUnits Is Nothing Then
                        LogLine "  LengthUnits: " & partUnits.LengthUnits
                        LogLine "  LengthDisplayUnits: " & partUnits.LengthDisplayUnits
                    End If

                    foundPlate = True
                End If
            End If
        End If
    Next

    If Not foundPlate Then
        LogLine "No plate parts found in assembly"
    End If

    LogLine ""

    ' === PHASE 5: BOM REFRESH EXPERIMENT ===
    LogLine "=== PHASE 5: BOM REFRESH EXPERIMENT ==="
    LogLine ""

    LogLine "Testing various refresh methods..."
    LogLine ""

    ' Method 1: BOMView.Renumber
    LogLine "Method 1: BOMView.Renumber()..."
    On Error Resume Next
    bomView.Renumber
    If Err.Number = 0 Then
        LogLine "  SUCCESS: Renumber completed"
    Else
        LogLine "  FAILED: " & Err.Description
        Err.Clear
    End If

    ' Method 2: Document Update
    LogLine "Method 2: asmDoc.Update()..."
    asmDoc.Update
    If Err.Number = 0 Then
        LogLine "  SUCCESS: Update completed"
    Else
        LogLine "  FAILED: " & Err.Description
        Err.Clear
    End If

    ' Method 3: Rebuild2
    LogLine "Method 3: asmDoc.Rebuild2(True)..."
    asmDoc.Rebuild2 True
    If Err.Number = 0 Then
        LogLine "  SUCCESS: Rebuild2 completed"
    Else
        LogLine "  FAILED: " & Err.Description
        Err.Clear
    End If

    ' Method 4: Save
    LogLine "Method 4: asmDoc.Save()..."
    asmDoc.Save
    If Err.Number = 0 Then
        LogLine "  SUCCESS: Save completed"
    Else
        LogLine "  FAILED: " & Err.Description
        Err.Clear
    End If

    LogLine ""

    ' === PHASE 6: HYPOTHESIS TESTING ===
    LogLine "=== PHASE 6: HYPOTHESIS TESTING ==="
    LogLine ""

    LogLine "HYPOTHESIS 1: BOM uses LinearDimensionPrecision for display")
    LogLine "  Current assembly LinearDimensionPrecision: " & asmDoc.ComponentDefinition.Parameters.LinearDimensionPrecision
    LogLine "  If BOM shows decimals, this hypothesis is WRONG"
    LogLine ""

    LogLine "HYPOTHESIS 2: BOM display is controlled by BOMQuantity.Precision")
    If bomView.BOMRows.Count > 0 Then
        Set bomRow = bomView.BOMRows.Item(1)
        On Error Resume Next
        Set bomQty = bomRow.BOMQuantity
        If Err.Number = 0 Then
            Dim qtyPrecision
            qtyPrecision = bomQty.Precision
            If Err.Number = 0 Then
                LogLine "  First row BOMQuantity.Precision: " & qtyPrecision
                LogLine "  If this controls display, we can set it directly"
            Else
                LogLine "  BOMQuantity.Precision property not found"
                Err.Clear
            End If
        End If
        Err.Clear
    End If
    LogLine ""

    LogLine "HYPOTHESIS 3: BOM display uses UnitsOfMeasure.LengthDisplayUnits")
    Dim asmUnits
    Set asmUnits = asmDoc.UnitsOfMeasure
    LogLine "  Current LengthDisplayUnits: " & asmUnits.LengthDisplayUnits
    LogLine "  If toggling this refreshes BOM, it's the key"
    LogLine ""

    LogLine "HYPOTHESIS 4: BOM display is cached at application level")
    LogLine "  If only document reopen works, cache is application-level"
    LogLine ""

    ' === PHASE 7: RECOMMENDATIONS ===
    LogLine "=== PHASE 7: RECOMMENDATIONS ==="
    LogLine ""

    LogLine "BASED ON THIS INVESTIGATION:"
    LogLine ""
    LogLine "1. Check BOM manually - does it show decimals or whole numbers?"
    LogLine "2. If decimals: LinearDimensionPrecision is NOT controlling BOM display"
    LogLine "3. If whole numbers: Our scripts work, but cache invalidation is the issue"
    LogLine ""
    LogLine "NEXT STEPS:"
    LogLine "1. Run this diagnostic on an assembly with BOM showing decimals"
    LogLine "2. Check if BOMQuantity.Precision exists and what value it has"
    LogLine "3. Try manually toggling LengthDisplayUnits in Document Settings"
    LogLine "4. See if that refreshes the BOM display"
    LogLine ""

    ShowReport

End Sub

Sub LogLine(msg)
    m_Report = m_Report & msg & vbCrLf
    WScript.Echo msg
End Sub

Sub ShowReport()
    ' Save report to file
    Dim fso, file, reportPath
    Set fso = CreateObject("Scripting.FileSystemObject")

    reportPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\BOM_NUCLEAR_DIAGNOSTIC_REPORT.txt"

    Set file = fso.CreateTextFile(reportPath, True)
    file.Write m_Report
    file.Close

    WScript.Echo ""
    WScript.Echo "=== REPORT SAVED ==="
    WScript.Echo "Location: " & reportPath

    MsgBox "Nuclear diagnostic complete!" & vbCrLf & vbCrLf & _
           "Report saved to:" & vbCrLf & _
           reportPath, vbInformation, "BOM Nuclear Diagnostic"
End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_Nuclear_Deep_Investigation.vbs