' Diagnose BOM Formula System
' Deep investigation into how BOM formulas are stored and evaluated
' This helps understand WHY formulas don't re-evaluate programmatically
' Author: Quintin de Bruin © 2026
'
' INVESTIGATES:
' - BOMRow object structure
' - BOMQuantity property type
' - Parameter/formula references in BOM cells
' - Display precision settings at BOM level
' - Any cached formatting properties
'
' OUTPUTS: Detailed diagnostic report

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp
Dim m_Report

Sub Main()
    On Error Resume Next

    m_Report = ""

    LogLine "=== BOM FORMULA SYSTEM DIAGNOSTIC ==="
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
        LogLine "ERROR: Not an assembly"
        ShowReport
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    LogLine "Assembly: " & asmDoc.DisplayName
    LogLine "Path: " & asmDoc.FullFileName
    LogLine ""

    ' === SECTION 1: ASSEMBLY PRECISION SETTINGS ===
    LogLine "=== SECTION 1: ASSEMBLY PRECISION SETTINGS ==="
    LogLine ""

    Dim params
    Set params = asmDoc.ComponentDefinition.Parameters
    If Not params Is Nothing Then
        LogLine "Parameters.LinearDimensionPrecision: " & params.LinearDimensionPrecision
        LogLine "Parameters.DimensionDisplayType: " & params.DimensionDisplayType
        LogLine "Parameters.DisplayParameterAsExpression: " & params.DisplayParameterAsExpression
    Else
        LogLine "ERROR: Could not access Parameters"
    End If

    Dim unitsOfMeasure
    Set unitsOfMeasure = asmDoc.UnitsOfMeasure
    If Not unitsOfMeasure Is Nothing Then
        LogLine "UnitsOfMeasure.LengthUnits: " & unitsOfMeasure.LengthUnits
        LogLine "UnitsOfMeasure.LengthDisplayUnits: " & unitsOfMeasure.LengthDisplayUnits
    Else
        LogLine "ERROR: Could not access UnitsOfMeasure"
    End If

    LogLine ""

    ' === SECTION 2: BOM OBJECT INVESTIGATION ===
    LogLine "=== SECTION 2: BOM OBJECT INVESTIGATION ==="
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

    If Not bom.StructuredViewEnabled Then
        LogLine "WARNING: Structured view not enabled - enabling it..."
        bom.StructuredViewEnabled = True
        If Err.Number <> 0 Then
            LogLine "ERROR: Could not enable structured view - " & Err.Description
            Err.Clear
        End If
    End If

    ' === SECTION 3: BOMVIEW OBJECT PROPERTIES ===
    LogLine "=== SECTION 3: BOMVIEW OBJECT PROPERTIES ==="
    LogLine ""

    Dim bomView
    Set bomView = bom.BOMViews.Item("Structured")
    If bomView Is Nothing Then
        LogLine "ERROR: Could not access Structured BOMView"
        ShowReport
        Exit Sub
    End If

    LogLine "BOMView.ViewState: " & bomView.ViewState
    LogLine "BOMView.BOMRowsCount: " & bomView.BOMRows.Count
    LogLine ""

    ' Check if BOMView has any precision-related properties
    LogLine "Investigating BOMView for precision properties..."
    On Error Resume Next

    ' Try to access various possible properties
    Dim testProp
    testProp = bomView.Precision
    If Err.Number = 0 Then
        LogLine "  BOMView.Precision: " & testProp
    Else
        LogLine "  BOMView.Precision: (property not found)"
        Err.Clear
    End If

    testProp = bomView.NumberFormat
    If Err.Number = 0 Then
        LogLine "  BOMView.NumberFormat: " & testProp
    Else
        LogLine "  BOMView.NumberFormat: (property not found)"
        Err.Clear
    End If

    LogLine ""

    ' === SECTION 4: BOMROW INVESTIGATION (FIRST 5 ROWS) ===
    LogLine "=== SECTION 4: BOMROW INVESTIGATION (FIRST 5 ROWS) ==="
    LogLine ""

    Dim rowCount
    rowCount = bomView.BOMRows.Count
    If rowCount > 5 Then rowCount = 5

    Dim i
    For i = 1 To rowCount
        Dim bomRow
        Set bomRow = bomView.BOMRows.Item(i)

        LogLine "--- BOMRow " & i & " ---"
        LogLine "  ComponentDefinition: " & bomRow.ComponentDefinitions.Item(1).Definition.Name
        LogLine "  ItemQuantity: " & bomRow.ItemQuantity
        LogLine "  TotalQuantity: " & bomRow.TotalQuantity

        ' Check BOMQuantity object properties
        On Error Resume Next

        Dim bomQty
        Set bomQty = bomRow.BOMQuantity
        If Err.Number = 0 Then
            LogLine "  BOMQuantity object exists"

            ' Try to access precision
            testProp = bomQty.Precision
            If Err.Number = 0 Then
                LogLine "    BOMQuantity.Precision: " & testProp
            Else
                LogLine "    BOMQuantity.Precision: (property not found)"
                Err.Clear
            End If

            ' Try to access display format
            testProp = bomQty.DisplayFormat
            If Err.Number = 0 Then
                LogLine "    BOMQuantity.DisplayFormat: " & testProp
            Else
                LogLine "    BOMQuantity.DisplayFormat: (property not found)"
                Err.Clear
            End If

            ' Try to access expression
            testProp = bomQty.Expression
            If Err.Number = 0 Then
                LogLine "    BOMQuantity.Expression: " & testProp
            Else
                LogLine "    BOMQuantity.Expression: (property not found)"
                Err.Clear
            End If
        Else
            LogLine "  BOMQuantity: (object not accessible)"
            Err.Clear
        End If

        LogLine ""
    Next

    ' === SECTION 5: PART PRECISION INVESTIGATION ===
    LogLine "=== SECTION 5: PART PRECISION INVESTIGATION (SAMPLE PLATE) ==="
    LogLine ""

    ' Find first plate part
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim foundPlate
    foundPlate = False

    Dim j
    For j = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(j)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                ' Check if it's a plate
                Dim partNumber
                partNumber = ""
                On Error Resume Next
                partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                Err.Clear

                If InStr(UCase(partNumber), "PL") > 0 Then
                    LogLine "Sample plate part: " & partNumber
                    LogLine "  Full path: " & refDoc.FullFileName

                    ' Get part precision settings
                    Dim partParams
                    Set partParams = refDoc.ComponentDefinition.Parameters
                    If Not partParams Is Nothing Then
                        LogLine "  Part LinearDimensionPrecision: " & partParams.LinearDimensionPrecision
                        LogLine "  Part DimensionDisplayType: " & partParams.DimensionDisplayType
                        LogLine "  Part DisplayParameterAsExpression: " & partParams.DisplayParameterAsExpression
                    End If

                    foundPlate = True
                    Exit For
                End If
            End If
        End If
    Next

    If Not foundPlate Then
        LogLine "No plate parts found in assembly"
    End If

    LogLine ""

    ' === SECTION 6: RECOMMENDATIONS ===
    LogLine "=== SECTION 6: RECOMMENDATIONS ==="
    LogLine ""

    LogLine "Based on investigation, try these approaches:"
    LogLine ""
    LogLine "APPROACH 1: LengthDisplayUnits toggle (not just LengthUnits)"
    LogLine "  - Toggle UnitsOfMeasure.LengthDisplayUnits instead"
    LogLine "  - Display units affect how formulas are formatted"
    LogLine ""
    LogLine "APPROACH 2: BOMView.Renumber() method"
    LogLine "  - Forces complete BOM rebuild"
    LogLine "  - Should trigger formula re-evaluation"
    LogLine ""
    LogLine "APPROACH 3: Nuclear document reopen"
    LogLine "  - Close and reopen assembly"
    LogLine "  - Guarantees formula re-evaluation from scratch"
    LogLine ""
    LogLine "APPROACH 4: Direct BOMQuantity manipulation"
    LogLine "  - If BOMQuantity has Precision property, set it directly"
    LogLine "  - Bypass formula system entirely"
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

    reportPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\BOM_DIAGNOSTIC_REPORT.txt"

    Set file = fso.CreateTextFile(reportPath, True)
    file.Write m_Report
    file.Close

    WScript.Echo ""
    WScript.Echo "=== REPORT SAVED ==="
    WScript.Echo "Location: " & reportPath

    MsgBox "Diagnostic complete!" & vbCrLf & vbCrLf & _
           "Report saved to:" & vbCrLf & _
           reportPath, vbInformation, "BOM Diagnostic"
End Sub

Main
