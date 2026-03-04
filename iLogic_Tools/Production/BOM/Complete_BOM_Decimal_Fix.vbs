' Complete BOM Decimal Fix - Precision + Stock Number
' Sets precision AND fixes Stock Number display
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291
Const kZeroDecimalPlaceLinearPrecision = 0
Const kMillimeterLengthUnits = 11269
Const kCentimeterLengthUnits = 11266

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== COMPLETE BOM DECIMAL FIX ==="
    WScript.Echo "Sets precision AND fixes Stock Number display"
    WScript.Echo ""

    ' Connect to Inventor
    Set m_InventorApp = GetObject(, "Inventor.Application")
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Inventor not running"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument Is Nothing Then
        WScript.Echo "ERROR: No active document"
        Exit Sub
    End If

    If m_InventorApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        WScript.Echo "ERROR: Not an assembly"
        Exit Sub
    End If

    Dim asmDoc
    Set asmDoc = m_InventorApp.ActiveDocument
    WScript.Echo "Assembly: " & asmDoc.DisplayName
    WScript.Echo ""

    ' Step 1: Set precision on all plate parts
    WScript.Echo "STEP 1: Setting LinearDimensionPrecision = 0 on all plate parts..."
    WScript.Echo ""

    Dim precisionUpdated
    precisionUpdated = SetPrecisionOnPlates(asmDoc)

    WScript.Echo "Precision updated on " & precisionUpdated & " parts"
    WScript.Echo ""

    ' Step 2: Fix Stock Number display
    WScript.Echo "STEP 2: Fixing Stock Number display to show whole numbers..."
    WScript.Echo ""

    Dim stockNumFixed
    stockNumFixed = FixStockNumbers(asmDoc)

    WScript.Echo "Stock Numbers fixed on " & stockNumFixed & " parts"
    WScript.Echo ""

    ' Step 3: Force BOM refresh
    WScript.Echo "STEP 3: Forcing BOM refresh..."
    ForceBOMRefresh asmDoc

    WScript.Echo ""
    WScript.Echo "=== COMPLETE FIX FINISHED ==="
    WScript.Echo ""
    WScript.Echo "BOM should now show whole numbers in Stock Number column."
    WScript.Echo "If not, close and reopen the assembly."

End Sub

Function SetPrecisionOnPlates(asmDoc)
    On Error Resume Next

    SetPrecisionOnPlates = 0

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                Dim desc
                desc = ""
                On Error Resume Next
                desc = refDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear

                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
                    ' Set precision
                    Dim compDef
                    Set compDef = refDoc.ComponentDefinition

                    Dim params
                    Set params = compDef.Parameters

                    If Not params Is Nothing Then
                        params.LinearDimensionPrecision = kZeroDecimalPlaceLinearPrecision
                        params.DimensionDisplayType = 34821 ' Display as value
                        params.DisplayParameterAsExpression = True

                        ' Force units refresh event
                        ForceUnitsRefreshEvent refDoc

                        refDoc.Save

                        SetPrecisionOnPlates = SetPrecisionOnPlates + 1
                    End If
                End If
            End If
        End If
    Next

End Function

Sub ForceUnitsRefreshEvent(partDoc)
    On Error Resume Next

    Dim unitsOfMeasure
    Set unitsOfMeasure = partDoc.UnitsOfMeasure

    If Not unitsOfMeasure Is Nothing Then
        Dim originalUnits
        originalUnits = unitsOfMeasure.LengthUnits

        ' Toggle units to trigger cache invalidation
        If originalUnits = kMillimeterLengthUnits Then
            unitsOfMeasure.LengthUnits = kCentimeterLengthUnits
        Else
            unitsOfMeasure.LengthUnits = kMillimeterLengthUnits
        End If

        partDoc.Update

        ' Restore original
        unitsOfMeasure.LengthUnits = originalUnits
        partDoc.Update
    End If

End Sub

Function FixStockNumbers(asmDoc)
    On Error Resume Next

    FixStockNumbers = 0

    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

        If Not occ.Suppressed Then
            Dim refDoc
            Set refDoc = occ.Definition.Document

            If LCase(Right(refDoc.FullFileName, 4)) = ".ipt" Then
                Dim desc2
                desc2 = ""
                On Error Resume Next
                desc2 = refDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear

                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc2), "PL") > 0 Or InStr(UCase(desc2), "VRN") > 0 Or InStr(UCase(desc2), "S355JR") > 0 Then
                    ' Fix Stock Number
                    If FixSingleStockNumber(refDoc) Then
                        refDoc.Save
                        FixStockNumbers = FixStockNumbers + 1
                    End If
                End If
            End If
        End If
    Next

End Function

Function FixSingleStockNumber(partDoc)
    On Error Resume Next

    FixSingleStockNumber = False

    ' Get design tracking properties
    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then Exit Function

    ' Get Stock Number property
    Dim stockNumProp
    Set stockNumProp = designProps.Item("Stock Number")

    If Err.Number <> 0 Then Exit Function

    ' Get current value
    Dim currentValue
    currentValue = stockNumProp.Value

    ' Parse and round decimal numbers
    Dim newValue
    newValue = currentValue

    ' Replace "number.decimals mm" with "number mm"
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "(\d+)\.(\d+)( mm)"
    regex.Global = True

    Dim matches
    Set matches = regex.Execute(currentValue)

    Dim match
    For Each match In matches
        Dim fullMatch
        fullMatch = match.Value

        Dim integerPart
        integerPart = match.SubMatches(0)

        Dim unitPart
        unitPart = match.SubMatches(2)

        Dim roundedMatch
        roundedMatch = integerPart & unitPart

        newValue = Replace(newValue, fullMatch, roundedMatch)
    Next

    ' Update if changed
    If newValue <> currentValue Then
        stockNumProp.Value = newValue
        If Err.Number = 0 Then
            FixSingleStockNumber = True
        Else
            Err.Clear
        End If
    End If

End Function

Sub ForceBOMRefresh(asmDoc)
    On Error Resume Next

    Dim bom
    Set bom = asmDoc.ComponentDefinition.BOM

    If Not bom Is Nothing Then
        If bom.StructuredViewEnabled Then
            Dim bomView
            Set bomView = bom.BOMViews.Item("Structured")

            If Not bomView Is Nothing Then
                bomView.Renumber
            End If
        End If
    End If

    asmDoc.Update
    asmDoc.Rebuild2 True

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Complete_BOM_Decimal_Fix.vbs