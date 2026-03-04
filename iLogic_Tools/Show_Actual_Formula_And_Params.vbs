' Show Actual iProperty Formula and Parameters
' Displays the exact formula and parameter names
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== SHOW ACTUAL IPROPERTY FORMULA ==="
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

    ' Find first plate part
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim foundPlate
    foundPlate = False

    Dim i
    For i = 1 To occurrences.Count
        Dim occ
        Set occ = occurrences.Item(i)

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
                    WScript.Echo "--- FIRST PLATE PART FOUND ---"
                    WScript.Echo "Part Number: " & partNumber
                    WScript.Echo "File: " & refDoc.FullFileName
                    WScript.Echo ""

                    ShowActualFormulaAndParams refDoc
                    foundPlate = True
                    Exit For
                End If
            End If
        End If
    Next

    If Not foundPlate Then
        WScript.Echo "No plate parts found!"
        WScript.Echo ""
        WScript.Echo "Please open an assembly with plate parts (containing 'PL' in part number)"
    End If

End Sub

Sub ShowActualFormulaAndParams(partDoc)
    On Error Resume Next

    WScript.Echo "=== STOCK NUMBER IPROPERTY ==="

    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")

    If Err.Number = 0 Then
        Dim stockNumProp
        Set stockNumProp = designProps.Item("Stock Number")

        If Err.Number = 0 Then
            WScript.Echo "Stock Number Value: " & stockNumProp.Value

            Dim formula
            formula = stockNumProp.Expression
            If Err.Number = 0 And formula <> "" Then
                WScript.Echo "Stock Number Formula: " & formula
            Else
                WScript.Echo "Stock Number Formula: (none - static value)"
                Err.Clear
            End If
        Else
            WScript.Echo "Stock Number: (not found)"
            Err.Clear
        End If
    Else
        WScript.Echo "Design Tracking Properties: (not accessible)"
        Err.Clear
    End If

    WScript.Echo ""
    WScript.Echo "=== ALL PARAMETERS IN PART ==="

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    If Not params Is Nothing Then
        WScript.Echo "Total Parameters: " & params.Count
        WScript.Echo ""

        Dim param
        For Each param In params
            WScript.Echo "Parameter: " & param.Name & " = " & param.Value & " (Expr: " & param.Expression & ")"
        Next
    Else
        WScript.Echo "Parameters: (not accessible)"
    End If

    WScript.Echo ""
    WScript.Echo "=== SHEET METAL SPECIFIC ==="

    If compDef.IsSheetMetalPart Then
        WScript.Echo "Is Sheet Metal Part: YES"

        Dim sheetMetalComp
        Set sheetMetalComp = compDef.SheetMetalComponent

        If Not sheetMetalComp Is Nothing Then
            WScript.Echo "Thickness: " & sheetMetalComp.Thickness.Value & " (Expr: " & sheetMetalComp.Thickness.Expression & ")"

            ' Try to get extents
            On Error Resume Next
            Dim length, width
            length = sheetMetalComp.Length
            width = sheetMetalComp.Width

            If Err.Number = 0 Then
                WScript.Echo "Length: " & length
                WScript.Echo "Width: " & width
            Else
                WScript.Echo "Length/Width: (not directly accessible)"
                Err.Clear
            End If
        End If
    Else
        WScript.Echo "Is Sheet Metal Part: NO"
    End If

    WScript.Echo ""
    WScript.Echo "=== ANALYSIS ==="
    WScript.Echo ""
    WScript.Echo "Look at the formula and parameter names above."
    WScript.Echo "The formula uses parameter names like <ParameterName>"
    WScript.Echo "Find the actual parameter names that correspond to length, width, thickness"
    WScript.Echo "Then modify the formula to use Round(<ActualParameterName>)"

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Show_Actual_Formula_And_Params.vbs