' BOM DIAGNOSTIC - Show Current State
' Just shows current parameter values and iProperty formulas
' No modifications, just investigation
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BOM DIAGNOSTIC - CURRENT STATE ==="
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
                Dim desc
                desc = ""
                On Error Resume Next
                desc = refDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
                Err.Clear

                ' Check Description for PL, VRN, or S355JR
                If InStr(UCase(desc), "PL") > 0 Or InStr(UCase(desc), "VRN") > 0 Or InStr(UCase(desc), "S355JR") > 0 Then
                    Dim partNumber
                    partNumber = ""
                    On Error Resume Next
                    partNumber = refDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
                    Err.Clear
                    
                    WScript.Echo "--- FIRST PLATE FOUND ---"
                    WScript.Echo "Part Number: " & partNumber
                    WScript.Echo "Description: " & desc
                    WScript.Echo "File: " & refDoc.FullFileName
                    WScript.Echo ""

                    ShowPartDetails refDoc
                    foundPlate = True
                    Exit For
                End If
            End If
        End If
    Next

    If Not foundPlate Then
        WScript.Echo "No plate parts found!"
    End If

End Sub

Sub ShowPartDetails(partDoc)
    On Error Resume Next

    ' Show iProperties
    WScript.Echo "=== IPROPERTIES ==="

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

    ' Show parameters
    WScript.Echo "=== PARAMETERS ==="

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    If Not params Is Nothing Then
        WScript.Echo "LinearDimensionPrecision: " & params.LinearDimensionPrecision
        WScript.Echo ""

        WScript.Echo "All Parameters:"
        Dim param
        For Each param In params
            WScript.Echo "  " & param.Name & " = " & param.Value & " (Expr: " & param.Expression & ")"
        Next
    Else
        WScript.Echo "Parameters: (not accessible)"
    End If

    WScript.Echo ""

    ' Show sheet metal specific
    WScript.Echo "=== SHEET METAL INFO ==="

    If compDef.IsSheetMetalPart Then
        WScript.Echo "Is Sheet Metal Part: YES"

        ' Try to get sheet metal component
        Dim sheetMetalComp
        Set sheetMetalComp = compDef.SheetMetalComponent

        If Not sheetMetalComp Is Nothing Then
            WScript.Echo "Sheet Metal Thickness: " & sheetMetalComp.Thickness.Value
            WScript.Echo "Sheet Metal Thickness Expr: " & sheetMetalComp.Thickness.Expression
        End If
    Else
        WScript.Echo "Is Sheet Metal Part: NO"
    End If

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_Diagnostic_Show_Current.vbs