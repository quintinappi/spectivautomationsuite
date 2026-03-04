' BOM iProperty Formula Investigation
' Investigate how Stock Number iProperty formulas work
' Check parameter values vs display formatting
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== BOM IPROPERTY FORMULA INVESTIGATION ==="
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

    ' Find plate parts
    Dim occurrences
    Set occurrences = asmDoc.ComponentDefinition.Occurrences

    Dim plateCount
    plateCount = 0

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
                    plateCount = plateCount + 1

                    WScript.Echo "--- PLATE PART " & plateCount & " ---"
                    WScript.Echo "Part: " & partNumber
                    WScript.Echo "Path: " & refDoc.FullFileName

                    ' Check iProperties
                    InvestigatePartProperties refDoc

                    WScript.Echo ""

                    ' Only check first few plates
                    If plateCount >= 3 Then Exit For
                End If
            End If
        End If
    Next

    WScript.Echo "=== ANALYSIS ==="
    WScript.Echo ""
    WScript.Echo "The issue: Stock Number iProperty uses parameter VALUES, not display formatting"
    WScript.Echo ""
    WScript.Echo "SOLUTION OPTIONS:"
    WScript.Echo "1. Round parameter values to whole numbers"
    WScript.Echo "2. Modify iProperty formula to use ROUND() function"
    WScript.Echo "3. Create custom iProperty with integer formatting"
    WScript.Echo ""

End Sub

Sub InvestigatePartProperties(partDoc)
    On Error Resume Next

    ' Check design tracking properties
    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")

    If Err.Number = 0 Then
        Dim stockNumber
        stockNumber = designProps.Item("Stock Number").Value
        WScript.Echo "Stock Number: " & stockNumber

        ' Check if Stock Number has a formula
        Dim stockNumProp
        Set stockNumProp = designProps.Item("Stock Number")

        Dim formula
        formula = stockNumProp.Expression
        If Err.Number = 0 And formula <> "" Then
            WScript.Echo "Stock Number Formula: " & formula
        Else
            WScript.Echo "Stock Number: (no formula, static value)"
            Err.Clear
        End If
    End If
    Err.Clear

    ' Check parameters used in formula
    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    WScript.Echo "LinearDimensionPrecision: " & params.LinearDimensionPrecision

    ' Look for sheet metal parameters
    Dim param
    For Each param In params
        Dim paramName
        paramName = LCase(param.Name)

        If InStr(paramName, "length") > 0 Or InStr(paramName, "width") > 0 Or InStr(paramName, "thickness") > 0 Then
            WScript.Echo "Parameter: " & param.Name & " = " & param.Value & " (Expression: " & param.Expression & ")"
        End If
    Next

    ' Check custom properties
    Dim customProps
    Set customProps = partDoc.PropertySets.Item("Inventor User Defined Properties")

    If Err.Number = 0 Then
        WScript.Echo "Custom Properties:"
        Dim j
        For j = 1 To customProps.Count
            Dim customProp
            Set customProp = customProps.Item(j)
            WScript.Echo "  " & customProp.Name & ": " & customProp.Value
        Next
    Else
        WScript.Echo "No custom properties"
        Err.Clear
    End If

End Sub

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\BOM_iProperty_Formula_Investigation.vbs