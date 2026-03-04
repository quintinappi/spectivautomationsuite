' Fix iLogic Stock Number Formula
' Modifies the iProperty formula to use ROUND() functions
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FIX iLOGIC STOCK NUMBER FORMULA ==="
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

    Dim updatedCount
    updatedCount = 0

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
                    WScript.Echo "--- FIXING FORMULA: " & partNumber & " ---"

                    If FixiLogicFormula(refDoc) Then
                        updatedCount = updatedCount + 1
                        refDoc.Save
                        WScript.Echo "  Formula updated and saved"
                    Else
                        WScript.Echo "  No changes needed or failed"
                    End If

                    WScript.Echo ""
                End If
            End If
        End If
    Next

    WScript.Echo "=== FORMULA FIX COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount
    WScript.Echo ""
    WScript.Echo "Stock Number formulas now use ROUND() functions."
    WScript.Echo "BOM should show whole numbers."

End Sub

Function FixiLogicFormula(partDoc)
    On Error Resume Next

    FixiLogicFormula = False

    ' Get design tracking properties
    Dim designProps
    Set designProps = partDoc.PropertySets.Item("Design Tracking Properties")

    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not access design properties"
        Exit Function
    End If

    ' Get Stock Number property
    Dim stockNumProp
    Set stockNumProp = designProps.Item("Stock Number")

    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: Could not access Stock Number property"
        Exit Function
    End If

    ' Get current formula
    Dim currentFormula
    currentFormula = stockNumProp.Expression

    If Err.Number <> 0 Or currentFormula = "" Then
        WScript.Echo "  No formula found"
        Exit Function
    End If

    WScript.Echo "  Current formula: " & currentFormula

    ' Check if it's the expected iLogic formula
    If InStr(currentFormula, "<sheet metal length>") = 0 Then
        WScript.Echo "  Not the expected iLogic formula"
        Exit Function
    End If

    ' Check if already has ROUND functions
    If InStr(LCase(currentFormula), "round(") > 0 Then
        WScript.Echo "  Formula already has ROUND functions"
        Exit Function
    End If

    ' Replace parameter references with ROUND versions
    ' From: =<sheet metal length> x <sheet metal width> x <thickness> PL S355JR
    ' To: =Round(<sheet metal length>) x Round(<sheet metal width>) x Round(<thickness>) PL S355JR

    Dim newFormula
    newFormula = currentFormula

    newFormula = Replace(newFormula, "<sheet metal length>", "Round(<sheet metal length>)")
    newFormula = Replace(newFormula, "<sheet metal width>", "Round(<sheet metal width>)")
    newFormula = Replace(newFormula, "<thickness>", "Round(<thickness>)")

    ' If formula changed, update it
    If newFormula <> currentFormula Then
        WScript.Echo "  New formula: " & newFormula

        stockNumProp.Expression = newFormula

        If Err.Number = 0 Then
            WScript.Echo "  Formula updated successfully"
            FixiLogicFormula = True
        Else
            WScript.Echo "  ERROR: Failed to update formula - " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "  No parameter references found to round"
    End If

End Function

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Fix_iLogic_Stock_Number_Formula.vbs