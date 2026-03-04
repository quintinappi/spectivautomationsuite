' Fix iLogic Formula with Format Function
' Uses Format() instead of Round() to display whole numbers
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FIX iLOGIC FORMULA WITH FORMAT ==="
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
                    WScript.Echo "--- FIXING: " & partNumber & " ---"

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

    WScript.Echo "=== FORMAT FIX COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount
    WScript.Echo ""
    WScript.Echo "Stock Number formulas now use Format() for whole numbers."

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

    ' Check if already has Format functions
    If InStr(LCase(currentFormula), "format(") > 0 Then
        WScript.Echo "  Formula already has Format functions"
        Exit Function
    End If

    ' Replace parameter references with Format versions
    ' From: =<sheet metal length> x <sheet metal width> x <thickness> PL S355JR
    ' To: =Format(<sheet metal length>, "0") x Format(<sheet metal width>, "0") x Format(<thickness>, "0") PL S355JR

    Dim newFormula
    newFormula = currentFormula

    newFormula = Replace(newFormula, "<sheet metal length>", "Format(<sheet metal length>, ""0"")")
    newFormula = Replace(newFormula, "<sheet metal width>", "Format(<sheet metal width>, ""0"")")
    newFormula = Replace(newFormula, "<thickness>", "Format(<thickness>, ""0"")")

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
        WScript.Echo "  No parameter references found to format"
    End If

End Function

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Fix_iLogic_Formula_With_Format.vbs