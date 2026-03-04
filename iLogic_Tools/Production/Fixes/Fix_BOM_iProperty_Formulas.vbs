' Fix BOM iProperty Formulas - Round Values to Whole Numbers
' Updates Stock Number iProperty formulas to use ROUND() for whole number display
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FIX BOM IPROPERTY FORMULAS ==="
    WScript.Echo "This will update Stock Number formulas to show whole numbers"
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
                    WScript.Echo "--- UPDATING: " & partNumber & " ---"

                    If UpdateStockNumberFormula(refDoc) Then
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

    WScript.Echo "=== UPDATE COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount
    WScript.Echo ""
    WScript.Echo "The BOM should now show whole numbers in Stock Number column."
    WScript.Echo "If not, try closing and reopening the assembly."

End Sub

Function UpdateStockNumberFormula(partDoc)
    On Error Resume Next

    UpdateStockNumberFormula = False

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
        WScript.Echo "  No formula found (static value)"
        Exit Function
    End If

    WScript.Echo "  Current formula: " & currentFormula

    ' Check if formula already has rounding
    If InStr(LCase(currentFormula), "round(") > 0 Or InStr(LCase(currentFormula), "int(") > 0 Then
        WScript.Echo "  Formula already has rounding - skipping"
        Exit Function
    End If

    ' Pattern: look for parameter references like <Parameter Name>
    ' We need to wrap them with Round()

    Dim newFormula
    newFormula = currentFormula

    ' Replace parameter references with Round(parameter)
    ' Pattern: <Parameter Name> -> Round(<Parameter Name>)

    Dim regex
    Set regex = New RegExp
    regex.Pattern = "<[^>]+>"
    regex.Global = True

    Dim matches
    Set matches = regex.Execute(currentFormula)

    Dim match
    For Each match In matches
        Dim paramRef
        paramRef = match.Value

        ' Skip if it's already wrapped
        If InStr(newFormula, "Round(" & paramRef) = 0 Then
            newFormula = Replace(newFormula, paramRef, "Round(" & paramRef & ")")
        End If
    Next

    ' If formula changed, update it
    If newFormula <> currentFormula Then
        WScript.Echo "  New formula: " & newFormula

        stockNumProp.Expression = newFormula

        If Err.Number = 0 Then
            WScript.Echo "  Formula updated successfully"
            UpdateStockNumberFormula = True
        Else
            WScript.Echo "  ERROR: Failed to update formula - " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "  No parameter references found to round"
    End If

End Function

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Fix_BOM_iProperty_Formulas.vbs