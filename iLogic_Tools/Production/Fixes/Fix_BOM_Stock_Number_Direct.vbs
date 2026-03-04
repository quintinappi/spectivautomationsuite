' Fix BOM Stock Number - Direct Value Modification
' Directly modifies Stock Number iProperty to show rounded values
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== FIX BOM STOCK NUMBER - DIRECT ==="
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

                    If FixStockNumberDirect(refDoc) Then
                        updatedCount = updatedCount + 1
                        refDoc.Save
                        WScript.Echo "  Stock Number fixed and saved"
                    Else
                        WScript.Echo "  No changes needed or failed"
                    End If

                    WScript.Echo ""
                End If
            End If
        End If
    Next

    WScript.Echo "=== FIX COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount
    WScript.Echo ""
    WScript.Echo "BOM Stock Number should now show whole numbers."

End Sub

Function FixStockNumberDirect(partDoc)
    On Error Resume Next

    FixStockNumberDirect = False

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

    ' Get current value
    Dim currentValue
    currentValue = stockNumProp.Value

    WScript.Echo "  Current Stock Number: " & currentValue

    ' Parse the value and round numbers
    ' Format: "3850.000 mm x 1365.000 mm x 6 mm PL S355JR"
    ' Should become: "3850 mm x 1365 mm x 6 mm PL S355JR"

    Dim newValue
    newValue = currentValue

    ' Replace decimal numbers with rounded versions
    ' Pattern: number with decimals followed by " mm"
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

    ' If value changed, update it
    If newValue <> currentValue Then
        WScript.Echo "  New Stock Number: " & newValue

        stockNumProp.Value = newValue

        If Err.Number = 0 Then
            WScript.Echo "  Stock Number updated successfully"
            FixStockNumberDirect = True
        Else
            WScript.Echo "  ERROR: Failed to update Stock Number - " & Err.Description
            Err.Clear
        End If
    Else
        WScript.Echo "  Stock Number already shows whole numbers"
    End If

End Function

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Fix_BOM_Stock_Number_Direct.vbs