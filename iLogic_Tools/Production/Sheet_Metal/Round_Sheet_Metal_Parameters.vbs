' Round Sheet Metal Parameters to Whole Numbers
' Rounds sheet metal length, width, thickness parameters to whole numbers
' This ensures iProperty formulas display whole numbers
' Author: Quintin de Bruin © 2026

Option Explicit

Const kAssemblyDocumentObject = 12291

Dim m_InventorApp

Sub Main()
    On Error Resume Next

    WScript.Echo "=== ROUND SHEET METAL PARAMETERS ==="
    WScript.Echo "This will round length, width, thickness to whole numbers"
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
                    WScript.Echo "--- ROUNDING: " & partNumber & " ---"

                    If RoundSheetMetalParameters(refDoc) Then
                        updatedCount = updatedCount + 1
                        refDoc.Save
                        WScript.Echo "  Parameters rounded and saved"
                    Else
                        WScript.Echo "  No changes needed or failed"
                    End If

                    WScript.Echo ""
                End If
            End If
        End If
    Next

    WScript.Echo "=== ROUNDING COMPLETE ==="
    WScript.Echo "Parts updated: " & updatedCount
    WScript.Echo ""
    WScript.Echo "Sheet metal parameters are now whole numbers."
    WScript.Echo "BOM Stock Number should show whole numbers."

End Sub

Function RoundSheetMetalParameters(partDoc)
    On Error Resume Next

    RoundSheetMetalParameters = False

    Dim compDef
    Set compDef = partDoc.ComponentDefinition

    Dim params
    Set params = compDef.Parameters

    If params Is Nothing Then
        WScript.Echo "  ERROR: Could not access parameters"
        Exit Function
    End If

    ' Look for sheet metal parameters to round
    Dim paramNames
    paramNames = Array("Sheet Metal Length", "Sheet Metal Width", "Thickness", "Length", "Width", "Height", "Depth")

    Dim param
    Dim changed
    changed = False

    For Each param In params
        Dim paramName
        paramName = param.Name

        ' Check if this parameter should be rounded
        Dim shouldRound
        shouldRound = False

        Dim nameCheck
        For Each nameCheck In paramNames
            If InStr(LCase(paramName), LCase(nameCheck)) > 0 Then
                shouldRound = True
                Exit For
            End If
        Next

        If shouldRound Then
            ' Check if value has decimals
            Dim currentValue
            currentValue = param.Value

            Dim roundedValue
            roundedValue = Round(currentValue)

            If Abs(currentValue - roundedValue) > 0.001 Then ' Has decimals
                WScript.Echo "  Rounding " & paramName & ": " & currentValue & " -> " & roundedValue

                param.Value = roundedValue
                If Err.Number = 0 Then
                    changed = True
                Else
                    WScript.Echo "    ERROR: Failed to round - " & Err.Description
                    Err.Clear
                End If
            Else
                WScript.Echo "  " & paramName & " already whole number: " & currentValue
            End If
        End If
    Next

    RoundSheetMetalParameters = changed

End Function

Main</content>
<parameter name="filePath">c:\Users\Quintin\Documents\Spectiv\3. Working\FINAL_PRODUCTION_SCRIPTS 1 Oct 2025\iLogic_Tools\Round_Sheet_Metal_Parameters.vbs